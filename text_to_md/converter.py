from __future__ import annotations

from dataclasses import dataclass
from html import escape as escape_html
from io import BytesIO
import re
from zipfile import BadZipFile, ZipFile

from docx import Document
from docx.document import Document as DocumentObject
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.text.paragraph import CT_P
from docx.oxml.table import CT_Tbl
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph
from pypdf import PdfReader
from pypdf.errors import PdfReadError


WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
HEADING_STYLE = re.compile(r"^(?:heading|nadpis)\s*(\d+)$", re.IGNORECASE)
MESSAGE_LABEL = re.compile(
    r"^(?:(?:obsah|tělo)\s+e-?mailu|předmět\s+(?:e-?mailu|mailu)|"
    r"pre-?header(?:\s+e-?mailu)?|"
    r"text\s+(?:sms|push)|titulek\s+push)\s*:?\s*$",
    re.IGNORECASE,
)
EMAIL_BODY_LABEL = re.compile(
    r"^(?:obsah|tělo)\s+e-?mailu\s*:?\s*$",
    re.IGNORECASE,
)
EMAIL_SUBJECT_LABEL = re.compile(
    r"^předmět\s+(?:e-?mailu|mailu)\s*:?\s*$",
    re.IGNORECASE,
)
EMAIL_PREHEADER_LABEL = re.compile(
    r"^pre-?header(?:\s+e-?mailu)?\s*:?\s*$",
    re.IGNORECASE,
)
EMAIL_INLINE_FIELD = re.compile(
    r"^(?P<label>předmět\s+(?:e-?mailu|mailu)|"
    r"pre-?header(?:\s+e-?mailu)?)\s*:\s*(?P<value>.+)$",
    re.IGNORECASE,
)
EMAIL_VARIANT = re.compile(r"^varianta pro .+$", re.IGNORECASE)
COMMUNICATION_HEADING = re.compile(
    r"^(?:\d+(?:\.\d+)*\s+)?(?P<channel>sms|e-?mail|push(?:\s+notifikace)?)$",
    re.IGNORECASE,
)
FIELD_LABEL = re.compile(r"^([^:\n]{1,80}):\s*(.+)$")


@dataclass(frozen=True)
class WordComment:
    comment_id: str
    author: str
    text: str
    quoted_text: str


def _iter_blocks(document: DocumentObject):
    for child in document.element.body.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, document)
        elif isinstance(child, CT_Tbl):
            yield Table(child, document)


def _is_formatted_title(paragraph: Paragraph) -> bool:
    runs = [run for run in paragraph.runs if run.text.strip()]
    return (
        paragraph.paragraph_format.alignment == WD_ALIGN_PARAGRAPH.CENTER
        and bool(runs)
        and all(run.bold for run in runs)
        and all(run.font.size is not None and run.font.size.pt >= 14 for run in runs)
    )


def _list_marker(paragraph: Paragraph, numbering_root) -> str:
    properties = paragraph._p.pPr
    numbering = properties.numPr if properties is not None else None
    if numbering is None or numbering.numId is None:
        return ""

    num_id = numbering.numId.val
    level = numbering.ilvl.val if numbering.ilvl is not None else 0
    num = next(
        (item for item in numbering_root.num_lst if item.numId == num_id),
        None,
    )
    if num is None:
        return ""

    abstract_id = num.abstractNumId.val
    abstract = next(
        (
            item for item in numbering_root.xpath("./w:abstractNum")
            if item.get(qn("w:abstractNumId")) == str(abstract_id)
        ),
        None,
    )
    if abstract is None:
        return ""

    level_element = next(
        (
            item for item in abstract.findall(qn("w:lvl"))
            if item.get(qn("w:ilvl")) == str(level)
        ),
        None,
    )
    if level_element is None:
        return ""

    format_element = level_element.find(qn("w:numFmt"))
    if format_element is None:
        return ""
    marker = "- " if format_element.get(qn("w:val")) == "bullet" else "1. "
    return "  " * level + marker


def _paragraph_to_markdown(
    paragraph: Paragraph,
    numbering_root,
    heading_offset: int = 0,
    is_title: bool = False,
) -> str:
    raw_text = paragraph.text.strip()
    if not raw_text:
        return ""

    text = raw_text.replace("\n", "  \n")

    if is_title:
        return f"# {raw_text}"

    style_name = paragraph.style.name if paragraph.style else ""
    if style_name.lower() == "title":
        return f"# {raw_text}"

    heading = HEADING_STYLE.match(style_name)
    if heading:
        level = min(max(int(heading.group(1)) + heading_offset, 1), 6)
        return f'{"#" * level} {text}'

    message_label = MESSAGE_LABEL.match(raw_text)
    if message_label:
        return f"**{raw_text.rstrip(':').strip()}**"

    list_marker = _list_marker(paragraph, numbering_root)
    if not list_marker:
        if style_name.lower() in {"list bullet", "seznam s odrážkami"}:
            list_marker = "- "
        elif style_name.lower() in {"list number", "číslovaný seznam"}:
            list_marker = "1. "

    field_label = FIELD_LABEL.match(raw_text)
    if field_label:
        return (
            f"{list_marker}**{_clean_field_label(field_label.group(1))}:** "
            f"{field_label.group(2)}"
        )
    return f"{list_marker}{text}"


def _escape_table_cell(text: str) -> str:
    return text.replace("|", r"\|").replace("\n", "<br>").strip()


def _clean_field_label(text: str) -> str:
    return re.sub(r"\s*\([^()]*\)", "", text).strip()


def _find_communication_title(document_blocks) -> Paragraph | None:
    for index, block in enumerate(document_blocks):
        if not isinstance(block, Paragraph):
            continue
        style_name = block.style.name if block.style else ""
        if not HEADING_STYLE.match(style_name) or "název komunikace" not in block.text.casefold():
            continue
        for candidate in document_blocks[index + 1:]:
            if isinstance(candidate, Paragraph):
                candidate_style = candidate.style.name if candidate.style else ""
                if HEADING_STYLE.match(candidate_style):
                    break
                text = candidate.text.strip()
                if not text or text.casefold().startswith("do názvu komunikace"):
                    continue
                return candidate
    return None


def _table_to_markdown(table: Table) -> str:
    rows = [
        [cell.text for cell in row.cells]
        for row in table.rows
    ]
    if not rows:
        return ""

    column_count = max(len(row) for row in rows)
    normalized_rows = [
        row + [""] * (column_count - len(row))
        for row in rows
    ]
    if column_count == 2:
        items = []
        for label, value in normalized_rows:
            if not label and not value:
                continue
            label_lines = [
                _escape_table_cell(_clean_field_label(line))
                for line in label.splitlines()
                if line.strip()
            ]
            value_lines = [
                _escape_table_cell(line)
                for line in value.splitlines()
                if line.strip()
            ]
            if label_lines and value_lines:
                value_lines = label_lines[1:] + value_lines
                item = f"- **{label_lines[0]}:**"
                if value_lines:
                    item += " " + "<br>".join(value_lines)
            elif label_lines:
                item = f"- **{label_lines[0]}:**"
                if len(label_lines) > 1:
                    item += " " + "<br>".join(label_lines[1:])
            else:
                item = f"- {'<br>'.join(value_lines)}"
            items.append(item)
        return "\n".join(items)

    normalized_rows = [
        [_escape_table_cell(cell) for cell in row]
        for row in normalized_rows
    ]
    headers = [f"Sloupec {index + 1}" for index in range(column_count)]

    markdown_rows = [
        "| " + " | ".join(headers) + " |",
        "| " + " | ".join("---" for _ in headers) + " |",
    ]
    markdown_rows.extend(
        "| " + " | ".join(row) + " |"
        for row in normalized_rows
    )
    return "\n".join(markdown_rows)


def _is_message_boundary(block) -> bool:
    if not isinstance(block, Paragraph):
        return True
    text = block.text.strip()
    style_name = block.style.name if block.style else ""
    return (
        bool(HEADING_STYLE.match(style_name))
        or bool(MESSAGE_LABEL.match(text))
        or "\n" in text
        or bool(EMAIL_VARIANT.match(text))
        or text.lower() in {"český jazyk", "anglický jazyk"}
    )


def _is_email_body_boundary(block) -> bool:
    if not isinstance(block, Paragraph):
        return True
    text = block.text.strip()
    style_name = block.style.name if block.style else ""
    return (
        bool(HEADING_STYLE.match(style_name))
        or bool(MESSAGE_LABEL.match(text))
        or bool(EMAIL_INLINE_FIELD.match(text))
        or bool(EMAIL_VARIANT.match(text))
        or text.casefold() in {
            "český jazyk",
            "anglický jazyk",
            "cz varianta",
            "en varianta",
        }
    )


def _collect_email_body(
    document_blocks,
    start_index: int,
) -> tuple[list[Paragraph], int]:
    paragraphs = []
    index = start_index
    while (
        index < len(document_blocks)
        and not _is_email_body_boundary(document_blocks[index])
    ):
        block = document_blocks[index]
        if isinstance(block, Paragraph) and block.text.strip():
            paragraphs.append(block)
        index += 1
    return paragraphs, index


def _message_content_block(paragraphs: list[Paragraph]) -> str:
    content = "\n\n".join(paragraph.text.strip() for paragraph in paragraphs)
    return f"```text\n{content}\n```"


def _inline_runs_to_html(
    paragraph: Paragraph,
    anchors: list[tuple[int, str]],
) -> str:
    runs = paragraph.runs
    if not runs or "".join(run.text for run in runs) != paragraph.text:
        runs = []

    rendered_runs = [
        (run.text, run.bold, run.italic, run.underline)
        for run in runs
        if run.text
    ] or [(paragraph.text, False, False, False)]

    grouped_positions: dict[tuple[int, int], list[int]] = {}
    quote_spans: dict[str, tuple[int, int]] = {}
    for number, quoted_text in anchors:
        span = quote_spans.get(quoted_text)
        if span is None:
            start = paragraph.text.find(quoted_text)
            if start < 0:
                continue
            span = (start, start + len(quoted_text))
            quote_spans[quoted_text] = span
        grouped_positions.setdefault(span, []).append(number)

    anchor_positions = []
    last_end = -1
    for (start, end), numbers in sorted(
        grouped_positions.items(),
        key=lambda item: (item[0][0], item[0][1]),
    ):
        if start < last_end:
            continue
        anchor_positions.extend((start, end, number) for number in numbers)
        last_end = end

    opening: dict[int, list[int]] = {}
    closing: dict[int, list[int]] = {}
    for start, end, number in anchor_positions:
        opening.setdefault(start, []).append(number)
        closing.setdefault(end, []).append(number)

    output = []
    absolute_offset = 0
    for text, bold, italic, underline in rendered_runs:
        run_end = absolute_offset + len(text)
        boundaries = sorted(
            {
                absolute_offset,
                run_end,
                *(pos for pos in opening if absolute_offset < pos < run_end),
                *(pos for pos in closing if absolute_offset < pos < run_end),
            }
        )
        for start, end in zip(boundaries, boundaries[1:]):
            output.extend(f"[{number}]" for number in opening.get(start, []))
            fragment = escape_html(
                text[start - absolute_offset:end - absolute_offset],
                quote=False,
            ).replace("\n", "<br>")
            if underline:
                fragment = f"<u>{fragment}</u>"
            if italic:
                fragment = f"<i>{fragment}</i>"
            if bold:
                fragment = f"<b>{fragment}</b>"
            output.append(fragment)
            output.extend(f"[/{number}]" for number in reversed(closing.get(end, [])))
        absolute_offset = run_end

    return "".join(output)


def _email_body_markdown(
    paragraphs: list[Paragraph],
    numbering_root,
    comments: list[WordComment],
) -> tuple[str, list[WordComment]]:
    anchor_ids = {
        element.get(qn("w:id"))
        for paragraph in paragraphs
        for element in paragraph._p.iter()
        if element.tag == qn("w:commentRangeStart")
    }
    body_comments = [
        comment for comment in comments
        if comment.comment_id in anchor_ids
    ]
    comment_numbers = {
        comment.comment_id: number
        for number, comment in enumerate(body_comments, start=1)
    }

    html_blocks = []
    list_tag = None
    list_items: list[str] = []

    def flush_list() -> None:
        nonlocal list_tag
        if list_tag is not None:
            if list_items:
                html_blocks.append(
                    f"<{list_tag}>"
                    + "".join(f"<li>{item}</li>" for item in list_items)
                    + f"</{list_tag}>"
                )
            list_items.clear()
            list_tag = None

    for paragraph in paragraphs:
        text = paragraph.text.strip()
        if not text:
            continue

        paragraph_comment_ids = {
            element.get(qn("w:id"))
            for element in paragraph._p.iter()
            if element.tag == qn("w:commentRangeStart")
        }
        anchors = [
            (comment_numbers[comment.comment_id], comment.quoted_text)
            for comment in body_comments
            if comment.comment_id in paragraph_comment_ids
        ]

        marker = _list_marker(paragraph, numbering_root)
        style_name = paragraph.style.name.lower() if paragraph.style else ""
        if not marker:
            if style_name in {"list bullet", "seznam s odrážkami"}:
                marker = "- "
            elif style_name in {"list number", "číslovaný seznam"}:
                marker = "1. "

        if marker:
            current_list_tag = "ul" if marker.strip() == "-" else "ol"
            if list_tag != current_list_tag:
                flush_list()
                list_tag = current_list_tag
            list_items.append(_inline_runs_to_html(paragraph, anchors))
        else:
            flush_list()
            html_blocks.append(
                f"<div>{_inline_runs_to_html(paragraph, anchors)}</div>"
            )

    flush_list()
    html_content = "".join(html_blocks)
    return f"```text\n{html_content}\n```", body_comments


def _email_comments_markdown(comments: list[WordComment]) -> str:
    lines = ["#### Komentáře k tělu e-mailu"]
    for number, comment in enumerate(comments, start=1):
        lines.extend(
            [
                "",
                f"**[{number}] Komentář**",
                "",
                "```text",
                comment.text,
                "```",
            ]
        )
    return "\n".join(lines)


def _push_message_markdown(
    paragraph: Paragraph,
    seen_segments: set[str],
) -> tuple[str, str] | None:
    lines = [line.strip() for line in paragraph.text.splitlines() if line.strip()]
    if len(lines) < 3:
        return None

    segment_match = re.fullmatch(r"([A-Z]{2,3}):", lines[0])
    if segment_match:
        segment = segment_match.group(1)
        lines = lines[1:]
    else:
        segment = ""
    if len(lines) < 2:
        return None

    blocks = []
    if segment and segment not in seen_segments:
        blocks.append(f"- **Segmentace:** {segment}")
    body = "\n".join(lines[1:])
    blocks.extend(
        [
            "**Titulek PUSH**",
            f"```text\n{lines[0]}\n```",
            "**Text PUSH**",
            f"```text\n{body}\n```",
        ]
    )
    return "\n\n".join(blocks), segment


def _extract_comments(document_xml: bytes, comments_xml: bytes) -> list[WordComment]:
    from xml.etree import ElementTree

    main_root = ElementTree.fromstring(document_xml)
    comments_root = ElementTree.fromstring(comments_xml)
    ns = {"w": WORD_NS}

    quoted_text: dict[str, list[str]] = {}
    active_ids: set[str] = set()
    for element in main_root.iter():
        if element.tag == f"{{{WORD_NS}}}commentRangeStart":
            comment_id = element.attrib.get(f"{{{WORD_NS}}}id")
            if comment_id is not None:
                active_ids.add(comment_id)
                quoted_text.setdefault(comment_id, [])
        elif element.tag == f"{{{WORD_NS}}}commentRangeEnd":
            comment_id = element.attrib.get(f"{{{WORD_NS}}}id")
            if comment_id is not None:
                active_ids.discard(comment_id)
        elif element.tag == f"{{{WORD_NS}}}t" and element.text:
            for comment_id in active_ids:
                quoted_text.setdefault(comment_id, []).append(element.text)

    comments = []
    for comment in comments_root.findall(".//w:comment", ns):
        comment_id = comment.attrib.get(f"{{{WORD_NS}}}id", "")
        text = "".join(node.text or "" for node in comment.findall(".//w:t", ns)).strip()
        if not text:
            continue
        comments.append(
            WordComment(
                comment_id=comment_id,
                author=comment.attrib.get(f"{{{WORD_NS}}}author", ""),
                text=text,
                quoted_text="".join(quoted_text.get(comment_id, [])).strip(),
            )
        )
    return comments


def _comments_to_markdown(comments: list[WordComment]) -> str:
    if not comments:
        return ""

    lines = ["## Komentáře ve Wordu"]
    for comment in comments:
        author = f" — {comment.author}" if comment.author else ""
        lines.extend(["", f"- **Komentář {comment.comment_id}{author}:** {comment.text}"])
        if comment.quoted_text:
            lines.append(f"  - **Označený text:** {comment.quoted_text}")
    return "\n".join(lines)


def _add_fallback_title(markdown: str, filename: str) -> str:
    title = filename.rsplit(".", 1)[0].strip() or "Dokument"
    if not markdown:
        return f"# {title}\n"
    if not any(line.startswith("# ") for line in markdown.splitlines()):
        markdown = f"# {title}\n\n{markdown}"
    return markdown.rstrip() + "\n"


def convert_txt(content: bytes, filename: str) -> tuple[str, list[str]]:
    try:
        text = content.decode("utf-8-sig").strip()
    except UnicodeDecodeError as error:
        raise ValueError("Textový soubor musí být kódovaný v UTF-8.") from error
    return _add_fallback_title(text, filename), []


def convert_pdf(content: bytes, filename: str) -> tuple[str, list[str]]:
    try:
        reader = PdfReader(BytesIO(content), strict=True)
        if reader.is_encrypted:
            raise ValueError("Šifrované PDF dokumenty nejsou podporovány.")
        text = "\n\n".join(
            page.extract_text() or ""
            for page in reader.pages
        ).strip()
    except ValueError:
        raise
    except PdfReadError as error:
        raise ValueError("Soubor není platný nebo čitelný dokument PDF.") from error

    warnings = [
        "Převod PDF zachovává pouze extrahovatelný text; rozložení, tabulky a obrázky se nemusí zachovat. OCR naskenovaných stránek se neprovádí."
    ]
    if not text:
        warnings.append("V dokumentu nebyl nalezen extrahovatelný text.")
    return _add_fallback_title(text, filename), warnings


def convert_docx(content: bytes, filename: str) -> tuple[str, list[str]]:
    try:
        document = Document(BytesIO(content))
        with ZipFile(BytesIO(content)) as archive:
            document_xml = archive.read("word/document.xml")
            comments_xml = (
                archive.read("word/comments.xml")
                if "word/comments.xml" in archive.namelist()
                else None
            )
    except (BadZipFile, KeyError, ValueError) as error:
        raise ValueError("Soubor není platný dokument DOCX.") from error

    document_blocks = list(_iter_blocks(document))
    title_paragraph = next(
        (
            block for block in document_blocks
            if isinstance(block, Paragraph) and block.text.strip()
        ),
        None,
    )
    formatted_source_title = (
        title_paragraph is not None and _is_formatted_title(title_paragraph)
    )
    communication_title = _find_communication_title(document_blocks)
    formatted_title = formatted_source_title or communication_title is not None
    skip_source_title = (
        communication_title is not None
        and title_paragraph is not None
        and (
            formatted_source_title
            or (title_paragraph.style is not None and title_paragraph.style.name.lower() == "title")
        )
    )
    comments = (
        _extract_comments(document_xml, comments_xml)
        if comments_xml is not None
        else []
    )

    blocks = []
    if communication_title is not None:
        blocks.append(f"# {communication_title.text.strip()}")
    block_index = 0
    push_heading_level = None
    push_segments = set()
    current_channel = None
    channel_heading_level = None
    used_comment_ids: set[str] = set()

    def append_email_body(paragraphs: list[Paragraph]) -> None:
        if not paragraphs:
            return
        body_markdown, body_comments = _email_body_markdown(
            paragraphs,
            document.part.numbering_part.element,
            comments,
        )
        blocks.extend(["**Tělo emailu**", body_markdown])
        if body_comments:
            blocks.append(_email_comments_markdown(body_comments))
            used_comment_ids.update(
                comment.comment_id for comment in body_comments
            )

    while block_index < len(document_blocks):
        block = document_blocks[block_index]
        if block is communication_title or (
            skip_source_title and block is title_paragraph
        ):
            block_index += 1
            continue
        if isinstance(block, Paragraph):
            style_name = block.style.name if block.style else ""
            heading = HEADING_STYLE.match(style_name)
            if heading:
                level = int(heading.group(1))
                channel_match = COMMUNICATION_HEADING.match(block.text.strip())
                if channel_match:
                    channel = channel_match.group("channel").lower()
                    if channel.startswith("push"):
                        channel = "PUSH"
                        push_heading_level = level
                        push_segments.clear()
                    elif channel == "sms":
                        channel = "SMS"
                        push_heading_level = None
                        push_segments.clear()
                    else:
                        channel = "Email"
                        push_heading_level = None
                        push_segments.clear()
                    current_channel = channel
                    channel_heading_level = level
                    blocks.append(f"## {channel}")
                    block_index += 1
                    continue
                if (
                    channel_heading_level is not None
                    and level <= channel_heading_level
                ):
                    current_channel = None
                    channel_heading_level = None
                if push_heading_level is not None and level <= push_heading_level:
                    push_heading_level = None

            raw_text = block.text.strip()
            language_variants = {
                "český jazyk": "CZ varianta",
                "anglický jazyk": "EN varianta",
                "cz varianta": "CZ varianta",
                "en varianta": "EN varianta",
            }
            variant = language_variants.get(raw_text.casefold())
            if current_channel is not None and variant is not None:
                blocks.append(f"### {variant}")
                block_index += 1
                continue

            email_variant = EMAIL_VARIANT.match(raw_text)
            if current_channel == "Email" and email_variant:
                blocks.append(f"#### {raw_text}")
                block_index += 1
                continue

            inline_email_field = EMAIL_INLINE_FIELD.match(raw_text)
            if inline_email_field:
                raw_label = inline_email_field.group("label").casefold()
                label = (
                    "Předmět emailu"
                    if raw_label.startswith("předmět")
                    else "Preheader emailu"
                )
                blocks.extend(
                    [
                        f"**{label}**",
                        f"```text\n{inline_email_field.group('value').strip()}\n```",
                    ]
                )
                if current_channel == "Email":
                    content_paragraphs, next_index = _collect_email_body(
                        document_blocks,
                        block_index + 1,
                    )
                    append_email_body(content_paragraphs)
                    block_index = next_index
                else:
                    block_index += 1
                continue

            if MESSAGE_LABEL.match(raw_text):
                if EMAIL_BODY_LABEL.match(raw_text):
                    blocks.append("**Tělo emailu**")
                    content_paragraphs, next_index = _collect_email_body(
                        document_blocks,
                        block_index + 1,
                    )
                    if content_paragraphs:
                        body_markdown, body_comments = _email_body_markdown(
                            content_paragraphs,
                            document.part.numbering_part.element,
                            comments,
                        )
                        blocks.append(body_markdown)
                        if body_comments:
                            blocks.append(_email_comments_markdown(body_comments))
                            used_comment_ids.update(
                                comment.comment_id for comment in body_comments
                            )
                    block_index = next_index
                    continue

                if (
                    EMAIL_SUBJECT_LABEL.match(raw_text)
                    or EMAIL_PREHEADER_LABEL.match(raw_text)
                ):
                    is_subject = bool(EMAIL_SUBJECT_LABEL.match(raw_text))
                    label = "Předmět emailu" if is_subject else "Preheader emailu"
                    blocks.append(f"**{label}**")
                    value_index = block_index + 1
                    if (
                        value_index < len(document_blocks)
                        and not _is_email_body_boundary(document_blocks[value_index])
                    ):
                        value = document_blocks[value_index]
                        if isinstance(value, Paragraph) and value.text.strip():
                            blocks.append(f"```text\n{value.text.strip()}\n```")
                        value_index += 1
                    if current_channel == "Email":
                        content_paragraphs, next_index = _collect_email_body(
                            document_blocks,
                            value_index,
                        )
                        append_email_body(content_paragraphs)
                        block_index = next_index
                    else:
                        block_index = value_index
                    continue

                label = _paragraph_to_markdown(
                    block,
                    document.part.numbering_part.element,
                    heading_offset=1 if formatted_title else 0,
                )
                blocks.append(label)
                content_paragraphs = []
                next_index = block_index + 1
                while (
                    next_index < len(document_blocks)
                    and not _is_message_boundary(document_blocks[next_index])
                ):
                    content_block = document_blocks[next_index]
                    if isinstance(content_block, Paragraph) and content_block.text.strip():
                        content_paragraphs.append(content_block)
                    next_index += 1

                if content_paragraphs:
                    blocks.append(_message_content_block(content_paragraphs))
                block_index = next_index
                continue

            if heading:
                markdown = _paragraph_to_markdown(
                    block,
                    document.part.numbering_part.element,
                    heading_offset=1 if formatted_title else 0,
                    is_title=(
                        formatted_title
                        and communication_title is None
                        and block is title_paragraph
                    ),
                )
            elif "\n" in block.text and push_heading_level is not None:
                push_message = _push_message_markdown(
                    block,
                    seen_segments=push_segments,
                )
                if push_message is not None:
                    markdown, segment = push_message
                    if segment:
                        push_segments.add(segment)
                else:
                    markdown = _message_content_block([block])
            elif "\n" in block.text:
                markdown = _message_content_block([block])
            else:
                markdown = _paragraph_to_markdown(
                    block,
                    document.part.numbering_part.element,
                    heading_offset=1 if formatted_title else 0,
                    is_title=(
                        formatted_title
                        and communication_title is None
                        and block is title_paragraph
                    ),
                )
        else:
            markdown = _table_to_markdown(block)
        if markdown:
            blocks.append(markdown)
        block_index += 1

    markdown = _add_fallback_title("\n\n".join(blocks), filename)

    comments_markdown = _comments_to_markdown(
        [comment for comment in comments if comment.comment_id not in used_comment_ids]
    )

    warnings = [
        "Převod zachovává text, nadpisy, odrážky a tabulky. Zvýraznění uvnitř odstavců se nepřenáší a obchodní význam polí ani názvy proměnných se neodhadují."
    ]
    if comments_xml is None:
        warnings.append("Dokument neobsahuje vložené komentáře Wordu.")

    if comments_markdown:
        markdown = markdown.rstrip() + "\n\n" + comments_markdown + "\n"
    return markdown, warnings
