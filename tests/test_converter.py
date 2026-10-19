import asyncio
from io import BytesIO
import unittest
from xml.etree import ElementTree
from zipfile import ZIP_DEFLATED, ZipFile

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from fastapi import UploadFile
from pypdf import PdfWriter

from text_to_md.app import convert as convert_upload
from text_to_md.converter import _extract_comments, convert_docx, convert_pdf, convert_txt


class ConvertEndpointTests(unittest.TestCase):
    def test_uses_fixed_markdown_filename(self) -> None:
        document = Document()
        document.add_paragraph("Dokument pro test názvu výstupu.")
        output = BytesIO()
        document.save(output)
        upload = UploadFile(
            filename="jiny-nazev-vstupu.docx",
            file=BytesIO(output.getvalue()),
        )

        response = asyncio.run(convert_upload(upload))

        self.assertEqual(response.filename, "pozadavek-na-sablonu.md")


class ConvertDocxTests(unittest.TestCase):
    def make_document(self) -> bytes:
        document = Document()
        document.add_heading("Upozornění na neprovedené platby", level=1)
        document.add_heading("E-mail", level=2)
        document.add_paragraph("Předmět: <licenceName> - Upozornění")
        document.add_paragraph("Dobrý den, <licenceName>.")
        split_token = document.add_paragraph()
        split_token.add_run("<licence")
        split_token.add_run("Name>").bold = True
        table = document.add_table(rows=2, cols=2)
        table.cell(0, 0).text = "Kanál"
        table.cell(0, 1).text = "E-mail"
        table.cell(1, 0).text = "Jazyk"
        table.cell(1, 1).text = "Čeština"
        output = BytesIO()
        document.save(output)
        return output.getvalue()

    def test_preserves_headings_text_and_tables(self) -> None:
        markdown, warnings = convert_docx(self.make_document(), "zadani.docx")

        self.assertIn("# Upozornění na neprovedené platby", markdown)
        self.assertIn("## Email", markdown)
        self.assertIn("**Předmět:** <licenceName> - Upozornění", markdown)
        self.assertIn("<licenceName>", markdown)
        self.assertIn("- **Kanál:** E-mail", markdown)
        self.assertIn("- **Jazyk:** Čeština", markdown)
        self.assertNotIn("Preheader", markdown)
        self.assertTrue(any("neodhadují" in warning for warning in warnings))

    def test_adds_fallback_title_when_word_has_no_title(self) -> None:
        document = Document()
        document.add_paragraph("Obyčejný odstavec.")
        output = BytesIO()
        document.save(output)

        markdown, _ = convert_docx(output.getvalue(), "bez-nadpisu.docx")

        self.assertTrue(markdown.startswith("# bez-nadpisu\n"))

    def test_uses_communication_name_as_title_and_cleans_table_labels(self) -> None:
        document = Document()
        title = document.add_paragraph()
        title.alignment = 1
        title_run = title.add_run("Zadávací dokumentace pro test")
        title_run.bold = True
        title_run.font.size = 16 * 12700
        document.add_heading("2.1 Název komunikace", level=2)
        document.add_paragraph(
            "Do názvu komunikace se vyplní obchodní název."
        )
        document.add_paragraph("Testovací oznámení")
        document.add_heading("2.1.1 Email", level=3)
        table = document.add_table(rows=2, cols=2)
        table.cell(0, 0).text = "Tribe (kam spadá požadavek)"
        table.cell(0, 1).text = "TEST"
        table.cell(1, 0).text = "Zdrojový systém (systém, který odesílá request)"
        table.cell(1, 1).text = "TEST_APP"
        output = BytesIO()
        document.save(output)

        markdown, _ = convert_docx(output.getvalue(), "test.docx")

        self.assertTrue(markdown.startswith("# Testovací oznámení\n"))
        self.assertNotIn("Zadávací dokumentace", markdown)
        self.assertIn("- **Tribe:** TEST", markdown)
        self.assertIn("- **Zdrojový systém:** TEST_APP", markdown)
        self.assertNotIn("kam spadá požadavek", markdown)

    def test_uses_formatted_cover_title_and_preserves_bullets(self) -> None:
        document = Document()
        title = document.add_paragraph()
        title.alignment = 1
        run = title.add_run("Název zadávací dokumentace")
        run.bold = True
        run.font.size = 16 * 12700
        document.add_heading("1 Obchodní zadání", level=1)
        bullet = document.add_paragraph("Český jazyk", style="List Paragraph")
        numbering = document.part.numbering_part.element
        abstract_ids = [
            int(item.get(qn("w:abstractNumId")))
            for item in numbering.xpath("./w:abstractNum")
        ]
        abstract_id = max(abstract_ids, default=-1) + 1
        abstract = OxmlElement("w:abstractNum")
        abstract.set(qn("w:abstractNumId"), str(abstract_id))
        level = OxmlElement("w:lvl")
        level.set(qn("w:ilvl"), "0")
        number_format = OxmlElement("w:numFmt")
        number_format.set(qn("w:val"), "bullet")
        level.append(number_format)
        level_text = OxmlElement("w:lvlText")
        level_text.set(qn("w:val"), "•")
        level.append(level_text)
        abstract.append(level)
        numbering.append(abstract)
        num_id = max((item.numId for item in numbering.num_lst), default=0) + 1
        num = OxmlElement("w:num")
        num.set(qn("w:numId"), str(num_id))
        abstract_reference = OxmlElement("w:abstractNumId")
        abstract_reference.set(qn("w:val"), str(abstract_id))
        num.append(abstract_reference)
        numbering.append(num)
        properties = bullet._p.get_or_add_pPr()
        num_properties = OxmlElement("w:numPr")
        level_reference = OxmlElement("w:ilvl")
        level_reference.set(qn("w:val"), "0")
        num_reference = OxmlElement("w:numId")
        num_reference.set(qn("w:val"), str(num_id))
        num_properties.extend([level_reference, num_reference])
        properties.append(num_properties)
        document.add_paragraph("Příklad běžného odstavce", style="List Paragraph")
        output = BytesIO()
        document.save(output)

        markdown, _ = convert_docx(output.getvalue(), "zadani.docx")

        self.assertIn("# Název zadávací dokumentace", markdown)
        self.assertIn("## 1 Obchodní zadání", markdown)
        self.assertIn("- Český jazyk", markdown)
        self.assertIn("Příklad běžného odstavce", markdown)
        self.assertNotIn("- Příklad běžného odstavce", markdown)

    def test_formats_form_fields_and_message_text(self) -> None:
        document = Document()
        document.add_paragraph("Číslo JIRA epicu: ABC-123")
        label = document.add_paragraph()
        label.add_run("Text SMS").bold = True
        document.add_paragraph("Dobrý den, <jméno>.")
        document.add_paragraph("Vaše banka")
        document.add_paragraph("Obsah emailu:")
        document.add_paragraph("Dobrý den,")
        document.add_paragraph("Toto je text e-mailu.")
        document.add_paragraph("Preheader emailu:")
        document.add_paragraph("Krátký náhled zprávy.")
        document.add_paragraph("CS:\nTitulek zprávy\nText zprávy")
        document.add_heading("Další část", level=1)
        output = BytesIO()
        document.save(output)

        markdown, _ = convert_docx(output.getvalue(), "komunikace.docx")

        self.assertIn("**Číslo JIRA epicu:** ABC-123", markdown)
        self.assertIn(
            "**Text SMS**\n\n```text\nDobrý den, <jméno>.\n\nVaše banka\n```",
            markdown,
        )
        self.assertIn(
            "**Tělo emailu**\n\n"
            "```text\n<div>Dobrý den,</div><div>Toto je text e-mailu.</div>\n```",
            markdown,
        )
        self.assertIn(
            "**Preheader emailu**\n\n```text\nKrátký náhled zprávy.\n```",
            markdown,
        )
        self.assertIn("```text\nCS:\nTitulek zprávy\nText zprávy\n```", markdown)

    def test_formats_email_fields_html_and_anchored_comments(self) -> None:
        document = Document()
        document.add_heading("2.1.2 Email", level=1)
        document.add_paragraph("Český jazyk", style="List Paragraph")
        document.add_paragraph("Předmět emailu: Předmět mailu")
        document.add_paragraph("Preheader emailu:")
        document.add_paragraph("Krátký náhled")
        body = document.add_paragraph()
        body.add_run("Tento text je ")
        body.add_run("tučně").bold = True
        body.add_run(" a obsahuje ")
        body.add_run("<contractNum>")
        output = BytesIO()
        document.save(output)

        markdown, _ = convert_docx(
            self.add_comment_to_docx(output.getvalue()),
            "email.docx",
        )

        self.assertIn("## Email\n\n### CZ varianta", markdown)
        self.assertIn(
            "**Předmět emailu**\n\n```text\nPředmět mailu\n```",
            markdown,
        )
        self.assertIn(
            "**Preheader emailu**\n\n```text\nKrátký náhled\n```",
            markdown,
        )
        self.assertIn(
            "**Tělo emailu**\n\n"
            "```text\n<div>Tento text je <b>tučně</b> a obsahuje "
            "[1]&lt;contractNum&gt;[/1]</div>\n```",
            markdown,
        )
        self.assertIn(
            "#### Komentáře k tělu e-mailu\n\n"
            "**[1] Komentář**\n\n```text\nČíslo smlouvy\n```",
            markdown,
        )

    def test_infers_email_body_after_subject_when_no_body_label_exists(self) -> None:
        document = Document()
        document.add_heading("Email", level=1)
        document.add_paragraph("CZ varianta", style="List Paragraph")
        document.add_paragraph("Předmět emailu: Testovací předmět")
        document.add_paragraph("První odstavec e-mailu.")
        document.add_paragraph("Druhý odstavec e-mailu.")
        output = BytesIO()
        document.save(output)

        markdown, _ = convert_docx(output.getvalue(), "email.docx")

        self.assertIn(
            "**Tělo emailu**\n\n"
            "```text\n<div>První odstavec e-mailu.</div>"
            "<div>Druhý odstavec e-mailu.</div>\n```",
            markdown,
        )

    def add_comment_to_docx(self, content: bytes) -> bytes:
        with ZipFile(BytesIO(content)) as source:
            files = {
                item.filename: source.read(item.filename)
                for item in source.infolist()
            }

        namespace = f"{{{_WORD_NS}}}"
        root = ElementTree.fromstring(files["word/document.xml"])
        parents = {
            child: parent
            for parent in root.iter()
            for child in parent
        }
        anchor_run = next(
            run
            for run in root.iter(f"{namespace}r")
            if "".join(
                text.text or ""
                for text in run.iter(f"{namespace}t")
            ) == "<contractNum>"
        )
        paragraph = parents[anchor_run]
        index = list(paragraph).index(anchor_run)
        start = ElementTree.Element(f"{namespace}commentRangeStart")
        start.set(f"{namespace}id", "7")
        end = ElementTree.Element(f"{namespace}commentRangeEnd")
        end.set(f"{namespace}id", "7")
        paragraph.insert(index, start)
        paragraph.insert(index + 2, end)
        reference_run = ElementTree.Element(f"{namespace}r")
        reference = ElementTree.SubElement(
            reference_run,
            f"{namespace}commentReference",
        )
        reference.set(f"{namespace}id", "7")
        paragraph.insert(index + 3, reference_run)
        files["word/document.xml"] = ElementTree.tostring(
            root,
            encoding="utf-8",
            xml_declaration=True,
        )

        comments = ElementTree.Element(f"{namespace}comments")
        comment = ElementTree.SubElement(
            comments,
            f"{namespace}comment",
            {f"{namespace}id": "7", f"{namespace}author": "Tester"},
        )
        comment_paragraph = ElementTree.SubElement(comment, f"{namespace}p")
        comment_run = ElementTree.SubElement(comment_paragraph, f"{namespace}r")
        comment_text = ElementTree.SubElement(comment_run, f"{namespace}t")
        comment_text.text = "Číslo smlouvy"
        files["word/comments.xml"] = ElementTree.tostring(
            comments,
            encoding="utf-8",
            xml_declaration=True,
        )

        result = BytesIO()
        with ZipFile(result, "w", ZIP_DEFLATED) as archive:
            for filename, data in files.items():
                archive.writestr(filename, data)
        return result.getvalue()

    def test_splits_push_title_from_message_text(self) -> None:
        document = Document()
        document.add_heading("PUSH", level=1)
        document.add_paragraph("CS:\nČeský titulek\nČeský text zprávy")
        document.add_paragraph("CS:\nEnglish title\nEnglish message text")
        output = BytesIO()
        document.save(output)

        markdown, _ = convert_docx(output.getvalue(), "push.docx")

        self.assertEqual(markdown.count("- **Segmentace:** CS"), 1)
        self.assertEqual(markdown.count("**Titulek PUSH**"), 2)
        self.assertEqual(markdown.count("**Text PUSH**"), 2)
        self.assertIn("```text\nČeský titulek\n```", markdown)
        self.assertIn("```text\nČeský text zprávy\n```", markdown)
        self.assertIn("```text\nEnglish title\n```", markdown)
        self.assertIn("```text\nEnglish message text\n```", markdown)

    def test_extracts_word_comment_and_its_anchor_text(self) -> None:
        document_xml = f"""\
<w:document xmlns:w="{_WORD_NS}">
  <w:body><w:p>
    <w:commentRangeStart w:id="7"/>
    <w:r><w:t>licenceName</w:t></w:r>
    <w:commentRangeEnd w:id="7"/>
  </w:p></w:body>
</w:document>""".encode()
        comments_xml = f"""\
<w:comments xmlns:w="{_WORD_NS}">
  <w:comment w:id="7" w:author="Tester"><w:p><w:r><w:t>Proměnná</w:t></w:r></w:p></w:comment>
</w:comments>""".encode()

        comments = _extract_comments(document_xml, comments_xml)

        self.assertEqual(len(comments), 1)
        self.assertEqual(comments[0].text, "Proměnná")
        self.assertEqual(comments[0].quoted_text, "licenceName")


class ConvertOtherFormatsTests(unittest.TestCase):
    def test_converts_utf8_text_and_adds_fallback_title(self) -> None:
        markdown, warnings = convert_txt(
            b"\xef\xbb\xbfPrvn\xc3\xad \xc5\x99\xc3\xa1dek.\nDruh\xc3\xbd \xc5\x99\xc3\xa1dek.",
            "poznamky.txt",
        )

        self.assertEqual(markdown, "# poznamky\n\nPrvní řádek.\nDruhý řádek.\n")
        self.assertEqual(warnings, [])

    def test_rejects_text_that_is_not_utf8(self) -> None:
        with self.assertRaisesRegex(ValueError, "UTF-8"):
            convert_txt(b"\xff", "poznamky.txt")

    def test_converts_pdf_without_extractable_text_with_warning(self) -> None:
        writer = PdfWriter()
        writer.add_blank_page(width=72, height=72)
        output = BytesIO()
        writer.write(output)

        markdown, warnings = convert_pdf(output.getvalue(), "sken.pdf")

        self.assertEqual(markdown, "# sken\n")
        self.assertTrue(any("OCR" in warning for warning in warnings))
        self.assertTrue(any("nebyl nalezen" in warning for warning in warnings))

    def test_rejects_invalid_pdf(self) -> None:
        with self.assertRaisesRegex(ValueError, "platný nebo čitelný"):
            convert_pdf(b"not a PDF", "vadny.pdf")


_WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


if __name__ == "__main__":
    unittest.main()
