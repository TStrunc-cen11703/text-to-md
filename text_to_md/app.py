from pathlib import Path, PurePath
from zipfile import BadZipFile

from docx.opc.exceptions import PackageNotFoundError
from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.responses import FileResponse
from lxml.etree import XMLSyntaxError
from pydantic import BaseModel

from text_to_md.converter import convert_docx, convert_pdf, convert_txt


MAX_UPLOAD_BYTES = 20 * 1024 * 1024

app = FastAPI(
    title="Text to Markdown",
    description="Converts uploaded text documents to structure-preserving Markdown.",
    version="0.1.0",
)


class ConversionResponse(BaseModel):
    filename: str
    markdown: str
    warnings: list[str]


@app.get("/", include_in_schema=False)
def home() -> FileResponse:
    return FileResponse(
        Path(__file__).with_name("index.html"),
        media_type="text/html",
    )


@app.get("/health")
def health() -> dict[str, str]:
    return {"status": "ok"}


@app.post("/convert", response_model=ConversionResponse)
async def convert(file: UploadFile = File(...)) -> ConversionResponse:
    uploaded_name = (file.filename or "").replace("\\", "/").rsplit("/", 1)[-1]
    extension = PurePath(uploaded_name).suffix.lower()
    converters = {
        ".docx": convert_docx,
        ".pdf": convert_pdf,
        ".txt": convert_txt,
    }
    converter = converters.get(extension)
    if converter is None:
        raise HTTPException(
            status_code=415,
            detail="Podporované formáty jsou .docx, .pdf a .txt.",
        )

    content = await file.read(MAX_UPLOAD_BYTES + 1)
    if len(content) > MAX_UPLOAD_BYTES:
        raise HTTPException(status_code=413, detail="Soubor překračuje limit 20 MB.")

    try:
        markdown, warnings = converter(content, uploaded_name)
    except (BadZipFile, PackageNotFoundError, XMLSyntaxError, ValueError) as error:
        raise HTTPException(
            status_code=422,
            detail=f"Soubor se nepodařilo přečíst jako platný dokument {extension}.",
        ) from error

    output_name = f"{PurePath(uploaded_name).stem or 'dokument'}.md"
    return ConversionResponse(
        filename=output_name,
        markdown=markdown,
        warnings=warnings,
    )
