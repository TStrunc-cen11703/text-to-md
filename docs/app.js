import { loadPyodide } from "https://cdn.jsdelivr.net/pyodide/v314.0.7/full/pyodide.mjs";

const MAX_UPLOAD_BYTES = 20 * 1024 * 1024;
const PYODIDE_BASE = "https://cdn.jsdelivr.net/pyodide/v314.0.7/full/";
const INPUT_PATH = "/tmp/text-to-md-input";

const form = document.querySelector("#convert-form");
const fileInput = document.querySelector("#document");
const convertButton = document.querySelector("#convert-button");
const status = document.querySelector("#status");
const warnings = document.querySelector("#warnings");
const result = document.querySelector("#result");
const outputName = document.querySelector("#output-name");
const markdownOutput = document.querySelector("#markdown-output");
const downloadButton = document.querySelector("#download-button");
const copyButton = document.querySelector("#copy-button");

let pyodidePromise;

function setStatus(message, state = "") {
  status.textContent = message;
  status.dataset.state = state;
}

async function loadConverter() {
  if (!pyodidePromise) {
    pyodidePromise = (async () => {
      setStatus("Připravuji převodní prostředí; při prvním spuštění to může chvíli trvat…");
      const pyodide = await loadPyodide({ indexURL: PYODIDE_BASE });
      await pyodide.loadPackage(["lxml", "micropip"]);

      const micropip = pyodide.pyimport("micropip");
      try {
        await micropip.install(["python-docx>=1.1,<2", "pypdf>=5,<7"]);
      } finally {
        micropip.destroy();
      }

      const sourceResponse = await fetch("./converter.py");
      if (!sourceResponse.ok) {
        throw new Error("Nepodařilo se načíst převodní logiku.");
      }
      pyodide.FS.writeFile(
        "/home/pyodide/converter.py",
        await sourceResponse.text(),
      );
      pyodide.runPython("import sys; sys.path.insert(0, '/home/pyodide')");
      return pyodide;
    })().catch((error) => {
      pyodidePromise = undefined;
      throw error;
    });
  }
  return pyodidePromise;
}

async function convertFile(pyodide, file) {
  pyodide.FS.writeFile(INPUT_PATH, new Uint8Array(await file.arrayBuffer()));
  pyodide.globals.set("uploaded_filename", file.name);

  try {
    const response = await pyodide.runPythonAsync(`
import json
from pathlib import PurePath
from converter import OUTPUT_FILENAME, convert_docx, convert_pdf, convert_txt

filename = uploaded_filename.replace("\\\\", "/").rsplit("/", 1)[-1]
extension = PurePath(filename).suffix.lower()
converters = {
    ".docx": convert_docx,
    ".pdf": convert_pdf,
    ".txt": convert_txt,
}
converter = converters.get(extension)
if converter is None:
    raise ValueError("Podporované formáty jsou .docx, .pdf a .txt.")

with open("${INPUT_PATH}", "rb") as source:
    markdown, conversion_warnings = converter(source.read(), filename)

json.dumps({
    "filename": OUTPUT_FILENAME,
    "markdown": markdown,
    "warnings": conversion_warnings,
}, ensure_ascii=False)
`);
    return JSON.parse(response);
  } finally {
    pyodide.FS.unlink(INPUT_PATH);
    pyodide.globals.delete("uploaded_filename");
  }
}

form.addEventListener("submit", async (event) => {
  event.preventDefault();
  const file = fileInput.files[0];
  if (!file) {
    setStatus("Nejprve vyberte soubor.", "error");
    return;
  }
  if (file.size > MAX_UPLOAD_BYTES) {
    setStatus("Soubor překračuje limit 20 MB.", "error");
    return;
  }

  convertButton.disabled = true;
  result.hidden = true;
  warnings.replaceChildren();

  try {
    const pyodide = await loadConverter();
    setStatus("Probíhá převod…");
    const converted = await convertFile(pyodide, file);
    outputName.textContent = converted.filename;
    markdownOutput.value = converted.markdown;
    for (const warning of converted.warnings) {
      const item = document.createElement("li");
      item.textContent = warning;
      warnings.append(item);
    }
    result.hidden = false;
    setStatus("Převod je hotový.", "success");
  } catch (error) {
    const message = error instanceof Error
      ? error.message
      : "Převod se nezdařil. Zkuste to prosím znovu.";
    setStatus(message, "error");
  } finally {
    convertButton.disabled = false;
  }
});

downloadButton.addEventListener("click", () => {
  const blob = new Blob([markdownOutput.value], {
    type: "text/markdown;charset=utf-8",
  });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = outputName.textContent;
  link.hidden = true;
  document.body.append(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
});

copyButton.addEventListener("click", async () => {
  try {
    await navigator.clipboard.writeText(markdownOutput.value);
    setStatus("Markdown byl zkopírován do schránky.", "success");
  } catch {
    setStatus("Kopírování se nepodařilo. Označte text v náhledu a zkopírujte jej ručně.", "error");
  }
});
