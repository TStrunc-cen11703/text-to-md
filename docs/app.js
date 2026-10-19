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


      const sourceResponse = await fetch("./converter.py", { cache: "no-store" });
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
