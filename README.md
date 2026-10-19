# Text to MD

Samostatná služba pro převod textových dokumentů do Markdownu.

## Podporovaný vstup

Přijímá dokumenty Microsoft Word `.docx`, textové soubory `.txt` kódované
v UTF-8 a textová PDF `.pdf`. Převod DOCX zachovává text, nadpisy, tabulky
a vložené komentáře Wordu. Formátovaný titulní řádek a odrážkové i číslované
seznamy se převádějí do odpovídající struktury Markdownu. Dvou-sloupcové
formulářové tabulky se převedou na odrážky s tučně označenými názvy polí.
Jako hlavní nadpis se použije název komunikace, pokud je v dokumentu uveden;
vysvětlující texty v závorkách se z názvů formulářových polí odstraní.
V e-mailu se předmět a preheader převedou jako textové bloky a tělo jako HTML
v bloku `text`; zachovává se tučné, kurzíva, podtržení, seznamy a ukotvené
komentáře Wordu. Preheader se převede pouze tehdy, když je uveden ve vstupu.
Text/titulek SMS či PUSH se také uzavře do bloku `text`. TXT se převede jako
prostý text. Z PDF se získá text; rozložení a tabulky se nemusí zachovat. OCR
naskenovaných PDF se neprovádí. Zvýraznění mimo tělo e-mailu se nepřenáší;
popisky oddělené dvojtečkou se však zobrazí tučně.

Převodník neodhaduje obchodní význam textů ani názvy proměnných. Výsledný
Markdown a upozornění vrací HTML klientovi k náhledu a dalšímu zpracování.
Dokument se během požadavku zpracuje v paměti a neukládá se na disk.

## Spuštění

```powershell
python -m pip install -r requirements.txt
python -m uvicorn text_to_md.app:app --reload
```

API dokumentace je po spuštění dostupná na `http://127.0.0.1:8000/docs`.

## API

`POST /convert` přijímá multipart pole `file` obsahující `.docx`, `.txt` nebo
`.pdf` a vrací JSON:

```json
{
  "filename": "zadani.md",
  "markdown": "# ...",
  "warnings": ["..."]
}
```

Limit nahrávaného souboru je 20 MB. `GET /health` slouží ke kontrole dostupnosti.

## Testy

```powershell
python -m unittest discover -s tests -v
```
