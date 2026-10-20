# Text to MD

Samostatná služba pro převod textových dokumentů do Markdownu.

## Podporovaný vstup

Přijímá dokumenty Microsoft Word `.docx`, textové soubory `.txt` kódované
v UTF-8 a textová PDF `.pdf`. Převod DOCX zachovává text, nadpisy, tabulky
a vložené komentáře Wordu. Formátovaný titulní řádek a odrážkové i číslované
seznamy se převádějí do odpovídající struktury Markdownu. Dvou-sloupcové
formulářové tabulky se převedou na odrážky s tučně označenými názvy polí.
Číslování kapitol na začátku nadpisů, například `1 Management summary`, se
ve výstupu odstraní. Nadpis druhé úrovně `Management summary` se převede na
`Základní informace`.
Sekce Email, SMS nebo PUSH se vynechá, pokud dokument uvádí, že příslušný
kanál není požadován, například `SMS není požadována`.
Sekce `Obchodní zadání` se vynechá, pokud její obsah uvádí, že zadání není
potřebné.
Sekce `Instrukce` v DOCX se vynechá včetně podsekcí a jejího textu až do
následujícího nadpisu stejné nebo vyšší úrovně.
Vyřadí se také pole `Instrukce:` s vysvětlujícím textem.
Kontakty Business garanta a garanta akceptace se v tabulkách i textových řádcích
převedou do samostatných sekcí; pokud jsou stejné, druhá sekce uvede „Stejný
jako Business garant“. Telefonní čísla se převedou bez mezer.
Popisek `Ukládání komunikace do TRACE` se zkrátí na `Ukládání do TRACE`.
U tohoto pole se zachová pouze samostatná hodnota `ANO` nebo `NE`; položky
`Trace_Type`, `Trace_Sub_Type` a `State` se vynechají.
Pole s KCP se převede na `Konzultace s KCP` a hodnotu `ANO` nebo `NE`;
prázdná hodnota se doplní jako `NE`.
Obecné instrukční texty šablony, například „Tato podkapitola se vyplní…“ nebo
„Zde se doplní Emailová textace…“, se vynechají; skutečný obsah a komentáře
k proměnným zůstávají zachovány.
Samostatná poznámka `* nepovinné` se z výstupu vynechá.
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
Markdown a upozornění se zobrazí v lokálním HTML rozhraní k náhledu a dalšímu
zpracování. Dokument se zpracuje lokálně a aplikace jej po ukončení požadavku
neuchovává.
Emotikony ve výstupním Markdownu se vždy zapisují jako znaky Unicode.

## Použití přímo na GitHubu

Po úspěšném nasazení přes GitHub Actions bude aplikace dostupná na
`https://tstrunc-cen11703.github.io/text-to-md/`. Repozitář i stránka jsou
veřejné. Vybraný dokument zpracovává Python převodník spuštěný v prohlížeči;
obsah dokumentu se neodesílá na GitHub ani na jiné servery. Při prvním použití
se z CDN a PyPI stáhne běhové prostředí a potřebné knihovny.
Výsledný soubor se vždy jmenuje `pozadavek-na-sablonu.md`.
Soubor lze do vyznačeného pole přetáhnout nebo jej vybrat tlačítkem.

## Spuštění

```powershell
python -m pip install -r requirements.txt
python -m uvicorn text_to_md.app:app --reload
```

GitHub v prohlížeči soubory `.bat` nespouští. Pokud otevíráte projekt na
GitHubu, zvolte **Code → Download ZIP**, archiv rozbalte a v Průzkumníku
Windows otevřete rozbalenou složku. Teprve tam dvojklikněte na
`spustit-prevodnik.bat`.

Při prvním spuštění skript vytvoří virtuální prostředí a nainstaluje
závislosti, takže je potřeba připojení k internetu. Poté otevře místní
převodník v prohlížeči. Okno serveru nechte otevřené; zavřením okna aplikaci
zastavíte.

Na stránce vyberte soubor a klikněte na **Převést na Markdown**. Hotový
soubor stáhnete tlačítkem **Stáhnout soubor .md**. Převod probíhá na vašem
počítači; dokument se neposílá na GitHub ani jiné externí služby.

API dokumentace je po spuštění dostupná na `http://127.0.0.1:8000/docs`.

## API

`POST /convert` přijímá multipart pole `file` obsahující `.docx`, `.txt` nebo
`.pdf` a vrací JSON:

```json
{
  "filename": "pozadavek-na-sablonu.md",
  "markdown": "# ...",
  "warnings": ["..."]
}
```

Název výstupního souboru je vždy `pozadavek-na-sablonu.md`, bez ohledu na název
vstupního souboru. Limit nahrávaného souboru je 20 MB. `GET /health` slouží ke
kontrole dostupnosti.

## Testy

```powershell
python -m unittest discover -s tests -v
```
