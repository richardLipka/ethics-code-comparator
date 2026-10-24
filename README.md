# Internationalisation Manager

Offline webová aplikace pro správu jazykových verzí: porovnání **staré a nové verze referenčního textu** (smlouvy, směrnice, kodexy, manuály – libovolný dokument rozdělený na odstavce v řádcích tabulky) a pro aktualizaci **všech jeho jazykových verzí**. Ukáže, které odstavce se změnily, pomůže s překladem (prompty pro Copilot, hromadné vložení odpovědi) a vyexportuje výsledky do Excelu, dokumentu pro tisk a shrnutí změn.

**▶ Online verze:** https://richardlipka.github.io/internationalisation-manager/ – i online verze zpracovává soubory výhradně ve vašem prohlížeči, nic se neodesílá na server.

---

## 🔒 Offline a bezpečnost dat
- Aplikace nevolá žádné LLM ani jiné externí služby a nesbírá telemetrii.
- Texty zůstávají v paměti a v lokálním úložišti vašeho prohlížeče; projekt lze uložit do souboru `.json`.
- Knihovny (`xlsx`, `diff`) jsou přibaleny ve složce `lib/`.

## 🚀 Spuštění
- **Bez instalace:** otevřete `index.html` v prohlížeči (Chrome, Edge, Firefox, Safari).
- **Přes lokální server:** `node serve.js` a otevřete `http://localhost:3000`. Jiný port: `node serve.js 8080` (nebo proměnná `PORT`). Server obsluhuje jen soubory ze složky projektu (bez skrytých souborů jako `.git`) a je dostupný jen z tohoto počítače; pro přístup ze sítě spusťte `HOST=0.0.0.0 node serve.js`.

---

## 📋 Postup práce

### Krok 1: Tabulka a výběr sloupců
- Výchozí importér čte `.xlsx`, `.xlsm`, `.xls`, `.ods` a `.csv`. U sešitu s více listy vyberete list. Další formáty lze doplnit importním adaptérem (viz níže).
- **CSV a `.txt`** se načtou v UTF-8 i ve starší kódové stránce Windows-1250 (např. „CSV (oddělený středníkem)“ z českého Excelu); oddělovač (čárka, středník, tabulátor) se pozná automaticky. Buňky CSV zůstanou textem – `0012` se nezmění na `12` a datum na číslo.
- **Buňky s datem** v sešitu se načtou jako datum `RRRR-MM-DD` (s časem, pokud ho formát buňky zobrazuje), ne jako pořadové číslo dne.
- **Word:** `.docx`, `.doc` (Word 97-2003 i dokument uložený Wordem jako RTF nebo HTML) a `.rtf` – každý odstavec (i buňka tabulky) je jeden řádek; z polí se převezme jen zobrazený výsledek.
- **Bílé znaky na začátku a konci buněk se odstraní** (dlouhé buňky začínající prázdnými řádky tak v náhledu nevypadají prázdné).
- Volba **„První řádek obsahuje názvy sloupců“** – bez záhlaví jsou sloupce pojmenovány A, B, C… a první řádek je považován za data.
- Vyberete dva **referenční sloupce**: starou a novou verzi textu (v libovolném jazyce).
- Pro každý další sloupec určíte, zda obsahuje **starou verzi některého jazyka** (použít), nebo se **ignoruje**. Aplikace jazyky podle názvů nehádá; předem zaškrtne jen sloupce, které obsahují souvislý text.
- U jazyka zadáte jeho **název** (výchozí je název sloupce bez čísla verze, např. „CS 2024“ → „CS“, u sloupce bez záhlaví „Sloupec C“).
- Pokud tabulka už obsahuje **sloupec s novou verzí jazyka** (třeba téměř prázdný), vyberete ho; jeho vyplněné buňky se převezmou jako rozpracovaný text ke kontrole. Jinak se při exportu vytvoří nový sloupec (název lze upravit). Pokud se referenční sloupce liší číslem verze (např. `EN 2024` / `EN 2026`) a v tabulce je odpovídající sloupec jazyka (`CS 2026`), spáruje se automaticky.
- Odstavce nemají nadpisy – identifikuje je **číslo řádku v Excelu**.
- V **náhledu řádků** lze zaškrtnutím „Přeskočit“ vyřadit řádky, které se nemají zpracovat (nadpisy, poznámky…); úplně prázdné řádky se přeskočí automaticky. Přeskočené řádky se nepřekládají a nejsou v nové verzi, dokumentu ani shrnutí, ale kompletní sešit je zachová beze změny na původním místě. Výběr se ukládá do projektu a změna výběru nepřesune rozpracované texty k jiným řádkům.

### Krok 2: Změny reference
- Statistika (beze změny, upraveno, nově přidáno, zrušeno) a tabulka referenčního textu se zvýrazněnými změnami.
- Shrnutí změn reference ke stažení (Markdown, Word, tisk/PDF).

### Krok 3: Aktualizace překladů
- Pro každý odstavec: rozdíl v referenci, stará verze jazyka, editor nové verze a živý rozdíl stará → nová verze jazyka. Nad texty je jen číslo řádku a typ změny – text buňky se neopakuje.
- **Pomocný referenční jazyk (Auxiliary Language Context):** volitelný výběr dalšího jazyka v panelu studia (např. anglická nová verze). Zobrazí se jako 5. sloupec karty vedle referenčního diffu s tlačítkem „Použít jako základ“ (vloží text přímo do editoru) a je automaticky zahrnut i do individuálních a hromadných promptů pro Copilot.
- **Ověření překladu (Quality Audit Prompt):** druhé tlačítko v editoru karty („🔍 Ověřit překlad“), které do schránky zkopíruje prompt pro jazykový model s požadavkem na posouzení věcné a významové shodnosti nového překladu vůči referenci a výpis případných rozdílů, nedostatků či návrhu úprav.
- Nezměněné odstavce se převezmou automaticky; zrušené se v nové verzi vynechají.
- Prompt pro Copilot pro jeden odstavec, hromadný prompt pro všechny změny a **Smart Paste** – odpověď se rozdělí podle značek řádků (`[#12]`, `Řádek 12:`…); nespárované značky se zobrazí a nikam se nezapíší.
- Schvalování po odstavcích (i `Ctrl + Enter`), filtry, hledání, přehled postupu všech jazyků, hromadné schválení nezměněných.

### Krok 4: Export
Nahoře vyberete **jazyk** (referenci nebo některý z jazyků) – použije se pro dokument k tisku, shrnutí změn a částečný sešit.
- **Kompletní sešit:** všechny původní sloupce i řádky; nová verze každého jazyka v jeho existujícím sloupci nebo v novém sloupci hned za starou verzí (existující data se nikdy nepřepíší). List `Stav_prekladu` ukazuje pro každý řádek a jazyk *Ověřeno / Čeká na kontrolu / Chybí překlad / Zrušeno*.
- **Aktualizovat původní sešit (zachová formát)** – u tabulky načtené z `.xlsx`/`.xlsm`: nové verze se zapíšou do **kopie původního souboru**, takže zůstane formátování, vzorce, sloučené buňky, šířky sloupců, ukotvení, ostatní listy, grafy i makra. Mění se jen zapisované buňky jednoho listu: existující sloupec nové verze se přepíše na místě, jinak se nový sloupec přidá **za poslední použitý sloupec** (vložení doprostřed by vyžadovalo přepsat všechny vzorce a odkazy) a převezme vzhled a šířku staré verze jazyka. Stav kontroly se zapíše na list `Stav_prekladu` (při opakované aktualizaci se nahradí). Po obnovení stránky nebo načtení projektu aplikace požádá o původní soubor a ověří, že k projektu patří.
- **Pouze nová verze:** odstavce platné v nové verzi, nová verze reference a všech jazyků, bez starých verzí.
- **Částečný sešit pro jeden jazyk:** číslo řádku, zvolené další sloupce, stará a nová verze vybraného jazyka a stav – např. pro předání překladateli. **Vyplněný sešit lze načíst zpět** tlačítkem „📂 Importovat překlady ze sešitu“ (Krok 3 nebo 4): texty se spárují podle sloupce `Řádek`/`Row` a sloupce nové verze jazyka, aplikace před převzetím ukáže, kolik textů se změní (a které řádky v projektu nejsou), a převzaté texty označí jako *Čeká na kontrolu*. Prázdné buňky a zrušené odstavce se nemění. Další sloupce se vybírají ve dvou skupinách: **nové verze** (nová reference a aktuální stav překladu ostatních jazyků, i když jejich sloupec v tabulce zatím neexistuje) a **původní sloupce** tabulky. Nové verze jsou v exportu na prvním místě; výchozí výběr je stará a nová reference.
- **Word (.docx):** u dokumentu načteného z Wordu se překlad zapíše do **kopie původního dokumentu** – každý odstavec si ponechá styl, číslování a formát prvního úseku textu; tabulky, obrázky, textová pole, záhlaví, zápatí a konce stránek zůstanou. Nepřeložené odstavce zůstanou v původním znění. U jiného zdroje se vytvoří nový dokument Wordu s hlavičkou z nastavení dokumentu.
- **Dokument pro tisk / PDF:** nová verze vybraného jazyka jako formátovaný dokument A4. Chybějící překlad je vyznačen a neúplný dokument je označen jako **KONCEPT** – nikdy se nepoužije zastaralý text. Hlavička (organizace, název, podtitul s `{LANG}`, verze, datum, úvod, podpisy, barva) je volitelná.
- **Shrnutí změn vybraného jazyka:** statistika, upravené, nové a zrušené odstavce (u jazyka jeho stará a nová verze a stav kontroly) – Markdown, Word nebo tisk/PDF.
- Po exportu aplikace upozorní, kolik odstavců ještě není schváleno nebo přeloženo.

### Formáty importu a exportu (adaptéry)
- Načítání souborů a všechny exporty jdou přes **adaptéry** – aplikace ani jádro nevolají knihovny konkrétních formátů přímo. Nový formát se přidá jedním souborem ve složce `js/io/` a jeho uvedením v `index.html`.
- Výchozí adaptéry: import Excel/CSV, Word a text; export tabulek do Excelu, zápis do původního sešitu, dokument k tisku/HTML a Word (.docx), shrnutí změn jako Markdown, Word a tisk.
- Pokud je zaregistrováno víc tabulkových exportérů, Krok 4 nabídne volbu **formátu tabulek** (výchozí zůstává Excel).
- Popis rozhraní a příklady: [docs/ADAPTERS.md](docs/ADAPTERS.md).

### Uložení projektu
- Tlačítka **💾 Uložit projekt** a **📂 Načíst projekt** jsou v záhlaví stránky ve všech krocích.
- Soubor `.json` obsahuje tabulku, výběr sloupců, názvy jazyků, rozpracované texty, stav kontroly a nastavení šablon. U každého jazyka se ukládá jen vaše práce; ostatní texty se při načtení doplní z tabulky, takže soubor je malý. Soubory starších verzí aplikace se načtou také.
- Načítaný soubor se nejprve zkontroluje. Pokud nejde o projekt překladu (jiný JSON, soubor jiné aplikace, omylem vybraná tabulka) nebo je poškozený, aplikace vysvětlí proč a **aktuální práce zůstane beze změny**. Drobné nesrovnalosti (např. soubor z novější verze, texty jazyka, který není vybraný, nesedící počet řádků) projekt nezablokují – zobrazí se jako upozornění.
- Před nahrazením rozpracované práce jiným projektem se aplikace zeptá. Soubor projektu lze otevřít i přetažením do oblasti pro nahrání v Kroku 1.
- Práce se průběžně ukládá i do prohlížeče (při zavření stránky okamžitě). Když se uložení nepodaří (např. plné úložiště), indikátor v záhlaví zčervená a aplikace vyzve k uložení projektu.

---

## 🧪 Testy

```bash
npm test
```

- **Jádro a adaptéry (61 testů):** `npm run test:unit` – klasifikace a diff, pracovní data jazyků (včetně existujících sloupců nové verze), prompty (překlad i ověření), Smart Paste, exporty tabulek (zapsané do `.xlsx` a načtené zpět), soubor projektu, shrnutí změn a dokument pro tisk, registry adaptérů (výchozí formáty, registrace nových), Excel importér (CSV ve Windows-1250, data), Word importér na skutečných souborech Wordu (`test/fixtures`), import vyplněného částečného sešitu, zápis do kopie původního sešitu a dokumentu Wordu (ověřeno na souborech vytvořených Excelem a Wordem), úplnost a obecnost lokalizace.
- **GUI v jsdom (53 testů):** `npm run test:gui` – testy klikají na skutečné ovládací prvky a čtou stažené soubory: výběr sloupců, nahrání `.xlsx`/`.csv` (bez záhlaví, s mezerami v buňkách, s existujícím sloupcem nové verze), studio, pomocný referenční jazyk, ověření překladu, uložení a načtení projektu, autosave (i pod původním názvem), exporty a tisk, nový importér a exportér zaregistrovaný za běhu, import překladů ze sešitu, přístupnost (audit [axe-core](https://github.com/dequelabs/axe-core) všech kroků a dialogů, ovládání klávesnicí, fokus v dialozích).
- **E2E (5 testů):** `npm run test:e2e` – celé průchody Kroky 1–4 pro vícejazyčný Excel, Word, text a jednosloupcový Excel.
- **Kontrola kódu:** `npm run lint` (ESLint) a `npm run typecheck` (typová kontrola jádra a adaptérů podle JSDoc přes TypeScript).
- Kontrola kódu i testy běží automaticky na GitHubu při každém push a pull requestu (`.github/workflows/test.yml`).

V aplikaci je také tlačítko 🧪 s interní diagnostikou přímo v prohlížeči.

---

## 📁 Struktura

```
├── index.html                 # Rozhraní (kroky 1–4, modální okna)
├── serve.js                   # Lokální HTTP server
├── css/app.css                # Styly (světlý/tmavý režim, responzivní rozvržení; barvy splňují kontrast WCAG AA)
├── js/app/                    # Řízení rozhraní a workflow (sdílený jmenný prostor App, pořadí načtení v index.html)
│   ├── core.js                # Stav, odkazy na prvky stránky, překlady, notifikace, autosave, společné funkce
│   ├── shell.js               # Jazyk rozhraní, motiv, navigace mezi kroky, dialogy
│   ├── project.js             # Soubor projektu, obnovení autosave, nový projekt
│   ├── step1-import.js        # Krok 1: načtení souboru, list, náhled a přeskočené řádky
│   ├── step1-columns.js       # Krok 1: výběr sloupců a jazyků
│   ├── step2-reference.js     # Krok 2: změny reference
│   ├── step3-studio.js        # Krok 3: postup, jazyky, filtry, navigace
│   ├── step3-cards.js         # Krok 3: karty odstavců, editor, prompty
│   ├── step3-batch.js         # Krok 3: hromadný prompt, Smart Paste, import překladů ze sešitu
│   ├── step4-export.js        # Krok 4: exporty a jejich nastavení
│   ├── settings.js            # Šablona promptu, diagnostika
│   └── main.js                # Spuštění aplikace
├── js/core/
│   ├── comparator-engine.js   # Diff, klasifikace, pracovní data, prompty, Smart Paste, modely exportů
│   ├── storage.js             # Autosave, soubor projektu (.json)
│   ├── i18n.js                # Čeština a angličtina
│   └── sample-generator.js    # Ukázková data
├── js/io/                     # Adaptéry importu a exportu
│   ├── importers.js           # Registr importérů a model tabulky
│   ├── excel-importer.js      # Výchozí importér: Excel / OpenDocument / CSV
│   ├── word-importer.js       # Word: .docx, .doc, .rtf
│   ├── ooxml.js               # Balíčky Office (.docx, .xlsx): zip, odstavce Wordu
│   ├── text-importer.js       # Text: .txt
│   ├── exporters.js           # Registr exportérů, stahování a tisk
│   ├── excel-exporter.js      # Výchozí export tabulek: Excel (.xlsx)
│   ├── html-document-exporter.js  # Dokument pro tisk / HTML
│   ├── docx-exporter.js       # Dokument Word (.docx), i do kopie původního dokumentu
│   ├── xlsx-update-exporter.js    # Zápis do kopie původního sešitu (zachová formát)
│   └── summary-exporters.js   # Shrnutí změn: Markdown, Word, tisk
├── docs/ADAPTERS.md           # Jak napsat vlastní importér / exportér
├── lib/                       # SheetJS a jsdiff (vlastní licence)
├── sample/                    # Ukázkový sešit a skript pro jeho vytvoření
├── test/                      # Testy jádra, GUI a E2E; test/fixtures: skutečné soubory Wordu a Excelu
├── eslint.config.js, jsconfig.json  # Kontrola kódu (ESLint, typová kontrola)
├── LICENSE                    # Licence MIT
└── THIRD_PARTY_NOTICES.md     # Licence přibalených knihoven
```

## 📄 Licence

Aplikace je šířena pod licencí [MIT](LICENSE) © 2026 Richard Lipka.

Přibalené knihovny ve složce `lib/` mají vlastní licence (viz [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md)): SheetJS Community Edition (Apache 2.0) a jsdiff (BSD 3-Clause).
