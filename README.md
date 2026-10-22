# Srovnávač verzí & Překladový asistent

Offline webová aplikace pro porovnání **staré a nové verze referenčního textu** (smlouvy, směrnice, kodexy, manuály – libovolný dokument rozdělený na odstavce v řádcích tabulky) a pro aktualizaci **všech jeho jazykových verzí**. Ukáže, které odstavce se změnily, pomůže s překladem (prompty pro Copilot, hromadné vložení odpovědi) a vyexportuje výsledky do Excelu, dokumentu pro tisk a shrnutí změn.

**▶ Online verze:** https://richardlipka.github.io/ethics-code-comparator/ – i online verze zpracovává soubory výhradně ve vašem prohlížeči, nic se neodesílá na server.

---

## 🔒 Offline a bezpečnost dat
- Aplikace nevolá žádné LLM ani jiné externí služby a nesbírá telemetrii.
- Texty zůstávají v paměti a v lokálním úložišti vašeho prohlížeče; projekt lze uložit do souboru `.json`.
- Knihovny (`xlsx`, `diff`) jsou přibaleny ve složce `lib/`.

## 🚀 Spuštění
- **Bez instalace:** otevřete `index.html` v prohlížeči (Chrome, Edge, Firefox, Safari).
- **Přes lokální server:** `node serve.js` a otevřete `http://localhost:3000`. Jiný port: `node serve.js 8080` (nebo proměnná `PORT`). Server obsluhuje jen soubory ze složky projektu.

---

## 📋 Postup práce

### Krok 1: Tabulka a výběr sloupců
- Podporované formáty `.xlsx`, `.xls`, `.csv` (CSV se čte jako UTF-8). U sešitu s více listy vyberete list.
- **Bílé znaky na začátku a konci buněk se odstraní** (dlouhé buňky začínající prázdnými řádky tak v náhledu nevypadají prázdné).
- Volba **„První řádek obsahuje názvy sloupců“** – bez záhlaví jsou sloupce pojmenovány A, B, C… a první řádek je považován za data.
- Vyberete dva **referenční sloupce**: starou a novou verzi textu (v libovolném jazyce).
- Pro každý další sloupec určíte, zda obsahuje **starou verzi některého jazyka** (použít), nebo se **ignoruje**. Aplikace jazyky podle názvů nehádá; předem zaškrtne jen sloupce, které obsahují souvislý text.
- U jazyka zadáte jeho **název** (výchozí je název sloupce bez čísla verze, např. „CS 2024“ → „CS“, u sloupce bez záhlaví „Sloupec C“).
- Pokud tabulka už obsahuje **sloupec s novou verzí jazyka** (třeba téměř prázdný), vyberete ho; jeho vyplněné buňky se převezmou jako rozpracovaný text ke kontrole. Jinak se při exportu vytvoří nový sloupec (název lze upravit). Pokud se referenční sloupce liší číslem verze (např. `EN 2024` / `EN 2026`) a v tabulce je odpovídající sloupec jazyka (`CS 2026`), spáruje se automaticky.
- Odstavce nemají nadpisy – identifikuje je **číslo řádku v Excelu**.

### Krok 2: Změny reference
- Statistika (beze změny, upraveno, nově přidáno, zrušeno) a tabulka referenčního textu se zvýrazněnými změnami.
- Shrnutí změn reference ke stažení (Markdown, Word, tisk/PDF).

### Krok 3: Aktualizace překladů
- Pro každý odstavec: rozdíl v referenci, stará verze jazyka, editor nové verze a živý rozdíl stará → nová verze jazyka. Nad texty je jen číslo řádku a typ změny – text buňky se neopakuje.
- Nezměněné odstavce se převezmou automaticky; zrušené se v nové verzi vynechají.
- Prompt pro Copilot pro jeden odstavec, hromadný prompt pro všechny změny a **Smart Paste** – odpověď se rozdělí podle značek řádků (`[#12]`, `Řádek 12:`…); nespárované značky se zobrazí a nikam se nezapíší.
- Schvalování po odstavcích (i `Ctrl + Enter`), filtry, hledání, přehled postupu všech jazyků, hromadné schválení nezměněných.

### Krok 4: Export
Nahoře vyberete **jazyk** (referenci nebo některý z jazyků) – použije se pro dokument k tisku, shrnutí změn a částečný sešit.
- **Kompletní sešit:** všechny původní sloupce i řádky; nová verze každého jazyka v jeho existujícím sloupci nebo v novém sloupci hned za starou verzí (existující data se nikdy nepřepíší). List `Stav_prekladu` ukazuje pro každý řádek a jazyk *Ověřeno / Čeká na kontrolu / Chybí překlad / Zrušeno*.
- **Pouze nová verze:** odstavce platné v nové verzi, nová verze reference a všech jazyků, bez starých verzí.
- **Částečný sešit pro jeden jazyk:** číslo řádku, zvolené referenční sloupce (výchozí stará a nová reference, lze vybrat jiné, např. angličtinu), stará a nová verze vybraného jazyka a stav – např. pro předání překladateli.
- **Dokument pro tisk / PDF:** nová verze vybraného jazyka jako formátovaný dokument A4. Chybějící překlad je vyznačen a neúplný dokument je označen jako **KONCEPT** – nikdy se nepoužije zastaralý text. Hlavička (organizace, název, podtitul s `{LANG}`, verze, datum, úvod, podpisy, barva) je volitelná.
- **Shrnutí změn vybraného jazyka:** statistika, upravené, nové a zrušené odstavce (u jazyka jeho stará a nová verze a stav kontroly) – Markdown, Word nebo tisk/PDF.
- Po exportu aplikace upozorní, kolik odstavců ještě není schváleno nebo přeloženo.

### Uložení projektu
- Tlačítka **💾 Uložit projekt** a **📂 Načíst projekt** jsou v záhlaví stránky ve všech krocích.
- Soubor `.json` obsahuje tabulku, výběr sloupců, názvy jazyků, rozpracované texty, stav kontroly a nastavení šablon. U každého jazyka se ukládá jen vaše práce; ostatní texty se při načtení doplní z tabulky, takže soubor je malý. Soubory starších verzí aplikace se načtou také.
- Práce se průběžně ukládá i do prohlížeče (při zavření stránky okamžitě). Když se uložení nepodaří (např. plné úložiště), indikátor v záhlaví zčervená a aplikace vyzve k uložení projektu.

---

## 🧪 Testy

```bash
npm test
```

- **Jádro (27 testů):** `npm run test:unit` – klasifikace a diff, pracovní data jazyků (včetně existujících sloupců nové verze), prompty, Smart Paste, všechny tři exporty do Excelu (zapsané do `.xlsx` a načtené zpět), soubor projektu, shrnutí změn pro referenci i jazyk, dokument pro tisk, úplnost a obecnost lokalizace.
- **GUI v jsdom (28 testů):** `npm run test:gui` – testy klikají na skutečné ovládací prvky a čtou výsledné soubory: výběr sloupců, nahrání `.xlsx`/`.csv` (bez záhlaví, s mezerami v buňkách, s existujícím sloupcem nové verze), studio, uložení a načtení projektu tlačítky v záhlaví, autosave, exporty a tisk.

V aplikaci je také tlačítko 🧪 s interní diagnostikou přímo v prohlížeči.

---

## 📁 Struktura

```
├── index.html                 # Rozhraní (kroky 1–4, modální okna)
├── serve.js                   # Lokální HTTP server
├── css/app.css                # Styly (světlý/tmavý režim, responzivní rozvržení)
├── js/app.js                  # Řízení rozhraní a workflow
├── js/core/
│   ├── comparator-engine.js   # Diff, klasifikace, pracovní data, prompty, Smart Paste, exporty, shrnutí změn
│   ├── storage.js             # Autosave, soubor projektu (.json)
│   ├── pdf-export.js          # Dokument pro tisk, shrnutí změn (Word/Markdown/tisk)
│   ├── i18n.js                # Čeština a angličtina
│   └── sample-generator.js    # Ukázková data
├── lib/                       # SheetJS a jsdiff (vlastní licence)
├── sample/                    # Ukázkový sešit a skript pro jeho vytvoření
├── test/                      # Testy jádra a GUI
├── LICENSE                    # Licence MIT
└── THIRD_PARTY_NOTICES.md     # Licence přibalených knihoven
```

## 📄 Licence

Aplikace je šířena pod licencí [MIT](LICENSE) © 2026 Richard Lipka.

Přibalené knihovny ve složce `lib/` mají vlastní licence (viz [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md)): SheetJS Community Edition (Apache 2.0) a jsdiff (BSD 3-Clause).
