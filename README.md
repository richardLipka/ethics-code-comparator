# Etický Kodex – Vícejazyčný Porovnávač & Synchronizátor (2024 / 2026)

Profesionální offline JavaScriptová webová aplikace pro porovnání referenčních verzí Etického kodexu (např. němčina 2024 vs. 2026), automatické sestavení stubů nových jazykových mutací a přípravu promptů pro LLM Copilot s vizuálním diffem.

**▶ Online verze:** https://richardlipka.github.io/ethics-code-comparator/ – i online verze zpracovává soubory výhradně ve vašem prohlížeči, nic se neodesílá na server.

---

## 🔒 100% Offline & Bezpečnost dat
- **Žádné externí volání:** Aplikace sama od sebe nevolá žádné LLM ani externí cloudová API.
- **Žádná telemetrie:** Veškerá data, texty kodexu a překlady zůstávají výhradně v operační paměti a lokálním úložišti vašeho prohlížeče.
- **Bezpečné pro důvěrné firemní dokumenty:** Všechny knihovny (`xlsx`, `diff`) jsou přibaleny lokálně ve složce `lib/`.

---

## 🌐 Dvojjazyčné rozhraní & Moderní lehký styl
- **Přepínání jazyků (Čeština / English):** Tlačítko v záhlaví aplikace umožňuje okamžité přepnutí celého uživatelského rozhraní mezi češtinou a angličtinou (včetně vygenerovaných instrukcí v promptech pro Copilot a exportech).
- **Lehký moderní styl (Light Theme):** Aplikace je primárně stylizována do čistého, vysoce čitelného moderního vzhledu s bílým pozadím a tmavým kontrastním textem (s možností přepnout na tmavý režim).
- **Plná šířka obrazovky (Full Width):** Rozhraní plně využívá celou šířku monitoru pro maximální komfort při čtení a porovnávání dlouhých právních textů.
- **Indikátor automatického ukládání:** V záhlaví aplikace je živý odznak informující o stavu uložení v prohlížeči vč. časového razítka.

---

## 🚀 Jak aplikaci spustit

Máte dvě rovnocenné možnosti:

### Možnost 1: Přímo v prohlížeči (Bez jakékoliv instalace)
1. Přejděte do složky projektu a **dvojklikem otevřete soubor `index.html`** v libovolném moderním prohlížeči (Chrome, Edge, Firefox, Safari).
2. Aplikace funguje okamžitě a zcela offline.

### Možnost 2: Přes lokální HTTP server
Pokud preferujete spuštění přes lokální server:
```bash
node serve.js
```
Poté otevřete v prohlížeči adresu: **`http://localhost:3000`**

Pokud je port 3000 obsazený, zadejte jiný (např. `node serve.js 8080` nebo proměnnou `PORT`). Server obsluhuje pouze soubory ze složky projektu.

---

## 📋 Hlavní funkce a workflow

### Krok 1: Nahrání souboru & Označení sloupců
- **Podpora formátů:** `.xlsx`, `.xls`, `.csv`.
- **Chytrá autodetekce:** Aplikace analyzuje hlavičky a sama předvybere referenční sloupce (např. *DE 2024* a *DE 2026*) i cílové jazyky (*CS 2024*, *EN 2024*, *FR 2024*...).
- **Živý náhled dat (Data Preview):** Přímo po načtení souboru se zobrazí tabulka s ukázkou prvních řádků načteného sešitu pro okamžité ověření správnosti.
- **Okamžité demo:** Tlačítko **"Vyzkoušet na ukázkovém kodexu"** ihned načte realistický etický kodex (15 článků, 12 jazyků).
- **Vzorový Excel ke stažení:** Tlačítko **"Stáhnout vzorový soubor (Demo XLSX)"** nebo skript `node sample/generate_sample.js`.

### Krok 2: Přehled referenčních změn (Němčina 2024 vs 2026)
- **KPI ukazatele:** Okamžitý souhrn: *Celkem odstavců*, *Beze změny*, *Upraveno*, *Nově přibylo (+)* a *Odstraněno (-)*.
- **Slovní diff:** Barevně zvýrazněné změny (zeleně přidaná slova `<ins>`, červeně přeškrtnutá odstraněná slova `<del>`).
- **Manažerský export přímo z Kroku 2:** Možnost okamžitého stažení souhrnu změn v Markdownu, Wordu (`.doc`) nebo tisku do PDF ještě před zahájením překladů.

### Krok 3: Pracovní plocha překladu (Jazyk po jazyku)
> **DŮLEŽITÉ: Plné zachování práce po jednotlivých buňkách!**  
> Všechna nová zrychlení jsou čistě aditivní. Původní precizní workflow buňku po buňce zůstává 100% zachováno.

- **4 sloupce vedle sebe v překladové laboratoři:**
  1. **Sloupec 1: Referenční německý diff (2024 → 2026)** s barevně odlišenými změnami (`<ins>` / `<del>`).
  2. **Sloupec 2: Původní schválený překlad (2024)** v daném jazyce.
  3. **Sloupec 3: Nový překlad pro rok 2026 (Editor)** s akčním panelem pro prompt do Copilotu a předvyplnění.
  4. **Sloupec 4: Živý diff vašeho překladu (2024 → 2026)** přímo vedle referenčního diffu pro okamžitou optickou kontrolu!
  *(K dispozici je i přepínač mezi 4 sloupci vedle sebe a 2×2 mřížkou).*
- **Individuální práce po buňkách (ZACHOVÁNA):**
  - Tlačítko **"Kopírovat prompt pro Copilot"** pro konkrétní buňku se všemi detaily.
  - Tlačítka rychlého předvyplnění (Původní 2024 / Referenční němčina 2026).
  - Checkbox pro označení konkrétního odstavce jako ověřený / schválený.
  - Okamžitý živý diff v reálném čase při psaní do textarea.
- **Klávesová zkratka CAT nástrojů (`Ctrl + Enter` / `Cmd + Enter`):**
  - Při psaní do editačního pole stiskněte `Ctrl + Enter` – odstavec se ihned označí jako ověřený, změny se uloží a kurzor automaticky skočí na další čekající odstavec.
- **Matice pokroku napříč jazyky (Cross-Language Matrix):**
  - Přehledné karty všech cílových jazyků s procenty dokončení a počty ověřených článků.
  - Tlačítko **"Schválit nezměněné ve všech jazycích"** pro bleskové dokončení stovek nezměněných odstavců najednou.
- **Hromadné schválení v aktivním jazyce:**
  - Tlačítko **"Schválit nezměněné"** označí všechny identické články daného jazyka jako hotové.
- **Dávkový prompt pro Copilot (Batch Prompt Modal):**
  - Vygeneruje jeden přehledný hromadný Markdown prompt se všemi upravenými a novými články najednou s jednoznačnými značkami `[Art. XX]`.
- **Chytré vložení dávky (Smart Multi-Paste Modal):**
  - Vložte celou odpověď z Copilotu – vestavěný parser automaticky rozpozná značky `[Art. XX]`, `Art. XX:`, Markdown nadpisy i JSON a roztřídí přeložené texty do příslušných buněk na jedno kliknutí.
- **Filtrování a vyhledávání:**
  - Filtry: *Vše*, *Pouze změněné a nové*, *Pouze čekající na schválení*, *Pouze schválené*.
  - Vyhledávání v textu s živým žlutým zvýrazněním nalezených výrazů.
- **Rychlá navigace:** `Alt + Šipka dolů` / `Alt + Šipka nahoru` pro skok na další/předchozí neschválený odstavec.

### Krok 4: Export výsledků & Manažerské reporty

#### 1. Manažerský souhrn změn (Managerial Change Summary)
Přehledný manažerský report s KPI tabulkou, přehledem změn a detailními tabulkami rozdílů dostupný ve 3 formátech:
- **Markdown (`.md`):** Čistý textový report ideální pro interní wikipedie, Confluence, GitHub nebo archivy.
- **Microsoft Word (`.doc`):** Nativní formátovaný Word dokument s deklarací `xmlns:w="urn:schemas-microsoft-com:office:word"`, profesionální typografií, barevnými tabulkami a nastaveným tiskovým profilem A4 na výšku.
- **Vektorové PDF:** Možnost okamžitého tisku nebo uložení do PDF z tiskového dialogu s formátováním hlavičky, patičky a tabulek.

#### 2. Přizpůsobitelná PDF šablona (PDF Template Customizer)
Aplikace obsahuje plně konfigurovatelnou šablonu pro tisk finálního kodexu v jednotlivých jazycích (tlačítko **"Upravit PDF šablonu"**):
- **Název společnosti (Company Name)**
- **Hlavní název dokumentu (Document Title)**
- **Podtitul s dynamickým zástupným symbolem `{LANG}` (Subtitle)**
- **Datum platnosti / účinnosti (Effective Date)**
- **Označení verze (Version Label)**
- **Vlastní úvodní preambule / prohlášení vedení (Custom Preamble)**
- **Firemní akcentní barva (Accent Color)** – výběr libovolné barvy pomocí color-pickeru
- **Konfigurace podpisového bloku** – možnost zapnout/vypnout a upravit funkce a jména podepisujících osob
- Všechny nastavené parametry šablony se ukládají jak do lokálního úložiště, tak do záložního `.json` projektu.
- Dokument nikdy nepoužije zastaralý text 2024 u upraveného nebo nového článku: chybějící překlad je vyznačen jako **[Překlad chybí]** a dokument, který ještě není kompletně přeložený a ověřený, nese červené označení **KONCEPT**.

#### 3. Excel sešity
- **Kompletní srovnávací sešit (2024 + 2026):**
  - Zachová veškeré původní sloupce i řádky a vedle každého sloupce 2024 vloží nový sloupec pro rok 2026.
  - Druhý list **`Stav_prekladu`** ukazuje pro každý článek a jazyk stav: *Ověřeno*, *Čeká na kontrolu*, *Chybí překlad* nebo *Zrušeno v 2026*.
  - Pokud by nový sloupec měl stejný název jako sloupec, který už v tabulce je (např. `CS 2026`), dostane příponu `(2)` – žádná data se nepřepíší.
- **Čistá verze 2026 (Finální publikace):**
  - Vygeneruje čistou tabulku pouze s platnými texty pro rok 2026, přičemž zrušené články jsou automaticky vynechány.
- Jazyky, které jste ve studiu vůbec neotevřeli, se exportují také: nezměněné články se převezmou z roku 2024, upravené a nové zůstanou prázdné.
- Po exportu aplikace upozorní, kolik odstavců ještě není ověřeno nebo přeloženo.

#### 4. Záloha projektu (.json) a automatické ukládání
- Umožňuje kdykoliv uložit rozpracovaný stav do souboru a později ho načíst zpět (včetně stavu ověření, rozpracovaných textů, šablony promptu i nastavení PDF šablony).
- Soubor obsahuje tabulku jednou a u každého jazyka jen vaši práci (nový text, stav ověření); referenční a původní texty se při načtení doplní z tabulky. Díky tomu je projekt i automatické uložení v prohlížeči několikanásobně menší. Starší soubory projektu (verze 1.0.0) se načtou beze změny.
- Při načtení se soubor kontroluje (platný JSON, tabulka, existence mapovaných sloupců) a případná chyba se srozumitelně zobrazí.
- Automatické ukládání do prohlížeče se při zavření stránky dokončí okamžitě. Pokud se uložení nepodaří (např. plné úložiště), indikátor v záhlaví zčervená a aplikace vyzve k exportu projektu.
- Soubory `.csv` se načítají jako UTF-8 (diakritika a přehlásky zůstanou zachovány). Načtením nového souboru se zahodí rozpracované překlady předchozího souboru.

---

## 🧪 Testy a ověření správnosti

Projekt obsahuje dvě komplexní testovací sady o **58 automatizovaných testech (100% zelené)**. Testy ověřují skutečné výstupy – např. GUI testy klikají na exportní tlačítka a čtou obsah vygenerovaného sešitu, staženého `.json` projektu či tiskového dokumentu:

### Spuštění všech testů přes npm:
```bash
npm test
```

Nebo jednotlivě:
- **Jednotkové & integrační testy jádra (33 testů):** `npm run test:unit` (`node test/comparator.test.js`)
  - Klasifikace řádků (`EMPTY`, `ADDED`, `REMOVED`, `UNCHANGED`, `MODIFIED`).
  - Přesnost slovního diffu v němčině i češtině (včetně diakritiky a přehlásek).
  - Bezpečnost HTML a sanitizace proti XSS.
  - Generování jednotlivých i dávkových promptů pro Copilot.
  - Inteligentní parsování dávkových odpovědí Copilotu s různými formáty značek.
  - Generování manažerského souhrnu změn pro Markdown a Word.
  - Uživatelská úprava a interpolace PDF šablony (branding, barva, preambule, podpisy).
  - XLSX binární export/import roundtrip vč. listu se stavem překladu, pořadí sloupců a ochrany proti kolizi názvů.
  - JSON záloha projektu roundtrip přes skutečný serializér (kompaktní formát, rehydratace, validace poškozených souborů).
  - Smart Paste: nespárované značky se nikam nezapíší, `1.2` ≠ `1.3`, roundtrip „hromadný prompt → odpověď → vložení“.
  - Formátovaný tiskový PDF export (bez zastaralých textů, označení konceptu, escapování polí šablony).
  - Úplnost lokalizace (všechny použité klíče existují v CS i EN).
- **Baterie GUI testů v reálném DOM (25 testů):** `npm run test:gui` (`node test/gui-battery.test.js`)
  - Ověřuje existenci skutečných dat a ovládacích prvků v DOMu:
  - Inicializace GUI a přepínání jazyků CZ/EN.
  - Detekce sloupců a konfigurovatelnost mapování.
  - Tabulka živého náhledu dat (Data Preview Table) v Kroku 1.
  - Vykreslení německých diffů s reálnými `<ins>` a `<del>` značkami v Kroku 2.
  - 4 sloupce překladatelského studia a přepínání jazykových záložek.
  - Živý přepočet diffu cílového jazyka v reálném čase při psaní do textarea.
  - Tlačítka pro předvyplnění původního textu a německé reference.
  - Matice pokroku napříč jazyky a hromadné schválení nezměněných.
  - Klávesová zkratka `Ctrl + Enter` (CAT accelerator) s automatickým skokem na další položku.
  - Dávkový prompt pro Copilot a modální okno pro chytré vložení dávky (Smart Multi-Paste).
  - Export manažerského souhrnu změn v Markdownu, Wordu i PDF.
  - Modální okno úpravy PDF šablony a promítnutí změn do generovaného HTML.
  - Automatické ukládání, obnova po znovuotevření a viditelné hlášení selhání ukládání.
  - Průběh ověřování, posun progress baru a záloha/obnova z JSON přes tlačítko a výběr souboru.
  - Excel export tlačítky v Kroku 4 (obsah obou sešitů), nahrání skutečného `.xlsx` a UTF-8 `.csv`.
  - Ochrana navigace před neplatným mapováním, zavírání modálních oken klávesou Escape.

### Diagnostika přímo v prohlížeči:
- V záhlaví aplikace klikněte na tlačítko **"🧪 Diagnostika a testy"**.
- Spustí se interní testovací modul přímo v běhovém prostředí vašeho prohlížeče a zobrazí zelené výsledky.

---

## 📁 Struktura souborů v projektu

```
Comparator/
├── index.html                   # Hlavní aplikační rozhraní (Kroky 1-4, modální okna)
├── serve.js                     # Lokální HTTP server (node serve.js)
├── package.json                 # Konfigurace npm projektu a testovacích skriptů
├── README.md                    # Podrobná dokumentace aplikace
├── LICENSE                      # Licence MIT
├── THIRD_PARTY_NOTICES.md       # Licence přibalených knihoven v lib/
├── css/
│   └── app.css                  # Styly, sticky toolbar, matice, modály, lehký/tmavý režim
├── js/
│   ├── app.js                   # Řídicí logika uživatelského rozhraní, klávesové zkratky a workflow
│   └── core/
│       ├── i18n.js              # Internacionalizace (Čeština a English slovníky)
│       ├── comparator-engine.js # Jádro diffů, dávkových promptů, parseru a manažerských reportů
│       ├── sample-generator.js  # Generátor 15 ukázkových článků ve 12 jazycích
│       ├── storage.js           # Ukládání do localStorage, JSON import/export vč. PDF šablony
│       └── pdf-export.js        # Vektorový tiskový PDF export, Word export a manažerské sestavy
├── lib/
│   ├── xlsx.full.min.js         # Lokální offline knihovna SheetJS pro Excel
│   └── diff.min.js              # Lokální offline knihovna jsdiff pro slovní porovnání
├── sample/
│   ├── generate_sample.js       # Node skript pro vytvoření vzorového Excelu
│   └── vzorec_eticky_kodex_2024_2026.xlsx # Fyzický vzorový sešit
└── test/
    ├── comparator.test.js       # Jednotková a integrační testovací sada jádra (33 testů)
    └── gui-battery.test.js      # Baterie testů grafického uživatelského rozhraní v DOM (25 testů)
```

---

## 📄 Licence

Aplikace je šířena pod licencí [MIT](LICENSE) © 2026 Richard Lipka.

Přibalené knihovny ve složce `lib/` mají vlastní licence (viz [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md)): SheetJS Community Edition (Apache 2.0, viz `lib/LICENSE-sheetjs-xlsx.txt`) a jsdiff (BSD 3-Clause, viz `lib/LICENSE-jsdiff.txt`).
