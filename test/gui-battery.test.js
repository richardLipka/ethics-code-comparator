/**
 * Baterie testů grafického uživatelského rozhraní (GUI Battery Tests)
 * 
 * Ověřuje existenci skutečných dat v DOM, nikoliv pouze návratové hodnoty funkcí:
 * 1. Inicializaci GUI, přepínání jazyků v reálném DOM (CZ/EN)
 * 2. Konfigurovatelnost sloupců v GUI (výběr referenčních sloupců, cílových jazyků, vyloučení z checkboxů)
 * 3. Přehledné zobrazení diffů v GUI (reálné <ins> a <del> značky, německý text, procenta změn)
 * 4. Překladatelské studio: přepínání záložek jazyků, 4-sloupcový layout, živý přepočet diffu cílového jazyka při psaní do textarea
 * 5. Předvyplňovací tlačítka v GUI pro upravené a přidané odstavce
 * 6. Kvalitně generované prompty pro Copilot (stručné, bez zbytečné vaty, s kompletním kontextem a pravidly)
 * 7. Vlastní šablony promptů v GUI
 * 8. Průběh ověřování a progress bary v reálném DOM
 * 9. Ukládání a načítání projektu (.json) s kompletní obnovou dat v DOM
 * 10. Formátovaný PDF export pro jednotlivé jazyky i hromadný export
 * 11. Komplexní a čistý Excel export
 */

const fs = require('fs');
const path = require('path');
const assert = require('assert');
const { JSDOM } = require('jsdom');

let passedTests = 0;
let failedTests = 0;
const tests = [];

// Tests are collected and run sequentially at the end, so they may be async
// (file reading, debounced autosave, delayed restore).
function runTest(name, fn) {
  tests.push({ name, fn });
}

async function runAllTests() {
  for (const test of tests) {
    try {
      await test.fn();
      console.log(`  ✓ ${test.name}`);
      passedTests++;
    } catch (err) {
      console.error(`  ✗ ${test.name}`);
      console.error(`    CHYBA: ${err.message}`);
      if (err.stack) {
        const lines = err.stack.split('\n').slice(1, 4).join('\n');
        console.error(`    ${lines}`);
      }
      failedTests++;
    }
  }
}

const plain = (value) => JSON.parse(JSON.stringify(value));
const sleep = (ms) => new Promise(resolve => setTimeout(resolve, ms));

function readBlobText(window, blob) {
  return new Promise((resolve, reject) => {
    const reader = new window.FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = () => reject(reader.error);
    reader.readAsText(blob);
  });
}

function toastTexts(document) {
  return Array.from(document.querySelectorAll('#toastContainer .toast')).map(t => t.textContent);
}

// Loads the sample, confirms the mapping and opens the CS 2024 tab of the studio
function openCzechStudio(window) {
  const document = window.document;
  document.getElementById('loadSampleBtn').click();
  document.getElementById('confirmMappingBtn').click();
  document.getElementById('goToTranslationBtn').click();
  const csTab = Array.from(document.querySelectorAll('#langTabsWrapper .lang-tab')).find(t => t.textContent.includes('CS 2024'));
  csTab.click();
}

function typeInto(window, textarea, text) {
  textarea.value = text;
  textarea.dispatchEvent(new window.Event('input'));
}

console.log('=== Spouštím baterii GUI testů pro Comparator ===\n');

// 1. Příprava JSDOM prostředí
const htmlPath = path.join(__dirname, '..', 'index.html');
const htmlSource = fs.readFileSync(htmlPath, 'utf8');

const diffJs = fs.readFileSync(path.join(__dirname, '..', 'lib', 'diff.min.js'), 'utf8');
const xlsxJs = fs.readFileSync(path.join(__dirname, '..', 'lib', 'xlsx.full.min.js'), 'utf8');
const i18nJs = fs.readFileSync(path.join(__dirname, '..', 'js', 'core', 'i18n.js'), 'utf8');
const engineJs = fs.readFileSync(path.join(__dirname, '..', 'js', 'core', 'comparator-engine.js'), 'utf8');
const sampleJs = fs.readFileSync(path.join(__dirname, '..', 'js', 'core', 'sample-generator.js'), 'utf8');
const storageJs = fs.readFileSync(path.join(__dirname, '..', 'js', 'core', 'storage.js'), 'utf8');
const pdfJs = fs.readFileSync(path.join(__dirname, '..', 'js', 'core', 'pdf-export.js'), 'utf8');
const appJs = fs.readFileSync(path.join(__dirname, '..', 'js', 'app.js'), 'utf8');

function createDomEnvironment(options) {
  const opts = options || {};
  const dom = new JSDOM(htmlSource, {
    url: 'http://localhost:3000/',
    runScripts: 'outside-only',
    pretendToBeVisual: true
  });

  const { window } = dom;

  // Mock clipboard and window APIs
  let lastCopiedText = '';
  window.navigator.clipboard = {
    writeText: async (text) => {
      lastCopiedText = text;
      return Promise.resolve();
    }
  };

  window.getLastCopiedText = () => lastCopiedText;
  window.alert = (msg) => { window.__lastAlert = msg; };
  // The "resume autosaved project?" prompt is declined unless a test asks for it
  window.confirm = (msg) => {
    window.__lastConfirm = msg;
    return /rozpracovaný projekt|saved project/.test(msg) ? !!opts.acceptResume : true;
  };
  window.scrollTo = () => {};
  window.Element.prototype.scrollIntoView = () => {};

  // Capture file downloads (Blob + <a download>) instead of navigating
  const blobs = {};
  let blobCounter = 0;
  window.__downloads = [];
  window.URL.createObjectURL = (blob) => {
    const url = 'blob:mock/' + (++blobCounter);
    blobs[url] = blob;
    return url;
  };
  window.URL.revokeObjectURL = () => {};
  window.HTMLAnchorElement.prototype.click = function () {
    window.__downloads.push({ name: this.download, blob: blobs[this.href] });
  };

  // Capture print windows
  window.__printWindows = [];
  window.open = () => {
    let html = '';
    const fake = {
      document: {
        readyState: 'complete',
        open() { html = ''; },
        write(s) { html += s; },
        close() {}
      },
      addEventListener() {},
      focus() {},
      print() { fake.printed = true; },
      get html() { return html; }
    };
    window.__printWindows.push(fake);
    return fake;
  };

  if (opts.beforeInit) opts.beforeInit(window);

  // Spuštění skriptů v pořadí z index.html
  window.eval(diffJs);
  window.eval(xlsxJs);
  window.eval(i18nJs);
  window.eval(engineJs);
  window.eval(sampleJs);
  window.eval(storageJs);
  window.eval(pdfJs);
  // Capture generated Excel workbooks (XLSX.writeFile would write to disk)
  window.__workbooks = [];
  window.XLSX.writeFile = (wb, name) => { window.__workbooks.push({ wb, name }); };

  window.eval(appJs);

  // Inicializace aplikace (DOMContentLoaded)
  window.document.dispatchEvent(new window.Event('DOMContentLoaded'));

  return { dom, window, document: window.document };
}

// BATERIE TESTŮ

// TEST 1: Inicializace GUI a přepínání jazyků
runTest('GUI Boot: Správné výchozí zobrazení a přepínání jazyků CZ/EN v DOM', () => {
  const { document } = createDomEnvironment();

  // Výchozí jazyk je čeština
  assert.strictEqual(document.documentElement.getAttribute('lang'), 'cs', 'Výchozí lang atribut musí být cs');
  
  const stepNav1 = document.getElementById('stepNav1');
  assert.ok(stepNav1.classList.contains('active'), 'Navigace kroku 1 je aktivní');

  const langBtn = document.getElementById('langToggleBtn');
  assert.ok(langBtn, 'Tlačítko přepnutí jazyka existuje');
  assert.strictEqual(langBtn.textContent.trim(), '🇬🇧 English', 'Tlačítko v češtině nabízí přepnutí do EN');

  // Kliknutí na přepínač jazyka -> EN
  langBtn.click();
  assert.strictEqual(document.documentElement.getAttribute('lang'), 'en', 'Lang atribut se musí změnit na en');
  assert.strictEqual(langBtn.textContent.trim(), '🇨🇿 Čeština', 'Tlačítko v EN nabízí přepnutí do CS');
  
  // Kontrola reálného textu v DOM po překladu
  const dropzoneTitle = document.querySelector('[data-i18n="dropzoneTitle"]');
  assert.strictEqual(dropzoneTitle.textContent.trim(), 'Drag & drop Excel file (.xlsx, .xls) here or click to browse', 'Text dropzone musí být v angličtině');

  // Přepnutí zpět do češtiny
  langBtn.click();
  assert.strictEqual(document.documentElement.getAttribute('lang'), 'cs', 'Lang atribut se musí vrátit na cs');
  assert.strictEqual(dropzoneTitle.textContent.trim(), 'Přetáhněte sem Excel soubor (.xlsx, .xls) nebo klikněte pro výběr', 'Text dropzone musí být v češtině');
});

// TEST 2: Načtení ukázkových dat a zobrazení reálných sloupců
runTest('Column Detection: Načtení dat zobrazí v GUI reálné sloupce a správné výchozí mapování', () => {
  const { document } = createDomEnvironment();

  const loadSampleBtn = document.getElementById('loadSampleBtn');
  loadSampleBtn.click();

  const mappingContainer = document.getElementById('mappingContainer');
  assert.strictEqual(mappingContainer.style.display, 'block', 'Mapovací kontejner se musí zobrazit');

  const refOldSelect = document.getElementById('refOldSelect');
  const refNewSelect = document.getElementById('refNewSelect');

  assert.strictEqual(refOldSelect.value, 'DE 2024 (Referenz)', 'refOldSelect musí mít vybranou DE 2024 (Referenz)');
  assert.strictEqual(refNewSelect.value, 'DE 2026 (Referenz)', 'refNewSelect musí mít vybranou DE 2026 (Referenz)');

  // Ověření všech 14 sloupců v možnostech selectu
  const options = Array.from(refOldSelect.options).map(o => o.value).filter(v => v);
  assert.strictEqual(options.length, 14, 'Select musí obsahovat všech 14 sloupců ze souboru');
  assert.ok(options.includes('CS 2024'), 'Obsahuje CS 2024');
  assert.ok(options.includes('EN 2024'), 'Obsahuje EN 2024');
  assert.ok(options.includes('FR 2024'), 'Obsahuje FR 2024');

  // Ověření cílových jazyků v checkboxy kontejneru (12 sloupců, mimo DE 2024 a DE 2026)
  const checkboxes = document.querySelectorAll('#targetColsContainer input[type="checkbox"]');
  assert.strictEqual(checkboxes.length, 12, 'Musí existovat 12 cílových jazykových checkboxů');

  const checkedValues = Array.from(checkboxes).filter(cb => cb.checked).map(cb => cb.value);
  assert.ok(checkedValues.includes('CS 2024'), 'Čeština je vybrána jako cílový jazyk');
  assert.ok(checkedValues.includes('EN 2024'), 'Angličtina je vybrána jako cílový jazyk');
});

// TEST 3: Konfigurovatelnost sloupců v GUI (změna referenčních sloupců a aktualizace checkboxů)
runTest('Column Reconfigurability: Změna referenčních sloupců dynamicky aktualizuje dostupné cílové jazyky', () => {
  const { window, document } = createDomEnvironment();

  document.getElementById('loadSampleBtn').click();

  const refOldSelect = document.getElementById('refOldSelect');
  const refNewSelect = document.getElementById('refNewSelect');

  // Uživatel změní refOldSelect na EN 2024
  refOldSelect.value = 'EN 2024';
  if (refOldSelect.onchange) refOldSelect.onchange();
  refOldSelect.dispatchEvent(new window.Event('change'));

  // EN 2024 nesmí být mezi cílovými checkboxy
  let checkboxes = Array.from(document.querySelectorAll('#targetColsContainer input[type="checkbox"]'));
  let cbValues = checkboxes.map(cb => cb.value);
  assert.strictEqual(cbValues.includes('EN 2024'), false, 'EN 2024 nesmí být mezi cílovými jazyky, pokud je referencí');
  assert.strictEqual(cbValues.includes('DE 2024 (Referenz)'), true, 'DE 2024 (Referenz) musí být nyní dostupná jako cílový jazyk');

  // Vrátíme na standardní německou referenci
  refOldSelect.value = 'DE 2024 (Referenz)';
  if (refOldSelect.onchange) refOldSelect.onchange();
  refOldSelect.dispatchEvent(new window.Event('change'));

  refNewSelect.value = 'DE 2026 (Referenz)';
  if (refNewSelect.onchange) refNewSelect.onchange();
  refNewSelect.dispatchEvent(new window.Event('change'));

  checkboxes = Array.from(document.querySelectorAll('#targetColsContainer input[type="checkbox"]'));
  cbValues = checkboxes.map(cb => cb.value);
  assert.strictEqual(cbValues.includes('DE 2024 (Referenz)'), false, 'DE 2024 nesmí být mezi cílovými checkboxy');
  assert.strictEqual(cbValues.includes('DE 2026 (Referenz)'), false, 'DE 2026 nesmí být mezi cílovými checkboxy');
  assert.strictEqual(cbValues.includes('CS 2024'), true, 'CS 2024 je mezi cílovými checkboxy');

  // Vybereme pouze CS 2024 a FR 2024
  checkboxes.forEach(cb => {
    cb.checked = (cb.value === 'CS 2024' || cb.value === 'FR 2024');
    cb.dispatchEvent(new window.Event('change'));
  });

  // Potvrdit mapování
  document.getElementById('confirmMappingBtn').click();

  const step2 = document.getElementById('step2Section');
  assert.strictEqual(step2.style.display, 'block', 'Krok 2 musí být po potvrzení aktivní');
});

// TEST 4: Krok 2 - Přehledné zobrazení německých diffů se skutečnými daty v GUI
runTest('Diff Display Step 2: Tabulka změn v Kroku 2 obsahuje reálná data, <ins> a <del> značky', () => {
  const { document } = createDomEnvironment();

  document.getElementById('loadSampleBtn').click();
  document.getElementById('confirmMappingBtn').click();

  // Ověření celkového počtu řádků v KPI kartách
  assert.strictEqual(document.getElementById('kpiTotalRows').textContent, '15', 'KPI celkem řádků musí být 15');
  
  const modifiedKpi = parseInt(document.getElementById('kpiModified').textContent, 10);
  assert.ok(modifiedKpi > 0, 'Počet upravených řádků musí být větší než 0');

  // Kontrola řádků tabulky v DOM
  const rows = document.querySelectorAll('#refDiffTableBody tr');
  assert.strictEqual(rows.length, 15, 'V tabulce musí být vygenerováno přesně 15 řádků');

  // Řádek 1 (Art. 01)
  const row1 = rows[0];
  const row1Label = row1.querySelector('td:nth-child(1)').textContent;
  assert.ok(row1Label.includes('Art. 01'), 'První řádek obsahuje identifikátor Art. 01');
  assert.ok(row1Label.includes('Präambel'), 'První řádek obsahuje název sekce Präambel');

  // Řádek 3 (Art. 03 - Chancengleichheit) obsahuje změny (Geschlechtsidentität, Alter, ausnahmslos vs Rasse)
  const row3 = rows[2];
  const diffBox3 = row3.querySelector('.diff-box');
  assert.ok(diffBox3, 'Diff box musí existovat');
  assert.ok(diffBox3.innerHTML.includes('<ins class="diff-ins"'), 'Obsahuje <ins class="diff-ins">');
  assert.ok(diffBox3.innerHTML.includes('Geschlechtsidentität'), 'Zeleně zvýrazněno slovo Geschlechtsidentität');
  assert.ok(diffBox3.innerHTML.includes('<del class="diff-del"'), 'Obsahuje <del class="diff-del">');
  assert.ok(diffBox3.innerHTML.includes('Rasse'), 'Červeně škrtnuto slovo Rasse');

  // Řádek 4 (Art. 04 - Dary 35 vs 50 Euro)
  const row4 = rows[3];
  const diffBox4 = row4.querySelector('.diff-box');
  assert.ok(diffBox4.innerHTML.includes('<ins class="diff-ins"'), 'Obsahuje ins tag');
  assert.ok(diffBox4.innerHTML.includes('50') && diffBox4.innerHTML.includes('diff-ins'), 'Zeleně zvýrazněno 50');
  assert.ok(diffBox4.innerHTML.includes('35') && diffBox4.innerHTML.includes('diff-del'), 'Červeně škrtnuto 35');

  // Řádek 8 (Art. 08 - AI Etika) byl přidán (2024 prázdné, 2026 text)
  const row8 = rows[7];
  const badge8 = row8.querySelector('.badge-added');
  assert.ok(badge8, 'Řádek 8 má odznak PŘIDÁNO');
  assert.ok(row8.innerHTML.includes('Künstliche Intelligenz'), 'Řádek 8 obsahuje text o umělé inteligenci');

  // Řádek 10 (Art. 10 - Pandemie) byl smazán (2024 text, 2026 prázdné)
  const row10 = rows[9];
  const badge10 = row10.querySelector('.badge-removed');
  assert.ok(badge10, 'Řádek 10 má odznak SMAZÁNO');
  assert.ok(row10.innerHTML.includes('<del class="diff-del"'), 'Smazaný řádek je přeškrtnut');
});

// TEST 5: Krok 3 - Překladatelské studio a 4-sloupcový layout s reálnými daty
runTest('Translation Studio Step 3: Zobrazení jazykových záložek, přepnutí na češtinu a 4 sloupce s reálným textem', () => {
  const { document } = createDomEnvironment();

  document.getElementById('loadSampleBtn').click();
  document.getElementById('confirmMappingBtn').click();
  document.getElementById('goToTranslationBtn').click();

  const step3 = document.getElementById('step3Section');
  assert.strictEqual(step3.style.display, 'block', 'Krok 3 musí být aktivní');

  // Ověření záložek jazyků v DOM
  const tabs = Array.from(document.querySelectorAll('#langTabsWrapper .lang-tab'));
  assert.ok(tabs.length >= 2, 'Záložky jazyků musí být vykresleny');
  
  // Klikneme na záložku CS 2024 (Čeština)
  const csTab = tabs.find(t => t.textContent.includes('CS 2024'));
  assert.ok(csTab, 'Záložka pro CS 2024 existuje');
  csTab.click();

  const activeTab = document.querySelector('#langTabsWrapper .lang-tab.active');
  assert.ok(activeTab && activeTab.textContent.includes('CS 2024'), 'Záložka CS 2024 je aktivní v DOM');

  // Ověření karet článků v DOM
  const cards = document.querySelectorAll('#articlesContainer .article-card');
  assert.strictEqual(cards.length, 15, 'Musí být vykresleno 15 karet článků');

  // Kontrola 4 sloupců pro Art. 04 (Dary)
  const card4 = document.getElementById('article-card-3');
  const subPanels = card4.querySelectorAll('.sub-panel');
  assert.strictEqual(subPanels.length, 4, 'Karta musí obsahovat přesně 4 sloupce');

  // Sloupec 1: Německý diff
  assert.ok(subPanels[0].innerHTML.includes('diff-ins'), 'Sloupec 1 obsahuje zvýrazněný německý diff');
  assert.ok(subPanels[0].innerHTML.includes('50'), 'Sloupec 1 obsahuje 50');

  // Sloupec 2: Původní ověřený český překlad
  assert.ok(subPanels[1].textContent.includes('Dary a pozvání'), 'Sloupec 2 obsahuje původní český překlad 2024');
  assert.ok(subPanels[1].textContent.includes('35 eur'), 'Sloupec 2 obsahuje 35 eur');

  // Sloupec 3: Editor překladu 2026 s textarea
  const textarea = subPanels[2].querySelector('textarea.target-textarea');
  assert.ok(textarea, 'Sloupec 3 obsahuje textarea editor');

  // Sloupec 4: Živý diff cílového jazyka
  const liveDiffBox = subPanels[3].querySelector('.live-target-diff');
  assert.ok(liveDiffBox, 'Sloupec 4 obsahuje box živého diffu');
});

// TEST 6: Živý přepočet diffu cílového jazyka při editaci v reálném DOM
runTest('Live Target Diff: Psaní do textarea okamžitě v reálném čase aktualizuje <ins> a <del> v DOM', () => {
  const { window, document } = createDomEnvironment();

  document.getElementById('loadSampleBtn').click();
  document.getElementById('confirmMappingBtn').click();
  document.getElementById('goToTranslationBtn').click();

  // Přepneme na záložku CS 2024
  const tabs = Array.from(document.querySelectorAll('#langTabsWrapper .lang-tab'));
  const csTab = tabs.find(t => t.textContent.includes('CS 2024'));
  csTab.click();

  // Vybereme kartu 4 (Art. 04 - Dary do 35 eur změněno na 50 eur)
  const card4 = document.getElementById('article-card-3');
  assert.ok(card4, 'Karta pro článek 4 existuje');

  const textarea = card4.querySelector('textarea.target-textarea');
  const liveDiffBox = card4.querySelector('.live-target-diff');

  // Původní český text má "35 eur"
  const oldText = card4.querySelectorAll('.sub-panel')[1].textContent;
  assert.ok(oldText.includes('35 eur'), 'Původní text obsahuje 35 eur');

  // Uživatel zadá upravený text s 50 eur
  const newCzechText = 'Dary a pozvání lze přijmout pouze tehdy, pokud jejich hodnota nepřesahuje 50 eur a byly předem schváleny přes portál compliance.';
  textarea.value = newCzechText;
  
  // Vyvoláme událost input (psaní uživatele)
  textarea.dispatchEvent(new window.Event('input'));

  // Okamžitá kontrola DOMu ve sloupci 4
  assert.ok(liveDiffBox.innerHTML.includes('<ins class="diff-ins"'), 'Sloupec 4 musí obsahovat zelený <ins> tag');
  assert.ok(liveDiffBox.innerHTML.includes('50') && liveDiffBox.innerHTML.includes('diff-ins'), 'Zeleně je zvýrazněno nově napsané 50');
  assert.ok(liveDiffBox.innerHTML.includes('<del class="diff-del"'), 'Sloupec 4 musí obsahovat červený <del> tag');
  assert.ok(liveDiffBox.innerHTML.includes('35') && liveDiffBox.innerHTML.includes('diff-del'), 'Červeně je škrtnuto původní 35');
});

// TEST 7: Předvyplňovací tlačítka v GUI (Kopírovat 2024 / Kopírovat DE)
runTest('Prefill Buttons: Tlačítka pro předvyplnění správně vloží text do textarea a vyvolají diff', () => {
  const { document } = createDomEnvironment();

  document.getElementById('loadSampleBtn').click();
  document.getElementById('confirmMappingBtn').click();
  document.getElementById('goToTranslationBtn').click();

  // Přepneme na záložku CS 2024
  const tabs = Array.from(document.querySelectorAll('#langTabsWrapper .lang-tab'));
  const csTab = tabs.find(t => t.textContent.includes('CS 2024'));
  csTab.click();

  // Článek 4 (Art. 04) je MODIFIED, má tlačítko .prefill-old-btn
  const card4 = document.getElementById('article-card-3');
  const textarea = card4.querySelector('textarea.target-textarea');
  const prefillOldBtn = card4.querySelector('.prefill-old-btn');

  assert.ok(prefillOldBtn, 'Tlačítko pro předvyplnění z 2024 existuje na upraveném článku');
  
  // Kliknutí na předvyplnit z 2024
  prefillOldBtn.click();

  assert.ok(textarea.value.includes('Dary a pozvání'), 'Textarea byla naplněna původním českým textem');
  assert.ok(textarea.value.includes('35 eur'), 'Textarea obsahuje 35 eur z roku 2024');

  // Sloupec 4 živý diff byl aktualizován
  const liveDiffBox = card4.querySelector('.live-target-diff');
  assert.ok(liveDiffBox.textContent.length > 0, 'Živý diff byl přepočítán');
});

// TEST 8: Kvalitně generované prompty pro Copilot (stručné, účelné, kompletní kontext)
runTest('Copilot Prompt Quality: Generovaný prompt je výstižný, neobsahuje zbytečný balast a obsahuje všechna data', () => {
  const { window, document } = createDomEnvironment();

  document.getElementById('loadSampleBtn').click();
  document.getElementById('confirmMappingBtn').click();
  document.getElementById('goToTranslationBtn').click();

  // Přepneme na záložku CS 2024
  const tabs = Array.from(document.querySelectorAll('#langTabsWrapper .lang-tab'));
  const csTab = tabs.find(t => t.textContent.includes('CS 2024'));
  csTab.click();

  const card4 = document.getElementById('article-card-3');
  const copyPromptBtn = card4.querySelector('.copy-prompt-btn');
  assert.ok(copyPromptBtn, 'Tlačítko pro kopírování promptu existuje');

  // Kliknutí na tlačítko pro kopírování promptu
  copyPromptBtn.click();

  const prompt = window.getLastCopiedText();
  assert.ok(prompt.length > 0, 'Prompt byl úspěšně vygenerován');

  // Ověření přesných klíčových součástí
  assert.ok(prompt.includes('CS 2024'), 'Prompt obsahuje cílový jazyk');
  assert.ok(prompt.includes('Art. 04'), 'Prompt obsahuje identifikátor odstavce');
  assert.ok(prompt.includes('35 Euro'), 'Prompt obsahuje původní německý text');
  assert.ok(prompt.includes('50 Euro'), 'Prompt obsahuje nový německý text');
  assert.ok(prompt.includes('35 eur'), 'Prompt obsahuje dosavadní ověřený překlad');
  
  // Ověření instrukcí pro minimální změny a čistý výstup
  assert.ok(prompt.includes('POUZE minimální nutné změny'), 'Obsahuje pravidlo minimálních změn');
  assert.ok(prompt.includes('POUZE vraťte výsledný text'), 'Obsahuje příkaz nevracet konverzační úvod');

  // Ověření, že prompt není zbytečně dlouhý (méně než 1600 znaků pro standardní článek)
  assert.ok(prompt.length < 1600, `Prompt je stručný a efektivní (délka: ${prompt.length} znaků)`);

  // Toast notifikace se zobrazila v DOM
  const toast = document.querySelector('#toastContainer .toast');
  assert.ok(toast, 'Toast notifikace se zobrazila v DOM');
});

// TEST 9: Vlastní šablona promptu v GUI
runTest('Custom Prompt Template: Změna šablony v modálním okně se projeví při generování promptu', () => {
  const { window, document } = createDomEnvironment();

  document.getElementById('loadSampleBtn').click();
  document.getElementById('confirmMappingBtn').click();
  document.getElementById('goToTranslationBtn').click();

  // Přepneme na záložku CS 2024
  const tabs = Array.from(document.querySelectorAll('#langTabsWrapper .lang-tab'));
  const csTab = tabs.find(t => t.textContent.includes('CS 2024'));
  csTab.click();

  // Nastavíme vlastní šablonu v nastavení
  const templateInput = document.getElementById('promptTemplateInput');
  const saveBtn = document.getElementById('savePromptTemplateBtn');

  templateInput.value = 'AI_TRANSLATE: {TARGET_LANG} | {ROW_ID} | NEW_DE: {REF_NEW} | OLD_TRANS: {TARGET_OLD}';
  saveBtn.click();

  // Vygenerujeme prompt pro kartu 4
  const card4 = document.getElementById('article-card-3');
  card4.querySelector('.copy-prompt-btn').click();

  const prompt = window.getLastCopiedText();
  assert.ok(prompt.startsWith('AI_TRANSLATE: CS 2024 | Art. 04'), 'Prompt se přesně řídí uživatelskou šablonou');
  assert.ok(prompt.includes('50 Euro'), 'Šablona správně dosadila {REF_NEW}');
});

// TEST 10: Schvalování překladů a progress bar v reálném DOM
runTest('Verification Workflow: Označení odstavce jako ověřený aktualizuje třídy karty a progress bar v DOM', () => {
  const { window, document } = createDomEnvironment();

  document.getElementById('loadSampleBtn').click();
  document.getElementById('confirmMappingBtn').click();
  document.getElementById('goToTranslationBtn').click();

  // Přepneme na záložku CS 2024
  const tabs = Array.from(document.querySelectorAll('#langTabsWrapper .lang-tab'));
  const csTab = tabs.find(t => t.textContent.includes('CS 2024'));
  csTab.click();

  // Vybereme kartu 4 (která je pending, protože je MODIFIED)
  const card4 = document.getElementById('article-card-3');
  const checkbox = card4.querySelector('.verify-checkbox');
  const progressBar = document.getElementById('currentLangProgressBar');
  const progressText = document.getElementById('currentLangProgressText');

  // Původní stav
  const initialWidth = progressBar.style.width;

  // Zaškrtneme jako ověřeno
  checkbox.checked = true;
  checkbox.dispatchEvent(new window.Event('change'));

  // Kontrola DOMu
  assert.ok(card4.classList.contains('status-verified'), 'Karta má třídu status-verified');
  assert.ok(!card4.classList.contains('status-pending'), 'Karta již nemá třídu status-pending');
  assert.ok(progressBar.style.width !== initialWidth, 'Progress bar se posunul');
  assert.ok(progressText.textContent.includes('hotovo'), 'Progress text indikuje hotové odstavce');
});

// TEST 11: Ukládání a načítání projektu (.json) s kompletní obnovou GUI stavu
runTest('Project Save & Load: Export do JSON a zpětné načtení kompletně obnoví reálná data a stav v GUI', async () => {
  const env1 = createDomEnvironment();

  env1.document.getElementById('loadSampleBtn').click();
  env1.document.getElementById('confirmMappingBtn').click();
  env1.document.getElementById('goToTranslationBtn').click();

  // Přepneme na záložku CS 2024
  const tabs = Array.from(env1.document.querySelectorAll('#langTabsWrapper .lang-tab'));
  const csTab = tabs.find(t => t.textContent.includes('CS 2024'));
  csTab.click();

  // Uživatel upraví článek 4
  const card4 = env1.document.getElementById('article-card-3');
  const textarea4 = card4.querySelector('textarea.target-textarea');
  textarea4.value = 'MŮJ TESTOVACÍ ČESKÝ PŘEKLAD 2026 PRO ČLÁNEK 4 - DARY 50 EUR';
  textarea4.dispatchEvent(new env1.window.Event('input'));

  const cb4 = card4.querySelector('.verify-checkbox');
  cb4.checked = true;
  cb4.dispatchEvent(new env1.window.Event('change'));

  // Uživatel upraví šablonu PDF (musí být součástí projektu)
  env1.window.__app.state.pdfTemplate = { companyName: 'Projektová Firma s.r.o.', accentColor: '#123456' };

  // Export projektu tlačítkem v Kroku 4 → zachycený stažený soubor
  env1.document.getElementById('goToExportBtn').click();
  env1.document.getElementById('exportProjectJsonBtn').click();
  const download = env1.window.__downloads.pop();
  assert.ok(download && /_projekt_\d{4}-\d{2}-\d{2}\.json$/.test(download.name), 'Stáhne se soubor projektu .json');
  const projectData = await readBlobText(env1.window, download.blob);
  assert.ok(projectData.includes('MŮJ TESTOVACÍ ČESKÝ PŘEKLAD 2026 PRO ČLÁNEK 4'), 'JSON obsahuje upravený text');

  // Nové čisté prostředí: načtení přes stejnou cestu jako výběr souboru v GUI
  const env2 = createDomEnvironment();
  env2.window.__app.loadProjectFile(new env2.window.File([projectData], download.name, { type: 'application/json' }));
  await sleep(50);
  assert.strictEqual(env2.window.__lastAlert, undefined, 'Načtení nesmí skončit chybou: ' + env2.window.__lastAlert);
  assert.ok(toastTexts(env2.document).some(t => t.includes('úspěšně načten')), 'Zobrazí se potvrzení o načtení');
  assert.deepStrictEqual(plain(env2.window.__app.state.pdfTemplate), { companyName: 'Projektová Firma s.r.o.', accentColor: '#123456' }, 'Šablona PDF byla obnovena');

  // Kontrola, že GUI přešlo do Kroku 3 a otevřelo naposledy aktivní jazyk
  assert.strictEqual(env2.document.getElementById('step3Section').style.display, 'block', 'Po obnovení je aktivní Krok 3');
  assert.ok(env2.document.querySelector('#langTabsWrapper .lang-tab.active').textContent.includes('CS 2024'), 'Aktivní je jazyk CS 2024');
  const restoredStub = env2.window.__app.state.translations['CS 2024'][3];
  assert.ok(restoredStub.targetOld.includes('35 eur'), 'Původní překlad 2024 byl doplněn z řádků tabulky');
  assert.strictEqual(restoredStub.classification, 'MODIFIED');

  // Kontrola konkrétních dat v DOMu
  const restoredCard4 = env2.document.getElementById('article-card-3');
  assert.ok(restoredCard4, 'Karta 4 byla obnovena');
  
  const restoredTextarea4 = restoredCard4.querySelector('textarea.target-textarea');
  assert.strictEqual(restoredTextarea4.value, 'MŮJ TESTOVACÍ ČESKÝ PŘEKLAD 2026 PRO ČLÁNEK 4 - DARY 50 EUR', 'Skutečný text v textarea byl věrně obnoven');

  const restoredCb4 = restoredCard4.querySelector('.verify-checkbox');
  assert.strictEqual(restoredCb4.checked, true, 'Checkbox ověření zůstal zaškrtnutý');
  assert.ok(restoredCard4.classList.contains('status-verified'), 'Karta má status-verified');
});

// TEST 12: Krok 4 - Formátovaný PDF export se skutečnými daty
runTest('Formatted PDF Export: Vygenerovaný tiskový HTML/PDF dokument obsahuje formátování A4 a texty kodexu', async () => {
  const { window, document } = createDomEnvironment();

  document.getElementById('loadSampleBtn').click();
  document.getElementById('confirmMappingBtn').click();
  document.getElementById('goToTranslationBtn').click();

  // Přepneme na záložku CS 2024
  const tabs = Array.from(document.querySelectorAll('#langTabsWrapper .lang-tab'));
  const csTab = tabs.find(t => t.textContent.includes('CS 2024'));
  csTab.click();

  // Upravíme článek 4
  const card4 = document.getElementById('article-card-3');
  const textarea4 = card4.querySelector('textarea.target-textarea');
  textarea4.value = 'Finální ověřený text článku 4 v češtině pro rok 2026 (dary do 50 EUR).';
  textarea4.dispatchEvent(new window.Event('input'));

  // Přechod do Kroku 4
  document.getElementById('goToExportBtn').click();
  assert.strictEqual(document.getElementById('step4Section').style.display, 'block', 'Krok 4 je zobrazen');

  // Tisk přes tlačítko v Kroku 4 pro vybraný jazyk → zachycené tiskové okno
  const pdfLangSelect = document.getElementById('pdfLangSelect');
  assert.strictEqual(pdfLangSelect.value, 'CS 2024', 'Výběr jazyka pro PDF je předvyplněn aktivním jazykem');
  document.getElementById('printPdfBtn').click();
  const printWindow = window.__printWindows.pop();
  assert.ok(printWindow, 'Otevře se tiskové okno');
  await sleep(350);
  assert.ok(printWindow.printed, 'Automaticky se vyvolá tiskový dialog');
  const printableHtml = printWindow.html;

  assert.ok(printableHtml.includes('size: A4'), 'Obsahuje formátovací CSS pravidla pro A4');
  assert.ok(printableHtml.toLowerCase().includes('etický kodex'), 'Obsahuje hlavní nadpis kodexu');
  assert.ok(printableHtml.includes('CS 2026'), 'Obsahuje informaci o jazykové mutaci');
  assert.ok(printableHtml.includes('Finální ověřený text článku 4 v češtině pro rok 2026 (dary do 50 EUR).'), 'Obsahuje skutečný editovaný text');
  assert.ok(printableHtml.includes('Vedení společnosti'), 'Obsahuje podpisový blok vedení');
  assert.ok(!printableHtml.includes('Telefax'), 'Neobsahuje smazaný článek 15 o telefaxu');
  // Ostatní upravené články nejsou přeložené ani ověřené → dokument je zřetelně koncept
  assert.ok(printableHtml.includes('KONCEPT'), 'Neúplný překlad je označen jako koncept');
  const art7Old = window.__app.state.rows[6]['CS 2024'];
  assert.ok(!printableHtml.includes(art7Old), 'Nepřeložený upravený článek nesmí převzít zastaralý text 2024');

  // Stažení HTML verze
  document.getElementById('downloadPdfHtmlBtn').click();
  const htmlDownload = window.__downloads.pop();
  assert.strictEqual(htmlDownload.name, 'Eticky_Kodex_2026_CS_2024.html');
});

// TEST 13: Excel export - tlačítka v Kroku 4 vytvoří skutečné sešity se správným obsahem
runTest('Excel Export: Tlačítka v Kroku 4 vygenerují kompletní a čistý sešit s upravenými texty a stavem ověření', () => {
  const { window } = createDomEnvironment();
  const document = window.document;
  const XLSX = window.XLSX;

  openCzechStudio(window);
  const card4 = document.getElementById('article-card-3');
  typeInto(window, card4.querySelector('textarea.target-textarea'), 'Dary do 50 EUR – nový text.');
  const cb4 = card4.querySelector('.verify-checkbox');
  cb4.checked = true;
  cb4.dispatchEvent(new window.Event('change'));
  // Článek 8 (nový) má jen rozpracovaný koncept
  typeInto(window, document.getElementById('article-card-7').querySelector('textarea.target-textarea'), 'Koncept AI článku');

  document.getElementById('goToExportBtn').click();

  // Komplexní sešit
  document.getElementById('downloadComprehensiveBtn').click();
  const comp = window.__workbooks.pop();
  assert.strictEqual(comp.name, 'vzorec_eticky_kodex_2024_2026_2024_2026_Srovnani.xlsx');
  assert.deepStrictEqual(plain(comp.wb.SheetNames), ['Srovnani_2024_2026_Komplet', 'Stav_prekladu']);
  const compRows = XLSX.utils.sheet_to_json(comp.wb.Sheets.Srovnani_2024_2026_Komplet, { defval: '' });
  assert.strictEqual(compRows.length, 15, 'Komplexní export zachovává všech 15 řádků');
  assert.strictEqual(compRows[3]['CS 2024'], window.__app.state.rows[3]['CS 2024'], 'Původní sloupec zůstal');
  assert.strictEqual(compRows[3]['CS 2026'], 'Dary do 50 EUR – nový text.', 'Nový sloupec obsahuje upravený text');
  assert.strictEqual(compRows[0]['CS 2026'], window.__app.state.rows[0]['CS 2024'], 'Nezměněný článek je převzat');
  assert.strictEqual(compRows[0]['FR 2026'], window.__app.state.rows[0]['FR 2024'], 'I neotevřený jazyk má převzaté nezměněné články');
  const statusRows = XLSX.utils.sheet_to_json(comp.wb.Sheets.Stav_prekladu, { defval: '' });
  assert.strictEqual(statusRows[3]['CS 2026'], 'Ověřeno');
  assert.strictEqual(statusRows[7]['CS 2026'], 'Čeká na kontrolu');
  assert.strictEqual(statusRows[2]['CS 2026'], 'Chybí překlad');
  assert.ok(toastTexts(document).some(t => /Pozor: \d+ odstavců/.test(t)), 'Uživatel je upozorněn na neověřené odstavce');

  // Čistý sešit 2026
  document.getElementById('downloadCleanBtn').click();
  const clean = window.__workbooks.pop();
  assert.strictEqual(clean.name, 'vzorec_eticky_kodex_2024_2026_2026_Final.xlsx');
  const cleanHeader = XLSX.utils.sheet_to_json(clean.wb.Sheets[clean.wb.SheetNames[0]], { header: 1 })[0];
  assert.ok(cleanHeader.includes('DE 2026 (Referenz)'), 'Název referenčního sloupce 2026 se nezdvojuje');
  assert.ok(!cleanHeader.includes('DE 2024 (Referenz)') && !cleanHeader.includes('CS 2024'), 'Bez sloupců 2024');
  const cleanRows = XLSX.utils.sheet_to_json(clean.wb.Sheets[clean.wb.SheetNames[0]], { defval: '' });
  assert.strictEqual(cleanRows.length, 13, 'Čistý export 2026 odfiltroval 2 smazané odstavce (13 řádků)');
  assert.ok(cleanRows.some(r => r['CS 2026'] === 'Dary do 50 EUR – nový text.'), 'Upravený text je v čisté verzi');
  assert.ok(!cleanRows.some(r => r['Artikel-Nr.'] === 'Art. 10'), 'Zrušený článek 10 chybí');
});

// TEST 14: Krok 1 - Interaktivní náhled dat Excelu (Data Preview Table)
runTest('Data Preview Table: Po načtení souboru se v Kroku 1 vykreslí tabulka s reálnými daty', () => {
  const { window } = createDomEnvironment();
  const document = window.document;

  const previewContainer = document.getElementById('dataPreviewContainer');
  const previewHead = document.getElementById('previewTableHead');
  const previewBody = document.getElementById('previewTableBody');

  // Před načtením je container skrytý
  assert.strictEqual(previewContainer.style.display, 'none');

  // Načtení sample dat
  document.getElementById('loadSampleBtn').click();

  // Po načtení je container zobrazen a tabulka naplněna
  assert.strictEqual(previewContainer.style.display, 'block', 'Data preview container je zobrazen');
  assert.ok(previewHead.innerHTML.includes('Artikel-Nr.'), 'Hlavička obsahuje Artikel-Nr.');
  assert.ok(previewHead.innerHTML.includes('DE 2024 (Referenz)'), 'Hlavička obsahuje DE 2024');
  assert.ok(previewHead.innerHTML.includes('CS 2024'), 'Hlavička obsahuje CS 2024');

  // Ověření řádků
  const trs = previewBody.querySelectorAll('tr');
  assert.strictEqual(trs.length, 4, 'Náhled zobrazuje první 4 řádky dat');
  assert.ok(previewBody.innerHTML.includes('Art. 01'), 'Tělo náhledu obsahuje Art. 01');
});

// TEST 15: Cross-Language Progress Matrix a globální schválení beze změny
runTest('Cross-Language Matrix & Bulk Approval: Matice pokroku a hromadné schválení nezměněných', () => {
  const { window } = createDomEnvironment();
  const document = window.document;

  document.getElementById('loadSampleBtn').click();
  document.getElementById('confirmMappingBtn').click();
  document.getElementById('goToTranslationBtn').click();

  const matrixCard = document.getElementById('crossLangMatrixCard');
  assert.strictEqual(matrixCard.style.display, 'block', 'Karta matice jazyků je zobrazena');

  const matrixGrid = document.getElementById('crossLangMatrixGrid');
  const pills = matrixGrid.querySelectorAll('.matrix-pill');
  assert.ok(pills.length >= 10, 'Matice obsahuje pill pro každý cílový jazyk');

  // Aktivní pill odpovídá aktivnímu jazyku
  const activePill = matrixGrid.querySelector('.matrix-pill.active');
  assert.ok(activePill, 'Existuje aktivní pill');

  // Nezměněné články jsou schváleny automaticky; uživatel jedno schválení v CS a jedno v EN zruší
  const csStubs = window.__app.state.translations['CS 2024'];
  const enStubs = window.__app.state.translations['EN 2024'];
  csStubs[0].isVerified = false;
  enStubs[1].isVerified = false;

  // Kliknutí na globální schválení nezměněných
  const globalApproveBtn = document.getElementById('approveAllUnchangedGlobalBtn');
  globalApproveBtn.click();

  // Ověření, že pro CS 2024 i ostatní jazyky došlo ke schválení nezměněných
  const unchangedCs = csStubs.filter(s => s.classification === 'UNCHANGED');
  assert.ok(unchangedCs.every(s => s.isVerified === true), 'Všechny nezměněné články v češtině jsou schváleny');
  assert.strictEqual(enStubs[1].isVerified, true, 'Schválení proběhlo i v jiném jazyce');
  assert.ok(toastTexts(document).some(t => t.includes('Schváleno 2 nezměněných')), 'Toast hlásí přesný počet (a ne surový klíč překladu)');

  // Aktualizace matice zobrazí nový počet schválených
  const csPill = Array.from(matrixGrid.querySelectorAll('.matrix-pill')).find(p => p.textContent.includes('CS 2024'));
  assert.ok(csPill && csPill.textContent.includes('7/15'), 'Pill pro CS 2024 zobrazuje 7 schválených z 15');
});

// TEST 16: Schválení nezměněných v aktivním jazyce a klávesová zkratka Ctrl+Enter (CAT)
runTest('CAT Shortcut & Active Language Bulk Approve: Klávesová zkratka Ctrl+Enter a hromadné schválení', () => {
  const { window } = createDomEnvironment();
  const document = window.document;

  document.getElementById('loadSampleBtn').click();
  document.getElementById('confirmMappingBtn').click();
  document.getElementById('goToTranslationBtn').click();

  // Přepnutí na češtinu (CS 2024)
  const csTab = Array.from(document.querySelectorAll('.lang-tab')).find(t => t.textContent.includes('CS 2024'));
  if (csTab) csTab.click();

  // Test tlačítka Schválit nezměněné v aktivním jazyce
  const approveActiveBtn = document.getElementById('approveUnchangedBtn');
  approveActiveBtn.click();

  const csStubs = window.__app.state.translations['CS 2024'];
  assert.strictEqual(csStubs[0].isVerified, true, 'Článek 1 (nezměněný) je ověřen');

  // Test CAT zkratky Ctrl + Enter na článku 4 (který je MODIFIED a dosud neověřen)
  assert.strictEqual(csStubs[3].isVerified, false, 'Článek 4 byl předtím neověřen');
  const card4 = document.getElementById('article-card-3');
  const textarea4 = card4.querySelector('.target-textarea');
  textarea4.value = 'Upravený text článku 4 potvrzený klávesou Ctrl+Enter.';
  textarea4.dispatchEvent(new window.Event('input'));

  // Simulace klávesové zkratky Ctrl + Enter
  const ctrlEnterEvent = new window.KeyboardEvent('keydown', {
    key: 'Enter',
    ctrlKey: true,
    bubbles: true,
    cancelable: true
  });
  textarea4.dispatchEvent(ctrlEnterEvent);

  // Ověření, že článek 4 byl okamžitě označen jako ověřený
  assert.strictEqual(csStubs[3].isVerified, true, 'Článek 4 je nyní ověřen díky Ctrl+Enter');
  assert.strictEqual(card4.classList.contains('status-verified'), true, 'Karta má třídu status-verified');
  const cb4 = card4.querySelector('.verify-checkbox');
  assert.strictEqual(cb4.checked, true, 'Checkbox je zaškrtnut');
});

// TEST 17: Dávkový prompt pro Copilot a Smart Multi-Paste v DOM
runTest('Batch Copilot Prompt & Smart Multi-Paste: Hromadné prompty a vkládání přeložených dávek v GUI', () => {
  const { window } = createDomEnvironment();
  const document = window.document;

  document.getElementById('loadSampleBtn').click();
  document.getElementById('confirmMappingBtn').click();
  document.getElementById('goToTranslationBtn').click();

  // Přepnutí na češtinu (CS 2024)
  const csTab = Array.from(document.querySelectorAll('.lang-tab')).find(t => t.textContent.includes('CS 2024'));
  if (csTab) csTab.click();

  // 1. Otevření modálního okna dávkového promptu
  const batchPromptBtn = document.getElementById('batchPromptBtn');
  batchPromptBtn.click();

  const batchPromptModal = document.getElementById('batchPromptModal');
  assert.strictEqual(batchPromptModal.classList.contains('open'), true, 'Modální okno dávkového promptu je otevřeno');

  const batchTextarea = document.getElementById('batchPromptTextarea');
  assert.ok(batchTextarea.value.includes('DÁVKOVÝ PŘEKLAD'), 'Textarea obsahuje dávkový prompt');
  assert.ok(batchTextarea.value.includes('Art. 04'), 'Obsahuje článek 4');
  assert.ok(batchTextarea.value.includes('Art. 08'), 'Obsahuje článek 8');

  // Zavření okna
  document.getElementById('closeBatchPromptBtn').click();
  assert.strictEqual(batchPromptModal.classList.contains('open'), false, 'Modální okno se zavřelo');

  // 2. Otevření Smart Multi-Paste okna
  const smartPasteBtn = document.getElementById('smartPasteBtn');
  smartPasteBtn.click();

  const smartPasteModal = document.getElementById('smartPasteModal');
  assert.strictEqual(smartPasteModal.classList.contains('open'), true, 'Smart Paste modální okno je otevřeno');

  // Vložení odpovědi z LLM
  const pasteInput = document.getElementById('smartPasteInput');
  pasteInput.value = `
[Art. 04]
Dary a pohostinnost: Přijímání darů je nově omezeno částkou 50 EUR ročně.

[Art. 06]
Ochrana údajů: Všechny systémy podléhají směrnici AI 2026.

[Art. 99]
Text článku, který v kodexu neexistuje.
  `;
  const art1Before = window.__app.state.translations['CS 2024'][0].targetNew;

  // Kliknutí na analyzovat (Parse)
  const parseBtn = document.getElementById('smartPasteParseBtn');
  parseBtn.click();

  // Ověření zobrazení rozpoznaných odstavců
  const matchesContainer = document.getElementById('smartPasteMatchesContainer');
  assert.strictEqual(matchesContainer.style.display, 'block', 'Kontejner rozpoznaných položek je zobrazen');

  const badges = matchesContainer.querySelectorAll('.matched-article-badge');
  assert.strictEqual(badges.length, 2, 'Byly rozpoznány přesně 2 odstavce');
  assert.ok(badges[0].textContent.includes('Art. 04'), 'První odznak je Art. 04');
  assert.ok(badges[1].textContent.includes('Art. 06'), 'Druhý odznak je Art. 06');
  const unmatchedInfo = matchesContainer.querySelector('.smart-paste-unmatched');
  assert.ok(unmatchedInfo && unmatchedInfo.textContent.includes('Art. 99'), 'Nespárovaná značka je uživateli ukázána');

  // Aplikace do studia
  const applyBtn = document.getElementById('smartPasteApplyBtn');
  assert.strictEqual(applyBtn.disabled, false, 'Tlačítko aplikovat je aktivní');
  applyBtn.click();

  // Ověření, že texty byly zapsány do DOM a stavu
  const csStubs = window.__app.state.translations['CS 2024'];
  assert.ok(csStubs[3].targetNew.includes('50 EUR ročně'), 'Článek 4 má nový text z Smart Paste');
  assert.ok(csStubs[5].targetNew.includes('směrnici AI 2026'), 'Článek 6 má nový text z Smart Paste');
  assert.strictEqual(csStubs[3].isVerified, true, 'Článek 4 byl automaticky ověřen');
  assert.strictEqual(csStubs[0].targetNew, art1Before, 'Nespárovaný text nepřepsal žádný jiný článek');
  assert.ok(!csStubs.some(s => s.targetNew.includes('neexistuje')), 'Text neexistujícího článku nebyl nikam vložen');
  assert.strictEqual(document.getElementById('article-card-3').querySelector('textarea').value, csStubs[3].targetNew, 'Karta v DOM ukazuje vložený text');
  assert.ok(toastTexts(document).some(t => t.includes('Vloženo 2 přeložených článků')), 'Toast hlásí počet vložených článků');
});

// TEST 18: Manažerský souhrn změn (Markdown, Word .doc, PDF) v GUI
runTest('Managerial Change Summary: Export souhrnu změn v Markdown, Wordu a PDF z Kroku 2 a 4', async () => {
  const { window } = createDomEnvironment();
  const document = window.document;

  document.getElementById('loadSampleBtn').click();
  document.getElementById('confirmMappingBtn').click();

  // Ověření existence tlačítek v Kroku 2
  const step2MdBtn = document.getElementById('step2ExportMdBtn');
  const step2DocBtn = document.getElementById('step2ExportDocBtn');
  const step2PdfBtn = document.getElementById('step2ExportPdfBtn');

  assert.ok(step2MdBtn && step2DocBtn && step2PdfBtn, 'Tlačítka pro manažerský export v Kroku 2 existují');

  // Skutečné stažení z Kroku 2
  step2MdBtn.click();
  const md = window.__downloads.pop();
  assert.strictEqual(md.name, 'Eticky_Kodex_Manazerske_Shrnuti_Zmen_2024_2026.md');
  const mdText = await readBlobText(window, md.blob);
  assert.ok(mdText.includes('| **Celkem hodnocených odstavců** | 15 | 100% |'), 'Markdown obsahuje tabulku KPI');
  assert.ok(mdText.includes('Art. 08'), 'Markdown obsahuje nově přidaný článek');

  step2DocBtn.click();
  const doc = window.__downloads.pop();
  assert.ok(doc.name.endsWith('.doc'));
  const docText = await readBlobText(window, doc.blob);
  assert.ok(docText.includes('urn:schemas-microsoft-com:office:word') && docText.includes('Art. 04'), 'Word dokument obsahuje data');

  step2PdfBtn.click();
  const mgrPrint = window.__printWindows.pop();
  assert.ok(mgrPrint && mgrPrint.html.includes('Manažerské shrnutí změn'), 'Tiskové okno obsahuje report');

  const toasts = toastTexts(document);
  assert.ok(toasts.some(t => t.includes('Markdown')) && toasts.some(t => t.includes('Word')), 'Toasty mají lidský text');
  assert.ok(!toasts.some(t => /mgrExport/.test(t)), 'Žádný surový klíč překladu');

  // Přechod do Kroku 4
  document.getElementById('goToTranslationBtn').click();
  document.getElementById('goToExportBtn').click();

  // Ověření existence tlačítek v Kroku 4
  const step4MdBtn = document.getElementById('downloadMgrMdBtn');
  const step4DocBtn = document.getElementById('downloadMgrDocBtn');
  const step4PdfBtn = document.getElementById('printMgrPdfBtn');

  assert.ok(step4MdBtn && step4DocBtn && step4PdfBtn, 'Tlačítka pro manažerský export v Kroku 4 existují');

  // Ověření generování obsahu manažerského souhrnu
  const summary = window.ComparatorEngine.generateManagerialChangeSummary({
    rows: window.__app.state.rows,
    headers: window.__app.state.headers,
    mappings: window.__app.state.mappings,
    companyName: 'Corporate Compliance AG',
    uiLang: 'cs'
  });

  assert.ok(summary.markdown.includes('Manažerské shrnutí změn'), 'Obsahuje markdown nadpis');
  assert.ok(summary.htmlDoc.includes('xmlns:w="urn:schemas-microsoft-com:office:word"'), 'Word dokument má Word namespace');
  assert.strictEqual(summary.stats.totalRows, 15);
});

// TEST 19: Uživatelská úprava PDF šablony v GUI
runTest('PDF Template Customization Modal: Úprava šablony, barev a podpisů v GUI', () => {
  const { window } = createDomEnvironment();
  const document = window.document;

  document.getElementById('loadSampleBtn').click();
  document.getElementById('confirmMappingBtn').click();
  document.getElementById('goToTranslationBtn').click();
  document.getElementById('goToExportBtn').click();

  // Otevření okna šablony
  const editTplBtn = document.getElementById('editPdfTemplateBtn');
  editTplBtn.click();

  const tplModal = document.getElementById('pdfTemplateModal');
  assert.strictEqual(tplModal.classList.contains('open'), true, 'Modální okno šablony PDF je otevřeno');

  // Úprava hodnot v GUI
  const compInput = document.getElementById('tplCompanyName');
  compInput.value = 'Holding Bohemia a.s.';

  const titleInput = document.getElementById('tplDocTitle');
  titleInput.value = 'FIREMNÍ ETICKÝ KODEX 2026';

  const accentInput = document.getElementById('tplAccentColor');
  accentInput.value = '#7c3aed';

  const dept1Input = document.getElementById('tplSignDept1');
  dept1Input.value = 'Oddělení udržitelnosti a ESG';

  // Uložení šablony
  document.getElementById('savePdfTemplateBtn').click();
  assert.strictEqual(tplModal.classList.contains('open'), false, 'Modální okno se zavřelo');

  // Ověření, že stav obsahuje novou šablonu
  const savedTpl = window.__app.state.pdfTemplate;
  assert.ok(savedTpl, 'Šablona je uložena ve stavu');
  assert.strictEqual(savedTpl.companyName, 'Holding Bohemia a.s.');
  assert.strictEqual(savedTpl.docTitle, 'FIREMNÍ ETICKÝ KODEX 2026');
  assert.strictEqual(savedTpl.accentColor, '#7c3aed');
  assert.strictEqual(savedTpl.signDept1, 'Oddělení udržitelnosti a ESG');

  // Ověření, že generovaný PDF dokument novou šablonu plně respektuje
  const docHtml = window.PdfExporter.generateDocumentHtml({
    rows: window.__app.state.rows,
    headers: window.__app.state.headers,
    mappings: window.__app.state.mappings,
    translations: window.__app.state.translations,
    targetCol: 'CS 2024',
    template: savedTpl,
    uiLang: 'cs'
  });

  assert.ok(docHtml.includes('Holding Bohemia a.s.'), 'Dokument obsahuje upravený název firmy');
  assert.ok(docHtml.includes('FIREMNÍ ETICKÝ KODEX 2026'), 'Dokument obsahuje upravený název kodexu');
  assert.ok(docHtml.includes('#7c3aed'), 'Dokument obsahuje upravenou accent barvu');
  assert.ok(docHtml.includes('Oddělení udržitelnosti a ESG'), 'Dokument obsahuje upravený podpisový blok');
});

// TEST 20: Automatické ukládání do LocalStorage a obnovení po znovuotevření aplikace
runTest('Auto-Save & Resume: Úpravy se uloží do prohlížeče a po znovuotevření se obnoví', async () => {
  const env1 = createDomEnvironment();
  const saveText = env1.document.getElementById('saveStatusText');
  assert.ok(saveText.textContent.includes('uloženo'), 'Výchozí stav je uloženo');

  openCzechStudio(env1.window);
  const card4 = env1.document.getElementById('article-card-3');
  typeInto(env1.window, card4.querySelector('textarea.target-textarea'), 'Automaticky uložený text čl. 4');
  assert.strictEqual(saveText.textContent, 'Ukládám...', 'Během čekání na uložení indikátor ukazuje ukládání');

  // Zavření stránky uloží rozpracovanou změnu okamžitě (bez čekání na debounce)
  env1.window.dispatchEvent(new env1.window.Event('pagehide'));
  assert.strictEqual(saveText.textContent, 'Vše uloženo');
  const stored = env1.window.localStorage.getItem('ethic_comparator_workspace_v1');
  assert.ok(stored && stored.includes('Automaticky uložený text čl. 4'), 'LocalStorage obsahuje poslední úpravu');

  // Nové otevření aplikace se stejným úložištěm a potvrzením obnovy
  const env2 = createDomEnvironment({
    acceptResume: true,
    beforeInit: (w) => w.localStorage.setItem('ethic_comparator_workspace_v1', stored)
  });
  await sleep(150);
  assert.ok(/rozpracovaný projekt/.test(env2.window.__lastConfirm), 'Uživatel je dotázán na obnovení');
  assert.strictEqual(env2.document.getElementById('step3Section').style.display, 'block', 'Obnoví se Krok 3');
  const restored = env2.document.getElementById('article-card-3').querySelector('textarea.target-textarea');
  assert.strictEqual(restored.value, 'Automaticky uložený text čl. 4');
});

runTest('Auto-Save Failure: Selhání ukládání (plné úložiště) je viditelně oznámeno', () => {
  const { window } = createDomEnvironment();
  const document = window.document;
  openCzechStudio(window);

  const storageProto = Object.getPrototypeOf(window.localStorage);
  const originalSetItem = storageProto.setItem;
  storageProto.setItem = function (key, value) {
    if (key === 'ethic_comparator_workspace_v1') {
      const err = new Error('QuotaExceededError');
      err.name = 'QuotaExceededError';
      throw err;
    }
    return originalSetItem.call(this, key, value);
  };
  const originalWarn = window.console.warn;
  window.console.warn = () => {};
  try {
    typeInto(window, document.getElementById('article-card-3').querySelector('textarea.target-textarea'), 'Text, který se neuloží');
    window.StorageManager.flushPendingSave();
  } finally {
    storageProto.setItem = originalSetItem;
    window.console.warn = originalWarn;
  }

  assert.ok(document.getElementById('saveIndicator').classList.contains('save-error'), 'Indikátor je v chybovém stavu');
  assert.ok(document.getElementById('saveStatusText').textContent.includes('Neuloženo'), 'Text indikátoru varuje');
  assert.ok(toastTexts(document).some(t => t.includes('exportem projektu')), 'Uživatel dostane radu, jak práci zálohovat');
});

// TEST 22: Nahrání skutečného Excel souboru + automatická detekce sloupců
runTest('Excel Upload: Nahrání vzorového .xlsx souboru detekuje referenční i cílové sloupce', async () => {
  const { window } = createDomEnvironment();
  const document = window.document;
  const bytes = fs.readFileSync(path.join(__dirname, '..', 'sample', 'vzorec_eticky_kodex_2024_2026.xlsx'));
  window.__app.processUploadedFile(new window.File([bytes], 'vzorec_eticky_kodex_2024_2026.xlsx'));
  await sleep(100);

  assert.strictEqual(window.__lastAlert, undefined, 'Bez chyby: ' + window.__lastAlert);
  const state = window.__app.state;
  assert.strictEqual(state.rows.length, 15);
  assert.strictEqual(state.headers[0], 'Artikel-Nr.', 'Pořadí sloupců odpovídá souboru');
  assert.strictEqual(state.mappings.refOldCol, 'DE 2024 (Referenz)');
  assert.strictEqual(state.mappings.refNewCol, 'DE 2026 (Referenz)');
  assert.deepStrictEqual(plain(state.mappings.targetCols), ['EN 2024', 'CS 2024', 'FR 2024', 'ES 2024', 'IT 2024', 'PL 2024', 'NL 2024', 'SK 2024', 'HU 2024', 'RO 2024'],
    'Cílové jazyky jsou přesně jazykové sloupce (ne ID ani oblast)');
  assert.strictEqual(document.getElementById('refOldSelect').value, 'DE 2024 (Referenz)');
});

runTest('CSV Upload: UTF-8 CSV zachová diakritiku a nový soubor zahodí překlady předchozího', async () => {
  const { window } = createDomEnvironment();
  const document = window.document;

  // Rozpracovaný projekt z ukázky
  openCzechStudio(window);
  typeInto(window, document.getElementById('article-card-0').querySelector('textarea.target-textarea'), 'STARÁ PRÁCE');

  const csv = '﻿ID;DE 2024;DE 2026;CS 2024\n1;Größe;Größe neu;Příliš žluťoučký kůň\n2;Gleich;Gleich;Stejné\n';
  window.__app.processUploadedFile(new window.File([csv], 'kodex.csv', { type: 'text/csv' }));
  await sleep(100);

  const state = window.__app.state;
  assert.strictEqual(state.rows[0]['DE 2024'], 'Größe', 'Přehlásky jsou zachovány');
  assert.strictEqual(state.rows[0]['CS 2024'], 'Příliš žluťoučký kůň', 'Česká diakritika je zachována');
  assert.deepStrictEqual(plain(state.headers), ['ID', 'DE 2024', 'DE 2026', 'CS 2024']);
  assert.deepStrictEqual(plain(state.translations), {}, 'Překlady předchozího souboru byly zahozeny');
  assert.strictEqual(state.mappings.refOldCol, 'DE 2024');
  assert.deepStrictEqual(plain(state.mappings.targetCols), ['CS 2024']);

  document.getElementById('confirmMappingBtn').click();
  document.getElementById('goToTranslationBtn').click();
  const stubs = state.translations['CS 2024'];
  assert.ok(!stubs.some(s => s.targetNew === 'STARÁ PRÁCE'), 'Stará práce se neobjeví u nových řádků');
  assert.strictEqual(stubs[1].targetNew, 'Stejné', 'Nezměněný řádek nového souboru je převzat');
});

// TEST 24: Navigace nepustí uživatele dál s neplatným mapováním
runTest('Step Guards: Krok 2/3 se neotevře bez platného mapování sloupců', () => {
  const { window } = createDomEnvironment();
  const document = window.document;

  document.getElementById('stepNav3').click();
  assert.notStrictEqual(document.getElementById('step1Section').style.display, 'none', 'Bez dat zůstává Krok 1');
  assert.strictEqual(document.getElementById('step3Section').style.display, 'none');

  document.getElementById('loadSampleBtn').click();
  const refNewSelect = document.getElementById('refNewSelect');
  refNewSelect.value = 'DE 2024 (Referenz)';
  refNewSelect.onchange();

  document.getElementById('stepNav2').click();
  assert.strictEqual(document.getElementById('step2Section').style.display, 'none', 'Stejné referenční sloupce nepustí do Kroku 2');
  assert.ok(toastTexts(document).some(t => t.includes('musí být rozdílné')), 'Uživatel vidí důvod');
});

runTest('Modals: Escape a klik na pozadí zavřou modální okno', () => {
  const { window } = createDomEnvironment();
  const document = window.document;
  const modal = document.getElementById('promptSettingsModal');

  document.getElementById('settingsPromptBtn').click();
  assert.ok(modal.classList.contains('open'));
  document.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
  assert.ok(!modal.classList.contains('open'), 'Escape zavře okno');

  document.getElementById('settingsPromptBtn').click();
  modal.dispatchEvent(new window.MouseEvent('click', { bubbles: true }));
  assert.ok(!modal.classList.contains('open'), 'Klik na pozadí zavře okno');
});

// Výsledky
runAllTests().then(() => {
  console.log('\n========================================');
  console.log(`VÝSLEDEK BATERIE GUI TESTŮ: ${passedTests} úspěšných, ${failedTests} chyb`);
  console.log('========================================\n');
  process.exit(failedTests > 0 ? 1 : 0);
});

