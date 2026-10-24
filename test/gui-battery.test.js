/**
 * Baterie testů grafického uživatelského rozhraní (GUI Battery Tests)
 *
 * Testy klikají na skutečné ovládací prvky v DOM (jsdom) a kontrolují výsledná data:
 * výběr sloupců, překladové studio, prompty a Smart Paste, ukládání projektu tlačítky v záhlaví,
 * autosave, exporty do Excelu (kompletní, nová verze, částečný), dokument pro tisk
 * a shrnutí změn pro vybraný jazyk.
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
  // ONLY=<text> runs just the tests whose name contains the text
  const only = process.env.ONLY;
  for (const test of tests.filter(t => !only || t.name.includes(only))) {
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

function readBlobBuffer(window, blob) {
  return new Promise((resolve, reject) => {
    const reader = new window.FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = () => reject(reader.error);
    reader.readAsArrayBuffer(blob);
  });
}

// The last downloaded file parsed as a workbook: { name, wb }
async function lastWorkbook(window) {
  const download = window.__downloads.pop();
  assert.ok(download && /\.xlsx$/.test(download.name), 'Stáhne se sešit .xlsx: ' + (download && download.name));
  const buffer = await readBlobBuffer(window, download.blob);
  return { name: download.name, wb: window.XLSX.read(new Uint8Array(buffer), { type: 'array' }) };
}

function toastTexts(document) {
  return Array.from(document.querySelectorAll('#toastContainer .toast')).map(t => t.textContent);
}

// Loads the sample, confirms the mapping and opens the CS tab of the studio
function openCzechStudio(window) {
  const document = window.document;
  document.getElementById('loadSampleBtn').click();
  document.getElementById('confirmMappingBtn').click();
  document.getElementById('goToTranslationBtn').click();
  const csTab = Array.from(document.querySelectorAll('#langTabsWrapper .lang-tab'))
    .find(t => t.querySelector('span').textContent === 'CS');
  csTab.click();
}

function SampleData() {
  return require('../js/core/sample-generator').getSampleData();
}

function typeInto(window, textarea, text) {
  textarea.value = text;
  textarea.dispatchEvent(new window.Event('input'));
}

console.log('=== Spouštím baterii GUI testů pro Comparator ===\n');

// 1. Příprava JSDOM prostředí
const htmlPath = path.join(__dirname, '..', 'index.html');
const htmlSource = fs.readFileSync(htmlPath, 'utf8');

const scriptSources = [...htmlSource.matchAll(/<script src="([^"]+)"><\/script>/g)]
  .map(m => fs.readFileSync(path.join(__dirname, '..', m[1]), 'utf8'));
const STORAGE_KEY = 'i18n_manager_workspace_v1';

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
    if (window.__confirmAnswer !== undefined) return window.__confirmAnswer;
    return /rozpracovaný projekt z|saved project from/.test(msg) ? !!opts.acceptResume : true;
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
  // Scripts in index.html order; opts.beforeApp may register extra adapters before the app starts
  scriptSources.forEach((src, i) => {
    if (i === scriptSources.length - 1 && opts.beforeApp) opts.beforeApp(window);
    window.eval(src);
  });

  // Inicializace aplikace (DOMContentLoaded)
  window.document.dispatchEvent(new window.Event('DOMContentLoaded'));

  return { dom, window, document: window.document };
}

// BATERIE TESTŮ

const NodeXLSX = require('xlsx');

function langRow(document, col) {
  return Array.from(document.querySelectorAll('#languagesTableBody tr')).find(tr => tr.getAttribute('data-col') === col);
}

function workbookFile(window, sheets, name) {
  const wb = NodeXLSX.utils.book_new();
  Object.keys(sheets).forEach(sheetName => {
    NodeXLSX.utils.book_append_sheet(wb, NodeXLSX.utils.aoa_to_sheet(sheets[sheetName]), sheetName);
  });
  const buf = NodeXLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
  return new window.File([buf], name || 'data.xlsx');
}

function createTestDocx(paragraphs) {
  const zlib = require('zlib');
  const bodyXml = paragraphs.map(p => '<w:p><w:r><w:t>' + p.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;') + '</w:t></w:r></w:p>').join('');
  const docXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>' + bodyXml + '</w:body></w:document>';
  const nameBuf = Buffer.from('word/document.xml', 'utf8');
  const compData = zlib.deflateRawSync(Buffer.from(docXml, 'utf8'));

  const lh = Buffer.alloc(30 + nameBuf.length + compData.length);
  lh.writeUInt32LE(0x04034b50, 0);
  lh.writeUInt16LE(20, 4);
  lh.writeUInt16LE(0, 6);
  lh.writeUInt16LE(8, 8);
  lh.writeUInt32LE(compData.length, 18);
  lh.writeUInt32LE(Buffer.byteLength(docXml), 22);
  lh.writeUInt16LE(nameBuf.length, 26);
  lh.writeUInt16LE(0, 28);
  nameBuf.copy(lh, 30);
  compData.copy(lh, 30 + nameBuf.length);

  const cd = Buffer.alloc(46 + nameBuf.length);
  cd.writeUInt32LE(0x02014b50, 0);
  cd.writeUInt16LE(20, 4);
  cd.writeUInt16LE(20, 6);
  cd.writeUInt16LE(0, 8);
  cd.writeUInt16LE(8, 10);
  cd.writeUInt32LE(compData.length, 20);
  cd.writeUInt32LE(Buffer.byteLength(docXml), 24);
  cd.writeUInt16LE(nameBuf.length, 28);
  cd.writeUInt32LE(0, 42);
  nameBuf.copy(cd, 46);

  const eocd = Buffer.alloc(22);
  eocd.writeUInt32LE(0x06054b50, 0);
  eocd.writeUInt16LE(1, 8);
  eocd.writeUInt16LE(1, 10);
  eocd.writeUInt32LE(cd.length, 12);
  eocd.writeUInt32LE(lh.length, 16);

  return Buffer.concat([lh, cd, eocd]);
}

function setSelect(window, select, value) {
  select.value = value;
  if (select.onchange) select.onchange();
  else select.dispatchEvent(new window.Event('change'));
}

// TEST 1: Inicializace GUI a přepínání jazyků
runTest('Aplikace se inicializuje jen jednou (i když DOMContentLoaded přijde znovu)', async () => {
  const { window, document } = createDomEnvironment();
  await sleep(20);
  document.dispatchEvent(new window.Event('DOMContentLoaded'));
  document.getElementById('langToggleBtn').click();
  assert.strictEqual(window.I18n.getLanguage(), 'en', 'Jedno kliknutí přepne jazyk právě jednou');
  document.getElementById('loadSampleBtn').click();
  assert.strictEqual(toastTexts(document).filter(t => t.includes('Sample data loaded')).length, 1, 'Akce proběhne jednou');
});

runTest('GUI Boot: výchozí zobrazení, obecné texty a přepínání jazyků CZ/EN', () => {
  const { document } = createDomEnvironment();

  assert.strictEqual(document.documentElement.getAttribute('lang'), 'cs');
  assert.ok(document.getElementById('stepNav1').classList.contains('active'));
  assert.strictEqual(document.querySelector('.brand-title').textContent, 'Internationalisation Manager');
  assert.ok(!/etick|kodex|němčin|2024|2026/i.test(document.body.textContent), 'Rozhraní nepředpokládá etický kodex, němčinu ani roky');

  const langBtn = document.getElementById('langToggleBtn');
  assert.strictEqual(langBtn.textContent.trim(), '🇬🇧 English');
  langBtn.click();
  assert.strictEqual(document.documentElement.getAttribute('lang'), 'en');
  assert.strictEqual(langBtn.textContent.trim(), '🇨🇿 Čeština');
  assert.strictEqual(document.querySelector('[data-i18n="dropzoneTitle"]').textContent.trim(), 'Drag & drop a file (Excel, OpenDocument, CSV, Word or text) here or click to browse');
  assert.strictEqual(document.getElementById('saveProjectBtn').textContent, '💾 Save project');
  langBtn.click();
  assert.strictEqual(document.getElementById('loadProjectBtn').textContent, '📂 Načíst projekt');
});

// TEST 2: Výběr sloupců – tabulka použitých a ignorovaných sloupců
runTest('Column Table: ukázková data – reference, 10 jazyků s názvy, ignorované sloupce', () => {
  const { document } = createDomEnvironment();
  document.getElementById('loadSampleBtn').click();

  assert.strictEqual(document.getElementById('mappingContainer').style.display, 'block');
  assert.strictEqual(document.getElementById('refOldSelect').value, 'DE 2024 (Referenz)');
  assert.strictEqual(document.getElementById('refNewSelect').value, 'DE 2026 (Referenz)');

  const rows = Array.from(document.querySelectorAll('#languagesTableBody tr'));
  assert.strictEqual(rows.length, 12, 'Každý sloupec kromě dvou referenčních má řádek');
  const used = rows.filter(r => r.querySelector('.lang-use-cb').checked).map(r => r.getAttribute('data-col'));
  assert.strictEqual(used.length, 10);
  assert.ok(!used.includes('Artikel-Nr.') && !used.includes('Bereich'), 'Ostatní sloupce jsou ignorovány');
  assert.strictEqual(langRow(document, 'CS 2024').querySelector('.lang-label-input').value, 'CS', 'Název jazyka bez čísla verze');
  assert.ok(langRow(document, 'Bereich').textContent.includes('ignorováno'));
  assert.strictEqual(langRow(document, 'CS 2024').querySelector('.lang-newcol-select').value, '', 'Bez existujícího sloupce se vytvoří nový');
  assert.strictEqual(langRow(document, 'CS 2024').querySelector('.lang-newname-input').placeholder, 'CS 2026');
});

runTest('Column Table: vypnutí jazyka, vlastní název, existující sloupec nové verze a změna reference', () => {
  const { window, document } = createDomEnvironment();
  const state = () => window.__app.state;
  document.getElementById('loadSampleBtn').click();

  // Vypnutí jazyka
  const plCb = langRow(document, 'PL 2024').querySelector('.lang-use-cb');
  plCb.checked = false;
  plCb.dispatchEvent(new window.Event('change'));
  assert.ok(!state().mappings.targetCols.includes('PL 2024'));
  assert.ok(langRow(document, 'PL 2024').textContent.includes('ignorováno'));

  // Vlastní název jazyka
  const label = langRow(document, 'CS 2024').querySelector('.lang-label-input');
  label.value = 'Čeština';
  label.dispatchEvent(new window.Event('input'));
  assert.strictEqual(state().mappings.languages['CS 2024'].label, 'Čeština');

  // Existující sloupec jako nová verze jazyka (nabízí jen nepoužité sloupce)
  const select = langRow(document, 'CS 2024').querySelector('.lang-newcol-select');
  const offered = Array.from(select.options).map(o => o.value);
  assert.ok(offered.includes('PL 2024') && offered.includes('Bereich'));
  assert.ok(!offered.includes('EN 2024') && !offered.includes('DE 2024 (Referenz)'), 'Použité a referenční sloupce nejsou nabízeny');
  setSelect(window, select, 'PL 2024');
  assert.strictEqual(state().mappings.languages['CS 2024'].newCol, 'PL 2024');
  const paired = langRow(document, 'PL 2024');
  assert.ok(paired.textContent.includes('nová verze jazyka Čeština'));
  assert.ok(paired.querySelector('.lang-use-cb').disabled, 'Sloupec nové verze nelze zároveň použít jako jazyk');

  // Změna reference: sloupec se stane referencí a zmizí z jazyků
  setSelect(window, document.getElementById('refOldSelect'), 'EN 2024');
  assert.ok(!state().mappings.targetCols.includes('EN 2024'));
  assert.ok(!langRow(document, 'EN 2024'), 'Referenční sloupec není v tabulce jazyků');
  assert.ok(langRow(document, 'DE 2024 (Referenz)'), 'Dřívější reference je teď běžný sloupec');
});

// TEST 4: Krok 2
runTest('Step 2: tabulka změn reference s čísly řádků a <ins>/<del>', () => {
  const { document } = createDomEnvironment();
  document.getElementById('loadSampleBtn').click();
  document.getElementById('confirmMappingBtn').click();

  assert.strictEqual(document.getElementById('kpiTotalRows').textContent, '15');
  assert.strictEqual(document.getElementById('kpiModified').textContent, '6');
  const rows = document.querySelectorAll('#refDiffTableBody tr');
  assert.strictEqual(rows.length, 15);
  assert.strictEqual(rows[0].querySelector('td').textContent.trim(), 'Řádek 2', 'Odstavec je označen číslem řádku v Excelu');
  assert.ok(!rows[0].querySelector('td').textContent.includes('Art.'), 'Nepředpokládá se sloupec s nadpisem');
  const diff4 = rows[3].querySelector('.diff-box').innerHTML;
  assert.ok(diff4.includes('<ins class="diff-ins"') && diff4.includes('<del class="diff-del"') && diff4.includes('50') && diff4.includes('35'));
  assert.ok(rows[7].querySelector('.badge-added') && rows[9].querySelector('.badge-removed'));
});

// TEST 5: Krok 3
runTest('Studio: záložky s názvy jazyků, karta bez nadpisů a bez duplicitního textu nad editorem', () => {
  const { window, document } = createDomEnvironment();
  openCzechStudio(window);

  const tabs = Array.from(document.querySelectorAll('#langTabsWrapper .lang-tab')).map(t => t.textContent.replace(/\s+/g, ' ').trim());
  assert.ok(tabs[0].startsWith('EN') && tabs.some(t => t.startsWith('CS ')), 'Záložky ukazují názvy jazyků');
  assert.ok(document.querySelector('#langTabsWrapper .lang-tab.active').textContent.includes('CS'));

  const cards = document.querySelectorAll('#articlesContainer .article-card');
  assert.strictEqual(cards.length, 15);
  const card4 = document.getElementById('article-card-3');
  const header = card4.querySelector('.article-header');
  assert.strictEqual(header.querySelector('.article-id').textContent, 'Řádek 5');
  assert.ok(!header.querySelector('.article-section'), 'Žádný nadpis odstavce');
  assert.ok(!header.textContent.includes('Geschenke'), 'Text buňky se nad editorem neopakuje');

  const panels = card4.querySelectorAll('.sub-panel');
  assert.strictEqual(panels.length, 4);
  assert.ok(panels[0].textContent.includes('Reference (DE 2026 (Referenz))'));
  assert.ok(panels[1].textContent.includes('Stará verze – CS') && panels[1].textContent.includes('35 eur'));
  assert.ok(panels[2].querySelector('textarea.target-textarea') && panels[2].textContent.includes('Nová verze – CS'));
  assert.ok(panels[3].querySelector('.live-target-diff').textContent.includes('Zatím bez nové verze'));
});

runTest('Studio: živý diff při psaní a předvyplňovací tlačítka', () => {
  const { window, document } = createDomEnvironment();
  openCzechStudio(window);
  const card4 = document.getElementById('article-card-3');
  const textarea = card4.querySelector('textarea.target-textarea');
  const liveDiff = card4.querySelector('.live-target-diff');

  typeInto(window, textarea, 'Dary a pozvání lze přijmout pouze tehdy, pokud jejich hodnota nepřesahuje 50 eur a byly předem schváleny přes portál compliance.');
  assert.ok(liveDiff.innerHTML.includes('<ins class="diff-ins"') && liveDiff.innerHTML.includes('50'));
  assert.ok(liveDiff.innerHTML.includes('<del class="diff-del"') && liveDiff.innerHTML.includes('35'));
  assert.strictEqual(window.__app.state.translations['CS 2024'][3].status, 'draft');

  card4.querySelector('.prefill-old-btn').click();
  assert.ok(textarea.value.includes('35 eur'));

  const card8 = document.getElementById('article-card-7');
  card8.querySelector('.prefill-ref-btn').click();
  assert.ok(card8.querySelector('textarea').value.includes('Künstlicher Intelligenz'));
});

runTest('Copilot prompt: obecný text, číslo řádku, reference i stará verze jazyka', () => {
  const { window, document } = createDomEnvironment();
  openCzechStudio(window);
  document.getElementById('article-card-3').querySelector('.copy-prompt-btn').click();

  const prompt = window.getLastCopiedText();
  ['Odstavec: #5', 'CS', '35 Euro', '50 Euro', '35 eur', 'POUZE minimální nutné změny', 'POUZE vraťte výsledný text']
    .forEach(part => assert.ok(prompt.includes(part), 'Chybí: ' + part));
  assert.ok(!/etick|kodex|němč/i.test(prompt.split('=== 1.')[0]), 'Instrukce nejsou vázané na druh dokumentu');
  assert.ok(prompt.length < 1600, `Prompt je stručný (${prompt.length} znaků)`);
  assert.ok(document.querySelector('#toastContainer .toast'));
});

runTest('Vlastní šablona promptu se použije a uloží', async () => {
  const { window, document } = createDomEnvironment();
  openCzechStudio(window);
  document.getElementById('settingsPromptBtn').click();
  const input = document.getElementById('promptTemplateInput');
  assert.ok(input.value.includes('{TARGET_LANG}') && !/etick|němč/i.test(input.value), 'Výchozí šablona je obecná');
  input.value = 'AI: {TARGET_LANG} | {ROW_ID} | {REF_LANG} | {REF_NEW}';
  document.getElementById('savePromptTemplateBtn').click();

  document.getElementById('article-card-3').querySelector('.copy-prompt-btn').click();
  const prompt = window.getLastCopiedText();
  assert.ok(prompt.startsWith('AI: CS | #5 | DE 2026 (Referenz) | '), prompt.slice(0, 60));
  assert.ok(prompt.includes('50 Euro'));
});

runTest('Schvalování: třídy karty a progress bar', () => {
  const { window, document } = createDomEnvironment();
  openCzechStudio(window);
  const card4 = document.getElementById('article-card-3');
  const bar = document.getElementById('currentLangProgressBar');
  const before = bar.style.width;
  const cb = card4.querySelector('.verify-checkbox');
  cb.checked = true;
  cb.dispatchEvent(new window.Event('change'));
  assert.ok(card4.classList.contains('status-verified') && !card4.classList.contains('status-pending'));
  assert.notStrictEqual(bar.style.width, before);
  assert.ok(document.getElementById('currentLangProgressText').textContent.includes('hotovo'));
});

// Uložení / načtení projektu tlačítky v záhlaví
runTest('Projekt: tlačítka v záhlaví (i v Kroku 3) uloží a načtou kompletní stav', async () => {
  const env1 = createDomEnvironment();
  const d1 = env1.document;
  d1.getElementById('loadSampleBtn').click();
  const label = langRow(d1, 'CS 2024').querySelector('.lang-label-input');
  label.value = 'Čeština';
  label.dispatchEvent(new env1.window.Event('input'));
  d1.getElementById('confirmMappingBtn').click();
  d1.getElementById('goToTranslationBtn').click();
  Array.from(d1.querySelectorAll('#langTabsWrapper .lang-tab')).find(t => t.textContent.includes('Čeština')).click();

  const card4 = d1.getElementById('article-card-3');
  typeInto(env1.window, card4.querySelector('textarea.target-textarea'), 'MŮJ TEXT ČLÁNKU 4');
  const cb = card4.querySelector('.verify-checkbox');
  cb.checked = true;
  cb.dispatchEvent(new env1.window.Event('change'));
  env1.window.__app.state.pdfTemplate = { docTitle: 'Můj dokument' };

  // Uložení z Kroku 3 tlačítkem v záhlaví
  assert.strictEqual(d1.getElementById('step3Section').style.display, 'block');
  d1.getElementById('saveProjectBtn').click();
  const download = env1.window.__downloads.pop();
  assert.ok(download && /_projekt_\d{4}-\d{2}-\d{2}\.json$/.test(download.name), 'Stáhne se soubor projektu');
  assert.ok(toastTexts(d1).some(t => t.includes(download.name)), 'Toast uvádí název souboru');
  const json = await readBlobText(env1.window, download.blob);

  // Načtení tlačítkem v záhlaví: otevře výběr souboru
  const env2 = createDomEnvironment();
  let pickerOpened = false;
  env2.document.getElementById('projectFileInput').click = () => { pickerOpened = true; };
  env2.document.getElementById('loadProjectBtn').click();
  assert.ok(pickerOpened, 'Tlačítko Načíst projekt otevře výběr souboru');
  env2.window.__app.loadProjectFile(new env2.window.File([json], download.name, { type: 'application/json' }));
  await sleep(50);

  assert.strictEqual(env2.window.__lastAlert, undefined, String(env2.window.__lastAlert));
  assert.strictEqual(env2.document.getElementById('step3Section').style.display, 'block');
  const active = env2.document.querySelector('#langTabsWrapper .lang-tab.active');
  assert.ok(active.textContent.includes('Čeština'), 'Název jazyka i aktivní jazyk jsou obnoveny');
  const restored = env2.document.getElementById('article-card-3');
  assert.strictEqual(restored.querySelector('textarea').value, 'MŮJ TEXT ČLÁNKU 4');
  assert.strictEqual(restored.querySelector('.verify-checkbox').checked, true);
  assert.deepStrictEqual(plain(env2.window.__app.state.pdfTemplate), { docTitle: 'Můj dokument' });
});

// Načítání souborů projektu: chybné a neočekávané soubory
async function loadProject(window, content, name) {
  window.__lastAlert = undefined;
  window.__app.loadProjectFile(new window.File([content], name || 'projekt.json', { type: 'application/json' }));
  await sleep(50);
}

runTest('Načtení projektu: JSON s jinou strukturou je odmítnut s varováním a rozpracovaná práce zůstane', async () => {
  const { window, document } = createDomEnvironment();
  openCzechStudio(window);
  typeInto(window, document.getElementById('article-card-3').querySelector('textarea.target-textarea'), 'Moje rozpracovaná práce');
  window.__lastConfirm = undefined;

  const cases = [
    ['{"hello":"Ahoj","bye":"Nashle"}', 'preklady.json', /preklady\.json.*jinou strukturou.*nejde o projekt překladu.*nalezeno: hello, bye/],
    ['[1, 2, 3]', 'seznam.json', /jinou strukturou/],
    ['{"app":"NejakaJinaAplikace","rows":[]}', 'jiny.json', /patří jiné aplikaci \(NejakaJinaAplikace\)/],
    ['{ tohle není json', 'rozbity.json', /není platný soubor projektu/],
    ['{"rows":[1,2],"mappings":{}}', 'poskozeny.json', /poškozený: řádek 1/]
  ];
  for (const [content, name, pattern] of cases) {
    await loadProject(window, content, name);
    assert.ok(pattern.test(window.__lastAlert), name + ': ' + window.__lastAlert);
    assert.ok(window.__lastAlert.endsWith('Aktuální práce zůstala beze změny.'), name);
  }
  assert.strictEqual(window.__lastConfirm, undefined, 'U neplatného souboru se na nahrazení práce vůbec neptá');
  assert.strictEqual(document.getElementById('step3Section').style.display, 'block', 'Zůstává Krok 3');
  assert.strictEqual(document.getElementById('article-card-3').querySelector('textarea').value, 'Moje rozpracovaná práce');
  assert.strictEqual(window.__app.state.translations['CS 2024'][3].targetNew, 'Moje rozpracovaná práce');
});

runTest('Načtení projektu: omylem vybraná tabulka dostane radu, kde ji nahrát (i v angličtině)', async () => {
  const { window, document } = createDomEnvironment();
  const bytes = fs.readFileSync(path.join(__dirname, '..', 'sample', 'ukazkovy_dokument.xlsx'));
  await loadProject(window, bytes, 'tabulka.xlsx');
  assert.ok(/tabulka\.xlsx.*je tabulka, ne soubor projektu.*Kroku 1/.test(window.__lastAlert), window.__lastAlert);

  document.getElementById('langToggleBtn').click();
  await loadProject(window, '{}', 'empty.json');
  assert.ok(/"empty\.json" is JSON with a different structure, not a translation project/.test(window.__lastAlert), window.__lastAlert);
  assert.ok(window.__lastAlert.endsWith('Your current work is unchanged.'));
});

runTest('Načtení projektu: nahrazení rozpracované práce je třeba potvrdit', async () => {
  const env1 = createDomEnvironment();
  openCzechStudio(env1.window);
  env1.window.__app.state.pdfTemplate = { docTitle: 'Projekt A' };
  env1.document.getElementById('saveProjectBtn').click();
  const projectA = await readBlobText(env1.window, env1.window.__downloads.pop().blob);

  const { window, document } = createDomEnvironment();
  openCzechStudio(window);
  typeInto(window, document.getElementById('article-card-3').querySelector('textarea.target-textarea'), 'Neuložená práce B');

  window.__confirmAnswer = false;
  await loadProject(window, projectA, 'projekt_A.json');
  assert.ok(/Načtením projektu „projekt_A\.json“ se nahradí aktuální rozpracovaná práce/.test(window.__lastConfirm));
  assert.strictEqual(document.getElementById('article-card-3').querySelector('textarea').value, 'Neuložená práce B', 'Po odmítnutí zůstane práce B');

  window.__confirmAnswer = true;
  await loadProject(window, projectA, 'projekt_A.json');
  assert.deepStrictEqual(plain(window.__app.state.pdfTemplate), { docTitle: 'Projekt A' }, 'Po potvrzení je načten projekt A');
  assert.strictEqual(window.__app.state.translations['CS 2024'][3].targetNew, '', 'Neuložená práce B byla nahrazena');
  assert.strictEqual(document.getElementById('step2Section').style.display, 'block', 'Projekt bez úprav se otevře v Kroku 2');
  assert.ok(toastTexts(document).some(t => t.includes('Projekt byl načten.')));
});

runTest('Načtení projektu: nesrovnalosti, které nebrání načtení, se zobrazí jako upozornění', async () => {
  const env1 = createDomEnvironment();
  openCzechStudio(env1.window);
  env1.document.getElementById('saveProjectBtn').click();
  const saved = JSON.parse(await readBlobText(env1.window, env1.window.__downloads.pop().blob));
  saved.version = '99.0.0';
  saved.translations['CS 2024'] = saved.translations['CS 2024'].slice(0, 10);
  saved.translations['Bereich'] = [{ targetNew: 'něco' }];

  const { window, document } = createDomEnvironment();
  await loadProject(window, JSON.stringify(saved), 'upraveny.json');
  assert.strictEqual(window.__lastAlert, undefined, String(window.__lastAlert));
  assert.ok(toastTexts(document).some(t => t.includes('Projekt byl načten.')));
  const warning = toastTexts(document).find(t => t.includes('načten s upozorněním'));
  assert.ok(warning, 'Zobrazí se upozornění');
  assert.ok(warning.includes('novější verze aplikace (99.0.0)'));
  assert.ok(warning.includes('jazyk CS má uloženo 10 textů, ale tabulka má 15 řádků'));
  assert.ok(warning.includes('(Bereich)'));
  assert.strictEqual(document.getElementById('step2Section').style.display === 'block' || document.getElementById('step3Section').style.display === 'block', true);
});

runTest('Načtení projektu: soubor .json přetažený do oblasti pro nahrání v Kroku 1 se otevře jako projekt', async () => {
  const env1 = createDomEnvironment();
  openCzechStudio(env1.window);
  typeInto(env1.window, env1.document.getElementById('article-card-3').querySelector('textarea.target-textarea'), 'Text z přetaženého projektu');
  env1.document.getElementById('saveProjectBtn').click();
  const json = await readBlobText(env1.window, env1.window.__downloads.pop().blob);

  const { window, document } = createDomEnvironment();
  window.__app.processUploadedFile(new window.File([json], 'projekt.json', { type: 'application/json' }));
  await sleep(50);
  assert.strictEqual(window.__lastAlert, undefined, String(window.__lastAlert));
  assert.strictEqual(document.getElementById('step3Section').style.display, 'block');
  assert.strictEqual(document.getElementById('article-card-3').querySelector('textarea').value, 'Text z přetaženého projektu');
});

// Dokument pro tisk
runTest('Dokument pro tisk: vybraný jazyk, upravený text, koncept, žádné nadpisy', async () => {
  const { window, document } = createDomEnvironment();
  openCzechStudio(window);
  typeInto(window, document.getElementById('article-card-3').querySelector('textarea.target-textarea'), 'Finální text odstavce 4.');
  document.getElementById('goToExportBtn').click();

  const select = document.getElementById('pdfLangSelect');
  const options = Array.from(select.options).map(o => o.textContent);
  assert.strictEqual(options[0], 'DE 2026 (Referenz) (reference – nová verze)');
  assert.ok(options.includes('CS'));
  assert.strictEqual(select.value, 'CS 2024', 'Předvybrán aktivní jazyk');

  document.getElementById('printPdfBtn').click();
  const win = window.__printWindows.pop();
  await sleep(350);
  assert.ok(win && win.printed);
  const html = win.html;
  assert.ok(html.includes('Finální text odstavce 4.'));
  assert.ok(html.includes('KONCEPT'));
  assert.ok(!html.includes(window.__app.state.rows[6]['CS 2024']), 'Zastaralý text se nepoužije');
  assert.ok(!html.includes('Art. 0') && !html.includes('Telefax'), 'Bez nadpisů ze sloupců a bez zrušených odstavců');
  assert.ok(html.includes('<div class="doc-title">ukazkovy_dokument</div>'));

  document.getElementById('downloadPdfHtmlBtn').click();
  assert.strictEqual(window.__downloads.pop().name, 'ukazkovy_dokument_CS.html');
});

// Excel exporty
runTest('Excel: kompletní a nová verze tlačítky v Kroku 4', async () => {
  const { window, document } = createDomEnvironment();
  const XLSX = window.XLSX;
  openCzechStudio(window);
  const card4 = document.getElementById('article-card-3');
  typeInto(window, card4.querySelector('textarea.target-textarea'), 'Dary do 50 EUR.');
  const cb = card4.querySelector('.verify-checkbox');
  cb.checked = true;
  cb.dispatchEvent(new window.Event('change'));
  document.getElementById('goToExportBtn').click();

  document.getElementById('downloadComprehensiveBtn').click();
  const comp = await lastWorkbook(window);
  assert.strictEqual(comp.name, 'ukazkovy_dokument_srovnani.xlsx');
  assert.deepStrictEqual(plain(comp.wb.SheetNames), ['Srovnani', 'Stav_prekladu']);
  const header = XLSX.utils.sheet_to_json(comp.wb.Sheets.Srovnani, { header: 1 })[0];
  assert.strictEqual(header.indexOf('CS 2026'), header.indexOf('CS 2024') + 1, 'Nová verze hned za starou');
  const rows = XLSX.utils.sheet_to_json(comp.wb.Sheets.Srovnani, { defval: '' });
  assert.strictEqual(rows[3]['CS 2026'], 'Dary do 50 EUR.');
  assert.strictEqual(rows[3]['CS 2024'], window.__app.state.rows[3]['CS 2024'], 'Stará verze zůstává');
  const status = XLSX.utils.sheet_to_json(comp.wb.Sheets.Stav_prekladu, { defval: '' });
  assert.strictEqual(status[3]['Řádek'], 5);
  assert.strictEqual(status[3].CS, 'Ověřeno');
  assert.ok(toastTexts(document).some(t => /Pozor: \d+ odstavců/.test(t)));

  document.getElementById('downloadCleanBtn').click();
  const clean = await lastWorkbook(window);
  assert.strictEqual(clean.name, 'ukazkovy_dokument_nova_verze.xlsx');
  const cleanRows = XLSX.utils.sheet_to_json(clean.wb.Sheets[clean.wb.SheetNames[0]], { defval: '' });
  assert.strictEqual(cleanRows.length, 13);
  assert.ok(!('CS 2024' in cleanRows[0]) && cleanRows.some(r => r['CS 2026'] === 'Dary do 50 EUR.'));
});

runTest('Excel: částečný sešit – nové verze ostatních jazyků (přednostně) i původní sloupce jako další sloupce', async () => {
  const { window, document } = createDomEnvironment();
  const XLSX = window.XLSX;
  openCzechStudio(window);
  typeInto(window, document.getElementById('article-card-3').querySelector('textarea.target-textarea'), 'Nový text 4.');
  // Angličtina má rozpracovaný nový překlad stejného odstavce
  Array.from(document.querySelectorAll('#langTabsWrapper .lang-tab')).find(t => t.querySelector('span').textContent === 'EN').click();
  typeInto(window, document.getElementById('article-card-3').querySelector('textarea.target-textarea'), 'Gifts up to EUR 50.');
  document.getElementById('goToExportBtn').click();
  const select = document.getElementById('pdfLangSelect');
  select.value = 'CS 2024';
  select.dispatchEvent(new window.Event('change'));

  const context = document.getElementById('partialContextContainer');
  const groups = Array.from(context.querySelectorAll('.partial-context-group'));
  assert.deepStrictEqual(groups.map(gr => gr.querySelector('.partial-context-title').textContent),
    ['Nové verze (aktuální stav překladu)', 'Původní sloupce tabulky'], 'Nové verze jsou uvedeny jako první');
  const newLabels = Array.from(groups[0].querySelectorAll('label')).map(l => l.textContent);
  assert.strictEqual(newLabels[0], 'DE 2026 (Referenz)');
  assert.ok(newLabels.includes('EN 2026 (EN)') && newLabels.includes('FR 2026 (FR)'), 'Nové sloupce všech ostatních jazyků');
  assert.ok(!newLabels.some(l => l.startsWith('CS 2026')), 'Vlastní nová verze exportovaného jazyka není mezi dalšími sloupci');
  const originalLabels = Array.from(groups[1].querySelectorAll('label')).map(l => l.textContent);
  assert.ok(originalLabels.includes('EN 2024') && originalLabels.includes('DE 2024 (Referenz)') && !originalLabels.includes('CS 2024'));

  const boxes = Array.from(context.querySelectorAll('input[type="checkbox"]'));
  assert.deepStrictEqual(boxes.filter(bx => bx.checked).map(bx => bx.value).sort(), ['col:DE 2024 (Referenz)', 'col:DE 2026 (Referenz)'], 'Výchozí: stará a nová reference');

  // Uživatel přidá novou verzi angličtiny a původní angličtinu, odebere starou referenci
  const tick = (value, checked) => {
    const bx = boxes.find(x => x.value === value);
    bx.checked = checked;
    bx.dispatchEvent(new window.Event('change'));
  };
  tick('col:EN 2024', true);
  tick('new:EN 2024', true);
  tick('col:DE 2024 (Referenz)', false);

  document.getElementById('downloadPartialBtn').click();
  const out = await lastWorkbook(window);
  assert.strictEqual(out.name, 'ukazkovy_dokument_CS.xlsx');
  const sheet = out.wb.Sheets[out.wb.SheetNames[0]];
  assert.deepStrictEqual(plain(XLSX.utils.sheet_to_json(sheet, { header: 1 })[0]),
    ['Řádek', 'DE 2026 (Referenz)', 'EN 2026', 'EN 2024', 'CS 2024', 'CS 2026', 'Stav'], 'Nové verze před původními sloupci');
  const data = XLSX.utils.sheet_to_json(sheet, { defval: '' });
  assert.strictEqual(data.length, 15);
  assert.strictEqual(data[3]['Řádek'], 5);
  assert.strictEqual(data[3]['EN 2026'], 'Gifts up to EUR 50.', 'Nová verze jiného jazyka = aktuální překlad');
  assert.strictEqual(data[0]['EN 2026'], data[0]['EN 2024'], 'Nezměněný odstavec převzat');
  assert.strictEqual(data[3]['CS 2026'], 'Nový text 4.');
  assert.strictEqual(data[3].Stav, 'Čeká na kontrolu');

  // Jiný jazyk ve výběru → kontext se přepočítá a výběr zůstane uložen
  select.value = 'FR 2024';
  select.dispatchEvent(new window.Event('change'));
  const values = Array.from(context.querySelectorAll('input')).map(bx => bx.value);
  assert.ok(values.includes('new:CS 2024') && values.includes('col:CS 2024'));
  assert.ok(Array.from(context.querySelectorAll('input')).find(bx => bx.value === 'new:EN 2024').checked);
});

// Náhled dat
runTest('Náhled dat: sloupec Přeskočit, čísla řádků, sloupce v pořadí souboru a postupné zobrazení dalších řádků', () => {
  const { document } = createDomEnvironment();
  const container = document.getElementById('dataPreviewContainer');
  assert.strictEqual(container.style.display, 'none');
  document.getElementById('loadSampleBtn').click();
  assert.strictEqual(container.style.display, 'block');
  const head = Array.from(document.querySelectorAll('#previewTableHead th')).map(th => th.textContent);
  assert.deepStrictEqual(head.slice(0, 4), ['Přeskočit', 'Řádek', 'Artikel-Nr.', 'Bereich']);
  const firstRow = document.querySelectorAll('#previewTableBody tr')[0].querySelectorAll('td');
  assert.ok(firstRow[0].querySelector('.preview-skip-cb') && !firstRow[0].querySelector('.preview-skip-cb').checked);
  assert.strictEqual(firstRow[1].textContent, '2');
  assert.strictEqual(document.querySelectorAll('#previewTableBody tr').length, 10, 'Na začátku se ukáže 10 řádků');

  const more = document.getElementById('previewShowMoreBtn');
  assert.strictEqual(more.textContent, 'Zobrazit dalších 5 řádků (zbývá 5)');
  more.click();
  assert.strictEqual(document.querySelectorAll('#previewTableBody tr').length, 15);
  assert.strictEqual(more.style.display, 'none');
});

runTest('Přeskočení řádků: řádek se nezpracuje, ale v kompletním sešitu zůstane beze změny', async () => {
  const { window, document } = createDomEnvironment();
  const XLSX = window.XLSX;
  document.getElementById('loadSampleBtn').click();

  // Řádek 4 (Art. 03, upravený) uživatel přeskočí
  const cb = document.querySelector('.preview-skip-cb[data-row="4"]');
  cb.checked = true;
  cb.dispatchEvent(new window.Event('change'));
  const st = window.__app.state;
  assert.strictEqual(st.rows.length, 14);
  assert.deepStrictEqual(plain(st.skippedRows.map(r => r.rowNumber)), [4]);
  assert.ok(document.querySelector('#previewTableBody tr[data-row="4"]').classList.contains('preview-row-skipped'));
  assert.ok(document.getElementById('previewRowsCount').textContent.includes('14 řádků · 1 přeskočeno'));

  document.getElementById('confirmMappingBtn').click();
  assert.strictEqual(document.getElementById('kpiTotalRows').textContent, '14', 'Statistika bez přeskočeného řádku');
  assert.strictEqual(document.getElementById('kpiModified').textContent, '5');
  document.getElementById('goToTranslationBtn').click();
  const labels = Array.from(document.querySelectorAll('#articlesContainer .article-id')).map(e => e.textContent);
  assert.ok(!labels.includes('Řádek 4') && labels.includes('Řádek 5'), 'Přeskočený řádek není v překladu');

  document.getElementById('goToExportBtn').click();
  document.getElementById('downloadComprehensiveBtn').click();
  const comp = await lastWorkbook(window);
  const rows = XLSX.utils.sheet_to_json(comp.wb.Sheets.Srovnani, { defval: '' });
  assert.strictEqual(rows.length, 15, 'Kompletní sešit má všechny řádky');
  assert.strictEqual(rows[2]['Artikel-Nr.'], 'Art. 03', 'Přeskočený řádek je na svém místě');
  assert.strictEqual(rows[2]['CS 2024'], SampleData()[2]['CS 2024'], 'a beze změny');
  assert.strictEqual(rows[2]['CS 2026'], '', 'Nová verze přeskočeného řádku zůstane prázdná');
  assert.strictEqual(rows[3]['CS 2026'], '', 'Další řádek je na správném místě');
  const status = XLSX.utils.sheet_to_json(comp.wb.Sheets.Stav_prekladu, { defval: '' });
  assert.deepStrictEqual([status[2]['Řádek'], status[2]['Změna reference'], status[2].CS], [4, 'Přeskočeno', 'Přeskočeno']);

  document.getElementById('downloadCleanBtn').click();
  const clean = XLSX.utils.sheet_to_json((await lastWorkbook(window)).wb.Sheets.Nova_verze, { defval: '' });
  assert.ok(!clean.some(r => r['Artikel-Nr.'] === 'Art. 03'), 'Nová verze přeskočený řádek neobsahuje');

  window.__app.setStep(4);
  document.getElementById('downloadPartialBtn').click();
  const partial = XLSX.utils.sheet_to_json((await lastWorkbook(window)).wb.Sheets.EN, { defval: '' });
  assert.strictEqual(partial.length, 14);
});

runTest('Přeskočení řádků: změna výběru později nepřesune rozpracované texty k jiným řádkům', () => {
  const { window, document } = createDomEnvironment();
  openCzechStudio(window);
  typeInto(window, document.getElementById('article-card-3').querySelector('textarea.target-textarea'), 'Text řádku 5');

  document.getElementById('stepNav1').click();
  const cb = document.querySelector('.preview-skip-cb[data-row="3"]');
  cb.checked = true;
  cb.dispatchEvent(new window.Event('change'));
  document.getElementById('stepNav3').click();
  const card = Array.from(document.querySelectorAll('#articlesContainer .article-card'))
    .find(c => c.querySelector('.article-id').textContent === 'Řádek 5');
  assert.strictEqual(card.querySelector('textarea').value, 'Text řádku 5', 'Text zůstal u řádku 5');

  // Zrušení přeskočení vrátí řádek 3 (s výchozím stavem) a text zůstane u řádku 5
  document.getElementById('stepNav1').click();
  const cb2 = document.querySelector('.preview-skip-cb[data-row="3"]');
  cb2.checked = false;
  cb2.dispatchEvent(new window.Event('change'));
  document.getElementById('stepNav3').click();
  assert.strictEqual(window.__app.state.rows.length, 15);
  assert.strictEqual(document.getElementById('article-card-3').querySelector('.article-id').textContent, 'Řádek 5');
  assert.strictEqual(document.getElementById('article-card-3').querySelector('textarea').value, 'Text řádku 5');
  assert.strictEqual(window.__app.state.translations['CS 2024'][1].targetNew, window.__app.state.rows[1]['CS 2024'], 'Vrácený nezměněný řádek má výchozí převzatý text');
});

runTest('Prázdné řádky se přeskočí automaticky a nelze je zapnout; projekt si přeskočené řádky pamatuje', async () => {
  const { window, document } = createDomEnvironment();
  window.__app.processUploadedFile(workbookFile(window, {
    Data: [
      ['EN old', 'EN new', 'CZ'],
      ['First paragraph text.', 'First paragraph text, new.', 'Text prvního odstavce.'],
      ['   ', '', ' '],
      ['Second paragraph text.', 'Second paragraph text.', 'Text druhého odstavce.'],
      ['NOTES – not for translation', '', '']
    ]
  }, 'texty.xlsx'));
  await sleep(100);
  const st = window.__app.state;
  assert.deepStrictEqual(plain(st.rowNumbers), [2, 4, 5]);
  assert.deepStrictEqual(plain(st.skippedRows.map(r => [r.rowNumber, r.auto])), [[3, true]]);
  assert.ok(toastTexts(document).some(t => t.includes('Prázdné řádky (1) byly automaticky přeskočeny')));
  const emptyCb = document.querySelector('.preview-skip-cb[data-row="3"]');
  assert.ok(emptyCb.checked && emptyCb.disabled, 'Prázdný řádek je přeskočen a nelze ho zapnout');
  assert.ok(document.querySelector('#previewTableBody tr[data-row="3"]').textContent.includes('prázdný'));

  // Řádek s poznámkou přeskočí uživatel
  const notesCb = document.querySelector('.preview-skip-cb[data-row="5"]');
  notesCb.checked = true;
  notesCb.dispatchEvent(new window.Event('change'));
  assert.deepStrictEqual(plain(st.rowNumbers), [2, 4]);

  const set = (id, v) => { const el = document.getElementById(id); el.value = v; el.onchange(); };
  set('refOldSelect', 'EN old');
  set('refNewSelect', 'EN new');
  document.getElementById('saveProjectBtn').click();
  const json = await readBlobText(window, window.__downloads.pop().blob);

  const env2 = createDomEnvironment();
  env2.window.__app.loadProjectFile(new env2.window.File([json], 'texty_projekt.json'));
  await sleep(50);
  assert.strictEqual(env2.window.__lastAlert, undefined, String(env2.window.__lastAlert));
  assert.deepStrictEqual(plain(env2.window.__app.state.skippedRows.map(r => [r.rowNumber, r.auto])), [[3, true], [5, false]]);
  env2.document.getElementById('stepNav1').click();
  assert.ok(env2.document.querySelector('.preview-skip-cb[data-row="5"]').checked, 'Po načtení projektu je řádek stále přeskočen');
});

runTest('Matice jazyků a hromadné schválení nezměněných', () => {
  const { window, document } = createDomEnvironment();
  openCzechStudio(window);
  const grid = document.getElementById('crossLangMatrixGrid');
  assert.strictEqual(grid.querySelectorAll('.matrix-pill').length, 10);

  const cs = window.__app.state.translations['CS 2024'];
  const en = window.__app.state.translations['EN 2024'];
  cs[0].isVerified = false;
  en[1].isVerified = false;
  document.getElementById('approveAllUnchangedGlobalBtn').click();
  assert.ok(cs.filter(s => s.classification === 'UNCHANGED').every(s => s.isVerified));
  assert.strictEqual(en[1].isVerified, true);
  assert.ok(toastTexts(document).some(t => t.includes('Schváleno 2 nezměněných')));
  const csPill = Array.from(grid.querySelectorAll('.matrix-pill')).find(p => p.querySelector('.matrix-pill-lang').textContent === 'CS');
  assert.ok(csPill.textContent.includes('7/15'));
});

runTest('Ctrl+Enter schválí odstavec', () => {
  const { window, document } = createDomEnvironment();
  openCzechStudio(window);
  const cs = window.__app.state.translations['CS 2024'];
  const card4 = document.getElementById('article-card-3');
  const ta = card4.querySelector('.target-textarea');
  typeInto(window, ta, 'Text potvrzený klávesou.');
  assert.strictEqual(cs[3].isVerified, false);
  ta.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'Enter', ctrlKey: true, bubbles: true, cancelable: true }));
  assert.strictEqual(cs[3].isVerified, true);
  assert.ok(card4.classList.contains('status-verified') && card4.querySelector('.verify-checkbox').checked);
});

runTest('Hromadný prompt a Smart Paste se značkami řádků', () => {
  const { window, document } = createDomEnvironment();
  openCzechStudio(window);

  document.getElementById('batchPromptBtn').click();
  const prompt = document.getElementById('batchPromptTextarea').value;
  assert.ok(prompt.includes('HROMADNÁ AKTUALIZACE PŘEKLADU'));
  assert.ok(prompt.includes('[#5]') && prompt.includes('[#9]'), 'Značky odpovídají číslům řádků');
  document.getElementById('closeBatchPromptBtn').click();

  document.getElementById('smartPasteBtn').click();
  const before = window.__app.state.translations['CS 2024'][0].targetNew;
  document.getElementById('smartPasteInput').value = '[#5]\nDary jsou omezeny na 50 EUR ročně.\n\n[#7]\nOchrana údajů podle AI 2026.\n\n[#99]\nNeexistující odstavec.';
  document.getElementById('smartPasteParseBtn').click();
  const container = document.getElementById('smartPasteMatchesContainer');
  const badges = Array.from(container.querySelectorAll('.matched-article-badge')).map(b => b.textContent);
  assert.strictEqual(badges.length, 2);
  assert.ok(badges[0].includes('#5') && badges[1].includes('#7'));
  assert.ok(container.querySelector('.smart-paste-unmatched').textContent.includes('#99'));

  document.getElementById('smartPasteApplyBtn').click();
  const cs = window.__app.state.translations['CS 2024'];
  assert.ok(cs[3].targetNew.includes('50 EUR ročně') && cs[3].isVerified);
  assert.ok(cs[5].targetNew.includes('AI 2026'));
  assert.strictEqual(cs[0].targetNew, before);
  assert.ok(!cs.some(s => s.targetNew.includes('Neexistující')));
  assert.ok(toastTexts(document).some(t => t.includes('Vloženo 2 odstavců')));
});

runTest('Shrnutí změn: Krok 2 pro referenci, Krok 4 pro vybraný jazyk', async () => {
  const { window, document } = createDomEnvironment();
  document.getElementById('loadSampleBtn').click();
  document.getElementById('confirmMappingBtn').click();

  document.getElementById('step2ExportMdBtn').click();
  const refMd = window.__downloads.pop();
  assert.strictEqual(refMd.name, 'ukazkovy_dokument_zmeny_DE_2026_Referenz.md');
  const refText = await readBlobText(window, refMd.blob);
  assert.ok(refText.includes('# Shrnutí změn – DE 2026 (Referenz)') && refText.includes('| **Celkem odstavců** | 15 | 100% |'));

  document.getElementById('goToTranslationBtn').click();
  Array.from(document.querySelectorAll('#langTabsWrapper .lang-tab')).find(t => t.textContent.includes('CS')).click();
  typeInto(window, document.getElementById('article-card-3').querySelector('textarea.target-textarea'), 'Dary nově do 50 eur.');
  document.getElementById('goToExportBtn').click();

  document.getElementById('downloadMgrMdBtn').click();
  const csMd = await readBlobText(window, window.__downloads.pop().blob);
  assert.ok(csMd.includes('# Shrnutí změn – CS (stará → nová verze)'));
  assert.ok(csMd.includes('Dary nově do 50 eur.'), 'Obsahuje novou verzi jazyka');
  assert.ok(csMd.includes('### Stav překladu'));
  assert.ok(!csMd.includes('Geschenke und Einladungen dürfen nur'), 'Obsahuje jazyk, ne referenci');

  document.getElementById('downloadMgrDocBtn').click();
  const doc = window.__downloads.pop();
  assert.ok(doc.name.endsWith('_CS.doc'));
  assert.ok((await readBlobText(window, doc.blob)).includes('urn:schemas-microsoft-com:office:word'));

  document.getElementById('printMgrPdfBtn').click();
  assert.ok(window.__printWindows.pop().html.includes('Shrnutí změn – CS'));
  assert.ok(!toastTexts(document).some(t => /mgrExport/.test(t)));
});

runTest('Hlavička dokumentu: modální okno bez přednastavených textů, uložení a použití', async () => {
  const { window, document } = createDomEnvironment();
  openCzechStudio(window);
  document.getElementById('goToExportBtn').click();
  document.getElementById('editPdfTemplateBtn').click();

  assert.ok(document.getElementById('pdfTemplateModal').classList.contains('open'));
  assert.strictEqual(document.getElementById('tplCompanyName').value, '');
  assert.strictEqual(document.getElementById('tplDocTitle').value, 'ukazkovy_dokument');
  assert.strictEqual(document.getElementById('tplDocSubtitle').value, '{LANG}');
  assert.strictEqual(document.getElementById('tplShowSignatures').checked, false);

  document.getElementById('tplCompanyName').value = 'Holding a.s.';
  document.getElementById('tplDocTitle').value = 'Provozní řád';
  document.getElementById('tplAccentColor').value = '#7c3aed';
  document.getElementById('tplShowSignatures').checked = true;
  document.getElementById('tplSignDept1').value = 'Ředitel';
  document.getElementById('savePdfTemplateBtn').click();
  assert.ok(!document.getElementById('pdfTemplateModal').classList.contains('open'));

  document.getElementById('printPdfBtn').click();
  const html = window.__printWindows.pop().html;
  ['Holding a.s.', 'Provozní řád', '#7c3aed', 'Ředitel', '<div class="doc-subtitle">CS</div>'].forEach(p => assert.ok(html.includes(p), p));
});

runTest('Autosave: uložení do prohlížeče a obnovení po znovuotevření', async () => {
  const env1 = createDomEnvironment();
  openCzechStudio(env1.window);
  typeInto(env1.window, env1.document.getElementById('article-card-3').querySelector('textarea.target-textarea'), 'Automaticky uložený text');
  assert.strictEqual(env1.document.getElementById('saveStatusText').textContent, 'Ukládám...');
  env1.window.dispatchEvent(new env1.window.Event('pagehide'));
  assert.strictEqual(env1.document.getElementById('saveStatusText').textContent, 'Vše uloženo');
  const stored = env1.window.localStorage.getItem(STORAGE_KEY);
  assert.ok(stored.includes('Automaticky uložený text'));

  const env2 = createDomEnvironment({ acceptResume: true, beforeInit: (w) => w.localStorage.setItem(STORAGE_KEY, stored) });
  await sleep(150);
  assert.strictEqual(env2.document.getElementById('step3Section').style.display, 'block');
  assert.strictEqual(env2.document.getElementById('article-card-3').querySelector('textarea').value, 'Automaticky uložený text');
});

runTest('Autosave pod původním názvem projektu se stále nabídne k obnovení', async () => {
  const env1 = createDomEnvironment();
  openCzechStudio(env1.window);
  typeInto(env1.window, env1.document.getElementById('article-card-3').querySelector('textarea.target-textarea'), 'Starší autosave');
  env1.window.dispatchEvent(new env1.window.Event('pagehide'));
  const stored = env1.window.localStorage.getItem(STORAGE_KEY);

  const env2 = createDomEnvironment({ acceptResume: true, beforeInit: (w) => w.localStorage.setItem('ethic_comparator_workspace_v1', stored) });
  await sleep(150);
  assert.strictEqual(env2.document.getElementById('article-card-3').querySelector('textarea').value, 'Starší autosave');
});

runTest('Adaptéry: nový importér a exportér se použijí bez úprav aplikace', async () => {
  const { window, document } = createDomEnvironment({
    beforeApp: (w) => {
      // Simple key = text importer for .tsv files (one language per column)
      w.Importers.register({
        id: 'test-tsv',
        label: 'TSV',
        extensions: ['.tsv'],
        readAs: 'text',
        parse: (content) => {
          const lines = content.trim().split('\n').map(l => l.split('\t'));
          return {
            sheetNames: ['TSV'],
            readSheet: () => ({
              headers: lines[0],
              rows: lines.slice(1).map(cells => Object.fromEntries(lines[0].map((h, i) => [h, cells[i] || '']))),
              rowNumbers: lines.slice(1).map((l, i) => i + 2)
            })
          };
        }
      });
      // CSV table exporter (first table only)
      w.Exporters.register({
        id: 'test-csv',
        label: 'CSV (test)',
        group: 'table',
        export: (payload) => ({
          content: [payload.tables[0].headers.join(';')].concat(payload.tables[0].rows.map(r => payload.tables[0].headers.map(h => r[h]).join(';'))).join('\n'),
          mimeType: 'text/csv',
          extension: '.csv'
        })
      });
    }
  });

  document.getElementById('dropzone').dispatchEvent(new window.MouseEvent('click'));
  assert.ok(document.getElementById('fileInput').getAttribute('accept').split(',').includes('.tsv'), 'Výběr souboru nabízí formát nového importéru');

  const tsv = 'EN old\tEN new\tCZ\nThe first paragraph text.\tThe first paragraph text, updated.\tText prvního odstavce.\n';
  window.__app.processUploadedFile(new window.File([tsv], 'texty.tsv'));
  await sleep(100);
  assert.strictEqual(window.__lastAlert, undefined, String(window.__lastAlert));
  assert.deepStrictEqual(plain(window.__app.state.headers), ['EN old', 'EN new', 'CZ']);

  const set = (id, v) => { const el = document.getElementById(id); el.value = v; el.onchange(); };
  set('refOldSelect', 'EN old');
  set('refNewSelect', 'EN new');
  document.getElementById('confirmMappingBtn').click();
  document.getElementById('goToTranslationBtn').click();
  document.getElementById('goToExportBtn').click();

  const wrapper = document.getElementById('tableFormatWrapper');
  const select = document.getElementById('tableFormatSelect');
  assert.notStrictEqual(wrapper.style.display, 'none', 'S více tabulkovými formáty se zobrazí výběr formátu');
  assert.deepStrictEqual(Array.from(select.options).map(o => o.value), ['xlsx', 'test-csv']);
  assert.strictEqual(select.value, 'xlsx', 'Výchozí formát je Excel');

  select.value = 'test-csv';
  select.dispatchEvent(new window.Event('change'));
  document.getElementById('downloadCleanBtn').click();
  const download = window.__downloads.pop();
  assert.strictEqual(download.name, 'texty_nova_verze.csv');
  const text = await readBlobText(window, download.blob);
  assert.ok(text.startsWith('EN new;'), text);
});

runTest('Adaptéry: nepodporovaný soubor dá srozumitelnou chybu se seznamem formátů', async () => {
  const { window } = createDomEnvironment();
  window.__app.processUploadedFile(new window.File(['%PDF-1.4'], 'smlouva.pdf'));
  await sleep(50);
  assert.ok(/smlouva\.pdf.*\.xlsx.*\.csv/.test(window.__lastAlert), String(window.__lastAlert));
});

runTest('Autosave: selhání ukládání je viditelně oznámeno', () => {
  const { window, document } = createDomEnvironment();
  openCzechStudio(window);
  const proto = Object.getPrototypeOf(window.localStorage);
  const original = proto.setItem;
  const originalWarn = window.console.warn;
  proto.setItem = function (key, value) {
    if (key === STORAGE_KEY) throw new Error('QuotaExceededError');
    return original.call(this, key, value);
  };
  window.console.warn = () => {};
  try {
    typeInto(window, document.getElementById('article-card-3').querySelector('textarea.target-textarea'), 'x');
    window.StorageManager.flushPendingSave();
  } finally {
    proto.setItem = original;
    window.console.warn = originalWarn;
  }
  assert.ok(document.getElementById('saveIndicator').classList.contains('save-error'));
  assert.ok(toastTexts(document).some(t => t.includes('Uložit projekt')));
});

// Nahrání souborů
runTest('Nahrání .xlsx: žádné hádání jazyků podle názvů – reference vybírá uživatel, jazyky jsou textové sloupce', async () => {
  const { window, document } = createDomEnvironment();
  const bytes = fs.readFileSync(path.join(__dirname, '..', 'sample', 'ukazkovy_dokument.xlsx'));
  window.__app.processUploadedFile(new window.File([bytes], 'data.xlsx'));
  await sleep(100);
  const st = window.__app.state;

  assert.strictEqual(st.rows.length, 15);
  assert.strictEqual(st.mappings.refOldCol, '', 'Reference se nevybírá automaticky');
  assert.strictEqual(document.getElementById('refOldSelect').value, '');
  assert.ok(st.mappings.targetCols.includes('DE 2024 (Referenz)') && !st.mappings.targetCols.includes('Artikel-Nr.') && !st.mappings.targetCols.includes('Bereich'),
    'Předvybrány jsou sloupce s delším textem, ne podle názvu');

  setSelect(window, document.getElementById('refOldSelect'), 'DE 2024 (Referenz)');
  setSelect(window, document.getElementById('refNewSelect'), 'DE 2026 (Referenz)');
  assert.strictEqual(st.mappings.targetCols.length, 10, 'Referenční sloupce zmizely z jazyků');
  document.getElementById('confirmMappingBtn').click();
  assert.strictEqual(document.getElementById('step2Section').style.display, 'block');
});

runTest('Nahrání .xlsx s existujícím (téměř prázdným) sloupcem nové verze: spárování, předvyplnění a export do stejného sloupce', async () => {
  const { window, document } = createDomEnvironment();
  const file = workbookFile(window, {
    List1: [
      ['EN old', 'EN new', 'CZ old', 'CZ new', 'Poznámka'],
      ['Same paragraph text here.', 'Same paragraph text here.', 'Stejný text odstavce.', '', 'x'],
      ['The limit is thirty-five euro.', 'The limit is fifty euro.', 'Limit je třicet pět eur.', 'Limit je padesát eur.', ''],
      ['Removed paragraph of the text.', '', 'Zrušený odstavec textu.', '', '']
    ]
  });
  window.__app.processUploadedFile(file);
  await sleep(100);
  const st = window.__app.state;

  setSelect(window, document.getElementById('refOldSelect'), 'EN old');
  setSelect(window, document.getElementById('refNewSelect'), 'EN new');
  assert.deepStrictEqual(plain(st.mappings.targetCols), ['CZ old']);
  // Bez společného čísla verze uživatel vybere existující sloupec ručně
  setSelect(window, langRow(document, 'CZ old').querySelector('.lang-newcol-select'), 'CZ new');
  assert.ok(langRow(document, 'CZ new').textContent.includes('nová verze jazyka CZ old'));

  document.getElementById('confirmMappingBtn').click();
  document.getElementById('goToTranslationBtn').click();
  const stubs = st.translations['CZ old'];
  assert.strictEqual(stubs[1].targetNew, 'Limit je padesát eur.', 'Text z existujícího sloupce je předvyplněn');
  assert.strictEqual(stubs[1].isVerified, false, 'a čeká na kontrolu');
  assert.strictEqual(document.getElementById('article-card-1').querySelector('textarea').value, 'Limit je padesát eur.');

  document.getElementById('goToExportBtn').click();
  document.getElementById('downloadComprehensiveBtn').click();
  const sheet = (await lastWorkbook(window)).wb.Sheets.Srovnani;
  assert.deepStrictEqual(plain(window.XLSX.utils.sheet_to_json(sheet, { header: 1 })[0]), ['EN old', 'EN new', 'CZ old', 'CZ new', 'Poznámka'], 'Žádný nový sloupec');
  const rows = window.XLSX.utils.sheet_to_json(sheet, { defval: '' });
  assert.strictEqual(rows[0]['CZ new'], 'Stejný text odstavce.', 'Prázdná buňka doplněna');
  assert.strictEqual(rows[1]['CZ new'], 'Limit je padesát eur.');
});

runTest('Nahrání .xlsx: automatické spárování podle čísla verze v referenci (CS 2024 → CS 2026)', async () => {
  const { window, document } = createDomEnvironment();
  window.__app.processUploadedFile(workbookFile(window, {
    Data: [
      ['EN 2024', 'EN 2026', 'CS 2024', 'CS 2026'],
      ['A longer paragraph that stays the same.', 'A longer paragraph that stays the same.', 'Delší odstavec, který zůstává stejný.', ''],
      ['Another long paragraph with old wording.', 'Another long paragraph with new wording.', 'Další dlouhý odstavec se starým zněním.', '']
    ]
  }));
  await sleep(100);
  setSelect(window, document.getElementById('refOldSelect'), 'EN 2024');
  setSelect(window, document.getElementById('refNewSelect'), 'EN 2026');
  const st = window.__app.state;
  assert.deepStrictEqual(plain(st.mappings.targetCols), ['CS 2024']);
  assert.strictEqual(st.mappings.languages['CS 2024'].newCol, 'CS 2026');
  assert.strictEqual(langRow(document, 'CS 2024').querySelector('.lang-label-input').value, 'CS');
});

runTest('List bez řádku záhlaví: sloupce A, B, C…, první řádek jsou data, název jazyka zadá uživatel', async () => {
  const { window, document } = createDomEnvironment();
  window.__app.processUploadedFile(workbookFile(window, {
    Data: [
      ['Old paragraph one text.', 'New paragraph one text.', 'Starý text odstavce jedna.'],
      ['Paragraph two stays as it is.', 'Paragraph two stays as it is.', 'Odstavec dva zůstává.']
    ]
  }));
  await sleep(100);
  const st = window.__app.state;
  assert.strictEqual(st.rows.length, 1, 'S výchozím nastavením je první řádek záhlaví');

  const headerCb = document.getElementById('hasHeaderRowCheckbox');
  headerCb.checked = false;
  headerCb.dispatchEvent(new window.Event('change'));
  assert.strictEqual(st.rows.length, 2, 'První řádek je teď data');
  assert.deepStrictEqual(plain(st.headers), ['A', 'B', 'C']);
  assert.deepStrictEqual(plain(st.rowNumbers), [1, 2]);
  assert.strictEqual(document.getElementById('refOldSelect').options[1].textContent, 'Sloupec A');

  setSelect(window, document.getElementById('refOldSelect'), 'A');
  setSelect(window, document.getElementById('refNewSelect'), 'B');
  const label = langRow(document, 'C').querySelector('.lang-label-input');
  assert.strictEqual(label.value, 'Sloupec C');
  label.value = 'Čeština';
  label.dispatchEvent(new window.Event('input'));
  document.getElementById('confirmMappingBtn').click();
  document.getElementById('goToTranslationBtn').click();
  assert.ok(document.querySelector('#langTabsWrapper .lang-tab').textContent.includes('Čeština'));
  assert.strictEqual(document.getElementById('article-card-0').querySelector('.article-id').textContent, 'Řádek 1');
});

runTest('CSV: UTF-8 diakritika, odstranění bílých znaků na okrajích buněk a zahození předchozí práce', async () => {
  const { window, document } = createDomEnvironment();
  openCzechStudio(window);
  typeInto(window, document.getElementById('article-card-0').querySelector('textarea.target-textarea'), 'STARÁ PRÁCE');

  const longCell = '\n\n   ' + 'Velmi dlouhý text bez zalomení řádku. '.repeat(10).trim() + '   \n';
  const csv = '﻿Ref old;Ref new;Čeština\nGröße alt;Größe neu;' + '"' + longCell + '"' + '\n';
  window.__app.processUploadedFile(new window.File([csv], 'data.csv', { type: 'text/csv' }));
  await sleep(100);

  const st = window.__app.state;
  assert.strictEqual(st.rows[0]['Ref new'], 'Größe neu');
  assert.ok(st.rows[0]['Čeština'].startsWith('Velmi dlouhý text') && st.rows[0]['Čeština'].endsWith('řádku.'), 'Bílé znaky na začátku a konci odstraněny');
  assert.deepStrictEqual(plain(st.translations), {});
  const previewCell = document.querySelectorAll('#previewTableBody tr')[0].querySelectorAll('td')[4];
  assert.ok(previewCell.textContent.startsWith('Velmi dlouhý text'), 'Náhled neukazuje prázdnou buňku');
});

runTest('Navigace: Krok 2/3 se neotevře bez platného výběru sloupců', () => {
  const { window, document } = createDomEnvironment();
  document.getElementById('stepNav3').click();
  assert.strictEqual(document.getElementById('step3Section').style.display, 'none');

  document.getElementById('loadSampleBtn').click();
  setSelect(window, document.getElementById('refNewSelect'), 'DE 2024 (Referenz)');
  document.getElementById('stepNav2').click();
  assert.strictEqual(document.getElementById('step2Section').style.display, 'none');
  assert.ok(toastTexts(document).some(t => t.includes('různé sloupce')));
});

runTest('Modální okna: Escape a klik na pozadí je zavřou', () => {
  const { window, document } = createDomEnvironment();
  const modal = document.getElementById('promptSettingsModal');
  document.getElementById('settingsPromptBtn').click();
  assert.ok(modal.classList.contains('open'));
  document.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
  assert.ok(!modal.classList.contains('open'));
  document.getElementById('settingsPromptBtn').click();
  modal.dispatchEvent(new window.MouseEvent('click', { bubbles: true }));
  assert.ok(!modal.classList.contains('open'));
});

runTest('Pomocný referenční jazyk: výběr zobrazí 5. sloupec, kopírování do editoru a vložení do promptu', async () => {
  const { window, document } = createDomEnvironment();
  openCzechStudio(window);

  const auxSelect = document.getElementById('auxLangSelect');
  assert.ok(auxSelect, 'Selektor pomocného jazyka existuje v panelu nástrojů');

  const options = Array.from(auxSelect.options);
  assert.ok(options.length >= 2, 'Selektor obsahuje volbu žádný a další jazyky');
  assert.strictEqual(options[0].value, '', 'První volba je prázdná');
  assert.ok(!options.some(o => o.value === 'CS 2024'), 'Aktivní cílový jazyk (CS) není nabízen jako pomocný jazyk');
  assert.ok(options.some(o => o.value === 'EN 2024'), 'Jiné jazyky (např. EN) jsou nabízeny');

  // Výchozí stav karet: bez 5. sloupce
  const card0 = document.getElementById('article-card-0');
  const grid0 = card0.querySelector('.article-grid');
  assert.ok(!grid0.classList.contains('has-aux'), 'Výchozí rozložení nemá třídu has-aux');
  assert.ok(!card0.querySelector('.aux-panel'), 'Výchozí karta neobsahuje pomocný panel');

  // Výběr pomocného jazyka EN 2024
  setSelect(window, auxSelect, 'EN 2024');
  assert.strictEqual(window.__app.state.auxiliaryCol, 'EN 2024', 'Stav aplikace zaznamenal pomocný jazyk');

  const card0After = document.getElementById('article-card-0');
  const grid0After = card0After.querySelector('.article-grid');
  assert.ok(grid0After.classList.contains('has-aux'), 'Karta má po výběru pomocného jazyka třídu has-aux');

  const auxPanel = card0After.querySelector('.aux-panel');
  assert.ok(auxPanel, 'Karta nyní obsahuje pomocný panel');
  assert.ok(auxPanel.querySelector('.aux-diff-box'), 'Pomocný panel obsahuje textový diff-box');
  assert.ok(auxPanel.querySelector('.copy-aux-btn'), 'Pomocný panel obsahuje tlačítko pro zkopírování do editoru');

  const auxText = auxPanel.querySelector('.aux-diff-box').textContent;
  assert.ok(auxText.length > 0, 'Pomocný panel obsahuje text daného odstavce');

  // Kopírování do editoru tlačítkem v panelu
  const targetTextarea = card0After.querySelector('textarea.target-textarea');
  const copyAuxBtn = auxPanel.querySelector('.copy-aux-btn');
  copyAuxBtn.click();
  assert.strictEqual(targetTextarea.value, auxText, 'Text z pomocného jazyka byl zkopírován do cílového editoru');
  assert.ok(toastTexts(document).some(t => t.includes('vložen do editoru') || t.includes('inserted into the editor')), 'Toast potvrdil vložení textu');

  // Zkopírování promptu pro Copilot s pomocným jazykem
  const copyPromptBtn = card0After.querySelector('.copy-prompt-btn');
  copyPromptBtn.click();
  await sleep(10);
  const promptText = window.getLastCopiedText();
  assert.ok(promptText.includes('POMOCNÝ PŘEKLAD') || promptText.includes('AUXILIARY TRANSLATION'), 'Copilot prompt obsahuje blok pomocného překladu');
  assert.ok(promptText.includes(auxText.substring(0, 20)), 'Prompt obsahuje text z pomocného jazyka');

  // Přepnutí na záložku EN: pomocný jazyk se automaticky resetuje, protože EN je nově aktivní
  const enTab = Array.from(document.querySelectorAll('#langTabsWrapper .lang-tab'))
    .find(t => t.querySelector('span').textContent === 'EN');
  enTab.click();
  assert.strictEqual(window.__app.state.auxiliaryCol, '', 'Při přepnutí na tab pomocného jazyka se výběr resetuje');
  assert.strictEqual(document.getElementById('auxLangSelect').value, '');
  assert.ok(!document.getElementById('article-card-0').querySelector('.article-grid').classList.contains('has-aux'));

  const enAuxOptions = Array.from(document.getElementById('auxLangSelect').options);
  assert.ok(!enAuxOptions.some(o => o.value === 'EN 2024'), 'Nyní v EN tabu není nabízen EN');
  assert.ok(enAuxOptions.some(o => o.value === 'CS 2024'), 'Naopak CS je nyní nabízen');
});

runTest('Ověření překladu: tlačítko zkopíruje prompt pro kontrolu věrnosti a shodnosti významu do schránky', async () => {
  const { window, document } = createDomEnvironment();
  openCzechStudio(window);

  const card0 = document.getElementById('article-card-0');
  const verifyBtn = card0.querySelector('.copy-verify-prompt-btn');
  assert.ok(verifyBtn, 'Tlačítko ověření překladu existuje na kartě odstavce');
  assert.ok(verifyBtn.textContent.includes('Ověřit překlad'), 'Tlačítko má český popisek pro ověření překladu');

  const textarea = card0.querySelector('textarea.target-textarea');

  // 1. Prázdný překlad: varovný toast
  typeInto(window, textarea, '');
  verifyBtn.click();
  await sleep(10);
  assert.ok(toastTexts(document).some(t => t.includes('Nejprve vložte') || t.includes('Please enter')), 'Toast varuje při prázdném překladu');

  // 2. Vyplněný překlad: zkopírování ověřovacího promptu do schránky
  const myTranslation = 'Schválené nové znění pro odstavec 1.';
  typeInto(window, textarea, myTranslation);
  verifyBtn.click();
  await sleep(10);

  assert.ok(toastTexts(document).some(t => t.includes('Prompt pro ověření překladu byl zkopírován') || t.includes('verification prompt copied')), 'Toast potvrdil zkopírování ověřovacího promptu');

  const promptText = window.getLastCopiedText();
  assert.ok(promptText.includes('Ověřte věcnou a významovou shodnost'), 'Prompt obsahuje požadavek na ověření shodnosti');
  assert.ok(promptText.includes('=== 1. REFERENČNÍ TEXT'), 'Prompt obsahuje referenční text');
  assert.ok(promptText.includes('=== 2. KONTROLOVANÝ NOVÝ PŘEKLAD'), 'Prompt obsahuje kontrolovaný nový překlad');
  assert.ok(promptText.includes(myTranslation), 'Prompt obsahuje samotný text překladu');
  assert.ok(promptText.includes('ROZDÍLY A NEDOSTATKY'), 'Prompt obsahuje požadavek na výpis rozdílů nebo nedostatků');
  assert.ok(promptText.includes('NÁVRH ÚPRAVY'), 'Prompt obsahuje požadavek na návrh úpravy');

  // 3. Přepnutí do angličtiny
  document.getElementById('langToggleBtn').click();
  const card0En = document.getElementById('article-card-0');
  const verifyBtnEn = card0En.querySelector('.copy-verify-prompt-btn');
  assert.ok(verifyBtnEn.textContent.includes('Verify translation'), 'V angličtině má tlačítko anglický text');
  verifyBtnEn.click();
  await sleep(10);
  const promptEn = window.getLastCopiedText();
  assert.ok(promptEn.includes('Verify the semantic equivalence'), 'Anglický prompt obsahuje úkol ověření');
  assert.ok(promptEn.includes('=== 2. NEW TRANSLATION UNDER REVIEW'), 'Anglický prompt obsahuje sekci nového překladu');
  assert.ok(promptEn.includes('DISCREPANCIES & DEFECTS'), 'Anglický prompt obsahuje sekci nedostatků');
});

runTest('Nově přidaný jazyk (bez původního sloupce): přidání v Kroku 1, 1-sloupcový layout ve studiu, prostý překlad a export', async () => {
  const { window, document } = createDomEnvironment();
  document.getElementById('loadSampleBtn').click();
  await sleep(100);

  const nameInput = document.getElementById('newLangNameInput');
  const colNameInput = document.getElementById('newLangColNameInput');
  const addBtn = document.getElementById('addNewLangBtn');
  assert.ok(nameInput && addBtn, 'Ovládací prvky pro přidání jazyka existují v Kroku 1');

  nameInput.value = 'Polština';
  colNameInput.value = 'PL';
  addBtn.click();
  await sleep(50);

  const rowNew = document.querySelector('tr.lang-row-new');
  assert.ok(rowNew, 'Řádek pro nově přidaný jazyk je vykreslen v tabulce');
  assert.ok(rowNew.textContent.includes('Polština'));
  assert.ok(rowNew.querySelector('.badge-info').textContent.includes('nový jazyk'));
  assert.ok(rowNew.querySelector('.remove-lang-btn'), 'Obsahuje tlačítko pro odebrání');

  document.getElementById('confirmMappingBtn').click();
  document.getElementById('goToTranslationBtn').click();
  await sleep(50);

  const tabs = Array.from(document.querySelectorAll('#langTabsWrapper .lang-tab'));
  const newTab = tabs.find(t => t.textContent.includes('Polština'));
  assert.ok(newTab, 'Záložka nového jazyka existuje');
  newTab.click();
  await sleep(50);

  const card0 = document.getElementById('article-card-0');
  assert.ok(card0, 'Karta odstavce existuje');
  const grid = card0.querySelector('.article-grid');
  assert.ok(grid.classList.contains('layout-new-lang'), 'Mřížka používá layout-new-lang');
  assert.strictEqual(card0.querySelector('.cardTargetOldTitle'), null, 'Stará verze chybí');
  assert.strictEqual(card0.querySelector('.live-target-diff'), null, 'Rozdílový panel chybí');
  const textarea = card0.querySelector('textarea.target-textarea');
  assert.ok(textarea, 'Editor je přítomen');

  card0.querySelector('.copy-prompt-btn').click();
  await sleep(20);
  const copied = window.getLastCopiedText();
  assert.ok(copied.includes('Přeložte text do jazyka: Polština'), 'Prompt je prostý překlad');
  assert.ok(!copied.includes('STARÁ VERZE'), 'Neobsahuje staré znění reference');
  assert.ok(!copied.includes('SCHVÁLENÝ PŘEKLAD'), 'Neobsahuje původní schválený překlad');

  // Hromadné schválení neoznačí prázdné nepřeložené řádky u nového jazyka
  document.getElementById('approveUnchangedBtn').click();
  await sleep(20);
  const card0After = document.getElementById('article-card-0');
  assert.strictEqual(card0After.classList.contains('status-verified'), false, 'Nepřeložený odstavec nového jazyka se neoznačí jako schválený');

  const curTextarea = card0After.querySelector('textarea.target-textarea');
  typeInto(window, curTextarea, 'Przetłumaczony tekst.');
  const cb = card0After.querySelector('.verify-checkbox');
  cb.checked = true;
  cb.dispatchEvent(new window.Event('change'));
  await sleep(20);
  assert.ok(card0After.classList.contains('status-verified'), 'Po ručním schválení je hotovo');

  document.getElementById('goToExportBtn').click();
  await sleep(50);
  const compFile = window.__app.exportTables('comprehensive');
  assert.ok(compFile, 'Kompletní export úspěšně proběhl');
  const cleanFile = window.__app.exportTables('clean');
  assert.ok(cleanFile, 'Export nové verze úspěšně proběhl');

  const jsonStr = window.StorageManager.serializeProjectJson(window.__app.state);
  const parsed = window.StorageManager.parseProjectJson(jsonStr);
  assert.ok(Object.keys(parsed.mappings.languages).some(k => parsed.mappings.languages[k].isNewLanguage && parsed.mappings.languages[k].label === 'Polština'));
});

runTest('Nahrání textového souboru (.txt): rozdělení na odstavce jako referenční sloupec, přidání jazyka a export', async () => {
  const { window, document } = createDomEnvironment();
  const txtContent = 'První odstavec textu k překladu.\n\nDruhý odstavec obsahující další informace.\n\nTřetí odstavec dokumentu.';
  const file = new window.File([txtContent], 'predpis.txt');

  window.__app.processUploadedFile(file);
  await sleep(100);

  const st = window.__app.state;
  assert.deepStrictEqual(plain(st.headers), ['Reference']);
  assert.strictEqual(st.rows.length, 3);
  assert.strictEqual(st.rows[0]['Reference'], 'První odstavec textu k překladu.');
  assert.strictEqual(st.rows[1]['Reference'], 'Druhý odstavec obsahující další informace.');
  assert.strictEqual(st.rows[2]['Reference'], 'Třetí odstavec dokumentu.');

  // Možnosti listů i záhlaví jsou pro text skryty
  assert.strictEqual(document.getElementById('sheetOptionsWrapper').style.display, 'none');

  // Výběr referenčních sloupců: refNewCol = Reference, refOldCol = ''
  assert.strictEqual(document.getElementById('refNewSelect').value, 'Reference');
  assert.strictEqual(document.getElementById('refOldSelect').value, '');
  assert.strictEqual(document.getElementById('refOldSelect').options[0].textContent, '— žádná (pouze nová reference) —');

  // Bez cílového jazyka nelze pokračovat
  document.getElementById('confirmMappingBtn').click();
  await sleep(20);
  assert.strictEqual(document.getElementById('step2Section').style.display, 'none');

  // Přidání cílového jazyka
  document.getElementById('newLangNameInput').value = 'Polština';
  document.getElementById('newLangColNameInput').value = 'PL';
  document.getElementById('addNewLangBtn').click();
  await sleep(50);

  // Potvrzení mapování a přímý přechod do překladového studia (Krok 2 se přeskakuje)
  document.getElementById('confirmMappingBtn').click();
  await sleep(50);
  assert.strictEqual(document.getElementById('step2Section').style.display, 'none');
  assert.strictEqual(document.getElementById('step3Section').style.display, 'block');

  const card0 = document.getElementById('article-card-0');
  assert.ok(card0, 'Karta prvního odstavce existuje');
  assert.strictEqual(card0.querySelector('.diff-box').textContent.trim(), 'První odstavec textu k překladu.');
  assert.strictEqual(card0.querySelector('.cardTargetOldTitle'), null, 'Stará reference ani starý překlad se nezobrazují');

  // Přeložení prvního odstavce
  const textarea = card0.querySelector('textarea.target-textarea');
  typeInto(window, textarea, 'Pierwszy akapit tekstu.');
  const cb = card0.querySelector('.verify-checkbox');
  cb.checked = true;
  cb.dispatchEvent(new window.Event('change'));
  await sleep(20);
  assert.ok(card0.classList.contains('status-verified'));

  // Krok 4: Export
  document.getElementById('goToExportBtn').click();
  await sleep(50);
  const cleanFile = window.__app.exportTables('clean');
  assert.ok(cleanFile, 'Export nové verze proběhl');
  const wbClean = (await lastWorkbook(window)).wb;
  const sheetClean = wbClean.Sheets[wbClean.SheetNames[0]];
  const cleanHeaders = window.XLSX.utils.sheet_to_json(sheetClean, { header: 1 })[0];
  assert.deepStrictEqual(plain(cleanHeaders), ['Reference', 'PL']);
  const cleanRows = window.XLSX.utils.sheet_to_json(sheetClean, { defval: '' });
  assert.strictEqual(cleanRows[0]['PL'], 'Pierwszy akapit tekstu.');
});

runTest('Nahrání Word dokumentu (.docx a .doc): odstavce jako referenční buňky, čistý přehled změn a překlad', async () => {
  const { window, document } = createDomEnvironment();

  // Test .docx
  const docxBuf = createTestDocx(['První odstavec z Wordu.', 'Druhý odstavec z Wordu.']);
  const docxFile = new window.File([docxBuf], 'dokument.docx');

  window.__app.processUploadedFile(docxFile);
  await sleep(100);

  let st = window.__app.state;
  assert.deepStrictEqual(plain(st.headers), ['Reference']);
  assert.strictEqual(st.rows.length, 2);
  assert.strictEqual(st.rows[0]['Reference'], 'První odstavec z Wordu.');
  assert.strictEqual(st.rows[1]['Reference'], 'Druhý odstavec z Wordu.');
  assert.strictEqual(document.getElementById('sheetOptionsWrapper').style.display, 'none');

  // Test .doc (HTML formát Wordu)
  const htmlDoc = '<html xmlns:w="urn:schemas-microsoft-com:office:word"><body><p>HTML odstavec 1.</p><p>HTML odstavec 2.</p></body></html>';
  const docFile = new window.File([htmlDoc], 'dokument.doc');
  window.__app.processUploadedFile(docFile);
  await sleep(100);

  st = window.__app.state;
  assert.strictEqual(st.rows.length, 2);
  assert.strictEqual(st.rows[0]['Reference'], 'HTML odstavec 1.');
  assert.strictEqual(st.rows[1]['Reference'], 'HTML odstavec 2.');

  // Přímý režim: skrytí staré reference, zobrazení rychlého výběru jazyků a skrytí kroku 2 v navigaci
  assert.strictEqual(document.getElementById('refOldColCard').style.display, 'none');
  assert.strictEqual(document.getElementById('quickLangsWrapper').style.display, 'block');
  assert.strictEqual(document.getElementById('stepNav2').style.display, 'none');

  // Přidání jazyka a dokončení workflow
  document.getElementById('newLangNameInput').value = 'Francouzština';
  document.getElementById('newLangColNameInput').value = 'FR';
  document.getElementById('addNewLangBtn').click();
  await sleep(50);

  // Potvrzení mapování přejde přímo do studia (Krok 3)
  document.getElementById('confirmMappingBtn').click();
  await sleep(50);
  assert.strictEqual(document.getElementById('step2Section').style.display, 'none');
  assert.strictEqual(document.getElementById('step3Section').style.display, 'block');

  const card0 = document.getElementById('article-card-0');
  assert.ok(card0.querySelector('.article-grid').classList.contains('layout-new-lang'));

  // Tlačítko Zpět do nastavení jazyků v záhlaví studia
  const backBtn = document.getElementById('backToSettingsBtn');
  assert.strictEqual(backBtn.style.display, 'inline-flex');
  backBtn.click();
  await sleep(20);
  assert.strictEqual(document.getElementById('step1Section').style.display, 'block');
});

runTest('Nahrání Excelu s jediným sloupcem: automatická nová reference bez staré, přidání jazyka a export', async () => {
  const { window, document } = createDomEnvironment();
  const file = workbookFile(window, {
    Data: [
      ['Výchozí text'],
      ['První odstavec předpisu.'],
      ['Druhý odstavec předpisu.']
    ]
  }, 'jednosloupcovy.xlsx');

  window.__app.processUploadedFile(file);
  await sleep(100);

  const st = window.__app.state;
  assert.deepStrictEqual(plain(st.headers), ['Výchozí text']);
  assert.strictEqual(st.rows.length, 2);
  assert.strictEqual(document.getElementById('refNewSelect').value, 'Výchozí text');
  assert.strictEqual(document.getElementById('refOldSelect').value, '');

  // Pro Excel zůstávají možnosti listů a záhlaví dostupné
  assert.strictEqual(document.getElementById('sheetOptionsWrapper').style.display, 'flex');

  // Přidání cílového jazyka
  document.getElementById('newLangNameInput').value = 'Italština';
  document.getElementById('newLangColNameInput').value = 'IT';
  document.getElementById('addNewLangBtn').click();
  await sleep(50);

  document.getElementById('confirmMappingBtn').click();
  document.getElementById('goToTranslationBtn').click();
  await sleep(50);

  const card0 = document.getElementById('article-card-0');
  typeInto(window, card0.querySelector('textarea.target-textarea'), 'Primo paragrafo.');
  card0.querySelector('.verify-checkbox').click();
  await sleep(20);

  document.getElementById('goToExportBtn').click();
  await sleep(50);
  const compFile = window.__app.exportTables('comprehensive');
  assert.ok(compFile, 'Export srovnání proběhl');
  const wbComp = (await lastWorkbook(window)).wb;
  const sheetComp = wbComp.Sheets[wbComp.SheetNames[0]];
  const compHeaders = window.XLSX.utils.sheet_to_json(sheetComp, { header: 1 })[0];
  assert.deepStrictEqual(plain(compHeaders), ['Výchozí text', 'IT']);
  const compRows = window.XLSX.utils.sheet_to_json(sheetComp, { defval: '' });
  assert.strictEqual(compRows[0]['IT'], 'Primo paragrafo.');
});

runTest('Rychlý výběr jazyků pro přímý překlad (tlačítka rychlých jazyků a automatické přidání vepsaného jazyka)', async () => {
  const { window, document } = createDomEnvironment();
  const file = new window.File(['Odstavec 1\n\nOdstavec 2'], 'soubor.txt');

  window.__app.processUploadedFile(file);
  await sleep(100);

  let quickBtns = document.querySelectorAll('.quick-lang-btn');
  assert.ok(quickBtns.length >= 8, 'Rychlá tlačítka jazyků jsou vygenerována');

  // Kliknutí na tlačítko přidá jazyk
  quickBtns[0].click(); // CS
  await sleep(30);
  quickBtns = document.querySelectorAll('.quick-lang-btn');
  assert.ok(quickBtns[0].classList.contains('active'));
  assert.strictEqual(window.__app.state.mappings.targetCols.length, 1);

  // Kliknutí na další tlačítko
  quickBtns[2].click(); // EN
  await sleep(30);
  quickBtns = document.querySelectorAll('.quick-lang-btn');
  assert.strictEqual(window.__app.state.mappings.targetCols.length, 2);

  // Znovu kliknutí na první tlačítko jazyk odebere
  quickBtns[0].click();
  await sleep(30);
  quickBtns = document.querySelectorAll('.quick-lang-btn');
  assert.ok(!quickBtns[0].classList.contains('active'));
  assert.strictEqual(window.__app.state.mappings.targetCols.length, 1);

  // Napsání názvu jazyka do inputu a kliknutí na confirmMappingBtn jej automaticky přidá
  document.getElementById('newLangNameInput').value = 'Španělština';
  document.getElementById('confirmMappingBtn').click();
  await sleep(50);

  assert.strictEqual(window.__app.state.mappings.targetCols.length, 2);
  assert.strictEqual(document.getElementById('step3Section').style.display, 'block');
});

// Výsledky
runTest('Import překladů ze sešitu: vyplněný částečný sešit se po potvrzení převezme ke kontrole', async () => {
  const { window, document } = createDomEnvironment();
  openCzechStudio(window);
  const state = window.__app.state;
  const stubs = state.translations['CS 2024'];
  const idx = state.rowNumbers.indexOf(5);
  assert.strictEqual(stubs[idx].classification, 'MODIFIED');
  stubs[idx].isVerified = true;
  const unchangedIdx = state.rowNumbers.indexOf(2);

  // Buttons in Step 3 and Step 4 open the file picker
  let picked = 0;
  document.getElementById('translationsFileInput').click = () => { picked++; };
  document.getElementById('importTranslationsBtn').click();
  window.__app.setStep(4);
  document.getElementById('importTranslationsExportBtn').click();
  assert.strictEqual(picked, 2);

  // A file without the row column is explained and changes nothing
  await window.__app.importTranslationsFile(workbookFile(window, { CS: [['CS 2026'], ['x']] }, 'bez_radku.xlsx'));
  assert.ok(/Řádek/.test(window.__lastAlert), 'Vysvětlení chybějícího sloupce Řádek');

  // Declined confirmation changes nothing
  const file = () => workbookFile(window, {
    CS: [['Řádek', 'DE 2026 (Referenz)', 'CS 2026', 'Stav'],
      [2, '', stubs[unchangedIdx].targetNew, 'Ověřeno'],
      [5, '', 'Dárky do 50 EUR.', 'Ověřeno'],
      [999, '', 'Neznámý řádek', '']]
  }, 'ukazkovy_dokument_CS.xlsx');
  window.__confirmAnswer = false;
  await window.__app.importTranslationsFile(file());
  assert.notStrictEqual(stubs[idx].targetNew, 'Dárky do 50 EUR.');
  assert.ok(/CS 2026/.test(window.__lastConfirm) && /999/.test(window.__lastConfirm), 'Potvrzení uvádí sloupec a neznámé řádky');

  window.__confirmAnswer = true;
  const count = await window.__app.importTranslationsFile(file());
  assert.strictEqual(count, 1);
  const updated = window.__app.state.translations['CS 2024'][idx];
  assert.strictEqual(updated.targetNew, 'Dárky do 50 EUR.');
  assert.strictEqual(updated.isVerified, false, 'Převzatý text čeká na kontrolu');
  assert.ok(toastTexts(document).some(t => /Převzato 1/.test(t)));

  // The same file again: nothing new
  await window.__app.importTranslationsFile(file());
  assert.ok(toastTexts(document).some(t => /žádné nové ani změněné/.test(t)));
});

runTest('Přístupnost: axe bez nálezů ve všech krocích a v dialogu (kontrast se ověřuje v prohlížeči)', async () => {
  const axeSource = fs.readFileSync(require.resolve('axe-core/axe.min.js'), 'utf8');
  const { window, document } = createDomEnvironment();
  window.eval(axeSource);
  const audit = async (label) => {
    const res = await window.axe.run(document, { rules: { 'color-contrast': { enabled: false } } });
    const found = res.violations.map(v => `${label}: ${v.id} (${v.nodes.length}×, ${v.nodes[0].html.slice(0, 80)})`);
    assert.strictEqual(found.join(' | '), '', 'Porušení přístupnosti');
  };
  await audit('start');
  openCzechStudio(window);
  await audit('krok 3');
  window.__app.setStep(1);
  await audit('krok 1');
  window.__app.setStep(2);
  await audit('krok 2');
  window.__app.setStep(4);
  await audit('krok 4');
  document.getElementById('smartPasteBtn').click();
  await audit('Smart Paste');
});

runTest('Přístupnost: ovládání klávesnicí, stav tlačítek a fokus v dialogu', async () => {
  const { window, document } = createDomEnvironment();
  // The upload area opens the file picker with Enter
  let picked = 0;
  document.getElementById('fileInput').click = () => { picked++; };
  document.getElementById('dropzone').dispatchEvent(new window.KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
  assert.strictEqual(picked, 1, 'Enter na oblasti pro nahrání otevře výběr souboru');

  openCzechStudio(window);
  assert.strictEqual(document.getElementById('stepNav3').getAttribute('aria-current'), 'step');
  assert.strictEqual(document.getElementById('stepNav1').hasAttribute('aria-current'), false);
  ['stepNav1', 'stepNav2', 'stepNav3', 'stepNav4'].forEach(id => assert.strictEqual(document.getElementById(id).tagName, 'BUTTON'));
  assert.ok(Array.from(document.querySelectorAll('#langTabsWrapper .lang-tab')).every(tab => tab.tagName === 'BUTTON'));
  assert.strictEqual(document.querySelector('#langTabsWrapper .lang-tab.active').getAttribute('aria-pressed'), 'true');

  const chip = document.querySelector('.filter-chip[data-status="pending"]');
  chip.click();
  assert.strictEqual(chip.getAttribute('aria-pressed'), 'true');
  assert.strictEqual(document.querySelector('.filter-chip[data-status="all"]').getAttribute('aria-pressed'), 'false');

  // Focus moves into the dialog and back to the button that opened it
  const opener = document.getElementById('smartPasteBtn');
  opener.focus();
  opener.click();
  await sleep(0);
  assert.strictEqual(document.activeElement, document.getElementById('smartPasteInput'), 'Fokus v dialogu');
  document.dispatchEvent(new window.KeyboardEvent('keydown', { key: 'Escape' }));
  await sleep(0);
  assert.strictEqual(document.activeElement, opener, 'Fokus se vrátí na tlačítko');

  // Accessible names follow the interface language
  document.getElementById('langToggleBtn').click();
  assert.strictEqual(document.getElementById('themeToggleBtn').getAttribute('aria-label'), 'Toggle dark / light mode');
  assert.strictEqual(document.getElementById('closeSmartPasteBtn').getAttribute('aria-label'), 'Close');
});

function fixtureFile(window, name) {
  return new window.File([fs.readFileSync(path.join(__dirname, 'fixtures', name))], name);
}

function chooseFile(window, input, file) {
  Object.defineProperty(input, 'files', { value: [file], configurable: true });
  input.dispatchEvent(new window.Event('change'));
}

async function loadRichWorkbook(window) {
  const document = window.document;
  await window.__app.processUploadedFile(fixtureFile(window, 'excel-rich.xlsx'));
  setSelect(window, document.getElementById('refOldSelect'), 'DE 2024');
  setSelect(window, document.getElementById('refNewSelect'), 'DE 2026');
  // The short sample texts are not pre-selected as a language; tick it like a user would
  const cb = langRow(document, 'CS 2024').querySelector('.lang-use-cb');
  if (!cb.checked) { cb.checked = true; cb.dispatchEvent(new window.Event('change')); }
  assert.deepStrictEqual(plain(window.__app.state.mappings.targetCols), ['CS 2024']);
  document.getElementById('confirmMappingBtn').click();
  window.__app.setStep(3);
  typeInto(window, document.getElementById('article-card-0').querySelector('textarea.target-textarea'), 'Nový text jedna.');
  window.__app.setStep(4);
}

runTest('Export do původního sešitu: tlačítko v Kroku 4 zapíše novou verzi do kopie nahraného souboru', async () => {
  const { window, document } = createDomEnvironment();
  await loadRichWorkbook(window);
  assert.strictEqual(document.getElementById('updateOriginalWrapper').style.display, 'flex');

  document.getElementById('updateOriginalBtn').click();
  const out = await lastWorkbook(window);
  assert.strictEqual(out.name, 'excel-rich_srovnani.xlsx');
  assert.deepStrictEqual(plain(out.wb.SheetNames), ['Data', 'Jiný', 'Stav_prekladu']);
  assert.strictEqual(out.wb.Sheets.Data.I1.v, 'CS 2026');
  assert.strictEqual(out.wb.Sheets.Data.I2.v, 'Nový text jedna.');
  assert.ok(toastTexts(document).some(t => /s formátováním původního souboru/.test(t)));
});

runTest('Export do původního sešitu po načtení projektu: aplikace požádá o původní soubor a ověří, že patří k projektu', async () => {
  const first = createDomEnvironment();
  await loadRichWorkbook(first.window);
  const project = JSON.parse(first.window.StorageManager.serializeProjectJson(first.window.__app.state));
  // A new session (page reloaded) opens the project file
  const { window, document } = createDomEnvironment();
  window.__app.restoreProjectFromData(window.StorageManager.parseProjectJson(project));
  window.__app.setStep(4);
  assert.strictEqual(window.__app.state.originalFile, null, 'Původní soubor není součástí projektu');

  const input = document.getElementById('originalFileInput');
  let opened = 0;
  input.click = () => { opened++; };
  document.getElementById('updateOriginalBtn').click();
  assert.strictEqual(opened, 1, 'Otevře se výběr souboru');
  assert.ok(toastTexts(document).some(t => /Vyberte původní soubor excel-rich\.xlsx/.test(t)));

  // A different workbook is refused
  chooseFile(window, input, workbookFile(window, { Data: [['ID', 'DE 2024', 'DE 2026', 'CS 2024'], ['A1', 'x', 'y', 'z']] }, 'jiny.xlsx'));
  await sleep(50);
  assert.ok(/neodpovídá projektu/.test(window.__lastAlert), 'Cizí soubor odmítnut: ' + window.__lastAlert);
  assert.strictEqual(window.__downloads.length, 0);

  document.getElementById('updateOriginalBtn').click();
  chooseFile(window, input, fixtureFile(window, 'excel-rich.xlsx'));
  await sleep(50);
  const out = await lastWorkbook(window);
  assert.strictEqual(out.wb.Sheets.Data.I2.v, 'Nový text jedna.');

  // The selected file is remembered for the next export
  document.getElementById('updateOriginalBtn').click();
  assert.strictEqual(opened, 2, 'Podruhé se už neptá');
  assert.ok(window.__downloads.length === 1);
});

runTest('Word export: dokument z Wordu se přeloží v kopii originálu, jiný zdroj vytvoří nový dokument', async () => {
  const { window, document } = createDomEnvironment();
  await window.__app.processUploadedFile(new window.File([createTestDocx(['Odstavec jedna.', 'Odstavec dva.'])], 'smlouva.docx'));
  document.getElementById('newLangNameInput').value = 'EN';
  document.getElementById('addNewLangBtn').click();
  document.getElementById('confirmMappingBtn').click();
  typeInto(window, document.getElementById('article-card-0').querySelector('textarea.target-textarea'), 'Paragraph one.');
  window.__app.setStep(4);
  assert.strictEqual(document.getElementById('docxOriginalHint').style.display, 'block');
  const select = document.getElementById('pdfLangSelect');
  setSelect(window, select, Array.from(select.options).find(o => o.textContent.trim().startsWith('EN')).value);

  document.getElementById('downloadDocxBtn').click();
  const download = window.__downloads.pop();
  assert.strictEqual(download.name, 'smlouva_EN.docx');
  const buffer = await readBlobBuffer(window, download.blob);
  const source = window.WordImporter.parse(buffer, { fileName: 'x.docx' });
  assert.deepStrictEqual(plain(window.Importers.readSheet(source, source.sheetNames[0]).rows.map(r => r.Reference)), ['Paragraph one.', 'Odstavec dva.']);
  assert.ok(toastTexts(document).some(t => /1 odstavců bez překladu zůstalo v původním znění/.test(t)));

  // Excel source: a new Word document
  const env2 = createDomEnvironment();
  openCzechStudio(env2.window);
  env2.window.__app.setStep(4);
  assert.strictEqual(env2.document.getElementById('docxOriginalHint').style.display, 'none');
  assert.strictEqual(env2.document.getElementById('updateOriginalWrapper').style.display, 'flex', 'Ukázka je sešit .xlsx');
  env2.document.getElementById('downloadDocxBtn').click();
  const generated = env2.window.__downloads.pop();
  assert.ok(/\.docx$/.test(generated.name));
  const gsrc = env2.window.WordImporter.parse(await readBlobBuffer(env2.window, generated.blob), { fileName: 'g.docx' });
  assert.ok(env2.window.Importers.readSheet(gsrc, gsrc.sheetNames[0]).rows.length > 5);
});

runAllTests().then(() => {
  console.log('\n========================================');
  console.log(`VÝSLEDEK BATERIE GUI TESTŮ: ${passedTests} úspěšných, ${failedTests} chyb`);
  console.log('========================================\n');
  process.exit(failedTests > 0 ? 1 : 0);
});
