/**
 * Komplexní End-to-End (E2E) integrační testy pro Comparator / Internationalisation Manager.
 *
 * Ověřuje oba klíčové toky zpracování od načtení souboru až po export:
 * 1. Vícejazyčný Excel sešit: stará reference vs nová reference, stávající jazyky + nově přidaný jazyk,
 *    přehled změn (Krok 2), překladové studio (Krok 3) a exporty do Excelu/PDF (Krok 4).
 * 2. Dokument s jediným výchozím textem k přímému překladu (Word .docx/.doc, textový .txt, 1-sloupcový Excel):
 *    skrytí staré reference, vynechání Kroku 2, rychlá volba jazyků, přímý přechod do studia,
 *    navigace zpět do nastavení a exporty.
 * 3. Přepínání mezi oběma režimy bez kontaminace stavu a ovládacích prvků.
 */

const fs = require('fs');
const path = require('path');
const assert = require('assert');
const { JSDOM } = require('jsdom');
const NodeXLSX = require('xlsx');

let passedTests = 0;
let failedTests = 0;
const tests = [];

function runTest(name, fn) {
  tests.push({ name, fn });
}

async function runAllTests() {
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

function readBlobBuffer(window, blob) {
  return new Promise((resolve, reject) => {
    const reader = new window.FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = () => reject(reader.error);
    reader.readAsArrayBuffer(blob);
  });
}

async function lastWorkbook(window) {
  const download = window.__downloads.pop();
  assert.ok(download && /\.xlsx$/.test(download.name), 'Stáhne se sešit .xlsx: ' + (download && download.name));
  const buffer = await readBlobBuffer(window, download.blob);
  return { name: download.name, wb: window.XLSX.read(new Uint8Array(buffer), { type: 'array' }) };
}

function typeInto(window, textarea, text) {
  textarea.value = text;
  textarea.dispatchEvent(new window.Event('input'));
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

const htmlPath = path.join(__dirname, '..', 'index.html');
const htmlSource = fs.readFileSync(htmlPath, 'utf8');

const scriptSources = [...htmlSource.matchAll(/<script src="([^"]+)"><\/script>/g)]
  .map(m => fs.readFileSync(path.join(__dirname, '..', m[1]), 'utf8'));

function createDomEnvironment(options) {
  const opts = options || {};
  const dom = new JSDOM(htmlSource, {
    url: 'http://localhost:3000/',
    runScripts: 'outside-only',
    pretendToBeVisual: true
  });

  const { window } = dom;

  let lastCopiedText = '';
  window.navigator.clipboard = {
    writeText: async (text) => {
      lastCopiedText = text;
      return Promise.resolve();
    }
  };

  window.getLastCopiedText = () => lastCopiedText;
  window.alert = (msg) => { window.__lastAlert = msg; };
  window.confirm = (msg) => {
    window.__lastConfirm = msg;
    if (window.__confirmAnswer !== undefined) return window.__confirmAnswer;
    return /rozpracovaný projekt z|saved project from/.test(msg) ? !!opts.acceptResume : true;
  };
  window.scrollTo = () => {};
  window.Element.prototype.scrollIntoView = () => {};

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

  scriptSources.forEach((src, i) => {
    if (i === scriptSources.length - 1 && opts.beforeApp) opts.beforeApp(window);
    window.eval(src);
  });

  window.document.dispatchEvent(new window.Event('DOMContentLoaded'));

  return { dom, window, document: window.document };
}

console.log('=== Spouštím End-to-End (E2E) testy pro vícejazyčné a jednosloupcové toky ===\n');

// ------------------------------------------------------------------------------------------------
// TEST 1: Kompletní E2E průchod vícejazyčným Excel souborem (stará + nová reference + stávající jazyky + nový jazyk)
// ------------------------------------------------------------------------------------------------
runTest('E2E: Vícejazyčný Excel sešit (srovnání staré a nové reference, stávající jazyky i nově přidaný jazyk, Kroky 1-4)', async () => {
  const { window, document } = createDomEnvironment();

  // 1. Příprava vícejazyčného sešitu
  const testData = [
    ['Ref_Old', 'Ref_New', 'CS_Old', 'SK_Old'],
    ['Společné ustanovení pro zaměstnance.', 'Společné ustanovení pro zaměstnance.', 'Společné ustanovení pro zaměstnance.', 'Spoločné ustanovenie pre zamestnancov.'], // UNCHANGED
    ['Všechny zprávy se posílají poštou.', 'Všechny zprávy se posílají elektronicky.', 'Všechny zprávy se posílají poštou.', 'Všetky správy sa posielajú poštou.'], // MODIFIED
    ['', 'Nové pravidlo pro práci na dálku.', '', ''], // ADDED
    ['Zrušené ustanovení o telegramech.', '', 'Zrušené ustanovení o telegramech.', 'Zrušené ustanovenie o telegramoch.'] // REMOVED
  ];
  const file = workbookFile(window, { Dokument: testData }, 'vicejazycny_predpis.xlsx');

  // KROK 1: Nahrání
  window.__app.processUploadedFile(file);
  await sleep(100);

  const st = window.__app.state;
  assert.deepStrictEqual(plain(st.headers), ['Ref_Old', 'Ref_New', 'CS_Old', 'SK_Old']);
  assert.strictEqual(st.rows.length, 4);

  // Ověření vícejazyčného rozhraní v Kroku 1
  assert.strictEqual(document.getElementById('refOldColCard').style.display, '', 'Karta staré reference je zobrazena');
  assert.strictEqual(document.getElementById('quickLangsWrapper').style.display, 'none', 'Rychlý výběr jazyků pro přímý překlad je skryt');
  assert.strictEqual(document.getElementById('languagesTableWrapper').style.display, '', 'Tabulka sloupců původního souboru je zobrazena');
  assert.strictEqual(document.getElementById('stepNav2').style.display, '', 'Krok 2 je v navigaci zobrazen');
  assert.strictEqual(document.getElementById('stepNumber3').textContent, '3', 'Krok 3 má číslo 3');
  assert.strictEqual(document.getElementById('stepNumber4').textContent, '4', 'Krok 4 má číslo 4');

  // Nastavení referencí
  const refOldSelect = document.getElementById('refOldSelect');
  const refNewSelect = document.getElementById('refNewSelect');
  refOldSelect.value = 'Ref_Old';
  refOldSelect.dispatchEvent(new window.Event('change'));
  refNewSelect.value = 'Ref_New';
  refNewSelect.dispatchEvent(new window.Event('change'));

  // Přidání NOVÉHO jazyka (Polština), který v původním souboru nebyl
  document.getElementById('newLangNameInput').value = 'Polština';
  document.getElementById('newLangColNameInput').value = 'PL';
  document.getElementById('addNewLangBtn').click();
  await sleep(50);

  // KROK 2: Potvrzení mapování a přechod na přehled změn
  document.getElementById('confirmMappingBtn').click();
  await sleep(50);
  assert.strictEqual(document.getElementById('step1Section').style.display, 'none');
  assert.strictEqual(document.getElementById('step2Section').style.display, 'block');

  // Ověření statistiky v Kroku 2
  assert.strictEqual(document.getElementById('kpiTotalRows').textContent, '4');
  assert.strictEqual(document.getElementById('kpiUnchanged').textContent, '1');
  assert.strictEqual(document.getElementById('kpiModified').textContent, '1');
  assert.strictEqual(document.getElementById('kpiAdded').textContent, '1');
  assert.strictEqual(document.getElementById('kpiRemoved').textContent, '1');

  // KROK 3: Přechod do překladového studia
  document.getElementById('goToTranslationBtn').click();
  await sleep(50);
  assert.strictEqual(document.getElementById('step2Section').style.display, 'none');
  assert.strictEqual(document.getElementById('step3Section').style.display, 'block');

  // Matice a záložky obsahují CS_Old, SK_Old i nově přidanou Polštinu
  const langTabs = Array.from(document.querySelectorAll('#langTabsWrapper .lang-tab'));
  assert.strictEqual(langTabs.length, 3, 'K dispozici jsou 3 záložky jazyků');

  // A) Kontrola stávajícího jazyka (CS_Old)
  langTabs[0].click();
  await sleep(30);

  // Řádek 1 (UNCHANGED) je automaticky schválen
  const card1 = document.getElementById('article-card-0');
  assert.ok(card1.classList.contains('status-verified'), 'Nezměněný řádek je automaticky schválen');

  // Řádek 2 (MODIFIED) má diff a editor s tlačítky Copilot promptu
  const card2 = document.getElementById('article-card-1');
  assert.ok(!card2.classList.contains('status-verified'));
  const promptBtn = card2.querySelector('.copy-prompt-btn');
  promptBtn.click();
  await sleep(20);
  const copiedPrompt = window.getLastCopiedText();
  assert.ok(copiedPrompt.includes('Všechny zprávy se posílají poštou.'), 'Prompt obsahuje původní text');
  assert.ok(copiedPrompt.includes('elektronicky'), 'Prompt obsahuje novou změnu');

  // Zadání překladu pro MODIFIED řádek v CS
  const csTextarea = card2.querySelector('textarea.target-textarea');
  typeInto(window, csTextarea, 'Všechny zprávy se posílají elektronicky.');
  const csVerifyCb = card2.querySelector('.verify-checkbox');
  csVerifyCb.checked = true;
  csVerifyCb.dispatchEvent(new window.Event('change'));
  await sleep(20);
  assert.ok(card2.classList.contains('status-verified'));

  // B) Kontrola nově přidaného jazyka (Polština - PL)
  langTabs[2].click();
  await sleep(30);

  const plCard1 = document.getElementById('article-card-0');
  assert.ok(plCard1.querySelector('.article-grid').classList.contains('layout-new-lang'), 'Nový jazyk má 1-sloupcový layout');
  const plPromptBtn = plCard1.querySelector('.copy-prompt-btn');
  plPromptBtn.click();
  await sleep(20);
  const plPrompt = window.getLastCopiedText();
  assert.ok(plPrompt.includes('POKYNY:'), 'Prompt pro nový jazyk obsahuje standardní pokyny');
  assert.ok(!plPrompt.includes('=== 1. REFERENČNÍ TEXT – STARÁ VERZE ==='), 'Prompt pro nový jazyk neobsahuje diff staré verze');

  // Přeložení prvního odstavce do PL
  const plTextarea = plCard1.querySelector('textarea.target-textarea');
  typeInto(window, plTextarea, 'Wspólne postanowienie dla pracowników.');
  const plVerifyCb = plCard1.querySelector('.verify-checkbox');
  plVerifyCb.checked = true;
  plVerifyCb.dispatchEvent(new window.Event('change'));
  await sleep(20);
  assert.ok(plCard1.classList.contains('status-verified'));

  // KROK 4: Export výsledků
  document.getElementById('goToExportBtn').click();
  await sleep(50);
  assert.strictEqual(document.getElementById('step3Section').style.display, 'none');
  assert.strictEqual(document.getElementById('step4Section').style.display, 'block');

  // Test kompletního exportu srovnání
  document.getElementById('downloadComprehensiveBtn').click();
  await sleep(50);
  const compWb = (await lastWorkbook(window)).wb;
  assert.ok(compWb.SheetNames.includes('Srovnani'), 'Obsahuje list Srovnani');
  assert.ok(compWb.SheetNames.includes('Stav_prekladu'), 'Obsahuje list Stav_prekladu');

  const compSheet = compWb.Sheets['Srovnani'];
  const compHeaders = NodeXLSX.utils.sheet_to_json(compSheet, { header: 1 })[0];
  assert.ok(compHeaders.includes('Ref_Old'));
  assert.ok(compHeaders.includes('Ref_New'));
  assert.ok(compHeaders.includes('PL'), 'Nový jazyk PL je zařazen v exportu');

  // Test čisté nové verze
  document.getElementById('downloadCleanBtn').click();
  await sleep(50);
  const cleanWb = (await lastWorkbook(window)).wb;
  const cleanSheet = cleanWb.Sheets['Nova_verze'];
  const cleanRows = NodeXLSX.utils.sheet_to_json(cleanSheet, { defval: '' });
  assert.strictEqual(cleanRows[0]['PL'], 'Wspólne postanowienie dla pracowników.');
});

// ------------------------------------------------------------------------------------------------
// TEST 2: Kompletní E2E průchod dokumentem Word (.docx) v režimu přímého překladu
// ------------------------------------------------------------------------------------------------
runTest('E2E: Dokument Word (.docx) v režimu přímého překladu (vynechání Kroku 2, rychlá volba, návrat zpět a export)', async () => {
  const { window, document } = createDomEnvironment();

  const docxData = createTestDocx([
    'Všeobecné obchodní podmínky společnosti.',
    'Dodací lhůta činí 14 pracovních dnů od objednání.',
    'Záruka na dodané zboží je 24 měsíců.'
  ]);
  const file = new window.File([docxData], 'podminky.docx');

  // KROK 1: Nahrání
  window.__app.processUploadedFile(file);
  await sleep(100);

  const st = window.__app.state;
  assert.deepStrictEqual(plain(st.headers), ['Reference']);
  assert.strictEqual(st.rows.length, 3);

  // Ověření přímého režimu v Kroku 1
  assert.strictEqual(document.getElementById('refOldColCard').style.display, 'none', 'Karta staré reference je v přímém režimu skryta');
  assert.strictEqual(document.getElementById('quickLangsWrapper').style.display, 'block', 'Rychlý výběr jazyků je zobrazen');
  assert.strictEqual(document.getElementById('languagesTableWrapper').style.display, 'none', 'Původní tabulka sloupců je skryta');
  assert.strictEqual(document.getElementById('stepNav2').style.display, 'none', 'Krok 2 je v navigaci skryt');
  assert.strictEqual(document.getElementById('stepNumber3').textContent, '2', 'Krok 3 je přečíslován na 2');
  assert.strictEqual(document.getElementById('stepNumber4').textContent, '3', 'Krok 4 je přečíslován na 3');

  // Rychlá volba cílových jazyků kliknutím na tlačítka
  const quickBtns = Array.from(document.querySelectorAll('.quick-lang-btn'));
  assert.ok(quickBtns.length >= 8);

  // Vybereme CS a EN
  const csBtn = quickBtns.find(b => b.textContent.includes('Čeština'));
  const enBtn = quickBtns.find(b => b.textContent.includes('English'));
  csBtn.click();
  await sleep(30);
  enBtn.click();
  await sleep(30);

  assert.strictEqual(st.mappings.targetCols.length, 2, 'Vybrány 2 cílové jazyky');

  // Přímý přechod do překladu (Krok 2 se přeskakuje)
  document.getElementById('confirmMappingBtn').click();
  await sleep(50);
  assert.strictEqual(document.getElementById('step1Section').style.display, 'none');
  assert.strictEqual(document.getElementById('step2Section').style.display, 'none', 'Krok 2 byl přeskočen');
  assert.strictEqual(document.getElementById('step3Section').style.display, 'block', 'Otevřeno studio');

  // V záhlaví studia je tlačítko návratu do nastavení jazyků
  const backBtn = document.getElementById('backToSettingsBtn');
  assert.strictEqual(backBtn.style.display, 'inline-flex');

  // Ověření karty odstavce: čistý výchozí text bez diff odznaků
  const card0 = document.getElementById('article-card-0');
  const subPanelTitle = card0.querySelector('.sub-panel-title span').textContent;
  assert.ok(subPanelTitle.includes('Výchozí text'));
  assert.strictEqual(card0.querySelector('.diff-box').textContent.trim(), 'Všeobecné obchodní podmínky společnosti.');

  // Překlad odstavce v prvním jazyce
  typeInto(window, card0.querySelector('textarea.target-textarea'), 'General Terms and Conditions of the company.');
  const cb = card0.querySelector('.verify-checkbox');
  cb.checked = true;
  cb.dispatchEvent(new window.Event('change'));
  await sleep(20);
  assert.ok(card0.classList.contains('status-verified'));

  // Test návratu zpět do nastavení jazyků
  backBtn.click();
  await sleep(30);
  assert.strictEqual(document.getElementById('step1Section').style.display, 'block');

  // Přidáme další jazyk vepsáním do inputu a kliknutím na Přejít k překladu (auto-add)
  document.getElementById('newLangNameInput').value = 'Slovenčina';
  document.getElementById('confirmMappingBtn').click();
  await sleep(50);

  assert.strictEqual(st.mappings.targetCols.length, 3, 'Nyní jsou 3 jazyky');
  assert.strictEqual(document.getElementById('step3Section').style.display, 'block');

  // KROK 4: Export
  document.getElementById('goToExportBtn').click();
  await sleep(50);
  assert.strictEqual(document.getElementById('step4Section').style.display, 'block');

  document.getElementById('downloadCleanBtn').click();
  await sleep(50);
  const cleanWb = (await lastWorkbook(window)).wb;
  const cleanSheet = cleanWb.Sheets['Nova_verze'];
  const cleanHeaders = NodeXLSX.utils.sheet_to_json(cleanSheet, { header: 1 })[0];
  assert.ok(cleanHeaders.includes('Reference'));
  assert.ok(cleanHeaders.length >= 3, 'Obsahuje referenci a přeložené jazyky');
});

// ------------------------------------------------------------------------------------------------
// TEST 3: Kompletní E2E průchod textovým souborem (.txt)
// ------------------------------------------------------------------------------------------------
runTest('E2E: Textový soubor (.txt) v režimu přímého překladu s dávkovým promptem a exportem', async () => {
  const { window, document } = createDomEnvironment();

  const txtContent = 'První odstavec textu.\n\nDruhý odstavec textu s instrukcemi.\n\nTřetí odstavec textu.';
  const file = new window.File([txtContent], 'smlouva.txt');

  window.__app.processUploadedFile(file);
  await sleep(100);

  assert.strictEqual(window.__app.state.headers.length, 1);
  assert.strictEqual(window.__app.state.rows.length, 3);
  assert.strictEqual(document.getElementById('refOldColCard').style.display, 'none');

  // Přidání cílového jazyka
  document.getElementById('newLangNameInput').value = 'Francouzština';
  document.getElementById('newLangColNameInput').value = 'FR';
  document.getElementById('confirmMappingBtn').click();
  await sleep(50);

  assert.strictEqual(document.getElementById('step3Section').style.display, 'block');

  // Otevření a zkopírování hromadného promptu
  document.getElementById('batchPromptBtn').click();
  await sleep(30);
  assert.ok(document.getElementById('batchPromptModal').classList.contains('open'));
  const batchPromptText = document.getElementById('batchPromptTextarea').value;
  assert.ok(batchPromptText.includes('[#1]') || batchPromptText.includes('Řádek 1') || batchPromptText.includes('První odstavec textu.'));

  // Zavření modálu
  document.getElementById('closeBatchPromptBtn').click();
  await sleep(20);

  // Krok 4: Export
  document.getElementById('goToExportBtn').click();
  await sleep(50);
  document.getElementById('downloadCleanBtn').click();
  await sleep(50);
  const cleanWb = (await lastWorkbook(window)).wb;
  assert.ok(cleanWb.SheetNames.includes('Nova_verze'));
});

// ------------------------------------------------------------------------------------------------
// TEST 4: Kompletní E2E průchod jednosloupcovým Excel souborem (.xlsx)
// ------------------------------------------------------------------------------------------------
runTest('E2E: Jednosloupcový Excel (.xlsx) aktivuje přímý překlad a exportuje výsledky', async () => {
  const { window, document } = createDomEnvironment();

  const file = workbookFile(window, {
    Data: [
      ['Výchozí text'],
      ['Odstavec A.'],
      ['Odstavec B.']
    ]
  }, 'jeden_sloupec.xlsx');

  window.__app.processUploadedFile(file);
  await sleep(100);

  assert.strictEqual(window.__app.state.headers.length, 1);
  assert.strictEqual(document.getElementById('refOldColCard').style.display, 'none');
  assert.strictEqual(document.getElementById('quickLangsWrapper').style.display, 'block');

  // Vybereme rychlý jazyk
  const quickBtns = Array.from(document.querySelectorAll('.quick-lang-btn'));
  quickBtns[0].click(); // První jazyk
  await sleep(30);

  document.getElementById('confirmMappingBtn').click();
  await sleep(50);
  assert.strictEqual(document.getElementById('step3Section').style.display, 'block');

  // Export
  document.getElementById('goToExportBtn').click();
  await sleep(50);
  document.getElementById('downloadComprehensiveBtn').click();
  await sleep(50);
  const wb = (await lastWorkbook(window)).wb;
  assert.ok(wb.SheetNames.includes('Srovnani'));
});

// ------------------------------------------------------------------------------------------------
// TEST 5: Přepínání mezi jednosloupcovým a vícejazyčným souborem v rámci jedné relace
// ------------------------------------------------------------------------------------------------
runTest('E2E: Přepínání mezi Word/Txt a vícejazyčným Excelem zcela resetuje a adaptuje UI bez kontaminace', async () => {
  const { window, document } = createDomEnvironment();

  // Fáze 1: Nahrajeme Word dokument (.docx)
  const docxFile = new window.File([createTestDocx(['Text 1', 'Text 2'])], 'doc1.docx');
  window.__app.processUploadedFile(docxFile);
  await sleep(100);

  assert.strictEqual(document.getElementById('refOldColCard').style.display, 'none', 'Fáze 1: stará reference skryta');
  assert.strictEqual(document.getElementById('quickLangsWrapper').style.display, 'block', 'Fáze 1: rychlé jazyky zobrazeny');
  assert.strictEqual(document.getElementById('stepNav2').style.display, 'none', 'Fáze 1: Krok 2 skryt v navigaci');
  assert.strictEqual(document.getElementById('stepNumber3').textContent, '2');

  // Vybereme jazyk a přejdeme do studia
  const quickBtns = Array.from(document.querySelectorAll('.quick-lang-btn'));
  quickBtns[0].click();
  await sleep(30);
  document.getElementById('confirmMappingBtn').click();
  await sleep(50);
  assert.strictEqual(document.getElementById('step3Section').style.display, 'block');
  assert.strictEqual(document.getElementById('backToSettingsBtn').style.display, 'inline-flex');

  // Fáze 2: Uživatel nahraje vícejazyčný Excel sešit s 4 sloupci
  const excelFile = workbookFile(window, {
    List: [
      ['R_Old', 'R_New', 'CS_Old', 'SK_Old'],
      ['Původní text společného ustanovení.', 'Nový upravený text společného ustanovení.', 'Starý schválený překlad tohoto odstavce.', 'Starý schválený preklad tohto odseku.']
    ]
  }, 'vicejazycny.xlsx');

  window.__app.processUploadedFile(excelFile);
  await sleep(100);

  // Ověření, že se aplikace zcela vrátila do vícejazyčného režimu
  assert.strictEqual(document.getElementById('step1Section').style.display, 'block', 'Zpět v Kroku 1');
  assert.strictEqual(document.getElementById('refOldColCard').style.display, '', 'Karta staré reference opět viditelná');
  assert.strictEqual(document.getElementById('quickLangsWrapper').style.display, 'none', 'Rychlé jazyky skryty');
  assert.strictEqual(document.getElementById('languagesTableWrapper').style.display, '', 'Tabulka sloupců opět viditelná');
  assert.strictEqual(document.getElementById('stepNav2').style.display, '', 'Krok 2 opět zobrazen v navigaci');
  assert.strictEqual(document.getElementById('stepNumber3').textContent, '3', 'Krok 3 má zpět číslo 3');
  assert.strictEqual(document.getElementById('stepNumber4').textContent, '4', 'Krok 4 má zpět číslo 4');

  // Nastavíme reference pro vícejazyčný sešit
  const refOldSelect = document.getElementById('refOldSelect');
  const refNewSelect = document.getElementById('refNewSelect');
  refOldSelect.value = 'R_Old';
  refOldSelect.dispatchEvent(new window.Event('change'));
  refNewSelect.value = 'R_New';
  refNewSelect.dispatchEvent(new window.Event('change'));

  // Potvrzení mapování vícejazyčného sešitu přejde do Kroku 2 (nikoliv do Kroku 3!)
  document.getElementById('confirmMappingBtn').click();
  await sleep(50);
  assert.strictEqual(document.getElementById('step2Section').style.display, 'block', 'Přešlo do Kroku 2');
  assert.strictEqual(document.getElementById('step3Section').style.display, 'none');

  // Fáze 3: Uživatel nahraje obyčejný textový soubor (.txt)
  const txtFile = new window.File(['Jediný odstavec textu.'], 'prostotext.txt');
  window.__app.processUploadedFile(txtFile);
  await sleep(100);

  assert.strictEqual(document.getElementById('refOldColCard').style.display, 'none', 'Fáze 3: stará reference opět skryta');
  assert.strictEqual(document.getElementById('quickLangsWrapper').style.display, 'block', 'Fáze 3: rychlé jazyky opět zobrazeny');
  assert.strictEqual(document.getElementById('stepNav2').style.display, 'none', 'Fáze 3: Krok 2 v navigaci skryt');
  assert.strictEqual(document.getElementById('stepNumber3').textContent, '2');
  assert.strictEqual(window.__app.state.mappings.targetCols.length, 0, 'Vyčištěny předchozí jazyky');
});

runAllTests().then(() => {
  console.log('\n========================================');
  console.log(`VÝSLEDEK E2E TESTŮ: ${passedTests} úspěšných, ${failedTests} chyb`);
  console.log('========================================\n');
  process.exit(failedTests > 0 ? 1 : 0);
});
