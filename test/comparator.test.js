/**
 * Automated Test Suite for ComparatorEngine, StorageManager, I18n and the import/export adapters
 * Run with: node test/comparator.test.js
 */

const assert = require('assert');
const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');
const ComparatorEngine = require('../js/core/comparator-engine');
const StorageManager = require('../js/core/storage');
const SampleGenerator = require('../js/core/sample-generator');
const I18n = require('../js/core/i18n');
const Importers = require('../js/io/importers');
const ExcelImporter = require('../js/io/excel-importer');
const WordImporter = require('../js/io/word-importer');
const TextImporter = require('../js/io/text-importer');
const Exporters = require('../js/io/exporters');
const ExcelExporter = require('../js/io/excel-exporter');
const HtmlDocumentExporter = require('../js/io/html-document-exporter');
const SummaryExporters = require('../js/io/summary-exporters');

// What the application does: a format-independent model from the engine, rendered by the
// default export adapter
function exportWorkbook(options) {
  const result = ComparatorEngine.prepareExportTables(options);
  return { workbook: ExcelExporter.buildWorkbook({ tables: result.tables }), unverifiedCount: result.unverifiedCount };
}

function summaryOf(params) {
  const model = ComparatorEngine.buildChangeSummary(params);
  return Object.assign({}, model, {
    markdown: SummaryExporters.renderSummaryMarkdown(model),
    htmlDoc: SummaryExporters.renderSummaryHtml(model)
  });
}

function documentHtml(options) {
  return HtmlDocumentExporter.renderDocumentHtml(ComparatorEngine.buildDocumentModel(options));
}

let passCount = 0;
let failCount = 0;

function runTest(name, fn) {
  try {
    fn();
    console.log(`  ✓ ${name}`);
    passCount++;
  } catch (err) {
    console.error(`  ✗ ${name}`);
    console.error(`    Error: ${err.message}`);
    failCount++;
  }
}

// Sample data: reference DE 2024 / DE 2026, languages "EN 2024", "CS 2024", ...
function sample() {
  return {
    rows: SampleGenerator.getSampleData(),
    headers: SampleGenerator.getSampleHeaders(),
    mappings: Object.assign({ languages: {} }, SampleGenerator.getSampleDefaultMappings())
  };
}

// A small generic table: no ID/section columns, an almost empty column with the new version
// of one language and an unrelated notes column
function genericTable() {
  const headers = ['Reference old', 'Reference new', 'Czech', 'Czech new', 'Polish', 'Notes'];
  const rows = [
    { 'Reference old': 'Same text.', 'Reference new': 'Same text.', Czech: 'Stejný text.', 'Czech new': '', Polish: 'Ten sam tekst.', Notes: 'a' },
    { 'Reference old': 'Limit is 35 EUR.', 'Reference new': 'Limit is 50 EUR.', Czech: 'Limit je 35 EUR.', 'Czech new': 'Limit je 50 EUR.', Polish: 'Limit to 35 EUR.', Notes: '' },
    { 'Reference old': '', 'Reference new': 'Brand new paragraph.', Czech: '', 'Czech new': '', Polish: '', Notes: '' },
    { 'Reference old': 'Removed paragraph.', 'Reference new': '', Czech: 'Zrušený odstavec.', 'Czech new': '', Polish: 'Usunięty akapit.', Notes: '' }
  ];
  const mappings = {
    refOldCol: 'Reference old',
    refNewCol: 'Reference new',
    targetCols: ['Czech', 'Polish'],
    languages: {
      Czech: { label: 'Čeština', newCol: 'Czech new', newColName: '' },
      Polish: { label: 'Polština', newCol: '', newColName: 'Polish NEW' }
    }
  };
  return { headers, rows, mappings, rowNumbers: [2, 3, 5, 6] };
}

console.log('=== Spouštím testy jádra ===\n');

// 1. Klasifikace řádků
runTest('classifyRow: EMPTY pro prázdné buňky nebo pouze bílé znaky', () => {
  assert.strictEqual(ComparatorEngine.classifyRow('', ''), 'EMPTY');
  assert.strictEqual(ComparatorEngine.classifyRow('   ', '\n\t '), 'EMPTY');
  assert.strictEqual(ComparatorEngine.classifyRow(null, undefined), 'EMPTY');
});

runTest('classifyRow: ADDED / REMOVED / UNCHANGED / MODIFIED', () => {
  assert.strictEqual(ComparatorEngine.classifyRow('', 'Neuer Absatz.'), 'ADDED');
  assert.strictEqual(ComparatorEngine.classifyRow('Alter Absatz.', '   '), 'REMOVED');
  assert.strictEqual(ComparatorEngine.classifyRow('Text äöüß', 'Text äöüß'), 'UNCHANGED');
  assert.strictEqual(ComparatorEngine.classifyRow('Řádek 1\r\nŘádek 2', 'Řádek 1\nŘádek 2'), 'UNCHANGED');
  assert.strictEqual(ComparatorEngine.classifyRow('  Text na kraji  ', 'Text na kraji'), 'UNCHANGED', 'Okrajové mezery nejsou změna');
  assert.strictEqual(ComparatorEngine.classifyRow('Frist 14 Tage.', 'Frist 7 Tage.'), 'MODIFIED');
});

// 2. Diff
runTest('computeDiff: beze změny i přesná detekce přidaných a odebraných slov (diakritika)', () => {
  const same = ComparatorEngine.computeDiff('Stejný text.', 'Stejný text.');
  assert.strictEqual(same.stats.hasChanges, false);
  assert.strictEqual(same.stats.similarity, 1);

  const res = ComparatorEngine.computeDiff('Lhůta je 14 kalendářních dnů.', 'Lhůta je 7 pracovních dnů.');
  assert.strictEqual(res.stats.hasChanges, true);
  assert.ok(res.summary.removedTerms.includes('14') && res.summary.removedTerms.includes('kalendářních'));
  assert.ok(res.summary.addedTerms.includes('7') && res.summary.addedTerms.includes('pracovních'));
});

runTest('renderDiffHtml: escapuje HTML a používá <ins>/<del>; side-by-side vrací dva panely', () => {
  const inline = ComparatorEngine.renderDiffHtml(ComparatorEngine.computeDiff('Pozor na <script>alert(1)</script>', 'Pozor na <b>bezpečný</b> kód'), 'inline');
  assert.ok(!inline.includes('<script>') && inline.includes('&lt;'));
  assert.ok(inline.includes('<del class="diff-del"') && inline.includes('<ins class="diff-ins"'));

  const sbs = ComparatorEngine.renderDiffHtml(ComparatorEngine.computeDiff('Stara verze', 'Nova verze'), 'side-by-side');
  assert.ok(sbs.leftHtml.includes('<del') && !sbs.leftHtml.includes('<ins'));
  assert.ok(sbs.rightHtml.includes('<ins') && !sbs.rightHtml.includes('<del'));
});

// 3. Pracovní data jazyka
runTest('generateTargetStub: pravidla pro UNCHANGED, MODIFIED, ADDED, REMOVED', () => {
  const rows = [
    { 'R old': 'Unverändert.', 'R new': 'Unverändert.', CS: 'Nezměněno.' },
    { 'R old': 'Alt.', 'R new': 'Neu.', CS: 'Starý.' },
    { 'R old': '', 'R new': 'Ganz neu.', CS: '' },
    { 'R old': 'Gelöscht.', 'R new': '', CS: 'Smazáno.' }
  ];
  const stubs = ComparatorEngine.generateTargetStub(rows, 'CS', 'R old', 'R new');
  assert.deepStrictEqual(stubs.map(s => [s.classification, s.targetNew, s.status, s.isVerified]), [
    ['UNCHANGED', 'Nezměněno.', 'auto_copied', true],
    ['MODIFIED', '', 'pending_review', false],
    ['ADDED', '', 'pending_translation', false],
    ['REMOVED', '', 'removed', true]
  ]);
});

runTest('generateTargetStub: existující sloupec nové verze předvyplní text ke kontrole', () => {
  const g = genericTable();
  const stubs = ComparatorEngine.generateTargetStub(g.rows, 'Czech', g.mappings.refOldCol, g.mappings.refNewCol, 'Czech new');
  assert.strictEqual(stubs[1].targetNew, 'Limit je 50 EUR.', 'Text z existujícího sloupce je převzat');
  assert.strictEqual(stubs[1].status, 'imported');
  assert.strictEqual(stubs[1].isVerified, false, 'Převzatý text změněného odstavce je třeba zkontrolovat');
  assert.strictEqual(stubs[0].targetNew, 'Stejný text.', 'Prázdná buňka nezměněného odstavce se doplní starou verzí');
  assert.strictEqual(stubs[0].isVerified, true);
  assert.strictEqual(stubs[2].targetNew, '');
});

runTest('defaultNewColName / defaultLanguageLabel: odvození z čísla verze v referenčních sloupcích', () => {
  assert.strictEqual(ComparatorEngine.defaultNewColName('CS 2024', 'DE 2024', 'DE 2026'), 'CS 2026');
  assert.strictEqual(ComparatorEngine.defaultNewColName('Czech', 'Reference old', 'Reference new', ' (nová verze)'), 'Czech (nová verze)');
  assert.strictEqual(ComparatorEngine.defaultNewColName('Czech', 'Reference old', 'Reference new'), 'Czech (new)');
  assert.strictEqual(ComparatorEngine.defaultLanguageLabel('CS 2024', 'DE 2024'), 'CS');
  assert.strictEqual(ComparatorEngine.defaultLanguageLabel('Polština', 'EN old'), 'Polština');
});

// 4. Prompty
runTest('buildCopilotPrompt: obecný český prompt (bez zaměření na konkrétní dokument či jazyk)', () => {
  const prompt = ComparatorEngine.buildCopilotPrompt({
    targetLangName: 'Čeština', refLangName: 'EN', rowId: '#7', classification: 'MODIFIED',
    refOld: 'Limit is 35 EUR.', refNew: 'Limit is 50 EUR.', targetOld: 'Limit je 35 EUR.',
    diffResult: ComparatorEngine.computeDiff('Limit is 35 EUR.', 'Limit is 50 EUR.')
  });
  ['Čeština', '#7', 'Limit is 35 EUR.', 'Limit is 50 EUR.', 'Limit je 35 EUR.', 'POUZE minimální nutné změny', 'POUZE vraťte výsledný text', '(EN)']
    .forEach(part => assert.ok(prompt.includes(part), 'Chybí: ' + part));
  assert.ok(!/etick|kodex|němč|german|2024|2026/i.test(prompt), 'Prompt nesmí předpokládat konkrétní dokument, jazyk ani roky');

  const added = ComparatorEngine.buildCopilotPrompt({ targetLangName: 'Polština', rowId: '#9', classification: 'ADDED', refNew: 'Brand new paragraph.' });
  assert.ok(added.includes('ÚKOL: Přeložte nový odstavec') && added.includes('Brand new paragraph.'));
});

runTest('buildCopilotPrompt: anglický prompt a vlastní šablona se zástupnými symboly', () => {
  const en = ComparatorEngine.buildCopilotPrompt({
    promptLang: 'en', targetLangName: 'Spanish', rowId: '#4', classification: 'MODIFIED',
    refOld: 'a', refNew: 'b', targetOld: 'c', diffResult: ComparatorEngine.computeDiff('a', 'b')
  });
  assert.ok(en.includes('professional translator') && en.includes('CRITICAL INSTRUCTIONS') && en.includes('minimal necessary alterations'));

  const custom = ComparatorEngine.buildCopilotPrompt({
    targetLangName: 'CS', refLangName: 'EN', rowId: '#3', refOld: 'x', refNew: 'y', targetOld: 'z',
    customTemplate: '{TARGET_LANG}|{REF_LANG}|{ROW_ID}|{REF_OLD}|{REF_NEW}|{TARGET_OLD}|{UNKNOWN}'
  });
  assert.strictEqual(custom, 'CS|EN|#3|x|y|z|{UNKNOWN}');

  const withAux = ComparatorEngine.buildCopilotPrompt({
    targetLangName: 'Čeština', refLangName: 'Reference', rowId: '#7', classification: 'MODIFIED',
    refOld: 'Old text', refNew: 'New text', targetOld: 'Starý text',
    auxiliaryLangName: 'Polština', auxiliaryText: 'Nowy tekst po polsku.'
  });
  assert.ok(withAux.includes('POMOCNÝ PŘEKLAD PRO POROVNÁNÍ (Polština)') && withAux.includes('Nowy tekst po polsku.'));

  const withAuxEn = ComparatorEngine.buildCopilotPrompt({
    promptLang: 'en', targetLangName: 'Spanish', refLangName: 'Reference', rowId: '#7', classification: 'MODIFIED',
    refOld: 'Old text', refNew: 'New text', targetOld: 'Texto viejo',
    auxiliaryLangName: 'French', auxiliaryText: 'Nouveau texte en français.'
  });
  assert.ok(withAuxEn.includes('AUXILIARY TRANSLATION FOR REFERENCE (French)') && withAuxEn.includes('Nouveau texte en français.'));
});

runTest('buildVerificationPrompt: generování promptu pro ověření shodnosti a věrnosti překladu (CZ, EN, aux)', () => {
  const czPrompt = ComparatorEngine.buildVerificationPrompt({
    promptLang: 'cs',
    targetLangName: 'Čeština',
    refLangName: 'Reference',
    rowId: '#7',
    refNew: 'The payment must be completed within 14 days.',
    targetNew: 'Platba musí být dokončena do 14 dnů.',
    auxiliaryLangName: 'Polština',
    auxiliaryText: 'Płatność musi nastąpić w ciągu 14 dni.'
  });

  assert.ok(czPrompt.includes('revizor překladů'), 'Český úvod revizora');
  assert.ok(czPrompt.includes('Ověřte věcnou a významovou shodnost'), 'Úkol ověření shodnosti');
  assert.ok(czPrompt.includes('=== 1. REFERENČNÍ TEXT (Reference) ==='));
  assert.ok(czPrompt.includes('The payment must be completed within 14 days.'));
  assert.ok(czPrompt.includes('=== 2. KONTROLOVANÝ NOVÝ PŘEKLAD (Čeština) ==='));
  assert.ok(czPrompt.includes('Platba musí být dokončena do 14 dnů.'));
  assert.ok(czPrompt.includes('=== 3. POMOCNÝ REFERENČNÍ PŘEKLAD (Polština) ==='));
  assert.ok(czPrompt.includes('Płatność musi nastąpić w ciągu 14 dni.'));
  assert.ok(czPrompt.includes('ROZDÍLY A NEDOSTATKY'), 'Požadavek na výpis rozdílů a nedostatků');
  assert.ok(czPrompt.includes('NÁVRH ÚPRAVY'), 'Požadavek na návrh opravy');
  assert.ok(!/etick|kodex|němč|german|2024|2026/i.test(czPrompt), 'Obecný text bez zakázaných slov');

  const enPrompt = ComparatorEngine.buildVerificationPrompt({
    promptLang: 'en',
    targetLangName: 'Spanish',
    refLangName: 'English',
    rowId: 'Row 12',
    refNew: 'Contract is valid for 1 year.',
    targetNew: 'El contrato es válido por 1 año.'
  });
  assert.ok(enPrompt.includes('translation proofreader and language quality auditor'));
  assert.ok(enPrompt.includes('Verify the semantic equivalence'));
  assert.ok(enPrompt.includes('=== 1. REFERENCE TEXT (English) ==='));
  assert.ok(enPrompt.includes('=== 2. NEW TRANSLATION UNDER REVIEW (Spanish) ==='));
  assert.ok(enPrompt.includes('DISCREPANCIES & DEFECTS'));
  assert.ok(enPrompt.includes('SUGGESTED REVISION'));

  const custom = ComparatorEngine.buildVerificationPrompt({
    targetLangName: 'CS',
    refLangName: 'EN',
    rowId: '#1',
    refNew: 'Ref',
    targetNew: 'Trg',
    customTemplate: 'VERIFY:{TARGET_LANG}|{REF_LANG}|{ROW_ID}|{REF_NEW}|{TARGET_NEW}'
  });
  assert.strictEqual(custom, 'VERIFY:CS|EN|#1|Ref|Trg');
});

runTest('buildBatchCopilotPrompt: obsahuje jen změněné a nové odstavce se značkami čísel řádků a pomocným jazykem', () => {
  const g = genericTable();
  const stubs = ComparatorEngine.generateTargetStub(g.rows, 'Czech', g.mappings.refOldCol, g.mappings.refNewCol);
  const prompt = ComparatorEngine.buildBatchCopilotPrompt({
    promptLang: 'cs', targetCol: 'Czech', targetLangName: 'Čeština', refLangName: 'Reference new',
    auxiliaryCol: 'Polish', auxiliaryLangName: 'Polština',
    rows: g.rows, rowNumbers: g.rowNumbers, mappings: g.mappings, translations: { Czech: stubs }
  });
  assert.ok(prompt.includes('ODSTAVEC: [#3] (Typ: UPRAVENO'), 'Upravený odstavec na Excel řádku 3');
  assert.ok(prompt.includes('ODSTAVEC: [#5] (Typ: NOVĚ PŘIDÁNO)'), 'Nový odstavec na Excel řádku 5');
  assert.ok(prompt.includes('POMOCNÝ PŘEKLAD (Polština):'), 'Hromadný prompt obsahuje sekci pomocného jazyka');
  assert.ok(!prompt.includes('[#2]') && !prompt.includes('[#6]'), 'Nezměněné a zrušené odstavce nejsou v promptu');
  assert.ok(prompt.includes('Čeština') && !/etick|kodex|němč/i.test(prompt));

  const en = ComparatorEngine.buildBatchCopilotPrompt({
    promptLang: 'en', targetCol: 'Czech', targetLangName: 'Czech', rows: g.rows, rowNumbers: g.rowNumbers, mappings: g.mappings, translations: {}
  });
  assert.ok(en.includes('BATCH TRANSLATION UPDATE') && en.includes('PARAGRAPH: [#3]'));
});

// 5. Smart Paste
runTest('parseBatchTranslationResponse: značky [#N], "Řádek N:" i JSON; nespárované se nikam nezapíší', () => {
  const expected = [{ index: 0, id: '#2' }, { index: 1, id: '#3' }, { index: 2, id: '#5' }];
  const parsed = ComparatorEngine.parseBatchTranslationResponse('Úvod modelu\n\n[#3]\nText tři\n\n[#99]\nCizí text\n\nŘádek 5: Text pět', expected);
  assert.deepStrictEqual(parsed.map(p => [p.index, p.text]), [[1, 'Text tři'], [2, 'Text pět']]);
  assert.deepStrictEqual(parsed.unmatched, ['#99']);

  const json = ComparatorEngine.parseBatchTranslationResponse(JSON.stringify([{ id: '#2', text: 'A' }, { id: '#42', text: 'B' }]), expected);
  assert.deepStrictEqual(json.map(p => [p.index, p.text]), [[0, 'A']]);
  assert.deepStrictEqual(json.unmatched, ['#42']);

  assert.strictEqual(ComparatorEngine.idsMatch('1.2', '1.3'), false);
  assert.strictEqual(ComparatorEngine.idsMatch('#12', 'Row 12'), true);
});

runTest('Hromadný prompt → odpověď se stejnými značkami → Smart Paste přiřadí všechny odstavce správně', () => {
  const s = sample();
  const rowNumbers = s.rows.map((r, i) => i + 2);
  const stubs = ComparatorEngine.generateTargetStub(s.rows, 'CS 2024', s.mappings.refOldCol, s.mappings.refNewCol);
  const prompt = ComparatorEngine.buildBatchCopilotPrompt({
    promptLang: 'cs', targetCol: 'CS 2024', targetLangName: 'CS', rows: s.rows, rowNumbers, mappings: s.mappings, translations: { 'CS 2024': stubs }
  });
  const tags = [...prompt.matchAll(/^ODSTAVEC: \[(.+?)\] \(Typ/gm)].map(m => m[1]);
  const expectedIdx = stubs.map((st, i) => i).filter(i => ['MODIFIED', 'ADDED'].includes(stubs[i].classification));
  assert.strictEqual(tags.length, expectedIdx.length);

  const reply = tags.map((tag, k) => `[${tag}]\nPřeklad číslo ${k}`).join('\n\n');
  const parsed = ComparatorEngine.parseBatchTranslationResponse(reply, s.rows.map((r, i) => ({ index: i, id: ComparatorEngine.rowTag(rowNumbers, i) })));
  assert.deepStrictEqual(parsed.map(p => p.index), expectedIdx);
  assert.deepStrictEqual(parsed.unmatched, []);
});

// 6. Export do Excelu
runTest('prepareExportData (kompletní): překlad jde do existujícího sloupce nové verze, jinak do nového sloupce za starou verzí', () => {
  const g = genericTable();
  const translations = {
    Czech: ComparatorEngine.generateTargetStub(g.rows, 'Czech', g.mappings.refOldCol, g.mappings.refNewCol, 'Czech new'),
    Polish: ComparatorEngine.generateTargetStub(g.rows, 'Polish', g.mappings.refOldCol, g.mappings.refNewCol)
  };
  translations.Czech[2].targetNew = 'Úplně nový odstavec.';
  translations.Polish[1].targetNew = 'Limit to 50 EUR.';

  const res = ComparatorEngine.prepareExportData({ headers: g.headers, rows: g.rows, mappings: g.mappings, translations });
  assert.deepStrictEqual(res.headers, ['Reference old', 'Reference new', 'Czech', 'Czech new', 'Polish', 'Polish NEW', 'Notes'],
    'Existující sloupec se nezdvojí, nový sloupec je hned za starou verzí, ignorované sloupce zůstanou');
  assert.strictEqual(res.data[1]['Czech new'], 'Limit je 50 EUR.');
  assert.strictEqual(res.data[2]['Czech new'], 'Úplně nový odstavec.');
  assert.strictEqual(res.data[0]['Czech new'], 'Stejný text.', 'Nezměněný odstavec je doplněn do prázdné buňky');
  assert.strictEqual(res.data[1]['Polish NEW'], 'Limit to 50 EUR.');
  assert.strictEqual(res.data[0].Notes, 'a');
  assert.strictEqual(res.data[1].Czech, 'Limit je 35 EUR.', 'Stará verze zůstává');
});

runTest('prepareExportData (pouze nová verze): bez starých verzí a zrušených odstavců', () => {
  const g = genericTable();
  const res = ComparatorEngine.prepareExportData({ headers: g.headers, rows: g.rows, mappings: g.mappings, translations: {}, exportType: 'clean' });
  assert.deepStrictEqual(res.headers, ['Reference new', 'Notes', 'Czech new', 'Polish NEW']);
  assert.strictEqual(res.data.length, 3, 'Zrušený odstavec je vynechán');
  assert.strictEqual(res.data[1]['Czech new'], 'Limit je 50 EUR.');
  assert.strictEqual(res.data[1]['Polish NEW'], '', 'Nepřeložený změněný odstavec zůstane prázdný (žádný zastaralý text)');
  assert.strictEqual(ComparatorEngine.prepareExportData({ headers: g.headers, rows: g.rows, mappings: g.mappings, exportType: 'clean_2026' }).data.length, 3, 'Starý název režimu stále funguje');
});

runTest('prepareExportData: vytvořený sloupec nepřepíše existující sloupec stejného jména', () => {
  const headers = ['R1', 'R2', 'CS', 'CS (new)'];
  const rows = [{ R1: 'a', R2: 'b', CS: 'staré', 'CS (new)': 'JINÁ DATA' }];
  const mappings = { refOldCol: 'R1', refNewCol: 'R2', targetCols: ['CS'], languages: { CS: { label: 'CS', newCol: '', newColName: 'CS (new)' } } };
  const res = ComparatorEngine.prepareExportData({ headers, rows, mappings, translations: { CS: [{ classification: 'MODIFIED', targetNew: 'NOVÉ' }] } });
  assert.deepStrictEqual(res.headers, ['R1', 'R2', 'CS', 'CS (new) (2)', 'CS (new)']);
  assert.strictEqual(res.data[0]['CS (new)'], 'JINÁ DATA');
  assert.strictEqual(res.data[0]['CS (new) (2)'], 'NOVÉ');
});

runTest('Export tabulek + Excel adaptér: kompletní sešit zapsaný do .xlsx a zpět, včetně listu se stavem kontroly', () => {
  const g = genericTable();
  const result = exportWorkbook({
    rows: g.rows, headers: g.headers, mappings: g.mappings, rowNumbers: g.rowNumbers, translations: {}, exportType: 'comprehensive', uiLang: 'cs'
  });
  const wb = XLSX.read(XLSX.write(result.workbook, { type: 'buffer', bookType: 'xlsx' }), { type: 'buffer' });
  assert.deepStrictEqual(wb.SheetNames, ['Srovnani', 'Stav_prekladu']);
  const header = XLSX.utils.sheet_to_json(wb.Sheets.Srovnani, { header: 1 })[0];
  assert.deepStrictEqual(header, ['Reference old', 'Reference new', 'Czech', 'Czech new', 'Polish', 'Polish NEW', 'Notes']);
  const status = XLSX.utils.sheet_to_json(wb.Sheets.Stav_prekladu, { defval: '' });
  assert.deepStrictEqual(status.map(s => s['Řádek']), [2, 3, 5, 6], 'Stav je u skutečných čísel řádků');
  assert.deepStrictEqual(status.map(s => s['Čeština']), ['Ověřeno', 'Čeká na kontrolu', 'Chybí překlad', 'Zrušeno v nové verzi']);
  assert.strictEqual(result.unverifiedCount, 4, 'Čeština: 1 ke kontrole + 1 chybí; polština: 2 chybí');

  const en = exportWorkbook({ rows: g.rows, headers: g.headers, mappings: g.mappings, exportType: 'comprehensive', uiLang: 'en' });
  assert.deepStrictEqual(en.workbook.SheetNames, ['Comparison', 'Translation_Status']);
});

runTest('Export tabulek + Excel adaptér (částečný): číslo řádku, vybrané referenční sloupce, stará a nová verze jazyka a stav', () => {
  const g = genericTable();
  const translations = { Polish: ComparatorEngine.generateTargetStub(g.rows, 'Polish', g.mappings.refOldCol, g.mappings.refNewCol) };
  translations.Polish[1].targetNew = 'Limit to 50 EUR.';
  translations.Polish[1].isVerified = true;

  const result = exportWorkbook({
    rows: g.rows, headers: g.headers, mappings: g.mappings, rowNumbers: g.rowNumbers, translations,
    exportType: 'language', langCol: 'Polish', contextCols: ['Reference old', 'Reference new'], uiLang: 'cs'
  });
  assert.deepStrictEqual(result.workbook.SheetNames, ['Polština']);
  const sheet = result.workbook.Sheets['Polština'];
  assert.deepStrictEqual(XLSX.utils.sheet_to_json(sheet, { header: 1 })[0], ['Řádek', 'Reference old', 'Reference new', 'Polish', 'Polish NEW', 'Stav']);
  const data = XLSX.utils.sheet_to_json(sheet, { defval: '' });
  assert.deepStrictEqual(data[1], { 'Řádek': 3, 'Reference old': 'Limit is 35 EUR.', 'Reference new': 'Limit is 50 EUR.', Polish: 'Limit to 35 EUR.', 'Polish NEW': 'Limit to 50 EUR.', Stav: 'Ověřeno' });
  assert.strictEqual(data[2].Stav, 'Chybí překlad');

  // Nová verze jiného jazyka jako další sloupec: aktuální překlad, ne zastaralá buňka listu
  const czWork = ComparatorEngine.generateTargetStub(g.rows, 'Czech', g.mappings.refOldCol, g.mappings.refNewCol, 'Czech new');
  czWork[2].targetNew = 'Nový odstavec CZ';
  const withNew = ComparatorEngine.prepareLanguageExportData({
    rows: g.rows, headers: g.headers, mappings: g.mappings, translations: { Czech: czWork, Polish: translations.Polish },
    langCol: 'Polish', contextCols: [{ newOf: 'Czech' }, 'Czech new', 'Reference new', { newOf: 'Polish' }, { newOf: 'Nope' }]
  });
  assert.deepStrictEqual(withNew.headers, ['Řádek', 'Czech new', 'Reference new', 'Polish', 'Polish NEW', 'Stav'],
    'Sloupec nové verze se nezdvojí; vlastní jazyk a neznámé jazyky se vynechají');
  assert.deepStrictEqual(withNew.data.map(r => r['Czech new']), ['Stejný text.', 'Limit je 50 EUR.', 'Nový odstavec CZ', ''],
    'Hodnoty jsou z aktuální práce (doplněný nezměněný odstavec, nový text)');

  // Jazyk s existujícím sloupcem nové verze a jiné kontextové sloupce
  const cz = ComparatorEngine.prepareLanguageExportData({
    rows: g.rows, headers: g.headers, mappings: g.mappings, langCol: 'Czech', contextCols: ['Reference new', 'Polish', 'Czech']
  });
  assert.deepStrictEqual(cz.headers, ['Řádek', 'Reference new', 'Polish', 'Czech', 'Czech new', 'Stav'], 'Vlastní sloupce jazyka se jako kontext nezdvojí');
});

runTest('Přeskočené řádky: kompletní export je vloží na původní místo beze změny, ostatní exporty je vynechají', () => {
  const g = genericTable();
  // Řádek 4 (Excel) byl přeskočen – ve zpracovaných řádcích chybí
  const skippedRows = [{ rowNumber: 4, row: { 'Reference old': 'NADPIS', 'Reference new': 'NADPIS', Czech: 'NADPIS CZ', 'Czech new': 'stará data', Polish: '', Notes: 'pozn' } }];
  const tables = ComparatorEngine.prepareExportTables({
    rows: g.rows, headers: g.headers, mappings: g.mappings, rowNumbers: g.rowNumbers, translations: {}, skippedRows, uiLang: 'cs'
  }).tables;
  const main = tables[0].rows;
  assert.deepStrictEqual(main.map(r => r['Reference old']), ['Same text.', 'Limit is 35 EUR.', 'NADPIS', '', 'Removed paragraph.'], 'Řazení podle čísel řádků 2, 3, 4, 5, 6');
  assert.strictEqual(main[2]['Czech new'], 'stará data', 'Existující buňka přeskočeného řádku se nepřepíše');
  assert.strictEqual(main[2]['Polish NEW'], '', 'Nový sloupec přeskočeného řádku zůstane prázdný');
  assert.strictEqual(main[2].Notes, 'pozn');
  assert.deepStrictEqual(tables[1].rows.map(r => r['Řádek']), [2, 3, 4, 5, 6]);
  assert.deepStrictEqual([tables[1].rows[2]['Změna reference'], tables[1].rows[2]['Čeština']], ['Přeskočeno', 'Přeskočeno']);

  const clean = ComparatorEngine.prepareExportTables({ rows: g.rows, headers: g.headers, mappings: g.mappings, translations: {}, skippedRows, exportType: 'clean' }).tables;
  assert.ok(!clean[0].rows.some(r => r['Reference new'] === 'NADPIS'));
  const partial = ComparatorEngine.prepareExportTables({ rows: g.rows, headers: g.headers, mappings: g.mappings, translations: {}, skippedRows, exportType: 'language', langCol: 'Polish' }).tables;
  assert.strictEqual(partial[0].rows.length, 4);
});

// 7. Soubor projektu
runTest('StorageManager: uložení a načtení projektu zachová práci, jazyky, čísla řádků a nastavení', () => {
  const g = genericTable();
  const czech = ComparatorEngine.generateTargetStub(g.rows, 'Czech', g.mappings.refOldCol, g.mappings.refNewCol, 'Czech new');
  czech[1].isVerified = true;
  czech[1].status = 'verified';
  czech[2].targetNew = 'Rozpracovaný koncept.';
  czech[2].status = 'draft';
  const state = {
    fileName: 'data.xlsx', activeSheet: 'List1', hasHeaderRow: true, headers: g.headers, rows: g.rows, rowNumbers: g.rowNumbers,
    mappings: g.mappings, translations: { Czech: czech }, customPromptTemplate: 'X {TARGET_LANG}',
    pdfTemplate: { docTitle: 'Dokument', accentColor: '#0d9488' }, activeTargetCol: 'Czech',
    exportLangCol: 'Polish', partialContextCols: ['Reference new'], auxiliaryCol: 'Polish'
  };

  const json = StorageManager.serializeProjectJson(state);
  const file = JSON.parse(json);
  assert.strictEqual(file.app, 'InternationalisationManager');
  assert.strictEqual(file.translations.Czech[1].refOld, undefined, 'Kompaktní formát bez duplicitních textů');

  const restored = StorageManager.parseProjectJson(json);
  assert.deepStrictEqual(restored.mappings, g.mappings);
  assert.deepStrictEqual(restored.rowNumbers, g.rowNumbers);
  assert.deepStrictEqual(restored.pdfTemplate, state.pdfTemplate);
  assert.strictEqual(restored.exportLangCol, 'Polish');
  assert.strictEqual(restored.auxiliaryCol, 'Polish');
  assert.deepStrictEqual(restored.partialContextCols, ['Reference new']);
  const rehydrated = ComparatorEngine.rehydrateStubs(restored.rows, restored.translations.Czech, 'Czech', g.mappings.refOldCol, g.mappings.refNewCol, 'Czech new');
  assert.deepStrictEqual(rehydrated, czech);
});

runTest('StorageManager: soubory, které nejsou projektem nebo jsou poškozené, se odmítnou s kódem problému', () => {
  const valid = { app: 'InternationalisationManager', headers: ['A', 'B'], rows: [{ A: 'x', B: 'y' }], mappings: { refOldCol: 'A', refNewCol: 'B', targetCols: [] } };
  const codeOf = (content) => {
    try {
      StorageManager.parseProjectJson(typeof content === 'string' ? content : JSON.stringify(content));
      return 'OK';
    } catch (e) {
      return e.code;
    }
  };

  assert.strictEqual(codeOf(valid), 'OK');
  assert.strictEqual(codeOf({ ...valid, app: 'EthicCodeComparator' }), 'OK', 'Soubory pod původním názvem projektu se stále načtou');

  // Není to JSON (např. omylem vybraný sešit)
  assert.strictEqual(codeOf('{ not json'), 'INVALID_JSON');
  assert.strictEqual(codeOf('PK\u0003\u0004 binary spreadsheet'), 'INVALID_JSON');
  // JSON s jinou strukturou – nejde o projekt překladu
  assert.strictEqual(codeOf({}), 'NOT_A_PROJECT');
  assert.strictEqual(codeOf([1, 2, 3]), 'NOT_A_PROJECT');
  assert.strictEqual(codeOf('"text"'), 'NOT_A_PROJECT');
  assert.strictEqual(codeOf('null'), 'NOT_A_PROJECT');
  assert.strictEqual(codeOf({ hello: 'Ahoj', bye: 'Nashle' }), 'NOT_A_PROJECT', 'Překladový JSON jiné aplikace (klíč → text)');
  try {
    StorageManager.parseProjectJson(JSON.stringify({ name: 'pkg', version: '1.0.0', dependencies: {} }));
  } catch (e) {
    assert.strictEqual(e.params.FOUND, 'name, version, dependencies', 'Chyba uvádí, co soubor obsahuje');
  }
  assert.strictEqual(codeOf({ ...valid, app: 'JinaAplikace' }), 'FOREIGN_APP');
  // Projekt, ale poškozený
  assert.strictEqual(codeOf({ ...valid, rows: [] }), 'NO_ROWS');
  assert.strictEqual(codeOf({ ...valid, rows: 'x' }), 'NO_ROWS');
  assert.strictEqual(codeOf({ ...valid, rows: [1, 2] }), 'BAD_ROWS');
  assert.strictEqual(codeOf({ ...valid, rows: [['A', 'B']] }), 'BAD_ROWS');
  assert.strictEqual(codeOf({ ...valid, mappings: 'A,B' }), 'BAD_MAPPINGS');
  assert.strictEqual(codeOf({ ...valid, mappings: { refOldCol: 'A', refNewCol: 'B', targetCols: 'C' } }), 'BAD_MAPPINGS');
  assert.strictEqual(codeOf({ ...valid, mappings: { refOldCol: 'A', refNewCol: 'X' } }), 'MISSING_COLUMN');
  assert.strictEqual(codeOf({
    ...valid, headers: ['A', 'B', 'C'], rows: [{ A: 1, B: 2, C: 3 }],
    mappings: { refOldCol: 'A', refNewCol: 'B', targetCols: ['C'], languages: { C: { newCol: 'CHYBÍ' } } }
  }), 'MISSING_COLUMN', 'Neexistující sloupec nové verze');
});

runTest('StorageManager: načitatelný projekt s nesrovnalostmi vrátí upozornění a bezpečné hodnoty', () => {
  const ok = StorageManager.parseProjectJson(JSON.stringify({
    app: 'InternationalisationManager', headers: ['A', 'B'], rows: [{ A: 'x', B: 'y' }], mappings: { refOldCol: 'A', refNewCol: 'B', targetCols: [] }
  }));
  assert.deepStrictEqual(ok.warnings, [], 'Správný soubor nemá upozornění');

  const legacy = StorageManager.parseProjectJson('\uFEFF' + JSON.stringify({ rows: [{ A: 'x', B: 'y', C: 'z' }], mappings: { refOldCol: 'A', refNewCol: 'B', targetCols: ['C'] } }));
  assert.deepStrictEqual(legacy.headers, ['A', 'B', 'C']);
  assert.deepStrictEqual(legacy.mappings.languages, { C: { label: '', newCol: '', newColName: '' } });
  assert.deepStrictEqual(legacy.rowNumbers, [2]);
  assert.deepStrictEqual(legacy.warnings.map(w => w.code), ['NO_APP_ID'], 'Soubor bez označení aplikace');

  const odd = StorageManager.parseProjectJson(JSON.stringify({
    app: 'InternationalisationManager',
    version: '99.0.0',
    headers: ['A', 'B', 'C', 'D'],
    rows: [{ A: 'a', B: 'b', C: 'c', D: 'd' }, { A: 'a2', B: 'b2', C: 'c2', D: 'd2' }],
    rowNumbers: ['x', 'y'],
    mappings: { refOldCol: 'A', refNewCol: 'B', targetCols: ['C'], languages: { C: { label: 'Čeština' } } },
    translations: { C: [{ targetNew: 'jedna' }, 'chyba', { targetNew: 'navíc' }], D: [{ targetNew: 'neužitý' }] },
    pdfTemplate: 'blue',
    partialContextCols: 'A'
  }));
  assert.deepStrictEqual(odd.warnings.map(w => w.code), ['NEWER_VERSION', 'ROW_COUNT_MISMATCH', 'UNUSED_TRANSLATIONS', 'INVALID_TRANSLATIONS', 'INVALID_SETTINGS']);
  assert.deepStrictEqual(odd.warnings[1].params, { LANG: 'Čeština', COUNT: 3, ROWS: 2 });
  assert.strictEqual(odd.warnings[2].params.LIST, 'D');
  assert.deepStrictEqual(odd.translations, { C: [{ targetNew: 'jedna' }, null, { targetNew: 'navíc' }] }, 'Neplatné položky a nevybrané jazyky jsou vynechány');
  assert.deepStrictEqual(odd.rowNumbers, [2, 3], 'Neplatná čísla řádků se nahradí výchozími');
  assert.strictEqual(odd.pdfTemplate, null);
  assert.strictEqual(odd.partialContextCols, null);
});

runTest('StorageManager: uložený projekt se načte bez upozornění (skutečný serializér)', () => {
  const g = genericTable();
  const json = StorageManager.serializeProjectJson({
    fileName: 'data.xlsx', headers: g.headers, rows: g.rows, rowNumbers: g.rowNumbers, mappings: g.mappings,
    translations: { Czech: ComparatorEngine.generateTargetStub(g.rows, 'Czech', g.mappings.refOldCol, g.mappings.refNewCol, 'Czech new') },
    pdfTemplate: { docTitle: 'X' }, partialContextCols: ['Notes', { newOf: 'Czech' }],
    skippedRows: [{ rowNumber: 4, row: { 'Reference old': 'Nadpis' }, auto: false }]
  });
  const parsed = StorageManager.parseProjectJson(json);
  assert.deepStrictEqual(parsed.warnings, []);
  assert.deepStrictEqual(parsed.partialContextCols, ['Notes', { newOf: 'Czech' }], 'Výběr dalších sloupců včetně nových verzí');
  assert.deepStrictEqual(parsed.skippedRows, [{ rowNumber: 4, row: { 'Reference old': 'Nadpis' }, auto: false }], 'Přeskočené řádky jsou součástí projektu');
  const broken = StorageManager.parseProjectJson(JSON.stringify({ ...JSON.parse(json), skippedRows: [{ rowNumber: 'x' }] }));
  assert.deepStrictEqual(broken.skippedRows, []);
  assert.deepStrictEqual(broken.warnings.map(w => w.code), ['INVALID_SETTINGS']);
  assert.strictEqual(JSON.parse(json).version, StorageManager.PROJECT_VERSION);
});

runTest('rehydrateStubs: změna klasifikace (jiné referenční sloupce) zruší schválení, text zůstane', () => {
  const rows = [{ A: 'Alt', B: 'Neu', C: 'Alt', CS: 'Starý' }];
  const saved = [{ classification: 'MODIFIED', targetNew: 'Nový', status: 'verified', isVerified: true }];
  assert.strictEqual(ComparatorEngine.rehydrateStubs(rows, saved, 'CS', 'A', 'B')[0].isVerified, true);
  const changed = ComparatorEngine.rehydrateStubs(rows, saved, 'CS', 'A', 'C')[0];
  assert.strictEqual(changed.classification, 'UNCHANGED');
  assert.strictEqual(changed.isVerified, false);
  assert.strictEqual(changed.targetNew, 'Nový');
});

// 8. Shrnutí změn pro jeden jazyk
runTest('Shrnutí změn: reference – statistika, Markdown a Word s čísly řádků', () => {
  const s = sample();
  const summary = summaryOf({ rows: s.rows, mappings: s.mappings, documentName: 'Dokument', uiLang: 'cs' });
  assert.strictEqual(summary.isReference, true);
  assert.deepStrictEqual(summary.stats, { totalRows: 15, unchangedCount: 5, modifiedCount: 6, addedCount: 2, removedCount: 2, pctChanged: 67 });
  assert.ok(summary.markdown.includes('# Shrnutí změn – DE 2026 (Referenz) (stará → nová verze)'));
  assert.ok(summary.markdown.includes('| **Celkem odstavců** | 15 | 100% |'));
  assert.ok(summary.markdown.includes('### Řádek 5'), 'Odstavce jsou označeny číslem řádku');
  const generatedLines = summary.markdown.split('\n').filter(l => !l.startsWith('>') && !l.startsWith('- **'));
  assert.ok(!/etick|kodex|němč|Art\. 0/i.test(generatedLines.join('\n')), 'Žádné předpoklady o druhu dokumentu nebo nadpisech');
  assert.ok(summary.htmlDoc.includes('xmlns:w="urn:schemas-microsoft-com:office:word"'));
});

runTest('Shrnutí změn: vybraný jazyk – jeho stará a nová verze a stav kontroly', () => {
  const g = genericTable();
  const translations = { Polish: ComparatorEngine.generateTargetStub(g.rows, 'Polish', g.mappings.refOldCol, g.mappings.refNewCol) };
  translations.Polish[1].targetNew = 'Limit to 50 EUR.';
  translations.Polish[1].isVerified = true;

  const summary = summaryOf({
    rows: g.rows, mappings: g.mappings, translations, rowNumbers: g.rowNumbers, langCol: 'Polish', uiLang: 'cs'
  });
  assert.strictEqual(summary.language, 'Polština');
  assert.deepStrictEqual(summary.translationStatus, { verified: 2, review: 0, missing: 1 });
  assert.strictEqual(summary.modifiedItems[0].oldText, 'Limit to 35 EUR.', 'Upravený odstavec ukazuje starou verzi jazyka');
  assert.strictEqual(summary.modifiedItems[0].newText, 'Limit to 50 EUR.');
  assert.ok(summary.addedItems[0].missing, 'Nový odstavec bez překladu je vyznačen');
  assert.ok(summary.markdown.includes('Limit to 50 EUR.') && summary.markdown.includes('*Překlad chybí*'));
  assert.ok(!summary.markdown.includes('Limit is 50 EUR.'), 'Shrnutí jazyka neobsahuje referenční text');
  assert.ok(summary.htmlDoc.includes('Stav překladu'));

  const printable = SummaryExporters.renderSummaryHtml(ComparatorEngine.buildChangeSummary({ rows: g.rows, mappings: g.mappings, translations, langCol: 'Polish', uiLang: 'en' }), { printable: true });
  assert.ok(printable.includes('window.print()') && printable.includes('Change summary – Polština'));
});

runTest('Shrnutí změn: víceřádkový text zůstane celý v citaci a se zalomením ve Wordu', () => {
  const rows = [{ A: 'Zeile eins\nZeile zwei', B: 'Zeile eins\nZeile drei' }];
  const s = summaryOf({ rows, mappings: { refOldCol: 'A', refNewCol: 'B' } });
  assert.ok(s.markdown.includes('> Zeile eins\n> Zeile zwei'));
  assert.ok(s.markdown.includes('> Zeile eins\n> Zeile drei'));
  assert.ok(s.htmlDoc.includes('Zeile eins<br>Zeile zwei'));
});

// 9. Dokument pro tisk
runTest('Dokument pro tisk: dokument nové verze jazyka – čisté odstavce bez nadpisů, koncept při neúplném překladu', () => {
  const g = genericTable();
  const html = documentHtml({
    rows: g.rows, mappings: g.mappings, translations: {}, targetCol: 'Polish', documentName: 'Smlouva', uiLang: 'cs'
  });
  assert.ok(html.includes('size: A4') && html.includes('window.print()'));
  assert.ok(html.includes('<div class="doc-title">Smlouva</div>'), 'Název dokumentu z názvu souboru');
  assert.ok(html.includes('<div class="doc-subtitle">Polština</div>'));
  assert.ok(html.includes('>Ten sam tekst.</p>'), 'Nezměněný odstavec převzat');
  assert.ok(!html.includes('Limit to 35 EUR.'), 'Zastaralý text upraveného odstavce se neobjeví');
  assert.ok(!html.includes('Usunięty akapit.'), 'Zrušený odstavec chybí');
  assert.ok(html.includes('[Překlad chybí]') && html.includes('KONCEPT'));
  assert.ok(!/etick|kodex|compliance|certifik/i.test(html), 'Žádný text specifický pro etický kodex');
  assert.ok(!html.includes('<div class="sign-block">'), 'Podpisy jen na vyžádání');

  const ref = documentHtml({ rows: g.rows, mappings: g.mappings, targetCol: 'Reference new', uiLang: 'cs' });
  assert.ok(ref.includes('Brand new paragraph.') && !ref.includes('KONCEPT'));
});

runTest('Dokument pro tisk: vlastní hlavička (organizace, podtitul s {LANG}, verze, podpisy) a escapování', () => {
  const g = genericTable();
  const html = documentHtml({
    rows: g.rows, mappings: g.mappings, targetCol: 'Polish', uiLang: 'cs',
    template: {
      companyName: 'ACME a.s.', docTitle: 'Provozní řád', docSubtitle: 'Znění – {LANG}', versionLabel: '<b>v3</b>',
      effectiveDate: 'Platné od 1. 1.', accentColor: '#0d9488', introPreamble: 'Úvod.', showSignatures: true, signDept1: 'Ředitel', signDept2: 'Rada'
    }
  });
  ['ACME a.s.', 'Provozní řád', 'Znění – Polština', 'Platné od 1. 1.', '#0d9488', 'Úvod.', 'Ředitel', '<div class="sign-block">']
    .forEach(part => assert.ok(html.includes(part), 'Chybí: ' + part));
  assert.ok(html.includes('&lt;b&gt;v3&lt;/b&gt;') && !html.includes('<b>v3</b>'));
  assert.ok(!documentHtml({ rows: g.rows, mappings: g.mappings, targetCol: 'Polish', template: { accentColor: 'red;}body{x' } }).includes('red;}'),
    'Neplatná barva se nepoužije');
});

// 10. Integrace se vzorovým souborem
runTest('Integrace: ukázkový .xlsx – stuby a kompletní export', () => {
  const wb = XLSX.readFile(path.join(__dirname, '../sample/ukazkovy_dokument.xlsx'));
  const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]);
  assert.strictEqual(rows.length, 15);
  const s = sample();
  const stubs = ComparatorEngine.generateTargetStub(rows, 'CS 2024', s.mappings.refOldCol, s.mappings.refNewCol);
  stubs[3].targetNew = 'Dary do 50 eur.';
  const res = ComparatorEngine.prepareExportData({ rows, headers: Object.keys(rows[0]), mappings: s.mappings, translations: { 'CS 2024': stubs } });
  assert.strictEqual(res.data.length, 15);
  assert.strictEqual(res.data[3]['CS 2026'], 'Dary do 50 eur.', 'Výchozí název nového sloupce odvozen z čísla verze reference');
});

// 11. Adaptéry pro import a export
runTest('Importers: Excel adaptér je výchozí, přípona i MIME typ CSV se rozpoznají', () => {
  assert.ok(Importers.list().some(a => a.id === 'excel'));
  assert.strictEqual(Importers.findForFile('data.XLSX').id, 'excel');
  assert.strictEqual(Importers.findForFile('export', 'text/csv').id, 'excel');
  assert.strictEqual(Importers.findForFile('data.json'), null);
  assert.ok(Importers.acceptAttribute().split(',').includes('.csv'));
});

runTest('Excel adaptér: list se záhlavím, bez záhlaví a CSV se převedou na model tabulky (oříznutí, čísla řádků)', () => {
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([
    ['2024', 'Text', ''],
    ['a', '  okraj  ', 'x'],
    [],
    ['b', 'druhý', 'y']
  ]), 'List1');
  const buf = XLSX.write(wb, { type: 'array', bookType: 'xlsx' });
  const source = ExcelImporter.parse(buf, { fileName: 't.xlsx' });
  assert.deepStrictEqual(source.sheetNames, ['List1']);

  const withHeader = Importers.readSheet(source, 'List1', { hasHeaderRow: true });
  assert.deepStrictEqual(withHeader.headers, ['2024', 'Text', '__EMPTY'], 'Pořadí sloupců podle listu, i číselný název');
  assert.deepStrictEqual(withHeader.rowNumbers, [2, 4], 'Prázdný řádek se přeskočí, čísla řádků odpovídají Excelu');
  assert.strictEqual(withHeader.rows[0].Text, 'okraj', 'Bílé znaky na okrajích odstraněny');
  assert.strictEqual(withHeader.columnLetters.__EMPTY, 'C');

  const noHeader = Importers.readSheet(source, 'List1', { hasHeaderRow: false });
  assert.deepStrictEqual(noHeader.headers, ['A', 'B', 'C']);
  assert.deepStrictEqual(noHeader.rowNumbers, [1, 2, 4]);

  const csv = ExcelImporter.parse('\uFEFFR;N\nGröße;Nová\n', { fileName: 'x.csv' });
  assert.deepStrictEqual(Importers.readSheet(csv, csv.sheetNames[0]).rows, [{ R: 'Größe', N: 'Nová' }]);
  // Files are read as bytes; CSV is decoded by the adapter (UTF-8 or a legacy code page)
  assert.strictEqual(ExcelImporter.readAs, 'arraybuffer');
});

runTest('Importers: nový adaptér lze zaregistrovat a má přednost; model tabulky se normalizuje', () => {
  const custom = {
    id: 'test-json',
    label: 'JSON test',
    extensions: ['.json', '.csv'],
    readAs: 'text',
    parse: function (content) {
      const data = JSON.parse(content);
      return { sheetNames: ['Strings'], readSheet: () => ({ rows: data.map(d => ({ key: d.key, text: '  ' + d.text + ' ' })) }) };
    }
  };
  Importers.register(custom);
  try {
    assert.strictEqual(Importers.findForFile('strings.json').id, 'test-json');
    assert.strictEqual(Importers.findForFile('data.csv').id, 'test-json', 'Později registrovaný adaptér má přednost');
    const source = custom.parse('[{"key":"a","text":"Ahoj"}]');
    const table = Importers.readSheet(source, 'Strings');
    assert.deepStrictEqual(table, { headers: ['key', 'text'], rows: [{ key: 'a', text: 'Ahoj' }], rowNumbers: [2], columnLetters: {} });
    assert.throws(() => Importers.register({ id: 'bad' }), /parse/);
  } finally {
    Importers.unregister('test-json');
  }
  assert.strictEqual(Importers.findForFile('data.csv').id, 'excel');
});

runTest('Exporters: výchozí adaptéry podle skupin a registrace nového formátu', () => {
  assert.deepStrictEqual(Exporters.list('table').map(a => a.id), ['xlsx']);
  assert.strictEqual(Exporters.getDefault('table').id, 'xlsx');
  assert.strictEqual(Exporters.getDefault('document').id, 'html-print');
  assert.strictEqual(Exporters.getDefault('summary').id, 'summary-markdown');
  assert.deepStrictEqual(Exporters.list('summary').map(a => a.id), ['summary-markdown', 'summary-word', 'summary-print']);

  const csvAdapter = {
    id: 'test-csv', label: 'CSV', group: 'table',
    export: (payload) => ({ content: payload.tables[0].headers.join(';'), mimeType: 'text/csv', extension: '.csv' })
  };
  Exporters.register(csvAdapter);
  try {
    assert.deepStrictEqual(Exporters.list('table').map(a => a.id), ['xlsx', 'test-csv']);
    assert.strictEqual(Exporters.getDefault('table').id, 'xlsx', 'Výchozí formát zůstává');
    const g = genericTable();
    const tables = ComparatorEngine.prepareExportTables({ rows: g.rows, headers: g.headers, mappings: g.mappings, exportType: 'clean' }).tables;
    assert.strictEqual(Exporters.get('test-csv').export({ tables }).content, 'Reference new;Notes;Czech new;Polish NEW');
    assert.throws(() => Exporters.register({ id: 'bad', group: 'table' }), /export/);
  } finally {
    Exporters.unregister('test-csv');
  }
  assert.strictEqual(Exporters.safeFilePart('Dokument – změny (CS)'), 'Dokument_změny_CS');
});

runTest('Excel exportér: jeden list na tabulku, platné a unikátní názvy listů, šířky a filtr', () => {
  const content = Exporters.get('xlsx').export({
    tables: [
      { name: 'Jazyk: CS/EN [nový] s velmi dlouhým názvem listu', headers: ['A', 'B'], rows: [{ A: 1, B: 'dlouhý text' }] },
      { name: 'Jazyk: CS/EN [nový] s velmi dlouhým názvem listu', headers: ['X'], rows: [{ X: 'y' }] }
    ]
  });
  assert.strictEqual(content.extension, '.xlsx');
  const wb = XLSX.read(content.content, { type: 'array' });
  assert.strictEqual(wb.SheetNames.length, 2);
  assert.ok(wb.SheetNames.every(n => n.length <= 31 && !/[\\/?*[\]:]/.test(n)), JSON.stringify(wb.SheetNames));
  assert.notStrictEqual(wb.SheetNames[0], wb.SheetNames[1]);
  assert.deepStrictEqual(XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]), [{ A: 1, B: 'dlouhý text' }]);
});

runTest('Modely dokumentu a shrnutí neobsahují formátování – vykreslí je až adaptér', () => {
  const g = genericTable();
  const model = ComparatorEngine.buildDocumentModel({ rows: g.rows, mappings: g.mappings, rowNumbers: g.rowNumbers, targetCol: 'Polish', documentName: 'Smlouva' });
  assert.deepStrictEqual(model.paragraphs.map(p => [p.rowNumber, p.missing]), [[2, false], [3, true], [5, true]]);
  assert.strictEqual(model.title, 'Smlouva');
  assert.strictEqual(model.isDraft, true);
  assert.ok(!JSON.stringify(model).includes('<'), 'Model dokumentu je bez HTML');

  const summary = ComparatorEngine.buildChangeSummary({ rows: g.rows, mappings: g.mappings, rowNumbers: g.rowNumbers });
  assert.strictEqual(summary.markdown, undefined);
  assert.deepStrictEqual(summary.stats, { totalRows: 4, unchangedCount: 1, modifiedCount: 1, addedCount: 1, removedCount: 1, pctChanged: 75 });
  const md = Exporters.get('summary-markdown').export(summary);
  assert.strictEqual(md.extension, '.md');
  const word = Exporters.get('summary-word').export(summary);
  assert.ok(word.content.startsWith('\ufeff<!DOCTYPE html>') && word.extension === '.doc');
  assert.strictEqual(Exporters.get('html-file').export(model).extension, '.html');
});

// 12. Podpora pro nově přidané jazyky bez staré verze
runTest('Nový jazyk: isNewLanguage rozpozná __new__ id i příznak v konfiguraci', () => {
  const mappings = {
    languages: {
      '__new__French': { label: 'Francouzština', isNewLanguage: true },
      'Czech': { label: 'Čeština' }
    }
  };
  assert.strictEqual(ComparatorEngine.isNewLanguage(mappings, '__new__French'), true);
  assert.strictEqual(ComparatorEngine.isNewLanguage(mappings, 'Czech'), false);
  assert.strictEqual(ComparatorEngine.isNewLanguage(mappings, '__new__custom'), true);
});

runTest('Nový jazyk: generateTargetStub inicializuje odstavce pro překlad bez staré verze', () => {
  const g = genericTable();
  const stubs = ComparatorEngine.generateTargetStub(
    g.rows,
    '__new__French',
    g.mappings.refOldCol,
    g.mappings.refNewCol,
    '',
    true
  );
  assert.strictEqual(stubs.length, 4);
  assert.strictEqual(stubs[0].targetOld, '');
  assert.strictEqual(stubs[0].classification, 'UNCHANGED');
  assert.strictEqual(stubs[0].isVerified, false, 'U nového jazyka musí i nezměněná reference projít překladem');
  assert.strictEqual(stubs[0].status, 'pending_translation');

  assert.strictEqual(stubs[1].classification, 'MODIFIED');
  assert.strictEqual(stubs[1].isVerified, false);
  assert.strictEqual(stubs[1].status, 'pending_translation');

  assert.strictEqual(stubs[2].classification, 'ADDED');
  assert.strictEqual(stubs[2].isVerified, false);
  assert.strictEqual(stubs[2].status, 'pending_translation');

  assert.strictEqual(stubs[3].classification, 'REMOVED');
  assert.strictEqual(stubs[3].isVerified, true, 'Zrušený odstavec se nepřekládá');
});

runTest('Nový jazyk: buildCopilotPrompt vytvoří prostý překlad bez diffu a historie změn', () => {
  const promptCs = ComparatorEngine.buildCopilotPrompt({
    promptLang: 'cs',
    targetLangName: 'Francouzština',
    refLangName: 'Reference',
    rowId: '#2',
    classification: 'MODIFIED',
    refOld: 'Limit is 35 EUR.',
    refNew: 'Limit is 50 EUR.',
    targetOld: '',
    diffResult: ComparatorEngine.computeDiff('Limit is 35 EUR.', 'Limit is 50 EUR.'),
    isNewLanguage: true
  });
  assert.ok(promptCs.includes('Přeložte text do jazyka: Francouzština'), 'Obsahuje instrukci k překladu');
  assert.ok(promptCs.includes('Limit is 50 EUR.'), 'Obsahuje nový referenční text');
  assert.ok(!promptCs.includes('STARÉ ZNĚNÍ'), 'Neobsahuje staré referenční znění');
  assert.ok(!promptCs.includes('PŮVODNÍ SCHVÁLENÝ'), 'Neobsahuje původní překlad');
  assert.ok(!promptCs.includes('PŘEHLED ZMĚN'), 'Neobsahuje přehled změn');

  const promptEn = ComparatorEngine.buildCopilotPrompt({
    promptLang: 'en',
    targetLangName: 'French',
    refLangName: 'Reference',
    rowId: '#2',
    classification: 'MODIFIED',
    refOld: 'Limit is 35 EUR.',
    refNew: 'Limit is 50 EUR.',
    targetOld: '',
    diffResult: ComparatorEngine.computeDiff('Limit is 35 EUR.', 'Limit is 50 EUR.'),
    isNewLanguage: true
  });
  assert.ok(promptEn.includes('Translate the text into language: French'));
  assert.ok(promptEn.includes('Limit is 50 EUR.'));
  assert.ok(!promptEn.includes('OLD REFERENCE TEXT'));
});

runTest('Nový jazyk: export kompletní i nová verze s kombinací aktualizovaných a nových jazyků', () => {
  const g = genericTable();
  const newLangId = '__new__French';
  const mappings = {
    refOldCol: g.mappings.refOldCol,
    refNewCol: g.mappings.refNewCol,
    targetCols: ['Czech', newLangId],
    languages: {
      Czech: g.mappings.languages.Czech,
      [newLangId]: { label: 'Francouzština', newCol: '', newColName: 'French', isNewLanguage: true }
    }
  };
  const translations = {
    Czech: [
      { targetNew: 'Stejný text.', isVerified: true, classification: 'UNCHANGED' },
      { targetNew: 'Limit je 50 EUR.', isVerified: true, classification: 'MODIFIED' },
      { targetNew: 'Zbrusu nový.', isVerified: true, classification: 'ADDED' },
      { targetNew: '', isVerified: true, classification: 'REMOVED' }
    ],
    [newLangId]: [
      { targetNew: 'Même texte.', isVerified: true, classification: 'UNCHANGED' },
      { targetNew: 'La limite est de 50 EUR.', isVerified: true, classification: 'MODIFIED' },
      { targetNew: 'Tout nouveau paragraphe.', isVerified: true, classification: 'ADDED' },
      { targetNew: '', isVerified: true, classification: 'REMOVED' }
    ]
  };

  // 1. Kompletní export: Czech má starou + novou verzi, Francouzština má jen 1 sloupec přidaný na konec
  const comp = ComparatorEngine.prepareExportData({
    rows: g.rows,
    headers: g.headers,
    mappings: mappings,
    translations: translations,
    exportType: 'comprehensive'
  });
  assert.ok(comp.headers.includes('French'), 'Export obsahuje sloupec nového jazyka');
  assert.strictEqual(comp.headers.filter(h => h === 'French').length, 1, 'Pouze jeden sloupec pro nový jazyk');
  assert.strictEqual(comp.data[0]['French'], 'Même texte.');
  assert.strictEqual(comp.data[1]['French'], 'La limite est de 50 EUR.');
  assert.strictEqual(comp.data[2]['French'], 'Tout nouveau paragraphe.');
  assert.strictEqual(comp.data[3]['French'], '');

  // 2. Čistý export (jen nová verze - vynechá REMOVED a staré verze)
  const clean = ComparatorEngine.prepareExportData({
    rows: g.rows,
    headers: g.headers,
    mappings: mappings,
    translations: translations,
    exportType: 'clean'
  });
  assert.strictEqual(clean.data.length, 3, 'Vynechán 1 zrušený řádek');
  assert.deepStrictEqual(clean.headers, ['Reference new', 'Polish', 'Notes', 'Czech new', 'French']);
  assert.strictEqual(clean.data[0]['French'], 'Même texte.');
  assert.strictEqual(clean.data[1]['French'], 'La limite est de 50 EUR.');

  // 3. Částečný export pro nový jazyk (omits targetOld)
  const partial = ComparatorEngine.prepareLanguageExportData({
    rows: g.rows,
    headers: g.headers,
    mappings: mappings,
    translations: translations,
    targetCol: newLangId,
    contextCols: [mappings.refNewCol]
  });
  assert.deepStrictEqual(partial.headers, ['Řádek', 'Reference new', 'French', 'Stav']);
  assert.strictEqual(partial.data[0]['French'], 'Même texte.');
});

runTest('Nový jazyk: StorageManager uloží a načte projekt s novým jazykem bez MISSING_COLUMN', () => {
  const g = genericTable();
  const newLangId = '__new__French';
  const state = {
    fileName: 'test.xlsx',
    rows: g.rows,
    headers: g.headers,
    mappings: {
      refOldCol: g.mappings.refOldCol,
      refNewCol: g.mappings.refNewCol,
      targetCols: ['Czech', newLangId],
      languages: {
        Czech: g.mappings.languages.Czech,
        [newLangId]: { label: 'Francouzština', newCol: '', newColName: 'French', isNewLanguage: true }
      }
    },
    translations: {
      Czech: [],
      [newLangId]: []
    }
  };

  const json = StorageManager.serializeProjectJson(state);
  const parsed = StorageManager.parseProjectJson(json);
  assert.ok(parsed.mappings.targetCols.includes(newLangId));
  assert.strictEqual(parsed.mappings.languages[newLangId].isNewLanguage, true);
  assert.strictEqual(parsed.mappings.languages[newLangId].label, 'Francouzština');
  assert.strictEqual(parsed.mappings.languages[newLangId].newColName, 'French');
});

// Helper pro vytvoření minimálního .docx v paměti pro testy
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

runTest('WordImporter: načtení .docx vytvoří 1 referenční sloupec s odstavci', () => {
  const buf = createTestDocx(['První odstavec textu.', 'Druhý odstavec s & speciálním znakem.']);
  const source = WordImporter.parse(buf, { fileName: 'predpis.docx' });
  assert.strictEqual(source.sheetNames.length, 1);
  const table = Importers.readSheet(source, source.sheetNames[0]);
  assert.deepStrictEqual(table.headers, ['Reference']);
  assert.strictEqual(table.rows.length, 2);
  assert.strictEqual(table.rows[0]['Reference'], 'První odstavec textu.');
  assert.strictEqual(table.rows[1]['Reference'], 'Druhý odstavec s & speciálním znakem.');
  assert.deepStrictEqual(table.rowNumbers, [1, 2]);
});

runTest('WordImporter: načtení .doc (HTML formát Wordu) extrahuje odstavce', () => {
  const html = '<html xmlns:w="urn:schemas-microsoft-com:office:word"><body><p>Odstavec 1.</p><p>Odstavec 2.</p></body></html>';
  const source = WordImporter.parse(Buffer.from(html, 'utf8'), { fileName: 'dokument.doc' });
  const table = Importers.readSheet(source, source.sheetNames[0]);
  assert.deepStrictEqual(table.headers, ['Reference']);
  assert.strictEqual(table.rows.length, 2);
  assert.strictEqual(table.rows[0]['Reference'], 'Odstavec 1.');
  assert.strictEqual(table.rows[1]['Reference'], 'Odstavec 2.');
});

runTest('TextImporter: načtení .txt souboru odděleného prázdnými řádky i jednotlivými řádky', () => {
  const txt1 = 'Blok 1 řádek 1.\nBlok 1 řádek 2.\n\nBlok 2.';
  const s1 = TextImporter.parse(txt1, { fileName: 'text1.txt' });
  const t1 = Importers.readSheet(s1, s1.sheetNames[0]);
  assert.strictEqual(t1.rows.length, 2);
  assert.strictEqual(t1.rows[0]['Reference'], 'Blok 1 řádek 1.\nBlok 1 řádek 2.');
  assert.strictEqual(t1.rows[1]['Reference'], 'Blok 2.');

  const txt2 = 'Řádek 1.\nŘádek 2.\nŘádek 3.';
  const s2 = TextImporter.parse(txt2, { fileName: 'text2.txt' });
  const t2 = Importers.readSheet(s2, s2.sheetNames[0]);
  assert.strictEqual(t2.rows.length, 3);
  assert.strictEqual(t2.rows[0]['Reference'], 'Řádek 1.');
  assert.strictEqual(t2.rows[2]['Reference'], 'Řádek 3.');
});

runTest('Importers: registr obsahuje Excel, Word i Text a acceptAttribute zahrnuje všechny přípony', () => {
  const exts = Importers.acceptAttribute().split(',');
  ['.xlsx', '.csv', '.docx', '.doc', '.txt'].forEach(ext => {
    assert.ok(exts.includes(ext), 'Přípona ' + ext + ' je v acceptAttribute');
  });
  assert.strictEqual(Importers.findForFile('test.docx').id, 'word');
  assert.strictEqual(Importers.findForFile('test.doc').id, 'word');
  assert.strictEqual(Importers.findForFile('test.txt').id, 'text');
  assert.strictEqual(Importers.findForFile('test.xlsx').id, 'excel');
});

runTest('Export s jedinou referencí (refOldCol = ""): kompletní, čistý i částečný export', () => {
  const rows = [
    { Reference: 'First article.' },
    { Reference: 'Second article.' }
  ];
  const newLang = '__new__Polstina';
  const mappings = {
    refOldCol: '',
    refNewCol: 'Reference',
    targetCols: [newLang],
    languages: {
      [newLang]: { label: 'Polština', newCol: '', newColName: 'PL', isNewLanguage: true }
    }
  };
  const translations = {
    [newLang]: [
      { targetNew: 'Pierwszy artykuł.', isVerified: true, classification: 'ADDED' },
      { targetNew: 'Drugi artykuł.', isVerified: true, classification: 'ADDED' }
    ]
  };

  // Kompletní export
  const comp = ComparatorEngine.prepareExportData({
    rows: rows,
    headers: ['Reference'],
    mappings: mappings,
    translations: translations,
    exportType: 'comprehensive'
  });
  assert.deepStrictEqual(comp.headers, ['Reference', 'PL']);
  assert.strictEqual(comp.data.length, 2);
  assert.strictEqual(comp.data[0]['PL'], 'Pierwszy artykuł.');

  // Čistý export
  const clean = ComparatorEngine.prepareExportData({
    rows: rows,
    headers: ['Reference'],
    mappings: mappings,
    translations: translations,
    exportType: 'clean'
  });
  assert.deepStrictEqual(clean.headers, ['Reference', 'PL']);
  assert.strictEqual(clean.data.length, 2);

  // Částečný export pro vybraný jazyk
  const partial = ComparatorEngine.prepareLanguageExportData({
    rows: rows,
    headers: ['Reference'],
    mappings: mappings,
    translations: translations,
    targetCol: newLang,
    contextCols: ['Reference']
  });
  assert.deepStrictEqual(partial.headers, ['Řádek', 'Reference', 'PL', 'Stav']);
  assert.strictEqual(partial.data[0]['PL'], 'Pierwszy artykuł.');
});
runTest('i18n: všechny klíče z app.js a index.html existují v CS i EN; texty jsou obecné', () => {
  const appSrc = fs.readFileSync(path.join(__dirname, '../js/app.js'), 'utf8');
  const html = fs.readFileSync(path.join(__dirname, '../index.html'), 'utf8');
  const used = new Set([...appSrc.matchAll(/\bt\('([a-zA-Z0-9]+)'/g)].map(m => m[1]));
  [...html.matchAll(/data-i18n(?:-placeholder|-title)?="([a-zA-Z0-9]+)"/g)].forEach(m => used.add(m[1]));
  // keys passed indirectly
  ['toastExportUnverifiedClean', 'toastExportUnverifiedLanguage', 'alertDifferentRefCols', 'alertSelectTargetLangs'].forEach(k => used.add(k));

  assert.deepStrictEqual([...used].filter(k => !(k in I18n.translations.cs)), [], 'Chybějící české klíče');
  assert.deepStrictEqual([...used].filter(k => !(k in I18n.translations.en)), [], 'Chybějící anglické klíče');
  assert.deepStrictEqual(Object.keys(I18n.translations.cs).sort(), Object.keys(I18n.translations.en).sort(), 'cs/en parita');

  const allTexts = JSON.stringify(I18n.translations) + html.replace(/<script[\s\S]*?<\/script>/g, '');
  const specific = allTexts.match(/etick|ethic|kodex|němčin|german|deutsch|2024|2026/gi) || [];
  assert.deepStrictEqual(specific, [], 'Rozhraní nesmí předpokládat etický kodex, němčinu ani konkrétní roky');
});

// ---------------------------------------------------------------------------
// Import / export regressions
// ---------------------------------------------------------------------------

const BOM = String.fromCharCode(0xFEFF);

function fixtureBuffer(name) {
  const b = fs.readFileSync(path.join(__dirname, 'fixtures', name));
  return b.buffer.slice(b.byteOffset, b.byteOffset + b.length);
}

// Text of the Word fixtures (saved by Word as .doc 97-2003, .rtf and filtered HTML)
const WORD_FIXTURE_PARAGRAPHS = [
  'První odstavec: žluťoučký kůň úpěl ďábelské ódy.',
  'Druhý odstavec & <tagy> s "uvozovkami" — pomlčka.',
  'Třetí odstavec\ns ručním zalomením.',
  '2026 – rok v poli',
  'Buňka A',
  'Buňka B'
];

runTest('CSV: Windows-1250 se středníky, texty zůstanou textem (0012, datum)', () => {
  const cp1250 = { 'á': 0xE1, 'č': 0xE8, 'ř': 0xF8, 'ž': 0x9E, 'ť': 0x9D, 'ů': 0xF9, 'ý': 0xFD, 'í': 0xED };
  const text = 'Název;Kód;Datum\r\nžluťoučký řádek;0012;1.2.2024\r\n';
  const bytes = Uint8Array.from([...text].map(ch => cp1250[ch] || ch.charCodeAt(0)));
  const source = ExcelImporter.parse(bytes.buffer, { fileName: 'export.csv' });
  const table = Importers.readSheet(source, source.sheetNames[0], { hasHeaderRow: true });
  assert.deepStrictEqual(table.headers, ['Název', 'Kód', 'Datum']);
  assert.deepStrictEqual(table.rows[0], { 'Název': 'žluťoučký řádek', 'Kód': '0012', 'Datum': '1.2.2024' });

  // UTF-8 with BOM read as bytes (as the browser does) and as a string (custom callers)
  const utf8 = Buffer.from(BOM + text, 'utf8');
  const fromBytes = ExcelImporter.parse(utf8.buffer.slice(utf8.byteOffset, utf8.byteOffset + utf8.length), { fileName: 'export.csv' });
  assert.strictEqual(Importers.readSheet(fromBytes, fromBytes.sheetNames[0]).rows[0]['Název'], 'žluťoučký řádek');
  const fromString = ExcelImporter.parse(BOM + text, { fileName: 'export.csv' });
  assert.strictEqual(Importers.readSheet(fromString, fromString.sheetNames[0]).rows[0]['Kód'], '0012');
});

runTest('Importers.decodeText: UTF-8, BOM, UTF-16 a Windows-1250', () => {
  assert.strictEqual(Importers.decodeText(Buffer.from(BOM + 'čeština', 'utf8')), 'čeština');
  assert.strictEqual(Importers.decodeText(Buffer.from(BOM + 'čeština', 'utf16le')), 'čeština');
  assert.strictEqual(Importers.decodeText(Uint8Array.from([0xE8, 0x65, 0x9A, 0x74, 0x69, 0x6E, 0x61])), 'čeština');
  assert.strictEqual(Importers.decodeText(BOM + 'text'), 'text');
});

runTest('TextImporter: soubor ve Windows-1250 se načte s diakritikou', () => {
  const bytes = Uint8Array.from([0x50, 0xF8, 0xED, 0x6C, 0x69, 0x9A, 0x0D, 0x0A, 0x0D, 0x0A, 0x9E, 0x6C, 0x75, 0x9D, 0x6F, 0x75, 0xE8, 0x6B, 0xFD]);
  const source = TextImporter.parse(bytes.buffer, { fileName: 'text.txt' });
  assert.deepStrictEqual(Importers.readSheet(source, source.sheetNames[0]).rows.map(r => r.Reference), ['Příliš', 'žluťoučký']);
});

runTest('Excel: buňky s datem (formát d.m.yyyy z Excelu) se načtou jako datum ISO, ne jako číslo', () => {
  const source = ExcelImporter.parse(fixtureBuffer('excel-dates.xlsx'), { fileName: 'excel-dates.xlsx' });
  const row = Importers.readSheet(source, source.sheetNames[0], { hasHeaderRow: true }).rows[0];
  assert.deepStrictEqual(row, { Text: 'Odstavec', Platnost: '2024-02-01', 'Čas': '2024-02-01 12:00' });
});

runTest('WordImporter: skutečné soubory Wordu – .doc 97-2003, .rtf a HTML .doc ve Windows-1250', () => {
  ['word97.doc', 'word.rtf', 'word-html-cp1250.doc'].forEach(name => {
    const source = WordImporter.parse(fixtureBuffer(name), { fileName: name });
    const texts = Importers.readSheet(source, source.sheetNames[0]).rows.map(r => r.Reference);
    assert.deepStrictEqual(texts, WORD_FIXTURE_PARAGRAPHS, name);
  });
  assert.ok(Importers.findForFile('smlouva.rtf'), '.rtf má importér');
});

runTest('WordImporter: .docx – entity se dekódují jen jednou, zalomení a tabulátory', () => {
  const xml = '<w:document><w:body>' +
    '<w:p><w:pPr><w:tabs><w:tab w:val="left" w:pos="720"/></w:tabs></w:pPr><w:r><w:t>A &amp;lt;b&amp;gt; &amp; C</w:t></w:r></w:p>' +
    '<w:p><w:r><w:t>Řádek 1</w:t><w:br w:type="textWrapping"/><w:t>Řádek 2</w:t><w:tab/><w:t>X</w:t></w:r></w:p>' +
    '</w:body></w:document>';
  const zip = XLSX.CFB.utils.cfb_new();
  XLSX.CFB.utils.cfb_add(zip, '/word/document.xml', Buffer.from(xml, 'utf8'));
  const bytes = XLSX.CFB.write(zip, { type: 'array', fileType: 'zip' });
  const source = WordImporter.parse(bytes, { fileName: 'a.docx' });
  assert.deepStrictEqual(Importers.readSheet(source, source.sheetNames[0]).rows.map(r => r.Reference), ['A &lt;b&gt; & C', 'Řádek 1\nŘádek 2\tX']);
});

runTest('Excel export: sloupec bez názvu v záhlaví se nevyexportuje jako "__EMPTY"', () => {
  const wb = ExcelExporter.buildWorkbook({ tables: [{ name: 'S', headers: ['Text', '__EMPTY', '__EMPTY_1'], rows: [{ Text: 'a', __EMPTY: 'b', __EMPTY_1: 'c' }] }] });
  const aoa = XLSX.utils.sheet_to_json(wb.Sheets.S, { header: 1, defval: '' });
  assert.deepStrictEqual(aoa, [['Text', '', ''], ['a', 'b', 'c']]);
});

runTest('Soubor projektu: písmena sloupců bez názvu se uloží a načtou', () => {
  const state = {
    fileName: 't.xlsx', headers: ['A', 'B'], hasHeaderRow: false, columnLetters: { A: 'A', B: 'B' },
    rows: [{ A: 'x', B: 'y' }], rowNumbers: [1],
    mappings: { refOldCol: 'A', refNewCol: 'B', targetCols: [], languages: {} }, translations: {}
  };
  const loaded = StorageManager.parseProjectJson(StorageManager.serializeProjectJson(state));
  assert.deepStrictEqual(loaded.columnLetters, { A: 'A', B: 'B' });
  const bad = JSON.parse(StorageManager.serializeProjectJson(state));
  bad.columnLetters = 'x';
  const loadedBad = StorageManager.parseProjectJson(bad);
  assert.deepStrictEqual(loadedBad.columnLetters, {});
  assert.ok(loadedBad.warnings.some(w => w.code === 'INVALID_SETTINGS'));
});

runTest('Import překladů: částečný sešit vyplněný překladatelem se spáruje podle řádků', () => {
  const rows = [
    { 'Ref 1': 'A old', 'Ref 2': 'A new text', 'CS 1': 'A cs' },
    { 'Ref 1': 'B', 'Ref 2': 'B', 'CS 1': 'B cs' },
    { 'Ref 1': 'C removed', 'Ref 2': '', 'CS 1': 'C cs' },
    { 'Ref 1': '', 'Ref 2': 'D added', 'CS 1': '' }
  ];
  const headers = ['Ref 1', 'Ref 2', 'CS 1'];
  const rowNumbers = [2, 3, 4, 6];
  const mappings = { refOldCol: 'Ref 1', refNewCol: 'Ref 2', targetCols: ['CS 1'], languages: { 'CS 1': { label: 'CS', newCol: '', newColName: 'CS new' } } };
  const translations = {};
  const exported = ComparatorEngine.prepareExportTables({
    rows, headers, rowNumbers, mappings, translations, exportType: 'language', langCol: 'CS 1',
    contextCols: ['Ref 1', 'Ref 2'], uiLang: 'cs'
  });
  // Round trip through a real workbook, as the translator would send it back
  const wb = ExcelExporter.buildWorkbook({ tables: exported.tables });
  const aoa = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, defval: '' });
  const newIdx = aoa[0].indexOf('CS new');
  assert.ok(newIdx > 0, 'sloupec nové verze v exportu');
  aoa[1][newIdx] = 'A cs NEW';
  aoa[3][newIdx] = 'ignorováno (zrušený odstavec)';
  aoa[4][newIdx] = 'D cs';
  const unknown = aoa[0].map(() => '');
  unknown[0] = 99;
  unknown[newIdx] = 'neznámý řádek';
  aoa.push(unknown);
  const wb2 = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb2, XLSX.utils.aoa_to_sheet(aoa), 'CS');
  const source = ExcelImporter.parse(XLSX.write(wb2, { type: 'array', bookType: 'xlsx' }), { fileName: 'cs.xlsx' });
  const table = Importers.readSheet(source, 'CS', { hasHeaderRow: true });

  const match = ComparatorEngine.matchImportedTranslations({ table, rows, rowNumbers, mappings, translations });
  assert.strictEqual(match.error, '');
  assert.strictEqual(match.languages.length, 1);
  assert.strictEqual(match.languages[0].column, 'CS new');
  assert.deepStrictEqual(match.languages[0].updates.map(u => [u.index, u.rowNumber, u.text]), [[0, 2, 'A cs NEW'], [3, 6, 'D cs']]);
  assert.strictEqual(match.languages[0].unchanged, 1, 'nezměněný odstavec B');
  assert.deepStrictEqual(match.unknownRows, [99]);

  const noRow = ComparatorEngine.matchImportedTranslations({ table: { headers: ['CS new'], rows: [{ 'CS new': 'x' }] }, rows, rowNumbers, mappings, translations });
  assert.strictEqual(noRow.error, 'NO_ROW_COLUMN');
  const noLang = ComparatorEngine.matchImportedTranslations({ table: { headers: ['Row', 'CS 1'], rows: [{ Row: 2, 'CS 1': 'x' }] }, rows, rowNumbers, mappings, translations });
  assert.strictEqual(noLang.error, 'NO_LANGUAGE_COLUMN', 'stará verze jazyka se nikdy nečte jako nová');
});

console.log(`\n========================================`);
console.log(`VÝSLEDEK TESTŮ: ${passCount} úspěšných, ${failCount} chyb`);
console.log(`========================================\n`);

if (failCount > 0) {
  process.exit(1);
}
