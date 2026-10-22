/**
 * Automated Test Suite for ComparatorEngine, StorageManager, PdfExporter and I18n
 * Run with: node test/comparator.test.js
 */

const assert = require('assert');
const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');
const ComparatorEngine = require('../js/core/comparator-engine');
const StorageManager = require('../js/core/storage');
const PdfExporter = require('../js/core/pdf-export');
const SampleGenerator = require('../js/core/sample-generator');
const I18n = require('../js/core/i18n');

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
});

runTest('buildBatchCopilotPrompt: obsahuje jen změněné a nové odstavce se značkami čísel řádků', () => {
  const g = genericTable();
  const stubs = ComparatorEngine.generateTargetStub(g.rows, 'Czech', g.mappings.refOldCol, g.mappings.refNewCol);
  const prompt = ComparatorEngine.buildBatchCopilotPrompt({
    promptLang: 'cs', targetCol: 'Czech', targetLangName: 'Čeština', refLangName: 'Reference new',
    rows: g.rows, rowNumbers: g.rowNumbers, mappings: g.mappings, translations: { Czech: stubs }
  });
  assert.ok(prompt.includes('ODSTAVEC: [#3] (Typ: UPRAVENO'), 'Upravený odstavec na Excel řádku 3');
  assert.ok(prompt.includes('ODSTAVEC: [#5] (Typ: NOVĚ PŘIDÁNO)'), 'Nový odstavec na Excel řádku 5');
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

runTest('buildExportWorkbook: kompletní sešit zapsaný do .xlsx a zpět, včetně listu se stavem kontroly', () => {
  const g = genericTable();
  const result = ComparatorEngine.buildExportWorkbook(XLSX, {
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

  const en = ComparatorEngine.buildExportWorkbook(XLSX, { rows: g.rows, headers: g.headers, mappings: g.mappings, exportType: 'comprehensive', uiLang: 'en' });
  assert.deepStrictEqual(en.workbook.SheetNames, ['Comparison', 'Translation_Status']);
});

runTest('buildExportWorkbook (částečný): číslo řádku, vybrané referenční sloupce, stará a nová verze jazyka a stav', () => {
  const g = genericTable();
  const translations = { Polish: ComparatorEngine.generateTargetStub(g.rows, 'Polish', g.mappings.refOldCol, g.mappings.refNewCol) };
  translations.Polish[1].targetNew = 'Limit to 50 EUR.';
  translations.Polish[1].isVerified = true;

  const result = ComparatorEngine.buildExportWorkbook(XLSX, {
    rows: g.rows, headers: g.headers, mappings: g.mappings, rowNumbers: g.rowNumbers, translations,
    exportType: 'language', langCol: 'Polish', contextCols: ['Reference old', 'Reference new'], uiLang: 'cs'
  });
  assert.deepStrictEqual(result.workbook.SheetNames, ['Polština']);
  const sheet = result.workbook.Sheets['Polština'];
  assert.deepStrictEqual(XLSX.utils.sheet_to_json(sheet, { header: 1 })[0], ['Řádek', 'Reference old', 'Reference new', 'Polish', 'Polish NEW', 'Stav']);
  const data = XLSX.utils.sheet_to_json(sheet, { defval: '' });
  assert.deepStrictEqual(data[1], { 'Řádek': 3, 'Reference old': 'Limit is 35 EUR.', 'Reference new': 'Limit is 50 EUR.', Polish: 'Limit to 35 EUR.', 'Polish NEW': 'Limit to 50 EUR.', Stav: 'Ověřeno' });
  assert.strictEqual(data[2].Stav, 'Chybí překlad');

  // Jazyk s existujícím sloupcem nové verze a jiné kontextové sloupce
  const cz = ComparatorEngine.prepareLanguageExportData({
    rows: g.rows, headers: g.headers, mappings: g.mappings, langCol: 'Czech', contextCols: ['Reference new', 'Polish', 'Czech']
  });
  assert.deepStrictEqual(cz.headers, ['Řádek', 'Reference new', 'Polish', 'Czech', 'Czech new', 'Stav'], 'Vlastní sloupce jazyka se jako kontext nezdvojí');
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
    exportLangCol: 'Polish', partialContextCols: ['Reference new']
  };

  const json = StorageManager.serializeProjectJson(state);
  const file = JSON.parse(json);
  assert.strictEqual(file.app, 'EthicCodeComparator');
  assert.strictEqual(file.translations.Czech[1].refOld, undefined, 'Kompaktní formát bez duplicitních textů');

  const restored = StorageManager.parseProjectJson(json);
  assert.deepStrictEqual(restored.mappings, g.mappings);
  assert.deepStrictEqual(restored.rowNumbers, g.rowNumbers);
  assert.deepStrictEqual(restored.pdfTemplate, state.pdfTemplate);
  assert.strictEqual(restored.exportLangCol, 'Polish');
  assert.deepStrictEqual(restored.partialContextCols, ['Reference new']);
  const rehydrated = ComparatorEngine.rehydrateStubs(restored.rows, restored.translations.Czech, 'Czech', g.mappings.refOldCol, g.mappings.refNewCol, 'Czech new');
  assert.deepStrictEqual(rehydrated, czech);
});

runTest('StorageManager: chybné soubory jsou odmítnuty, starší soubory doplněny o výchozí hodnoty', () => {
  const valid = { app: 'EthicCodeComparator', headers: ['A', 'B'], rows: [{ A: 'x', B: 'y' }], mappings: { refOldCol: 'A', refNewCol: 'B', targetCols: [] } };
  assert.throws(() => StorageManager.parseProjectJson('{ not json'), /platný JSON/);
  assert.throws(() => StorageManager.parseProjectJson(JSON.stringify({ ...valid, app: 'Jina' })), /nepatří/);
  assert.throws(() => StorageManager.parseProjectJson(JSON.stringify({ ...valid, rows: [] })), /rows/);
  assert.throws(() => StorageManager.parseProjectJson(JSON.stringify({ ...valid, mappings: { refOldCol: 'A', refNewCol: 'X' } })), /"X"/);
  assert.throws(() => StorageManager.parseProjectJson(JSON.stringify({
    ...valid, headers: ['A', 'B', 'C'], rows: [{ A: 1, B: 2, C: 3 }],
    mappings: { refOldCol: 'A', refNewCol: 'B', targetCols: ['C'], languages: { C: { newCol: 'CHYBÍ' } } }
  })), /CHYBÍ/, 'Neexistující sloupec nové verze');

  const legacy = StorageManager.parseProjectJson('﻿' + JSON.stringify({ rows: [{ A: 'x', B: 'y', C: 'z' }], mappings: { refOldCol: 'A', refNewCol: 'B', targetCols: ['C'] } }));
  assert.deepStrictEqual(legacy.headers, ['A', 'B', 'C']);
  assert.deepStrictEqual(legacy.mappings.languages, { C: { label: '', newCol: '', newColName: '' } });
  assert.deepStrictEqual(legacy.rowNumbers, [2]);
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
runTest('generateManagerialChangeSummary: reference – statistika, Markdown a Word s čísly řádků', () => {
  const s = sample();
  const summary = ComparatorEngine.generateManagerialChangeSummary({ rows: s.rows, mappings: s.mappings, documentName: 'Dokument', uiLang: 'cs' });
  assert.strictEqual(summary.isReference, true);
  assert.deepStrictEqual(summary.stats, { totalRows: 15, unchangedCount: 5, modifiedCount: 6, addedCount: 2, removedCount: 2, pctChanged: 67 });
  assert.ok(summary.markdown.includes('# Shrnutí změn – DE 2026 (Referenz) (stará → nová verze)'));
  assert.ok(summary.markdown.includes('| **Celkem odstavců** | 15 | 100% |'));
  assert.ok(summary.markdown.includes('### Řádek 5'), 'Odstavce jsou označeny číslem řádku');
  const generatedLines = summary.markdown.split('\n').filter(l => !l.startsWith('>') && !l.startsWith('- **'));
  assert.ok(!/etick|kodex|němč|Art\. 0/i.test(generatedLines.join('\n')), 'Žádné předpoklady o druhu dokumentu nebo nadpisech');
  assert.ok(summary.htmlDoc.includes('xmlns:w="urn:schemas-microsoft-com:office:word"'));
});

runTest('generateManagerialChangeSummary: vybraný jazyk – jeho stará a nová verze a stav kontroly', () => {
  const g = genericTable();
  const translations = { Polish: ComparatorEngine.generateTargetStub(g.rows, 'Polish', g.mappings.refOldCol, g.mappings.refNewCol) };
  translations.Polish[1].targetNew = 'Limit to 50 EUR.';
  translations.Polish[1].isVerified = true;

  const summary = ComparatorEngine.generateManagerialChangeSummary({
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

  const printable = PdfExporter.generateManagerialReportHtml({ rows: g.rows, mappings: g.mappings, translations, langCol: 'Polish', uiLang: 'en' });
  assert.ok(printable.includes('window.print()') && printable.includes('Change summary – Polština'));
});

runTest('generateManagerialChangeSummary: víceřádkový text zůstane celý v citaci a se zalomením ve Wordu', () => {
  const rows = [{ A: 'Zeile eins\nZeile zwei', B: 'Zeile eins\nZeile drei' }];
  const s = ComparatorEngine.generateManagerialChangeSummary({ rows, mappings: { refOldCol: 'A', refNewCol: 'B' } });
  assert.ok(s.markdown.includes('> Zeile eins\n> Zeile zwei'));
  assert.ok(s.markdown.includes('> Zeile eins\n> Zeile drei'));
  assert.ok(s.htmlDoc.includes('Zeile eins<br>Zeile zwei'));
});

// 9. Dokument pro tisk
runTest('PdfExporter: dokument nové verze jazyka – čisté odstavce bez nadpisů, koncept při neúplném překladu', () => {
  const g = genericTable();
  const html = PdfExporter.generateDocumentHtml({
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

  const ref = PdfExporter.generateDocumentHtml({ rows: g.rows, mappings: g.mappings, targetCol: 'Reference new', uiLang: 'cs' });
  assert.ok(ref.includes('Brand new paragraph.') && !ref.includes('KONCEPT'));
});

runTest('PdfExporter: vlastní hlavička (organizace, podtitul s {LANG}, verze, podpisy) a escapování', () => {
  const g = genericTable();
  const html = PdfExporter.generateDocumentHtml({
    rows: g.rows, mappings: g.mappings, targetCol: 'Polish', uiLang: 'cs',
    template: {
      companyName: 'ACME a.s.', docTitle: 'Provozní řád', docSubtitle: 'Znění – {LANG}', versionLabel: '<b>v3</b>',
      effectiveDate: 'Platné od 1. 1.', accentColor: '#0d9488', introPreamble: 'Úvod.', showSignatures: true, signDept1: 'Ředitel', signDept2: 'Rada'
    }
  });
  ['ACME a.s.', 'Provozní řád', 'Znění – Polština', 'Platné od 1. 1.', '#0d9488', 'Úvod.', 'Ředitel', '<div class="sign-block">']
    .forEach(part => assert.ok(html.includes(part), 'Chybí: ' + part));
  assert.ok(html.includes('&lt;b&gt;v3&lt;/b&gt;') && !html.includes('<b>v3</b>'));
  assert.ok(!PdfExporter.generateDocumentHtml({ rows: g.rows, mappings: g.mappings, targetCol: 'Polish', template: { accentColor: 'red;}body{x' } }).includes('red;}'),
    'Neplatná barva se nepoužije');
});

// 10. Integrace se vzorovým souborem
runTest('Integrace: ukázkový .xlsx – stuby a kompletní export', () => {
  const wb = XLSX.readFile(path.join(__dirname, '../sample/vzorec_eticky_kodex_2024_2026.xlsx'));
  const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]);
  assert.strictEqual(rows.length, 15);
  const s = sample();
  const stubs = ComparatorEngine.generateTargetStub(rows, 'CS 2024', s.mappings.refOldCol, s.mappings.refNewCol);
  stubs[3].targetNew = 'Dary do 50 eur.';
  const res = ComparatorEngine.prepareExportData({ rows, headers: Object.keys(rows[0]), mappings: s.mappings, translations: { 'CS 2024': stubs } });
  assert.strictEqual(res.data.length, 15);
  assert.strictEqual(res.data[3]['CS 2026'], 'Dary do 50 eur.', 'Výchozí název nového sloupce odvozen z čísla verze reference');
});

// 11. Lokalizace a obecnost textů
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

console.log(`\n========================================`);
console.log(`VÝSLEDEK TESTŮ: ${passCount} úspěšných, ${failCount} chyb`);
console.log(`========================================\n`);

if (failCount > 0) {
  process.exit(1);
}
