/**
 * Automated Test Suite for ComparatorEngine
 * Run with: node test/comparator.test.js
 */

const assert = require('assert');
const ComparatorEngine = require('../js/core/comparator-engine');

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

console.log('=== Spouštím testy pro ComparatorEngine ===\n');

// 1. Testy klasifikace řádků (classifyRow)
runTest('classifyRow: EMPTY pro prázdné buňky nebo pouze bílé znaky', () => {
  assert.strictEqual(ComparatorEngine.classifyRow('', ''), 'EMPTY');
  assert.strictEqual(ComparatorEngine.classifyRow('   ', '\n\t '), 'EMPTY');
  assert.strictEqual(ComparatorEngine.classifyRow(null, undefined), 'EMPTY');
});

runTest('classifyRow: ADDED pro nový odstavec (2024 prázdný, 2026 text)', () => {
  assert.strictEqual(ComparatorEngine.classifyRow('', 'Neuer Absatz im Jahr 2026.'), 'ADDED');
  assert.strictEqual(ComparatorEngine.classifyRow('   ', 'Neuer Absatz.'), 'ADDED');
  assert.strictEqual(ComparatorEngine.classifyRow(null, 'Paragraph 15a'), 'ADDED');
});

runTest('classifyRow: REMOVED pro odstraněný odstavec (2024 text, 2026 prázdný)', () => {
  assert.strictEqual(ComparatorEngine.classifyRow('Alter Absatz aus 2024.', ''), 'REMOVED');
  assert.strictEqual(ComparatorEngine.classifyRow('Alter Absatz.', '   '), 'REMOVED');
  assert.strictEqual(ComparatorEngine.classifyRow('Paragraph 12', null), 'REMOVED');
});

runTest('classifyRow: UNCHANGED pro shodný text včetně shody po normalizaci konců řádků', () => {
  const text = 'Wir respektieren die Menschenrechte aller Mitarbeiter.';
  assert.strictEqual(ComparatorEngine.classifyRow(text, text), 'UNCHANGED');
  assert.strictEqual(ComparatorEngine.classifyRow('Text mit Umlauten: äöüß', 'Text mit Umlauten: äöüß'), 'UNCHANGED');
  // Normalizace konců řádků
  assert.strictEqual(ComparatorEngine.classifyRow('Řádek 1\r\nŘádek 2', 'Řádek 1\nŘádek 2'), 'UNCHANGED');
});

runTest('classifyRow: MODIFIED pro upravený text', () => {
  assert.strictEqual(
    ComparatorEngine.classifyRow('Geschenke bis 35 Euro.', 'Geschenke bis 50 Euro.'),
    'MODIFIED'
  );
  assert.strictEqual(
    ComparatorEngine.classifyRow('Frist beträgt 14 Tage.', 'Frist beträgt 7 Tage.'),
    'MODIFIED'
  );
});

// 2. Testy textového diffu (computeDiff)
runTest('computeDiff: detekce beze změny', () => {
  const res = ComparatorEngine.computeDiff('Stejný text.', 'Stejný text.');
  assert.strictEqual(res.stats.hasChanges, false);
  assert.strictEqual(res.stats.addedTokens, 0);
  assert.strictEqual(res.stats.removedTokens, 0);
  assert.strictEqual(res.stats.similarity, 1);
});

runTest('computeDiff: přesná detekce přidaných a odstraněných slov v němčině', () => {
  const oldText = 'Geschenke bis zu einem Wert von 35 Euro sind zulässig.';
  const newText = 'Geschenke bis zu einem Wert von 50 Euro sind genehmigungspflichtig.';
  const res = ComparatorEngine.computeDiff(oldText, newText);

  assert.strictEqual(res.stats.hasChanges, true);
  assert.ok(res.summary.removedTerms.includes('35'));
  assert.ok(res.summary.addedTerms.includes('50'));
  assert.ok(res.summary.removedTerms.includes('zulässig'));
  assert.ok(res.summary.addedTerms.includes('genehmigungspflichtig'));
});

runTest('computeDiff: bezpečné ošetření diakritiky v češtině a speciálních znaků', () => {
  const oldText = 'Lhůta pro nahlášení je 14 kalendářních dnů u ombudsmana.';
  const newText = 'Lhůta pro nahlášení je 7 pracovních dnů prostřednictvím portálu.';
  const res = ComparatorEngine.computeDiff(oldText, newText);

  assert.strictEqual(res.stats.hasChanges, true);
  assert.ok(res.summary.removedTerms.includes('14'));
  assert.ok(res.summary.addedTerms.includes('7'));
  assert.ok(res.summary.removedTerms.includes('kalendářních'));
  assert.ok(res.summary.addedTerms.includes('pracovních'));
});

// 3. Testy renderování HTML diffu (renderDiffHtml)
runTest('renderDiffHtml: sanitizace XSS a správné značky <ins> a <del>', () => {
  const oldText = 'Pozor na <script>alert(1)</script>';
  const newText = 'Pozor na <b>bezpečný</b> kód';
  const diffRes = ComparatorEngine.computeDiff(oldText, newText);
  const inlineHtml = ComparatorEngine.renderDiffHtml(diffRes, 'inline');

  assert.ok(!inlineHtml.includes('<script>'), 'Nesmí obsahovat neescapovaný script');
  assert.ok(inlineHtml.includes('&lt;') && inlineHtml.includes('&gt;'));
  assert.ok(inlineHtml.includes('<del class="diff-del"'));
  assert.ok(inlineHtml.includes('<ins class="diff-ins"'));
});

runTest('renderDiffHtml: side-by-side režim vrací levý a pravý panel', () => {
  const oldText = 'Stara verze';
  const newText = 'Nova verze';
  const diffRes = ComparatorEngine.computeDiff(oldText, newText);
  const sbs = ComparatorEngine.renderDiffHtml(diffRes, 'side-by-side');

  assert.ok(sbs.leftHtml.includes('<del class="diff-del"'));
  assert.ok(!sbs.leftHtml.includes('<ins class="diff-ins"'));
  assert.ok(sbs.rightHtml.includes('<ins class="diff-ins"'));
  assert.ok(!sbs.rightHtml.includes('<del class="diff-del"'));
});

// 4. Testy generování stubu pro cílový jazyk (generateTargetStub)
runTest('generateTargetStub: správná pravidla pro UNCHANGED, ADDED, REMOVED, MODIFIED', () => {
  const rows = [
    { id: 1, 'DE 2024': 'Unverändert Text.', 'DE 2026': 'Unverändert Text.', 'CS 2024': 'Nezměněný text.' },
    { id: 2, 'DE 2024': 'Alter Text.', 'DE 2026': 'Neuer Text.', 'CS 2024': 'Starý text.' },
    { id: 3, 'DE 2024': '', 'DE 2026': 'Ganz neu.', 'CS 2024': '' },
    { id: 4, 'DE 2024': 'Gelöscht.', 'DE 2026': '', 'CS 2024': 'Bude smazáno.' }
  ];

  const stubs = ComparatorEngine.generateTargetStub(rows, 'CS 2024', 'DE 2024', 'DE 2026');

  // Řádek 1: UNCHANGED -> automatické zkopírování 2024
  assert.strictEqual(stubs[0].classification, 'UNCHANGED');
  assert.strictEqual(stubs[0].targetNew, 'Nezměněný text.');
  assert.strictEqual(stubs[0].status, 'auto_copied');
  assert.strictEqual(stubs[0].isVerified, true);

  // Řádek 2: MODIFIED -> čeká na revizi
  assert.strictEqual(stubs[1].classification, 'MODIFIED');
  assert.strictEqual(stubs[1].targetNew, '');
  assert.strictEqual(stubs[1].status, 'pending_review');
  assert.strictEqual(stubs[1].isVerified, false);

  // Řádek 3: ADDED -> nový překlad
  assert.strictEqual(stubs[2].classification, 'ADDED');
  assert.strictEqual(stubs[2].status, 'pending_translation');
  assert.strictEqual(stubs[2].isVerified, false);

  // Řádek 4: REMOVED -> prázdný řádek
  assert.strictEqual(stubs[3].classification, 'REMOVED');
  assert.strictEqual(stubs[3].targetNew, '');
  assert.strictEqual(stubs[3].status, 'removed');
  assert.strictEqual(stubs[3].isVerified, true);
});

// 5. Testy generování promptu pro Copilot (buildCopilotPrompt)
runTest('buildCopilotPrompt: MODIFIED prompt obsahuje původní i nový německý text a stávající překlad', () => {
  const diffResult = ComparatorEngine.computeDiff(
    'Frist ist 14 Tage.',
    'Frist ist 7 Tage.'
  );

  const prompt = ComparatorEngine.buildCopilotPrompt({
    targetLangName: 'Čeština (Czech)',
    rowId: 'Článek 3 - Hlášení',
    classification: 'MODIFIED',
    refOld: 'Frist ist 14 Tage.',
    refNew: 'Frist ist 7 Tage.',
    targetOld: 'Lhůta je 14 dnů.',
    diffResult: diffResult
  });

  assert.ok(prompt.includes('Čeština (Czech)'));
  assert.ok(prompt.includes('Frist ist 14 Tage.'));
  assert.ok(prompt.includes('Frist ist 7 Tage.'));
  assert.ok(prompt.includes('Lhůta je 14 dnů.'));
  assert.ok(prompt.includes('minimální nutné změny'));
  assert.ok(prompt.includes('POUZE'));
});

runTest('buildCopilotPrompt: ADDED prompt správně žádá o nový překlad', () => {
  const prompt = ComparatorEngine.buildCopilotPrompt({
    targetLangName: 'Francouzština (French)',
    rowId: 'Článek 10 - AI Etika',
    classification: 'ADDED',
    refOld: '',
    refNew: 'Einsatz von Künstlicher Intelligenz muss transparent erfolgen.',
    targetOld: ''
  });

  assert.ok(prompt.includes('Francouzština (French)'));
  assert.ok(prompt.includes('Einsatz von Künstlicher Intelligenz muss transparent erfolgen.'));
  assert.ok(prompt.includes('ÚKOL: Přeložte nový odstavec'));
});

runTest('buildCopilotPrompt: anglická verze (promptLang: en) generuje anglické instrukce', () => {
  const prompt = ComparatorEngine.buildCopilotPrompt({
    promptLang: 'en',
    targetLangName: 'Spanish (ES)',
    rowId: 'Article 4 - Gifts',
    classification: 'MODIFIED',
    refOld: 'Geschenke bis 35 Euro.',
    refNew: 'Geschenke bis 50 Euro.',
    targetOld: 'Regalos hasta 35 euros.',
    diffResult: ComparatorEngine.computeDiff('Geschenke bis 35 Euro.', 'Geschenke bis 50 Euro.')
  });

  assert.ok(prompt.includes('expert legal and corporate translator'));
  assert.ok(prompt.includes('Spanish (ES)'));
  assert.ok(prompt.includes('CRITICAL INSTRUCTIONS'));
  assert.ok(prompt.includes('minimal necessary alterations'));
});

// 6. Testy přípravy exportu (prepareExportData)
runTest('prepareExportData: comprehensive export přidává nové sloupce vedle starých', () => {
  const headers = ['ID', 'DE 2024', 'DE 2026', 'CS 2024', 'EN 2024'];
  const rows = [
    { ID: '1', 'DE 2024': 'Alt', 'DE 2026': 'Neu', 'CS 2024': 'Staré CS', 'EN 2024': 'Old EN' }
  ];
  const mappings = {
    refOldCol: 'DE 2024',
    refNewCol: 'DE 2026',
    targetCols: ['CS 2024', 'EN 2024']
  };
  const translations = {
    'CS 2024': [{ targetNew: 'Nové CS', status: 'verified' }],
    'EN 2024': [{ targetNew: 'New EN', status: 'verified' }]
  };

  const exportObj = ComparatorEngine.prepareExportData({
    headers,
    rows,
    mappings,
    translations,
    exportType: 'comprehensive'
  });

  const row0 = exportObj.data[0];
  assert.strictEqual(row0['CS 2024'], 'Staré CS');
  assert.strictEqual(row0['CS 2026'], 'Nové CS');
  assert.strictEqual(row0['EN 2024'], 'Old EN');
  assert.strictEqual(row0['EN 2026'], 'New EN');
});

runTest('prepareExportData: clean_2026 export filtruje smazané řádky a generuje čistou tabulku', () => {
  const headers = ['ID', 'DE 2024', 'DE 2026', 'CS 2024'];
  const rows = [
    { ID: '1', 'DE 2024': 'Platný', 'DE 2026': 'Platný nově', 'CS 2024': 'Platný CS' },
    { ID: '2', 'DE 2024': 'Smazaný odstavec', 'DE 2026': '', 'CS 2024': 'Bude smazán' }
  ];
  const mappings = {
    refOldCol: 'DE 2024',
    refNewCol: 'DE 2026',
    targetCols: ['CS 2024']
  };
  const translations = {
    'CS 2024': [
      { targetNew: 'Upravený CS', status: 'verified' },
      { targetNew: '', status: 'removed' }
    ]
  };

  const exportObj = ComparatorEngine.prepareExportData({
    headers,
    rows,
    mappings,
    translations,
    exportType: 'clean_2026'
  });

  assert.strictEqual(exportObj.data.length, 1, 'Měl by zůstat pouze 1 řádek (smazaný je odfiltrován)');
  assert.strictEqual(exportObj.data[0]['ID'], '1');
  assert.strictEqual(exportObj.data[0]['CS 2026'], 'Upravený CS');
});

// 7. End-to-end Integrační test se skutečným souborem Excel
runTest('Integrační test: Načtení a zpracování sample/vzorec_eticky_kodex_2024_2026.xlsx', () => {
  const XLSX = require('xlsx');
  const path = require('path');
  const filePath = path.join(__dirname, '../sample/vzorec_eticky_kodex_2024_2026.xlsx');
  
  const workbook = XLSX.readFile(filePath);
  const firstSheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[firstSheetName];
  const rows = XLSX.utils.sheet_to_json(sheet);

  assert.strictEqual(rows.length, 15, 'Má obsahovat 15 článků');

  const refOldCol = 'DE 2024 (Referenz)';
  const refNewCol = 'DE 2026 (Referenz)';
  const targetCol = 'CS 2024';

  const classifications = rows.map(r => ComparatorEngine.classifyRow(r[refOldCol], r[refNewCol]));
  
  const counts = classifications.reduce((acc, c) => {
    acc[c] = (acc[c] || 0) + 1;
    return acc;
  }, {});

  // Ověření distribuce
  assert.ok(counts.UNCHANGED >= 5, 'Má mít minimálně 5 nezměněných odstavců');
  assert.ok(counts.MODIFIED >= 4, 'Má mít minimálně 4 upravené odstavce');
  assert.ok(counts.ADDED >= 2, 'Má mít 2 nové odstavce v 2026');
  assert.ok(counts.REMOVED >= 2, 'Má mít 2 odstraněné odstavce v 2026');

  // Stub generace
  const stubs = ComparatorEngine.generateTargetStub(rows, targetCol, refOldCol, refNewCol);
  assert.strictEqual(stubs.length, 15);

  // Zápis nového překladu pro Art. 04 (Dary & Bewirtung)
  const art4Index = rows.findIndex(r => r['Artikel-Nr.'] === 'Art. 04');
  stubs[art4Index].targetNew = 'Dary a pozvání lze přijmout pouze tehdy, pokud jejich hodnota nepřesahuje 50 eur a byly předem schváleny přes compliance portál.';
  stubs[art4Index].isVerified = true;

  // Export
  const headers = Object.keys(rows[0]);
  const exportRes = ComparatorEngine.prepareExportData({
    rows,
    headers,
    mappings: {
      refOldCol,
      refNewCol,
      targetCols: [targetCol]
    },
    translations: {
      [targetCol]: stubs
    },
    exportType: 'comprehensive'
  });

  assert.strictEqual(exportRes.data.length, 15);
  assert.strictEqual(exportRes.data[art4Index]['CS 2026'], stubs[art4Index].targetNew);
});

// 8. Test ukládání a načítání projektového JSON souboru přes skutečný serializér aplikace
runTest('StorageManager: serializeProjectJson → parseProjectJson → rehydrateStubs zachová veškerou práci (roundtrip)', () => {
  const StorageManager = require('../js/core/storage');
  const SampleGenerator = require('../js/core/sample-generator');

  const rows = SampleGenerator.getSampleData();
  const headers = SampleGenerator.getSampleHeaders();
  const mappings = SampleGenerator.getSampleDefaultMappings();
  const csStubs = ComparatorEngine.generateTargetStub(rows, 'CS 2024', mappings.refOldCol, mappings.refNewCol);
  const enStubs = ComparatorEngine.generateTargetStub(rows, 'EN 2024', mappings.refOldCol, mappings.refNewCol);

  csStubs[3].targetNew = 'Upravený text článku 4 v češtině.';
  csStubs[3].isVerified = true;
  csStubs[3].status = 'verified';
  csStubs[5].targetNew = 'Rozpracovaný koncept článku 6.';
  csStubs[5].status = 'draft';
  enStubs[0].isVerified = false; // uživatel zrušil automatické schválení

  const pdfTemplate = { companyName: 'ACME a.s.', docTitle: 'KODEX 2026', accentColor: '#0d9488', versionLabel: 'Edice 3', showSignatures: false };
  const state = {
    fileName: 'test_kodex.xlsx',
    activeSheet: 'Eticky_Kodex',
    headers, rows, mappings,
    translations: { 'CS 2024': csStubs, 'EN 2024': enStubs },
    customPromptTemplate: 'Vlastní prompt šablona {TARGET_LANG}',
    pdfTemplate,
    activeTargetCol: 'EN 2024'
  };

  const json = StorageManager.serializeProjectJson(state);
  const fileData = JSON.parse(json);
  assert.strictEqual(fileData.app, 'EthicCodeComparator');
  // Kompaktní formát: referenční a původní texty se neukládají duplicitně pro každý jazyk
  assert.strictEqual(fileData.translations['CS 2024'][3].refOld, undefined, 'Stub nesmí duplikovat refOld');
  assert.strictEqual(fileData.translations['CS 2024'][3].targetOld, undefined, 'Stub nesmí duplikovat targetOld');

  const restored = StorageManager.parseProjectJson(json);
  assert.strictEqual(restored.fileName, 'test_kodex.xlsx');
  assert.deepStrictEqual(restored.rows, rows);
  assert.deepStrictEqual(restored.headers, headers);
  assert.deepStrictEqual(restored.mappings, mappings);
  assert.deepStrictEqual(restored.pdfTemplate, pdfTemplate, 'Šablona PDF musí být součástí projektu');
  assert.strictEqual(restored.customPromptTemplate, 'Vlastní prompt šablona {TARGET_LANG}');
  assert.strictEqual(restored.activeTargetCol, 'EN 2024');

  // Rehydratace (stejně jako v aplikaci) obnoví plné stuby včetně textů z řádků
  const cs = ComparatorEngine.rehydrateStubs(restored.rows, restored.translations['CS 2024'], 'CS 2024', mappings.refOldCol, mappings.refNewCol);
  const en = ComparatorEngine.rehydrateStubs(restored.rows, restored.translations['EN 2024'], 'EN 2024', mappings.refOldCol, mappings.refNewCol);
  assert.deepStrictEqual(cs, csStubs, 'Rehydratované stuby CS musí být identické s původními');
  assert.deepStrictEqual(en, enStubs, 'Rehydratované stuby EN musí být identické s původními');
});

runTest('StorageManager: parseProjectJson odmítne poškozené nebo cizí soubory srozumitelnou chybou', () => {
  const StorageManager = require('../js/core/storage');
  const valid = { app: 'EthicCodeComparator', headers: ['A', 'B'], rows: [{ A: 'x', B: 'y' }], mappings: { refOldCol: 'A', refNewCol: 'B', targetCols: [] } };

  assert.throws(() => StorageManager.parseProjectJson('{ not json'), /platný JSON/);
  assert.throws(() => StorageManager.parseProjectJson(JSON.stringify({ ...valid, app: 'JinaAplikace' })), /nepatří/);
  assert.throws(() => StorageManager.parseProjectJson(JSON.stringify({ ...valid, rows: [] })), /rows/);
  assert.throws(() => StorageManager.parseProjectJson(JSON.stringify({ ...valid, mappings: null })), /mappings/);
  assert.throws(() => StorageManager.parseProjectJson(JSON.stringify({ ...valid, mappings: { refOldCol: 'A', refNewCol: 'NEEXISTUJE', targetCols: [] } })), /NEEXISTUJE/);

  // Starší soubor bez headers/targetCols a s BOM se načte s doplněnými výchozími hodnotami
  const legacy = StorageManager.parseProjectJson('﻿' + JSON.stringify({ rows: valid.rows, mappings: { refOldCol: 'A', refNewCol: 'B' } }));
  assert.deepStrictEqual(legacy.headers, ['A', 'B']);
  assert.deepStrictEqual(legacy.mappings.targetCols, []);
  assert.deepStrictEqual(legacy.translations, {});
});

runTest('rehydrateStubs: změna klasifikace řádku (jiné referenční sloupce) zruší jeho ověření, text zůstane', () => {
  const rows = [{ ID: '1', 'DE A': 'Alt', 'DE B': 'Neu', 'DE C': 'Alt', CS: 'Starý' }];
  const saved = [{ classification: 'MODIFIED', targetNew: 'Nový', status: 'verified', isVerified: true }];

  const same = ComparatorEngine.rehydrateStubs(rows, saved, 'CS', 'DE A', 'DE B');
  assert.strictEqual(same[0].isVerified, true);
  assert.strictEqual(same[0].targetOld, 'Starý', 'targetOld se doplní z řádků');

  const changed = ComparatorEngine.rehydrateStubs(rows, saved, 'CS', 'DE A', 'DE C'); // nyní UNCHANGED
  assert.strictEqual(changed[0].classification, 'UNCHANGED');
  assert.strictEqual(changed[0].isVerified, false, 'Ověření neplatí pro jinou klasifikaci');
  assert.strictEqual(changed[0].targetNew, 'Nový', 'Uživatelův text se neztratí');
});

// 9. Test exportu a opětovného importu binárního Excel souboru (.xlsx) přes buildExportWorkbook
runTest('XLSX: buildExportWorkbook zapíše a zpětně načte kompletní sešit vč. listu se stavem překladu', () => {
  const XLSX = require('xlsx');
  const SampleGenerator = require('../js/core/sample-generator');

  const rows = SampleGenerator.getSampleData();
  const headers = SampleGenerator.getSampleHeaders();
  const mappings = SampleGenerator.getSampleDefaultMappings();
  const stubs = ComparatorEngine.generateTargetStub(rows, 'CS 2024', mappings.refOldCol, mappings.refNewCol);

  stubs[5].targetNew = 'Korupce v roce 2026 je přísně zakázána.';
  stubs[5].isVerified = true;

  const result = ComparatorEngine.buildExportWorkbook(XLSX, {
    rows, headers, mappings,
    translations: { 'CS 2024': stubs }, // ostatní jazyky nebyly ve studiu otevřeny
    exportType: 'comprehensive',
    uiLang: 'cs'
  });

  const excelBuffer = XLSX.write(result.workbook, { type: 'buffer', bookType: 'xlsx' });
  const readWb = XLSX.read(excelBuffer, { type: 'buffer' });
  assert.deepStrictEqual(readWb.SheetNames, ['Srovnani_2024_2026_Komplet', 'Stav_prekladu']);

  // Pořadí sloupců: každý nový sloupec 2026 hned za svým sloupcem 2024
  const headerRow = XLSX.utils.sheet_to_json(readWb.Sheets.Srovnani_2024_2026_Komplet, { header: 1 })[0];
  assert.deepStrictEqual(headerRow.slice(0, 6), ['Artikel-Nr.', 'Bereich', 'DE 2024 (Referenz)', 'DE 2026 (Referenz)', 'EN 2024', 'EN 2026']);
  assert.strictEqual(headerRow.indexOf('CS 2026'), headerRow.indexOf('CS 2024') + 1);

  const reloadedRows = XLSX.utils.sheet_to_json(readWb.Sheets.Srovnani_2024_2026_Komplet, { defval: '' });
  assert.strictEqual(reloadedRows.length, 15);
  assert.strictEqual(reloadedRows[5]['CS 2026'], 'Korupce v roce 2026 je přísně zakázána.');
  // Neotevřený jazyk: nezměněný článek převzat z 2024, upravený zůstává prázdný (žádný zastaralý text)
  assert.strictEqual(reloadedRows[0]['EN 2026'], rows[0]['EN 2024']);
  assert.strictEqual(reloadedRows[3]['EN 2026'], '');

  const status = XLSX.utils.sheet_to_json(readWb.Sheets.Stav_prekladu, { defval: '' });
  assert.strictEqual(status.length, 15);
  assert.strictEqual(status[5]['CS 2026'], 'Ověřeno');
  assert.strictEqual(status[3]['CS 2026'], 'Chybí překlad');
  assert.strictEqual(status[9]['Změna reference'], 'Odstraněno');
  assert.ok(result.unverifiedCount > 0, 'Počet neověřených musí být nahlášen');
});

runTest('prepareExportData: nový sloupec 2026 nepřepíše existující sloupec stejného jména', () => {
  const headers = ['ID', 'DE 2024', 'DE 2026', 'CS 2024', 'CS 2026'];
  const rows = [{ ID: '1', 'DE 2024': 'a', 'DE 2026': 'b', 'CS 2024': 'staré', 'CS 2026': 'PŮVODNÍ SLOUPEC' }];
  const mappings = { refOldCol: 'DE 2024', refNewCol: 'DE 2026', targetCols: ['CS 2024'] };
  const translations = { 'CS 2024': [{ classification: 'MODIFIED', targetNew: 'NOVÝ PŘEKLAD', isVerified: true }] };

  const comp = ComparatorEngine.prepareExportData({ headers, rows, mappings, translations });
  assert.deepStrictEqual(comp.headers, ['ID', 'DE 2024', 'DE 2026', 'CS 2024', 'CS 2026 (2)', 'CS 2026']);
  assert.strictEqual(comp.data[0]['CS 2026'], 'PŮVODNÍ SLOUPEC');
  assert.strictEqual(comp.data[0]['CS 2026 (2)'], 'NOVÝ PŘEKLAD');

  const clean = ComparatorEngine.prepareExportData({ headers, rows, mappings, translations, exportType: 'clean_2026' });
  assert.strictEqual(new Set(clean.headers).size, clean.headers.length, 'Hlavičky musí být unikátní');
  assert.ok(clean.headers.some(h => clean.data[0][h] === 'NOVÝ PŘEKLAD'), 'Překlad nesmí zmizet');
});

runTest('prepareExportData: názvy sloupců 2026 a číselné hlavičky zachovají pořadí', () => {
  assert.strictEqual(ComparatorEngine.format2026ColName('DE 2026 (Referenz)'), 'DE 2026 (Referenz)');
  assert.strictEqual(ComparatorEngine.format2026ColName('CS 2024'), 'CS 2026');
  assert.strictEqual(ComparatorEngine.format2026ColName('Czech'), 'Czech (2026)');

  const headers = ['ID', 'DE alt', 'DE neu', '2024'];
  const rows = [{ ID: 1, 'DE alt': 'a', 'DE neu': 'a', '2024': 'x' }];
  const res = ComparatorEngine.prepareExportData({ headers, rows, mappings: { refOldCol: 'DE alt', refNewCol: 'DE neu', targetCols: ['2024'] } });
  assert.deepStrictEqual(res.headers, ['ID', 'DE alt', 'DE neu', '2024', '2026']);
  assert.strictEqual(res.data[0]['2026'], 'x', 'Nezměněný text se převezme');
  assert.strictEqual(res.data[0].ID, 1, 'Čísla zůstávají čísly');
});

// 10. Test generování hezky naformátovaného PDF / Printable HTML pro jazykovou mutaci
runTest('PdfExporter: Ověření generování formátovaného tiskového PDF dokumentu pro cílový jazyk', () => {
  const PdfExporter = require('../js/core/pdf-export');
  const SampleGenerator = require('../js/core/sample-generator');

  const rows = SampleGenerator.getSampleData();
  const headers = SampleGenerator.getSampleHeaders();
  const mappings = SampleGenerator.getSampleDefaultMappings();
  const stubs = ComparatorEngine.generateTargetStub(rows, 'CS 2024', mappings.refOldCol, mappings.refNewCol);
  stubs[3].targetNew = 'Oficiální český překlad článku 4 pro rok 2026 s limitem 50 eur.';

  const docHtml = PdfExporter.generateDocumentHtml({
    rows: rows,
    headers: headers,
    mappings: mappings,
    translations: { 'CS 2024': stubs },
    targetCol: 'CS 2024',
    companyName: 'Acme Global Ethics Group',
    uiLang: 'cs'
  });

  assert.ok(typeof docHtml === 'string' && docHtml.length > 1000, 'HTML musí být vygenerován');
  assert.ok(docHtml.includes('Acme Global Ethics Group'), 'Musí obsahovat název společnosti');
  assert.ok(docHtml.includes('ETICKÝ KODEX A PRAVIDLA CHOVÁNÍ 2026'), 'Musí obsahovat oficiální český titul');
  assert.ok(docHtml.includes('CS 2026'), 'Musí obsahovat označení jazyka');
  assert.ok(docHtml.includes('@media print') && docHtml.includes('@page'), 'Musí obsahovat tisková pravidla pro A4');
  assert.ok(docHtml.includes('window.print()'), 'Musí obsahovat tiskové tlačítko');
  assert.ok(docHtml.includes('Oficiální český překlad článku 4 pro rok 2026 s limitem 50 eur.'), 'Musí obsahovat schválený text');
  // Musí vynechat smazané články 10 a 11
  assert.ok(!docHtml.includes('Fax-Meldungen'), 'Smazaný článek faxových zpráv nesmí být v 2026 dokumentu');
  assert.ok(!docHtml.includes('Hygieneregeln & Pandemie'), 'Smazaný článek pandemických pravidel nesmí být v 2026 dokumentu');
});

// 11. Test generování dávkového promptu pro Copilot (Batch Prompt)
runTest('buildBatchCopilotPrompt: Dávkový prompt pro všechny upravené a přidané články', () => {
  const SampleGenerator = require('../js/core/sample-generator');
  const rows = SampleGenerator.getSampleData();
  const headers = SampleGenerator.getSampleHeaders();
  const mappings = SampleGenerator.getSampleDefaultMappings();
  const stubs = ComparatorEngine.generateTargetStub(rows, 'CS 2024', mappings.refOldCol, mappings.refNewCol);

  const batchPromptCs = ComparatorEngine.buildBatchCopilotPrompt({
    promptLang: 'cs',
    targetLangName: 'CS 2024',
    rows: rows,
    headers: headers,
    mappings: mappings,
    translations: { 'CS 2024': stubs },
    filterType: 'all_modified'
  });

  assert.ok(batchPromptCs.includes('DÁVKOVÝ PŘEKLAD A AKTUALIZACE ETICKÉHO KODEXU 2026'), 'Musí mít český nadpis');
  assert.ok(batchPromptCs.includes('Art. 04'), 'Musí obsahovat změněný článek 4');
  assert.ok(batchPromptCs.includes('Art. 08'), 'Musí obsahovat nově přidaný článek 8');
  assert.ok(!batchPromptCs.includes('ODSTAVEC: [Art. 01'), 'Nesmí obsahovat nezměněný článek 1');

  // Test anglické varianty dávkového promptu
  const batchPromptEn = ComparatorEngine.buildBatchCopilotPrompt({
    promptLang: 'en',
    targetLangName: 'EN 2024',
    rows: rows,
    headers: headers,
    mappings: mappings,
    translations: { 'CS 2024': stubs },
    filterType: 'all_modified'
  });

  assert.ok(batchPromptEn.includes('BATCH TRANSLATION & UPDATE'), 'Musí mít anglický nadpis');
  assert.ok(batchPromptEn.includes('CRITICAL INSTRUCTIONS'), 'Musí obsahovat anglické pokyny');
  assert.ok(batchPromptEn.includes('Art. 04'), 'Musí obsahovat článek 4');
});

// 12. Test inteligentního parsování odpovědi dávkového překladu (Smart Multi-Paste Parser)
runTest('parseBatchTranslationResponse: Parsování výstupu Copilotu s různými formáty značek', () => {
  const expectedList = [
    { index: 3, id: 'Art. 04' },
    { index: 5, id: 'Art. 06' },
    { index: 13, id: 'Art. 14' }
  ];

  // Test formátu se závorkami [Art. XX]
  const copilotOutputBracket = `
Zde jsou přeložené odstavce:

[Art. 04]
Dary a pohostinnost: Přijímání darů je povoleno pouze do symbolické hodnoty 50 EUR.

[Art. 06]
Ochrana údajů: Všechna osobní data podléhají GDPR a standardům AI 2026.

[Art. 14]
Etické využívání umělé inteligence: Zaměstnanci nesmí zadávat důvěrná data do neověřených LLM nástrojů.
  `;

  const parsedBracket = ComparatorEngine.parseBatchTranslationResponse(copilotOutputBracket, expectedList);
  assert.strictEqual(parsedBracket.length, 3, 'Musí úspěšně naparsovat všechny 3 odstavce');
  assert.strictEqual(parsedBracket[0].index, 3);
  assert.ok(parsedBracket[0].text.includes('50 EUR'));
  assert.strictEqual(parsedBracket[1].index, 5);
  assert.ok(parsedBracket[1].text.includes('GDPR'));
  assert.strictEqual(parsedBracket[2].index, 13);
  assert.ok(parsedBracket[2].text.includes('umělé inteligence'));

  // Test formátu s dvojtečkou a bez závorek (Art. 04: text)
  const copilotOutputColon = `
Art. 04: Přísný limit 50 EUR na dary.
Art. 06: Data podléhají novým pravidlům.
  `;

  const parsedColon = ComparatorEngine.parseBatchTranslationResponse(copilotOutputColon, expectedList);
  assert.strictEqual(parsedColon.length, 2);
  assert.strictEqual(parsedColon[0].index, 3);
  assert.strictEqual(parsedColon[0].text, 'Přísný limit 50 EUR na dary.');
});

// 13. Test manažerského souhrnu změn (Markdown & Word export)
runTest('generateManagerialChangeSummary: Generování manažerského přehledu změn ve formátech Markdown a Word', () => {
  const SampleGenerator = require('../js/core/sample-generator');
  const rows = SampleGenerator.getSampleData();
  const headers = SampleGenerator.getSampleHeaders();
  const mappings = SampleGenerator.getSampleDefaultMappings();

  const summaryCs = ComparatorEngine.generateManagerialChangeSummary({
    rows: rows,
    headers: headers,
    mappings: mappings,
    companyName: 'Acme Holding AG',
    uiLang: 'cs'
  });

  // KPI metriky: 5 beze změny, 6 upravených, 2 přidané, 2 odebrané (celkem 15)
  assert.strictEqual(summaryCs.stats.totalRows, 15);
  assert.strictEqual(summaryCs.stats.unchangedCount, 5);
  assert.strictEqual(summaryCs.stats.modifiedCount, 6);
  assert.strictEqual(summaryCs.stats.addedCount, 2);
  assert.strictEqual(summaryCs.stats.removedCount, 2);

  // Markdown výstup
  assert.ok(summaryCs.markdown.includes('Manažerské shrnutí změn (2024 → 2026)'), 'Musí obsahovat markdown nadpis');
  assert.ok(summaryCs.markdown.includes('Acme Holding AG'), 'Musí obsahovat název společnosti');
  assert.ok(summaryCs.markdown.includes('| **Celkem hodnocených odstavců** | 15 | 100% |'), 'Musí obsahovat tabulku metrik v markdownu');
  assert.ok(summaryCs.markdown.includes('Art. 04'), 'Musí obsahovat detail změněného článku 4');
  assert.ok(summaryCs.markdown.includes('Art. 08'), 'Musí obsahovat detail přidaného článku 8');
  assert.ok(summaryCs.markdown.includes('Art. 10'), 'Musí obsahovat detail zrušeného článku 10');

  // Word (.doc) výstup
  assert.ok(summaryCs.htmlDoc.includes('xmlns:w="urn:schemas-microsoft-com:office:word"'), 'Word soubor musí mít Word namespace hlavičky');
  assert.ok(summaryCs.htmlDoc.includes('mso-page-orientation'), 'Musí obsahovat tisková CSS pravidla pro Word');
  assert.ok(summaryCs.htmlDoc.includes('Acme Holding AG'), 'Musí obsahovat název společnosti');
  assert.ok(summaryCs.htmlDoc.includes('Art. 04'), 'Musí obsahovat články v HTML tabulkách pro Word');
});

// 14. Test úpravy PDF šablony (Branding, vlastní barvy, preambule a podpisy)
runTest('PdfExporter: Uživatelská úprava PDF šablony (branding, accent color, preambule, podpisy)', () => {
  const PdfExporter = require('../js/core/pdf-export');
  const SampleGenerator = require('../js/core/sample-generator');

  const rows = SampleGenerator.getSampleData();
  const headers = SampleGenerator.getSampleHeaders();
  const mappings = SampleGenerator.getSampleDefaultMappings();
  const stubs = ComparatorEngine.generateTargetStub(rows, 'CS 2024', mappings.refOldCol, mappings.refNewCol);

  const customTemplate = {
    companyName: 'Česká Finanční Skupina a.s.',
    docTitle: 'KODEX PROFESIONÁLNÍHO JEDNÁNÍ 2026',
    docSubtitle: 'Závazné znění pro region – {LANG}',
    effectiveDate: 'Účinnost: 15. března 2026',
    versionLabel: 'Edice 3.0 Enterprise',
    accentColor: '#0d9488',
    introPreamble: 'Tento upravený kodex představuje nejvyšší standard etiky pro všechny dceřiné společnosti.',
    showSignatures: true,
    signDept1: 'Oddělení vnitřního auditu',
    signDept2: 'Představenstvo a.s.'
  };

  const customHtml = PdfExporter.generateDocumentHtml({
    rows: rows,
    headers: headers,
    mappings: mappings,
    translations: { 'CS 2024': stubs },
    targetCol: 'CS 2024',
    template: customTemplate,
    uiLang: 'cs'
  });

  assert.ok(customHtml.includes('Česká Finanční Skupina a.s.'), 'Musí obsahovat upravený název firmy');
  assert.ok(customHtml.includes('KODEX PROFESIONÁLNÍHO JEDNÁNÍ 2026'), 'Musí obsahovat upravený titul dokumentu');
  assert.ok(customHtml.includes('Závazné znění pro region – CS 2026'), 'Musí obsahovat {LANG} nahrazený jazykem');
  assert.ok(customHtml.includes('Účinnost: 15. března 2026'), 'Musí obsahovat upravené datum');
  assert.ok(customHtml.includes('Edice 3.0 Enterprise'), 'Musí obsahovat verzi');
  assert.ok(customHtml.includes('#0d9488'), 'Musí obsahovat vlastní accent color');
  assert.ok(customHtml.includes('Oddělení vnitřního auditu'), 'Musí obsahovat upravené podpisy');
});

// 15. Smart Paste: nespárované značky se nikdy nezapíší do cizího řádku
runTest('parseBatchTranslationResponse: neznámé a podobné identifikátory se nepřiřadí špatnému článku', () => {
  const expected = [
    { index: 0, id: 'Art. 01' },
    { index: 1, id: '1.2' },
    { index: 2, id: '1.3' }
  ];
  const parsed = ComparatorEngine.parseBatchTranslationResponse('[Art. 99]\nCizí text\n\n[1.3]\nTřetí odstavec\n\n[Art. 1]\nPrvní', expected);

  assert.deepStrictEqual(parsed.map(p => [p.index, p.text]), [[2, 'Třetí odstavec'], [0, 'První']]);
  assert.deepStrictEqual(parsed.unmatched, ['Art. 99']);
  assert.strictEqual(ComparatorEngine.idsMatch('1.2', '1.3'), false);
  assert.strictEqual(ComparatorEngine.idsMatch('Art. 04', 'Artikel 4'), true);

  // JSON odpověď se stejnými pravidly
  const json = ComparatorEngine.parseBatchTranslationResponse(JSON.stringify([{ id: 'Art. 99', text: 'X' }, { id: '1.2', text: 'Druhý' }]), expected);
  assert.deepStrictEqual(json.map(p => [p.index, p.text]), [[1, 'Druhý']]);
  assert.deepStrictEqual(json.unmatched, ['Art. 99']);
});

runTest('Batch prompt → odpověď Copilota se stejnými značkami → Smart Paste přiřadí všechny články správně', () => {
  const SampleGenerator = require('../js/core/sample-generator');
  const rows = SampleGenerator.getSampleData();
  const headers = SampleGenerator.getSampleHeaders();
  const mappings = SampleGenerator.getSampleDefaultMappings();
  const stubs = ComparatorEngine.generateTargetStub(rows, 'CS 2024', mappings.refOldCol, mappings.refNewCol);

  const prompt = ComparatorEngine.buildBatchCopilotPrompt({
    promptLang: 'cs', targetLangName: 'CS 2024', rows, headers, mappings, translations: { 'CS 2024': stubs }
  });
  // Značky, které prompt požaduje, např. [Art. 04 (Geschenke & Einladungen)]
  const tags = [...prompt.matchAll(/^ODSTAVEC: \[(.+)\] \(Typ/gm)].map(m => m[1]);
  const expectedIdx = stubs.map((s, i) => i).filter(i => ['MODIFIED', 'ADDED'].includes(stubs[i].classification));
  assert.strictEqual(tags.length, expectedIdx.length, 'Prompt obsahuje každý změněný / nový článek');

  const reply = tags.map((tag, k) => `[${tag}]\nPřeklad číslo ${k}`).join('\n\n');
  const expectedList = rows.map((r, i) => ({ index: i, id: r['Artikel-Nr.'] }));
  const parsed = ComparatorEngine.parseBatchTranslationResponse(reply, expectedList);

  assert.deepStrictEqual(parsed.map(p => p.index), expectedIdx);
  assert.deepStrictEqual(parsed.map(p => p.text), tags.map((t, k) => `Překlad číslo ${k}`));
  assert.deepStrictEqual(parsed.unmatched, []);
});

// 16. PDF dokument nesmí tvrdit "oficiální znění" s neaktualizovaným textem
runTest('PdfExporter: nepřeložené články nejsou nahrazeny textem 2024 a dokument je označen jako koncept', () => {
  const PdfExporter = require('../js/core/pdf-export');
  const SampleGenerator = require('../js/core/sample-generator');
  const rows = SampleGenerator.getSampleData();
  const headers = SampleGenerator.getSampleHeaders();
  const mappings = SampleGenerator.getSampleDefaultMappings();

  // Jazyk nebyl ve studiu otevřen → změněný článek 4 nemá překlad
  const draftHtml = PdfExporter.generateDocumentHtml({
    rows, headers, mappings, translations: {}, targetCol: 'CS 2024', uiLang: 'cs',
    template: { versionLabel: '<b>v3</b>' }
  });
  assert.ok(!draftHtml.includes(rows[3]['CS 2024']), 'Zastaralý text 2024 upraveného článku se nesmí objevit');
  assert.ok(draftHtml.includes(rows[0]['CS 2024']), 'Nezměněný článek je převzat z 2024');
  assert.ok(draftHtml.includes('[Překlad chybí]'));
  assert.ok(draftHtml.includes('KONCEPT'));
  assert.ok(draftHtml.includes('&lt;b&gt;v3&lt;/b&gt;') && !draftHtml.includes('<b>v3</b>'), 'Pole šablony musí být escapována');

  // Všechno přeloženo a ověřeno → žádné varování
  const stubs = ComparatorEngine.generateTargetStub(rows, 'CS 2024', mappings.refOldCol, mappings.refNewCol);
  stubs.forEach(s => { if (!s.isVerified) { s.targetNew = 'Hotový překlad\ns novým řádkem'; s.isVerified = true; } });
  const finalHtml = PdfExporter.generateDocumentHtml({
    rows, headers, mappings, translations: { 'CS 2024': stubs }, targetCol: 'CS 2024', uiLang: 'cs'
  });
  assert.ok(!finalHtml.includes('KONCEPT') && !finalHtml.includes('[Překlad chybí]'));
  assert.ok(finalHtml.includes('white-space: pre-wrap'), 'Zalomení řádků v textu článku se zachová');
});

// 17. Víceřádkové texty v manažerském shrnutí
runTest('generateManagerialChangeSummary: víceřádkový text zůstane celý v citaci (Markdown) a se zalomením (Word)', () => {
  const rows = [{ ID: 'Art. 1', Sekce: 'Test', 'DE 2024': 'Zeile eins\nZeile zwei', 'DE 2026': 'Zeile eins\nZeile drei' }];
  const s = ComparatorEngine.generateManagerialChangeSummary({
    rows, headers: ['ID', 'Sekce', 'DE 2024', 'DE 2026'], mappings: { refOldCol: 'DE 2024', refNewCol: 'DE 2026' }
  });
  assert.ok(s.markdown.includes('> Zeile eins\n> Zeile zwei'));
  assert.ok(s.markdown.includes('> Zeile eins\n> Zeile drei'));
  assert.ok(s.htmlDoc.includes('Zeile eins<br>Zeile zwei'));
});

// 18. Lokalizace: každý klíč použitý v aplikaci existuje v obou jazycích
runTest('i18n: všechny klíče z app.js a index.html existují v češtině i angličtině', () => {
  const fs = require('fs');
  const path = require('path');
  const I18n = require('../js/core/i18n');
  const appSrc = fs.readFileSync(path.join(__dirname, '../js/app.js'), 'utf8');
  const html = fs.readFileSync(path.join(__dirname, '../index.html'), 'utf8');
  const used = new Set([...appSrc.matchAll(/\bt\('([a-zA-Z0-9]+)'/g)].map(m => m[1]));
  [...html.matchAll(/data-i18n(?:-placeholder|-title)?="([a-zA-Z0-9]+)"/g)].forEach(m => used.add(m[1]));

  const missingCs = [...used].filter(k => !(k in I18n.translations.cs));
  const missingEn = [...used].filter(k => !(k in I18n.translations.en));
  assert.deepStrictEqual(missingCs, [], 'Chybějící české klíče');
  assert.deepStrictEqual(missingEn, [], 'Chybějící anglické klíče');
  assert.deepStrictEqual(Object.keys(I18n.translations.cs).sort(), Object.keys(I18n.translations.en).sort(), 'cs/en parita');
});

console.log(`\n========================================`);
console.log(`VÝSLEDEK TESTŮ: ${passCount} úspěšných, ${failCount} chyb`);
console.log(`========================================\n`);

if (failCount > 0) {
  process.exit(1);
}


