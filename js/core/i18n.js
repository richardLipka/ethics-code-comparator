/**
 * i18n - Internationalization module (Czech & English)
 * Universal module: works in browser and Node.js.
 */
(function (root, factory) {
  if (typeof define === 'function' && define.amd) {
    define([], factory);
  } else if (typeof module === 'object' && module.exports) {
    module.exports = factory();
  } else {
    root.I18n = factory();
  }
}(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  let currentLang = 'cs';

  const translations = {
    cs: {
      // Header
      appTitle: 'Internationalisation Manager',
      appSubtitle: 'Správa jazykových verzí: porovnání staré a nové verze referenčního textu a aktualizace překladů',
      securityBadge: '🔒 100% offline',
      securityTooltip: 'Tato aplikace běží výhradně lokálně ve vašem prohlížeči. Žádná data se neodesílají na internet ani do externích API.',
      saveProjectBtn: '💾 Uložit projekt',
      saveProjectTitle: 'Stáhne kompletní rozpracovaný projekt jako soubor .json',
      loadProjectBtn: '📂 Načíst projekt',
      loadProjectTitle: 'Načte dříve uložený projekt (.json) a pokračuje v práci',
      diagnosticsBtn: '🧪',
      diagnosticsTitle: 'Diagnostika a testy',
      promptSettingsBtn: '⚙️ Šablona promptu',
      resetProjectBtn: '↺ Nový projekt',
      themeDark: '🌙',
      themeLight: '☀️',
      themeToggleTitle: 'Přepnout tmavý / světlý režim',
      langToggle: '🇬🇧 English',

      // Stepper
      step1Nav: '1. Soubor & sloupce',
      step2Nav: '2. Změny reference',
      step3Nav: '3. Překlad',
      step4Nav: '4. Export',

      // Step 1
      step1Title: 'Krok 1: Nahrání tabulky a výběr sloupců',
      step1Desc: 'Nahrajte sešit, ve kterém je stará a nová verze referenčního textu a dosavadní verze ostatních jazyků (jeden odstavec na řádek).',
      dropzoneTitle: 'Přetáhněte sem soubor (Excel, OpenDocument nebo CSV) nebo klikněte pro výběr',
      dropzoneSubtitle: 'Soubor je zpracován výhradně ve vašem prohlížeči, bez jakéhokoliv odesílání na server.',
      loadSampleBtn: '🚀 Vyzkoušet na ukázkových datech',
      downloadSampleBtn: '📥 Stáhnout ukázkový soubor (.xlsx)',
      sheetSelectLabel: 'List sešitu:',
      hasHeaderRowLabel: 'První řádek obsahuje názvy sloupců',
      mappingTitle: 'Výběr sloupců',
      refOldLabel: '1. Reference – stará verze:',
      refOldHelp: 'Sloupec se starou verzí textu, podle které vznikly dosavadní překlady.',
      refNewLabel: '2. Reference – nová verze:',
      refNewHelp: 'Sloupec s novou verzí textu. Porovnáním se určí změněné, nové a zrušené odstavce.',
      selectColumnPlaceholder: '— vyberte sloupec —',
      optNoOldReference: '— žádná (pouze nová reference) —',
      targetColsLabel: '3. Jazyky k aktualizaci:',
      targetColsHelp: 'Zaškrtněte sloupce se starou verzí jazyků, které chcete aktualizovat; ostatní sloupce se ignorují (v kompletním exportu zůstanou beze změny). Pokud tabulka už obsahuje sloupec pro novou verzi jazyka, vyberte ho – jinak se při exportu vytvoří nový.',
      thUse: 'Použít',
      thColumn: 'Sloupec (stará verze)',
      thLanguage: 'Název jazyka',
      thNewVersion: 'Sloupec pro novou verzi',
      optCreateNewCol: '➕ vytvořit nový sloupec',
      newColNamePlaceholder: 'Název nového sloupce',
      usedAsNewVersionOf: '→ nová verze jazyka {LANG}',
      ignoredColumn: 'ignorováno',
      columnLetterLabel: 'Sloupec {COL}',
      confirmMappingBtn: 'Pokračovat na přehled změn →',

      // Step 2
      step2Title: 'Krok 2: Změny v referenčním textu (stará → nová verze)',
      step2Desc: 'Odstavce jsou rozděleny na beze změny, upravené, nově přidané a zrušené.',
      goToTranslationBtn: 'Přejít k překladu →',
      kpiTotal: 'Celkem odstavců',
      kpiUnchanged: 'Beze změny',
      kpiModified: 'Upraveno',
      kpiAdded: '+ Nově přibylo',
      kpiRemoved: '- Zrušeno',
      thArticle: 'Řádek',
      thClassification: 'Klasifikace',
      thDiff: 'Referenční text se zvýrazněnými změnami (stará → nová verze)',
      step2MgrTitle: 'Shrnutí změn reference:',

      // Step 3
      step3Title: 'Krok 3: Aktualizace překladů',
      step3Desc: 'Nezměněné odstavce jsou převzaty automaticky. U změněných zkopírujte prompt pro Copilot nebo text upravte ručně a zkontrolujte rozdíl.',
      goToExportBtn: 'Přejít k exportu →',
      progressLabel: 'Průběh kontroly aktivního jazyka:',
      filterAll: 'Všechny odstavce',
      filterPending: '⏳ Ke zpracování',
      filterVerified: '✓ Schválené',
      filterModified: '✏️ Pouze upravené',
      filterAdded: '➕ Pouze nové',
      filterRemoved: '🗑️ Pouze zrušené',
      searchPlaceholder: 'Hledat v textu...',
      prevPendingBtn: '▲ Předchozí ke kontrole',
      nextPendingBtn: '▼ Další ke kontrole',
      layout4Cols: '▤ 4 sloupce vedle sebe',
      layout5Cols: '▤ 5 sloupců vedle sebe',
      layout2x2: '▦ 2×2 mřížka',
      auxLangLabel: '🌐 Pomocný jazyk:',
      auxLangSelectTitle: 'Zobrazit vedle reference další jazyk pro porovnání',
      auxLangNone: '-- Žádný (pouze reference) --',

      // Paragraph card
      cardRefPanelTitle: '1. Reference ({LANG}): stará → nová',
      cardAuxPanelTitle: 'Pomocný jazyk ({LANG})',
      cardTargetOldTitle: '2. Stará verze – {LANG}',
      cardTargetNewTitle: '3. Nová verze – {LANG}',
      cardTargetDiffTitle: '4. Změny v {LANG} (stará → nová)',
      copyPromptBtn: '📋 Kopírovat prompt pro Copilot',
      copyVerifyPromptBtn: '🔍 Ověřit překlad (prompt)',
      copyVerifyPromptTitle: 'Zkopírovat prompt pro ověření shodnosti významu a kvality překladu do schránky',
      copyAuxBtn: 'Použít jako základ',
      copyAuxTitle: 'Vložit text pomocného jazyka do editoru',
      prefillOldBtn: '↺ Předvyplnit starou verzí',
      prefillRefBtn: '↺ Vložit referenční text',
      statusDraft: 'Rozpracováno',
      statusOriginal: 'Původní znění',
      toastAuxCopied: 'Text pomocného jazyka byl vložen do editoru.',
      toastVerifyPromptCopied: 'Prompt pro ověření překladu byl zkopírován do schránky.',
      toastNoTranslationToVerify: 'Nejprve vložte nebo napište text nového překladu.',
      markVerified: '✓ Schváleno / Hotovo',
      markToVerify: 'Označit jako hotové',
      textareaPlaceholder: 'Vložte nebo napište novou verzi textu...',
      emptyOriginalText: '(Ve staré verzi neexistoval)',
      noChangesTargetDiff: 'Zatím beze změn oproti staré verzi',
      targetNotTranslatedYet: 'Zatím bez nové verze – rozdíl se zobrazí při psaní.',
      emptyCell: '(Prázdné)',
      match100: '100% shoda',
      changePct: 'Změna: {PCT}%',

      // Badges
      badgeUnchanged: 'Beze změny',
      badgeModified: 'Upraveno',
      badgeAdded: '+ Přidáno',
      badgeRemoved: '- Zrušeno',

      // Step 4
      step4Title: 'Krok 4: Export výsledků',
      step4Desc: 'Stáhněte sešity Excel, dokument pro tisk nebo shrnutí změn. Dokument, shrnutí a částečný sešit se vytvářejí pro jazyk vybraný níže.',
      exportLangLabel: 'Jazyk pro dokument, shrnutí změn a částečný export:',
      tableFormatLabel: 'Formát tabulek:',
      comprehensiveTitle: 'Kompletní sešit (stará + nová verze)',
      comprehensiveDesc: 'Všechny původní sloupce a řádky; nová verze každého jazyka je ve svém sloupci (existujícím nebo nově vloženém hned za starou verzi). Druhý list ukazuje stav kontroly.',
      comprehensiveBtn: '📥 Stáhnout kompletní sešit (.xlsx)',
      cleanTitle: 'Pouze nová verze',
      cleanDesc: 'Tabulka jen s odstavci platnými v nové verzi – nová verze reference a všech jazyků, bez starých verzí. Zrušené odstavce jsou vynechány.',
      cleanBtn: '📥 Stáhnout novou verzi (.xlsx)',
      partialTitle: 'Částečný sešit pro jeden jazyk',
      partialDesc: 'Číslo řádku, vybrané další sloupce (nové verze ostatních jazyků i původní sloupce), stará a nová verze vybraného jazyka a stav kontroly – např. pro předání překladateli.',
      partialContextLabel: 'Další sloupce v exportu:',
      partialGroupNew: 'Nové verze (aktuální stav překladu)',
      partialGroupOriginal: 'Původní sloupce tabulky',
      partialBtn: '📥 Stáhnout částečný sešit (.xlsx)',
      backupTitle: 'Soubor projektu (.json)',
      backupDesc: 'Uloží kompletní stav práce (texty, stav kontroly, nastavení sloupců a šablon). Tlačítka pro uložení a načtení projektu jsou také v záhlaví stránky.',
      backupBtn: '💾 Uložit projekt (.json)',
      pdfTitle: 'Dokument pro tisk / PDF',
      pdfDesc: 'Formátovaný dokument s novou verzí vybraného jazyka (A4). Nepřeložené odstavce jsou vyznačeny a neúplný dokument je označen jako koncept.',
      pdfPrintBtn: '🖨️ Otevřít tisk / Uložit do PDF',
      pdfDownloadHtmlBtn: '📥 Stáhnout HTML',
      pdfAllBtn: '📑 Tisk všech jazyků',
      pdfOptionRefNew: 'reference – nová verze',

      // Modals
      diagTitle: '🧪 Interní diagnostika',
      diagDesc: 'Ověřuje klasifikaci odstavců, textové rozdíly, bezpečnost HTML, generování pracovních dat a exportů přímo v tomto prohlížeči.',
      diagRerunBtn: 'Spustit testy znovu',
      promptSettingsTitle: '⚙️ Šablona promptu pro Copilot',
      promptSettingsDesc: 'Upravte instrukce pro jazykový model (terminologie, tón, pravidla). Zástupné symboly: {TARGET_LANG}, {REF_LANG}, {ROW_ID}, {REF_OLD}, {REF_NEW}, {TARGET_OLD}, {CHANGES_SUMMARY}, {CLASSIFICATION}.',
      promptResetBtn: 'Obnovit výchozí',
      promptSaveBtn: 'Uložit šablonu',

      // Productivity
      saveStatusSaved: 'Vše uloženo',
      saveStatusSaving: 'Ukládám...',
      previewTitle: '👁️ Náhled řádků tabulky',
      previewDesc: 'Zaškrtnutím „Přeskočit“ vyřadíte řádky, které se nemají zpracovat (např. nadpisy nebo poznámky). Prázdné řádky se přeskakují automaticky. Mezery na začátku a konci buněk jsou odstraněny.',
      previewSkipCol: 'Přeskočit',
      previewSkipTitle: 'Řádek se nebude překládat ani exportovat do nové verze; v kompletním sešitu zůstane beze změny.',
      previewEmptyRow: 'prázdný – přeskočen automaticky',
      previewSkippedCount: '{COUNT} přeskočeno',
      previewShowMore: 'Zobrazit dalších {COUNT} řádků (zbývá {TOTAL})',
      toastEmptyRowsSkipped: 'Prázdné řádky ({COUNT}) byly automaticky přeskočeny.',
      matrixTitle: '📊 Postup všech jazyků',
      matrixApproveAllGlobal: '⚡ Schválit nezměněné ve všech jazycích',
      btnApproveUnchanged: '⚡ Schválit nezměněné',
      btnBatchPrompt: '📦 Hromadný prompt pro změněné',
      btnSmartPaste: '📥 Hromadné vložení (Smart Paste)',
      ctrlEnterHint: 'Tip: V editačním poli stiskněte Ctrl + Enter pro schválení a skok na další odstavec.',

      // Smart Paste Modal
      smartPasteTitle: '📥 Hromadné vložení odpovědi Copilota (Smart Paste)',
      smartPasteDesc: 'Vložte odpověď z Copilota. Odstavce se přiřadí podle značek řádků ([#12], Řádek 12: apod.).',
      smartPastePlaceholder: 'Vložte sem celou odpověď od Copilota...\n\nNapř.:\n[#5]\nUpravený překlad odstavce...\n\n[#9]\nPřeklad nového odstavce...',
      smartPasteParseBtn: '🔍 Analyzovat a spárovat',
      smartPasteApplyBtn: '✓ Vložit spárované odstavce',
      smartPasteAutoVerify: 'Zároveň označit vložené odstavce jako schválené',
      smartPasteNoMatches: 'V textu nebyly nalezeny žádné značky odpovídající řádkům tabulky.',

      // Batch Prompt Modal
      batchPromptTitle: '📦 Hromadný prompt pro Copilot',
      batchPromptDesc: 'Prompt obsahuje všechny upravené a nově přidané odstavce aktivního jazyka. Zkopírujte jej do Copilota.',
      batchPromptCopyBtn: '📋 Zkopírovat do schránky',
      closeBtn: 'Zavřít',

      // PDF Template Modal
      editPdfTemplateBtn: '⚙️ Upravit hlavičku dokumentu',
      pdfTemplateTitle: '⚙️ Hlavička a vzhled tiskového dokumentu',
      pdfTemplateDesc: 'Nevyplněná pole se v dokumentu nezobrazí.',
      lblCompanyName: 'Organizace:',
      lblDocTitle: 'Název dokumentu:',
      lblDocSubtitle: 'Podtitul ({LANG} bude nahrazeno jazykem):',
      lblEffectiveDate: 'Datum / platnost:',
      lblVersionLabel: 'Označení verze:',
      lblAccentColor: 'Barva akcentu (HEX):',
      lblPreamble: 'Úvodní text:',
      lblShowSignatures: 'Zobrazit podpisový blok na konci dokumentu',
      lblSignDept1: 'Podpis 1:',
      lblSignDept2: 'Podpis 2:',
      btnSavePdfTemplate: 'Uložit',
      btnResetPdfTemplate: 'Obnovit výchozí',

      // Change summary
      mgrReportTitle: '📋 Shrnutí změn (vybraný jazyk)',
      mgrReportDesc: 'Přehled pro vedení nebo audit: statistika, upravené, nové a zrušené odstavce vybraného jazyka (u překladu i stav kontroly).',
      btnExportMd: '📝 Markdown (.md)',
      btnExportDoc: '📄 Word (.doc)',
      btnExportPdf: '🖨️ Tisk / PDF',

      // Toasts & alerts
      toastPromptCopied: 'Prompt pro Copilot byl zkopírován do schránky!',
      toastFileLoaded: 'Soubor "{NAME}" byl načten.',
      toastSampleLoaded: 'Ukázková data byla načtena.',
      toastSampleDownloaded: 'Ukázkový soubor byl stažen.',
      toastProjectSaved: 'Projekt byl uložen do souboru {NAME}.',
      toastProjectLoaded: 'Projekt byl načten.',
      toastMappingSaved: 'Výběr sloupců byl uložen.',
      toastAllApproved: 'Všechny zobrazené odstavce jsou již schválené!',
      toastWorkbookGenerated: 'Sešit "{NAME}" byl vygenerován.',
      alertSelectRefCols: 'Vyberte oba referenční sloupce (starou i novou verzi).',
      alertDifferentRefCols: 'Stará a nová verze reference musí být různé sloupce.',
      alertSelectTargetLangs: 'Označte alespoň jeden jazyk k aktualizaci.',
      alertConfirmReset: 'Opravdu chcete zahodit rozpracovaný projekt a začít znovu?',
      saveStatusError: 'Neuloženo – uložte projekt (.json)',
      toastSaveFailed: 'Automatické uložení do prohlížeče selhalo (pravděpodobně plné úložiště). Zálohujte práci tlačítkem „Uložit projekt“.',
      toastBulkApproved: 'Schváleno {COUNT} nezměněných odstavců.',
      batchPromptCopied: 'Hromadný prompt byl zkopírován do schránky!',
      smartPasteParsedCount: 'Spárované odstavce',
      smartPasteApplied: 'Vloženo {COUNT} odstavců.',
      smartPasteUnmatched: 'Nespárované značky (nebudou vloženy): {LIST}',
      mgrExportMdSuccess: 'Shrnutí změn bylo staženo jako Markdown.',
      mgrExportDocSuccess: 'Shrnutí změn bylo staženo jako dokument Word.',
      pdfTemplateSaved: 'Hlavička dokumentu byla uložena.',
      pdfTemplateReset: 'Hlavička dokumentu byla obnovena na výchozí hodnoty.',
      toastNoData: 'Nejsou načtena žádná data.',
      toastSelectTargetLang: 'Vyberte jazyk.',
      toastExportUnverified: 'Pozor: {COUNT} odstavců (součet za všechny jazyky) zatím není schváleno nebo chybí jejich nová verze – viz list se stavem kontroly.',
      toastExportUnverifiedClean: 'Pozor: {COUNT} odstavců (součet za všechny jazyky) zatím není schváleno nebo chybí jejich nová verze – export ještě není finální.',
      toastExportUnverifiedLanguage: 'Pozor: {COUNT} odstavců tohoto jazyka zatím není schváleno nebo chybí jejich nová verze (viz sloupec se stavem).',
      toastPdfSelectLang: 'Vyberte jazyk pro export.',
      toastPdfHtmlDownloaded: 'Dokument byl stažen jako HTML.',
      toastNoLanguages: 'Žádné jazyky k exportu.',
      confirmPrintAll: 'Otevřít tisková okna pro všech {COUNT} jazyků?',
      confirmResumeProject: 'V prohlížeči byl nalezen rozpracovaný projekt z {DATE}. Chcete v něm pokračovat?',
      alertNoSheets: 'Nahraný soubor neobsahuje žádné listy.',
      alertUnsupportedFile: 'Soubor "{NAME}" nelze načíst – podporované formáty: {FORMATS}.',
      alertEmptySheet: 'Vybraný list je prázdný.',
      alertFileReadError: 'Soubor se nepodařilo načíst: {MSG}',
      alertProjectLoadError: 'Projekt se nepodařilo načíst. {MSG}',
      projectErrInvalidJson: 'Soubor „{NAME}“ není platný soubor projektu (.json) – jeho obsah nelze přečíst jako JSON.',
      projectErrSpreadsheet: 'Soubor „{NAME}“ je tabulka, ne soubor projektu. Tabulku nahrajte v Kroku 1 (přetažením nebo kliknutím na oblast pro nahrání); „Načíst projekt“ slouží jen pro soubory uložené tlačítkem „Uložit projekt“.',
      projectErrNotProject: 'Soubor „{NAME}“ je JSON s jinou strukturou, nejde o projekt překladu (očekávána tabulka „rows“ a výběr sloupců „mappings“; nalezeno: {FOUND}). Vyberte soubor uložený tlačítkem „Uložit projekt“.',
      projectErrForeignApp: 'Soubor „{NAME}“ patří jiné aplikaci ({APP}) a nejde o projekt překladu.',
      projectErrNoRows: 'Projekt „{NAME}“ neobsahuje žádné řádky tabulky.',
      projectErrBadRows: 'Projekt „{NAME}“ je poškozený: řádek {ROW} tabulky nemá očekávaný formát.',
      projectErrBadMappings: 'Projekt „{NAME}“ je poškozený: výběr sloupců (mappings) nemá očekávaný formát.',
      projectErrMissingColumn: 'Projekt „{NAME}“ je poškozený: vybraný sloupec „{COL}“ v tabulce neexistuje.',
      projectLoadKeptState: 'Aktuální práce zůstala beze změny.',
      projectLoadedWithWarnings: 'Projekt byl načten s upozorněním:',
      projectWarnNoAppId: 'soubor není označen jako projekt této aplikace, byl načten podle struktury – zkontrolujte data',
      projectWarnNewerVersion: 'soubor uložila novější verze aplikace ({VERSION}), některá nastavení nemusí být načtena',
      projectWarnUnusedTranslations: 'rozpracované texty sloupců, které nejsou vybrané jako jazyk, byly vynechány ({LIST})',
      projectWarnRowCountMismatch: 'jazyk {LANG} má uloženo {COUNT} textů, ale tabulka má {ROWS} řádků – zkontrolujte přiřazení',
      projectWarnInvalidTranslations: 'část rozpracovaných textů měla neočekávaný formát a byla vynechána',
      projectWarnInvalidSettings: 'některá nastavení (šablony, export) měla neočekávaný formát a byla nahrazena výchozími',
      confirmReplaceProject: 'Načtením projektu „{NAME}“ se nahradí aktuální rozpracovaná práce. Pokračovat?',
      rowsLabel: 'řádků',
      rowLabel: 'Řádek',
      progressDone: 'hotovo',
      noTargetLang: 'Žádný jazyk k zobrazení.',
      noArticlesMatch: 'Žádné odstavce neodpovídají filtru nebo hledání.',
      charsLabel: 'znaků',
      newColSuffix: ' (nová verze)',
      addLanguageTitle: 'Přidat nový jazyk (není v původním souboru)',
      addLanguageDesc: 'Pro jazyk vytvořený navíc bude k dispozici jeden sloupec a prompt pro Copilot bude prostý překlad bez historie změn.',
      newLangNamePlaceholder: 'Název nového jazyka (např. Polština)',
      newLangColNamePlaceholder: 'Název sloupce v exportu (volitelné)',
      addNewLangBtn: '➕ Přidat jazyk',
      badgeNewLanguage: 'nový jazyk',
      noOldVersion: '(nově vytvářený – bez staré verze)',
      removeLangTitle: 'Odebrat tento jazyk',
      cardTargetNewSingleTitle: '2. Překlad – {LANG}',
      toastLangAlreadyExists: 'Jazyk s tímto názvem již v projektu existuje.',
      toastEnterLangName: 'Zadejte prosím název nového jazyka.',
      toastNewLangAdded: 'Nový jazyk „{NAME}“ byl přidán.',
      toastNewLangRemoved: 'Jazyk „{NAME}“ byl odebrán.',
      batchPromptDescNew: 'Prompt obsahuje všechny odstavce k překladu do aktivního jazyka (prostý překlad nové verze reference). Zkopírujte jej do Copilota.'
    },

    en: {
      // Header
      appTitle: 'Internationalisation Manager',
      appSubtitle: 'Manage language versions: compare the old and new reference text and update the translations',
      securityBadge: '🔒 100% offline',
      securityTooltip: 'This application runs entirely locally in your browser. No data is sent to the internet or any external API.',
      saveProjectBtn: '💾 Save project',
      saveProjectTitle: 'Downloads the complete work in progress as a .json file',
      loadProjectBtn: '📂 Load project',
      loadProjectTitle: 'Loads a previously saved project (.json) and continues the work',
      diagnosticsBtn: '🧪',
      diagnosticsTitle: 'Diagnostics and tests',
      promptSettingsBtn: '⚙️ Prompt template',
      resetProjectBtn: '↺ New project',
      themeDark: '🌙',
      themeLight: '☀️',
      themeToggleTitle: 'Toggle dark / light mode',
      langToggle: '🇨🇿 Čeština',

      // Stepper
      step1Nav: '1. File & columns',
      step2Nav: '2. Reference changes',
      step3Nav: '3. Translation',
      step4Nav: '4. Export',

      // Step 1
      step1Title: 'Step 1: Upload the table and choose columns',
      step1Desc: 'Upload a workbook with the old and the new version of the reference text and the current versions of the other languages (one paragraph per row).',
      dropzoneTitle: 'Drag & drop a file (Excel, OpenDocument or CSV) here or click to browse',
      dropzoneSubtitle: 'The file is processed entirely in your browser and never uploaded to a server.',
      loadSampleBtn: '🚀 Try with sample data',
      downloadSampleBtn: '📥 Download sample file (.xlsx)',
      sheetSelectLabel: 'Workbook sheet:',
      hasHeaderRowLabel: 'The first row contains column names',
      mappingTitle: 'Column selection',
      refOldLabel: '1. Reference – old version:',
      refOldHelp: 'Column with the old version of the text the current translations are based on.',
      refNewLabel: '2. Reference – new version:',
      refNewHelp: 'Column with the new version of the text. Comparing both finds modified, new and removed paragraphs.',
      selectColumnPlaceholder: '— select a column —',
      optNoOldReference: '— none (new reference only) —',
      targetColsLabel: '3. Languages to update:',
      targetColsHelp: 'Tick the columns holding the old version of the languages you want to update; other columns are ignored (they stay unchanged in the full export). If the table already has a column for the new version of a language, select it – otherwise a new one is created on export.',
      thUse: 'Use',
      thColumn: 'Column (old version)',
      thLanguage: 'Language name',
      thNewVersion: 'Column for the new version',
      optCreateNewCol: '➕ create a new column',
      newColNamePlaceholder: 'New column name',
      usedAsNewVersionOf: '→ new version of {LANG}',
      ignoredColumn: 'ignored',
      columnLetterLabel: 'Column {COL}',
      confirmMappingBtn: 'Continue to the change overview →',

      // Step 2
      step2Title: 'Step 2: Changes in the reference text (old → new version)',
      step2Desc: 'Paragraphs are classified as unchanged, modified, newly added and removed.',
      goToTranslationBtn: 'Go to translation →',
      kpiTotal: 'Total paragraphs',
      kpiUnchanged: 'Unchanged',
      kpiModified: 'Modified',
      kpiAdded: '+ Newly added',
      kpiRemoved: '- Removed',
      thArticle: 'Row',
      thClassification: 'Classification',
      thDiff: 'Reference text with highlighted changes (old → new version)',
      step2MgrTitle: 'Reference change summary:',

      // Step 3
      step3Title: 'Step 3: Update the translations',
      step3Desc: 'Unchanged paragraphs are taken over automatically. For changed ones copy the Copilot prompt or edit the text yourself and check the difference.',
      goToExportBtn: 'Go to export →',
      progressLabel: 'Review progress of the active language:',
      filterAll: 'All paragraphs',
      filterPending: '⏳ To do',
      filterVerified: '✓ Approved',
      filterModified: '✏️ Modified only',
      filterAdded: '➕ New only',
      filterRemoved: '🗑️ Removed only',
      searchPlaceholder: 'Search text...',
      prevPendingBtn: '▲ Previous to review',
      nextPendingBtn: '▼ Next to review',
      layout4Cols: '▤ 4 columns side by side',
      layout5Cols: '▤ 5 columns side by side',
      layout2x2: '▦ 2×2 grid',
      auxLangLabel: '🌐 Auxiliary language:',
      auxLangSelectTitle: 'Show another language alongside the reference for comparison',
      auxLangNone: '-- None (reference only) --',

      // Paragraph card
      cardRefPanelTitle: '1. Reference ({LANG}): old → new',
      cardAuxPanelTitle: 'Auxiliary language ({LANG})',
      cardTargetOldTitle: '2. Old version – {LANG}',
      cardTargetNewTitle: '3. New version – {LANG}',
      cardTargetDiffTitle: '4. Changes in {LANG} (old → new)',
      copyPromptBtn: '📋 Copy Copilot prompt',
      copyVerifyPromptBtn: '🔍 Verify translation (prompt)',
      copyVerifyPromptTitle: 'Copy prompt for translation accuracy and equivalence check to clipboard',
      copyAuxBtn: 'Use as base',
      copyAuxTitle: 'Insert auxiliary language text into the editor',
      prefillOldBtn: '↺ Prefill with old version',
      prefillRefBtn: '↺ Insert reference text',
      statusDraft: 'Draft',
      statusOriginal: 'Original text',
      toastAuxCopied: 'Auxiliary language text inserted into the editor.',
      toastVerifyPromptCopied: 'Translation verification prompt copied to clipboard.',
      toastNoTranslationToVerify: 'Please enter a new translation text first.',
      markVerified: '✓ Approved / Done',
      markToVerify: 'Mark as done',
      textareaPlaceholder: 'Paste or type the new version of the text...',
      emptyOriginalText: '(Did not exist in the old version)',
      noChangesTargetDiff: 'No changes from the old version yet',
      targetNotTranslatedYet: 'No new version yet – the difference appears as you type.',
      emptyCell: '(Empty)',
      match100: '100% match',
      changePct: 'Change: {PCT}%',

      // Badges
      badgeUnchanged: 'Unchanged',
      badgeModified: 'Modified',
      badgeAdded: '+ Added',
      badgeRemoved: '- Removed',

      // Step 4
      step4Title: 'Step 4: Export the results',
      step4Desc: 'Download Excel workbooks, a printable document or a change summary. The document, the summary and the partial workbook are created for the language selected below.',
      exportLangLabel: 'Language for the document, change summary and partial export:',
      tableFormatLabel: 'Table format:',
      comprehensiveTitle: 'Full workbook (old + new version)',
      comprehensiveDesc: 'All original columns and rows; the new version of every language is in its own column (existing, or inserted right after the old version). A second sheet shows the review status.',
      comprehensiveBtn: '📥 Download full workbook (.xlsx)',
      cleanTitle: 'New version only',
      cleanDesc: 'A table with the paragraphs valid in the new version only – the new reference and all languages, without the old versions. Removed paragraphs are left out.',
      cleanBtn: '📥 Download new version (.xlsx)',
      partialTitle: 'Partial workbook for one language',
      partialDesc: 'Row number, chosen additional columns (new versions of other languages and original columns), old and new version of the selected language and its review status – e.g. to hand over to a translator.',
      partialContextLabel: 'Additional columns in the export:',
      partialGroupNew: 'New versions (current translation work)',
      partialGroupOriginal: 'Original columns of the table',
      partialBtn: '📥 Download partial workbook (.xlsx)',
      backupTitle: 'Project file (.json)',
      backupDesc: 'Saves the complete state of your work (texts, review status, column and template settings). Save and load buttons are also in the page header.',
      backupBtn: '💾 Save project (.json)',
      pdfTitle: 'Printable document / PDF',
      pdfDesc: 'A formatted A4 document with the new version of the selected language. Untranslated paragraphs are marked and an incomplete document is labelled as a draft.',
      pdfPrintBtn: '🖨️ Open print / Save as PDF',
      pdfDownloadHtmlBtn: '📥 Download HTML',
      pdfAllBtn: '📑 Print all languages',
      pdfOptionRefNew: 'reference – new version',

      // Modals
      diagTitle: '🧪 Internal diagnostics',
      diagDesc: 'Checks paragraph classification, text diffs, HTML safety, working-data generation and exports right in this browser.',
      diagRerunBtn: 'Run tests again',
      promptSettingsTitle: '⚙️ Copilot prompt template',
      promptSettingsDesc: 'Adjust the instructions for the language model (terminology, tone, rules). Placeholders: {TARGET_LANG}, {REF_LANG}, {ROW_ID}, {REF_OLD}, {REF_NEW}, {TARGET_OLD}, {CHANGES_SUMMARY}, {CLASSIFICATION}.',
      promptResetBtn: 'Restore default',
      promptSaveBtn: 'Save template',

      // Productivity
      saveStatusSaved: 'All changes saved',
      saveStatusSaving: 'Saving...',
      previewTitle: '👁️ Table rows',
      previewDesc: 'Tick "Skip" to exclude rows that should not be processed (e.g. headings or notes). Empty rows are skipped automatically. Leading and trailing whitespace of cells is removed.',
      previewSkipCol: 'Skip',
      previewSkipTitle: 'The row will not be translated or exported to the new version; the full workbook keeps it unchanged.',
      previewEmptyRow: 'empty – skipped automatically',
      previewSkippedCount: '{COUNT} skipped',
      previewShowMore: 'Show {COUNT} more rows ({TOTAL} left)',
      toastEmptyRowsSkipped: '{COUNT} empty rows were skipped automatically.',
      matrixTitle: '📊 Progress of all languages',
      matrixApproveAllGlobal: '⚡ Approve unchanged in all languages',
      btnApproveUnchanged: '⚡ Approve unchanged',
      btnBatchPrompt: '📦 Batch prompt for changes',
      btnSmartPaste: '📥 Smart Paste',
      ctrlEnterHint: 'Tip: press Ctrl + Enter in the editor to approve the paragraph and jump to the next one.',

      // Smart Paste Modal
      smartPasteTitle: '📥 Smart batch paste of the Copilot reply',
      smartPasteDesc: 'Paste the Copilot reply. Paragraphs are matched by their row tags ([#12], Row 12: etc.).',
      smartPastePlaceholder: 'Paste the full Copilot reply here...\n\nExample:\n[#5]\nUpdated translation...\n\n[#9]\nTranslation of a new paragraph...',
      smartPasteParseBtn: '🔍 Analyze and match',
      smartPasteApplyBtn: '✓ Apply matched paragraphs',
      smartPasteAutoVerify: 'Also mark applied paragraphs as approved',
      smartPasteNoMatches: 'No tags matching table rows were found in the text.',

      // Batch Prompt Modal
      batchPromptTitle: '📦 Copilot batch prompt',
      batchPromptDesc: 'The prompt contains all modified and newly added paragraphs of the active language. Copy it into Copilot.',
      batchPromptCopyBtn: '📋 Copy to clipboard',
      closeBtn: 'Close',

      // PDF Template Modal
      editPdfTemplateBtn: '⚙️ Edit document header',
      pdfTemplateTitle: '⚙️ Printable document header and style',
      pdfTemplateDesc: 'Empty fields are not shown in the document.',
      lblCompanyName: 'Organization:',
      lblDocTitle: 'Document title:',
      lblDocSubtitle: 'Subtitle ({LANG} is replaced by the language):',
      lblEffectiveDate: 'Date / validity:',
      lblVersionLabel: 'Version label:',
      lblAccentColor: 'Accent color (HEX):',
      lblPreamble: 'Introduction:',
      lblShowSignatures: 'Show a signature block at the end of the document',
      lblSignDept1: 'Signature 1:',
      lblSignDept2: 'Signature 2:',
      btnSavePdfTemplate: 'Save',
      btnResetPdfTemplate: 'Restore defaults',

      // Change summary
      mgrReportTitle: '📋 Change summary (selected language)',
      mgrReportDesc: 'An overview for management or audit: statistics, modified, new and removed paragraphs of the selected language (for a translation also its review status).',
      btnExportMd: '📝 Markdown (.md)',
      btnExportDoc: '📄 Word (.doc)',
      btnExportPdf: '🖨️ Print / PDF',

      // Toasts & alerts
      toastPromptCopied: 'Copilot prompt copied to clipboard!',
      toastFileLoaded: 'File "{NAME}" loaded.',
      toastSampleLoaded: 'Sample data loaded.',
      toastSampleDownloaded: 'Sample file downloaded.',
      toastProjectSaved: 'Project saved to {NAME}.',
      toastProjectLoaded: 'Project loaded.',
      toastMappingSaved: 'Column selection saved.',
      toastAllApproved: 'All displayed paragraphs are already approved!',
      toastWorkbookGenerated: 'Workbook "{NAME}" generated.',
      alertSelectRefCols: 'Select both reference columns (old and new version).',
      alertDifferentRefCols: 'The old and new reference version must be different columns.',
      alertSelectTargetLangs: 'Tick at least one language to update.',
      alertConfirmReset: 'Do you really want to discard the work in progress and start over?',
      saveStatusError: 'Not saved – save the project (.json)',
      toastSaveFailed: 'Autosave to the browser failed (storage is probably full). Back up your work with "Save project".',
      toastBulkApproved: '{COUNT} unchanged paragraphs approved.',
      batchPromptCopied: 'Batch prompt copied to clipboard!',
      smartPasteParsedCount: 'Matched paragraphs',
      smartPasteApplied: '{COUNT} paragraphs applied.',
      smartPasteUnmatched: 'Unmatched tags (will not be applied): {LIST}',
      mgrExportMdSuccess: 'Change summary downloaded as Markdown.',
      mgrExportDocSuccess: 'Change summary downloaded as a Word document.',
      pdfTemplateSaved: 'Document header saved.',
      pdfTemplateReset: 'Document header reset to defaults.',
      toastNoData: 'No data loaded.',
      toastSelectTargetLang: 'Select a language.',
      toastExportUnverified: 'Note: {COUNT} paragraphs (total over all languages) are not approved yet or have no new version – see the review status sheet.',
      toastExportUnverifiedClean: 'Note: {COUNT} paragraphs (total over all languages) are not approved yet or have no new version – the export is not final yet.',
      toastExportUnverifiedLanguage: 'Note: {COUNT} paragraphs of this language are not approved yet or have no new version (see the status column).',
      toastPdfSelectLang: 'Select a language to export.',
      toastPdfHtmlDownloaded: 'Document downloaded as HTML.',
      toastNoLanguages: 'No languages available to export.',
      confirmPrintAll: 'Open print windows for all {COUNT} languages?',
      confirmResumeProject: 'A saved project from {DATE} was found. Do you want to resume it?',
      alertNoSheets: 'The uploaded file contains no sheets.',
      alertUnsupportedFile: 'The file "{NAME}" cannot be loaded – supported formats: {FORMATS}.',
      alertEmptySheet: 'The selected sheet is empty.',
      alertFileReadError: 'The file could not be read: {MSG}',
      alertProjectLoadError: 'The project could not be loaded. {MSG}',
      projectErrInvalidJson: 'The file "{NAME}" is not a valid project file (.json) – its content cannot be read as JSON.',
      projectErrSpreadsheet: 'The file "{NAME}" is a spreadsheet, not a project file. Upload spreadsheets in Step 1 (drop the file on the upload area or click it); "Load project" only opens files saved with "Save project".',
      projectErrNotProject: 'The file "{NAME}" is JSON with a different structure, not a translation project (expected a "rows" table and a "mappings" column selection; found: {FOUND}). Select a file saved with "Save project".',
      projectErrForeignApp: 'The file "{NAME}" belongs to another application ({APP}) and is not a translation project.',
      projectErrNoRows: 'The project "{NAME}" contains no table rows.',
      projectErrBadRows: 'The project "{NAME}" is damaged: table row {ROW} does not have the expected format.',
      projectErrBadMappings: 'The project "{NAME}" is damaged: the column selection (mappings) does not have the expected format.',
      projectErrMissingColumn: 'The project "{NAME}" is damaged: the selected column "{COL}" does not exist in the table.',
      projectLoadKeptState: 'Your current work is unchanged.',
      projectLoadedWithWarnings: 'The project was loaded with warnings:',
      projectWarnNoAppId: 'the file is not marked as a project of this application and was loaded by its structure – please check the data',
      projectWarnNewerVersion: 'the file was saved by a newer version of the application ({VERSION}); some settings may not be loaded',
      projectWarnUnusedTranslations: 'work for columns that are not selected as languages was skipped ({LIST})',
      projectWarnRowCountMismatch: 'language {LANG} has {COUNT} saved texts but the table has {ROWS} rows – please check the assignment',
      projectWarnInvalidTranslations: 'some saved texts had an unexpected format and were skipped',
      projectWarnInvalidSettings: 'some settings (templates, export) had an unexpected format and were replaced by defaults',
      confirmReplaceProject: 'Loading the project "{NAME}" replaces your current work in progress. Continue?',
      rowsLabel: 'rows',
      rowLabel: 'Row',
      progressDone: 'completed',
      noTargetLang: 'No language to display.',
      noArticlesMatch: 'No paragraphs match the current filter or search.',
      charsLabel: 'chars',
      newColSuffix: ' (new)',
      addLanguageTitle: 'Add a new language (not in original file)',
      addLanguageDesc: 'For an added language, a single column will be available and the prompt will be a plain translation without change history.',
      newLangNamePlaceholder: 'New language name (e.g. Polish)',
      newLangColNamePlaceholder: 'Export column name (optional)',
      addNewLangBtn: '➕ Add language',
      badgeNewLanguage: 'new language',
      noOldVersion: '(newly created – no old version)',
      removeLangTitle: 'Remove this language',
      cardTargetNewSingleTitle: '2. Translation – {LANG}',
      toastLangAlreadyExists: 'A language with this name already exists in the project.',
      toastEnterLangName: 'Please enter a name for the new language.',
      toastNewLangAdded: 'New language "{NAME}" has been added.',
      toastNewLangRemoved: 'Language "{NAME}" has been removed.',
      batchPromptDescNew: 'The prompt contains all paragraphs to translate into the active language (plain translation of the new reference version). Copy it into Copilot.'
    }
  };

  function initLanguage() {
    let saved = null;
    try {
      saved = (typeof window !== 'undefined' && window.localStorage)
        ? window.localStorage.getItem('comparator_lang')
        : null;
    } catch (e) {
      saved = null;
    }
    currentLang = saved === 'en' ? 'en' : 'cs';
    return currentLang;
  }

  function getLanguage() {
    return currentLang;
  }

  function setLanguage(lang) {
    currentLang = lang === 'en' ? 'en' : 'cs';
    try {
      if (typeof window !== 'undefined' && window.localStorage) {
        window.localStorage.setItem('comparator_lang', currentLang);
      }
    } catch (e) {
      // Language preference is a convenience only
    }
    return currentLang;
  }

  function toggleLanguage() {
    return setLanguage(currentLang === 'cs' ? 'en' : 'cs');
  }

  function t(key, params) {
    const langDict = translations[currentLang] || translations.cs;
    let text = langDict[key] || translations.cs[key] || key;

    if (params && typeof params === 'object') {
      Object.keys(params).forEach(function (paramKey) {
        text = text.replace(new RegExp('\\{' + paramKey + '\\}', 'g'), function () { return String(params[paramKey]); });
      });
    }

    return text;
  }

  return {
    initLanguage: initLanguage,
    getLanguage: getLanguage,
    setLanguage: setLanguage,
    toggleLanguage: toggleLanguage,
    t: t,
    translations: translations
  };
}));
