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
      appTitle: 'Etický Kodex – Srovnávač & Překladový Asistent',
      appSubtitle: 'Synchronizace referenční němčiny (2024 → 2026) a vícejazyčných mutací',
      securityBadge: '🔒 100% Offline & Bezpečné',
      securityTooltip: 'Tato aplikace běží výhradně lokálně ve vašem prohlížeči. Žádná data se neodesílají na internet ani do externích API.',
      diagnosticsBtn: '🧪 Diagnostika a testy',
      promptSettingsBtn: '⚙️ Šablona promptu',
      resetProjectBtn: '↺ Nový projekt',
      themeDark: '🌙 Tmavý režim',
      themeLight: '☀️ Světlý režim',
      langToggle: '🇬🇧 English',

      // Stepper
      step1Nav: '1. Soubor & Mapování sloupců',
      step2Nav: '2. Přehled referenčních změn (DE)',
      step3Nav: '3. Pracovní plocha překladu',
      step4Nav: '4. Export & Stažení výsledků',

      // Step 1
      step1Title: 'Krok 1: Nahrání Excel tabulky a označení sloupců',
      step1Desc: 'Nahrajte sešit s referenčními verzemi (DE 2024 a DE 2026) a dosavadními jazykovými překlady.',
      dropzoneTitle: 'Přetáhněte sem Excel soubor (.xlsx, .xls) nebo klikněte pro výběr',
      dropzoneSubtitle: 'Soubor je zpracován výhradně ve vašem prohlížeči, bez jakéhokoliv odesílání na server.',
      loadSampleBtn: '🚀 Vyzkoušet na ukázkovém kodexu (15 článků, 12 jazyků)',
      downloadSampleBtn: '📥 Stáhnout vzorový soubor (Demo XLSX)',
      importProjectBtn: '📂 Načíst rozpracovaný projekt (.json)',
      sheetSelectLabel: 'Vyberte list sešitu k analýze:',
      mappingTitle: 'Označení referenčních sloupců a cílových jazyků',
      autoDetectBtn: '🔍 Automaticky detekovat sloupce',
      refOldLabel: '1. Referenční sloupec 2024 (Stará němčina):',
      refOldHelp: 'Slouží k porovnání se starými jazykovými mutacemi.',
      refNewLabel: '2. Referenční sloupec 2026 (Nová němčina):',
      refNewHelp: 'Slouží k detekci změn, nových a smazaných odstavců.',
      targetColsLabel: '3. Cílové jazykové mutace k překladu (Verze 2024):',
      targetColsHelp: 'Zaškrtněte všechny sloupce s existujícími ověřenými překlady, pro které chcete sestavit novou verzi 2026.',
      confirmMappingBtn: 'Pokračovat na přehled změn referenční verze →',

      // Step 2
      step2Title: 'Krok 2: Přehled změn v německé referenční verzi (2024 vs 2026)',
      step2Desc: 'Automatická detekce odstavců: beze změny, upraveno, nově přidáno a odstraněno.',
      goToTranslationBtn: 'Přejít do pracovní plochy překladu →',
      kpiTotal: 'Celkem odstavců',
      kpiUnchanged: 'Beze změny',
      kpiModified: 'Upraveno',
      kpiAdded: '+ Nově přibylo',
      kpiRemoved: '- Odstraněno',
      thArticle: 'Článek / ID',
      thClassification: 'Klasifikace',
      thDiff: 'Znění v němčině & Zvýrazněný textový diff (2024 → 2026)',

      // Step 3
      step3Title: 'Krok 3: Pracovní plocha překladu & vizuální kontrola diffů',
      step3Desc: 'Nezměněné části jsou automaticky zkopírovány. U změn zkopírujte prompt pro Copilot a vložte výsledek s live diff kontrolou.',
      goToExportBtn: 'Přejít k exportu hotových verzí →',
      progressLabel: 'Průběh revize a schválení pro aktivní jazyk:',
      filterAll: 'Všechny odstavce',
      filterPending: '⏳ Ke zpracování',
      filterVerified: '✓ Schválené',
      filterModified: '✏️ Pouze upravené',
      filterAdded: '➕ Pouze nově přidané',
      filterRemoved: '🗑️ Pouze smazané',
      searchPlaceholder: 'Hledat text v odstavcích...',
      prevPendingBtn: '▲ Předchozí ke kontrole',
      nextPendingBtn: '▼ Další ke kontrole',
      layout4Cols: '▤ 4 sloupce vedle sebe (Plná šířka)',
      layout2x2: '▦ 2×2 mřížka',
      
      // Article Card
      cardRefPanelTitle: '1. Referenční němčina (DE 2024 → 2026 diff)',
      cardTargetOldTitle: '2. Původní ověřený překlad ({LANG} 2024)',
      cardTargetNewTitle: '3. Nový překlad pro rok 2026 ({LANG})',
      cardTargetDiffTitle: '4. Diff vašeho překladu ({LANG} 2024 → 2026)',
      cardCompareSummary: 'Vizuální kontrola shody změn (Německá reference vs Váš překlad)',
      cardCompareHint: 'Ověřte, zda změny v cílovém jazyce odpovídají referenci',
      copyPromptBtn: '📋 Kopírovat prompt pro Copilot',
      prefillOldBtn: '↺ Předvyplnit z 2024',
      prefillRefBtn: '↺ Vložit německý text',
      markVerified: '✓ Schváleno / Hotovo',
      markToVerify: 'Označit jako hotové',
      textareaPlaceholder: 'Vložte nebo napište aktualizovaný překlad pro rok 2026...',
      emptyOriginalText: '(V roce 2024 neexistoval)',
      noChangesTargetDiff: 'Zatím beze změn oproti 2024',
      targetNotTranslatedYet: 'Zatím bez nového překladu – rozdíl se zobrazí při psaní.',
      emptyCell: '(Prázdné)',
      match100: '100% shoda',
      changePct: 'Změna: {PCT}%',

      // Badges
      badgeUnchanged: 'Beze změny',
      badgeModified: 'Upraveno',
      badgeAdded: '+ Přidáno',
      badgeRemoved: '- Smazáno',
      badgeVerified: 'Schváleno',

      // Step 4
      step4Title: 'Krok 4: Export hotového Excel sešitu a záloha projektu',
      step4Desc: 'Vyberte preferovanou formu výstupu – kompletní srovnávací tabulku se všemi sloupci nebo čistou finální verzi 2026.',
      comprehensiveTitle: 'Kompletní srovnávací sešit (2024 + 2026)',
      comprehensiveDesc: 'Zachová všechny původní sloupce i řádky a vedle každého jazyka z roku 2024 vloží nový sloupec s verzí 2026. Ideální pro archivaci, audit a interní schvalování změn.',
      comprehensiveBtn: '📥 Stáhnout kompletní sešit (.xlsx)',
      cleanTitle: 'Čistá nová verze 2026 (Finální publikace)',
      cleanDesc: 'Vygeneruje čistou tabulku obsahující výhradně platné články pro rok 2026 ve všech jazycích. Odstraněné články jsou automaticky vynechány.',
      cleanBtn: '📥 Stáhnout čistou verzi 2026 (.xlsx)',
      backupTitle: 'Záložní soubor projektu (.json)',
      backupDesc: 'Uloží kompletní stav vaší práce včetně všech rozpracovaných textů a mapování do jednoho souboru, který můžete kdykoliv znovu načíst.',
      backupBtn: '💾 Exportovat projekt (.json)',
      pdfTitle: 'Formátované PDF dokumenty (pro každý jazyk)',
      pdfDesc: 'Vygeneruje reprezentativní, graficky upravený dokument kodexu pro rok 2026 s formátováním pro A4, záhlavím, číslováním a podpisovým blokem.',
      pdfSelectLang: 'Vyberte jazykovou mutaci pro PDF:',
      pdfPrintBtn: '🖨️ Otevřít tisk / Uložit do PDF',
      pdfDownloadHtmlBtn: '📥 Stáhnout tiskové HTML',
      pdfAllBtn: '📑 Hromadný tisk všech jazyků',

      // Modals
      diagTitle: '🧪 Interní diagnostika a testy správnosti',
      diagDesc: 'Ověřuje správnost klasifikace řádků (beze změny, přidáno, odstraněno, upraveno), textových diffů, bezpečnosti HTML, generování stubů a exportů přímo v tomto prohlížeči.',
      diagRerunBtn: 'Spustit testy znovu',
      promptSettingsTitle: '⚙️ Nastavení šablony promptu pro Copilot',
      promptSettingsDesc: 'Můžete upravit instrukce pro jazykový model Copilot (např. upřesnit firemní terminologii, tón komunikace či pravidla).',
      promptResetBtn: 'Obnovit výchozí',
      promptSaveBtn: 'Uložit šablonu',

      // Productivity & Accelerator Features
      saveStatusSaved: 'Vše uloženo',
      saveStatusSaving: 'Ukládám...',
      previewTitle: '👁️ Náhled prvních řádků načtené tabulky',
      previewDesc: 'Rychlá vizuální kontrola struktury sešitu a správnosti záhlaví.',
      matrixTitle: '📊 Přehled postupu všech jazyků',
      matrixApproveAllGlobal: '⚡ Schválit nezměněné ve všech jazycích',
      btnApproveUnchanged: '⚡ Schválit nezměněné',
      btnBatchPrompt: '📦 Hromadný prompt pro změněné',
      btnSmartPaste: '📥 Hromadné vložení (Smart Paste)',
      toastApproveUnchanged: 'Všechny nezměněné odstavce pro tento jazyk byly schváleny.',
      toastApproveGlobal: 'Všechny nezměněné odstavce ve všech jazycích byly schváleny.',
      ctrlEnterHint: 'Tip: V editačním poli můžete stisknout Ctrl + Enter pro schválení a skok na další.',

      // Smart Paste Modal
      smartPasteTitle: '📥 Hromadné vložení překladů od Copilota (Smart Paste)',
      smartPasteDesc: 'Vložte text odpovědi z Copilota. Systém automaticky detekuje čísla článků ([Art. 04], Art. 04: atd.) a přiřadí je.',
      smartPastePlaceholder: 'Vložte sem celou odpověď od Copilota...\n\nNapř:\n[Art. 04]\nUpravený překlad odstavce...\n\n[Art. 08]\nPřeklad nového článku...',
      smartPasteParseBtn: '🔍 Analyzovat a spárovat články',
      smartPasteApplyBtn: '✓ Aplikovat na vybrané články',
      smartPasteAutoVerify: 'Zároveň označit vložené články jako ověřené / schválené',
      smartPasteMatchesFound: 'Nalezeno {COUNT} odpovídajících článků k vložení',
      smartPasteNoMatches: 'V zadaném textu nebyly nalezeny žádné odpovídající značky článků.',

      // Batch Prompt Modal
      batchPromptTitle: '📦 Hromadný prompt pro Copilot',
      batchPromptDesc: 'Prompt obsahuje všechny upravené a nově přidané odstavce pro vybraný jazyk. Zkopírujte jej do Copilota.',
      batchPromptCopyBtn: '📋 Zkopírovat hromadný prompt do schránky',

      // PDF Template Modal
      editPdfTemplateBtn: '⚙️ Upravit šablonu & hlavičku PDF',
      pdfTemplateTitle: '⚙️ Nastavení šablony a parametrů PDF dokumentu',
      pdfTemplateDesc: 'Přizpůsobte firemní hlavičku, datum účinnosti, barvy a podpisový blok pro PDF export.',
      lblCompanyName: 'Název společnosti / organizace:',
      lblDocTitle: 'Hlavní název dokumentu:',
      lblDocSubtitle: 'Podtitul dokumentu ({LANG} bude nahrazeno jazykem):',
      lblEffectiveDate: 'Datum účinnosti:',
      lblVersionLabel: 'Označení verze:',
      lblAccentColor: 'Hlavní barva akcentu (HEX):',
      lblPreamble: 'Úvodní prohlášení / preambule:',
      lblShowSignatures: 'Zobrazit podpisový blok na konci dokumentu',
      lblSignDept1: 'Podpisový blok 1 (např. Compliance):',
      lblSignDept2: 'Podpisový blok 2 (např. Vedení):',
      btnSavePdfTemplate: 'Uložit šablonu PDF',
      btnResetPdfTemplate: 'Obnovit výchozí šablonu',

      // Managerial Summary Export
      mgrReportTitle: '📋 Manažerské shrnutí změn (Change Log 2024 → 2026)',
      mgrReportDesc: 'Oficiální report pro management, auditory a právní oddělení zachycující veškeré úpravy, nově zavedené články a zrušená ustanovení.',
      btnExportMd: '📝 Stáhnout Markdown (.md)',
      btnExportDoc: '📄 Stáhnout Word (.doc)',
      btnExportPdf: '🖨️ Vytisknout / Uložit PDF',

      // Toast & Alerts
      toastPromptCopied: 'Prompt pro Copilot byl zkopírován do schránky!',
      toastFileLoaded: 'Soubor "{NAME}" byl úspěšně načten.',
      toastSampleLoaded: 'Ukázkový kodex (15 článků, 12 jazyků) byl načten!',
      toastSampleDownloaded: 'Vzorový soubor byl stažen do vašeho počítače.',
      toastProjectSaved: 'Záložní soubor projektu byl exportován.',
      toastProjectLoaded: 'Projekt byl úspěšně načten ze zálohy!',
      toastMappingSaved: 'Mapování sloupců bylo uloženo.',
      toastAutoDetectSuccess: 'Sloupce byly automaticky detekovány!',
      toastAllApproved: 'Všechny zobrazené články jsou již schválené!',
      toastWorkbookGenerated: 'Sešit "{NAME}" byl úspěšně vygenerován!',
      alertSelectRefCols: 'Prosím vyberte oba referenční sloupce (starou i novou verzi).',
      alertDifferentRefCols: 'Starý a nový referenční sloupec musí být rozdílné.',
      alertSelectTargetLangs: 'Prosím označte alespoň jeden cílový jazyk k překladu.',
      alertConfirmReset: 'Opravdu chcete vyresetovat celý rozpracovaný projekt a začít znovu?',

      // Messages used by the application controller
      saveStatusError: 'Neuloženo – exportujte projekt (.json)',
      toastSaveFailed: 'Automatické uložení do prohlížeče selhalo (pravděpodobně plné úložiště). Zálohujte práci exportem projektu (.json).',
      toastBulkApproved: 'Schváleno {COUNT} nezměněných odstavců.',
      batchPromptCopied: 'Hromadný prompt byl zkopírován do schránky!',
      smartPasteParsedCount: 'Rozpoznané články',
      smartPasteApplied: 'Vloženo {COUNT} přeložených článků.',
      smartPasteUnmatched: 'Nespárované značky (nebudou vloženy): {LIST}',
      mgrExportMdSuccess: 'Manažerské shrnutí bylo staženo jako Markdown.',
      mgrExportDocSuccess: 'Manažerské shrnutí bylo staženo jako dokument Word.',
      pdfTemplateSaved: 'Šablona PDF byla uložena.',
      pdfTemplateReset: 'Šablona PDF byla obnovena na výchozí hodnoty.',
      toastNoData: 'Nejsou načtena žádná data k exportu.',
      toastSelectTargetLang: 'Vyberte cílový jazyk.',
      toastExportUnverified: 'Pozor: {COUNT} odstavců (součet za všechny jazyky) zatím není ověřeno nebo chybí jejich překlad – viz list „Stav_prekladu“.',
      toastExportUnverifiedClean: 'Pozor: {COUNT} odstavců (součet za všechny jazyky) zatím není ověřeno nebo chybí jejich překlad – čistá verze proto ještě není finální.',
      toastPdfSelectLang: 'Prosím vyberte jazyk pro export do PDF.',
      toastPdfHtmlDownloaded: 'Formátovaný tiskový HTML soubor byl stažen.',
      toastNoLanguages: 'Žádné jazyky k exportu.',
      confirmPrintAll: 'Chcete vygenerovat tisková okna pro všech {COUNT} jazykových mutací?',
      confirmResumeProject: 'V paměti prohlížeče byl nalezen rozpracovaný projekt z {DATE}. Chcete v něm pokračovat?',
      alertNoSheets: 'Nahraný soubor neobsahuje žádné listy.',
      alertEmptySheet: 'Vybraný list je prázdný.',
      alertFileReadError: 'Soubor se nepodařilo načíst: {MSG}',
      alertProjectLoadError: 'Projekt se nepodařilo načíst. {MSG}',
      rowsLabel: 'řádků',
      rowLabel: 'Řádek',
      progressDone: 'hotovo',
      noTargetLang: 'Žádný cílový jazyk k zobrazení.',
      noArticlesMatch: 'Žádné odstavce neodpovídají zadanému filtru nebo vyhledávání.',
      pdfOptionRefNew: 'Referenční němčina 2026',
      charsLabel: 'znaků'
    },

    en: {
      // Header
      appTitle: 'Code of Ethics – Comparator & Translation Assistant',
      appSubtitle: 'Synchronization of German reference versions (2024 → 2026) and multilingual translations',
      securityBadge: '🔒 100% Offline & Secure',
      securityTooltip: 'This application runs entirely locally in your browser. No data is sent to the internet or external APIs.',
      diagnosticsBtn: '🧪 Diagnostics & Tests',
      promptSettingsBtn: '⚙️ Prompt Template',
      resetProjectBtn: '↺ New Project',
      themeDark: '🌙 Dark Mode',
      themeLight: '☀️ Light Mode',
      langToggle: '🇨🇿 Čeština',

      // Stepper
      step1Nav: '1. File & Column Mapping',
      step2Nav: '2. Reference Changes (DE)',
      step3Nav: '3. Translation Workspace',
      step4Nav: '4. Export & Download',

      // Step 1
      step1Title: 'Step 1: Upload Excel Table & Map Columns',
      step1Desc: 'Upload a workbook containing the reference versions (DE 2024 & DE 2026) and existing language translations.',
      dropzoneTitle: 'Drag & drop Excel file (.xlsx, .xls) here or click to browse',
      dropzoneSubtitle: 'Files are processed exclusively in your browser without any server transmission.',
      loadSampleBtn: '🚀 Try with Sample Code of Conduct (15 articles, 12 languages)',
      downloadSampleBtn: '📥 Download Sample File (Demo XLSX)',
      importProjectBtn: '📂 Open Existing Project (.json)',
      sheetSelectLabel: 'Select workbook sheet to analyze:',
      mappingTitle: 'Designate Reference Columns and Target Languages',
      autoDetectBtn: '🔍 Auto-Detect Columns',
      refOldLabel: '1. Reference Column 2024 (Original German):',
      refOldHelp: 'Used as baseline to compare against legacy translations.',
      refNewLabel: '2. Reference Column 2026 (Updated German):',
      refNewHelp: 'Used to identify modifications, additions, and deletions.',
      targetColsLabel: '3. Target Translation Languages (2024 Version):',
      targetColsHelp: 'Check all columns with approved legacy translations for which you want to build the 2026 version.',
      confirmMappingBtn: 'Proceed to Reference Changes Overview →',

      // Step 2
      step2Title: 'Step 2: Reference Changes Overview (German 2024 vs 2026)',
      step2Desc: 'Automated paragraph categorization: unchanged, modified, newly added, and removed.',
      goToTranslationBtn: 'Open Translation Workspace →',
      kpiTotal: 'Total Paragraphs',
      kpiUnchanged: 'Unchanged',
      kpiModified: 'Modified',
      kpiAdded: '+ Newly Added',
      kpiRemoved: '- Removed',
      thArticle: 'Article / ID',
      thClassification: 'Status',
      thDiff: 'German Text & Highlighted Word Diff (2024 → 2026)',

      // Step 3
      step3Title: 'Step 3: Translation Workspace & Visual Diff Verification',
      step3Desc: 'Unchanged sections are auto-copied. For modified rows, copy the prompt for Copilot and paste the result with live diff validation.',
      goToExportBtn: 'Proceed to Export →',
      progressLabel: 'Review & approval progress for active language:',
      filterAll: 'All Paragraphs',
      filterPending: '⏳ Pending Review',
      filterVerified: '✓ Approved',
      filterModified: '✏️ Modified Only',
      filterAdded: '➕ Added Only',
      filterRemoved: '🗑️ Removed Only',
      searchPlaceholder: 'Search text across paragraphs...',
      prevPendingBtn: '▲ Previous Pending',
      nextPendingBtn: '▼ Next Pending',
      layout4Cols: '▤ 4-Column Side-by-Side (Full Width)',
      layout2x2: '▦ 2×2 Grid',

      // Article Card
      cardRefPanelTitle: '1. German Reference (DE 2024 → 2026 diff)',
      cardTargetOldTitle: '2. Approved Original Translation ({LANG} 2024)',
      cardTargetNewTitle: '3. Updated 2026 Translation ({LANG})',
      cardTargetDiffTitle: '4. Translation Diff ({LANG} 2024 → 2026)',
      cardCompareSummary: 'Visual Comparison (German Reference vs Target Translation)',
      cardCompareHint: 'Verify that changes in target language correspond accurately to the reference diff',
      copyPromptBtn: '📋 Copy Prompt for Copilot',
      prefillOldBtn: '↺ Prefill from 2024',
      prefillRefBtn: '↺ Insert German Text',
      markVerified: '✓ Approved / Done',
      markToVerify: 'Mark as Approved',
      textareaPlaceholder: 'Paste or type updated 2026 translation...',
      emptyOriginalText: '(Did not exist in 2024)',
      noChangesTargetDiff: 'No changes from 2024 yet',
      targetNotTranslatedYet: 'No new translation yet – the difference appears as you type.',
      emptyCell: '(Empty)',
      match100: '100% match',
      changePct: 'Change: {PCT}%',

      // Badges
      badgeUnchanged: 'Unchanged',
      badgeModified: 'Modified',
      badgeAdded: '+ Added',
      badgeRemoved: '- Removed',
      badgeVerified: 'Approved',

      // Step 4
      step4Title: 'Step 4: Export Updated Excel Workbook & Backup Project',
      step4Desc: 'Select your preferred output format – a comprehensive comparison table with all columns or a clean 2026 final publication.',
      comprehensiveTitle: 'Comprehensive Comparison Workbook (2024 + 2026)',
      comprehensiveDesc: 'Preserves all original columns and rows, inserting a new 2026 column next to each 2024 language. Ideal for audit trails and internal approvals.',
      comprehensiveBtn: '📥 Download Full Workbook (.xlsx)',
      cleanTitle: 'Clean 2026 Version (Final Publication)',
      cleanDesc: 'Generates a clean table containing only valid 2026 articles in all languages. Removed articles are automatically omitted.',
      cleanBtn: '📥 Download Clean 2026 (.xlsx)',
      backupTitle: 'Project Backup File (.json)',
      backupDesc: 'Saves the complete state of your work including all draft translations and column mappings into a single file to resume anytime.',
      backupBtn: '💾 Export Project (.json)',
      pdfTitle: 'Formatted PDF Documents (Per Language)',
      pdfDesc: 'Generates an elegant, professionally formatted A4 print-ready Code of Conduct document with header, article tags, and execution block.',
      pdfSelectLang: 'Select language edition for PDF:',
      pdfPrintBtn: '🖨️ Open Print / Save as PDF',
      pdfDownloadHtmlBtn: '📥 Download Printable HTML',
      pdfAllBtn: '📑 Batch Print All Languages',

      // Modals
      diagTitle: '🧪 Internal Diagnostics & Accuracy Tests',
      diagDesc: 'Verifies paragraph classification (unchanged, added, removed, modified), word diffs, HTML sanitization, stub generation, and Excel exports directly in this browser.',
      diagRerunBtn: 'Rerun Tests',
      promptSettingsTitle: '⚙️ Copilot Prompt Template Settings',
      promptSettingsDesc: 'Customize the instructions sent to the Copilot LLM (e.g. company terminology, tone of voice, or legal constraints).',
      promptResetBtn: 'Reset to Default',
      promptSaveBtn: 'Save Template',

      // Productivity & Accelerator Features
      saveStatusSaved: 'All changes saved',
      saveStatusSaving: 'Saving...',
      previewTitle: '👁️ Spreadsheet First Rows Preview',
      previewDesc: 'Quick visual check of workbook structure and column headers.',
      matrixTitle: '📊 All Languages Progress Overview',
      matrixApproveAllGlobal: '⚡ Approve unchanged across all languages',
      btnApproveUnchanged: '⚡ Approve Unchanged',
      btnBatchPrompt: '📦 Batch Prompt for Modified',
      btnSmartPaste: '📥 Smart Batch Paste',
      toastApproveUnchanged: 'All unchanged articles for this language marked as approved.',
      toastApproveGlobal: 'All unchanged articles across all languages approved.',
      ctrlEnterHint: 'Tip: Press Ctrl + Enter inside any editor to approve and jump to the next item.',

      // Smart Paste Modal
      smartPasteTitle: '📥 Smart Batch Translation Paste',
      smartPasteDesc: 'Paste Copilot\'s response. The system will automatically detect article IDs ([Art. 04], Art. 04: etc.) and map them.',
      smartPastePlaceholder: 'Paste full Copilot reply here...\n\nExample:\n[Art. 04]\nUpdated translated text...\n\n[Art. 08]\nNewly translated text...',
      smartPasteParseBtn: '🔍 Analyze and Match Articles',
      smartPasteApplyBtn: '✓ Apply to Matched Articles',
      smartPasteAutoVerify: 'Also mark applied articles as verified / approved',
      smartPasteMatchesFound: 'Found {COUNT} matching articles to apply',
      smartPasteNoMatches: 'No matching article tags found in pasted text.',

      // Batch Prompt Modal
      batchPromptTitle: '📦 Batch Prompt for Copilot',
      batchPromptDesc: 'Contains all modified and newly added articles for the active language. Copy this prompt to Copilot.',
      batchPromptCopyBtn: '📋 Copy Batch Prompt to Clipboard',

      // PDF Template Modal
      editPdfTemplateBtn: '⚙️ Edit PDF Template & Branding',
      pdfTemplateTitle: '⚙️ PDF Template & Branding Settings',
      pdfTemplateDesc: 'Customize corporate header, effective date, colors, and signature blocks for PDF export.',
      lblCompanyName: 'Company / Organization Name:',
      lblDocTitle: 'Main Document Title:',
      lblDocSubtitle: 'Document Subtitle ({LANG} will be replaced by language):',
      lblEffectiveDate: 'Effective Date:',
      lblVersionLabel: 'Version Label:',
      lblAccentColor: 'Primary Accent Color (HEX):',
      lblPreamble: 'Introductory Statement / Preamble:',
      lblShowSignatures: 'Show signature block at the end of document',
      lblSignDept1: 'Signature Block 1 (e.g. Compliance):',
      lblSignDept2: 'Signature Block 2 (e.g. Executive Board):',
      btnSavePdfTemplate: 'Save PDF Template',
      btnResetPdfTemplate: 'Reset to Default Template',

      // Managerial Summary Export
      mgrReportTitle: '📋 Management Change Summary (Change Log 2024 → 2026)',
      mgrReportDesc: 'Official audit and executive report capturing all modifications, additions, and repealed articles.',
      btnExportMd: '📝 Download Markdown (.md)',
      btnExportDoc: '📄 Download Word (.doc)',
      btnExportPdf: '🖨️ Print / Save as PDF',

      // Toast & Alerts
      toastPromptCopied: 'Copilot prompt copied to clipboard!',
      toastFileLoaded: 'File "{NAME}" was successfully loaded.',
      toastSampleLoaded: 'Sample code (15 articles, 12 languages) loaded!',
      toastSampleDownloaded: 'Sample file downloaded to your computer.',
      toastProjectSaved: 'Project backup file was exported.',
      toastProjectLoaded: 'Project successfully loaded from backup!',
      toastMappingSaved: 'Column mappings saved.',
      toastAutoDetectSuccess: 'Columns auto-detected successfully!',
      toastAllApproved: 'All displayed articles are already approved!',
      toastWorkbookGenerated: 'Workbook "{NAME}" generated successfully!',
      alertSelectRefCols: 'Please select both reference columns (2024 and 2026 versions).',
      alertDifferentRefCols: 'The old and new reference columns must be different.',
      alertSelectTargetLangs: 'Please select at least one target language for translation.',
      alertConfirmReset: 'Are you sure you want to reset the entire project and start over?',

      // Messages used by the application controller
      saveStatusError: 'Not saved – export the project (.json)',
      toastSaveFailed: 'Autosave to the browser failed (storage is probably full). Back up your work by exporting the project (.json).',
      toastBulkApproved: '{COUNT} unchanged articles approved.',
      batchPromptCopied: 'Batch prompt copied to clipboard!',
      smartPasteParsedCount: 'Recognized articles',
      smartPasteApplied: '{COUNT} translated articles applied.',
      smartPasteUnmatched: 'Unmatched tags (will not be applied): {LIST}',
      mgrExportMdSuccess: 'Management summary downloaded as Markdown.',
      mgrExportDocSuccess: 'Management summary downloaded as a Word document.',
      pdfTemplateSaved: 'PDF template saved.',
      pdfTemplateReset: 'PDF template reset to defaults.',
      toastNoData: 'No data loaded to export.',
      toastSelectTargetLang: 'Select a target language.',
      toastExportUnverified: 'Note: {COUNT} articles (total over all languages) are not verified yet or have no translation – see the "Translation_Status" sheet.',
      toastExportUnverifiedClean: 'Note: {COUNT} articles (total over all languages) are not verified yet or have no translation – the clean version is not final yet.',
      toastPdfSelectLang: 'Please select a language for PDF export.',
      toastPdfHtmlDownloaded: 'Printable HTML document downloaded.',
      toastNoLanguages: 'No languages available to export.',
      confirmPrintAll: 'Do you want to generate print windows for all {COUNT} language editions?',
      confirmResumeProject: 'A saved project from {DATE} was found. Do you want to resume it?',
      alertNoSheets: 'The uploaded file contains no sheets.',
      alertEmptySheet: 'The selected sheet is empty.',
      alertFileReadError: 'The file could not be read: {MSG}',
      alertProjectLoadError: 'The project could not be loaded. {MSG}',
      rowsLabel: 'rows',
      rowLabel: 'Row',
      progressDone: 'completed',
      noTargetLang: 'No target language selected.',
      noArticlesMatch: 'No articles match the current filter or search criteria.',
      pdfOptionRefNew: 'German 2026 Reference',
      charsLabel: 'chars'
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
