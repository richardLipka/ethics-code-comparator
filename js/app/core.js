/**
 * Shared state, DOM references, translations, notifications, autosave and helpers used by every step.
 * Part of the application controller; the parts share the `App` namespace and are loaded in the
 * order given in index.html (js/app/core.js first, js/app/main.js last).
 */
(function (App) {
  'use strict';

  // Application State
  const state = {
    currentStep: 1, // 1: Upload & Map, 2: Ref Overview, 3: Translation Studio, 4: Export
    importSource: null, // source returned by the import adapter (see js/io/importers.js)
    fileName: '',
    sheetNames: [],
    activeSheet: '',
    hasHeaderRow: true,
    headers: [],
    rows: [],
    rowNumbers: [], // Excel row number of every data row
    skippedRows: [], // rows excluded from processing: [{ rowNumber, row, auto }] (auto = empty row)
    previewLimit: 10, // rows shown in the Step 1 preview
    mappings: {
      refOldCol: '',
      refNewCol: '',
      targetCols: [], // old-version columns of the languages to update
      languages: {} // { [oldCol]: { label, newCol, newColName } }
    },
    exportLangCol: '',
    partialContextCols: null,
    tableExportFormat: '', // id of the table export adapter ('' = default)
    // Translations map: { [targetCol]: [ { rowIndex, targetNew, status, isVerified, notes } ] }
    translations: {},
    activeTargetCol: '',
    auxiliaryCol: '',
    filterStatus: 'all', // 'all', 'pending', 'verified', 'modified', 'added', 'removed'
    searchQuery: '',
    layoutMode: '4cols', // '4cols' | '2x2'
    customPromptTemplate: null,
    pdfTemplate: null, // Custom PDF branding and parameters
    theme: 'light' // Light theme default
  };

  // DOM Elements
  const elements = {
    langToggleBtn: document.getElementById('langToggleBtn'),
    themeToggleBtn: document.getElementById('themeToggleBtn'),
    diagnosticsBtn: document.getElementById('diagnosticsBtn'),
    settingsPromptBtn: document.getElementById('settingsPromptBtn'),
    resetProjectBtn: document.getElementById('resetProjectBtn'),
    loadSampleBtn: document.getElementById('loadSampleBtn'),
    downloadSampleBtn: document.getElementById('downloadSampleBtn'),
    
    // Auto-save Indicator
    saveIndicator: document.getElementById('saveIndicator'),
    saveStatusText: document.getElementById('saveStatusText'),

    // Dropzone & File Input
    dropzone: document.getElementById('dropzone'),
    fileInput: document.getElementById('fileInput'),
    projectFileInput: document.getElementById('projectFileInput'),
    saveProjectBtn: document.getElementById('saveProjectBtn'),
    loadProjectBtn: document.getElementById('loadProjectBtn'),
    
    // Steps Containers
    step1Section: document.getElementById('step1Section'),
    step2Section: document.getElementById('step2Section'),
    step3Section: document.getElementById('step3Section'),
    step4Section: document.getElementById('step4Section'),
    
    // Stepper Nav Items
    stepNav1: document.getElementById('stepNav1'),
    stepNav2: document.getElementById('stepNav2'),
    stepNav3: document.getElementById('stepNav3'),
    stepNav4: document.getElementById('stepNav4'),
    stepDivider1: document.getElementById('stepDivider1'),
    stepDivider2: document.getElementById('stepDivider2'),
    stepDivider3: document.getElementById('stepDivider3'),
    stepNumber2: document.getElementById('stepNumber2'),
    stepNumber3: document.getElementById('stepNumber3'),
    stepNumber4: document.getElementById('stepNumber4'),
    
    // Step 1 Controls
    sheetOptionsWrapper: document.getElementById('sheetOptionsWrapper'),
    sheetSelectWrapper: document.getElementById('sheetSelectWrapper'),
    sheetSelect: document.getElementById('sheetSelect'),
    hasHeaderRowCheckbox: document.getElementById('hasHeaderRowCheckbox'),
    mappingContainer: document.getElementById('mappingContainer'),
    refOldColCard: document.getElementById('refOldColCard'),
    refNewColCard: document.getElementById('refNewColCard'),
    refNewColLabel: document.getElementById('refNewColLabel'),
    refNewColHelp: document.getElementById('refNewColHelp'),
    directTransInfoBox: document.getElementById('directTransInfoBox'),
    targetColsCard: document.getElementById('targetColsCard'),
    targetColsTitle: document.getElementById('targetColsTitle'),
    targetColsHelp: document.getElementById('targetColsHelp'),
    quickLangsWrapper: document.getElementById('quickLangsWrapper'),
    quickLangsContainer: document.getElementById('quickLangsContainer'),
    selectedDirectLangsList: document.getElementById('selectedDirectLangsList'),
    languagesTableWrapper: document.getElementById('languagesTableWrapper'),
    refOldSelect: document.getElementById('refOldSelect'),
    refNewSelect: document.getElementById('refNewSelect'),
    languagesTableBody: document.getElementById('languagesTableBody'),
    newLangNameInput: document.getElementById('newLangNameInput'),
    newLangColNameInput: document.getElementById('newLangColNameInput'),
    addNewLangBtn: document.getElementById('addNewLangBtn'),
    confirmMappingBtn: document.getElementById('confirmMappingBtn'),
    dataPreviewContainer: document.getElementById('dataPreviewContainer'),
    previewTableHead: document.getElementById('previewTableHead'),
    previewTableBody: document.getElementById('previewTableBody'),
    previewRowsCount: document.getElementById('previewRowsCount'),
    previewShowMoreBtn: document.getElementById('previewShowMoreBtn'),
    backToSettingsBtn: document.getElementById('backToSettingsBtn'),

    // Step 2 Controls
    kpiTotalRows: document.getElementById('kpiTotalRows'),
    kpiUnchanged: document.getElementById('kpiUnchanged'),
    kpiModified: document.getElementById('kpiModified'),
    kpiAdded: document.getElementById('kpiAdded'),
    kpiRemoved: document.getElementById('kpiRemoved'),
    refDiffTableBody: document.getElementById('refDiffTableBody'),
    goToTranslationBtn: document.getElementById('goToTranslationBtn'),
    step2ExportMdBtn: document.getElementById('step2ExportMdBtn'),
    step2ExportDocBtn: document.getElementById('step2ExportDocBtn'),
    step2ExportPdfBtn: document.getElementById('step2ExportPdfBtn'),
    
    // Step 3 Controls
    crossLangMatrixCard: document.getElementById('crossLangMatrixCard'),
    crossLangMatrixGrid: document.getElementById('crossLangMatrixGrid'),
    approveAllUnchangedGlobalBtn: document.getElementById('approveAllUnchangedGlobalBtn'),
    langTabsWrapper: document.getElementById('langTabsWrapper'),
    searchInput: document.getElementById('searchInput'),
    filterChipsContainer: document.getElementById('filterChipsContainer'),
    approveUnchangedBtn: document.getElementById('approveUnchangedBtn'),
    batchPromptBtn: document.getElementById('batchPromptBtn'),
    smartPasteBtn: document.getElementById('smartPasteBtn'),
    importTranslationsBtn: document.getElementById('importTranslationsBtn'),
    importTranslationsExportBtn: document.getElementById('importTranslationsExportBtn'),
    translationsFileInput: document.getElementById('translationsFileInput'),
    layoutToggleBtn: document.getElementById('layoutToggleBtn'),
    auxLangSelect: document.getElementById('auxLangSelect'),
    articlesContainer: document.getElementById('articlesContainer'),
    currentLangProgressText: document.getElementById('currentLangProgressText'),
    currentLangProgressBar: document.getElementById('currentLangProgressBar'),
    prevPendingBtn: document.getElementById('prevPendingBtn'),
    nextPendingBtn: document.getElementById('nextPendingBtn'),
    goToExportBtn: document.getElementById('goToExportBtn'),

    // Step 4 Controls
    downloadComprehensiveBtn: document.getElementById('downloadComprehensiveBtn'),
    downloadCleanBtn: document.getElementById('downloadCleanBtn'),
    downloadPartialBtn: document.getElementById('downloadPartialBtn'),
    tableFormatWrapper: document.getElementById('tableFormatWrapper'),
    tableFormatSelect: document.getElementById('tableFormatSelect'),
    partialContextContainer: document.getElementById('partialContextContainer'),
    exportProjectJsonBtn: document.getElementById('exportProjectJsonBtn'),
    pdfLangSelect: document.getElementById('pdfLangSelect'),
    printPdfBtn: document.getElementById('printPdfBtn'),
    downloadPdfHtmlBtn: document.getElementById('downloadPdfHtmlBtn'),
    printAllPdfBtn: document.getElementById('printAllPdfBtn'),
    editPdfTemplateBtn: document.getElementById('editPdfTemplateBtn'),
    downloadMgrMdBtn: document.getElementById('downloadMgrMdBtn'),
    downloadMgrDocBtn: document.getElementById('downloadMgrDocBtn'),
    printMgrPdfBtn: document.getElementById('printMgrPdfBtn'),
    
    // Modals
    diagnosticsModal: document.getElementById('diagnosticsModal'),
    diagnosticsLog: document.getElementById('diagnosticsLog'),
    closeDiagnosticsBtn: document.getElementById('closeDiagnosticsBtn'),
    runDiagnosticsBtn: document.getElementById('runDiagnosticsBtn'),
    
    promptSettingsModal: document.getElementById('promptSettingsModal'),
    promptTemplateInput: document.getElementById('promptTemplateInput'),
    savePromptTemplateBtn: document.getElementById('savePromptTemplateBtn'),
    resetPromptTemplateBtn: document.getElementById('resetPromptTemplateBtn'),
    closePromptSettingsBtn: document.getElementById('closePromptSettingsBtn'),

    batchPromptModal: document.getElementById('batchPromptModal'),
    batchPromptTextarea: document.getElementById('batchPromptTextarea'),
    closeBatchPromptBtn: document.getElementById('closeBatchPromptBtn'),
    closeBatchPromptFooterBtn: document.getElementById('closeBatchPromptFooterBtn'),
    copyBatchPromptBtn: document.getElementById('copyBatchPromptBtn'),

    smartPasteModal: document.getElementById('smartPasteModal'),
    smartPasteInput: document.getElementById('smartPasteInput'),
    smartPasteParseBtn: document.getElementById('smartPasteParseBtn'),
    smartPasteAutoVerifyCheckbox: document.getElementById('smartPasteAutoVerifyCheckbox'),
    smartPasteMatchesContainer: document.getElementById('smartPasteMatchesContainer'),
    smartPasteMatchesTitle: document.getElementById('smartPasteMatchesTitle'),
    smartPasteMatchesList: document.getElementById('smartPasteMatchesList'),
    closeSmartPasteBtn: document.getElementById('closeSmartPasteBtn'),
    closeSmartPasteFooterBtn: document.getElementById('closeSmartPasteFooterBtn'),
    smartPasteApplyBtn: document.getElementById('smartPasteApplyBtn'),

    pdfTemplateModal: document.getElementById('pdfTemplateModal'),
    tplCompanyName: document.getElementById('tplCompanyName'),
    tplEffectiveDate: document.getElementById('tplEffectiveDate'),
    tplDocTitle: document.getElementById('tplDocTitle'),
    tplAccentColorPicker: document.getElementById('tplAccentColorPicker'),
    tplAccentColor: document.getElementById('tplAccentColor'),
    tplDocSubtitle: document.getElementById('tplDocSubtitle'),
    tplVersionLabel: document.getElementById('tplVersionLabel'),
    tplIntroPreamble: document.getElementById('tplIntroPreamble'),
    tplShowSignatures: document.getElementById('tplShowSignatures'),
    signaturesBlockInputs: document.getElementById('signaturesBlockInputs'),
    tplSignDept1: document.getElementById('tplSignDept1'),
    tplSignDept2: document.getElementById('tplSignDept2'),
    savePdfTemplateBtn: document.getElementById('savePdfTemplateBtn'),
    resetPdfTemplateBtn: document.getElementById('resetPdfTemplateBtn'),
    closePdfTemplateBtn: document.getElementById('closePdfTemplateBtn'),
    
    toastContainer: document.getElementById('toastContainer')
  };

  /**
   * Helper to translate key with optional params
   */
  function t(key, params) {
    return I18n.t(key, params);
  }

  /**
   * Show notification toast
   */
  function showToast(message, type) {
    const toast = document.createElement('div');
    toast.className = 'toast toast-' + (type || 'info');
    const icon = type === 'success' ? '✓' : (type === 'warning' ? '⚠' : 'ℹ');
    toast.innerHTML = '<span>' + icon + '</span> ' + ComparatorEngine.escapeHtml(message);
    if (type === 'warning') toast.setAttribute('role', 'alert');
    elements.toastContainer.appendChild(toast);

    setTimeout(function () {
      toast.style.opacity = '0';
      toast.style.transform = 'translateY(10px)';
      setTimeout(function () {
        if (toast.parentNode) toast.parentNode.removeChild(toast);
      }, 300);
    }, type === 'warning' ? 7000 : 3200);
  }

  /**
   * Persist state to storage and update header save indicator
   */
  let saveFailedNotified = false;

  function persistState(debounceMs) {
    if (elements.saveIndicator && elements.saveStatusText) {
      elements.saveIndicator.classList.add('saving');
      elements.saveStatusText.textContent = t('saveStatusSaving');
    }
    const delay = debounceMs !== undefined ? debounceMs : 300;
    StorageManager.debouncedSaveState(state, delay, updateSaveIndicator);
  }

  function updateSaveIndicator(ok) {
    if (elements.saveIndicator && elements.saveStatusText) {
      elements.saveIndicator.classList.remove('saving');
      elements.saveIndicator.classList.toggle('save-error', !ok);
      elements.saveStatusText.textContent = ok ? t('saveStatusSaved') : t('saveStatusError');
    }
    if (!ok && !saveFailedNotified) {
      saveFailedNotified = true;
      showToast(t('toastSaveFailed'), 'warning');
    } else if (ok) {
      saveFailedNotified = false;
    }
  }

  /**
   * LocalStorage helpers for UI preferences (storage may be blocked, e.g. in private mode).
   */
  function readPref(key, fallback) {
    try {
      return window.localStorage.getItem(key) || fallback;
    } catch (e) {
      return fallback;
    }
  }

  function writePref(key, value) {
    try {
      window.localStorage.setItem(key, value);
    } catch (e) {
      // Preferences are a convenience only
    }
  }

  // A table with a single column is translated directly (no old reference, Step 2 is skipped)
  function isDirectTranslationMode() {
    return !!(state.headers && state.headers.length === 1);
  }

  /**
   * Returns the i18n key of the first mapping problem, or '' when the mapping is usable.
   */
  function getMappingError(requireTargets) {
    const m = state.mappings;
    if (!m.refNewCol) return 'alertSelectRefCols';
    if (!isDirectTranslationMode() && m.refOldCol && m.refOldCol === m.refNewCol) return 'alertDifferentRefCols';
    if (requireTargets && (!m.targetCols || m.targetCols.length === 0)) {
      return isDirectTranslationMode() ? 'alertSelectTargetLangsDirect' : 'alertSelectTargetLangs';
    }
    return '';
  }

  /**
   * Human-readable column name: the header, or "Column C" when the sheet has no header row
   * or the header cell is empty.
   */
  function columnDisplayName(col) {
    if (ComparatorEngine.isNewLanguage(state.mappings, col)) {
      const cfg = state.mappings.languages && state.mappings.languages[col];
      return (cfg && (cfg.label || cfg.newColName)) || String(col).replace(/^__new__/, '');
    }
    if (col === 'Reference' && isDirectTranslationMode()) {
      return t('sourceTextColumn');
    }
    if (!state.hasHeaderRow || /^__EMPTY(_\d+)?$/.test(col)) {
      const letter = (state.columnLetters && state.columnLetters[col]) || (state.hasHeaderRow ? '?' : col);
      return t('columnLetterLabel', { COL: letter });
    }
    return String(col);
  }

  function rowNumberOf(idx) {
    return ComparatorEngine.rowNumberOf(state.rowNumbers, idx);
  }

  function rowLabelOf(idx) {
    return t('rowLabel') + ' ' + rowNumberOf(idx);
  }

  // Rows shown in the Step 1 data preview, initially and per "show more"
  const PREVIEW_INITIAL_ROWS = 10;
  const PREVIEW_MORE_ROWS = 20;

  function languageLabel(col) {
    const cfg = state.mappings.languages && state.mappings.languages[col];
    if (cfg && cfg.label) return cfg.label;
    if (ComparatorEngine.isNewLanguage(state.mappings, col)) {
      return (cfg && cfg.newColName) || String(col).replace(/^__new__/, '');
    }
    return ComparatorEngine.defaultLanguageLabel(columnDisplayName(col), state.mappings.refOldCol);
  }

  function defaultNewColNameFor(col) {
    return ComparatorEngine.defaultNewColName(col, state.mappings.refOldCol, state.mappings.refNewCol, t('newColSuffix'));
  }

  function fallbackCopyText(text, successToast) {
    const temp = document.createElement('textarea');
    temp.value = text;
    document.body.appendChild(temp);
    temp.select();
    document.execCommand('copy');
    document.body.removeChild(temp);
    showToast(successToast || t('toastPromptCopied'), 'success');
  }

  Object.assign(App, { state: state, elements: elements, t: t, showToast: showToast,
    persistState: persistState, readPref: readPref, writePref: writePref,
    isDirectTranslationMode: isDirectTranslationMode, getMappingError: getMappingError,
    columnDisplayName: columnDisplayName, rowNumberOf: rowNumberOf, rowLabelOf: rowLabelOf,
    PREVIEW_INITIAL_ROWS: PREVIEW_INITIAL_ROWS, PREVIEW_MORE_ROWS: PREVIEW_MORE_ROWS,
    languageLabel: languageLabel, defaultNewColNameFor: defaultNewColNameFor,
    fallbackCopyText: fallbackCopyText });
}(window.App = window.App || {}));
