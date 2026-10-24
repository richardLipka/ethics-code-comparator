/**
 * Application Controller - multi-language version comparator and translation assistant.
 * Bilingual (Czech & English), Full-width 4-column comparison layout, 100% offline.
 */

(function () {
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

  /**
   * Language Switcher (Czech & English)
   */
  function initLanguage() {
    I18n.initLanguage();
    applyLanguageToDOM();

    elements.langToggleBtn.addEventListener('click', function () {
      I18n.toggleLanguage();
      applyLanguageToDOM();
      
      // Re-render active screens to update dynamic labels
      if (state.currentStep === 1) {
        if (state.rows.length > 0) renderColumnMappingControls();
        renderDataPreview();
      } else if (state.currentStep === 2) {
        renderReferenceOverview();
      } else if (state.currentStep === 3) {
        renderTranslationStudio();
      } else if (state.currentStep === 4) {
        updatePdfLanguageDropdown();
      }

      showToast(I18n.getLanguage() === 'en' ? 'Language switched to English' : 'Jazyk byl přepnut do češtiny', 'info');
    });
  }

  function applyLanguageToDOM() {
    const lang = I18n.getLanguage();
    document.documentElement.setAttribute('lang', lang);

    // Update toggle button text
    elements.langToggleBtn.textContent = t('langToggle');

    // Update static elements with data-i18n
    document.querySelectorAll('[data-i18n]').forEach(function (el) {
      const key = el.getAttribute('data-i18n');
      if (key) {
        el.textContent = t(key);
      }
    });

    // Update elements with data-i18n-placeholder
    document.querySelectorAll('[data-i18n-placeholder]').forEach(function (el) {
      const key = el.getAttribute('data-i18n-placeholder');
      if (key) {
        el.setAttribute('placeholder', t(key));
      }
    });

    // Update elements with data-i18n-title
    document.querySelectorAll('[data-i18n-title]').forEach(function (el) {
      const key = el.getAttribute('data-i18n-title');
      if (key) {
        el.setAttribute('title', t(key));
      }
    });

    // Update theme toggle button text
    elements.themeToggleBtn.textContent = state.theme === 'dark' ? t('themeLight') : t('themeDark');

    // Update layout toggle text and auxiliary language options
    updateLayoutToggleText();
    updateAuxLangOptions();
  }

  /**
   * Theme Management (Light mode default)
   */
  function initTheme() {
    const savedTheme = readPref('comparator_theme', 'light');
    setTheme(savedTheme);

    elements.themeToggleBtn.addEventListener('click', function () {
      const current = document.documentElement.getAttribute('data-theme') || 'light';
      const next = current === 'dark' ? 'light' : 'dark';
      setTheme(next);
    });
  }

  function setTheme(theme) {
    state.theme = theme;
    document.documentElement.setAttribute('data-theme', theme);
    writePref('comparator_theme', theme);
    elements.themeToggleBtn.textContent = theme === 'dark' ? t('themeLight') : t('themeDark');
  }

  /**
   * Layout Mode Switcher (4-Columns Side-by-Side vs 2x2 Grid)
   */
  function updateLayoutToggleText() {
    if (elements.layoutToggleBtn) {
      if (state.layoutMode === '4cols') {
        elements.layoutToggleBtn.textContent = state.auxiliaryCol ? t('layout5Cols') : t('layout4Cols');
      } else {
        elements.layoutToggleBtn.textContent = t('layout2x2');
      }
    }
  }

  function initLayoutMode() {
    const saved = readPref('comparator_layout', '4cols');
    state.layoutMode = saved === '2x2' ? '2x2' : '4cols';
    updateLayoutToggleText();
    if (elements.layoutToggleBtn) {
      elements.layoutToggleBtn.addEventListener('click', function () {
        state.layoutMode = state.layoutMode === '4cols' ? '2x2' : '4cols';
        writePref('comparator_layout', state.layoutMode);
        updateLayoutToggleText();
        renderArticleCards();
      });
    }
  }

  /**
   * Stepper navigation
   */
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

  function setStep(stepNumber) {
    if (stepNumber > 1 && (!state.rows || state.rows.length === 0)) {
      showToast(t('toastNoData'), 'warning');
      return;
    }
    if (isDirectTranslationMode() && stepNumber === 2) {
      stepNumber = 3;
    }
    if (stepNumber > 1) {
      const mappingError = getMappingError(stepNumber >= 3);
      if (mappingError) {
        showToast(t(mappingError), 'warning');
        return;
      }
    }

    const stepChanged = state.currentStep !== stepNumber;
    state.currentStep = stepNumber;

    // Toggle active sections
    elements.step1Section.style.display = stepNumber === 1 ? 'block' : 'none';
    elements.step2Section.style.display = stepNumber === 2 ? 'block' : 'none';
    elements.step3Section.style.display = stepNumber === 3 ? 'block' : 'none';
    elements.step4Section.style.display = stepNumber === 4 ? 'block' : 'none';

    if (elements.backToSettingsBtn) {
      elements.backToSettingsBtn.style.display = (stepNumber === 3 && isDirectTranslationMode()) ? 'inline-flex' : 'none';
    }

    // Update nav classes
    [elements.stepNav1, elements.stepNav2, elements.stepNav3, elements.stepNav4].forEach(function (nav, idx) {
      if (idx + 1 === stepNumber) {
        nav.classList.add('active');
      } else {
        nav.classList.remove('active');
      }
    });

    if (stepNumber === 2) {
      renderReferenceOverview();
    } else if (stepNumber === 3) {
      ensureStubsInitialized();
      renderTranslationStudio();
    } else if (stepNumber === 4) {
      ensureStubsInitialized();
      updatePdfLanguageDropdown();
      renderTableFormatOptions();
    }

    // Staying in the same step keeps the position (e.g. the column selection revealed after loading a file)
    if (stepChanged) window.scrollTo({ top: 0, behavior: 'smooth' });
  }

  /**
   * Handle File Upload & Sheet Reading
   */
  function initFileUpload() {
    elements.fileInput.setAttribute('accept', Importers.acceptAttribute());

    elements.dropzone.addEventListener('click', function () {
      // Formats of all import adapters registered so far (adapters may be added later)
      elements.fileInput.setAttribute('accept', Importers.acceptAttribute());
      elements.fileInput.click();
    });

    elements.fileInput.addEventListener('change', function (e) {
      if (e.target.files && e.target.files[0]) {
        processUploadedFile(e.target.files[0]);
      }
      // Allow selecting the same file again
      e.target.value = '';
    });

    elements.dropzone.addEventListener('dragover', function (e) {
      e.preventDefault();
      elements.dropzone.classList.add('dragover');
    });

    elements.dropzone.addEventListener('dragleave', function () {
      elements.dropzone.classList.remove('dragover');
    });

    elements.dropzone.addEventListener('drop', function (e) {
      e.preventDefault();
      elements.dropzone.classList.remove('dragover');
      if (e.dataTransfer.files && e.dataTransfer.files[0]) {
        processUploadedFile(e.dataTransfer.files[0]);
      }
    });

    elements.sheetSelect.addEventListener('change', function () {
      loadSheet(elements.sheetSelect.value);
    });

    elements.previewShowMoreBtn.addEventListener('click', function () {
      state.previewLimit = (state.previewLimit || PREVIEW_INITIAL_ROWS) + PREVIEW_MORE_ROWS;
      renderDataPreview();
    });

    // Re-read the sheet when the user says the first row is (not) a header
    elements.hasHeaderRowCheckbox.addEventListener('change', function () {
      if (state.importSource && state.activeSheet) {
        loadSheet(state.activeSheet);
      }
    });

    // Sample data button
    elements.loadSampleBtn.addEventListener('click', function () {
      loadSampleDemo();
    });

    // Download physical sample excel
    elements.downloadSampleBtn.addEventListener('click', function () {
      downloadSampleExcelFile();
    });
  }

  /**
   * Project file buttons in the page header (available in every step)
   */
  function initProjectButtons() {
    elements.saveProjectBtn.addEventListener('click', function () {
      exportProjectFile();
    });

    elements.loadProjectBtn.addEventListener('click', function () {
      elements.projectFileInput.click();
    });

    elements.projectFileInput.addEventListener('change', function (e) {
      if (e.target.files && e.target.files[0]) {
        loadProjectFile(e.target.files[0]);
      }
      e.target.value = '';
    });
  }

  const PROJECT_ERROR_KEYS = {
    INVALID_JSON: 'projectErrInvalidJson',
    NOT_A_PROJECT: 'projectErrNotProject',
    FOREIGN_APP: 'projectErrForeignApp',
    NO_ROWS: 'projectErrNoRows',
    BAD_ROWS: 'projectErrBadRows',
    BAD_MAPPINGS: 'projectErrBadMappings',
    MISSING_COLUMN: 'projectErrMissingColumn'
  };

  const PROJECT_WARNING_KEYS = {
    NO_APP_ID: 'projectWarnNoAppId',
    NEWER_VERSION: 'projectWarnNewerVersion',
    UNUSED_TRANSLATIONS: 'projectWarnUnusedTranslations',
    ROW_COUNT_MISMATCH: 'projectWarnRowCountMismatch',
    INVALID_TRANSLATIONS: 'projectWarnInvalidTranslations',
    INVALID_SETTINGS: 'projectWarnInvalidSettings'
  };

  /**
   * Translated explanation of why a file cannot be loaded as a project.
   */
  function projectErrorMessage(err, fileName) {
    const params = Object.assign({ NAME: fileName || '' }, err && err.params);
    let key = err && PROJECT_ERROR_KEYS[err.code];
    // A spreadsheet picked in "Load project" by mistake
    if (err && err.code === 'INVALID_JSON' && Importers.findForFile(fileName || '')) key = 'projectErrSpreadsheet';
    const text = key ? t(key, params) : t('alertProjectLoadError', { MSG: err && err.message ? err.message : '' });
    return text + ' ' + t('projectLoadKeptState');
  }

  function hasWorkInProgress() {
    return Array.isArray(state.rows) && state.rows.length > 0;
  }

  /**
   * Loads a project file (.json) saved with "Save project". The file is validated first; if it
   * cannot be loaded, the current work stays untouched. Problems that did not prevent loading
   * are shown as a warning.
   */
  function loadProjectFile(file) {
    const reader = new FileReader();
    reader.onload = function (evt) {
      let parsed;
      try {
        parsed = StorageManager.parseProjectJson(evt.target.result);
      } catch (err) {
        console.warn(err);
        alert(projectErrorMessage(err, file.name));
        return;
      }
      if (hasWorkInProgress() && !confirm(t('confirmReplaceProject', { NAME: file.name }))) return;

      const backup = hasWorkInProgress() ? JSON.parse(StorageManager.serializeProjectJson(state)) : null;
      try {
        restoreProjectFromData(parsed);
      } catch (err) {
        // Restoring failed half-way: go back to the previous work
        console.error(err);
        if (backup) {
          restoreProjectFromData(StorageManager.parseProjectJson(backup));
        } else {
          resetWorkspaceForNewTable();
          state.rows = [];
          state.headers = [];
          setStep(1);
        }
        alert(t('alertProjectLoadError', { MSG: err.message }) + ' ' + t('projectLoadKeptState'));
        return;
      }
      persistState(0);
      showToast(t('toastProjectLoaded'), 'success');

      const warnings = (parsed.warnings || []).map(function (w) {
        return t(PROJECT_WARNING_KEYS[w.code] || w.code, w.params);
      });
      if (warnings.length > 0) {
        showToast(t('projectLoadedWithWarnings') + ' ' + warnings.join(' • '), 'warning');
      }
    };
    reader.onerror = function () {
      alert(t('alertProjectLoadError', { MSG: reader.error ? reader.error.message : file.name }));
    };
    reader.readAsText(file);
  }

  /**
   * Forget everything derived from a previously loaded table (translations of another file
   * must never be matched to the rows of a new one).
   */
  function resetWorkspaceForNewTable() {
    state.mappings = { refOldCol: '', refNewCol: '', targetCols: [], languages: {} };
    state.translations = {};
    state.activeTargetCol = '';
    state.auxiliaryCol = '';
    state.exportLangCol = '';
    state.partialContextCols = null;
    state.filterStatus = 'all';
    state.searchQuery = '';
    if (elements.searchInput) elements.searchInput.value = '';
    if (elements.filterChipsContainer) {
      elements.filterChipsContainer.querySelectorAll('.filter-chip').forEach(function (c) {
        c.classList.toggle('active', c.getAttribute('data-status') === 'all');
      });
    }
  }

  /**
   * Imports a file with the matching import adapter (see js/io/importers.js).
   */
  function processUploadedFile(file) {
    if (!Importers.findForFile(file.name, file.type) && /\.json$/i.test(file.name)) {
      loadProjectFile(file);
      return Promise.resolve();
    }
    return Importers.importFile(file).then(function (result) {
      state.fileName = file.name;
      state.importSource = result.source;
      state.sheetNames = result.source.sheetNames.slice();

      // Setup sheet select dropdown
      elements.sheetSelect.innerHTML = '';
      state.sheetNames.forEach(function (sheetName) {
        const opt = document.createElement('option');
        opt.value = sheetName;
        opt.textContent = sheetName;
        elements.sheetSelect.appendChild(opt);
      });

      const isSingleDoc = result.adapter && (result.adapter.id === 'word' || result.adapter.id === 'text');
      elements.sheetOptionsWrapper.style.display = isSingleDoc ? 'none' : 'flex';
      elements.sheetSelectWrapper.style.display = state.sheetNames.length > 1 ? 'block' : 'none';
      if (loadSheet(state.sheetNames[0])) {
        setStep(1);
        showToast(t('toastFileLoaded', { NAME: file.name }), 'success');
      }
    }).catch(function (err) {
      console.error('Import failed:', err);
      if (err && err.code === 'UNSUPPORTED') {
        alert(t('alertUnsupportedFile', { NAME: file.name, FORMATS: Importers.acceptAttribute().split(',').join(', ') }));
      } else if (err && err.code === 'NO_SHEETS') {
        alert(t('alertNoSheets'));
      } else {
        alert(t('alertFileReadError', { MSG: err && err.message ? err.message : file.name }));
      }
    });
  }

  function loadSheet(sheetName) {
    const hasHeader = elements.hasHeaderRowCheckbox.checked;
    const table = Importers.readSheet(state.importSource, sheetName, { hasHeaderRow: hasHeader });

    if (!table) {
      alert(t('alertEmptySheet'));
      return false;
    }

    state.activeSheet = sheetName;
    state.hasHeaderRow = hasHeader;
    elements.sheetSelect.value = sheetName;
    state.headers = table.headers;
    state.columnLetters = table.columnLetters;
    state.translations = {};
    state.rows = [];
    state.skippedRows = [];
    state.previewLimit = PREVIEW_INITIAL_ROWS;
    applyRowSelection(table.rows.map(function (row, idx) {
      const empty = isEmptyRow(row);
      return { rowNumber: table.rowNumbers[idx], row: row, skipped: empty, auto: empty };
    }));
    resetWorkspaceForNewTable();
    selectDefaultLanguageColumns();
    if (state.skippedRows.length > 0) {
      showToast(t('toastEmptyRowsSkipped', { COUNT: state.skippedRows.length }), 'info');
    }

    renderColumnMappingControls();
    renderDataPreview();
    revealMapping();
    persistState(300);
    return true;
  }

  // The mapping appears below the upload area; bring it into view so the next step is obvious
  function revealMapping() {
    if (typeof elements.mappingContainer.scrollIntoView === 'function') {
      elements.mappingContainer.scrollIntoView({ behavior: 'smooth', block: 'start' });
    }
  }

  function loadSampleDemo() {
    state.fileName = 'ukazkovy_dokument.xlsx';
    state.importSource = null;
    state.hasHeaderRow = true;
    state.rows = SampleGenerator.getSampleData();
    state.headers = SampleGenerator.getSampleHeaders();
    state.rowNumbers = state.rows.map(function (r, idx) { return idx + 2; });
    state.skippedRows = [];
    state.previewLimit = PREVIEW_INITIAL_ROWS;
    state.columnLetters = {};
    state.sheetNames = ['Data'];
    state.activeSheet = 'Data';
    resetWorkspaceForNewTable();

    elements.sheetOptionsWrapper.style.display = 'none';

    const sampleDefaults = SampleGenerator.getSampleDefaultMappings();
    state.mappings.refOldCol = sampleDefaults.refOldCol;
    state.mappings.refNewCol = sampleDefaults.refNewCol;
    sampleDefaults.targetCols.forEach(function (col) { setLanguageUsed(col, true); });

    renderColumnMappingControls();
    renderDataPreview();
    revealMapping();
    setStep(1);

    showToast(t('toastSampleLoaded'), 'success');
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

  /**
   * Step 1: Data preview. Every row can be marked as skipped (headings, notes, …); rows that
   * are completely empty are skipped automatically. Skipped rows are not processed, but the full
   * workbook keeps them unchanged at their original place.
   */
  const PREVIEW_INITIAL_ROWS = 10;
  const PREVIEW_MORE_ROWS = 20;

  function isEmptyRow(row) {
    return Object.keys(row).every(function (k) {
      const v = row[k];
      return v === '' || v === null || v === undefined || (typeof v === 'string' && !v.trim());
    });
  }

  // All rows of the sheet in their original order: [{ rowNumber, row, skipped, auto }]
  function allTableEntries() {
    const entries = state.rows.map(function (row, idx) {
      return { rowNumber: rowNumberOf(idx), row: row, skipped: false, auto: false };
    });
    (state.skippedRows || []).forEach(function (s) {
      entries.push({ rowNumber: s.rowNumber, row: s.row, skipped: true, auto: !!s.auto });
    });
    return entries.sort(function (a, b) { return a.rowNumber - b.rowNumber; });
  }

  /**
   * Splits the sheet into processed rows and skipped rows. Translation work is kept by row
   * number, so (un)skipping rows later never moves texts to other rows.
   */
  function applyRowSelection(entries) {
    const workByRow = {};
    Object.keys(state.translations || {}).forEach(function (col) {
      workByRow[col] = {};
      (state.translations[col] || []).forEach(function (stub, idx) {
        if (stub) workByRow[col][rowNumberOf(idx)] = stub;
      });
    });

    const active = entries.filter(function (e) { return !e.skipped; });
    state.rows = active.map(function (e) { return e.row; });
    state.rowNumbers = active.map(function (e) { return e.rowNumber; });
    state.skippedRows = entries.filter(function (e) { return e.skipped; }).map(function (e) {
      return { rowNumber: e.rowNumber, row: e.row, auto: !!e.auto };
    });

    Object.keys(workByRow).forEach(function (col) {
      state.translations[col] = state.rowNumbers.map(function (n) { return workByRow[col][n] || null; });
    });
  }

  function setRowSkipped(rowNumber, skipped) {
    const entries = allTableEntries();
    const entry = entries.find(function (e) { return e.rowNumber === rowNumber; });
    if (!entry || entry.auto) return;
    entry.skipped = skipped;
    applyRowSelection(entries);
    renderDataPreview();
    persistState(300);
  }

  function renderDataPreview() {
    if (!elements.dataPreviewContainer) return;
    const entries = allTableEntries();
    if (entries.length === 0 || !state.headers || state.headers.length === 0) {
      elements.dataPreviewContainer.style.display = 'none';
      return;
    }

    elements.dataPreviewContainer.style.display = 'block';
    const skippedCount = entries.length - state.rows.length;
    if (elements.previewRowsCount) {
      elements.previewRowsCount.textContent = `(${state.rows.length} ${t('rowsLabel')}` +
        (skippedCount > 0 ? ` · ${t('previewSkippedCount', { COUNT: skippedCount })}` : '') + ')';
    }

    const esc = ComparatorEngine.escapeHtml;
    const thStyle = 'padding:0.5rem 0.75rem; font-size:0.8rem; background:var(--bg-tertiary); border:1px solid var(--border-color);';
    elements.previewTableHead.innerHTML = '<tr>' +
      `<th style="${thStyle}" title="${esc(t('previewSkipTitle'))}">${esc(t('previewSkipCol'))}</th>` +
      `<th style="${thStyle}">${esc(t('rowLabel'))}</th>` +
      state.headers.map(function (h) {
        return `<th style="${thStyle}">${esc(columnDisplayName(h))}</th>`;
      }).join('') + '</tr>';

    // Whitespace is collapsed so long single-line cells stay readable
    const tdStyle = 'padding:0.45rem 0.75rem; font-size:0.8rem; border:1px solid var(--border-color); color:var(--text-secondary); max-width:240px; overflow:hidden; text-overflow:ellipsis; white-space:nowrap;';
    const limit = state.previewLimit || PREVIEW_INITIAL_ROWS;
    elements.previewTableBody.innerHTML = entries.slice(0, limit).map(function (e) {
      const cells = state.headers.map(function (h) {
        let val = e.row[h] !== undefined ? String(e.row[h]).replace(/\s+/g, ' ').trim() : '';
        const full = val;
        if (val.length > 60) val = val.substring(0, 57) + '...';
        return `<td style="${tdStyle}" title="${esc(full.slice(0, 500))}">${esc(val)}</td>`;
      }).join('');
      const skipTitle = e.auto ? t('previewEmptyRow') : t('previewSkipTitle');
      return `<tr class="preview-row${e.skipped ? ' preview-row-skipped' : ''}" data-row="${e.rowNumber}">` +
        `<td style="${tdStyle} text-align:center;"><input type="checkbox" class="preview-skip-cb" data-row="${e.rowNumber}"` +
        `${e.skipped ? ' checked' : ''}${e.auto ? ' disabled' : ''} title="${esc(skipTitle)}" aria-label="${esc(t('previewSkipCol'))} ${e.rowNumber}"></td>` +
        `<td style="${tdStyle} font-family:var(--font-mono);">${e.rowNumber}${e.auto ? ` <span class="preview-empty-badge">${esc(t('previewEmptyRow'))}</span>` : ''}</td>` +
        cells + '</tr>';
    }).join('');

    elements.previewTableBody.querySelectorAll('.preview-skip-cb').forEach(function (cb) {
      cb.addEventListener('change', function () {
        setRowSkipped(parseInt(cb.getAttribute('data-row'), 10), cb.checked);
      });
    });

    const remaining = entries.length - limit;
    elements.previewShowMoreBtn.style.display = remaining > 0 ? 'inline-flex' : 'none';
    elements.previewShowMoreBtn.textContent = t('previewShowMore', { COUNT: Math.min(remaining, PREVIEW_MORE_ROWS), TOTAL: remaining });
  }

  function downloadSampleExcelFile() {
    const exporter = Exporters.getDefault('table');
    Exporters.run(exporter.id, {
      tables: [{ name: 'Data', headers: SampleGenerator.getSampleHeaders(), rows: SampleGenerator.getSampleData() }]
    }, { baseName: 'ukazkovy_dokument', uiLang: I18n.getLanguage() });
    showToast(t('toastSampleDownloaded'), 'success');
  }

  /**
   * Column selection: two reference columns, then for every other column whether it holds the
   * old version of a language (used) or is ignored, the language name and which column holds or
   * will hold its new version.
   */

  // Columns with longer text in a reasonable share of rows are pre-selected as languages;
  // the actual languages are never guessed from the header names.
  function selectDefaultLanguageColumns() {
    if (state.headers.length === 1) {
      state.mappings.refNewCol = state.headers[0];
      state.mappings.refOldCol = '';
      setLanguageUsed(state.headers[0], false);
      return;
    }
    state.headers.forEach(function (h) {
      let filled = 0;
      let totalWords = 0;
      state.rows.forEach(function (row) {
        const v = row[h];
        if (typeof v === 'string' && v) {
          filled++;
          totalWords += v.split(/\s+/).filter(Boolean).length;
        }
      });
      const isText = filled >= Math.max(1, state.rows.length * 0.5) && totalWords / Math.max(filled, 1) >= 3;
      if (isText) setLanguageUsed(h, true);
    });
  }

  function ensureLanguageConfig(col) {
    if (!state.mappings.languages) state.mappings.languages = {};
    if (!state.mappings.languages[col]) {
      const isNew = ComparatorEngine.isNewLanguage(state.mappings, col);
      state.mappings.languages[col] = { label: '', newCol: '', newColName: '', isNewLanguage: isNew };
    }
    return state.mappings.languages[col];
  }

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

  function newColumnOwner(col) {
    return state.mappings.targetCols.find(function (c) {
      return state.mappings.languages[c] && state.mappings.languages[c].newCol === col;
    }) || '';
  }

  function setLanguageUsed(col, used) {
    const targets = state.mappings.targetCols;
    if (used) {
      ensureLanguageConfig(col);
      if (targets.indexOf(col) === -1) targets.push(col);
      // Keep original sheet columns in header order; put newly added languages at the end
      targets.sort(function (a, b) {
        const idxA = state.headers.indexOf(a);
        const idxB = state.headers.indexOf(b);
        if (idxA !== -1 && idxB !== -1) return idxA - idxB;
        if (idxA !== -1) return -1;
        if (idxB !== -1) return 1;
        return 0;
      });
    } else {
      const pos = targets.indexOf(col);
      if (pos !== -1) targets.splice(pos, 1);
    }
  }

  function addNewLanguage(name, colName) {
    const trimmedName = (name || '').trim();
    if (!trimmedName) {
      showToast(t('toastEnterLangName'), 'warning');
      if (elements.newLangNameInput) elements.newLangNameInput.focus();
      return null;
    }

    const existing = state.mappings.targetCols.some(function (c) {
      const info = ComparatorEngine.getLanguageInfo(state.mappings, c);
      return (info.label && info.label.toLowerCase() === trimmedName.toLowerCase()) ||
             (info.newColName && info.newColName.toLowerCase() === trimmedName.toLowerCase()) ||
             c.toLowerCase() === trimmedName.toLowerCase();
    });
    if (existing) {
      showToast(t('toastLangAlreadyExists'), 'warning');
      return null;
    }

    const id = '__new__' + trimmedName.replace(/[^a-zA-Z0-9_\u00C0-\u024F\u1E00-\u1EFF]/g, '_') + '_' + Date.now();
    const exportColName = (colName || '').trim() || trimmedName;

    if (!state.mappings.languages) state.mappings.languages = {};
    state.mappings.languages[id] = {
      label: trimmedName,
      newCol: '',
      newColName: exportColName,
      isNewLanguage: true
    };

    setLanguageUsed(id, true);
    renderLanguageTable();
    if (isDirectTranslationMode()) {
      renderQuickLangs();
      renderSelectedDirectLangs();
    }
    persistState(300);
    showToast(t('toastNewLangAdded', { NAME: trimmedName }), 'success');
    if (elements.newLangNameInput) elements.newLangNameInput.value = '';
    if (elements.newLangColNameInput) elements.newLangColNameInput.value = '';
    return id;
  }

  function removeNewLanguage(col) {
    const info = ComparatorEngine.getLanguageInfo(state.mappings, col);
    const name = info.label || col;
    setLanguageUsed(col, false);
    delete state.mappings.languages[col];
    delete state.translations[col];
    if (state.activeTargetCol === col) {
      state.activeTargetCol = state.mappings.targetCols[0] || '';
    }
    if (state.exportLangCol === col) {
      state.exportLangCol = state.mappings.targetCols[0] || '';
    }
    renderLanguageTable();
    if (isDirectTranslationMode()) {
      renderQuickLangs();
      renderSelectedDirectLangs();
    }
    persistState(300);
    showToast(t('toastNewLangRemoved', { NAME: name }), 'info');
  }

  const QUICK_LANGS = [
    { code: 'CS', label: 'Čeština', flag: '🇨🇿' },
    { code: 'SK', label: 'Slovenčina', flag: '🇸🇰' },
    { code: 'EN', label: 'English', flag: '🇬🇧' },
    { code: 'DE', label: 'DE', flag: '🇩🇪' },
    { code: 'PL', label: 'Polski', flag: '🇵🇱' },
    { code: 'FR', label: 'Français', flag: '🇫🇷' },
    { code: 'ES', label: 'Español', flag: '🇪🇸' },
    { code: 'IT', label: 'Italiano', flag: '🇮🇹' },
    { code: 'UK', label: 'Українська', flag: '🇺🇦' }
  ];

  function isLanguageSelected(labelOrCode) {
    const needle = (labelOrCode || '').toLowerCase();
    return state.mappings.targetCols.some(function (c) {
      const info = ComparatorEngine.getLanguageInfo(state.mappings, c);
      const label = (info.label || '').toLowerCase();
      const newColName = (info.newColName || '').toLowerCase();
      const colId = String(c).toLowerCase();
      return label === needle || newColName === needle || colId === needle;
    });
  }

  function findColForLanguage(labelOrCode) {
    const needle = (labelOrCode || '').toLowerCase();
    return state.mappings.targetCols.find(function (c) {
      const info = ComparatorEngine.getLanguageInfo(state.mappings, c);
      const label = (info.label || '').toLowerCase();
      const newColName = (info.newColName || '').toLowerCase();
      const colId = String(c).toLowerCase();
      return label === needle || newColName === needle || colId === needle;
    }) || '';
  }

  function renderQuickLangs() {
    if (!elements.quickLangsContainer) return;
    elements.quickLangsContainer.innerHTML = '';
    QUICK_LANGS.forEach(function (item) {
      const selected = isLanguageSelected(item.label) || isLanguageSelected(item.code);
      const btn = document.createElement('button');
      btn.type = 'button';
      btn.className = 'quick-lang-btn' + (selected ? ' active' : '');
      btn.innerHTML = `<span class="quick-lang-flag">${item.flag}</span> <span class="quick-lang-name">${ComparatorEngine.escapeHtml(item.label)}</span>` + (selected ? ' ✓' : '');
      btn.addEventListener('click', function () {
        if (selected) {
          const col = findColForLanguage(item.label) || findColForLanguage(item.code);
          if (col) removeNewLanguage(col);
        } else {
          addNewLanguage(item.label, item.label);
        }
      });
      elements.quickLangsContainer.appendChild(btn);
    });
  }

  function renderSelectedDirectLangs() {
    if (!elements.selectedDirectLangsList) return;
    const esc = ComparatorEngine.escapeHtml;
    const cols = state.mappings.targetCols || [];
    if (cols.length === 0) {
      elements.selectedDirectLangsList.innerHTML = `<span style="color:var(--text-muted); font-size:0.85rem; font-style:italic;">${esc(t('noLangsSelectedYet'))}</span>`;
      return;
    }
    elements.selectedDirectLangsList.innerHTML = '';
    cols.forEach(function (col) {
      const info = ComparatorEngine.getLanguageInfo(state.mappings, col);
      const label = info.label || col;
      const chip = document.createElement('div');
      chip.className = 'selected-lang-chip';
      chip.innerHTML = `<span>🌐 ${esc(label)}</span><button type="button" class="remove-direct-lang-btn" title="${esc(t('removeLangTitle'))}">✕</button>`;
      chip.querySelector('.remove-direct-lang-btn').addEventListener('click', function () {
        removeNewLanguage(col);
      });
      elements.selectedDirectLangsList.appendChild(chip);
    });
  }

  /**
   * Store final labels and new-column names, so exports and the project file are explicit.
   */
  function finalizeLanguageConfig() {
    state.mappings.targetCols.forEach(function (col) {
      const cfg = ensureLanguageConfig(col);
      if (!cfg.label) cfg.label = languageLabel(col);
      if (cfg.newCol && state.headers.indexOf(cfg.newCol) === -1) cfg.newCol = '';
      if (!cfg.newColName) cfg.newColName = defaultNewColNameFor(col);
    });
  }

  /**
   * When the reference columns differ by a version number (e.g. "EN 2024" / "EN 2026") and the
   * sheet already has the matching column for a language ("CS 2026"), use it as its new version.
   */
  function autoPairNewColumns() {
    state.mappings.targetCols.slice().forEach(function (col) {
      const cfg = ensureLanguageConfig(col);
      if (cfg.newCol) return;
      const candidate = ComparatorEngine.defaultNewColName(col, state.mappings.refOldCol, state.mappings.refNewCol, '\u0000');
      if (candidate.indexOf('\u0000') !== -1 || candidate === col || state.headers.indexOf(candidate) === -1) return;
      if (candidate === state.mappings.refOldCol || candidate === state.mappings.refNewCol || newColumnOwner(candidate)) return;
      setLanguageUsed(candidate, false);
      cfg.newCol = candidate;
    });
  }

  function onReferenceChanged() {
    const m = state.mappings;
    [m.refOldCol, m.refNewCol].forEach(function (ref) {
      if (!ref) return;
      setLanguageUsed(ref, false);
      m.targetCols.forEach(function (c) {
        if (m.languages[c] && m.languages[c].newCol === ref) m.languages[c].newCol = '';
      });
    });
    autoPairNewColumns();
    renderLanguageTable();
  }

  function fillColumnSelect(select, value, allowNone) {
    select.innerHTML = '';
    const placeholder = document.createElement('option');
    placeholder.value = '';
    placeholder.textContent = allowNone ? t('optNoOldReference') : t('selectColumnPlaceholder');
    select.appendChild(placeholder);
    state.headers.forEach(function (h) {
      const opt = document.createElement('option');
      opt.value = h;
      opt.textContent = columnDisplayName(h);
      select.appendChild(opt);
    });
    select.value = value || '';
  }

  function renderColumnMappingControls() {
    elements.mappingContainer.style.display = 'block';
    const isDirect = isDirectTranslationMode();

    if (isDirect) {
      if (elements.refOldColCard) elements.refOldColCard.style.display = 'none';
      if (elements.directTransInfoBox) elements.directTransInfoBox.style.display = 'block';
      if (elements.refNewColLabel) elements.refNewColLabel.textContent = t('sourceTextLabel');
      if (elements.refNewColHelp) elements.refNewColHelp.textContent = t('sourceTextCardHelp');
      if (elements.targetColsTitle) elements.targetColsTitle.textContent = t('targetLangsDirectTitle');
      if (elements.targetColsHelp) elements.targetColsHelp.textContent = t('targetLangsDirectHelp');
      if (elements.quickLangsWrapper) elements.quickLangsWrapper.style.display = 'block';
      if (elements.languagesTableWrapper) elements.languagesTableWrapper.style.display = 'none';
      if (elements.confirmMappingBtn) elements.confirmMappingBtn.textContent = t('goToTranslationDirect');

      if (elements.stepNav2) elements.stepNav2.style.display = 'none';
      if (elements.stepDivider2) elements.stepDivider2.style.display = 'none';
      if (elements.stepNumber3) elements.stepNumber3.textContent = '2';
      if (elements.stepNumber4) elements.stepNumber4.textContent = '3';

      renderQuickLangs();
      renderSelectedDirectLangs();
    } else {
      if (elements.refOldColCard) elements.refOldColCard.style.display = '';
      if (elements.directTransInfoBox) elements.directTransInfoBox.style.display = 'none';
      if (elements.refNewColLabel) elements.refNewColLabel.textContent = t('refNewLabel');
      if (elements.refNewColHelp) elements.refNewColHelp.textContent = t('refNewHelp');
      if (elements.targetColsTitle) elements.targetColsTitle.textContent = t('targetColsLabel');
      if (elements.targetColsHelp) elements.targetColsHelp.textContent = t('targetColsHelp');
      if (elements.quickLangsWrapper) elements.quickLangsWrapper.style.display = 'none';
      if (elements.languagesTableWrapper) elements.languagesTableWrapper.style.display = '';
      if (elements.confirmMappingBtn) elements.confirmMappingBtn.textContent = t('confirmMappingBtn');

      if (elements.stepNav2) elements.stepNav2.style.display = '';
      if (elements.stepDivider2) elements.stepDivider2.style.display = '';
      if (elements.stepNumber3) elements.stepNumber3.textContent = '3';
      if (elements.stepNumber4) elements.stepNumber4.textContent = '4';
    }

    fillColumnSelect(elements.refOldSelect, state.mappings.refOldCol, true);
    fillColumnSelect(elements.refNewSelect, state.mappings.refNewCol, false);

    elements.refOldSelect.onchange = function () {
      state.mappings.refOldCol = elements.refOldSelect.value;
      onReferenceChanged();
    };

    elements.refNewSelect.onchange = function () {
      state.mappings.refNewCol = elements.refNewSelect.value;
      onReferenceChanged();
    };

    renderLanguageTable();
  }

  function firstSampleText(col) {
    const row = state.rows.find(function (r) { return r[col] !== '' && r[col] !== undefined && r[col] !== null; });
    return row ? String(row[col]).replace(/\s+/g, ' ').slice(0, 200) : '';
  }

  function renderLanguageTable() {
    const tbody = elements.languagesTableBody;
    const m = state.mappings;
    const esc = ComparatorEngine.escapeHtml;
    tbody.innerHTML = '';

    state.headers.forEach(function (h) {
      if (h === m.refOldCol || h === m.refNewCol) return;
      const used = m.targetCols.indexOf(h) !== -1;
      const owner = used ? '' : newColumnOwner(h);
      const tr = document.createElement('tr');
      tr.className = used ? 'lang-row-used' : (owner ? 'lang-row-paired' : 'lang-row-ignored');
      tr.setAttribute('data-col', h);

      let html = `<td><input type="checkbox" class="lang-use-cb" ${used ? 'checked' : ''} ${owner ? 'disabled' : ''}></td>`;
      html += `<td class="lang-col-name" title="${esc(firstSampleText(h))}">${esc(columnDisplayName(h))}</td>`;

      if (used) {
        const cfg = ensureLanguageConfig(h);
        const candidates = state.headers.filter(function (c) {
          return c !== h && c !== m.refOldCol && c !== m.refNewCol && m.targetCols.indexOf(c) === -1 &&
            (!newColumnOwner(c) || newColumnOwner(c) === h);
        });
        html += `<td><input type="text" class="form-input lang-label-input" value="${esc(languageLabel(h))}"></td>`;
        html += '<td><div class="new-col-cell"><select class="form-select lang-newcol-select">' +
          `<option value="">${esc(t('optCreateNewCol'))}</option>` +
          candidates.map(function (c) {
            return `<option value="${esc(c)}" ${cfg.newCol === c ? 'selected' : ''}>${esc(columnDisplayName(c))}</option>`;
          }).join('') + '</select>' +
          (cfg.newCol ? '' : `<input type="text" class="form-input lang-newname-input" value="${esc(cfg.newColName || '')}" placeholder="${esc(defaultNewColNameFor(h))}" title="${esc(t('newColNamePlaceholder'))}">`) +
          '</div></td>';
      } else if (owner) {
        html += `<td colspan="2">${esc(t('usedAsNewVersionOf', { LANG: languageLabel(owner) }))}</td>`;
      } else {
        html += `<td colspan="2">${esc(t('ignoredColumn'))}</td>`;
      }
      tr.innerHTML = html;
      tbody.appendChild(tr);

      const cb = tr.querySelector('.lang-use-cb');
      cb.addEventListener('change', function () {
        setLanguageUsed(h, cb.checked);
        renderLanguageTable();
      });
      const labelInput = tr.querySelector('.lang-label-input');
      if (labelInput) {
        labelInput.addEventListener('input', function () {
          ensureLanguageConfig(h).label = labelInput.value.trim();
        });
      }
      const newColSelect = tr.querySelector('.lang-newcol-select');
      if (newColSelect) {
        newColSelect.addEventListener('change', function () {
          ensureLanguageConfig(h).newCol = newColSelect.value;
          renderLanguageTable();
        });
      }
      const newNameInput = tr.querySelector('.lang-newname-input');
      if (newNameInput) {
        newNameInput.addEventListener('input', function () {
          ensureLanguageConfig(h).newColName = newNameInput.value.trim();
        });
      }
    });

    // Render newly added languages (not in original sheet headers)
    state.mappings.targetCols.filter(function (c) {
      return ComparatorEngine.isNewLanguage(state.mappings, c);
    }).forEach(function (newColId) {
      const cfg = ensureLanguageConfig(newColId);
      const tr = document.createElement('tr');
      tr.className = 'lang-row-used lang-row-new';
      tr.setAttribute('data-col', newColId);

      let html = '<td><input type="checkbox" class="lang-use-cb" checked></td>';
      html += `<td class="lang-col-name">${esc(languageLabel(newColId))} <span class="badge badge-info" style="font-size:0.7rem; padding:0.15rem 0.4rem; vertical-align:middle;">${esc(t('badgeNewLanguage'))}</span><small style="display:block; color:var(--text-muted); font-size:0.75rem;">${esc(t('noOldVersion'))}</small></td>`;
      html += `<td><input type="text" class="form-input lang-label-input" value="${esc(languageLabel(newColId))}"></td>`;
      html += `<td><div class="new-col-cell"><input type="text" class="form-input lang-newname-input" value="${esc(cfg.newColName || '')}" placeholder="${esc(languageLabel(newColId))}" title="${esc(t('newColNamePlaceholder'))}"><button type="button" class="btn btn-sm btn-outline remove-lang-btn" data-col="${esc(newColId)}" title="${esc(t('removeLangTitle'))}" style="color:#ef4444; border-color:var(--border-color); padding:0.3rem 0.6rem;">✕</button></div></td>`;

      tr.innerHTML = html;
      tbody.appendChild(tr);

      const cb = tr.querySelector('.lang-use-cb');
      cb.addEventListener('change', function () {
        if (!cb.checked) {
          removeNewLanguage(newColId);
        }
      });
      const labelInput = tr.querySelector('.lang-label-input');
      if (labelInput) {
        labelInput.addEventListener('input', function () {
          ensureLanguageConfig(newColId).label = labelInput.value.trim();
        });
      }
      const newNameInput = tr.querySelector('.lang-newname-input');
      if (newNameInput) {
        newNameInput.addEventListener('input', function () {
          ensureLanguageConfig(newColId).newColName = newNameInput.value.trim();
        });
      }
      const removeBtn = tr.querySelector('.remove-lang-btn');
      if (removeBtn) {
        removeBtn.addEventListener('click', function () {
          removeNewLanguage(newColId);
        });
      }
    });
  }

  /**
   * Initialize translation stubs for target languages
   */
  function ensureStubsInitialized() {
    finalizeLanguageConfig();
    state.mappings.targetCols.forEach(function (targetCol) {
      // Rebuild every stub from the current rows/mapping and merge the saved work into it, so
      // reference texts, the old target text and classifications are always up to date
      // (also after changing the reference columns or loading a compact project file).
      // Text already present in the language's new-version column is used as the starting point.
      const isNew = ComparatorEngine.isNewLanguage(state.mappings, targetCol);
      state.translations[targetCol] = ComparatorEngine.rehydrateStubs(
        state.rows,
        state.translations[targetCol],
        targetCol,
        state.mappings.refOldCol,
        state.mappings.refNewCol,
        state.mappings.languages[targetCol].newCol,
        isNew
      );
    });

    if (!state.activeTargetCol || !state.mappings.targetCols.includes(state.activeTargetCol)) {
      state.activeTargetCol = state.mappings.targetCols[0] || '';
    }

    persistState(400);
  }

  /**
   * Step 2: Render Reference Changes Overview
   */
  function renderReferenceOverview() {
    const refOldCol = state.mappings.refOldCol;
    const refNewCol = state.mappings.refNewCol;

    let unchangedCount = 0;
    let modifiedCount = 0;
    let addedCount = 0;
    let removedCount = 0;

    elements.refDiffTableBody.innerHTML = '';

    state.rows.forEach(function (row, idx) {
      const refOld = refOldCol ? row[refOldCol] : '';
      const refNew = row[refNewCol];
      const classification = ComparatorEngine.classifyRow(refOld, refNew);

      if (classification === 'UNCHANGED') unchangedCount++;
      else if (classification === 'MODIFIED') modifiedCount++;
      else if (classification === 'ADDED') addedCount++;
      else if (classification === 'REMOVED') removedCount++;

      const tr = document.createElement('tr');
      tr.style.borderBottom = '1px solid var(--border-color)';

      // Paragraphs have no headings of their own; the Excel row number identifies them
      const rowLabel = rowLabelOf(idx);

      // Diff preview
      let diffHtml = '';
      if (!refOldCol) {
        diffHtml = '<div class="diff-box" style="color:var(--text-primary);">' + ComparatorEngine.escapeHtml(refNew) + '</div>';
      } else if (classification === 'UNCHANGED') {
        diffHtml = '<div class="diff-unchanged" style="max-height:80px; overflow-y:auto;">' + ComparatorEngine.escapeHtml(refOld) + '</div>';
      } else if (classification === 'ADDED') {
        diffHtml = '<div class="diff-box"><ins class="diff-ins">' + ComparatorEngine.escapeHtml(refNew) + '</ins></div>';
      } else if (classification === 'REMOVED') {
        diffHtml = '<div class="diff-box"><del class="diff-del">' + ComparatorEngine.escapeHtml(refOld) + '</del></div>';
      } else {
        const diffRes = ComparatorEngine.computeDiff(refOld, refNew);
        diffHtml = '<div class="diff-box">' + ComparatorEngine.renderDiffHtml(diffRes, 'inline') + '</div>';
      }

      tr.innerHTML = `
        <td style="padding:0.75rem; font-weight:700; font-family:var(--font-mono); white-space:nowrap; vertical-align:top;">${ComparatorEngine.escapeHtml(rowLabel)}</td>
        <td style="padding:0.75rem; vertical-align:top;">${renderBadgeHtml(classification)}</td>
        <td style="padding:0.75rem; vertical-align:top;">${diffHtml}</td>
      `;

      elements.refDiffTableBody.appendChild(tr);
    });

    elements.kpiTotalRows.textContent = state.rows.length;
    elements.kpiUnchanged.textContent = unchangedCount;
    elements.kpiModified.textContent = modifiedCount;
    elements.kpiAdded.textContent = addedCount;
    elements.kpiRemoved.textContent = removedCount;
  }

  function renderBadgeHtml(classification) {
    switch (classification) {
      case 'UNCHANGED':
        return `<span class="badge badge-unchanged">${t('badgeUnchanged')}</span>`;
      case 'MODIFIED':
        return `<span class="badge badge-modified">${t('badgeModified')}</span>`;
      case 'ADDED':
        return `<span class="badge badge-added">${t('badgeAdded')}</span>`;
      case 'REMOVED':
        return `<span class="badge badge-removed">${t('badgeRemoved')}</span>`;
      default:
        return `<span class="badge">${classification}</span>`;
    }
  }

  /**
   * Step 3: Render Translation Studio (Full-width side-by-side)
   */
  function renderTranslationStudio() {
    updateAuxLangOptions();
    updateLayoutToggleText();
    renderCrossLanguageMatrix();
    renderLanguageTabs();
    renderArticleCards();
    updateLanguageProgress();
  }

  /**
   * Render Cross-Language Progress Matrix
   */
  function renderCrossLanguageMatrix() {
    if (!elements.crossLangMatrixCard || !elements.crossLangMatrixGrid) return;
    const targetCols = state.mappings.targetCols || [];
    if (targetCols.length === 0) {
      elements.crossLangMatrixCard.style.display = 'none';
      return;
    }

    elements.crossLangMatrixCard.style.display = 'block';
    elements.crossLangMatrixGrid.innerHTML = '';

    targetCols.forEach(function (col) {
      const stubs = state.translations[col] || [];
      const total = stubs.length;
      const verified = stubs.filter(function (s) { return s.isVerified; }).length;
      const pct = total > 0 ? Math.round((verified / total) * 100) : 0;
      const isComplete = verified === total && total > 0;
      const isActive = col === state.activeTargetCol;

      const pill = document.createElement('div');
      pill.className = 'matrix-pill' + (isActive ? ' active' : '') + (isComplete ? ' complete' : '');
      pill.innerHTML = `
        <span class="matrix-pill-lang">${ComparatorEngine.escapeHtml(languageLabel(col))}</span>
        <span class="matrix-pill-pct">${pct}% (${verified}/${total})</span>
      `;

      pill.addEventListener('click', function () {
        state.activeTargetCol = col;
        updateAuxLangOptions();
        renderLanguageTabs();
        renderArticleCards();
        updateLanguageProgress();
        renderCrossLanguageMatrix();
      });

      elements.crossLangMatrixGrid.appendChild(pill);
    });
  }

  /**
   * Bulk approve unchanged articles in currently active language
   */
  function approveUnchangedInActiveLanguage() {
    const col = state.activeTargetCol;
    if (!col || !state.translations[col]) return;
    const isNew = ComparatorEngine.isNewLanguage(state.mappings, col);

    const stubs = state.translations[col];
    let count = 0;
    stubs.forEach(function (stub) {
      if (stub.classification === 'UNCHANGED' && !stub.isVerified) {
        if (isNew && !stub.targetNew) return;
        stub.isVerified = true;
        stub.status = 'verified';
        count++;
      }
    });

    renderArticleCards();
    updateLanguageProgress();
    renderLanguageTabs();
    renderCrossLanguageMatrix();
    persistState(150);

    showToast(t('toastBulkApproved', { COUNT: count }), 'success');
  }

  /**
   * Global bulk approve unchanged articles across ALL languages
   */
  function approveAllUnchangedGlobal() {
    const targetCols = state.mappings.targetCols || [];
    let count = 0;

    targetCols.forEach(function (col) {
      const isNew = ComparatorEngine.isNewLanguage(state.mappings, col);
      const stubs = state.translations[col] || [];
      stubs.forEach(function (stub) {
        if (stub.classification === 'UNCHANGED' && !stub.isVerified) {
          if (isNew && !stub.targetNew) return;
          stub.isVerified = true;
          stub.status = 'verified';
          count++;
        }
      });
    });

    renderArticleCards();
    updateLanguageProgress();
    renderLanguageTabs();
    renderCrossLanguageMatrix();
    persistState(150);

    showToast(t('toastBulkApproved', { COUNT: count }), 'success');
  }

  function renderLanguageTabs() {
    elements.langTabsWrapper.innerHTML = '';

    state.mappings.targetCols.forEach(function (col) {
      const tab = document.createElement('div');
      tab.className = 'lang-tab' + (col === state.activeTargetCol ? ' active' : '');

      const stubs = state.translations[col] || [];
      const verifiedCount = stubs.filter(function (s) { return s.isVerified; }).length;
      const total = stubs.length;

      tab.innerHTML = `
        <span>${ComparatorEngine.escapeHtml(languageLabel(col))}</span>
        <span class="lang-tab-badge">${verifiedCount}/${total}</span>
      `;

      tab.addEventListener('click', function () {
        state.activeTargetCol = col;
        updateAuxLangOptions();
        renderCrossLanguageMatrix();
        renderLanguageTabs();
        renderArticleCards();
        updateLanguageProgress();
      });

      elements.langTabsWrapper.appendChild(tab);
    });
  }

  function updateLanguageProgress() {
    const col = state.activeTargetCol;
    if (!col || !state.translations[col]) return;

    const stubs = state.translations[col];
    const total = stubs.length;
    const verified = stubs.filter(function (s) { return s.isVerified; }).length;
    const pct = total > 0 ? Math.round((verified / total) * 100) : 0;

    elements.currentLangProgressText.textContent = `${verified} / ${total} ${t('progressDone')} (${pct}%)`;
    elements.currentLangProgressBar.style.width = pct + '%';
  }

  /**
   * Auxiliary Language Selector (Context helper)
   */
  function updateAuxLangOptions() {
    if (!elements.auxLangSelect) return;
    const select = elements.auxLangSelect;
    const currentVal = state.auxiliaryCol;
    select.innerHTML = '';

    const optNone = document.createElement('option');
    optNone.value = '';
    optNone.textContent = t('auxLangNone');
    select.appendChild(optNone);

    const availableCols = (state.mappings && state.mappings.targetCols) || [];
    availableCols.forEach(function (col) {
      if (col === state.activeTargetCol) return;
      const opt = document.createElement('option');
      opt.value = col;
      opt.textContent = languageLabel(col);
      select.appendChild(opt);
    });

    if (currentVal && currentVal !== state.activeTargetCol && availableCols.includes(currentVal)) {
      select.value = currentVal;
    } else {
      select.value = '';
      state.auxiliaryCol = '';
    }
  }

  function initAuxLangSelector() {
    if (!elements.auxLangSelect) return;
    elements.auxLangSelect.addEventListener('change', function () {
      state.auxiliaryCol = elements.auxLangSelect.value;
      persistState(0);
      updateLayoutToggleText();
      renderArticleCards();
    });
  }

  function highlightQuery(text, query) {
    if (!text) return '';
    const safe = ComparatorEngine.escapeHtml(text);
    if (!query) return safe;
    const escaped = query.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const regex = new RegExp('(' + escaped + ')', 'gi');
    return safe.replace(regex, '<mark class="search-highlight">$1</mark>');
  }

  /**
   * Column 4 diff of the language (old → new version). An empty editor is shown as "not translated
   * yet" rather than as the whole old text being deleted (except for removed paragraphs).
   */
  function renderTargetDiffHtml(stub) {
    const hint = function (key) {
      return `<span style="color:var(--text-muted); font-style:italic;">${t(key)}</span>`;
    };
    if (!String(stub.targetNew || '').trim() && stub.classification !== 'REMOVED') {
      return hint('targetNotTranslatedYet');
    }
    const html = ComparatorEngine.renderDiffHtml(ComparatorEngine.computeDiff(stub.targetOld, stub.targetNew), 'inline');
    return html || hint('noChangesTargetDiff');
  }

  function renderArticleCards() {
    const col = state.activeTargetCol;
    if (!col || !state.translations[col]) {
      elements.articlesContainer.innerHTML = '<div style="padding:2rem; text-align:center; color:var(--text-muted);">' + t('noTargetLang') + '</div>';
      return;
    }

    elements.articlesContainer.innerHTML = '';
    const stubs = state.translations[col];
    const isNewLang = ComparatorEngine.isNewLanguage(state.mappings, col);
    const esc = ComparatorEngine.escapeHtml;
    const langLabel = esc(languageLabel(col));
    const refLabel = esc(columnDisplayName(state.mappings.refNewCol));
    const auxCol = (state.auxiliaryCol && state.auxiliaryCol !== col && state.translations[state.auxiliaryCol])
      ? state.auxiliaryCol
      : '';
    const auxLabel = auxCol ? esc(languageLabel(auxCol)) : '';

    let renderedCount = 0;

    stubs.forEach(function (stub, idx) {
      const classification = stub.classification;

      // Filtering logic
      if (state.filterStatus === 'pending' && stub.isVerified) return;
      if (state.filterStatus === 'verified' && !stub.isVerified) return;
      if (state.filterStatus === 'modified' && classification !== 'MODIFIED') return;
      if (state.filterStatus === 'added' && classification !== 'ADDED') return;
      if (state.filterStatus === 'removed' && classification !== 'REMOVED') return;

      // Search query filter
      if (state.searchQuery) {
        const q = state.searchQuery.toLowerCase();
        const fullContent = (rowLabelOf(idx) + ' ' + stub.refOld + ' ' + stub.refNew + ' ' + stub.targetOld + ' ' + stub.targetNew).toLowerCase();
        if (!fullContent.includes(q)) return;
      }

      renderedCount++;

      const card = document.createElement('div');
      card.className = 'article-card ' + (stub.isVerified ? 'status-verified' : 'status-pending');
      card.id = 'article-card-' + idx;

      // Paragraphs have no headings; only the row number is shown above the texts
      const rowTitle = rowLabelOf(idx);

      // 1. Reference Diff (old -> new)
      const refDiffRes = ComparatorEngine.computeDiff(stub.refOld, stub.refNew);
      const refDiffHtml = ComparatorEngine.renderDiffHtml(refDiffRes, 'inline');

      // 2. Auxiliary Reference Column (Optional context)
      let colAuxHtml = '';
      if (auxCol && state.translations[auxCol]) {
        const auxStub = state.translations[auxCol][idx];
        const auxText = (auxStub && auxStub.targetNew) ? auxStub.targetNew : (auxStub ? auxStub.targetOld : (state.rows[idx] ? state.rows[idx][auxCol] : ''));
        let auxBadge = '';
        if (auxStub && auxStub.isVerified) {
          auxBadge = `<span class="badge badge-success" style="font-size:0.7rem; padding:0.15rem 0.4rem;">${t('markVerified')}</span>`;
        } else if (auxStub && auxStub.targetNew && auxStub.targetNew !== auxStub.targetOld) {
          auxBadge = `<span class="badge badge-warning" style="font-size:0.7rem; padding:0.15rem 0.4rem;">${t('statusDraft')}</span>`;
        } else if (auxStub && auxStub.classification === 'UNCHANGED') {
          auxBadge = `<span class="badge badge-neutral" style="font-size:0.7rem; padding:0.15rem 0.4rem;">${t('badgeUnchanged')}</span>`;
        } else if (auxText) {
          auxBadge = `<span class="badge badge-neutral" style="font-size:0.7rem; padding:0.15rem 0.4rem;">${t('statusOriginal')}</span>`;
        }

        colAuxHtml = `
          <div class="sub-panel aux-panel">
            <div class="sub-panel-title">
              <span>🌐 ${t('cardAuxPanelTitle', { LANG: auxLabel })}</span>
              ${auxBadge}
            </div>
            <div class="diff-box aux-diff-box">${esc(auxText) || `<span style="color:var(--text-muted); font-style:italic;">${t('emptyCell')}</span>`}</div>
            ${auxText ? `
            <div class="editor-toolbar" style="margin-top:0.35rem; justify-content:flex-end;">
              <button class="btn btn-sm btn-outline copy-aux-btn" data-index="${idx}" title="${t('copyAuxTitle')}">
                📋 ${t('copyAuxBtn')}
              </button>
            </div>` : ''}
          </div>
        `;
      }

      // 3. Old version of the target language
      const col2OldHtml = isNewLang ? '' : `
        <div class="sub-panel">
          <div class="sub-panel-title">
            <span>${t('cardTargetOldTitle', { LANG: langLabel })}</span>
          </div>
          <div class="diff-box" style="background-color:var(--bg-tertiary); color:var(--text-secondary);">${ComparatorEngine.escapeHtml(stub.targetOld) || `<span style="color:var(--text-muted); font-style:italic;">${t('emptyOriginalText')}</span>`}</div>
        </div>
      `;

      // 4. New version of the target language (editor)
      const editorTitle = isNewLang ? t('cardTargetNewSingleTitle', { LANG: langLabel }) : t('cardTargetNewTitle', { LANG: langLabel });
      const col3NewEditorHtml = `
        <div class="sub-panel">
          <div class="sub-panel-title">
            <span>${editorTitle}</span>
          </div>
          <textarea class="form-textarea target-textarea" data-index="${idx}" placeholder="${t('textareaPlaceholder')}">${ComparatorEngine.escapeHtml(stub.targetNew)}</textarea>
          <div class="editor-toolbar">
            <button class="btn btn-sm btn-primary copy-prompt-btn" data-index="${idx}">
              ${t('copyPromptBtn')}
            </button>
            <button class="btn btn-sm btn-outline copy-verify-prompt-btn" data-index="${idx}" title="${t('copyVerifyPromptTitle')}">
              ${t('copyVerifyPromptBtn')}
            </button>
            ${(!isNewLang && classification === 'MODIFIED') ? `<button class="btn btn-sm btn-outline prefill-old-btn" data-index="${idx}">${t('prefillOldBtn')}</button>` : ''}
            ${(classification === 'ADDED' || isNewLang) ? `<button class="btn btn-sm btn-outline prefill-ref-btn" data-index="${idx}">${t('prefillRefBtn')}</button>` : ''}
          </div>
        </div>
      `;

      // 5. Target Diff Live Preview
      const targetDiffHtml = renderTargetDiffHtml(stub);

      const col4DiffHtml = isNewLang ? '' : `
        <div class="sub-panel">
          <div class="sub-panel-title">
            <span>${t('cardTargetDiffTitle', { LANG: langLabel })}</span>
          </div>
          <div class="diff-box live-target-diff" id="target-diff-${idx}">${targetDiffHtml}</div>
        </div>
      `;

      // Card Header with optional search highlight
      const rowTitleHtml = highlightQuery(rowTitle, state.searchQuery);

      const headerHtml = `
        <div class="article-header">
          <div class="article-meta">
            <span class="article-id">${rowTitleHtml}</span>
            ${renderBadgeHtml(classification)}
          </div>
          <div>
            <label style="display:inline-flex; align-items:center; gap:0.4rem; cursor:pointer; font-size:0.85rem; font-weight:600; color:${stub.isVerified ? '#047857' : 'var(--text-muted)'};">
              <input type="checkbox" class="verify-checkbox" data-index="${idx}" ${stub.isVerified ? 'checked' : ''} style="accent-color:#10b981; width:16px; height:16px;">
              <span>${stub.isVerified ? t('markVerified') : t('markToVerify')}</span>
            </label>
          </div>
        </div>
      `;

      // Column 1: Reference diff (or clean text if no old reference / direct translation)
      const hasOldRef = !!stub.refOld;
      const isDirect = isDirectTranslationMode();
      const refTitle = isDirect ? t('cardSourceTextPanelTitle') : t('cardRefPanelTitle', { LANG: refLabel });
      const col1RefHtml = `
        <div class="sub-panel">
          <div class="sub-panel-title">
            <span>${refTitle}</span>
            ${(!hasOldRef || isDirect) ? '' : (refDiffRes.stats.hasChanges ? `<span style="font-size:0.75rem; color:#b45309; font-weight:600;">${t('changePct', { PCT: Math.round((1 - refDiffRes.stats.similarity) * 100) })}</span>` : `<span style="font-size:0.75rem; color:#15803d; font-weight:600;">${t('match100')}</span>`)}
          </div>
          <div class="diff-box">${((!hasOldRef || isDirect) ? ComparatorEngine.escapeHtml(stub.refNew) : refDiffHtml) || `<span style="color:var(--text-muted);">${t('emptyCell')}</span>`}</div>
        </div>
      `;

      // Full-Width 4-Columns or 2x2 Layout Grid (or 1-Column Layout for newly added language)
      const hasAux = !!colAuxHtml;
      const gridClass = isNewLang
        ? ('layout-new-lang' + (hasAux ? ' has-aux' : ''))
        : ((state.layoutMode === '2x2' ? 'layout-2x2' : 'layout-4cols') + (hasAux ? ' has-aux' : ''));

      card.innerHTML = `
        ${headerHtml}
        <div class="article-grid ${gridClass}">
          ${col1RefHtml}
          ${colAuxHtml}
          ${col2OldHtml}
          ${col3NewEditorHtml}
          ${col4DiffHtml}
        </div>
      `;

      elements.articlesContainer.appendChild(card);
    });

    if (renderedCount === 0) {
      elements.articlesContainer.innerHTML = '<div style="padding:3rem; text-align:center; color:var(--text-muted); background:var(--bg-secondary); border-radius:var(--radius-lg); border:1px solid var(--border-color);">' + t('noArticlesMatch') + '</div>';
    }

    attachCardEventListeners();
  }

  function attachCardEventListeners() {
    const col = state.activeTargetCol;
    if (!col) return;
    const stubs = state.translations[col];

    // Verification checkboxes
    elements.articlesContainer.querySelectorAll('.verify-checkbox').forEach(function (cb) {
      cb.addEventListener('change', function () {
        const idx = parseInt(cb.getAttribute('data-index'), 10);
        stubs[idx].isVerified = cb.checked;
        stubs[idx].status = cb.checked ? 'verified' : (stubs[idx].targetNew ? 'draft' : 'pending');

        const card = document.getElementById('article-card-' + idx);
        if (card) {
          if (cb.checked) {
            card.classList.remove('status-pending');
            card.classList.add('status-verified');
          } else {
            card.classList.remove('status-verified');
            card.classList.add('status-pending');
          }
        }

        updateLanguageProgress();
        renderLanguageTabs();
        renderCrossLanguageMatrix();
        persistState(150);
      });
    });

    // Textarea inputs + CAT shortcut (Ctrl + Enter)
    elements.articlesContainer.querySelectorAll('.target-textarea').forEach(function (textarea) {
      const idx = parseInt(textarea.getAttribute('data-index'), 10);

      // CAT keyboard shortcut: Ctrl+Enter (or Cmd+Enter) approves and moves to next pending
      textarea.addEventListener('keydown', function (e) {
        if ((e.ctrlKey || e.metaKey) && e.key === 'Enter') {
          e.preventDefault();
          stubs[idx].isVerified = true;
          stubs[idx].status = 'verified';

          const card = document.getElementById('article-card-' + idx);
          if (card) {
            card.classList.remove('status-pending');
            card.classList.add('status-verified');
            const cb = card.querySelector('.verify-checkbox');
            if (cb) cb.checked = true;
          }

          updateLanguageProgress();
          renderLanguageTabs();
          renderCrossLanguageMatrix();
          persistState(150);

          jumpToPending('next');
          showToast(t('markVerified') + ' ✓', 'success');
        }
      });

      textarea.addEventListener('input', function () {
        stubs[idx].targetNew = textarea.value;
        if (!stubs[idx].isVerified) {
          stubs[idx].status = textarea.value.trim() ? 'draft' : 'pending';
        }

        // Live update of visual diff in Column 4
        const liveDiffBox = document.getElementById('target-diff-' + idx);
        if (liveDiffBox) {
          liveDiffBox.innerHTML = renderTargetDiffHtml(stubs[idx]);
        }

        persistState(400);
      });
    });

    // Copy Prompt buttons
    elements.articlesContainer.querySelectorAll('.copy-prompt-btn').forEach(function (btn) {
      btn.addEventListener('click', function () {
        const idx = parseInt(btn.getAttribute('data-index'), 10);
        copyPromptForIndex(idx);
      });
    });

    // Copy Verification Prompt buttons
    elements.articlesContainer.querySelectorAll('.copy-verify-prompt-btn').forEach(function (btn) {
      btn.addEventListener('click', function () {
        const idx = parseInt(btn.getAttribute('data-index'), 10);
        copyVerifyPromptForIndex(idx);
      });
    });

    // Prefill with the old version
    elements.articlesContainer.querySelectorAll('.prefill-old-btn').forEach(function (btn) {
      btn.addEventListener('click', function () {
        const idx = parseInt(btn.getAttribute('data-index'), 10);
        const textarea = elements.articlesContainer.querySelector(`.target-textarea[data-index="${idx}"]`);
        if (textarea) {
          textarea.value = stubs[idx].targetOld;
          stubs[idx].targetNew = stubs[idx].targetOld;
          textarea.dispatchEvent(new Event('input'));
          showToast(t('prefillOldBtn'), 'info');
        }
      });
    });

    // Prefill with the reference text
    elements.articlesContainer.querySelectorAll('.prefill-ref-btn').forEach(function (btn) {
      btn.addEventListener('click', function () {
        const idx = parseInt(btn.getAttribute('data-index'), 10);
        const textarea = elements.articlesContainer.querySelector(`.target-textarea[data-index="${idx}"]`);
        if (textarea) {
          textarea.value = stubs[idx].refNew;
          stubs[idx].targetNew = stubs[idx].refNew;
          textarea.dispatchEvent(new Event('input'));
          showToast(t('prefillRefBtn'), 'info');
        }
      });
    });

    // Copy Auxiliary text into the editor
    elements.articlesContainer.querySelectorAll('.copy-aux-btn').forEach(function (btn) {
      btn.addEventListener('click', function () {
        const idx = parseInt(btn.getAttribute('data-index'), 10);
        const auxCol = state.auxiliaryCol;
        if (!auxCol || !state.translations[auxCol]) return;
        const auxStub = state.translations[auxCol][idx];
        const auxText = (auxStub && auxStub.targetNew) ? auxStub.targetNew : (auxStub ? auxStub.targetOld : (state.rows[idx] ? state.rows[idx][auxCol] : ''));
        if (!auxText) return;
        const textarea = elements.articlesContainer.querySelector(`.target-textarea[data-index="${idx}"]`);
        if (textarea) {
          textarea.value = auxText;
          stubs[idx].targetNew = auxText;
          textarea.dispatchEvent(new Event('input', { bubbles: true }));
          showToast(t('toastAuxCopied'), 'info');
        }
      });
    });
  }

  function copyPromptForIndex(idx) {
    const col = state.activeTargetCol;
    const stub = state.translations[col][idx];
    const refDiffRes = ComparatorEngine.computeDiff(stub.refOld, stub.refNew);

    let auxText = '';
    let auxLangName = '';
    if (state.auxiliaryCol && state.translations[state.auxiliaryCol]) {
      const auxStub = state.translations[state.auxiliaryCol][idx];
      auxText = (auxStub && auxStub.targetNew) ? auxStub.targetNew : (auxStub ? auxStub.targetOld : (state.rows[idx] ? state.rows[idx][state.auxiliaryCol] : ''));
      auxLangName = languageLabel(state.auxiliaryCol);
    }

    const isNew = ComparatorEngine.isNewLanguage(state.mappings, col);
    const promptText = ComparatorEngine.buildCopilotPrompt({
      promptLang: I18n.getLanguage(),
      targetLangName: languageLabel(col),
      refLangName: columnDisplayName(state.mappings.refNewCol),
      auxiliaryLangName: auxLangName,
      auxiliaryText: auxText,
      rowId: ComparatorEngine.rowTag(state.rowNumbers, idx),
      classification: stub.classification,
      refOld: stub.refOld,
      refNew: stub.refNew,
      targetOld: stub.targetOld,
      diffResult: refDiffRes,
      customTemplate: state.customPromptTemplate,
      isNewLanguage: isNew
    });

    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(promptText).then(function () {
        showToast(t('toastPromptCopied'), 'success');
      }).catch(function () {
        fallbackCopyText(promptText);
      });
    } else {
      fallbackCopyText(promptText);
    }
  }

  function copyVerifyPromptForIndex(idx) {
    const col = state.activeTargetCol;
    if (!col || !state.translations[col]) return;
    const stub = state.translations[col][idx];
    if (!stub) return;

    const currentText = (stub.targetNew !== undefined && stub.targetNew !== null) ? stub.targetNew.trim() : '';
    if (!currentText) {
      showToast(t('toastNoTranslationToVerify'), 'warning');
      return;
    }

    let auxText = '';
    let auxLangName = '';
    if (state.auxiliaryCol && state.translations[state.auxiliaryCol]) {
      const auxStub = state.translations[state.auxiliaryCol][idx];
      auxText = (auxStub && auxStub.targetNew) ? auxStub.targetNew : (auxStub ? auxStub.targetOld : (state.rows[idx] ? state.rows[idx][state.auxiliaryCol] : ''));
      auxLangName = languageLabel(state.auxiliaryCol);
    }

    const promptText = ComparatorEngine.buildVerificationPrompt({
      promptLang: I18n.getLanguage(),
      targetLangName: languageLabel(col),
      refLangName: columnDisplayName(state.mappings.refNewCol),
      auxiliaryLangName: auxLangName,
      auxiliaryText: auxText,
      rowId: ComparatorEngine.rowTag(state.rowNumbers, idx),
      refNew: stub.refNew,
      targetNew: currentText
    });

    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(promptText).then(function () {
        showToast(t('toastVerifyPromptCopied'), 'success');
      }).catch(function () {
        fallbackCopyText(promptText, t('toastVerifyPromptCopied'));
      });
    } else {
      fallbackCopyText(promptText, t('toastVerifyPromptCopied'));
    }
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

  /**
   * Filters and Search Handling
   */
  function initFilters() {
    elements.filterChipsContainer.querySelectorAll('.filter-chip').forEach(function (chip) {
      chip.addEventListener('click', function () {
        elements.filterChipsContainer.querySelectorAll('.filter-chip').forEach(function (c) {
          c.classList.remove('active');
        });
        chip.classList.add('active');
        state.filterStatus = chip.getAttribute('data-status');
        renderArticleCards();
      });
    });

    // Re-rendering recomputes two diffs per card, so do it once typing pauses
    let searchTimer = null;
    elements.searchInput.addEventListener('input', function (e) {
      const value = e.target.value.trim();
      clearTimeout(searchTimer);
      searchTimer = setTimeout(function () {
        if (value === state.searchQuery) return;
        state.searchQuery = value;
        renderArticleCards();
      }, 200);
    });

    elements.nextPendingBtn.addEventListener('click', function () {
      jumpToPending('next');
    });

    elements.prevPendingBtn.addEventListener('click', function () {
      jumpToPending('prev');
    });

    document.addEventListener('keydown', function (e) {
      if (state.currentStep !== 3) return;
      if (e.altKey && e.key === 'ArrowDown') {
        e.preventDefault();
        jumpToPending('next');
      } else if (e.altKey && e.key === 'ArrowUp') {
        e.preventDefault();
        jumpToPending('prev');
      }
    });
  }

  function jumpToPending(direction) {
    const col = state.activeTargetCol;
    if (!col || !state.translations[col]) return;

    const cards = Array.from(elements.articlesContainer.querySelectorAll('.article-card'));
    if (cards.length === 0) return;

    const pendingCards = cards.filter(function (card) {
      return card.classList.contains('status-pending');
    });

    if (pendingCards.length === 0) {
      showToast(t('toastAllApproved'), 'success');
      return;
    }

    let targetCard = null;

    if (direction === 'next') {
      targetCard = pendingCards.find(function (c) {
        return c.getBoundingClientRect().top > 120;
      }) || pendingCards[0];
    } else {
      const reversed = pendingCards.slice().reverse();
      targetCard = reversed.find(function (c) {
        return c.getBoundingClientRect().top < 50;
      }) || pendingCards[pendingCards.length - 1];
    }

    if (targetCard) {
      if (typeof targetCard.scrollIntoView === 'function') {
        targetCard.scrollIntoView({ behavior: 'smooth', block: 'center' });
      }
      const ta = targetCard.querySelector('.target-textarea');
      if (ta) ta.focus();
    }
  }

  /**
   * Step 4: Export Logic
   * One language selector drives the printable document, the change summary and the partial
   * workbook: the reference (new version) or any of the languages.
   */
  function documentBaseName() {
    return (state.fileName || 'dokument').replace(/\.[^/.]+$/, '');
  }

  function updatePdfLanguageDropdown() {
    if (!elements.pdfLangSelect) return;
    const select = elements.pdfLangSelect;
    select.innerHTML = '';

    if (state.mappings.refNewCol) {
      const optRef = document.createElement('option');
      optRef.value = state.mappings.refNewCol;
      optRef.textContent = isDirectTranslationMode()
        ? columnDisplayName(state.mappings.refNewCol)
        : `${columnDisplayName(state.mappings.refNewCol)} (${t('pdfOptionRefNew')})`;
      select.appendChild(optRef);
    }

    state.mappings.targetCols.forEach(function (col) {
      const opt = document.createElement('option');
      opt.value = col;
      opt.textContent = languageLabel(col);
      select.appendChild(opt);
    });

    const valid = function (col) {
      return col && (col === state.mappings.refNewCol || state.mappings.targetCols.indexOf(col) !== -1);
    };
    select.value = valid(state.exportLangCol) ? state.exportLangCol
      : (valid(state.activeTargetCol) ? state.activeTargetCol : (select.options[0] ? select.options[0].value : ''));
    state.exportLangCol = select.value;
    renderPartialContextOptions();
  }

  function selectedExportLang() {
    const col = elements.pdfLangSelect ? elements.pdfLangSelect.value : state.exportLangCol;
    if (!col) showToast(t('toastPdfSelectLang'), 'warning');
    return col;
  }

  /**
   * Columns that accompany the selected language in the partial workbook, in two groups:
   * new versions first (the reference's new version and the current translation work of the
   * other languages), then the original columns of the sheet. Default: old and new reference.
   * Entries are column names, or { newOf: <old-version column> } for a language's new version.
   */
  function contextKey(entry) {
    return typeof entry === 'string' ? 'col:' + entry : 'new:' + entry.newOf;
  }

  function renderPartialContextOptions() {
    const container = elements.partialContextContainer;
    if (!container) return;
    const m = state.mappings;
    const langCol = state.exportLangCol;
    const isReference = langCol === m.refNewCol;
    if (!Array.isArray(state.partialContextCols)) {
      state.partialContextCols = [m.refOldCol, m.refNewCol].filter(Boolean);
    }

    const languageNewCols = m.targetCols.map(function (c) { return m.languages[c] && m.languages[c].newCol; }).filter(Boolean);
    const newVersions = [];
    if (!isReference && m.refNewCol) {
      newVersions.push({ entry: m.refNewCol, label: columnDisplayName(m.refNewCol) });
    }
    m.targetCols.forEach(function (c) {
      if (c === langCol) return;
      const cfg = m.languages[c] || {};
      const colName = cfg.newCol || cfg.newColName || defaultNewColNameFor(c);
      newVersions.push({ entry: { newOf: c }, label: columnDisplayName(colName) + ' (' + languageLabel(c) + ')' });
    });

    const own = isReference ? [m.refOldCol, m.refNewCol] : [langCol, m.languages[langCol] && m.languages[langCol].newCol];
    const originals = state.headers
      .filter(function (h) {
        return own.indexOf(h) === -1 && h !== m.refNewCol && languageNewCols.indexOf(h) === -1;
      })
      .map(function (h) { return { entry: h, label: columnDisplayName(h) }; });

    const selectedKeys = state.partialContextCols.map(contextKey);
    const displayOrder = newVersions.concat(originals).map(function (o) { return contextKey(o.entry); });
    container.innerHTML = '';
    [[t('partialGroupNew'), newVersions], [t('partialGroupOriginal'), originals]].forEach(function (group) {
      if (group[1].length === 0) return;
      const box = document.createElement('div');
      box.className = 'partial-context-group';
      const title = document.createElement('div');
      title.className = 'partial-context-title';
      title.textContent = group[0];
      const items = document.createElement('div');
      items.className = 'partial-context-items';
      group[1].forEach(function (option) {
        const key = contextKey(option.entry);
        const label = document.createElement('label');
        label.className = 'checkbox-label';
        const cb = document.createElement('input');
        cb.type = 'checkbox';
        cb.value = key;
        cb.checked = selectedKeys.indexOf(key) !== -1;
        cb.addEventListener('change', function () {
          const list = state.partialContextCols;
          const pos = list.map(contextKey).indexOf(key);
          if (cb.checked && pos === -1) list.push(option.entry);
          if (!cb.checked && pos !== -1) list.splice(pos, 1);
          // Exported in the order shown (new versions first); entries not offered for the current
          // language (e.g. its own columns) keep their place at the end
          const rank = function (entry) {
            const i = displayOrder.indexOf(contextKey(entry));
            return i === -1 ? displayOrder.length : i;
          };
          list.sort(function (x, y) { return rank(x) - rank(y); });
          persistState(300);
        });
        label.appendChild(cb);
        label.appendChild(document.createTextNode(option.label));
        items.appendChild(label);
      });
      box.appendChild(title);
      box.appendChild(items);
      container.appendChild(box);
    });
  }

  /**
   * Change summary (Markdown, Word .doc, PDF) of one language: from Step 2 always the reference,
   * from Step 4 the language selected in the export selector.
   */
  function exportManagerialSummary(format, langCol) {
    if (!state.rows || state.rows.length === 0 || getMappingError(false)) {
      showToast(t(state.rows.length ? getMappingError(false) : 'toastNoData'), 'warning');
      return;
    }
    const col = langCol || state.mappings.refNewCol;
    if (col !== state.mappings.refNewCol) ensureStubsInitialized();

    const summary = ComparatorEngine.buildChangeSummary({
      rows: state.rows,
      mappings: state.mappings,
      translations: state.translations,
      rowNumbers: state.rowNumbers,
      langCol: col,
      documentName: documentBaseName(),
      uiLang: I18n.getLanguage()
    });
    const adapterId = { markdown: 'summary-markdown', word: 'summary-word', pdf: 'summary-print' }[format] || format;
    const baseName = documentBaseName() + '_zmeny_' + summary.language;
    Exporters.run(adapterId, summary, { baseName: baseName, uiLang: I18n.getLanguage() });

    if (format === 'markdown') {
      showToast(t('mgrExportMdSuccess'), 'success');
    } else if (format === 'word') {
      showToast(t('mgrExportDocSuccess'), 'success');
    }
  }

  function exportDocument(adapterId, col) {
    const model = ComparatorEngine.buildDocumentModel({
      rows: state.rows,
      mappings: state.mappings,
      translations: state.translations,
      rowNumbers: state.rowNumbers,
      targetCol: col,
      documentName: documentBaseName(),
      uiLang: I18n.getLanguage(),
      template: state.pdfTemplate
    });
    return Exporters.run(adapterId, model, { baseName: documentBaseName() + '_' + model.language, uiLang: I18n.getLanguage() });
  }

  function initExports() {
    elements.downloadComprehensiveBtn.addEventListener('click', function () {
      exportTables('comprehensive');
    });

    elements.downloadCleanBtn.addEventListener('click', function () {
      exportTables('clean');
    });

    elements.downloadPartialBtn.addEventListener('click', function () {
      const col = selectedExportLang();
      if (col) exportTables('language', col);
    });

    elements.exportProjectJsonBtn.addEventListener('click', function () {
      exportProjectFile();
    });

    elements.tableFormatSelect.addEventListener('change', function () {
      state.tableExportFormat = elements.tableFormatSelect.value;
    });

    elements.pdfLangSelect.addEventListener('change', function () {
      state.exportLangCol = elements.pdfLangSelect.value;
      renderPartialContextOptions();
      persistState(300);
    });

    // Change summary: Step 2 = reference, Step 4 = selected language
    elements.step2ExportMdBtn.addEventListener('click', function () { exportManagerialSummary('markdown'); });
    elements.step2ExportDocBtn.addEventListener('click', function () { exportManagerialSummary('word'); });
    elements.step2ExportPdfBtn.addEventListener('click', function () { exportManagerialSummary('pdf'); });
    elements.downloadMgrMdBtn.addEventListener('click', function () {
      const col = selectedExportLang();
      if (col) exportManagerialSummary('markdown', col);
    });
    elements.downloadMgrDocBtn.addEventListener('click', function () {
      const col = selectedExportLang();
      if (col) exportManagerialSummary('word', col);
    });
    elements.printMgrPdfBtn.addEventListener('click', function () {
      const col = selectedExportLang();
      if (col) exportManagerialSummary('pdf', col);
    });

    // Printable document of the selected language
    elements.printPdfBtn.addEventListener('click', function () {
      const col = selectedExportLang();
      if (col) exportDocument('html-print', col);
    });

    elements.downloadPdfHtmlBtn.addEventListener('click', function () {
      const col = selectedExportLang();
      if (!col) return;
      exportDocument('html-file', col);
      showToast(t('toastPdfHtmlDownloaded'), 'success');
    });

    // Print all languages
    elements.printAllPdfBtn.addEventListener('click', function () {
      const allCols = [];
      if (state.mappings.refNewCol) allCols.push(state.mappings.refNewCol);
      allCols.push.apply(allCols, state.mappings.targetCols);

      if (allCols.length === 0) {
        showToast(t('toastNoLanguages'), 'warning');
        return;
      }

      if (confirm(t('confirmPrintAll', { COUNT: allCols.length }))) {
        allCols.forEach(function (col, i) {
          setTimeout(function () {
            exportDocument('html-print', col);
          }, i * 400);
        });
      }
    });
  }

  /**
   * Batch Prompt Generation Modal
   */
  function initBatchPromptModal() {
    if (elements.batchPromptBtn) {
      elements.batchPromptBtn.addEventListener('click', function () {
        openBatchPromptModal();
      });
    }

    function closeModal() {
      if (elements.batchPromptModal) elements.batchPromptModal.classList.remove('open');
    }

    if (elements.closeBatchPromptBtn) elements.closeBatchPromptBtn.addEventListener('click', closeModal);
    if (elements.closeBatchPromptFooterBtn) elements.closeBatchPromptFooterBtn.addEventListener('click', closeModal);

    if (elements.copyBatchPromptBtn) {
      elements.copyBatchPromptBtn.addEventListener('click', function () {
        const text = elements.batchPromptTextarea.value;
        if (navigator.clipboard && navigator.clipboard.writeText) {
          navigator.clipboard.writeText(text).then(function () {
            showToast(t('batchPromptCopied'), 'success');
          }).catch(function () {
            fallbackCopyText(text);
          });
        } else {
          fallbackCopyText(text);
        }
      });
    }
  }

  function openBatchPromptModal() {
    const col = state.activeTargetCol;
    if (!col || !state.translations[col]) {
      showToast(t('toastSelectTargetLang'), 'warning');
      return;
    }

    const isNew = ComparatorEngine.isNewLanguage(state.mappings, col);
    const promptText = ComparatorEngine.buildBatchCopilotPrompt({
      promptLang: I18n.getLanguage(),
      targetCol: col,
      targetLangName: languageLabel(col),
      refLangName: columnDisplayName(state.mappings.refNewCol),
      auxiliaryCol: state.auxiliaryCol || '',
      auxiliaryLangName: state.auxiliaryCol ? languageLabel(state.auxiliaryCol) : '',
      rows: state.rows,
      rowNumbers: state.rowNumbers,
      mappings: state.mappings,
      translations: state.translations,
      isNewLanguage: isNew
    });

    const descEl = elements.batchPromptModal ? elements.batchPromptModal.querySelector('.modal-body p') : null;
    if (descEl) {
      descEl.textContent = isNew ? t('batchPromptDescNew') : t('batchPromptDesc');
    }

    elements.batchPromptTextarea.value = promptText;
    elements.batchPromptModal.classList.add('open');
  }

  /**
   * Smart Multi-Paste Modal
   */
  let smartParsedMatches = [];

  function initSmartPasteModal() {
    if (elements.smartPasteBtn) {
      elements.smartPasteBtn.addEventListener('click', function () {
        smartParsedMatches = [];
        elements.smartPasteInput.value = '';
        elements.smartPasteMatchesContainer.style.display = 'none';
        elements.smartPasteMatchesList.innerHTML = '';
        elements.smartPasteApplyBtn.disabled = true;
        elements.smartPasteModal.classList.add('open');
      });
    }

    function closeModal() {
      if (elements.smartPasteModal) elements.smartPasteModal.classList.remove('open');
    }

    if (elements.closeSmartPasteBtn) elements.closeSmartPasteBtn.addEventListener('click', closeModal);
    if (elements.closeSmartPasteFooterBtn) elements.closeSmartPasteFooterBtn.addEventListener('click', closeModal);

    if (elements.smartPasteParseBtn) {
      elements.smartPasteParseBtn.addEventListener('click', function () {
        const text = elements.smartPasteInput.value.trim();
        if (!text) {
          showToast(t('smartPasteNoMatches'), 'warning');
          return;
        }

        const col = state.activeTargetCol;
        const stubs = state.translations[col] || [];
        // Paragraphs are identified by their Excel row number ("#12")
        const expectedList = stubs.map(function (stub, idx) {
          return { index: idx, id: ComparatorEngine.rowTag(state.rowNumbers, idx) };
        });

        const parsed = ComparatorEngine.parseBatchTranslationResponse(text, expectedList);
        smartParsedMatches = parsed;
        const unmatchedHtml = parsed.unmatched && parsed.unmatched.length > 0
          ? `<div class="smart-paste-unmatched">⚠ ${ComparatorEngine.escapeHtml(t('smartPasteUnmatched', { LIST: parsed.unmatched.join(', ') }))}</div>`
          : '';

        if (parsed.length === 0) {
          elements.smartPasteMatchesContainer.style.display = 'block';
          elements.smartPasteMatchesTitle.textContent = '';
          elements.smartPasteMatchesList.innerHTML = `<span style="color:var(--text-muted); font-size:0.85rem;">${t('smartPasteNoMatches')}</span>` + unmatchedHtml;
          elements.smartPasteApplyBtn.disabled = true;
          return;
        }

        elements.smartPasteMatchesContainer.style.display = 'block';
        elements.smartPasteMatchesTitle.textContent = `${t('smartPasteParsedCount')} (${parsed.length}):`;
        elements.smartPasteMatchesList.innerHTML = parsed.map(function (m) {
          return `<span class="matched-article-badge">✓ ${ComparatorEngine.escapeHtml(m.id)} (${m.text.length} ${t('charsLabel')})</span>`;
        }).join('') + unmatchedHtml;

        elements.smartPasteApplyBtn.disabled = false;
      });
    }

    if (elements.smartPasteApplyBtn) {
      elements.smartPasteApplyBtn.addEventListener('click', function () {
        if (!smartParsedMatches || smartParsedMatches.length === 0) return;

        const col = state.activeTargetCol;
        const stubs = state.translations[col];
        const autoVerify = elements.smartPasteAutoVerifyCheckbox ? elements.smartPasteAutoVerifyCheckbox.checked : true;
        let appliedCount = 0;

        smartParsedMatches.forEach(function (item) {
          if (stubs[item.index]) {
            stubs[item.index].targetNew = item.text;
            if (autoVerify) {
              stubs[item.index].isVerified = true;
              stubs[item.index].status = 'verified';
            } else if (!stubs[item.index].isVerified) {
              stubs[item.index].status = 'draft';
            }
            appliedCount++;
          }
        });
        smartParsedMatches = [];
        elements.smartPasteApplyBtn.disabled = true;

        closeModal();
        renderArticleCards();
        updateLanguageProgress();
        renderLanguageTabs();
        renderCrossLanguageMatrix();
        persistState(150);

        showToast(t('smartPasteApplied', { COUNT: appliedCount }), 'success');
      });
    }
  }

  /**
   * Translations read back from a workbook (a partial workbook filled in by a translator): texts
   * are matched by row number and the language's new-version column, shown for confirmation and
   * taken over as "awaiting review".
   */
  function initTranslationImport() {
    const input = elements.translationsFileInput;
    if (!input) return;
    [elements.importTranslationsBtn, elements.importTranslationsExportBtn].forEach(function (btn) {
      if (!btn) return;
      btn.addEventListener('click', function () {
        if (!state.rows || state.rows.length === 0) {
          showToast(t('toastNoData'), 'warning');
          return;
        }
        const mappingError = getMappingError(true);
        if (mappingError) {
          showToast(t(mappingError), 'warning');
          return;
        }
        input.setAttribute('accept', Importers.acceptAttribute());
        input.click();
      });
    });
    input.addEventListener('change', function (e) {
      if (e.target.files && e.target.files[0]) importTranslationsFile(e.target.files[0]);
      e.target.value = '';
    });
  }

  function importTranslationsFile(file) {
    ensureStubsInitialized();
    return Importers.importFile(file).then(function (result) {
      // The first table of the file with a row-number column and a language column is used
      let match = null;
      let error = '';
      result.source.sheetNames.some(function (name) {
        const table = Importers.readSheet(result.source, name, { hasHeaderRow: true });
        if (!table) return false;
        const m = ComparatorEngine.matchImportedTranslations({
          table: table,
          rows: state.rows,
          rowNumbers: state.rowNumbers,
          mappings: state.mappings,
          translations: state.translations
        });
        if (!m.error) match = m;
        else if (error !== 'NO_LANGUAGE_COLUMN') error = m.error;
        return !!match;
      });

      if (!match) {
        if (error === 'NO_LANGUAGE_COLUMN') {
          const names = [];
          state.mappings.targetCols.forEach(function (col) {
            const cfg = state.mappings.languages[col] || {};
            [cfg.newCol, cfg.newColName, languageLabel(col)].forEach(function (n) {
              if (n && names.indexOf(n) === -1) names.push(n);
            });
          });
          alert(t('importTranslationsNoLangCol', { NAME: file.name, LIST: names.join(', ') }));
        } else {
          alert(t('importTranslationsNoRowCol', { NAME: file.name }));
        }
        return 0;
      }

      const withUpdates = match.languages.filter(function (l) { return l.updates.length > 0; });
      if (withUpdates.length === 0) {
        showToast(t('importTranslationsNothing', { NAME: file.name }), 'info');
        return 0;
      }
      const lines = match.languages.map(function (l) {
        return t('importTranslationsLine', { LANG: l.label, COL: l.column, COUNT: l.updates.length, SAME: l.unchanged });
      });
      if (match.unknownRows.length > 0) {
        const shown = match.unknownRows.slice(0, 20).join(', ') + (match.unknownRows.length > 20 ? ', …' : '');
        lines.push(t('importTranslationsUnknownRows', { LIST: shown }));
      }
      if (!confirm(t('confirmImportTranslations', { NAME: file.name, LIST: lines.join('\n') }))) return 0;

      let count = 0;
      withUpdates.forEach(function (l) {
        const stubs = state.translations[l.col] || [];
        l.updates.forEach(function (u) {
          const stub = stubs[u.index];
          if (!stub) return;
          stub.targetNew = u.text;
          stub.isVerified = false;
          stub.status = 'draft';
          count++;
        });
      });
      renderTranslationStudio();
      persistState(150);
      showToast(t('toastTranslationsImported', { COUNT: count }), 'success');
      return count;
    }).catch(function (err) {
      console.error('Import of translations failed:', err);
      if (err && err.code === 'UNSUPPORTED') {
        alert(t('alertUnsupportedFile', { NAME: file.name, FORMATS: Importers.acceptAttribute().split(',').join(', ') }));
      } else {
        alert(t('alertFileReadError', { MSG: err && err.message ? err.message : file.name }));
      }
      return 0;
    });
  }

  /**
   * PDF Template Customization Modal
   */
  function initPdfTemplateSettings() {
    if (!elements.editPdfTemplateBtn) return;

    elements.editPdfTemplateBtn.addEventListener('click', function () {
      openPdfTemplateModal();
    });

    function closeModal() {
      if (elements.pdfTemplateModal) elements.pdfTemplateModal.classList.remove('open');
    }

    if (elements.closePdfTemplateBtn) elements.closePdfTemplateBtn.addEventListener('click', closeModal);

    // Two-way sync for color picker
    if (elements.tplAccentColorPicker && elements.tplAccentColor) {
      elements.tplAccentColorPicker.addEventListener('input', function () {
        elements.tplAccentColor.value = elements.tplAccentColorPicker.value;
      });
      elements.tplAccentColor.addEventListener('input', function () {
        elements.tplAccentColorPicker.value = elements.tplAccentColor.value;
      });
    }

    // Toggle signatures inputs
    if (elements.tplShowSignatures && elements.signaturesBlockInputs) {
      elements.tplShowSignatures.addEventListener('change', function () {
        elements.signaturesBlockInputs.style.display = elements.tplShowSignatures.checked ? 'grid' : 'none';
      });
    }

    // Save Template
    if (elements.savePdfTemplateBtn) {
      elements.savePdfTemplateBtn.addEventListener('click', function () {
        state.pdfTemplate = {
          companyName: elements.tplCompanyName.value.trim(),
          effectiveDate: elements.tplEffectiveDate.value.trim(),
          docTitle: elements.tplDocTitle.value.trim(),
          accentColor: /^#[0-9a-f]{3,8}$/i.test(elements.tplAccentColor.value.trim()) ? elements.tplAccentColor.value.trim() : '#2563eb',
          docSubtitle: elements.tplDocSubtitle.value.trim(),
          versionLabel: elements.tplVersionLabel ? elements.tplVersionLabel.value.trim() : '',
          introPreamble: elements.tplIntroPreamble.value.trim(),
          showSignatures: elements.tplShowSignatures.checked,
          signDept1: elements.tplSignDept1.value.trim(),
          signDept2: elements.tplSignDept2.value.trim()
        };

        persistState(100);
        closeModal();
        showToast(t('pdfTemplateSaved'), 'success');
      });
    }

    // Reset Template
    if (elements.resetPdfTemplateBtn) {
      elements.resetPdfTemplateBtn.addEventListener('click', function () {
        state.pdfTemplate = null;
        persistState(100);
        openPdfTemplateModal();
        showToast(t('pdfTemplateReset'), 'info');
      });
    }
  }

  function openPdfTemplateModal() {
    const tpl = state.pdfTemplate || {};
    const value = function (v, fallback) { return v !== undefined && v !== null ? v : fallback; };

    elements.tplCompanyName.value = value(tpl.companyName, '');
    elements.tplEffectiveDate.value = value(tpl.effectiveDate, '');
    elements.tplDocTitle.value = value(tpl.docTitle, documentBaseName());
    elements.tplAccentColor.value = tpl.accentColor || '#2563eb';
    elements.tplAccentColorPicker.value = tpl.accentColor || '#2563eb';
    elements.tplDocSubtitle.value = value(tpl.docSubtitle, '{LANG}');
    elements.tplVersionLabel.value = value(tpl.versionLabel, '');
    elements.tplIntroPreamble.value = value(tpl.introPreamble, '');
    elements.tplShowSignatures.checked = tpl.showSignatures === true;
    elements.signaturesBlockInputs.style.display = elements.tplShowSignatures.checked ? 'grid' : 'none';
    elements.tplSignDept1.value = value(tpl.signDept1, '');
    elements.tplSignDept2.value = value(tpl.signDept2, '');

    elements.pdfTemplateModal.classList.add('open');
  }

  /**
   * Table exports (full, new version only, one language) through the selected table exporter
   * (default: Excel, see js/io/excel-exporter.js).
   */
  function exportTables(exportType, langCol) {
    if (!state.rows || state.rows.length === 0) {
      showToast(t('toastNoData'), 'warning');
      return null;
    }
    const mappingError = getMappingError(true);
    if (mappingError) {
      showToast(t(mappingError), 'warning');
      return null;
    }
    ensureStubsInitialized();

    const result = ComparatorEngine.prepareExportTables({
      rows: state.rows,
      headers: state.headers,
      rowNumbers: state.rowNumbers,
      mappings: state.mappings,
      translations: state.translations,
      exportType: exportType,
      langCol: langCol,
      contextCols: state.partialContextCols || [state.mappings.refOldCol, state.mappings.refNewCol].filter(Boolean),
      skippedRows: state.skippedRows,
      uiLang: I18n.getLanguage()
    });

    const suffix = exportType === 'clean' ? 'nova_verze'
      : (exportType === 'language' ? (langCol === state.mappings.refNewCol ? columnDisplayName(langCol) : languageLabel(langCol)) : 'srovnani');
    const adapter = Exporters.get(state.tableExportFormat) || Exporters.getDefault('table');
    const outcome = Exporters.run(adapter.id, { tables: result.tables }, {
      baseName: documentBaseName() + '_' + suffix,
      uiLang: I18n.getLanguage()
    });

    if (outcome.fileName) showToast(t('toastWorkbookGenerated', { NAME: outcome.fileName }), 'success');
    if (result.unverifiedCount > 0) {
      const key = exportType === 'clean' ? 'toastExportUnverifiedClean'
        : (exportType === 'language' ? 'toastExportUnverifiedLanguage' : 'toastExportUnverified');
      showToast(t(key, { COUNT: result.unverifiedCount }), 'warning');
    }
    return outcome.fileName || null;
  }

  /**
   * Format selector for table exports; shown only when more than one table exporter is registered.
   */
  function renderTableFormatOptions() {
    const select = elements.tableFormatSelect;
    if (!select) return;
    const adapters = Exporters.list('table');
    if (!Exporters.get(state.tableExportFormat)) state.tableExportFormat = Exporters.getDefault('table').id;
    select.innerHTML = '';
    adapters.forEach(function (a) {
      const opt = document.createElement('option');
      opt.value = a.id;
      opt.textContent = a.label;
      select.appendChild(opt);
    });
    select.value = state.tableExportFormat;
    elements.tableFormatWrapper.style.display = adapters.length > 1 ? 'flex' : 'none';
  }

  function exportProjectFile() {
    if (!state.rows || state.rows.length === 0) {
      showToast(t('toastNoData'), 'warning');
      return null;
    }
    // Make sure the file contains the latest edits and stubs
    if (state.mappings.targetCols.length > 0 && !getMappingError(true)) {
      ensureStubsInitialized();
    }
    const fileName = StorageManager.exportProjectAsJson(state);
    showToast(t('toastProjectSaved', { NAME: fileName }), 'success');
    return fileName;
  }

  /**
   * In-Browser Unit Diagnostics Runner
   */
  function initDiagnostics() {
    elements.diagnosticsBtn.addEventListener('click', function () {
      elements.diagnosticsModal.classList.add('open');
      runBrowserTests();
    });

    elements.closeDiagnosticsBtn.addEventListener('click', function () {
      elements.diagnosticsModal.classList.remove('open');
    });

    elements.runDiagnosticsBtn.addEventListener('click', function () {
      runBrowserTests();
    });
  }

  function runBrowserTests() {
    elements.diagnosticsLog.innerHTML = `<div style="color:var(--text-muted);">${I18n.getLanguage() === 'en' ? 'Running diagnostic suite...' : 'Spouštím interní diagnostickou sadu...'}</div>`;

    setTimeout(function () {
      const logs = [];
      let pass = 0;
      let fail = 0;

      function assertEq(name, actual, expected) {
        if (actual === expected) {
          logs.push(`<div style="color:#15803d; font-weight:600;">✓ ${ComparatorEngine.escapeHtml(name)}</div>`);
          pass++;
        } else {
          logs.push(`<div style="color:#b91c1c; font-weight:bold;">✗ ${ComparatorEngine.escapeHtml(name)}<br><small style="color:var(--text-muted); padding-left:1.5rem;">Expected: ${expected}, Got: ${actual}</small></div>`);
          fail++;
        }
      }

      function assertOk(name, cond, details) {
        if (cond) {
          logs.push(`<div style="color:#15803d; font-weight:600;">✓ ${ComparatorEngine.escapeHtml(name)}</div>`);
          pass++;
        } else {
          logs.push(`<div style="color:#b91c1c; font-weight:bold;">✗ ${ComparatorEngine.escapeHtml(name)} ${details ? '<br><small>' + details + '</small>' : ''}</div>`);
          fail++;
        }
      }

      // 1. Classification
      assertEq('classifyRow(empty, empty) -> EMPTY', ComparatorEngine.classifyRow('', ''), 'EMPTY');
      assertEq('classifyRow(empty, text) -> ADDED', ComparatorEngine.classifyRow('', 'Nový text'), 'ADDED');
      assertEq('classifyRow(text, empty) -> REMOVED', ComparatorEngine.classifyRow('Starý text', ''), 'REMOVED');
      assertEq('classifyRow(text, text) -> UNCHANGED', ComparatorEngine.classifyRow('Etika 2024', 'Etika 2024'), 'UNCHANGED');
      assertEq('classifyRow(diff, diff) -> MODIFIED', ComparatorEngine.classifyRow('14 dnů', '7 dnů'), 'MODIFIED');

      // 2. Diffing
      const diffRes = ComparatorEngine.computeDiff('Lhůta 14 dnů.', 'Lhůta 7 dnů.');
      assertOk('computeDiff detects changes', diffRes.stats.hasChanges);
      assertOk('computeDiff contains removed token 14', diffRes.summary.removedTerms.includes('14'));
      assertOk('computeDiff contains added token 7', diffRes.summary.addedTerms.includes('7'));

      // 3. HTML Escaping
      const safeDiff = ComparatorEngine.renderDiffHtml(ComparatorEngine.computeDiff('<b>test</b>', '<i>test</i>'), 'inline');
      assertOk('renderDiffHtml sanitizes HTML tags', !safeDiff.includes('<script>') && safeDiff.includes('&lt;'));

      // 4. Stubs
      const sampleRows = SampleGenerator.getSampleData();
      const stubs = ComparatorEngine.generateTargetStub(sampleRows, 'CS 2024', 'DE 2024 (Referenz)', 'DE 2026 (Referenz)');
      assertEq('Stub generation for 15 sample paragraphs', stubs.length, 15);
      assertEq('Unchanged paragraph 1 marked auto_copied', stubs[0].status, 'auto_copied');

      // 5. Prompt Builder (Bilingual test)
      const promptEn = ComparatorEngine.buildCopilotPrompt({
        promptLang: 'en',
        targetLangName: 'English',
        classification: 'MODIFIED',
        refOld: 'DE Alt',
        refNew: 'DE Neu',
        targetOld: 'EN Old',
        diffResult: diffRes
      });
      assertOk('buildCopilotPrompt generates English prompt', promptEn.includes('professional translator') && promptEn.includes('CRITICAL INSTRUCTIONS'));

      // 6. Export prep
      const exportData = ComparatorEngine.prepareExportData({
        rows: sampleRows,
        headers: SampleGenerator.getSampleHeaders(),
        mappings: SampleGenerator.getSampleDefaultMappings(),
        translations: { 'CS 2024': stubs },
        exportType: 'clean'
      });
      assertOk('clean export leaves out removed paragraphs', exportData.data.length < sampleRows.length);

      const summaryHtml = `
        <div style="padding:0.75rem; background:var(--bg-tertiary); border-radius:var(--radius-md); margin-bottom:1rem; font-weight:700;">
          ${I18n.getLanguage() === 'en' ? 'Diagnostics Result:' : 'Výsledek diagnostiky:'} <span style="color:#15803d;">${pass} ${I18n.getLanguage() === 'en' ? 'passed' : 'úspěšných'}</span>, <span style="color:${fail > 0 ? '#b91c1c' : '#15803d'};">${fail} ${I18n.getLanguage() === 'en' ? 'failed' : 'chyb'}</span>
        </div>
      `;

      elements.diagnosticsLog.innerHTML = summaryHtml + logs.join('');
    }, 50);
  }

  /**
   * Prompt Template Settings Modal
   */
  function initPromptSettings() {
    elements.settingsPromptBtn.addEventListener('click', function () {
      elements.promptSettingsModal.classList.add('open');
      elements.promptTemplateInput.value = state.customPromptTemplate || getDefaultPromptSample();
    });

    elements.closePromptSettingsBtn.addEventListener('click', function () {
      elements.promptSettingsModal.classList.remove('open');
    });

    elements.savePromptTemplateBtn.addEventListener('click', function () {
      const value = elements.promptTemplateInput.value.trim();
      // Saving the untouched sample keeps the built-in (classification-specific) prompts
      state.customPromptTemplate = value && value !== getDefaultPromptSample().trim() ? value : null;
      persistState(0);
      elements.promptSettingsModal.classList.remove('open');
      showToast(t('promptSaveBtn'), 'success');
    });

    elements.resetPromptTemplateBtn.addEventListener('click', function () {
      state.customPromptTemplate = null;
      elements.promptTemplateInput.value = getDefaultPromptSample();
      persistState(0);
      showToast(t('promptResetBtn'), 'info');
    });
  }

  function getDefaultPromptSample() {
    if (I18n.getLanguage() === 'en') {
      return 'You are an experienced professional translator.\n\nTASK: Update the approved translation in language {TARGET_LANG} to match the new version of the reference text.\nParagraph: {ROW_ID}\n\n1. REFERENCE – OLD VERSION ({REF_LANG}):\n{REF_OLD}\n\n2. REFERENCE – NEW VERSION:\n{REF_NEW}\n{CHANGES_SUMMARY}\n\n3. CURRENT APPROVED TRANSLATION ({TARGET_LANG}):\n{TARGET_OLD}\n\nINSTRUCTIONS:\nMake ONLY the minimal changes reflecting the reference difference. Keep the existing tone and terminology. Output only the final text.';
    }
    return 'Jste zkušený odborný překladatel.\n\nÚKOL: Aktualizujte schválený překlad v jazyce {TARGET_LANG} podle nové verze referenčního textu.\nOdstavec: {ROW_ID}\n\n1. REFERENCE – STARÁ VERZE ({REF_LANG}):\n{REF_OLD}\n\n2. REFERENCE – NOVÁ VERZE:\n{REF_NEW}\n{CHANGES_SUMMARY}\n\n3. DOSAVADNÍ SCHVÁLENÝ PŘEKLAD ({TARGET_LANG}):\n{TARGET_OLD}\n\nPOKYNY:\nProveďte POUZE minimální změny odpovídající rozdílu v referenci. Zachovejte zavedenou terminologii a styl. Vraťte výhradně výsledný text.';
  }

  /**
   * Reset & Restore Project
   */
  function initProjectReset() {
    elements.resetProjectBtn.addEventListener('click', function () {
      if (confirm(t('alertConfirmReset'))) {
        StorageManager.clearState();
        location.reload();
      }
    });
  }

  function restoreProjectFromData(data) {
    resetWorkspaceForNewTable();
    state.importSource = null;
    state.fileName = data.fileName || '';
    state.activeSheet = data.activeSheet || '';
    state.sheetNames = state.activeSheet ? [state.activeSheet] : [];
    state.rows = Array.isArray(data.rows) ? data.rows : [];
    state.headers = Array.isArray(data.headers) && data.headers.length > 0
      ? data.headers
      : Object.keys(state.rows[0] || {});
    state.hasHeaderRow = data.hasHeaderRow !== false;
    state.columnLetters = data.columnLetters || {};
    state.rowNumbers = Array.isArray(data.rowNumbers) && data.rowNumbers.length === state.rows.length
      ? data.rowNumbers
      : state.rows.map(function (r, idx) { return idx + 2; });
    state.skippedRows = Array.isArray(data.skippedRows) ? data.skippedRows : [];
    state.previewLimit = PREVIEW_INITIAL_ROWS;
    const mappings = data.mappings || {};
    state.mappings = {
      refOldCol: mappings.refOldCol || '',
      refNewCol: mappings.refNewCol || '',
      targetCols: Array.isArray(mappings.targetCols) ? mappings.targetCols.slice() : [],
      languages: {}
    };
    state.mappings.targetCols.forEach(function (col) {
      const saved = (mappings.languages && mappings.languages[col]) || {};
      state.mappings.languages[col] = {
        label: saved.label || '',
        newCol: saved.newCol || '',
        newColName: saved.newColName || '',
        isNewLanguage: saved.isNewLanguage === true || (typeof col === 'string' && col.indexOf('__new__') === 0)
      };
    });
    state.exportLangCol = data.exportLangCol || '';
    state.partialContextCols = Array.isArray(data.partialContextCols) ? data.partialContextCols.slice() : null;
    state.translations = data.translations && typeof data.translations === 'object' ? data.translations : {};
    state.customPromptTemplate = data.customPromptTemplate || null;
    state.pdfTemplate = data.pdfTemplate || null;
    state.activeTargetCol = data.activeTargetCol || state.mappings.targetCols[0] || '';
    state.auxiliaryCol = (typeof data.auxiliaryCol === 'string') ? data.auxiliaryCol : '';
    elements.sheetOptionsWrapper.style.display = 'none';
    elements.hasHeaderRowCheckbox.checked = state.hasHeaderRow;

    renderColumnMappingControls();
    renderDataPreview();

    // An incomplete mapping (e.g. autosave made while still in Step 1) resumes in Step 1
    if (getMappingError(true)) {
      setStep(1);
      return;
    }

    ensureStubsInitialized();

    // Resume in the studio when the user has done any work beyond the automatic defaults
    const hasEdits = state.mappings.targetCols.some(function (col) {
      return (state.translations[col] || []).some(function (it) {
        return it && (it.classification === 'MODIFIED' || it.classification === 'ADDED'
          ? (it.targetNew || it.isVerified)
          : false);
      });
    });

    setStep(hasEdits ? 3 : 2);
  }

  function checkStoredStateOnLoad() {
    const stored = StorageManager.loadState();
    if (!stored || !Array.isArray(stored.rows) || stored.rows.length === 0) return;

    const msg = t('confirmResumeProject', { DATE: new Date(stored.timestamp).toLocaleString() });
    if (!confirm(msg)) return;
    try {
      restoreProjectFromData(StorageManager.parseProjectJson(stored));
    } catch (err) {
      console.error(err);
      alert(projectErrorMessage(err, ''));
    }
  }

  /**
   * The header stays sticky only while it fits on one line; the sticky filter toolbar of the
   * studio sits right below it (CSS variable --sticky-header-h).
   */
  function initStickyHeader() {
    const header = document.querySelector('.app-header');
    if (!header) return;
    function update() {
      header.classList.remove('header-static');
      const tall = header.offsetHeight > 80;
      header.classList.toggle('header-static', tall);
      document.documentElement.style.setProperty('--sticky-header-h', (tall ? 0 : header.offsetHeight) + 'px');
    }
    update();
    window.addEventListener('resize', update);
    // Labels change length when the UI language is switched
    elements.langToggleBtn.addEventListener('click', update);
  }

  /**
   * Close the top-most open modal with Escape or a click on its backdrop.
   */
  function initModalDismissal() {
    document.querySelectorAll('.modal-overlay').forEach(function (overlay) {
      overlay.addEventListener('click', function (e) {
        if (e.target === overlay) overlay.classList.remove('open');
      });
    });
    document.addEventListener('keydown', function (e) {
      if (e.key !== 'Escape') return;
      const open = document.querySelectorAll('.modal-overlay.open');
      if (open.length > 0) {
        open[open.length - 1].classList.remove('open');
        e.preventDefault();
      }
    });
  }

  /**
   * Main App Initialization (runs once even if DOMContentLoaded is delivered more than once,
   * otherwise every listener would be registered twice)
   */
  let appInitialized = false;

  function initApp() {
    if (appInitialized) return;
    appInitialized = true;
    initLanguage();
    initTheme();
    initLayoutMode();
    initFileUpload();
    initFilters();
    initExports();
    initDiagnostics();
    initPromptSettings();
    initBatchPromptModal();
    initSmartPasteModal();
    initTranslationImport();
    initPdfTemplateSettings();
    initAuxLangSelector();
    initProjectButtons();
    initProjectReset();

    // Add new language controls
    if (elements.addNewLangBtn) {
      elements.addNewLangBtn.addEventListener('click', function () {
        const name = elements.newLangNameInput ? elements.newLangNameInput.value : '';
        const colName = elements.newLangColNameInput ? elements.newLangColNameInput.value : '';
        addNewLanguage(name, colName);
      });
    }
    if (elements.newLangNameInput) {
      elements.newLangNameInput.addEventListener('keydown', function (e) {
        if (e.key === 'Enter') {
          e.preventDefault();
          const name = elements.newLangNameInput.value;
          const colName = elements.newLangColNameInput ? elements.newLangColNameInput.value : '';
          addNewLanguage(name, colName);
        }
      });
    }
    if (elements.newLangColNameInput) {
      elements.newLangColNameInput.addEventListener('keydown', function (e) {
        if (e.key === 'Enter') {
          e.preventDefault();
          const name = elements.newLangNameInput ? elements.newLangNameInput.value : '';
          const colName = elements.newLangColNameInput.value;
          addNewLanguage(name, colName);
        }
      });
    }

    // Bulk action buttons
    if (elements.approveUnchangedBtn) {
      elements.approveUnchangedBtn.addEventListener('click', approveUnchangedInActiveLanguage);
    }
    if (elements.approveAllUnchangedGlobalBtn) {
      elements.approveAllUnchangedGlobalBtn.addEventListener('click', approveAllUnchangedGlobal);
    }

    // Stepper navigation clicks
    elements.stepNav1.addEventListener('click', function () { setStep(1); });
    elements.stepNav2.addEventListener('click', function () { setStep(2); });
    elements.stepNav3.addEventListener('click', function () { setStep(3); });
    elements.stepNav4.addEventListener('click', function () { setStep(4); });

    // Step 1 Confirm button
    elements.confirmMappingBtn.addEventListener('click', function () {
      state.mappings.refOldCol = elements.refOldSelect.value;
      state.mappings.refNewCol = elements.refNewSelect.value;

      if (elements.newLangNameInput && elements.newLangNameInput.value.trim()) {
        addNewLanguage(elements.newLangNameInput.value.trim(), elements.newLangColNameInput ? elements.newLangColNameInput.value.trim() : '');
      }

      const mappingError = getMappingError(true);
      if (mappingError) {
        alert(t(mappingError));
        return;
      }

      finalizeLanguageConfig();
      if (isDirectTranslationMode()) {
        setStep(3);
      } else {
        setStep(2);
      }
      showToast(t('toastMappingSaved'), 'success');
    });

    elements.goToTranslationBtn.addEventListener('click', function () {
      if (state.mappings.targetCols.length === 0) {
        alert(t(isDirectTranslationMode() ? 'alertSelectTargetLangsDirect' : 'alertSelectTargetLangs'));
        setStep(1);
        return;
      }
      setStep(3);
    });

    if (elements.backToSettingsBtn) {
      elements.backToSettingsBtn.addEventListener('click', function () {
        setStep(1);
      });
    }

    elements.goToExportBtn.addEventListener('click', function () {
      setStep(4);
    });

    initModalDismissal();
    initStickyHeader();

    // Write a pending (debounced) autosave before the page goes away
    window.addEventListener('pagehide', function () { StorageManager.flushPendingSave(); });
    document.addEventListener('visibilitychange', function () {
      if (document.visibilityState === 'hidden') StorageManager.flushPendingSave();
    });

    // Expose for testing & diagnostics
    if (typeof window !== 'undefined') {
      window.__app = {
        state: state,
        elements: elements,
        setStep: setStep,
        addNewLanguage: addNewLanguage,
        removeNewLanguage: removeNewLanguage,
        restoreProjectFromData: restoreProjectFromData,
        importTranslationsFile: importTranslationsFile,
        initApp: initApp,
        approveUnchangedInActiveLanguage: approveUnchangedInActiveLanguage,
        approveAllUnchangedGlobal: approveAllUnchangedGlobal,
        updateAuxLangOptions: updateAuxLangOptions,
        openBatchPromptModal: openBatchPromptModal,
        openPdfTemplateModal: openPdfTemplateModal,
        exportManagerialSummary: exportManagerialSummary,
        exportTables: exportTables,
        renderTableFormatOptions: renderTableFormatOptions,
        exportProjectFile: exportProjectFile,
        loadProjectFile: loadProjectFile,
        processUploadedFile: processUploadedFile,
        checkStoredStateOnLoad: checkStoredStateOnLoad
      };
    }

    // Check for autosaved session
    setTimeout(checkStoredStateOnLoad, 100);
  }

  // Run on DOM ready
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', initApp);
  } else {
    initApp();
  }
})();
