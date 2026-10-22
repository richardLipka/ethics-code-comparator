/**
 * Application Controller - multi-language version comparator and translation assistant.
 * Bilingual (Czech & English), Full-width 4-column comparison layout, 100% offline.
 */

(function () {
  'use strict';

  // Application State
  const state = {
    currentStep: 1, // 1: Upload & Map, 2: Ref Overview, 3: Translation Studio, 4: Export
    workbook: null,
    fileName: '',
    sheetNames: [],
    activeSheet: '',
    hasHeaderRow: true,
    headers: [],
    rows: [],
    rowNumbers: [], // Excel row number of every data row
    mappings: {
      refOldCol: '',
      refNewCol: '',
      targetCols: [], // old-version columns of the languages to update
      languages: {} // { [oldCol]: { label, newCol, newColName } }
    },
    exportLangCol: '',
    partialContextCols: null,
    // Translations map: { [targetCol]: [ { rowIndex, targetNew, status, isVerified, notes } ] }
    translations: {},
    activeTargetCol: '',
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
    
    // Step 1 Controls
    sheetOptionsWrapper: document.getElementById('sheetOptionsWrapper'),
    sheetSelectWrapper: document.getElementById('sheetSelectWrapper'),
    sheetSelect: document.getElementById('sheetSelect'),
    hasHeaderRowCheckbox: document.getElementById('hasHeaderRowCheckbox'),
    mappingContainer: document.getElementById('mappingContainer'),
    refOldSelect: document.getElementById('refOldSelect'),
    refNewSelect: document.getElementById('refNewSelect'),
    languagesTableBody: document.getElementById('languagesTableBody'),
    confirmMappingBtn: document.getElementById('confirmMappingBtn'),
    dataPreviewContainer: document.getElementById('dataPreviewContainer'),
    previewTableHead: document.getElementById('previewTableHead'),
    previewTableBody: document.getElementById('previewTableBody'),
    previewRowsCount: document.getElementById('previewRowsCount'),

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
    layoutToggleBtn: document.getElementById('layoutToggleBtn'),
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

    // Update layout toggle text
    if (elements.layoutToggleBtn) {
      elements.layoutToggleBtn.textContent = state.layoutMode === '4cols' ? t('layout4Cols') : t('layout2x2');
    }
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
  function initLayoutMode() {
    const saved = readPref('comparator_layout', '4cols');
    state.layoutMode = saved === '2x2' ? '2x2' : '4cols';
    if (elements.layoutToggleBtn) {
      elements.layoutToggleBtn.textContent = state.layoutMode === '4cols' ? t('layout4Cols') : t('layout2x2');
      elements.layoutToggleBtn.addEventListener('click', function () {
        state.layoutMode = state.layoutMode === '4cols' ? '2x2' : '4cols';
        writePref('comparator_layout', state.layoutMode);
        elements.layoutToggleBtn.textContent = state.layoutMode === '4cols' ? t('layout4Cols') : t('layout2x2');
        renderArticleCards();
      });
    }
  }

  /**
   * Stepper navigation
   */
  /**
   * Returns the i18n key of the first mapping problem, or '' when the mapping is usable.
   */
  function getMappingError(requireTargets) {
    const m = state.mappings;
    if (!m.refOldCol || !m.refNewCol) return 'alertSelectRefCols';
    if (m.refOldCol === m.refNewCol) return 'alertDifferentRefCols';
    if (requireTargets && (!m.targetCols || m.targetCols.length === 0)) return 'alertSelectTargetLangs';
    return '';
  }

  function setStep(stepNumber) {
    if (stepNumber > 1 && (!state.rows || state.rows.length === 0)) {
      showToast(t('toastNoData'), 'warning');
      return;
    }
    if (stepNumber > 1) {
      const mappingError = getMappingError(stepNumber >= 3);
      if (mappingError) {
        showToast(t(mappingError), 'warning');
        return;
      }
    }

    state.currentStep = stepNumber;

    // Toggle active sections
    elements.step1Section.style.display = stepNumber === 1 ? 'block' : 'none';
    elements.step2Section.style.display = stepNumber === 2 ? 'block' : 'none';
    elements.step3Section.style.display = stepNumber === 3 ? 'block' : 'none';
    elements.step4Section.style.display = stepNumber === 4 ? 'block' : 'none';

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
    }

    window.scrollTo({ top: 0, behavior: 'smooth' });
  }

  /**
   * Handle File Upload & Sheet Reading
   */
  function initFileUpload() {
    elements.dropzone.addEventListener('click', function () {
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

    // Re-read the sheet when the user says the first row is (not) a header
    elements.hasHeaderRowCheckbox.addEventListener('change', function () {
      if (state.workbook && state.activeSheet) {
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

  function loadProjectFile(file) {
    const reader = new FileReader();
    reader.onload = function (evt) {
      try {
        const parsed = StorageManager.parseProjectJson(evt.target.result);
        restoreProjectFromData(parsed);
        persistState(0);
        showToast(t('toastProjectLoaded'), 'success');
      } catch (err) {
        console.error(err);
        alert(t('alertProjectLoadError', { MSG: err.message }));
      }
    };
    reader.onerror = function () {
      alert(t('alertProjectLoadError', { MSG: reader.error ? reader.error.message : '' }));
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

  function processUploadedFile(file) {
    const reader = new FileReader();
    // CSV is decoded as UTF-8 text; reading it as bytes would garble diacritics
    const isCsv = /\.csv$/i.test(file.name) || file.type === 'text/csv';

    reader.onload = function (e) {
      try {
        const workbook = isCsv
          ? XLSX.read(String(e.target.result).replace(/^﻿/, ''), { type: 'string' })
          : XLSX.read(new Uint8Array(e.target.result), { type: 'array' });

        if (!workbook.SheetNames || workbook.SheetNames.length === 0) {
          alert(t('alertNoSheets'));
          return;
        }

        state.fileName = file.name;
        state.workbook = workbook;
        state.sheetNames = workbook.SheetNames;

        // Setup sheet select dropdown
        elements.sheetSelect.innerHTML = '';
        state.sheetNames.forEach(function (sheetName) {
          const opt = document.createElement('option');
          opt.value = sheetName;
          opt.textContent = sheetName;
          elements.sheetSelect.appendChild(opt);
        });

        elements.sheetOptionsWrapper.style.display = 'flex';
        elements.sheetSelectWrapper.style.display = state.sheetNames.length > 1 ? 'block' : 'none';
        if (loadSheet(state.sheetNames[0])) {
          showToast(t('toastFileLoaded', { NAME: file.name }), 'success');
        }
      } catch (err) {
        console.error('Chyba při parsování Excel souboru:', err);
        alert(t('alertFileReadError', { MSG: err.message }));
      }
    };
    reader.onerror = function () {
      alert(t('alertFileReadError', { MSG: reader.error ? reader.error.message : file.name }));
    };

    if (isCsv) {
      reader.readAsText(file, 'UTF-8');
    } else {
      reader.readAsArrayBuffer(file);
    }
  }

  // Cells often carry stray spaces or line breaks around the text (a cell may even look empty
  // in a narrow preview), so text is trimmed on import
  function trimCell(value) {
    return typeof value === 'string' ? value.trim() : (value === undefined || value === null ? '' : value);
  }

  function loadSheet(sheetName) {
    const sheet = state.workbook.Sheets[sheetName];
    const hasHeader = elements.hasHeaderRowCheckbox.checked;
    const range = sheet && sheet['!ref'] ? XLSX.utils.decode_range(sheet['!ref']) : null;
    const rawRows = range
      ? XLSX.utils.sheet_to_json(sheet, hasHeader ? { defval: '' } : { header: 'A', defval: '' })
      : [];

    if (!rawRows || rawRows.length === 0) {
      alert(t('alertEmptySheet'));
      return false;
    }

    const letters = [];
    for (let c = range.s.c; c <= range.e.c; c++) letters.push(XLSX.utils.encode_col(c));

    let headers;
    const columnLetters = {};
    if (hasHeader) {
      // Header order of the sheet itself (Object.keys would move numeric headers first)
      const headerRow = XLSX.utils.sheet_to_json(sheet, { header: 1, range: range.s.r, defval: '' })[0] || [];
      const keys = Object.keys(rawRows[0]);
      const ordered = [];
      headerRow.forEach(function (h, i) {
        const key = keys.find(function (k) { return k === String(h); }) ||
          keys.find(function (k) { return /^__EMPTY/.test(k) && ordered.indexOf(k) === -1; });
        if (key !== undefined && ordered.indexOf(key) === -1) {
          ordered.push(key);
          columnLetters[key] = letters[i];
        }
      });
      headers = ordered.length === keys.length ? ordered : keys;
    } else {
      headers = letters;
      letters.forEach(function (l) { columnLetters[l] = l; });
    }

    state.activeSheet = sheetName;
    state.hasHeaderRow = hasHeader;
    elements.sheetSelect.value = sheetName;
    state.headers = headers;
    state.columnLetters = columnLetters;
    state.rowNumbers = rawRows.map(function (r, i) {
      return typeof r.__rowNum__ === 'number' ? r.__rowNum__ + 1 : i + (hasHeader ? 2 : 1);
    });
    state.rows = rawRows.map(function (r) {
      const row = {};
      headers.forEach(function (h) { row[h] = trimCell(r[h]); });
      return row;
    });
    resetWorkspaceForNewTable();
    selectDefaultLanguageColumns();

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
    state.workbook = null;
    state.hasHeaderRow = true;
    state.rows = SampleGenerator.getSampleData();
    state.headers = SampleGenerator.getSampleHeaders();
    state.rowNumbers = state.rows.map(function (r, idx) { return idx + 2; });
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

    showToast(t('toastSampleLoaded'), 'success');
  }

  /**
   * Human-readable column name: the header, or "Column C" when the sheet has no header row
   * or the header cell is empty.
   */
  function columnDisplayName(col) {
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
   * Step 1: Render interactive data preview table
   */
  function renderDataPreview() {
    if (!elements.dataPreviewContainer) return;
    if (!state.rows || state.rows.length === 0 || !state.headers || state.headers.length === 0) {
      elements.dataPreviewContainer.style.display = 'none';
      return;
    }

    elements.dataPreviewContainer.style.display = 'block';
    if (elements.previewRowsCount) {
      elements.previewRowsCount.textContent = `(${state.rows.length} ${t('rowsLabel')})`;
    }

    const thStyle = 'padding:0.5rem 0.75rem; font-size:0.8rem; background:var(--bg-tertiary); border:1px solid var(--border-color);';
    elements.previewTableHead.innerHTML = '<tr>' + `<th style="${thStyle}">${ComparatorEngine.escapeHtml(t('rowLabel'))}</th>` +
      state.headers.map(function (h) {
        return `<th style="${thStyle}">${ComparatorEngine.escapeHtml(columnDisplayName(h))}</th>`;
      }).join('') + '</tr>';

    // Sample rows (first 4); whitespace is collapsed so long single-line cells stay readable
    const tdStyle = 'padding:0.45rem 0.75rem; font-size:0.8rem; border:1px solid var(--border-color); color:var(--text-secondary); max-width:240px; overflow:hidden; text-overflow:ellipsis; white-space:nowrap;';
    elements.previewTableBody.innerHTML = state.rows.slice(0, 4).map(function (row, idx) {
      return '<tr>' + `<td style="${tdStyle} font-family:var(--font-mono);">${rowNumberOf(idx)}</td>` + state.headers.map(function (h) {
        let val = row[h] !== undefined ? String(row[h]).replace(/\s+/g, ' ').trim() : '';
        const full = val;
        if (val.length > 60) val = val.substring(0, 57) + '...';
        return `<td style="${tdStyle}" title="${ComparatorEngine.escapeHtml(full.slice(0, 500))}">${ComparatorEngine.escapeHtml(val)}</td>`;
      }).join('') + '</tr>';
    }).join('');
  }

  function downloadSampleExcelFile() {
    const data = SampleGenerator.getSampleData();
    const headers = SampleGenerator.getSampleHeaders();
    const worksheet = XLSX.utils.json_to_sheet(data, { header: headers });

    worksheet['!cols'] = headers.map(function (h) {
      let maxLen = h.length;
      data.forEach(function (r) {
        const val = r[h] ? String(r[h]) : '';
        if (val.length > maxLen) maxLen = Math.min(val.length, 60);
      });
      return { wch: Math.max(maxLen + 3, 14) };
    });

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, worksheet, 'Data');
    XLSX.writeFile(wb, 'ukazkovy_dokument.xlsx');
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
      state.mappings.languages[col] = { label: '', newCol: '', newColName: '' };
    }
    return state.mappings.languages[col];
  }

  function languageLabel(col) {
    const cfg = state.mappings.languages && state.mappings.languages[col];
    return (cfg && cfg.label) || ComparatorEngine.defaultLanguageLabel(columnDisplayName(col), state.mappings.refOldCol);
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
      // Keep the sheet's column order
      targets.sort(function (a, b) { return state.headers.indexOf(a) - state.headers.indexOf(b); });
    } else {
      const pos = targets.indexOf(col);
      if (pos !== -1) targets.splice(pos, 1);
    }
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

  function fillColumnSelect(select, value) {
    select.innerHTML = '';
    const placeholder = document.createElement('option');
    placeholder.value = '';
    placeholder.textContent = t('selectColumnPlaceholder');
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
    fillColumnSelect(elements.refOldSelect, state.mappings.refOldCol);
    fillColumnSelect(elements.refNewSelect, state.mappings.refNewCol);

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
      state.translations[targetCol] = ComparatorEngine.rehydrateStubs(
        state.rows,
        state.translations[targetCol],
        targetCol,
        state.mappings.refOldCol,
        state.mappings.refNewCol,
        state.mappings.languages[targetCol].newCol
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
      const refOld = row[refOldCol];
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
      if (classification === 'UNCHANGED') {
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

    const stubs = state.translations[col];
    let count = 0;
    stubs.forEach(function (stub) {
      if (stub.classification === 'UNCHANGED' && !stub.isVerified) {
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
      const stubs = state.translations[col] || [];
      stubs.forEach(function (stub) {
        if (stub.classification === 'UNCHANGED' && !stub.isVerified) {
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
    const esc = ComparatorEngine.escapeHtml;
    const langLabel = esc(languageLabel(col));
    const refLabel = esc(columnDisplayName(state.mappings.refNewCol));

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

      // 2. Target Diff (old -> new version of the language)
      const targetDiffHtml = renderTargetDiffHtml(stub);

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

      // Column 1: Reference diff
      const col1RefHtml = `
        <div class="sub-panel">
          <div class="sub-panel-title">
            <span>${t('cardRefPanelTitle', { LANG: refLabel })}</span>
            ${refDiffRes.stats.hasChanges ? `<span style="font-size:0.75rem; color:#b45309; font-weight:600;">${t('changePct', { PCT: Math.round((1 - refDiffRes.stats.similarity) * 100) })}</span>` : `<span style="font-size:0.75rem; color:#15803d; font-weight:600;">${t('match100')}</span>`}
          </div>
          <div class="diff-box">${refDiffHtml || `<span style="color:var(--text-muted);">${t('emptyCell')}</span>`}</div>
        </div>
      `;

      // Column 2: Old version of the language
      const col2OldHtml = `
        <div class="sub-panel">
          <div class="sub-panel-title">
            <span>${t('cardTargetOldTitle', { LANG: langLabel })}</span>
          </div>
          <div class="diff-box" style="background-color:var(--bg-tertiary); color:var(--text-secondary);">${ComparatorEngine.escapeHtml(stub.targetOld) || `<span style="color:var(--text-muted); font-style:italic;">${t('emptyOriginalText')}</span>`}</div>
        </div>
      `;

      // Column 3: New version of the language (editor)
      const col3NewEditorHtml = `
        <div class="sub-panel">
          <div class="sub-panel-title">
            <span>${t('cardTargetNewTitle', { LANG: langLabel })}</span>
          </div>
          <textarea class="form-textarea target-textarea" data-index="${idx}" placeholder="${t('textareaPlaceholder')}">${ComparatorEngine.escapeHtml(stub.targetNew)}</textarea>
          <div class="editor-toolbar">
            <button class="btn btn-sm btn-primary copy-prompt-btn" data-index="${idx}">
              ${t('copyPromptBtn')}
            </button>
            ${classification === 'MODIFIED' ? `<button class="btn btn-sm btn-outline prefill-old-btn" data-index="${idx}">${t('prefillOldBtn')}</button>` : ''}
            ${classification === 'ADDED' ? `<button class="btn btn-sm btn-outline prefill-ref-btn" data-index="${idx}">${t('prefillRefBtn')}</button>` : ''}
          </div>
        </div>
      `;

      // Column 4: Translation Diff (Old Target vs New Target Live Preview)
      const col4DiffHtml = `
        <div class="sub-panel">
          <div class="sub-panel-title">
            <span>${t('cardTargetDiffTitle', { LANG: langLabel })}</span>
          </div>
          <div class="diff-box live-target-diff" id="target-diff-${idx}">${targetDiffHtml}</div>
        </div>
      `;

      // Full-Width 4-Columns or 2x2 Layout Grid
      const gridClass = state.layoutMode === '2x2' ? 'layout-2x2' : 'layout-4cols';

      card.innerHTML = `
        ${headerHtml}
        <div class="article-grid ${gridClass}">
          ${col1RefHtml}
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
  }

  function copyPromptForIndex(idx) {
    const col = state.activeTargetCol;
    const stub = state.translations[col][idx];
    const refDiffRes = ComparatorEngine.computeDiff(stub.refOld, stub.refNew);

    const promptText = ComparatorEngine.buildCopilotPrompt({
      promptLang: I18n.getLanguage(),
      targetLangName: languageLabel(col),
      refLangName: columnDisplayName(state.mappings.refNewCol),
      rowId: ComparatorEngine.rowTag(state.rowNumbers, idx),
      classification: stub.classification,
      refOld: stub.refOld,
      refNew: stub.refNew,
      targetOld: stub.targetOld,
      diffResult: refDiffRes,
      customTemplate: state.customPromptTemplate
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

  function fallbackCopyText(text) {
    const temp = document.createElement('textarea');
    temp.value = text;
    document.body.appendChild(temp);
    temp.select();
    document.execCommand('copy');
    document.body.removeChild(temp);
    showToast(t('toastPromptCopied'), 'success');
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
      optRef.textContent = `${columnDisplayName(state.mappings.refNewCol)} (${t('pdfOptionRefNew')})`;
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
   * Checkboxes of the columns that accompany the selected language in the partial workbook
   * (by default the reference old and new version).
   */
  function renderPartialContextOptions() {
    const container = elements.partialContextContainer;
    if (!container) return;
    const langCol = state.exportLangCol;
    const isReference = langCol === state.mappings.refNewCol;
    const own = isReference
      ? [state.mappings.refOldCol, state.mappings.refNewCol]
      : [langCol, state.mappings.languages[langCol] && state.mappings.languages[langCol].newCol];
    if (!Array.isArray(state.partialContextCols)) {
      state.partialContextCols = [state.mappings.refOldCol, state.mappings.refNewCol].filter(Boolean);
    }

    container.innerHTML = '';
    state.headers.forEach(function (h) {
      if (own.indexOf(h) !== -1) return;
      const label = document.createElement('label');
      label.className = 'checkbox-label';
      const cb = document.createElement('input');
      cb.type = 'checkbox';
      cb.value = h;
      cb.checked = state.partialContextCols.indexOf(h) !== -1;
      cb.addEventListener('change', function () {
        const list = state.partialContextCols;
        const pos = list.indexOf(h);
        if (cb.checked && pos === -1) list.push(h);
        if (!cb.checked && pos !== -1) list.splice(pos, 1);
        persistState(300);
      });
      label.appendChild(cb);
      label.appendChild(document.createTextNode(columnDisplayName(h)));
      container.appendChild(label);
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

    const options = {
      rows: state.rows,
      mappings: state.mappings,
      translations: state.translations,
      rowNumbers: state.rowNumbers,
      langCol: col,
      documentName: documentBaseName(),
      uiLang: I18n.getLanguage()
    };

    if (format === 'markdown') {
      PdfExporter.downloadManagerialMarkdown(options);
      showToast(t('mgrExportMdSuccess'), 'success');
    } else if (format === 'word') {
      PdfExporter.downloadManagerialWordDoc(options);
      showToast(t('mgrExportDocSuccess'), 'success');
    } else if (format === 'pdf') {
      PdfExporter.printManagerialReport(options);
    }
  }

  function documentOptions(col) {
    return {
      rows: state.rows,
      mappings: state.mappings,
      translations: state.translations,
      rowNumbers: state.rowNumbers,
      targetCol: col,
      documentName: documentBaseName(),
      uiLang: I18n.getLanguage(),
      template: state.pdfTemplate
    };
  }

  function initExports() {
    elements.downloadComprehensiveBtn.addEventListener('click', function () {
      downloadExcelWorkbook('comprehensive');
    });

    elements.downloadCleanBtn.addEventListener('click', function () {
      downloadExcelWorkbook('clean');
    });

    elements.downloadPartialBtn.addEventListener('click', function () {
      const col = selectedExportLang();
      if (col) downloadExcelWorkbook('language', col);
    });

    elements.exportProjectJsonBtn.addEventListener('click', function () {
      exportProjectFile();
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
      if (col) PdfExporter.printLanguageMutation(documentOptions(col));
    });

    elements.downloadPdfHtmlBtn.addEventListener('click', function () {
      const col = selectedExportLang();
      if (!col) return;
      PdfExporter.downloadPrintableHtml(documentOptions(col));
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
            PdfExporter.printLanguageMutation(documentOptions(col));
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

    const promptText = ComparatorEngine.buildBatchCopilotPrompt({
      promptLang: I18n.getLanguage(),
      targetCol: col,
      targetLangName: languageLabel(col),
      refLangName: columnDisplayName(state.mappings.refNewCol),
      rows: state.rows,
      rowNumbers: state.rowNumbers,
      mappings: state.mappings,
      translations: state.translations
    });

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

  function downloadExcelWorkbook(exportType, langCol) {
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

    const result = ComparatorEngine.buildExportWorkbook(XLSX, {
      rows: state.rows,
      headers: state.headers,
      rowNumbers: state.rowNumbers,
      mappings: state.mappings,
      translations: state.translations,
      exportType: exportType,
      langCol: langCol,
      contextCols: state.partialContextCols || [state.mappings.refOldCol, state.mappings.refNewCol],
      uiLang: I18n.getLanguage()
    });

    const safe = function (text) { return String(text).replace(/[\\/:*?"<>|]+/g, '_'); };
    const suffix = exportType === 'clean' ? 'nova_verze'
      : (exportType === 'language' ? safe(langCol === state.mappings.refNewCol ? columnDisplayName(langCol) : languageLabel(langCol)) : 'srovnani');
    const outputFileName = `${safe(documentBaseName())}_${suffix}.xlsx`;

    XLSX.writeFile(result.workbook, outputFileName);
    showToast(t('toastWorkbookGenerated', { NAME: outputFileName }), 'success');
    if (result.unverifiedCount > 0) {
      const key = exportType === 'clean' ? 'toastExportUnverifiedClean'
        : (exportType === 'language' ? 'toastExportUnverifiedLanguage' : 'toastExportUnverified');
      showToast(t(key, { COUNT: result.unverifiedCount }), 'warning');
    }
    return outputFileName;
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
    state.workbook = null;
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
        newColName: saved.newColName || ''
      };
    });
    state.exportLangCol = data.exportLangCol || '';
    state.partialContextCols = Array.isArray(data.partialContextCols) ? data.partialContextCols.slice() : null;
    state.translations = data.translations && typeof data.translations === 'object' ? data.translations : {};
    state.customPromptTemplate = data.customPromptTemplate || null;
    state.pdfTemplate = data.pdfTemplate || null;
    state.activeTargetCol = data.activeTargetCol || state.mappings.targetCols[0] || '';
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
        return it.classification === 'MODIFIED' || it.classification === 'ADDED'
          ? (it.targetNew || it.isVerified)
          : false;
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
      alert(t('alertProjectLoadError', { MSG: err.message }));
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
   * Main App Initialization
   */
  function initApp() {
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
    initPdfTemplateSettings();
    initProjectButtons();
    initProjectReset();

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

      if (!state.mappings.refOldCol || !state.mappings.refNewCol) {
        alert(t('alertSelectRefCols'));
        return;
      }
      if (state.mappings.refOldCol === state.mappings.refNewCol) {
        alert(t('alertDifferentRefCols'));
        return;
      }
      if (state.mappings.targetCols.length === 0) {
        alert(t('alertSelectTargetLangs'));
        return;
      }

      finalizeLanguageConfig();
      setStep(2);
      showToast(t('toastMappingSaved'), 'success');
    });

    elements.goToTranslationBtn.addEventListener('click', function () {
      setStep(3);
    });

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
        restoreProjectFromData: restoreProjectFromData,
        initApp: initApp,
        approveUnchangedInActiveLanguage: approveUnchangedInActiveLanguage,
        approveAllUnchangedGlobal: approveAllUnchangedGlobal,
        openBatchPromptModal: openBatchPromptModal,
        openPdfTemplateModal: openPdfTemplateModal,
        exportManagerialSummary: exportManagerialSummary,
        downloadExcelWorkbook: downloadExcelWorkbook,
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
