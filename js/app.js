/**
 * Application Controller - Code of Ethics Multi-Language Comparator
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
    headers: [],
    rows: [],
    mappings: {
      refOldCol: '',
      refNewCol: '',
      targetCols: []
    },
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
    importProjectBtn: document.getElementById('importProjectBtn'),
    
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
    sheetSelectWrapper: document.getElementById('sheetSelectWrapper'),
    sheetSelect: document.getElementById('sheetSelect'),
    mappingContainer: document.getElementById('mappingContainer'),
    refOldSelect: document.getElementById('refOldSelect'),
    refNewSelect: document.getElementById('refNewSelect'),
    targetColsContainer: document.getElementById('targetColsContainer'),
    confirmMappingBtn: document.getElementById('confirmMappingBtn'),
    autoDetectMappingBtn: document.getElementById('autoDetectMappingBtn'),
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

    // Sample data button
    elements.loadSampleBtn.addEventListener('click', function () {
      loadSampleDemo();
    });

    // Download physical sample excel
    elements.downloadSampleBtn.addEventListener('click', function () {
      downloadSampleExcelFile();
    });

    // Project JSON Import/Export
    elements.importProjectBtn.addEventListener('click', function () {
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
    state.mappings = { refOldCol: '', refNewCol: '', targetCols: [] };
    state.translations = {};
    state.activeTargetCol = '';
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

  function loadSheet(sheetName) {
    const sheet = state.workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(sheet, { defval: '' });

    if (!rows || rows.length === 0) {
      alert(t('alertEmptySheet'));
      return false;
    }

    state.activeSheet = sheetName;
    elements.sheetSelect.value = sheetName;
    state.rows = rows;
    // Header order of the sheet itself (Object.keys would move numeric headers first)
    const headerRow = XLSX.utils.sheet_to_json(sheet, { header: 1, range: 0 })[0] || [];
    const keys = Object.keys(rows[0]);
    const ordered = headerRow.map(function (h, i) { return keys.find(function (k) { return k === String(h); }) || keys[i]; })
      .filter(function (k, i, arr) { return k !== undefined && arr.indexOf(k) === i; });
    state.headers = ordered.length === keys.length ? ordered : keys;
    resetWorkspaceForNewTable();

    renderColumnMappingControls();
    autoDetectColumnMappings();
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
    state.fileName = 'vzorec_eticky_kodex_2024_2026.xlsx';
    state.workbook = null;
    state.rows = SampleGenerator.getSampleData();
    state.headers = SampleGenerator.getSampleHeaders();
    state.sheetNames = ['Eticky_Kodex'];
    state.activeSheet = 'Eticky_Kodex';
    resetWorkspaceForNewTable();

    elements.sheetSelectWrapper.style.display = 'none';

    const sampleDefaults = SampleGenerator.getSampleDefaultMappings();
    state.mappings.refOldCol = sampleDefaults.refOldCol;
    state.mappings.refNewCol = sampleDefaults.refNewCol;
    state.mappings.targetCols = sampleDefaults.targetCols.slice();

    renderColumnMappingControls();

    elements.refOldSelect.value = state.mappings.refOldCol;
    elements.refNewSelect.value = state.mappings.refNewCol;
    renderTargetCheckboxes();
    renderDataPreview();
    revealMapping();

    showToast(t('toastSampleLoaded'), 'success');
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

    // Header row
    elements.previewTableHead.innerHTML = '<tr>' + state.headers.map(function (h) {
      return `<th style="padding:0.5rem 0.75rem; font-size:0.8rem; background:var(--bg-tertiary); border:1px solid var(--border-color);">${ComparatorEngine.escapeHtml(h)}</th>`;
    }).join('') + '</tr>';

    // Sample rows (first 4)
    const sampleRows = state.rows.slice(0, 4);
    elements.previewTableBody.innerHTML = sampleRows.map(function (row) {
      return '<tr>' + state.headers.map(function (h) {
        let val = row[h] !== undefined ? String(row[h]) : '';
        if (val.length > 60) val = val.substring(0, 57) + '...';
        return `<td style="padding:0.45rem 0.75rem; font-size:0.8rem; border:1px solid var(--border-color); color:var(--text-secondary); max-width:240px; overflow:hidden; text-overflow:ellipsis; white-space:nowrap;">${ComparatorEngine.escapeHtml(val)}</td>`;
      }).join('') + '</tr>';
    }).join('');
  }

  function downloadSampleExcelFile() {
    const data = SampleGenerator.getSampleData();
    const headers = SampleGenerator.getSampleHeaders();
    const worksheet = XLSX.utils.json_to_sheet(data, { header: headers });

    const colWidths = headers.map(function (h) {
      let maxLen = h.length;
      data.forEach(function (r) {
        const val = r[h] ? String(r[h]) : '';
        if (val.length > maxLen) maxLen = Math.min(val.length, 60);
      });
      return { wch: Math.max(maxLen + 3, 14) };
    });
    worksheet['!cols'] = colWidths;

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, worksheet, 'Eticky_Kodex');
    XLSX.writeFile(wb, 'vzorec_eticky_kodex_2024_2026.xlsx');
    showToast(t('toastSampleDownloaded'), 'success');
  }

  /**
   * Column Mapping UI & Heuristics
   */
  function renderColumnMappingControls() {
    elements.mappingContainer.style.display = 'block';

    elements.refOldSelect.innerHTML = `<option value="">-- ${t('refOldLabel')} --</option>`;
    elements.refNewSelect.innerHTML = `<option value="">-- ${t('refNewLabel')} --</option>`;

    state.headers.forEach(function (h) {
      const optOld = document.createElement('option');
      optOld.value = h;
      optOld.textContent = h;
      elements.refOldSelect.appendChild(optOld);

      const optNew = document.createElement('option');
      optNew.value = h;
      optNew.textContent = h;
      elements.refNewSelect.appendChild(optNew);
    });

    elements.refOldSelect.onchange = function () {
      state.mappings.refOldCol = elements.refOldSelect.value;
      renderTargetCheckboxes();
    };

    elements.refNewSelect.onchange = function () {
      state.mappings.refNewCol = elements.refNewSelect.value;
      renderTargetCheckboxes();
    };

    renderTargetCheckboxes();
  }

  function renderTargetCheckboxes() {
    elements.targetColsContainer.innerHTML = '';
    state.headers.forEach(function (h) {
      if (h === state.mappings.refOldCol || h === state.mappings.refNewCol) {
        return;
      }

      const label = document.createElement('label');
      label.className = 'target-lang-item';

      const cb = document.createElement('input');
      cb.type = 'checkbox';
      cb.value = h;
      cb.checked = state.mappings.targetCols.includes(h);

      cb.addEventListener('change', function () {
        updateSelectedTargetCols();
      });

      label.appendChild(cb);
      label.appendChild(document.createTextNode(' ' + h));
      elements.targetColsContainer.appendChild(label);
    });
  }

  function updateSelectedTargetCols() {
    const selected = [];
    const cbs = elements.targetColsContainer.querySelectorAll('input[type="checkbox"]:checked');
    cbs.forEach(function (cb) {
      selected.push(cb.value);
    });
    state.mappings.targetCols = selected;
  }

  const GERMAN_TOKENS = ['de', 'deu', 'ger', 'german', 'deutsch', 'nemcina', 'němčina', 'referenz', 'reference', 'ref'];
  const OLD_TOKENS = ['2024', 'alt', 'old', 'stara', 'stará', 'puvodni', 'původní'];
  const NEW_TOKENS = ['2026', 'neu', 'new', 'nova', 'nová'];
  const LANGUAGE_TOKENS = ['en', 'cs', 'cz', 'sk', 'fr', 'es', 'it', 'pl', 'nl', 'hu', 'ro', 'pt', 'sv', 'da', 'fi', 'no', 'el',
    'bg', 'hr', 'sl', 'sr', 'lt', 'lv', 'et', 'ru', 'uk', 'tr', 'zh', 'ja', 'ko', 'ar', 'he',
    'english', 'czech', 'french', 'spanish', 'italian', 'polish', 'dutch', 'slovak', 'hungarian', 'romanian'];

  function headerTokens(header) {
    return String(header).toLowerCase().split(/[^\p{L}\p{N}]+/u).filter(Boolean);
  }

  function hasAnyToken(tokens, list) {
    return tokens.some(function (tok) { return list.indexOf(tok) !== -1; });
  }

  function autoDetectColumnMappings() {
    const headers = state.headers;
    let foundRefOld = '';
    let foundRefNew = '';
    const candidateTargets = [];

    headers.forEach(function (h) {
      const tokens = headerTokens(h);
      const isGerman = hasAnyToken(tokens, GERMAN_TOKENS) && !hasAnyToken(tokens, LANGUAGE_TOKENS);
      if (isGerman && !foundRefOld && hasAnyToken(tokens, OLD_TOKENS)) {
        foundRefOld = h;
      } else if (isGerman && !foundRefNew && hasAnyToken(tokens, NEW_TOKENS)) {
        foundRefNew = h;
      } else if (!isGerman && !hasAnyToken(tokens, NEW_TOKENS) &&
        (hasAnyToken(tokens, LANGUAGE_TOKENS) || (hasAnyToken(tokens, ['2024']) && tokens.length > 1))) {
        candidateTargets.push(h);
      }
    });

    if (foundRefOld) {
      state.mappings.refOldCol = foundRefOld;
      elements.refOldSelect.value = foundRefOld;
    }
    if (foundRefNew) {
      state.mappings.refNewCol = foundRefNew;
      elements.refNewSelect.value = foundRefNew;
    }

    state.mappings.targetCols = candidateTargets.filter(function (h) {
      return h !== state.mappings.refOldCol && h !== state.mappings.refNewCol;
    });
    renderTargetCheckboxes();

    if (foundRefOld && foundRefNew) {
      showToast(t('toastAutoDetectSuccess'), 'success');
    }
  }

  /**
   * Initialize translation stubs for target languages
   */
  function ensureStubsInitialized() {
    state.mappings.targetCols.forEach(function (targetCol) {
      // Rebuild every stub from the current rows/mapping and merge the saved work into it, so
      // reference texts, the 2024 target text and classifications are always up to date
      // (also after changing the reference columns or loading a compact project file).
      state.translations[targetCol] = ComparatorEngine.rehydrateStubs(
        state.rows,
        state.translations[targetCol],
        targetCol,
        state.mappings.refOldCol,
        state.mappings.refNewCol
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

      // Row ID / Section
      const idCol = state.headers[0] || 'ID';
      const sectionCol = state.headers[1] || '';
      const rowLabel = (row[idCol] ? row[idCol] + ' ' : '') + (row[sectionCol] ? '(' + row[sectionCol] + ')' : (t('rowLabel') + ' ') + (idx + 1));

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
        <span class="matrix-pill-lang">${ComparatorEngine.escapeHtml(col)}</span>
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
        <span>${ComparatorEngine.escapeHtml(col)}</span>
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
   * Column 4 diff of the translation (2024 → 2026). An empty editor is shown as "not translated
   * yet" rather than as the whole 2024 text being deleted (except for removed articles).
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
    const idCol = state.headers[0];
    const sectionCol = state.headers[1];

    let renderedCount = 0;

    stubs.forEach(function (stub, idx) {
      const row = state.rows[idx];
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
        const fullContent = (row[idCol] + ' ' + row[sectionCol] + ' ' + stub.refOld + ' ' + stub.refNew + ' ' + stub.targetOld + ' ' + stub.targetNew).toLowerCase();
        if (!fullContent.includes(q)) return;
      }

      renderedCount++;

      const card = document.createElement('div');
      card.className = 'article-card ' + (stub.isVerified ? 'status-verified' : 'status-pending');
      card.id = 'article-card-' + idx;

      const rowTitle = (row[idCol] ? row[idCol] : (t('rowLabel') + ' ') + (idx + 1));
      const sectionTitle = (row[sectionCol] ? row[sectionCol] : '');

      // 1. Reference Diff (German 2024 -> 2026)
      const refDiffRes = ComparatorEngine.computeDiff(stub.refOld, stub.refNew);
      const refDiffHtml = ComparatorEngine.renderDiffHtml(refDiffRes, 'inline');

      // 2. Target Diff (Target 2024 -> Target 2026)
      const targetDiffHtml = renderTargetDiffHtml(stub);

      // Card Header with optional search highlight
      const rowTitleHtml = highlightQuery(rowTitle, state.searchQuery);
      const sectionTitleHtml = highlightQuery(sectionTitle, state.searchQuery);

      const headerHtml = `
        <div class="article-header">
          <div class="article-meta">
            <span class="article-id">${rowTitleHtml}</span>
            ${sectionTitle ? `<span class="article-section">${sectionTitleHtml}</span>` : ''}
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

      // Column 1: Reference German Diff
      const col1RefHtml = `
        <div class="sub-panel">
          <div class="sub-panel-title">
            <span>${t('cardRefPanelTitle')}</span>
            ${refDiffRes.stats.hasChanges ? `<span style="font-size:0.75rem; color:#b45309; font-weight:600;">${t('changePct', { PCT: Math.round((1 - refDiffRes.stats.similarity) * 100) })}</span>` : `<span style="font-size:0.75rem; color:#15803d; font-weight:600;">${t('match100')}</span>`}
          </div>
          <div class="diff-box">${refDiffHtml || `<span style="color:var(--text-muted);">${t('emptyCell')}</span>`}</div>
        </div>
      `;

      // Column 2: Approved Original Target Translation (2024)
      const col2OldHtml = `
        <div class="sub-panel">
          <div class="sub-panel-title">
            <span>${t('cardTargetOldTitle', { LANG: col })}</span>
          </div>
          <div class="diff-box" style="background-color:var(--bg-tertiary); color:var(--text-secondary);">${ComparatorEngine.escapeHtml(stub.targetOld) || `<span style="color:var(--text-muted); font-style:italic;">${t('emptyOriginalText')}</span>`}</div>
        </div>
      `;

      // Column 3: Updated 2026 Target Translation (Editor)
      const col3NewEditorHtml = `
        <div class="sub-panel">
          <div class="sub-panel-title">
            <span>${t('cardTargetNewTitle', { LANG: col })}</span>
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
            <span>${t('cardTargetDiffTitle', { LANG: col })}</span>
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

    // Prefill from 2024
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

    // Prefill from German reference
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
    const row = state.rows[idx];
    const idCol = state.headers[0];
    const sectionCol = state.headers[1];

    const rowId = (row[idCol] ? row[idCol] : (t('rowLabel') + ' ') + (idx + 1)) + (row[sectionCol] ? ' (' + row[sectionCol] + ')' : '');
    const refDiffRes = ComparatorEngine.computeDiff(stub.refOld, stub.refNew);

    const promptText = ComparatorEngine.buildCopilotPrompt({
      promptLang: I18n.getLanguage(),
      targetLangName: col,
      rowId: rowId,
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
   */
  function updatePdfLanguageDropdown() {
    if (!elements.pdfLangSelect) return;
    elements.pdfLangSelect.innerHTML = '';

    // 1. Reference German 2026
    if (state.mappings.refNewCol) {
      const optRef = document.createElement('option');
      optRef.value = state.mappings.refNewCol;
      optRef.textContent = `🇩🇪 ${state.mappings.refNewCol} (${t('pdfOptionRefNew')})`;
      elements.pdfLangSelect.appendChild(optRef);
    }

    // 2. Target language editions
    if (state.mappings.targetCols && state.mappings.targetCols.length > 0) {
      state.mappings.targetCols.forEach(function (col) {
        const opt = document.createElement('option');
        opt.value = col;
        const cleanName = col.replace(/2024/i, '2026');
        opt.textContent = `🌐 ${cleanName}`;
        elements.pdfLangSelect.appendChild(opt);
      });
    }

    if (state.activeTargetCol && state.mappings.targetCols.includes(state.activeTargetCol)) {
      elements.pdfLangSelect.value = state.activeTargetCol;
    }
  }

  /**
   * Export Managerial Summary (Markdown, Word .doc, PDF)
   */
  function exportManagerialSummary(format) {
    if (!state.rows || state.rows.length === 0 || !state.mappings.refOldCol || !state.mappings.refNewCol) {
      showToast(t('alertSelectRefCols'), 'warning');
      return;
    }

    const companyName = (state.pdfTemplate && state.pdfTemplate.companyName)
      ? state.pdfTemplate.companyName
      : (state.fileName ? state.fileName.replace(/\.[^/.]+$/, '') : 'Corporate Group');

    const options = {
      rows: state.rows,
      headers: state.headers,
      mappings: state.mappings,
      companyName: companyName,
      uiLang: I18n.getLanguage(),
      template: state.pdfTemplate
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

  function initExports() {
    elements.downloadComprehensiveBtn.addEventListener('click', function () {
      downloadExcelWorkbook('comprehensive');
    });

    elements.downloadCleanBtn.addEventListener('click', function () {
      downloadExcelWorkbook('clean_2026');
    });

    elements.exportProjectJsonBtn.addEventListener('click', function () {
      exportProjectFile();
    });

    // Managerial Summary Exports (Step 2 and Step 4)
    if (elements.step2ExportMdBtn) {
      elements.step2ExportMdBtn.addEventListener('click', function () { exportManagerialSummary('markdown'); });
    }
    if (elements.step2ExportDocBtn) {
      elements.step2ExportDocBtn.addEventListener('click', function () { exportManagerialSummary('word'); });
    }
    if (elements.step2ExportPdfBtn) {
      elements.step2ExportPdfBtn.addEventListener('click', function () { exportManagerialSummary('pdf'); });
    }
    if (elements.downloadMgrMdBtn) {
      elements.downloadMgrMdBtn.addEventListener('click', function () { exportManagerialSummary('markdown'); });
    }
    if (elements.downloadMgrDocBtn) {
      elements.downloadMgrDocBtn.addEventListener('click', function () { exportManagerialSummary('word'); });
    }
    if (elements.printMgrPdfBtn) {
      elements.printMgrPdfBtn.addEventListener('click', function () { exportManagerialSummary('pdf'); });
    }

    // PDF Print / Save as PDF
    if (elements.printPdfBtn) {
      elements.printPdfBtn.addEventListener('click', function () {
        const selectedCol = elements.pdfLangSelect ? elements.pdfLangSelect.value : state.activeTargetCol;
        if (!selectedCol) {
          showToast(t('toastPdfSelectLang'), 'warning');
          return;
        }

        PdfExporter.printLanguageMutation({
          rows: state.rows,
          headers: state.headers,
          mappings: state.mappings,
          translations: state.translations,
          targetCol: selectedCol,
          companyName: state.fileName ? state.fileName.replace(/\.[^/.]+$/, '') : 'Corporate Group',
          uiLang: I18n.getLanguage(),
          template: state.pdfTemplate
        });
      });
    }

    // PDF Download HTML
    if (elements.downloadPdfHtmlBtn) {
      elements.downloadPdfHtmlBtn.addEventListener('click', function () {
        const selectedCol = elements.pdfLangSelect ? elements.pdfLangSelect.value : state.activeTargetCol;
        if (!selectedCol) {
          showToast(t('toastPdfSelectLang'), 'warning');
          return;
        }

        PdfExporter.downloadPrintableHtml({
          rows: state.rows,
          headers: state.headers,
          mappings: state.mappings,
          translations: state.translations,
          targetCol: selectedCol,
          companyName: state.fileName ? state.fileName.replace(/\.[^/.]+$/, '') : 'Corporate Group',
          uiLang: I18n.getLanguage(),
          template: state.pdfTemplate
        });
        showToast(t('toastPdfHtmlDownloaded'), 'success');
      });
    }

    // Batch Print All Languages
    if (elements.printAllPdfBtn) {
      elements.printAllPdfBtn.addEventListener('click', function () {
        const allCols = [];
        if (state.mappings.refNewCol) allCols.push(state.mappings.refNewCol);
        if (state.mappings.targetCols) allCols.push(...state.mappings.targetCols);

        if (allCols.length === 0) {
          showToast(t('toastNoLanguages'), 'warning');
          return;
        }

        if (confirm(t('confirmPrintAll', { COUNT: allCols.length }))) {
          allCols.forEach(function (col, i) {
            setTimeout(function () {
              PdfExporter.printLanguageMutation({
                rows: state.rows,
                headers: state.headers,
                mappings: state.mappings,
                translations: state.translations,
                targetCol: col,
                companyName: state.fileName ? state.fileName.replace(/\.[^/.]+$/, '') : 'Corporate Group',
                uiLang: I18n.getLanguage(),
                template: state.pdfTemplate
              });
            }, i * 400);
          });
        }
      });
    }
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
      targetLangName: col,
      rows: state.rows,
      headers: state.headers,
      mappings: state.mappings,
      translations: state.translations,
      filterType: 'all_modified'
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
        const idCol = state.headers[0];

        const expectedList = stubs.map(function (stub, idx) {
          const row = state.rows[idx] || {};
          return {
            index: idx,
            id: row[idCol] || ('Art. ' + (idx + 1))
          };
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
          accentColor: /^#[0-9a-f]{3,8}$/i.test(elements.tplAccentColor.value.trim()) ? elements.tplAccentColor.value.trim() : '#1e3a8a',
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
    const isEn = I18n.getLanguage() === 'en';
    const defaultCompany = state.fileName ? state.fileName.replace(/\.[^/.]+$/, '') : 'Corporate Group';
    const tpl = state.pdfTemplate || {};

    elements.tplCompanyName.value = tpl.companyName !== undefined ? tpl.companyName : defaultCompany;
    elements.tplEffectiveDate.value = tpl.effectiveDate !== undefined ? tpl.effectiveDate : (isEn ? 'Effective Date: January 1, 2026' : 'Účinnost od: 1. ledna 2026');
    elements.tplDocTitle.value = tpl.docTitle !== undefined ? tpl.docTitle : (isEn ? 'CODE OF CONDUCT & ETHICS 2026' : 'ETICKÝ KODEX A PRAVIDLA CHOVÁNÍ 2026');
    elements.tplAccentColor.value = tpl.accentColor || '#1e3a8a';
    elements.tplAccentColorPicker.value = tpl.accentColor || '#1e3a8a';
    elements.tplDocSubtitle.value = tpl.docSubtitle !== undefined ? tpl.docSubtitle : (isEn ? 'Official Certified Edition – {LANG}' : 'Oficiální certifikované znění – {LANG}');
    if (elements.tplVersionLabel) {
      elements.tplVersionLabel.value = tpl.versionLabel !== undefined ? tpl.versionLabel : (isEn ? 'Version 2.6 – Compliance Approved' : 'Verze 2.6 – Schváleno compliance');
    }
    elements.tplIntroPreamble.value = tpl.introPreamble !== undefined ? tpl.introPreamble : (isEn 
      ? 'This document represents the official and binding Code of Conduct and Ethics for the 2026 cycle. All provisions herein supersede previous editions and apply to all operational entities and employees.'
      : 'Tento dokument představuje oficiální a závazný Etický kodex pro období od roku 2026. Veškerá zde uvedená ustanovení nahrazují dřívější znění a jsou závazná pro všechny zaměstnance i vedoucí pracovníky.');
    elements.tplShowSignatures.checked = tpl.showSignatures !== false;
    elements.signaturesBlockInputs.style.display = elements.tplShowSignatures.checked ? 'grid' : 'none';
    elements.tplSignDept1.value = tpl.signDept1 !== undefined ? tpl.signDept1 : (isEn ? 'Corporate Compliance & Legal' : 'Oddělení Compliance & Právní');
    elements.tplSignDept2.value = tpl.signDept2 !== undefined ? tpl.signDept2 : (isEn ? 'Executive Board' : 'Vedení společnosti');

    elements.pdfTemplateModal.classList.add('open');
  }

  function downloadExcelWorkbook(exportType) {
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
      mappings: state.mappings,
      translations: state.translations,
      exportType: exportType,
      uiLang: I18n.getLanguage()
    });

    const baseName = (state.fileName || 'Eticky_Kodex').replace(/\.[^/.]+$/, '').replace(/[\\/:*?"<>|]+/g, '_');
    const suffix = exportType === 'clean_2026' ? '2026_Final' : '2024_2026_Srovnani';
    const outputFileName = `${baseName}_${suffix}.xlsx`;

    XLSX.writeFile(result.workbook, outputFileName);
    showToast(t('toastWorkbookGenerated', { NAME: outputFileName }), 'success');
    if (result.unverifiedCount > 0) {
      const key = exportType === 'clean_2026' ? 'toastExportUnverifiedClean' : 'toastExportUnverified';
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
    showToast(t('toastProjectSaved'), 'success');
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
      assertEq('Stub generation for 15 sample articles', stubs.length, 15);
      assertEq('Unchanged Article 1 marked auto_copied', stubs[0].status, 'auto_copied');

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
      assertOk('buildCopilotPrompt generates English prompt', promptEn.includes('expert legal') && promptEn.includes('CRITICAL INSTRUCTIONS'));

      // 6. Export prep
      const exportData = ComparatorEngine.prepareExportData({
        rows: sampleRows,
        headers: SampleGenerator.getSampleHeaders(),
        mappings: SampleGenerator.getSampleDefaultMappings(),
        translations: { 'CS 2024': stubs },
        exportType: 'clean_2026'
      });
      assertOk('clean_2026 export filters deleted articles', exportData.data.length < sampleRows.length);

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
      return 'You are an expert corporate translator specializing in Codes of Ethics.\n\nTASK: Update the approved translation for 2026 in language: {TARGET_LANG}.\nIdentifier: {ROW_ID}\n\n1. GERMAN SOURCE (2024):\n{REF_OLD}\n\n2. GERMAN SOURCE (2026):\n{REF_NEW}\n\n3. APPROVED TRANSLATION (2024):\n{TARGET_OLD}\n\nINSTRUCTIONS:\nMake ONLY minimal necessary alterations. Preserve existing tone and terminology. Output only the final updated text.';
    }
    return 'Jste špičkový odborný překladatel firemních a právních dokumentů (Etický kodex).\n\nÚKOL: Aktualizujte stávající schválený překlad odstavce pro rok 2026 v jazyce: {TARGET_LANG}.\nIdentifikátor: {ROW_ID}\n\n1. PŮVODNÍ NĚMECKÝ TEXT (2024):\n{REF_OLD}\n\n2. NOVÝ NĚMECKÝ TEXT (2026):\n{REF_NEW}\n\n3. DOSAVADNÍ OVĚŘENÝ PŘEKLAD (2024):\n{TARGET_OLD}\n\nPOKYNY:\nProveďte POUZE minimální nutné změny odpovídající německému rozdílu. Zachovejte zavedenou terminologii a styl. Vraťte výhradně čistý přeložený text.';
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
    const mappings = data.mappings || {};
    state.mappings = {
      refOldCol: mappings.refOldCol || '',
      refNewCol: mappings.refNewCol || '',
      targetCols: Array.isArray(mappings.targetCols) ? mappings.targetCols.slice() : []
    };
    state.translations = data.translations && typeof data.translations === 'object' ? data.translations : {};
    state.customPromptTemplate = data.customPromptTemplate || null;
    state.pdfTemplate = data.pdfTemplate || null;
    state.activeTargetCol = data.activeTargetCol || state.mappings.targetCols[0] || '';
    elements.sheetSelectWrapper.style.display = 'none';

    renderColumnMappingControls();
    elements.refOldSelect.value = state.mappings.refOldCol;
    elements.refNewSelect.value = state.mappings.refNewCol;

    const checkboxes = elements.targetColsContainer.querySelectorAll('input[type="checkbox"]');
    checkboxes.forEach(function (cb) {
      cb.checked = state.mappings.targetCols.includes(cb.value);
    });

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
      updateSelectedTargetCols();

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

      setStep(2);
      showToast(t('toastMappingSaved'), 'success');
    });

    elements.autoDetectMappingBtn.addEventListener('click', function () {
      autoDetectColumnMappings();
    });

    elements.goToTranslationBtn.addEventListener('click', function () {
      setStep(3);
    });

    elements.goToExportBtn.addEventListener('click', function () {
      setStep(4);
    });

    initModalDismissal();

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
