/**
 * Project file (save / load / validate), autosave restore and "new project".
 * Part of the application controller; the parts share the `App` namespace and are loaded in the
 * order given in index.html (js/app/core.js first, js/app/main.js last).
 */
(function (App) {
  'use strict';

  const { elements, getMappingError, persistState, PREVIEW_INITIAL_ROWS, showToast, state, t, setStep } = App;

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

  function exportProjectFile() {
    if (!state.rows || state.rows.length === 0) {
      showToast(t('toastNoData'), 'warning');
      return null;
    }
    // Make sure the file contains the latest edits and stubs
    if (state.mappings.targetCols.length > 0 && !getMappingError(true)) {
      App.ensureStubsInitialized();
    }
    const fileName = StorageManager.exportProjectAsJson(state);
    showToast(t('toastProjectSaved', { NAME: fileName }), 'success');
    return fileName;
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

    App.renderColumnMappingControls();
    App.renderDataPreview();

    // An incomplete mapping (e.g. autosave made while still in Step 1) resumes in Step 1
    if (getMappingError(true)) {
      setStep(1);
      return;
    }

    App.ensureStubsInitialized();

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

  Object.assign(App, { initProjectButtons: initProjectButtons, loadProjectFile: loadProjectFile,
    resetWorkspaceForNewTable: resetWorkspaceForNewTable, exportProjectFile: exportProjectFile,
    initProjectReset: initProjectReset, restoreProjectFromData: restoreProjectFromData,
    checkStoredStateOnLoad: checkStoredStateOnLoad });
}(window.App = window.App || {}));
