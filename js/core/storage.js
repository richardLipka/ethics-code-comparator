/**
 * StorageManager - Handles autosave to LocalStorage and JSON project import/export.
 * Universal module: works in browser and Node.js.
 */
(function (root, factory) {
  if (typeof define === 'function' && define.amd) {
    define([], factory);
  } else if (typeof module === 'object' && module.exports) {
    module.exports = factory();
  } else {
    root.StorageManager = factory();
  }
}(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  const STORAGE_KEY = 'i18n_manager_workspace_v1';
  const PROJECT_APP_ID = 'InternationalisationManager';
  const PROJECT_VERSION = '1.3.0';
  // Identifiers used before the project was renamed; still read so older work loads
  const LEGACY_STORAGE_KEYS = ['ethic_comparator_workspace_v1'];
  const LEGACY_APP_IDS = ['EthicCodeComparator'];
  let debounceTimer = null;
  let pendingSave = null; // { state, callback } of the debounced save that has not run yet

  function isLocalStorageAvailable() {
    try {
      const test = '__storage_test__';
      if (typeof window === 'undefined' || !window.localStorage) return false;
      window.localStorage.setItem(test, test);
      window.localStorage.removeItem(test);
      return true;
    } catch (e) {
      return false;
    }
  }

  /**
   * Keeps only the per-row work of each translation stub. Reference texts, the 2024 target
   * text and the row index are rebuilt from `rows` on load, so they are not stored once per
   * language (this keeps project files and the autosave far below the LocalStorage quota).
   */
  function compactTranslations(translations) {
    const out = {};
    Object.keys(translations || {}).forEach(function (col) {
      const stubs = translations[col];
      if (!Array.isArray(stubs)) return;
      out[col] = stubs.map(function (s) {
        if (!s) return null;
        const item = {
          classification: s.classification,
          targetNew: s.targetNew || '',
          status: s.status,
          isVerified: !!s.isVerified
        };
        if (s.notes) item.notes = s.notes;
        return item;
      });
    });
    return out;
  }

  function buildPersistedState(state) {
    return {
      fileName: state.fileName,
      activeSheet: state.activeSheet,
      hasHeaderRow: state.hasHeaderRow !== false,
      headers: state.headers,
      rows: state.rows,
      rowNumbers: state.rowNumbers,
      mappings: state.mappings, // { refOldCol, refNewCol, targetCols, languages: { [oldCol]: { label, newCol, newColName } } }
      translations: compactTranslations(state.translations), // { [targetCol]: [ { classification, targetNew, status, isVerified, notes } ] }
      customPromptTemplate: state.customPromptTemplate,
      pdfTemplate: state.pdfTemplate,
      activeTargetCol: state.activeTargetCol,
      exportLangCol: state.exportLangCol,
      partialContextCols: state.partialContextCols
    };
  }

  function saveState(state) {
    if (!isLocalStorageAvailable()) return false;
    try {
      const data = buildPersistedState(state);
      data.timestamp = Date.now();
      window.localStorage.setItem(STORAGE_KEY, JSON.stringify(data));
      return true;
    } catch (e) {
      console.warn('Nepodařilo se uložit stav do LocalStorage (možné překročení kvóty):', e);
      return false;
    }
  }

  /**
   * Debounced autosave. `callback(ok)` is called once the save actually ran.
   */
  function debouncedSaveState(state, delayMs, callback) {
    const delay = delayMs !== undefined ? delayMs : 500;
    if (debounceTimer) {
      clearTimeout(debounceTimer);
    }
    pendingSave = { state: state, callback: callback };
    debounceTimer = setTimeout(flushPendingSave, delay);
  }

  /**
   * Runs a pending debounced save immediately (e.g. when the page is being closed).
   */
  function flushPendingSave() {
    if (debounceTimer) {
      clearTimeout(debounceTimer);
      debounceTimer = null;
    }
    if (!pendingSave) return null;
    const job = pendingSave;
    pendingSave = null;
    const ok = saveState(job.state);
    if (typeof job.callback === 'function') job.callback(ok);
    return ok;
  }

  function loadState() {
    if (!isLocalStorageAvailable()) return null;
    try {
      let raw = window.localStorage.getItem(STORAGE_KEY);
      for (let i = 0; !raw && i < LEGACY_STORAGE_KEYS.length; i++) {
        raw = window.localStorage.getItem(LEGACY_STORAGE_KEYS[i]);
      }
      if (!raw) return null;
      return JSON.parse(raw);
    } catch (e) {
      console.warn('Nepodařilo se načíst stav z LocalStorage:', e);
      return null;
    }
  }

  function clearState() {
    if (debounceTimer) {
      clearTimeout(debounceTimer);
      debounceTimer = null;
    }
    pendingSave = null;
    if (!isLocalStorageAvailable()) return;
    try {
      [STORAGE_KEY].concat(LEGACY_STORAGE_KEYS).forEach(function (key) {
        window.localStorage.removeItem(key);
      });
    } catch (e) {
      console.warn('Chyba při mazání LocalStorage:', e);
    }
  }

  function serializeProjectJson(state) {
    const persisted = buildPersistedState(state);
    const projectData = Object.assign({
      app: PROJECT_APP_ID,
      version: PROJECT_VERSION,
      exportedAt: new Date().toISOString()
    }, persisted);
    projectData.fileName = persisted.fileName || 'projekt';
    return JSON.stringify(projectData, null, 2);
  }

  function getProjectFileName(state, date) {
    const safeName = (state.fileName || 'dokument')
      .replace(/\.[^/.]+$/, '')
      .replace(/[\\/:*?"<>|]+/g, '_');
    return `${safeName}_projekt_${(date || new Date()).toISOString().slice(0, 10)}.json`;
  }

  function downloadBlob(blob, fileName) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = fileName;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    // Revoking synchronously can cancel the download in some browsers
    setTimeout(function () { URL.revokeObjectURL(url); }, 1000);
  }

  function exportProjectAsJson(state) {
    const jsonString = serializeProjectJson(state);
    const blob = new Blob([jsonString], { type: 'application/json;charset=utf-8' });
    const fileName = getProjectFileName(state);
    downloadBlob(blob, fileName);
    return fileName;
  }

  /**
   * Error thrown for a file that cannot be loaded as a project. `code` identifies the problem
   * (the UI shows a translated message), `params` carries details for that message.
   *   INVALID_JSON     – not JSON at all (e.g. a spreadsheet picked by mistake)
   *   NOT_A_PROJECT    – JSON with a different structure (no table and column selection)
   *   FOREIGN_APP      – a JSON project of another application
   *   NO_ROWS          – a project without table rows
   *   BAD_ROWS         – rows are not table rows (objects)
   *   BAD_MAPPINGS     – the column selection is not usable
   *   MISSING_COLUMN   – a selected column does not exist in the table
   */
  function projectError(code, message, params) {
    const err = new Error('Project file: ' + message);
    err.code = code;
    err.params = params || {};
    return err;
  }

  function isPlainObject(value) {
    return !!value && typeof value === 'object' && !Array.isArray(value);
  }

  function compareVersions(a, b) {
    const pa = String(a || '0').split('.').map(function (n) { return parseInt(n, 10) || 0; });
    const pb = String(b || '0').split('.').map(function (n) { return parseInt(n, 10) || 0; });
    for (let i = 0; i < Math.max(pa.length, pb.length); i++) {
      if ((pa[i] || 0) !== (pb[i] || 0)) return (pa[i] || 0) - (pb[i] || 0);
    }
    return 0;
  }

  /**
   * Parses and validates a project file (or an autosaved state).
   * Throws a projectError for anything that would otherwise break the application. Problems that
   * do not prevent loading are returned in `result.warnings` ([{ code, params }]):
   *   NO_APP_ID, NEWER_VERSION, UNUSED_TRANSLATIONS, ROW_COUNT_MISMATCH, INVALID_TRANSLATIONS,
   *   INVALID_SETTINGS
   */
  function parseProjectJson(jsonContent) {
    let parsed;
    if (typeof jsonContent === 'string') {
      try {
        parsed = JSON.parse(jsonContent.replace(/^﻿/, ''));
      } catch (e) {
        throw projectError('INVALID_JSON', 'not valid JSON (' + e.message + ')', { MSG: e.message });
      }
    } else {
      parsed = jsonContent;
    }

    const warnings = [];
    const hasProjectShape = isPlainObject(parsed) && ('rows' in parsed || 'mappings' in parsed);
    const isOwnApp = isPlainObject(parsed) && (parsed.app === PROJECT_APP_ID || LEGACY_APP_IDS.indexOf(parsed.app) !== -1);

    if (isPlainObject(parsed) && parsed.app && !isOwnApp) {
      throw projectError('FOREIGN_APP', 'belongs to another application (' + parsed.app + ')', { APP: String(parsed.app) });
    }
    if (!isOwnApp && !hasProjectShape) {
      const found = isPlainObject(parsed) ? Object.keys(parsed).slice(0, 6).join(', ') : (Array.isArray(parsed) ? '[…]' : typeof parsed);
      throw projectError('NOT_A_PROJECT', 'different structure, not a translation project', { FOUND: found || '{}' });
    }
    if (!isOwnApp) {
      warnings.push({ code: 'NO_APP_ID', params: {} });
    }
    if (isOwnApp && parsed.version && compareVersions(parsed.version, PROJECT_VERSION) > 0) {
      warnings.push({ code: 'NEWER_VERSION', params: { VERSION: String(parsed.version) } });
    }

    if (!Array.isArray(parsed.rows) || parsed.rows.length === 0) {
      throw projectError('NO_ROWS', 'no table rows (rows)');
    }
    const badRow = parsed.rows.findIndex(function (r) { return !isPlainObject(r); });
    if (badRow !== -1) {
      throw projectError('BAD_ROWS', 'row ' + (badRow + 1) + ' is not a table row', { ROW: badRow + 1 });
    }
    if (!isPlainObject(parsed.mappings)) {
      throw projectError('BAD_MAPPINGS', 'no column selection (mappings)');
    }

    const headers = Array.isArray(parsed.headers) && parsed.headers.length > 0
      ? parsed.headers.map(String)
      : Object.keys(parsed.rows[0]);
    const m = parsed.mappings;
    const isColumnName = function (v) { return v === undefined || v === null || typeof v === 'string'; };
    if (!isColumnName(m.refOldCol) || !isColumnName(m.refNewCol) ||
      (m.targetCols !== undefined && (!Array.isArray(m.targetCols) || m.targetCols.some(function (c) { return typeof c !== 'string'; })))) {
      throw projectError('BAD_MAPPINGS', 'column selection has an unexpected format');
    }
    const mappings = {
      refOldCol: m.refOldCol || '',
      refNewCol: m.refNewCol || '',
      targetCols: Array.isArray(m.targetCols) ? m.targetCols.slice() : [],
      languages: {}
    };
    const savedLanguages = isPlainObject(m.languages) ? m.languages : {};
    mappings.targetCols.forEach(function (col) {
      const lang = isPlainObject(savedLanguages[col]) ? savedLanguages[col] : {};
      mappings.languages[col] = {
        label: typeof lang.label === 'string' ? lang.label : '',
        newCol: typeof lang.newCol === 'string' ? lang.newCol : '',
        newColName: typeof lang.newColName === 'string' ? lang.newColName : ''
      };
    });
    const languageNewCols = mappings.targetCols.map(function (col) { return mappings.languages[col].newCol; });
    [mappings.refOldCol, mappings.refNewCol].concat(mappings.targetCols, languageNewCols).forEach(function (col) {
      if (col && headers.indexOf(col) === -1) {
        throw projectError('MISSING_COLUMN', 'selected column "' + col + '" does not exist in the table', { COL: col });
      }
    });

    // Translation work: only for selected languages, as lists of row entries
    const translations = {};
    const unused = [];
    let invalidTranslations = false;
    if (parsed.translations !== undefined && !isPlainObject(parsed.translations)) {
      invalidTranslations = true;
    } else if (parsed.translations) {
      Object.keys(parsed.translations).forEach(function (col) {
        const list = parsed.translations[col];
        if (mappings.targetCols.indexOf(col) === -1) {
          if (Array.isArray(list) && list.some(function (s) { return s && (s.targetNew || s.isVerified === true); })) unused.push(col);
          return;
        }
        if (!Array.isArray(list)) {
          invalidTranslations = true;
          return;
        }
        if (list.some(function (s) { return s !== null && !isPlainObject(s); })) invalidTranslations = true;
        translations[col] = list.map(function (s) { return isPlainObject(s) ? s : null; });
        if (list.length !== parsed.rows.length) {
          const label = mappings.languages[col].label || col;
          warnings.push({ code: 'ROW_COUNT_MISMATCH', params: { LANG: label, COUNT: list.length, ROWS: parsed.rows.length } });
        }
      });
    }
    if (unused.length > 0) warnings.push({ code: 'UNUSED_TRANSLATIONS', params: { LIST: unused.join(', ') } });
    if (invalidTranslations) warnings.push({ code: 'INVALID_TRANSLATIONS', params: {} });

    const result = {};
    Object.keys(parsed).forEach(function (k) { result[k] = parsed[k]; });
    result.headers = headers;
    result.mappings = mappings;
    result.translations = translations;
    result.rowNumbers = Array.isArray(parsed.rowNumbers) && parsed.rowNumbers.length === parsed.rows.length &&
      parsed.rowNumbers.every(function (n) { return typeof n === 'number'; })
      ? parsed.rowNumbers
      : parsed.rows.map(function (r, idx) { return idx + 2; });

    // Optional settings of an unexpected type are dropped instead of breaking the UI
    let invalidSettings = false;
    if (result.pdfTemplate !== undefined && result.pdfTemplate !== null && !isPlainObject(result.pdfTemplate)) {
      result.pdfTemplate = null;
      invalidSettings = true;
    }
    if (result.customPromptTemplate !== undefined && result.customPromptTemplate !== null && typeof result.customPromptTemplate !== 'string') {
      result.customPromptTemplate = null;
      invalidSettings = true;
    }
    if (result.partialContextCols !== undefined && result.partialContextCols !== null &&
      (!Array.isArray(result.partialContextCols) || result.partialContextCols.some(function (c) { return typeof c !== 'string'; }))) {
      result.partialContextCols = null;
      invalidSettings = true;
    }
    if (invalidSettings) warnings.push({ code: 'INVALID_SETTINGS', params: {} });

    result.warnings = warnings;
    return result;
  }

  return {
    isLocalStorageAvailable: isLocalStorageAvailable,
    saveState: saveState,
    debouncedSaveState: debouncedSaveState,
    flushPendingSave: flushPendingSave,
    loadState: loadState,
    clearState: clearState,
    compactTranslations: compactTranslations,
    serializeProjectJson: serializeProjectJson,
    exportProjectJson: serializeProjectJson,
    exportProjectAsJson: exportProjectAsJson,
    getProjectFileName: getProjectFileName,
    downloadBlob: downloadBlob,
    parseProjectJson: parseProjectJson,
    PROJECT_VERSION: PROJECT_VERSION
  };
}));
