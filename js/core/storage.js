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

  const STORAGE_KEY = 'ethic_comparator_workspace_v1';
  const PROJECT_APP_ID = 'EthicCodeComparator';
  const PROJECT_VERSION = '1.1.0';
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
      headers: state.headers,
      rows: state.rows,
      mappings: state.mappings,
      translations: compactTranslations(state.translations), // { [targetCol]: [ { classification, targetNew, status, isVerified, notes } ] }
      customPromptTemplate: state.customPromptTemplate,
      pdfTemplate: state.pdfTemplate,
      activeTargetCol: state.activeTargetCol
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
      const raw = window.localStorage.getItem(STORAGE_KEY);
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
      window.localStorage.removeItem(STORAGE_KEY);
    } catch (e) {
      console.warn('Chyba při mazání LocalStorage:', e);
    }
  }

  function serializeProjectJson(state) {
    const persisted = buildPersistedState(state);
    const projectData = {
      app: PROJECT_APP_ID,
      version: PROJECT_VERSION,
      exportedAt: new Date().toISOString(),
      fileName: persisted.fileName || 'kodex_projekt',
      activeSheet: persisted.activeSheet,
      headers: persisted.headers,
      rows: persisted.rows,
      mappings: persisted.mappings,
      translations: persisted.translations,
      customPromptTemplate: persisted.customPromptTemplate,
      pdfTemplate: persisted.pdfTemplate,
      activeTargetCol: persisted.activeTargetCol
    };
    return JSON.stringify(projectData, null, 2);
  }

  function getProjectFileName(state, date) {
    const safeName = (state.fileName || 'kodex')
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
   * Parses and validates a project file (or an autosaved state). Throws a descriptive error
   * for anything that would otherwise break the application when restored.
   */
  function parseProjectJson(jsonContent) {
    let parsed;
    try {
      parsed = typeof jsonContent === 'string'
        ? JSON.parse(jsonContent.replace(/^﻿/, ''))
        : jsonContent;
    } catch (e) {
      throw new Error('Chyba při čtení JSON projektu: soubor není platný JSON (' + e.message + ').');
    }

    function fail(reason) {
      throw new Error('Chyba při čtení JSON projektu: ' + reason);
    }

    if (!parsed || typeof parsed !== 'object' || Array.isArray(parsed)) {
      fail('neplatný formát souboru projektu.');
    }
    if (parsed.app && parsed.app !== PROJECT_APP_ID) {
      fail('soubor nepatří této aplikaci (' + parsed.app + ').');
    }
    if (!Array.isArray(parsed.rows) || parsed.rows.length === 0) {
      fail('chybí řádky tabulky (rows).');
    }
    if (!parsed.mappings || typeof parsed.mappings !== 'object') {
      fail('chybí mapování sloupců (mappings).');
    }

    const headers = Array.isArray(parsed.headers) && parsed.headers.length > 0
      ? parsed.headers
      : Object.keys(parsed.rows[0] || {});
    const mappings = {
      refOldCol: parsed.mappings.refOldCol || '',
      refNewCol: parsed.mappings.refNewCol || '',
      targetCols: Array.isArray(parsed.mappings.targetCols) ? parsed.mappings.targetCols.slice() : []
    };
    [mappings.refOldCol, mappings.refNewCol].concat(mappings.targetCols).forEach(function (col) {
      if (col && headers.indexOf(col) === -1) {
        fail('mapovaný sloupec "' + col + '" v tabulce neexistuje.');
      }
    });

    const translations = {};
    if (parsed.translations && typeof parsed.translations === 'object') {
      Object.keys(parsed.translations).forEach(function (col) {
        if (Array.isArray(parsed.translations[col])) {
          translations[col] = parsed.translations[col];
        }
      });
    }

    const result = {};
    Object.keys(parsed).forEach(function (k) { result[k] = parsed[k]; });
    result.headers = headers;
    result.mappings = mappings;
    result.translations = translations;
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
    parseProjectJson: parseProjectJson
  };
}));
