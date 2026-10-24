/**
 * Application start: wires all parts together once the page is loaded.
 * Part of the application controller; the parts share the `App` namespace and are loaded in the
 * order given in index.html (js/app/core.js first, js/app/main.js last).
 */
(function (App) {
  'use strict';

  const { elements, getMappingError, isDirectTranslationMode, showToast, state, t, initLanguage,
    initLayoutMode, initModalDismissal, initStickyHeader, initTheme, setStep, exportProjectFile,
    checkStoredStateOnLoad, initProjectButtons, initProjectReset, loadProjectFile, restoreProjectFromData,
    initFileUpload, processUploadedFile, addNewLanguage, finalizeLanguageConfig, removeNewLanguage,
    approveAllUnchangedGlobal, approveUnchangedInActiveLanguage, initAuxLangSelector, initFilters,
    updateAuxLangOptions, importTranslationsFile, initBatchPromptModal, initSmartPasteModal,
    initTranslationImport, openBatchPromptModal, exportManagerialSummary, exportTables, initExports,
    initPdfTemplateSettings, openPdfTemplateModal, renderTableFormatOptions, initDiagnostics,
    initPromptSettings } = App;

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
}(window.App = window.App || {}));
