/**
 * Step 3: batch prompt, Smart Paste and import of translations from a workbook.
 * Part of the application controller; the parts share the `App` namespace and are loaded in the
 * order given in index.html (js/app/core.js first, js/app/main.js last).
 */
(function (App) {
  'use strict';

  const { columnDisplayName, elements, fallbackCopyText, getMappingError, languageLabel, persistState,
    showToast, state, t, ensureStubsInitialized, renderCrossLanguageMatrix, renderLanguageTabs,
    renderTranslationStudio, updateLanguageProgress, renderArticleCards } = App;

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

  Object.assign(App, { initBatchPromptModal: initBatchPromptModal, openBatchPromptModal: openBatchPromptModal,
    initSmartPasteModal: initSmartPasteModal, initTranslationImport: initTranslationImport,
    importTranslationsFile: importTranslationsFile });
}(window.App = window.App || {}));
