/**
 * Step 3: paragraph cards – rendering, editing, approval and per-paragraph prompts.
 * Part of the application controller; the parts share the `App` namespace and are loaded in the
 * order given in index.html (js/app/core.js first, js/app/main.js last).
 */
(function (App) {
  'use strict';

  const { columnDisplayName, elements, fallbackCopyText, isDirectTranslationMode, languageLabel, persistState,
    rowLabelOf, showToast, state, t, renderBadgeHtml, highlightQuery, jumpToPending,
    renderCrossLanguageMatrix, renderLanguageTabs, updateLanguageProgress } = App;

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

  Object.assign(App, { renderArticleCards: renderArticleCards });
}(window.App = window.App || {}));
