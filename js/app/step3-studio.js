/**
 * Step 3: translation studio – progress, language tabs, auxiliary language, filters and navigation.
 * Part of the application controller; the parts share the `App` namespace and are loaded in the
 * order given in index.html (js/app/core.js first, js/app/main.js last).
 */
(function (App) {
  'use strict';

  const { elements, languageLabel, persistState, showToast, state, t, updateLayoutToggleText } = App;

  /**
   * Step 3: Render Translation Studio (Full-width side-by-side)
   */
  function renderTranslationStudio() {
    updateAuxLangOptions();
    updateLayoutToggleText();
    renderCrossLanguageMatrix();
    renderLanguageTabs();
    App.renderArticleCards();
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
        App.renderArticleCards();
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

    App.renderArticleCards();
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

    App.renderArticleCards();
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
        App.renderArticleCards();
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
      App.renderArticleCards();
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
        App.renderArticleCards();
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
        App.renderArticleCards();
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

  Object.assign(App, { renderTranslationStudio: renderTranslationStudio,
    renderCrossLanguageMatrix: renderCrossLanguageMatrix,
    approveUnchangedInActiveLanguage: approveUnchangedInActiveLanguage,
    approveAllUnchangedGlobal: approveAllUnchangedGlobal, renderLanguageTabs: renderLanguageTabs,
    updateLanguageProgress: updateLanguageProgress, updateAuxLangOptions: updateAuxLangOptions,
    initAuxLangSelector: initAuxLangSelector, highlightQuery: highlightQuery, initFilters: initFilters,
    jumpToPending: jumpToPending });
}(window.App = window.App || {}));
