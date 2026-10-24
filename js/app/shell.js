/**
 * Page shell: interface language, theme, layout toggle, step navigation, sticky header and modal dismissal.
 * Part of the application controller; the parts share the `App` namespace and are loaded in the
 * order given in index.html (js/app/core.js first, js/app/main.js last).
 */
(function (App) {
  'use strict';

  const { elements, getMappingError, isDirectTranslationMode, readPref, showToast, state, t, writePref } = App;

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
        if (state.rows.length > 0) App.renderColumnMappingControls();
        App.renderDataPreview();
      } else if (state.currentStep === 2) {
        App.renderReferenceOverview();
      } else if (state.currentStep === 3) {
        App.renderTranslationStudio();
      } else if (state.currentStep === 4) {
        App.updatePdfLanguageDropdown();
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
    App.updateAuxLangOptions();
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
        App.renderArticleCards();
      });
    }
  }

  /**
   * Stepper navigation: shows a step after checking that the data and column selection allow it.
   */
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
      App.renderReferenceOverview();
    } else if (stepNumber === 3) {
      App.ensureStubsInitialized();
      App.renderTranslationStudio();
    } else if (stepNumber === 4) {
      App.ensureStubsInitialized();
      App.updatePdfLanguageDropdown();
      App.renderTableFormatOptions();
    }

    // Staying in the same step keeps the position (e.g. the column selection revealed after loading a file)
    if (stepChanged) window.scrollTo({ top: 0, behavior: 'smooth' });
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

  Object.assign(App, { initLanguage: initLanguage, initTheme: initTheme,
    updateLayoutToggleText: updateLayoutToggleText, initLayoutMode: initLayoutMode, setStep: setStep,
    initStickyHeader: initStickyHeader, initModalDismissal: initModalDismissal });
}(window.App = window.App || {}));
