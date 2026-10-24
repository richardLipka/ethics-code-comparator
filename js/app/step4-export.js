/**
 * Step 4: exports – workbooks, printable document, change summary and their settings.
 * Part of the application controller; the parts share the `App` namespace and are loaded in the
 * order given in index.html (js/app/core.js first, js/app/main.js last).
 */
(function (App) {
  'use strict';

  const { columnDisplayName, defaultNewColNameFor, elements, getMappingError, isDirectTranslationMode,
    languageLabel, persistState, showToast, state, t, exportProjectFile, ensureStubsInitialized } = App;

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
      optRef.textContent = isDirectTranslationMode()
        ? columnDisplayName(state.mappings.refNewCol)
        : `${columnDisplayName(state.mappings.refNewCol)} (${t('pdfOptionRefNew')})`;
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
   * Columns that accompany the selected language in the partial workbook, in two groups:
   * new versions first (the reference's new version and the current translation work of the
   * other languages), then the original columns of the sheet. Default: old and new reference.
   * Entries are column names, or { newOf: <old-version column> } for a language's new version.
   */
  function contextKey(entry) {
    return typeof entry === 'string' ? 'col:' + entry : 'new:' + entry.newOf;
  }

  function renderPartialContextOptions() {
    const container = elements.partialContextContainer;
    if (!container) return;
    const m = state.mappings;
    const langCol = state.exportLangCol;
    const isReference = langCol === m.refNewCol;
    if (!Array.isArray(state.partialContextCols)) {
      state.partialContextCols = [m.refOldCol, m.refNewCol].filter(Boolean);
    }

    const languageNewCols = m.targetCols.map(function (c) { return m.languages[c] && m.languages[c].newCol; }).filter(Boolean);
    const newVersions = [];
    if (!isReference && m.refNewCol) {
      newVersions.push({ entry: m.refNewCol, label: columnDisplayName(m.refNewCol) });
    }
    m.targetCols.forEach(function (c) {
      if (c === langCol) return;
      const cfg = m.languages[c] || {};
      const colName = cfg.newCol || cfg.newColName || defaultNewColNameFor(c);
      newVersions.push({ entry: { newOf: c }, label: columnDisplayName(colName) + ' (' + languageLabel(c) + ')' });
    });

    const own = isReference ? [m.refOldCol, m.refNewCol] : [langCol, m.languages[langCol] && m.languages[langCol].newCol];
    const originals = state.headers
      .filter(function (h) {
        return own.indexOf(h) === -1 && h !== m.refNewCol && languageNewCols.indexOf(h) === -1;
      })
      .map(function (h) { return { entry: h, label: columnDisplayName(h) }; });

    const selectedKeys = state.partialContextCols.map(contextKey);
    const displayOrder = newVersions.concat(originals).map(function (o) { return contextKey(o.entry); });
    container.innerHTML = '';
    [[t('partialGroupNew'), newVersions], [t('partialGroupOriginal'), originals]].forEach(function (group) {
      if (group[1].length === 0) return;
      const box = document.createElement('div');
      box.className = 'partial-context-group';
      const title = document.createElement('div');
      title.className = 'partial-context-title';
      title.textContent = group[0];
      const items = document.createElement('div');
      items.className = 'partial-context-items';
      group[1].forEach(function (option) {
        const key = contextKey(option.entry);
        const label = document.createElement('label');
        label.className = 'checkbox-label';
        const cb = document.createElement('input');
        cb.type = 'checkbox';
        cb.value = key;
        cb.checked = selectedKeys.indexOf(key) !== -1;
        cb.addEventListener('change', function () {
          const list = state.partialContextCols;
          const pos = list.map(contextKey).indexOf(key);
          if (cb.checked && pos === -1) list.push(option.entry);
          if (!cb.checked && pos !== -1) list.splice(pos, 1);
          // Exported in the order shown (new versions first); entries not offered for the current
          // language (e.g. its own columns) keep their place at the end
          const rank = function (entry) {
            const i = displayOrder.indexOf(contextKey(entry));
            return i === -1 ? displayOrder.length : i;
          };
          list.sort(function (x, y) { return rank(x) - rank(y); });
          persistState(300);
        });
        label.appendChild(cb);
        label.appendChild(document.createTextNode(option.label));
        items.appendChild(label);
      });
      box.appendChild(title);
      box.appendChild(items);
      container.appendChild(box);
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

    const summary = ComparatorEngine.buildChangeSummary({
      rows: state.rows,
      mappings: state.mappings,
      translations: state.translations,
      rowNumbers: state.rowNumbers,
      langCol: col,
      documentName: documentBaseName(),
      uiLang: I18n.getLanguage()
    });
    const adapterId = { markdown: 'summary-markdown', word: 'summary-word', pdf: 'summary-print' }[format] || format;
    const baseName = documentBaseName() + '_zmeny_' + summary.language;
    Exporters.run(adapterId, summary, { baseName: baseName, uiLang: I18n.getLanguage() });

    if (format === 'markdown') {
      showToast(t('mgrExportMdSuccess'), 'success');
    } else if (format === 'word') {
      showToast(t('mgrExportDocSuccess'), 'success');
    }
  }

  function exportDocument(adapterId, col) {
    const model = ComparatorEngine.buildDocumentModel({
      rows: state.rows,
      mappings: state.mappings,
      translations: state.translations,
      rowNumbers: state.rowNumbers,
      targetCol: col,
      documentName: documentBaseName(),
      uiLang: I18n.getLanguage(),
      template: state.pdfTemplate
    });
    return Exporters.run(adapterId, model, { baseName: documentBaseName() + '_' + model.language, uiLang: I18n.getLanguage() });
  }

  function initExports() {
    elements.downloadComprehensiveBtn.addEventListener('click', function () {
      exportTables('comprehensive');
    });

    elements.downloadCleanBtn.addEventListener('click', function () {
      exportTables('clean');
    });

    elements.downloadPartialBtn.addEventListener('click', function () {
      const col = selectedExportLang();
      if (col) exportTables('language', col);
    });

    elements.exportProjectJsonBtn.addEventListener('click', function () {
      exportProjectFile();
    });

    elements.tableFormatSelect.addEventListener('change', function () {
      state.tableExportFormat = elements.tableFormatSelect.value;
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
      if (col) exportDocument('html-print', col);
    });

    elements.downloadPdfHtmlBtn.addEventListener('click', function () {
      const col = selectedExportLang();
      if (!col) return;
      exportDocument('html-file', col);
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
            exportDocument('html-print', col);
          }, i * 400);
        });
      }
    });
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

  /**
   * Table exports (full, new version only, one language) through the selected table exporter
   * (default: Excel, see js/io/excel-exporter.js).
   */
  function exportTables(exportType, langCol) {
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

    const result = ComparatorEngine.prepareExportTables({
      rows: state.rows,
      headers: state.headers,
      rowNumbers: state.rowNumbers,
      mappings: state.mappings,
      translations: state.translations,
      exportType: exportType,
      langCol: langCol,
      contextCols: state.partialContextCols || [state.mappings.refOldCol, state.mappings.refNewCol].filter(Boolean),
      skippedRows: state.skippedRows,
      uiLang: I18n.getLanguage()
    });

    const suffix = exportType === 'clean' ? 'nova_verze'
      : (exportType === 'language' ? (langCol === state.mappings.refNewCol ? columnDisplayName(langCol) : languageLabel(langCol)) : 'srovnani');
    const adapter = Exporters.get(state.tableExportFormat) || Exporters.getDefault('table');
    const outcome = Exporters.run(adapter.id, { tables: result.tables }, {
      baseName: documentBaseName() + '_' + suffix,
      uiLang: I18n.getLanguage()
    });

    if (outcome.fileName) showToast(t('toastWorkbookGenerated', { NAME: outcome.fileName }), 'success');
    if (result.unverifiedCount > 0) {
      const key = exportType === 'clean' ? 'toastExportUnverifiedClean'
        : (exportType === 'language' ? 'toastExportUnverifiedLanguage' : 'toastExportUnverified');
      showToast(t(key, { COUNT: result.unverifiedCount }), 'warning');
    }
    return outcome.fileName || null;
  }

  /**
   * Format selector for table exports; shown only when more than one table exporter is registered.
   */
  function renderTableFormatOptions() {
    const select = elements.tableFormatSelect;
    if (!select) return;
    const adapters = Exporters.list('table');
    if (!Exporters.get(state.tableExportFormat)) state.tableExportFormat = Exporters.getDefault('table').id;
    select.innerHTML = '';
    adapters.forEach(function (a) {
      const opt = document.createElement('option');
      opt.value = a.id;
      opt.textContent = a.label;
      select.appendChild(opt);
    });
    select.value = state.tableExportFormat;
    elements.tableFormatWrapper.style.display = adapters.length > 1 ? 'flex' : 'none';
  }

  Object.assign(App, { updatePdfLanguageDropdown: updatePdfLanguageDropdown,
    exportManagerialSummary: exportManagerialSummary, initExports: initExports,
    initPdfTemplateSettings: initPdfTemplateSettings, openPdfTemplateModal: openPdfTemplateModal,
    exportTables: exportTables, renderTableFormatOptions: renderTableFormatOptions });
}(window.App = window.App || {}));
