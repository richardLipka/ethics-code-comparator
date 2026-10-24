/**
 * Step 1: reading a file through the import adapters, sheet selection, data preview and skipped rows.
 * Part of the application controller; the parts share the `App` namespace and are loaded in the
 * order given in index.html (js/app/core.js first, js/app/main.js last).
 */
(function (App) {
  'use strict';

  const { columnDisplayName, elements, persistState, PREVIEW_INITIAL_ROWS, PREVIEW_MORE_ROWS, rowNumberOf,
    showToast, state, t, setStep, loadProjectFile, resetWorkspaceForNewTable } = App;

  /**
   * Handle File Upload & Sheet Reading
   */
  function initFileUpload() {
    elements.fileInput.setAttribute('accept', Importers.acceptAttribute());

    elements.dropzone.addEventListener('click', function () {
      // Formats of all import adapters registered so far (adapters may be added later)
      elements.fileInput.setAttribute('accept', Importers.acceptAttribute());
      elements.fileInput.click();
    });

    elements.fileInput.addEventListener('change', function (e) {
      if (e.target.files && e.target.files[0]) {
        processUploadedFile(e.target.files[0]);
      }
      // Allow selecting the same file again
      e.target.value = '';
    });

    // The upload area is a button for keyboard users too
    elements.dropzone.addEventListener('keydown', function (e) {
      if (e.target === elements.dropzone && (e.key === 'Enter' || e.key === ' ')) {
        e.preventDefault();
        elements.dropzone.click();
      }
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

    elements.previewShowMoreBtn.addEventListener('click', function () {
      state.previewLimit = (state.previewLimit || PREVIEW_INITIAL_ROWS) + PREVIEW_MORE_ROWS;
      renderDataPreview();
    });

    // Re-read the sheet when the user says the first row is (not) a header
    elements.hasHeaderRowCheckbox.addEventListener('change', function () {
      if (state.importSource && state.activeSheet) {
        loadSheet(state.activeSheet);
      }
    });

    // Sample data button
    elements.loadSampleBtn.addEventListener('click', function () {
      loadSampleDemo();
    });

    // Download physical sample excel
    elements.downloadSampleBtn.addEventListener('click', function () {
      downloadSampleExcelFile();
    });
  }

  /**
   * Imports a file with the matching import adapter (see js/io/importers.js).
   */
  function processUploadedFile(file) {
    if (!Importers.findForFile(file.name, file.type) && /\.json$/i.test(file.name)) {
      loadProjectFile(file);
      return Promise.resolve();
    }
    return Importers.importFile(file).then(function (result) {
      state.fileName = file.name;
      state.importSource = result.source;
      state.sheetNames = result.source.sheetNames.slice();

      // Setup sheet select dropdown
      elements.sheetSelect.innerHTML = '';
      state.sheetNames.forEach(function (sheetName) {
        const opt = document.createElement('option');
        opt.value = sheetName;
        opt.textContent = sheetName;
        elements.sheetSelect.appendChild(opt);
      });

      const isSingleDoc = result.adapter && (result.adapter.id === 'word' || result.adapter.id === 'text');
      elements.sheetOptionsWrapper.style.display = isSingleDoc ? 'none' : 'flex';
      elements.sheetSelectWrapper.style.display = state.sheetNames.length > 1 ? 'block' : 'none';
      if (loadSheet(state.sheetNames[0])) {
        setStep(1);
        showToast(t('toastFileLoaded', { NAME: file.name }), 'success');
      }
    }).catch(function (err) {
      console.error('Import failed:', err);
      if (err && err.code === 'UNSUPPORTED') {
        alert(t('alertUnsupportedFile', { NAME: file.name, FORMATS: Importers.acceptAttribute().split(',').join(', ') }));
      } else if (err && err.code === 'NO_SHEETS') {
        alert(t('alertNoSheets'));
      } else {
        alert(t('alertFileReadError', { MSG: err && err.message ? err.message : file.name }));
      }
    });
  }

  function loadSheet(sheetName) {
    const hasHeader = elements.hasHeaderRowCheckbox.checked;
    const table = Importers.readSheet(state.importSource, sheetName, { hasHeaderRow: hasHeader });

    if (!table) {
      alert(t('alertEmptySheet'));
      return false;
    }

    state.activeSheet = sheetName;
    state.hasHeaderRow = hasHeader;
    elements.sheetSelect.value = sheetName;
    state.headers = table.headers;
    state.columnLetters = table.columnLetters;
    state.translations = {};
    state.rows = [];
    state.skippedRows = [];
    state.previewLimit = PREVIEW_INITIAL_ROWS;
    applyRowSelection(table.rows.map(function (row, idx) {
      const empty = isEmptyRow(row);
      return { rowNumber: table.rowNumbers[idx], row: row, skipped: empty, auto: empty };
    }));
    resetWorkspaceForNewTable();
    App.selectDefaultLanguageColumns();
    if (state.skippedRows.length > 0) {
      showToast(t('toastEmptyRowsSkipped', { COUNT: state.skippedRows.length }), 'info');
    }

    App.renderColumnMappingControls();
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
    state.fileName = 'ukazkovy_dokument.xlsx';
    state.importSource = null;
    state.hasHeaderRow = true;
    state.rows = SampleGenerator.getSampleData();
    state.headers = SampleGenerator.getSampleHeaders();
    state.rowNumbers = state.rows.map(function (r, idx) { return idx + 2; });
    state.skippedRows = [];
    state.previewLimit = PREVIEW_INITIAL_ROWS;
    state.columnLetters = {};
    state.sheetNames = ['Data'];
    state.activeSheet = 'Data';
    resetWorkspaceForNewTable();

    elements.sheetOptionsWrapper.style.display = 'none';

    const sampleDefaults = SampleGenerator.getSampleDefaultMappings();
    state.mappings.refOldCol = sampleDefaults.refOldCol;
    state.mappings.refNewCol = sampleDefaults.refNewCol;
    sampleDefaults.targetCols.forEach(function (col) { App.setLanguageUsed(col, true); });

    App.renderColumnMappingControls();
    renderDataPreview();
    revealMapping();
    setStep(1);

    showToast(t('toastSampleLoaded'), 'success');
  }

  /**
   * Step 1: Data preview. Every row can be marked as skipped (headings, notes, …); rows that
   * are completely empty are skipped automatically. Skipped rows are not processed, but the full
   * workbook keeps them unchanged at their original place.
   */
  function isEmptyRow(row) {
    return Object.keys(row).every(function (k) {
      const v = row[k];
      return v === '' || v === null || v === undefined || (typeof v === 'string' && !v.trim());
    });
  }

  // All rows of the sheet in their original order: [{ rowNumber, row, skipped, auto }]
  function allTableEntries() {
    const entries = state.rows.map(function (row, idx) {
      return { rowNumber: rowNumberOf(idx), row: row, skipped: false, auto: false };
    });
    (state.skippedRows || []).forEach(function (s) {
      entries.push({ rowNumber: s.rowNumber, row: s.row, skipped: true, auto: !!s.auto });
    });
    return entries.sort(function (a, b) { return a.rowNumber - b.rowNumber; });
  }

  /**
   * Splits the sheet into processed rows and skipped rows. Translation work is kept by row
   * number, so (un)skipping rows later never moves texts to other rows.
   */
  function applyRowSelection(entries) {
    const workByRow = {};
    Object.keys(state.translations || {}).forEach(function (col) {
      workByRow[col] = {};
      (state.translations[col] || []).forEach(function (stub, idx) {
        if (stub) workByRow[col][rowNumberOf(idx)] = stub;
      });
    });

    const active = entries.filter(function (e) { return !e.skipped; });
    state.rows = active.map(function (e) { return e.row; });
    state.rowNumbers = active.map(function (e) { return e.rowNumber; });
    state.skippedRows = entries.filter(function (e) { return e.skipped; }).map(function (e) {
      return { rowNumber: e.rowNumber, row: e.row, auto: !!e.auto };
    });

    Object.keys(workByRow).forEach(function (col) {
      state.translations[col] = state.rowNumbers.map(function (n) { return workByRow[col][n] || null; });
    });
  }

  function setRowSkipped(rowNumber, skipped) {
    const entries = allTableEntries();
    const entry = entries.find(function (e) { return e.rowNumber === rowNumber; });
    if (!entry || entry.auto) return;
    entry.skipped = skipped;
    applyRowSelection(entries);
    renderDataPreview();
    persistState(300);
  }

  function renderDataPreview() {
    if (!elements.dataPreviewContainer) return;
    const entries = allTableEntries();
    if (entries.length === 0 || !state.headers || state.headers.length === 0) {
      elements.dataPreviewContainer.style.display = 'none';
      return;
    }

    elements.dataPreviewContainer.style.display = 'block';
    const skippedCount = entries.length - state.rows.length;
    if (elements.previewRowsCount) {
      elements.previewRowsCount.textContent = `(${state.rows.length} ${t('rowsLabel')}` +
        (skippedCount > 0 ? ` · ${t('previewSkippedCount', { COUNT: skippedCount })}` : '') + ')';
    }

    const esc = ComparatorEngine.escapeHtml;
    elements.previewTableHead.innerHTML = '<tr>' +
      `<th title="${esc(t('previewSkipTitle'))}">${esc(t('previewSkipCol'))}</th>` +
      `<th>${esc(t('rowLabel'))}</th>` +
      state.headers.map(function (h) {
        return `<th>${esc(columnDisplayName(h))}</th>`;
      }).join('') + '</tr>';

    // Whitespace is collapsed so long single-line cells stay readable
    const limit = state.previewLimit || PREVIEW_INITIAL_ROWS;
    elements.previewTableBody.innerHTML = entries.slice(0, limit).map(function (e) {
      const cells = state.headers.map(function (h) {
        let val = e.row[h] !== undefined ? String(e.row[h]).replace(/\s+/g, ' ').trim() : '';
        const full = val;
        if (val.length > 60) val = val.substring(0, 57) + '...';
        return `<td title="${esc(full.slice(0, 500))}">${esc(val)}</td>`;
      }).join('');
      const skipTitle = e.auto ? t('previewEmptyRow') : t('previewSkipTitle');
      return `<tr class="preview-row${e.skipped ? ' preview-row-skipped' : ''}" data-row="${e.rowNumber}">` +
        `<td class="cell-center"><input type="checkbox" class="preview-skip-cb" data-row="${e.rowNumber}"` +
        `${e.skipped ? ' checked' : ''}${e.auto ? ' disabled' : ''} title="${esc(skipTitle)}" aria-label="${esc(t('previewSkipCol'))} ${e.rowNumber}"></td>` +
        `<td class="mono">${e.rowNumber}${e.auto ? ` <span class="preview-empty-badge">${esc(t('previewEmptyRow'))}</span>` : ''}</td>` +
        cells + '</tr>';
    }).join('');

    elements.previewTableBody.querySelectorAll('.preview-skip-cb').forEach(function (cb) {
      cb.addEventListener('change', function () {
        setRowSkipped(parseInt(cb.getAttribute('data-row'), 10), cb.checked);
      });
    });

    const remaining = entries.length - limit;
    elements.previewShowMoreBtn.style.display = remaining > 0 ? 'inline-flex' : 'none';
    elements.previewShowMoreBtn.textContent = t('previewShowMore', { COUNT: Math.min(remaining, PREVIEW_MORE_ROWS), TOTAL: remaining });
  }

  function downloadSampleExcelFile() {
    const exporter = Exporters.getDefault('table');
    Exporters.run(exporter.id, {
      tables: [{ name: 'Data', headers: SampleGenerator.getSampleHeaders(), rows: SampleGenerator.getSampleData() }]
    }, { baseName: 'ukazkovy_dokument', uiLang: I18n.getLanguage() });
    showToast(t('toastSampleDownloaded'), 'success');
  }

  Object.assign(App, { initFileUpload: initFileUpload, processUploadedFile: processUploadedFile,
    renderDataPreview: renderDataPreview });
}(window.App = window.App || {}));
