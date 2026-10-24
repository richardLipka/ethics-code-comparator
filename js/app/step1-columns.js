/**
 * Step 1: column selection – reference columns, languages, their new-version columns and new languages.
 * Part of the application controller; the parts share the `App` namespace and are loaded in the
 * order given in index.html (js/app/core.js first, js/app/main.js last).
 */
(function (App) {
  'use strict';

  const { columnDisplayName, defaultNewColNameFor, elements, isDirectTranslationMode, languageLabel,
    persistState, showToast, state, t } = App;

  /**
   * Column selection: two reference columns, then for every other column whether it holds the
   * old version of a language (used) or is ignored, the language name and which column holds or
   * will hold its new version.
   */

  // Columns with longer text in a reasonable share of rows are pre-selected as languages;
  // the actual languages are never guessed from the header names.
  function selectDefaultLanguageColumns() {
    if (state.headers.length === 1) {
      state.mappings.refNewCol = state.headers[0];
      state.mappings.refOldCol = '';
      setLanguageUsed(state.headers[0], false);
      return;
    }
    state.headers.forEach(function (h) {
      let filled = 0;
      let totalWords = 0;
      state.rows.forEach(function (row) {
        const v = row[h];
        if (typeof v === 'string' && v) {
          filled++;
          totalWords += v.split(/\s+/).filter(Boolean).length;
        }
      });
      const isText = filled >= Math.max(1, state.rows.length * 0.5) && totalWords / Math.max(filled, 1) >= 3;
      if (isText) setLanguageUsed(h, true);
    });
  }

  function ensureLanguageConfig(col) {
    if (!state.mappings.languages) state.mappings.languages = {};
    if (!state.mappings.languages[col]) {
      const isNew = ComparatorEngine.isNewLanguage(state.mappings, col);
      state.mappings.languages[col] = { label: '', newCol: '', newColName: '', isNewLanguage: isNew };
    }
    return state.mappings.languages[col];
  }

  function newColumnOwner(col) {
    return state.mappings.targetCols.find(function (c) {
      return state.mappings.languages[c] && state.mappings.languages[c].newCol === col;
    }) || '';
  }

  function setLanguageUsed(col, used) {
    const targets = state.mappings.targetCols;
    if (used) {
      ensureLanguageConfig(col);
      if (targets.indexOf(col) === -1) targets.push(col);
      // Keep original sheet columns in header order; put newly added languages at the end
      targets.sort(function (a, b) {
        const idxA = state.headers.indexOf(a);
        const idxB = state.headers.indexOf(b);
        if (idxA !== -1 && idxB !== -1) return idxA - idxB;
        if (idxA !== -1) return -1;
        if (idxB !== -1) return 1;
        return 0;
      });
    } else {
      const pos = targets.indexOf(col);
      if (pos !== -1) targets.splice(pos, 1);
    }
  }

  function addNewLanguage(name, colName) {
    const trimmedName = (name || '').trim();
    if (!trimmedName) {
      showToast(t('toastEnterLangName'), 'warning');
      if (elements.newLangNameInput) elements.newLangNameInput.focus();
      return null;
    }

    const existing = state.mappings.targetCols.some(function (c) {
      const info = ComparatorEngine.getLanguageInfo(state.mappings, c);
      return (info.label && info.label.toLowerCase() === trimmedName.toLowerCase()) ||
             (info.newColName && info.newColName.toLowerCase() === trimmedName.toLowerCase()) ||
             c.toLowerCase() === trimmedName.toLowerCase();
    });
    if (existing) {
      showToast(t('toastLangAlreadyExists'), 'warning');
      return null;
    }

    const id = '__new__' + trimmedName.replace(/[^a-zA-Z0-9_\u00C0-\u024F\u1E00-\u1EFF]/g, '_') + '_' + Date.now();
    const exportColName = (colName || '').trim() || trimmedName;

    if (!state.mappings.languages) state.mappings.languages = {};
    state.mappings.languages[id] = {
      label: trimmedName,
      newCol: '',
      newColName: exportColName,
      isNewLanguage: true
    };

    setLanguageUsed(id, true);
    renderLanguageTable();
    if (isDirectTranslationMode()) {
      renderQuickLangs();
      renderSelectedDirectLangs();
    }
    persistState(300);
    showToast(t('toastNewLangAdded', { NAME: trimmedName }), 'success');
    if (elements.newLangNameInput) elements.newLangNameInput.value = '';
    if (elements.newLangColNameInput) elements.newLangColNameInput.value = '';
    return id;
  }

  function removeNewLanguage(col) {
    const info = ComparatorEngine.getLanguageInfo(state.mappings, col);
    const name = info.label || col;
    setLanguageUsed(col, false);
    delete state.mappings.languages[col];
    delete state.translations[col];
    if (state.activeTargetCol === col) {
      state.activeTargetCol = state.mappings.targetCols[0] || '';
    }
    if (state.exportLangCol === col) {
      state.exportLangCol = state.mappings.targetCols[0] || '';
    }
    renderLanguageTable();
    if (isDirectTranslationMode()) {
      renderQuickLangs();
      renderSelectedDirectLangs();
    }
    persistState(300);
    showToast(t('toastNewLangRemoved', { NAME: name }), 'info');
  }

  const QUICK_LANGS = [
    { code: 'CS', label: 'Čeština', flag: '🇨🇿' },
    { code: 'SK', label: 'Slovenčina', flag: '🇸🇰' },
    { code: 'EN', label: 'English', flag: '🇬🇧' },
    { code: 'DE', label: 'DE', flag: '🇩🇪' },
    { code: 'PL', label: 'Polski', flag: '🇵🇱' },
    { code: 'FR', label: 'Français', flag: '🇫🇷' },
    { code: 'ES', label: 'Español', flag: '🇪🇸' },
    { code: 'IT', label: 'Italiano', flag: '🇮🇹' },
    { code: 'UK', label: 'Українська', flag: '🇺🇦' }
  ];

  function isLanguageSelected(labelOrCode) {
    const needle = (labelOrCode || '').toLowerCase();
    return state.mappings.targetCols.some(function (c) {
      const info = ComparatorEngine.getLanguageInfo(state.mappings, c);
      const label = (info.label || '').toLowerCase();
      const newColName = (info.newColName || '').toLowerCase();
      const colId = String(c).toLowerCase();
      return label === needle || newColName === needle || colId === needle;
    });
  }

  function findColForLanguage(labelOrCode) {
    const needle = (labelOrCode || '').toLowerCase();
    return state.mappings.targetCols.find(function (c) {
      const info = ComparatorEngine.getLanguageInfo(state.mappings, c);
      const label = (info.label || '').toLowerCase();
      const newColName = (info.newColName || '').toLowerCase();
      const colId = String(c).toLowerCase();
      return label === needle || newColName === needle || colId === needle;
    }) || '';
  }

  function renderQuickLangs() {
    if (!elements.quickLangsContainer) return;
    elements.quickLangsContainer.innerHTML = '';
    QUICK_LANGS.forEach(function (item) {
      const selected = isLanguageSelected(item.label) || isLanguageSelected(item.code);
      const btn = document.createElement('button');
      btn.type = 'button';
      btn.className = 'quick-lang-btn' + (selected ? ' active' : '');
      btn.innerHTML = `<span class="quick-lang-flag">${item.flag}</span> <span class="quick-lang-name">${ComparatorEngine.escapeHtml(item.label)}</span>` + (selected ? ' ✓' : '');
      btn.addEventListener('click', function () {
        if (selected) {
          const col = findColForLanguage(item.label) || findColForLanguage(item.code);
          if (col) removeNewLanguage(col);
        } else {
          addNewLanguage(item.label, item.label);
        }
      });
      elements.quickLangsContainer.appendChild(btn);
    });
  }

  function renderSelectedDirectLangs() {
    if (!elements.selectedDirectLangsList) return;
    const esc = ComparatorEngine.escapeHtml;
    const cols = state.mappings.targetCols || [];
    if (cols.length === 0) {
      elements.selectedDirectLangsList.innerHTML = `<span style="color:var(--text-muted); font-size:0.85rem; font-style:italic;">${esc(t('noLangsSelectedYet'))}</span>`;
      return;
    }
    elements.selectedDirectLangsList.innerHTML = '';
    cols.forEach(function (col) {
      const info = ComparatorEngine.getLanguageInfo(state.mappings, col);
      const label = info.label || col;
      const chip = document.createElement('div');
      chip.className = 'selected-lang-chip';
      chip.innerHTML = `<span>🌐 ${esc(label)}</span><button type="button" class="remove-direct-lang-btn" title="${esc(t('removeLangTitle'))}">✕</button>`;
      chip.querySelector('.remove-direct-lang-btn').addEventListener('click', function () {
        removeNewLanguage(col);
      });
      elements.selectedDirectLangsList.appendChild(chip);
    });
  }

  /**
   * Store final labels and new-column names, so exports and the project file are explicit.
   */
  function finalizeLanguageConfig() {
    state.mappings.targetCols.forEach(function (col) {
      const cfg = ensureLanguageConfig(col);
      if (!cfg.label) cfg.label = languageLabel(col);
      if (cfg.newCol && state.headers.indexOf(cfg.newCol) === -1) cfg.newCol = '';
      if (!cfg.newColName) cfg.newColName = defaultNewColNameFor(col);
    });
  }

  /**
   * When the reference columns differ by a version number (e.g. "EN 2024" / "EN 2026") and the
   * sheet already has the matching column for a language ("CS 2026"), use it as its new version.
   */
  function autoPairNewColumns() {
    state.mappings.targetCols.slice().forEach(function (col) {
      const cfg = ensureLanguageConfig(col);
      if (cfg.newCol) return;
      const candidate = ComparatorEngine.defaultNewColName(col, state.mappings.refOldCol, state.mappings.refNewCol, '\u0000');
      if (candidate.indexOf('\u0000') !== -1 || candidate === col || state.headers.indexOf(candidate) === -1) return;
      if (candidate === state.mappings.refOldCol || candidate === state.mappings.refNewCol || newColumnOwner(candidate)) return;
      setLanguageUsed(candidate, false);
      cfg.newCol = candidate;
    });
  }

  function onReferenceChanged() {
    const m = state.mappings;
    [m.refOldCol, m.refNewCol].forEach(function (ref) {
      if (!ref) return;
      setLanguageUsed(ref, false);
      m.targetCols.forEach(function (c) {
        if (m.languages[c] && m.languages[c].newCol === ref) m.languages[c].newCol = '';
      });
    });
    autoPairNewColumns();
    renderLanguageTable();
  }

  function fillColumnSelect(select, value, allowNone) {
    select.innerHTML = '';
    const placeholder = document.createElement('option');
    placeholder.value = '';
    placeholder.textContent = allowNone ? t('optNoOldReference') : t('selectColumnPlaceholder');
    select.appendChild(placeholder);
    state.headers.forEach(function (h) {
      const opt = document.createElement('option');
      opt.value = h;
      opt.textContent = columnDisplayName(h);
      select.appendChild(opt);
    });
    select.value = value || '';
  }

  function renderColumnMappingControls() {
    elements.mappingContainer.style.display = 'block';
    const isDirect = isDirectTranslationMode();

    if (isDirect) {
      if (elements.refOldColCard) elements.refOldColCard.style.display = 'none';
      if (elements.directTransInfoBox) elements.directTransInfoBox.style.display = 'block';
      if (elements.refNewColLabel) elements.refNewColLabel.textContent = t('sourceTextLabel');
      if (elements.refNewColHelp) elements.refNewColHelp.textContent = t('sourceTextCardHelp');
      if (elements.targetColsTitle) elements.targetColsTitle.textContent = t('targetLangsDirectTitle');
      if (elements.targetColsHelp) elements.targetColsHelp.textContent = t('targetLangsDirectHelp');
      if (elements.quickLangsWrapper) elements.quickLangsWrapper.style.display = 'block';
      if (elements.languagesTableWrapper) elements.languagesTableWrapper.style.display = 'none';
      if (elements.confirmMappingBtn) elements.confirmMappingBtn.textContent = t('goToTranslationDirect');

      if (elements.stepNav2) elements.stepNav2.style.display = 'none';
      if (elements.stepDivider2) elements.stepDivider2.style.display = 'none';
      if (elements.stepNumber3) elements.stepNumber3.textContent = '2';
      if (elements.stepNumber4) elements.stepNumber4.textContent = '3';

      renderQuickLangs();
      renderSelectedDirectLangs();
    } else {
      if (elements.refOldColCard) elements.refOldColCard.style.display = '';
      if (elements.directTransInfoBox) elements.directTransInfoBox.style.display = 'none';
      if (elements.refNewColLabel) elements.refNewColLabel.textContent = t('refNewLabel');
      if (elements.refNewColHelp) elements.refNewColHelp.textContent = t('refNewHelp');
      if (elements.targetColsTitle) elements.targetColsTitle.textContent = t('targetColsLabel');
      if (elements.targetColsHelp) elements.targetColsHelp.textContent = t('targetColsHelp');
      if (elements.quickLangsWrapper) elements.quickLangsWrapper.style.display = 'none';
      if (elements.languagesTableWrapper) elements.languagesTableWrapper.style.display = '';
      if (elements.confirmMappingBtn) elements.confirmMappingBtn.textContent = t('confirmMappingBtn');

      if (elements.stepNav2) elements.stepNav2.style.display = '';
      if (elements.stepDivider2) elements.stepDivider2.style.display = '';
      if (elements.stepNumber3) elements.stepNumber3.textContent = '3';
      if (elements.stepNumber4) elements.stepNumber4.textContent = '4';
    }

    fillColumnSelect(elements.refOldSelect, state.mappings.refOldCol, true);
    fillColumnSelect(elements.refNewSelect, state.mappings.refNewCol, false);

    elements.refOldSelect.onchange = function () {
      state.mappings.refOldCol = elements.refOldSelect.value;
      onReferenceChanged();
    };

    elements.refNewSelect.onchange = function () {
      state.mappings.refNewCol = elements.refNewSelect.value;
      onReferenceChanged();
    };

    renderLanguageTable();
  }

  function firstSampleText(col) {
    const row = state.rows.find(function (r) { return r[col] !== '' && r[col] !== undefined && r[col] !== null; });
    return row ? String(row[col]).replace(/\s+/g, ' ').slice(0, 200) : '';
  }

  function renderLanguageTable() {
    const tbody = elements.languagesTableBody;
    const m = state.mappings;
    const esc = ComparatorEngine.escapeHtml;
    tbody.innerHTML = '';

    state.headers.forEach(function (h) {
      if (h === m.refOldCol || h === m.refNewCol) return;
      const used = m.targetCols.indexOf(h) !== -1;
      const owner = used ? '' : newColumnOwner(h);
      const tr = document.createElement('tr');
      tr.className = used ? 'lang-row-used' : (owner ? 'lang-row-paired' : 'lang-row-ignored');
      tr.setAttribute('data-col', h);

      let html = `<td><input type="checkbox" class="lang-use-cb" ${used ? 'checked' : ''} ${owner ? 'disabled' : ''}></td>`;
      html += `<td class="lang-col-name" title="${esc(firstSampleText(h))}">${esc(columnDisplayName(h))}</td>`;

      if (used) {
        const cfg = ensureLanguageConfig(h);
        const candidates = state.headers.filter(function (c) {
          return c !== h && c !== m.refOldCol && c !== m.refNewCol && m.targetCols.indexOf(c) === -1 &&
            (!newColumnOwner(c) || newColumnOwner(c) === h);
        });
        html += `<td><input type="text" class="form-input lang-label-input" value="${esc(languageLabel(h))}"></td>`;
        html += '<td><div class="new-col-cell"><select class="form-select lang-newcol-select">' +
          `<option value="">${esc(t('optCreateNewCol'))}</option>` +
          candidates.map(function (c) {
            return `<option value="${esc(c)}" ${cfg.newCol === c ? 'selected' : ''}>${esc(columnDisplayName(c))}</option>`;
          }).join('') + '</select>' +
          (cfg.newCol ? '' : `<input type="text" class="form-input lang-newname-input" value="${esc(cfg.newColName || '')}" placeholder="${esc(defaultNewColNameFor(h))}" title="${esc(t('newColNamePlaceholder'))}">`) +
          '</div></td>';
      } else if (owner) {
        html += `<td colspan="2">${esc(t('usedAsNewVersionOf', { LANG: languageLabel(owner) }))}</td>`;
      } else {
        html += `<td colspan="2">${esc(t('ignoredColumn'))}</td>`;
      }
      tr.innerHTML = html;
      tbody.appendChild(tr);

      const cb = tr.querySelector('.lang-use-cb');
      cb.addEventListener('change', function () {
        setLanguageUsed(h, cb.checked);
        renderLanguageTable();
      });
      const labelInput = tr.querySelector('.lang-label-input');
      if (labelInput) {
        labelInput.addEventListener('input', function () {
          ensureLanguageConfig(h).label = labelInput.value.trim();
        });
      }
      const newColSelect = tr.querySelector('.lang-newcol-select');
      if (newColSelect) {
        newColSelect.addEventListener('change', function () {
          ensureLanguageConfig(h).newCol = newColSelect.value;
          renderLanguageTable();
        });
      }
      const newNameInput = tr.querySelector('.lang-newname-input');
      if (newNameInput) {
        newNameInput.addEventListener('input', function () {
          ensureLanguageConfig(h).newColName = newNameInput.value.trim();
        });
      }
    });

    // Render newly added languages (not in original sheet headers)
    state.mappings.targetCols.filter(function (c) {
      return ComparatorEngine.isNewLanguage(state.mappings, c);
    }).forEach(function (newColId) {
      const cfg = ensureLanguageConfig(newColId);
      const tr = document.createElement('tr');
      tr.className = 'lang-row-used lang-row-new';
      tr.setAttribute('data-col', newColId);

      let html = '<td><input type="checkbox" class="lang-use-cb" checked></td>';
      html += `<td class="lang-col-name">${esc(languageLabel(newColId))} <span class="badge badge-info" style="font-size:0.7rem; padding:0.15rem 0.4rem; vertical-align:middle;">${esc(t('badgeNewLanguage'))}</span><small style="display:block; color:var(--text-muted); font-size:0.75rem;">${esc(t('noOldVersion'))}</small></td>`;
      html += `<td><input type="text" class="form-input lang-label-input" value="${esc(languageLabel(newColId))}"></td>`;
      html += `<td><div class="new-col-cell"><input type="text" class="form-input lang-newname-input" value="${esc(cfg.newColName || '')}" placeholder="${esc(languageLabel(newColId))}" title="${esc(t('newColNamePlaceholder'))}"><button type="button" class="btn btn-sm btn-outline remove-lang-btn" data-col="${esc(newColId)}" title="${esc(t('removeLangTitle'))}" style="color:#ef4444; border-color:var(--border-color); padding:0.3rem 0.6rem;">✕</button></div></td>`;

      tr.innerHTML = html;
      tbody.appendChild(tr);

      const cb = tr.querySelector('.lang-use-cb');
      cb.addEventListener('change', function () {
        if (!cb.checked) {
          removeNewLanguage(newColId);
        }
      });
      const labelInput = tr.querySelector('.lang-label-input');
      if (labelInput) {
        labelInput.addEventListener('input', function () {
          ensureLanguageConfig(newColId).label = labelInput.value.trim();
        });
      }
      const newNameInput = tr.querySelector('.lang-newname-input');
      if (newNameInput) {
        newNameInput.addEventListener('input', function () {
          ensureLanguageConfig(newColId).newColName = newNameInput.value.trim();
        });
      }
      const removeBtn = tr.querySelector('.remove-lang-btn');
      if (removeBtn) {
        removeBtn.addEventListener('click', function () {
          removeNewLanguage(newColId);
        });
      }
    });
  }

  /**
   * Initialize translation stubs for target languages
   */
  function ensureStubsInitialized() {
    finalizeLanguageConfig();
    state.mappings.targetCols.forEach(function (targetCol) {
      // Rebuild every stub from the current rows/mapping and merge the saved work into it, so
      // reference texts, the old target text and classifications are always up to date
      // (also after changing the reference columns or loading a compact project file).
      // Text already present in the language's new-version column is used as the starting point.
      const isNew = ComparatorEngine.isNewLanguage(state.mappings, targetCol);
      state.translations[targetCol] = ComparatorEngine.rehydrateStubs(
        state.rows,
        state.translations[targetCol],
        targetCol,
        state.mappings.refOldCol,
        state.mappings.refNewCol,
        state.mappings.languages[targetCol].newCol,
        isNew
      );
    });

    if (!state.activeTargetCol || !state.mappings.targetCols.includes(state.activeTargetCol)) {
      state.activeTargetCol = state.mappings.targetCols[0] || '';
    }

    persistState(400);
  }

  Object.assign(App, { selectDefaultLanguageColumns: selectDefaultLanguageColumns,
    setLanguageUsed: setLanguageUsed, addNewLanguage: addNewLanguage, removeNewLanguage: removeNewLanguage,
    finalizeLanguageConfig: finalizeLanguageConfig, renderColumnMappingControls: renderColumnMappingControls,
    ensureStubsInitialized: ensureStubsInitialized });
}(window.App = window.App || {}));
