/**
 * ComparatorEngine - Core logic for text diffing, row classification,
 * translation stub assembly, prompt generation, exports and change summaries.
 * Universal module: works in browser and Node.js.
 */
(function (root, factory) {
  if (typeof define === 'function' && define.amd) {
    define(['diff'], factory);
  } else if (typeof module === 'object' && module.exports) {
    const Diff = require('diff');
    module.exports = factory(Diff);
  } else {
    root.ComparatorEngine = factory(root.Diff);
  }
}(typeof self !== 'undefined' ? self : this, function (Diff) {
  'use strict';

  /**
   * Escape HTML to prevent injection and render text safely.
   */
  function escapeHtml(str) {
    if (str === null || str === undefined) return '';
    return String(str)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#039;');
  }

  /**
   * Escape HTML and keep line breaks (for Word/print documents that do not honour pre-wrap).
   */
  function escapeHtmlMultiline(str) {
    return escapeHtml(str).replace(/\r?\n/g, '<br>');
  }

  /**
   * Markdown blockquote that keeps every line of a multi-line text inside the quote.
   */
  function mdQuote(text) {
    return String(text === null || text === undefined ? '' : text)
      .split(/\r\n|\r|\n/)
      .map(function (line) { return line ? '> ' + line : '>'; })
      .join('\n');
  }

  /**
   * Normalize text by trimming and normalizing internal whitespace.
   */
  function cleanText(text) {
    if (text === null || text === undefined) return '';
    return String(text).trim();
  }

  /**
   * Standardize whitespace for comparison (collapses multiple spaces/newlines).
   */
  function normalizeForComparison(text) {
    if (!text) return '';
    return String(text).replace(/\r\n/g, '\n').replace(/\r/g, '\n').trim();
  }

  /**
   * Classify a row based on old and new reference text.
   * Categories:
   *  - 'EMPTY': both cells are empty
   *  - 'ADDED': old is empty, new has text (new paragraph added in 2026)
   *  - 'REMOVED': old has text, new is empty (paragraph removed in 2026)
   *  - 'UNCHANGED': both have text and are identical
   *  - 'MODIFIED': both have text but differ
   */
  function classifyRow(refOld, refNew) {
    const oldClean = cleanText(refOld);
    const newClean = cleanText(refNew);

    const oldEmpty = oldClean.length === 0;
    const newEmpty = newClean.length === 0;

    if (oldEmpty && newEmpty) {
      return 'EMPTY';
    }
    if (oldEmpty && !newEmpty) {
      return 'ADDED';
    }
    if (!oldEmpty && newEmpty) {
      return 'REMOVED';
    }

    if (normalizeForComparison(oldClean) === normalizeForComparison(newClean)) {
      return 'UNCHANGED';
    }

    return 'MODIFIED';
  }

  /**
   * Compute word-level diff between two text strings using jsdiff.
   */
  function computeDiff(oldText, newText, options) {
    const opts = options || {};
    const textA = oldText !== null && oldText !== undefined ? String(oldText) : '';
    const textB = newText !== null && newText !== undefined ? String(newText) : '';

    if (!Diff || !Diff.diffWordsWithSpace) {
      throw new Error('Knihovna jsdiff (Diff) není načtena!');
    }

    const chunks = Diff.diffWordsWithSpace(textA, textB);

    let addedTokens = 0;
    let removedTokens = 0;
    let unchangedTokens = 0;
    let addedChars = 0;
    let removedChars = 0;

    const addedList = [];
    const removedList = [];

    chunks.forEach(function (part) {
      const trimmed = part.value.trim();
      const tokenLen = trimmed ? trimmed.split(/\s+/).length : 0;

      if (part.added) {
        addedTokens += tokenLen;
        addedChars += part.value.length;
        if (trimmed) addedList.push(trimmed);
      } else if (part.removed) {
        removedTokens += tokenLen;
        removedChars += part.value.length;
        if (trimmed) removedList.push(trimmed);
      } else {
        unchangedTokens += tokenLen;
      }
    });

    const hasChanges = addedTokens > 0 || removedTokens > 0 || addedChars > 0 || removedChars > 0;
    const totalTokens = addedTokens + removedTokens + unchangedTokens;
    const similarity = totalTokens > 0 ? (unchangedTokens / totalTokens) : 1;

    return {
      chunks: chunks,
      stats: {
        addedTokens: addedTokens,
        removedTokens: removedTokens,
        unchangedTokens: unchangedTokens,
        addedChars: addedChars,
        removedChars: removedChars,
        similarity: Math.round(similarity * 100) / 100,
        hasChanges: hasChanges
      },
      summary: {
        addedTerms: addedList,
        removedTerms: removedList
      }
    };
  }

  /**
   * Render diff chunks as HTML.
   * mode: 'inline' | 'side-by-side'
   */
  function renderDiffHtml(diffResult, mode) {
    if (!diffResult || !diffResult.chunks) return '';
    const viewMode = mode || 'inline';

    if (viewMode === 'side-by-side') {
      let leftHtml = '';
      let rightHtml = '';

      diffResult.chunks.forEach(function (part) {
        const val = escapeHtml(part.value);
        if (part.removed) {
          leftHtml += '<del class="diff-del" title="Odstraněno">' + val + '</del>';
        } else if (part.added) {
          rightHtml += '<ins class="diff-ins" title="Přidáno">' + val + '</ins>';
        } else {
          leftHtml += '<span class="diff-unchanged">' + val + '</span>';
          rightHtml += '<span class="diff-unchanged">' + val + '</span>';
        }
      });

      return {
        leftHtml: leftHtml,
        rightHtml: rightHtml
      };
    }

    // Default inline mode
    let inlineHtml = '';
    diffResult.chunks.forEach(function (part) {
      const val = escapeHtml(part.value);
      if (part.added) {
        inlineHtml += '<ins class="diff-ins" title="Přidáno v nové verzi">' + val + '</ins>';
      } else if (part.removed) {
        inlineHtml += '<del class="diff-del" title="Odstraněno ze staré verze">' + val + '</del>';
      } else {
        inlineHtml += '<span class="diff-unchanged">' + val + '</span>';
      }
    });

    return inlineHtml;
  }

  /**
   * Excel row number of a data row (rowNumbers[idx] when the sheet had gaps or no header row).
   */
  function rowNumberOf(rowNumbers, idx) {
    return rowNumbers && rowNumbers[idx] ? rowNumbers[idx] : idx + 2;
  }

  /**
   * Language-neutral paragraph tag used in prompts and Smart Paste, e.g. "#12".
   */
  function rowTag(rowNumbers, idx) {
    return '#' + rowNumberOf(rowNumbers, idx);
  }

  function firstNumberToken(text) {
    const m = String(text || '').match(/\d{2,}/);
    return m ? m[0] : '';
  }

  /**
   * Default name of a newly created "new version" column. When the two reference headers differ
   * by a number (e.g. "EN 2024" / "EN 2026"), the same substitution is applied to the old column
   * ("CS 2024" → "CS 2026"); otherwise a suffix is appended.
   */
  function defaultNewColName(oldCol, refOldCol, refNewCol, suffix) {
    const oldToken = firstNumberToken(refOldCol);
    const newToken = firstNumberToken(refNewCol);
    const col = String(oldCol);
    if (oldToken && newToken && oldToken !== newToken && col.indexOf(oldToken) !== -1) {
      return col.split(oldToken).join(newToken);
    }
    return col + (suffix || ' (new)');
  }

  /**
   * Default language label of a column: the header without the version number the reference
   * columns use ("CS 2024" → "CS"), or the header itself.
   */
  function defaultLanguageLabel(col, refOldCol) {
    const oldToken = firstNumberToken(refOldCol);
    const text = String(col);
    if (oldToken && text.indexOf(oldToken) !== -1) {
      const stripped = text.split(oldToken).join('').replace(/[\s()\-–_]+$/, '').replace(/^[\s()\-–_]+/, '').trim();
      if (stripped) return stripped;
    }
    return text;
  }

  /**
   * Settings of one language (keyed by its old-version column).
   * newCol: an existing column with the new version ('' = a new column is created on export).
   */
  function getLanguageInfo(mappings, col) {
    const languages = (mappings && mappings.languages) || {};
    const info = languages[col] || {};
    return {
      col: col,
      label: info.label || col,
      newCol: info.newCol || '',
      newColName: info.newColName || defaultNewColName(col, mappings && mappings.refOldCol, mappings && mappings.refNewCol)
    };
  }

  /**
   * Label of any column usable as an export language (the reference new column or a language).
   */
  function getLanguageLabel(mappings, col) {
    if (mappings && col === mappings.refNewCol) return String(col);
    return getLanguageInfo(mappings, col).label;
  }

  /**
   * Generate initial translation stub for a target language.
   *
   * @param {Array<Object>} rows - All table rows
   * @param {string} targetCol - Column with the old version of the language
   * @param {string} refOldCol - Reference column, old version
   * @param {string} refNewCol - Reference column, new version
   * @param {string} [newCol] - Existing column that already holds (part of) the new version
   * @returns {Array<Object>} Stubs for each row
   */
  function generateTargetStub(rows, targetCol, refOldCol, refNewCol, newCol) {
    if (!rows || !Array.isArray(rows)) return [];

    return rows.map(function (row, idx) {
      const refOld = row[refOldCol];
      const refNew = row[refNewCol];
      const targetOld = row[targetCol] !== undefined && row[targetCol] !== null ? String(row[targetCol]) : '';
      const existingNew = newCol && row[newCol] !== undefined && row[newCol] !== null ? String(row[newCol]) : '';
      const classification = classifyRow(refOld, refNew);

      let initialNewText = '';
      let status = 'pending';
      let isVerified = false;

      switch (classification) {
        case 'UNCHANGED':
          // Auto-copy the approved old translation
          initialNewText = targetOld;
          status = 'auto_copied';
          isVerified = true;
          break;

        case 'REMOVED':
          initialNewText = '';
          status = 'removed';
          isVerified = true;
          break;

        case 'ADDED':
          initialNewText = '';
          status = 'pending_translation';
          isVerified = false;
          break;

        case 'MODIFIED':
          initialNewText = '';
          status = 'pending_review';
          isVerified = false;
          break;

        case 'EMPTY':
        default:
          initialNewText = '';
          status = 'empty';
          isVerified = true;
          break;
      }

      // Text already present in the new-version column of the sheet wins, but has to be reviewed
      // unless it is simply the unchanged old translation
      if (existingNew.trim()) {
        const sameAsAuto = classification === 'UNCHANGED' && existingNew === targetOld;
        initialNewText = existingNew;
        if (!sameAsAuto) {
          status = 'imported';
          isVerified = false;
        }
      }

      return {
        rowIndex: idx,
        targetCol: targetCol,
        classification: classification,
        refOld: refOld !== undefined && refOld !== null ? String(refOld) : '',
        refNew: refNew !== undefined && refNew !== null ? String(refNew) : '',
        targetOld: targetOld,
        targetNew: initialNewText,
        status: status,
        isVerified: isVerified,
        notes: ''
      };
    });
  }

  function buildChangesSummary(diffResult, lang, limit) {
    if (!diffResult || !diffResult.summary) return '';
    const added = diffResult.summary.addedTerms.filter(function (t) { return t.length > 1; });
    const removed = diffResult.summary.removedTerms.filter(function (t) { return t.length > 1; });
    let out = '';
    if (lang === 'en') {
      if (added.length > 0) out += '\n[Added in the reference]: ' + added.slice(0, limit).join(', ');
      if (removed.length > 0) out += '\n[Removed from the reference]: ' + removed.slice(0, limit).join(', ');
    } else {
      if (added.length > 0) out += '\n[Přidáno v referenci]: ' + added.slice(0, limit).join(', ');
      if (removed.length > 0) out += '\n[Odebráno z reference]: ' + removed.slice(0, limit).join(', ');
    }
    return out.trim();
  }

  /**
   * Build a structured prompt for Copilot / LLM (supports Czech and English).
   */
  function buildCopilotPrompt(params) {
    const lang = params.promptLang || 'cs';
    const targetLang = params.targetLangName || (lang === 'en' ? 'target language' : 'cílový jazyk');
    const refLang = params.refLangName || (lang === 'en' ? 'reference' : 'reference');
    const classification = params.classification || 'MODIFIED';
    const rowId = params.rowId || ((lang === 'en' ? 'Row ' : 'Řádek ') + (params.rowIndex !== undefined ? params.rowIndex + 2 : ''));
    const refOld = cleanText(params.refOld);
    const refNew = cleanText(params.refNew);
    const targetOld = cleanText(params.targetOld);
    const changesSummary = buildChangesSummary(params.diffResult, lang, 10);

    // Check if user provided a custom prompt template
    if (params.customTemplate && typeof params.customTemplate === 'string' && params.customTemplate.trim()) {
      const values = {
        TARGET_LANG: targetLang,
        REF_LANG: refLang,
        ROW_ID: rowId,
        REF_OLD: refOld,
        REF_NEW: refNew,
        TARGET_OLD: targetOld,
        CHANGES_SUMMARY: changesSummary,
        CLASSIFICATION: classification
      };
      return params.customTemplate.replace(/\{([A-Z_]+)\}/g, function (all, key) {
        return Object.prototype.hasOwnProperty.call(values, key) ? values[key] : all;
      }).trim();
    }

    if (lang === 'en') {
      if (classification === 'ADDED') {
        return [
          'You are an experienced professional translator.',
          'TASK: Translate the new paragraph into language: ' + targetLang + '.',
          'Paragraph: ' + rowId,
          '',
          'INSTRUCTIONS:',
          '1. Keep the tone, style and terminology of the existing document in ' + targetLang + '.',
          '2. Output ONLY the translated text (no introduction, comments or quotes).',
          '',
          '=== NEW REFERENCE TEXT (' + refLang + ') ===',
          refNew
        ].join('\n');
      }

      return [
        'You are an experienced professional translator.',
        'TASK: Update the approved translation in language ' + targetLang + ' so that it matches the new version of the reference text.',
        'Paragraph: ' + rowId,
        '',
        '=== CRITICAL INSTRUCTIONS ===',
        '1. Make ONLY the minimal necessary alterations that reflect the difference between the old and the new reference text.',
        '2. Keep all unchanged sentences, established terminology and phrasing untouched.',
        '3. Output ONLY the updated text (no introduction, comments, quotes or markdown).',
        '',
        '=== 1. REFERENCE TEXT – OLD VERSION (' + refLang + ') ===',
        refOld,
        '',
        '=== 2. REFERENCE TEXT – NEW VERSION ===',
        refNew + (changesSummary ? '\n' + changesSummary : ''),
        '',
        '=== 3. CURRENT APPROVED TRANSLATION (' + targetLang + ') ===',
        targetOld
      ].join('\n');
    }

    // Czech default
    if (classification === 'ADDED') {
      return [
        'Jste zkušený odborný překladatel.',
        'ÚKOL: Přeložte nový odstavec do jazyka: ' + targetLang + '.',
        'Odstavec: ' + rowId,
        '',
        'POKYNY:',
        '1. Dodržte styl, tón a terminologii stávajícího dokumentu v jazyce ' + targetLang + '.',
        '2. POUZE vraťte přeložený text – bez úvodu, komentářů a uvozovek.',
        '',
        '=== NOVÝ REFERENČNÍ TEXT (' + refLang + ') ===',
        refNew
      ].join('\n');
    }

    return [
      'Jste zkušený odborný překladatel.',
      'ÚKOL: Aktualizujte schválený překlad v jazyce ' + targetLang + ' tak, aby odpovídal nové verzi referenčního textu.',
      'Odstavec: ' + rowId,
      '',
      '=== DŮLEŽITÉ POKYNY ===',
      '1. Proveďte POUZE minimální nutné změny odpovídající rozdílu mezi starou a novou verzí referenčního textu.',
      '2. Ponechte nezměněné věty, formulace a zavedenou terminologii beze změny.',
      '3. POUZE vraťte výsledný text – bez úvodu, komentářů, uvozovek a markdownu.',
      '',
      '=== 1. REFERENČNÍ TEXT – STARÁ VERZE (' + refLang + ') ===',
      refOld,
      '',
      '=== 2. REFERENČNÍ TEXT – NOVÁ VERZE ===',
      refNew + (changesSummary ? '\n' + changesSummary : ''),
      '',
      '=== 3. DOSAVADNÍ SCHVÁLENÝ PŘEKLAD (' + targetLang + ') ===',
      targetOld
    ].join('\n');
  }

  /**
   * Returns `name`, or a suffixed variant of it, that is not yet in `taken`
   * (so a generated column never overwrites an existing column of the same name).
   */
  function uniqueColName(name, taken) {
    let candidate = name;
    let n = 2;
    while (taken.indexOf(candidate) !== -1) {
      candidate = name + ' (' + n + ')';
      n++;
    }
    taken.push(candidate);
    return candidate;
  }

  function cellValue(value) {
    return value !== undefined && value !== null ? value : '';
  }

  /**
   * Merges saved per-row translation work (targetNew, isVerified, status, notes) into freshly
   * generated stubs, so reference texts and the old target text always come from the current rows.
   * Rows whose reference classification changed since the work was saved lose their verification.
   */
  function rehydrateStubs(rows, savedStubs, targetCol, refOldCol, refNewCol, newCol) {
    const fresh = generateTargetStub(rows, targetCol, refOldCol, refNewCol, newCol);
    if (!Array.isArray(savedStubs)) return fresh;

    return fresh.map(function (stub, idx) {
      const saved = savedStubs[idx];
      if (!saved || typeof saved !== 'object') return stub;

      const classificationChanged = saved.classification && saved.classification !== stub.classification;
      if (typeof saved.targetNew === 'string') stub.targetNew = saved.targetNew;
      if (typeof saved.notes === 'string') stub.notes = saved.notes;

      if (classificationChanged) {
        stub.isVerified = false;
        stub.status = stub.targetNew ? 'draft' : stub.status;
      } else {
        if (typeof saved.isVerified === 'boolean') stub.isVerified = saved.isVerified;
        if (typeof saved.status === 'string' && saved.status) stub.status = saved.status;
      }
      return stub;
    });
  }

  /**
   * Returns the translation stubs for a language, generating defaults when the language was
   * never opened in the translation studio (unchanged rows then carry the old translation).
   */
  function resolveStubs(rows, translations, targetCol, mappings) {
    const existing = translations && translations[targetCol];
    if (Array.isArray(existing) && existing.length === rows.length) {
      return existing;
    }
    return rehydrateStubs(rows, existing, targetCol, mappings.refOldCol, mappings.refNewCol,
      getLanguageInfo(mappings, targetCol).newCol);
  }

  /**
   * Machine-readable status of one translated row: 'verified', 'removed', 'missing' or 'review'.
   */
  function getStubState(stub) {
    if (!stub) return 'missing';
    if (stub.classification === 'REMOVED' || stub.classification === 'EMPTY') return 'removed';
    if (stub.isVerified) return 'verified';
    return stub.targetNew && String(stub.targetNew).trim() ? 'review' : 'missing';
  }

  function sheetLabels(uiLang) {
    return uiLang === 'en'
      ? { row: 'Row', change: 'Reference change', status: 'Status',
          verified: 'Verified', review: 'Awaiting review', missing: 'Translation missing', removed: 'Removed in new version',
          UNCHANGED: 'Unchanged', MODIFIED: 'Modified', ADDED: 'Added', REMOVED: 'Removed', EMPTY: 'Empty',
          comprehensive: 'Comparison', clean: 'New_version', statusSheet: 'Translation_Status' }
      : { row: 'Řádek', change: 'Změna reference', status: 'Stav',
          verified: 'Ověřeno', review: 'Čeká na kontrolu', missing: 'Chybí překlad', removed: 'Zrušeno v nové verzi',
          UNCHANGED: 'Beze změny', MODIFIED: 'Upraveno', ADDED: 'Přidáno', REMOVED: 'Odstraněno', EMPTY: 'Prázdné',
          comprehensive: 'Srovnani', clean: 'Nova_verze', statusSheet: 'Stav_prekladu' };
  }

  /**
   * Prepares workbook data for export.
   *  - 'comprehensive': all original columns; every language gets its new version either in its
   *    existing new-version column or in a column inserted right after the old version.
   *  - 'clean': rows valid in the new version only; reference new version + new versions of all
   *    languages (+ other, ignored columns), without the old versions.
   */
  function prepareExportData(options) {
    const rawRows = options.rows || [];
    const headers = options.headers || [];
    const mappings = options.mappings || {};
    const translations = options.translations || {}; // { [targetCol]: [ { targetNew, status, ... } ] }
    const exportType = options.exportType === 'clean_2026' ? 'clean' : (options.exportType || 'comprehensive');
    const labels = sheetLabels(options.uiLang);
    const targetCols = (mappings.targetCols || []).filter(function (col) { return headers.indexOf(col) !== -1; });

    const infos = {};
    const stubsByCol = {};
    targetCols.forEach(function (col) {
      const info = getLanguageInfo(mappings, col);
      if (info.newCol && headers.indexOf(info.newCol) === -1) info.newCol = '';
      infos[col] = info;
      stubsByCol[col] = resolveStubs(rawRows, translations, col, mappings);
    });
    const existingNewCols = {}; // existing column → language col
    targetCols.forEach(function (col) {
      if (infos[col].newCol) existingNewCols[infos[col].newCol] = col;
    });

    function translatedText(col, idx) {
      const item = stubsByCol[col][idx];
      return item ? cellValue(item.targetNew) : '';
    }

    if (exportType === 'clean') {
      // Export only rows valid in the new version, omitting removed ones
      const taken = [];
      const sourceCols = []; // [{ from, to }] other columns + reference new version
      headers.forEach(function (h) {
        if (h === mappings.refOldCol || targetCols.indexOf(h) !== -1 || existingNewCols[h]) return;
        sourceCols.push({ from: h, to: uniqueColName(h, taken) });
      });
      const targetNames = targetCols.map(function (col) {
        return { col: col, to: uniqueColName(infos[col].newCol || infos[col].newColName, taken) };
      });

      const cleanHeaders = sourceCols.map(function (c) { return c.to; })
        .concat(targetNames.map(function (t) { return t.to; }));

      const cleanRows = [];
      rawRows.forEach(function (row, idx) {
        const classification = classifyRow(row[mappings.refOldCol], row[mappings.refNewCol]);
        if (classification === 'REMOVED' || classification === 'EMPTY') {
          return; // Skip removed or empty rows in the clean new version
        }

        const newRow = {};
        sourceCols.forEach(function (c) {
          newRow[c.to] = cellValue(row[c.from]);
        });
        targetNames.forEach(function (t) {
          newRow[t.to] = translatedText(t.col, idx);
        });
        cleanRows.push(newRow);
      });

      return {
        sheetName: labels.clean,
        headers: cleanHeaders,
        data: cleanRows,
        stubsByCol: stubsByCol
      };
    }

    // Comprehensive mode
    const taken = headers.slice();
    const newColNames = {};
    targetCols.forEach(function (col) {
      newColNames[col] = infos[col].newCol || uniqueColName(infos[col].newColName, taken);
    });

    const compHeaders = [];
    headers.forEach(function (h) {
      compHeaders.push(h);
      if (targetCols.indexOf(h) !== -1 && !infos[h].newCol) compHeaders.push(newColNames[h]);
    });

    const comprehensiveRows = rawRows.map(function (row, idx) {
      const newRow = {};
      headers.forEach(function (h) {
        newRow[h] = existingNewCols[h] ? translatedText(existingNewCols[h], idx) : cellValue(row[h]);
        if (targetCols.indexOf(h) !== -1 && !infos[h].newCol) {
          newRow[newColNames[h]] = translatedText(h, idx);
        }
      });
      return newRow;
    });

    return {
      sheetName: labels.comprehensive,
      headers: compHeaders,
      data: comprehensiveRows,
      newColNames: newColNames,
      stubsByCol: stubsByCol
    };
  }

  /**
   * Per-row verification status table (one column per language), the second sheet of the
   * comprehensive workbook.
   */
  function prepareStatusSheetData(options, exportData) {
    const rows = options.rows || [];
    const mappings = options.mappings || {};
    const labels = sheetLabels(options.uiLang);
    const cols = Object.keys(exportData.stubsByCol);
    const taken = [labels.row, labels.change];
    const colTitles = cols.map(function (col) { return uniqueColName(getLanguageInfo(mappings, col).label, taken); });

    const data = rows.map(function (row, idx) {
      const classification = classifyRow(row[mappings.refOldCol], row[mappings.refNewCol]);
      const out = {};
      out[labels.row] = rowNumberOf(options.rowNumbers, idx);
      out[labels.change] = labels[classification] || classification;
      cols.forEach(function (col, k) {
        out[colTitles[k]] = labels[getStubState(exportData.stubsByCol[col][idx])];
      });
      return out;
    });

    return {
      sheetName: labels.statusSheet,
      headers: [labels.row, labels.change].concat(colTitles),
      data: data
    };
  }

  /**
   * Partial workbook for one language: row number, chosen context columns (e.g. the reference),
   * old and new version of the language and its status. For the reference itself (langCol ===
   * refNewCol) the old and new reference columns are exported.
   */
  function prepareLanguageExportData(options) {
    const rows = options.rows || [];
    const headers = options.headers || [];
    const mappings = options.mappings || {};
    const labels = sheetLabels(options.uiLang);
    const langCol = options.langCol;
    const isReference = langCol === mappings.refNewCol;
    const info = isReference ? null : getLanguageInfo(mappings, langCol);
    const oldCol = isReference ? mappings.refOldCol : langCol;
    const newColExisting = isReference ? mappings.refNewCol : (info.newCol && headers.indexOf(info.newCol) !== -1 ? info.newCol : '');
    const stubs = isReference ? null : resolveStubs(rows, options.translations || {}, langCol, mappings);

    const contextCols = (options.contextCols || []).filter(function (c) {
      return headers.indexOf(c) !== -1 && c !== oldCol && c !== newColExisting;
    });
    const taken = [labels.row];
    const contextNames = contextCols.map(function (c) { return uniqueColName(c, taken); });
    const oldName = uniqueColName(oldCol, taken);
    const newName = uniqueColName(newColExisting || info.newColName, taken);
    const statusName = isReference ? null : uniqueColName(labels.status, taken);

    const data = rows.map(function (row, idx) {
      const out = {};
      out[labels.row] = rowNumberOf(options.rowNumbers, idx);
      contextCols.forEach(function (c, k) { out[contextNames[k]] = cellValue(row[c]); });
      out[oldName] = cellValue(row[oldCol]);
      out[newName] = isReference ? cellValue(row[mappings.refNewCol]) : cellValue(stubs[idx] && stubs[idx].targetNew);
      if (statusName) out[statusName] = labels[getStubState(stubs[idx])];
      return out;
    });

    const label = isReference ? String(langCol) : info.label;
    return {
      sheetName: String(label).replace(/[\\/?*[\]:]/g, '_').slice(0, 31) || 'Export',
      headers: [labels.row].concat(contextNames, [oldName, newName], statusName ? [statusName] : []),
      data: data,
      stubsByCol: stubs ? { [langCol]: stubs } : {}
    };
  }

  function buildSheet(XLSX, sheetData) {
    const ws = XLSX.utils.json_to_sheet(sheetData.data, { header: sheetData.headers });
    ws['!cols'] = sheetData.headers.map(function (h) {
      let maxLen = String(h).length;
      sheetData.data.forEach(function (r) {
        const len = r[h] !== undefined && r[h] !== null ? String(r[h]).length : 0;
        if (len > maxLen) maxLen = len;
      });
      return { wch: Math.min(Math.max(maxLen + 3, 10), 63) };
    });
    if (ws['!ref']) ws['!autofilter'] = { ref: ws['!ref'] };
    return ws;
  }

  /**
   * Builds the export workbook with SheetJS (passed in so this works in Node and the browser).
   * exportType: 'comprehensive' (+ status sheet), 'clean', or 'language' (partial, one language).
   */
  function buildExportWorkbook(XLSX, options) {
    const exportData = options.exportType === 'language'
      ? prepareLanguageExportData(options)
      : prepareExportData(options);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, buildSheet(XLSX, exportData), exportData.sheetName);

    let unverifiedCount = 0;
    Object.keys(exportData.stubsByCol).forEach(function (col) {
      exportData.stubsByCol[col].forEach(function (s) {
        const st = getStubState(s);
        if (st === 'review' || st === 'missing') unverifiedCount++;
      });
    });

    if (exportData.newColNames) {
      const statusData = prepareStatusSheetData(options, exportData);
      XLSX.utils.book_append_sheet(workbook, buildSheet(XLSX, statusData), statusData.sheetName);
    }

    return { workbook: workbook, exportData: exportData, unverifiedCount: unverifiedCount };
  }

  /**
   * Build a consolidated batch prompt for all modified/added paragraphs of one language.
   */
  function buildBatchCopilotPrompt(params) {
    const lang = params.promptLang || 'cs';
    const targetLang = params.targetLangName || (lang === 'en' ? 'target language' : 'cílový jazyk');
    const refLang = params.refLangName || 'reference';
    let items = params.items;
    if (!items || items.length === 0) {
      items = [];
      const rows = params.rows || [];
      const mappings = params.mappings || {};
      const translations = params.translations || {};
      const targetCol = params.targetCol || params.targetLangName;
      const stubs = (translations && targetCol && translations[targetCol]) ? translations[targetCol] : [];

      rows.forEach(function (row, idx) {
        const stub = stubs[idx];
        const refOld = row[mappings.refOldCol] !== undefined ? String(row[mappings.refOldCol]) : '';
        const refNew = row[mappings.refNewCol] !== undefined ? String(row[mappings.refNewCol]) : '';
        const classification = stub ? stub.classification : classifyRow(refOld, refNew);

        if (classification === 'MODIFIED' || classification === 'ADDED') {
          items.push({
            index: idx,
            rowId: rowTag(params.rowNumbers, idx),
            classification: classification,
            refOld: refOld,
            refNew: refNew,
            targetOld: stub ? stub.targetOld : (row[targetCol] || ''),
            diffResult: computeDiff(refOld, refNew)
          });
        }
      });
    }

    if (items.length === 0) {
      return lang === 'en'
        ? 'No modified or added paragraphs found to translate.'
        : 'Nebyly nalezeny žádné změněné ani nově přidané odstavce k překladu.';
    }

    const t = lang === 'en'
      ? {
        title: '=== BATCH TRANSLATION UPDATE ===',
        role: 'You are an experienced professional translator.',
        task: 'TASK: Provide updated or new translations of the following ' + items.length + ' paragraphs in language: ' + targetLang + '.',
        rulesTitle: '=== CRITICAL INSTRUCTIONS ===',
        rules: [
          '1. For EACH paragraph output its exact tag (e.g. [#12]) on its own line, followed immediately by the translated text.',
          '2. For MODIFIED paragraphs make ONLY the minimal changes reflecting the reference difference; keep unchanged phrases identical to the current translation.',
          '3. For ADDED paragraphs translate the new reference text faithfully.',
          '4. Do NOT output any introduction, commentary or filler – only the tagged paragraphs.'
        ],
        para: 'PARAGRAPH', added: 'NEWLY ADDED', modified: 'MODIFIED', typ: 'Type',
        newRef: 'NEW REFERENCE TEXT (' + refLang + '):', oldRef: 'OLD REFERENCE TEXT (' + refLang + '):',
        updRef: 'NEW REFERENCE TEXT:', approved: 'CURRENT APPROVED TRANSLATION (' + targetLang + '):'
      }
      : {
        title: '=== HROMADNÁ AKTUALIZACE PŘEKLADU ===',
        role: 'Jste zkušený odborný překladatel.',
        task: 'ÚKOL: Připravte aktualizované nebo nové překlady následujících ' + items.length + ' odstavců v jazyce: ' + targetLang + '.',
        rulesTitle: '=== PŘÍSNÉ POKYNY PRO FORMÁT ODPOVĚDI ===',
        rules: [
          '1. U KAŽDÉHO odstavce uveďte na samostatném řádku jeho přesnou značku (např. [#12]) a pod ni přeložený text.',
          '2. U odstavců typu UPRAVENO proveďte POUZE minimální změny odpovídající rozdílu v referenci; nezměněné formulace ponechte shodné s dosavadním překladem.',
          '3. U odstavců typu PŘIDÁNO přeložte nový referenční text věrně.',
          '4. Nevypisujte žádné úvody, komentáře ani zdvořilostní fráze – pouze označené odstavce.'
        ],
        para: 'ODSTAVEC', added: 'NOVĚ PŘIDÁNO', modified: 'UPRAVENO', typ: 'Typ',
        newRef: 'NOVÝ REFERENČNÍ TEXT (' + refLang + '):', oldRef: 'PŮVODNÍ REFERENČNÍ TEXT (' + refLang + '):',
        updRef: 'NOVÝ REFERENČNÍ TEXT:', approved: 'DOSAVADNÍ SCHVÁLENÝ PŘEKLAD (' + targetLang + '):'
      };

    const header = [t.title, t.role, t.task, '', t.rulesTitle].concat(t.rules, ['']).join('\n');

    const body = items.map(function (it, idx) {
      const idTag = it.rowId || ('#' + (idx + 1));
      if (it.classification === 'ADDED') {
        return [
          '----------------------------------------',
          t.para + ': [' + idTag + '] (' + t.typ + ': ' + t.added + ')',
          t.newRef,
          cleanText(it.refNew),
          ''
        ].join('\n');
      }

      const summary = buildChangesSummary(it.diffResult || computeDiff(it.refOld, it.refNew), lang, 8).replace(/\n/g, ' ');
      return [
        '----------------------------------------',
        t.para + ': [' + idTag + '] (' + t.typ + ': ' + t.modified + ')' + (summary ? ' ' + summary : ''),
        t.oldRef,
        cleanText(it.refOld),
        t.updRef,
        cleanText(it.refNew),
        t.approved,
        cleanText(it.targetOld),
        ''
      ].join('\n');
    }).join('\n');

    return header + '\n' + body;
  }

  // Batch prompt tags look like "Art. 04 (Section)" and sections may contain parentheses
  // themselves ("Art. 08 (KI (Neu 2026))"); the section part is not part of the ID.
  function stripSectionSuffix(id) {
    const str = String(id).trim();
    if (str.charAt(str.length - 1) !== ')') return str;
    let depth = 0;
    for (let i = str.length - 1; i >= 0; i--) {
      const ch = str.charAt(i);
      if (ch === ')') depth++;
      else if (ch === '(') {
        depth--;
        if (depth === 0) {
          const head = str.slice(0, i).trim();
          return head || str; // an ID that is only "(...)" stays as it is
        }
      }
    }
    return str;
  }

  function normalizeId(id) {
    return stripSectionSuffix(id).toLowerCase().replace(/[^\p{L}\p{N}]/gu, '');
  }

  function numericSignature(id) {
    const groups = stripSectionSuffix(id).match(/\d+/g);
    return groups ? groups.map(function (g) { return parseInt(g, 10); }).join('.') : '';
  }

  /**
   * Checks whether two article identifiers match (e.g. 'Art. 04' and 'Artikel 4').
   * All number groups must agree, so '1.2' never matches '1.3'.
   */
  function idsMatch(id1, id2) {
    if (!id1 || !id2) return false;
    const s1 = normalizeId(id1);
    if (s1 && s1 === normalizeId(id2)) return true;
    const n1 = numericSignature(id1);
    return n1 !== '' && n1 === numericSignature(id2);
  }

  function findExpected(expectedList, tag) {
    if (!expectedList || !Array.isArray(expectedList)) return null;
    function expId(exp) {
      return typeof exp === 'string' ? exp : (exp.id || exp.rowId || '');
    }
    const normTag = normalizeId(tag);
    // Exact (normalized) matches win over numeric matches
    return expectedList.find(function (exp) { return normTag && normalizeId(expId(exp)) === normTag; }) ||
      expectedList.find(function (exp) { return idsMatch(expId(exp), tag); }) ||
      null;
  }

  /**
   * Parse batch response from LLM / Copilot into mapped translations.
   * Matches identifiers like [Art. 04], Art. 04:, === Art. 04 ===, or JSON.
   * When an expected list is given, tags that match no expected article are not returned
   * (they are listed in `result.unmatched`), so text can never land in the wrong row.
   */
  function parseBatchTranslationResponse(responseText, expectedList) {
    const resultArray = [];
    const matchesMap = {};
    const unmatched = [];
    const hasExpected = Array.isArray(expectedList) && expectedList.length > 0;
    resultArray.unmatched = unmatched;

    function pushResult(tag, text, fallbackIndex) {
      const foundExpected = findExpected(expectedList, tag);
      if (hasExpected && !foundExpected) {
        unmatched.push(tag);
        return;
      }
      const resolvedIndex = foundExpected && typeof foundExpected === 'object' && foundExpected.index !== undefined
        ? foundExpected.index : fallbackIndex;
      const resolvedId = foundExpected && typeof foundExpected === 'object' && (foundExpected.id || foundExpected.rowId)
        ? (foundExpected.id || foundExpected.rowId) : (typeof foundExpected === 'string' ? foundExpected : tag);
      // A later block for the same article replaces an earlier one
      for (let k = resultArray.length - 1; k >= 0; k--) {
        if (resultArray[k].index === resolvedIndex) resultArray.splice(k, 1);
      }
      resultArray.push({ index: resolvedIndex, id: resolvedId, text: text });
    }

    if (!responseText || typeof responseText !== 'string') {
      resultArray.matches = matchesMap;
      resultArray.count = 0;
      return resultArray;
    }

    const text = responseText.trim();

    // 1. Try parsing JSON if user or LLM returned JSON array
    if (text.startsWith('[') && text.endsWith(']')) {
      try {
        const parsed = JSON.parse(text);
        if (Array.isArray(parsed)) {
          parsed.forEach(function (it, idx) {
            const key = it.id || it.article || it.identifier || it.rowId;
            const val = it.text || it.translation || it.content;
            if (key && val) {
              const cleanedKey = cleanIdentifier(key);
              matchesMap[cleanedKey] = String(val).trim();
              pushResult(cleanedKey, String(val).trim(), idx);
            }
          });
          if (resultArray.length > 0 || unmatched.length > 0) {
            resultArray.matches = matchesMap;
            resultArray.count = resultArray.length;
            return resultArray;
          }
        }
      } catch (e) {
        // Fallback to text parsing
      }
    }

    // Regex pattern matching tags: [Art. 04], Art. 04:, === Art. 04 ===, ### Art. 04, etc.
    const sectionRegex = /(?:^|\n)\s*(?:\[([^\]]+)\]|={3,}\s*\[?([^=\]\n]+)\]?\s*={3,}|#{1,4}\s*\[?([^#\n\]]+)\]?|(?:Art(?:ikel)?|Čl(?:ánek)?|Article|Row|Řádek)\s*\.?\s*(\d+[a-z]?)\s*[:\-])([^\n]*)/gi;

    let match;
    const slices = [];
    while ((match = sectionRegex.exec(text)) !== null) {
      let rawTag = (match[1] || match[2] || match[3] || '').trim();
      if (!rawTag && match[4]) {
        rawTag = 'Art. ' + match[4].trim();
      }
      const inlineRest = (match[5] || '').trim();
      slices.push({
        tag: cleanIdentifier(rawTag),
        matchIndex: match.index,
        textStartIndex: match.index + match[0].length,
        inlineRest: inlineRest
      });
    }

    if (slices.length > 0) {
      for (let i = 0; i < slices.length; i++) {
        const cur = slices[i];
        const next = slices[i + 1];
        const endIdx = next ? next.matchIndex : text.length;
        let blockText = text.slice(cur.textStartIndex, endIdx).trim();

        // If there was text on the same line as tag, prepend it
        if (cur.inlineRest) {
          blockText = (cur.inlineRest + '\n' + blockText).trim();
        }

        // Clean out any trailing separators like ----------------
        blockText = blockText.replace(/-{3,}$/g, '').trim();

        if (cur.tag && blockText) {
          matchesMap[cur.tag] = blockText;
          pushResult(cur.tag, blockText, i);
        }
      }
    }

    resultArray.matches = matchesMap;
    resultArray.count = resultArray.length;
    return resultArray;
  }

  function cleanIdentifier(id) {
    if (!id) return '';
    return String(id)
      .replace(/^[\[\s=]+|[\]\s=:]+$/g, '')
      .replace(/\s+/g, ' ')
      .trim();
  }

  /**
   * Change summary (old → new version) for ONE language: the reference (langCol === refNewCol or
   * omitted) or a translated language (langCol = its old-version column, new text = translation).
   * Returns KPIs, item lists, Markdown and a Word/print-ready HTML document.
   */
  function generateManagerialChangeSummary(params) {
    const rows = params.rows || [];
    const mappings = params.mappings || {};
    const isEn = (params.uiLang || 'cs') === 'en';
    const langCol = params.langCol || mappings.refNewCol;
    const isReference = langCol === mappings.refNewCol;
    const label = getLanguageLabel(mappings, langCol);
    const documentName = params.documentName || params.companyName || '';
    const stubs = isReference ? null : resolveStubs(rows, params.translations || {}, langCol, mappings);

    const L = isEn
      ? {
        title: 'Change summary – ' + label + ' (old → new version)', document: 'Document', language: 'Language',
        reference: 'reference', refCols: 'Reference columns', date: 'Date', kpiTitle: '1. Change statistics',
        metric: 'Category', count: 'Count', share: 'Share', total: 'Total paragraphs', unchanged: 'Unchanged',
        modified: 'Modified', added: 'Added', removed: 'Removed', scope: 'Revision scope',
        scopeText: 'of paragraphs were changed, added or removed.', statusTitle: 'Translation status',
        verified: 'Verified', review: 'Awaiting review', missing: 'Translation missing',
        modTitle: '2. Modified paragraphs', addTitle: '3. Added paragraphs', remTitle: '4. Removed paragraphs',
        row: 'Row', similarity: 'Similarity', addedTerms: 'Added', removedTerms: 'Removed',
        oldVersion: 'Old version', newVersion: 'New version', newVersionDiff: 'New version (changes highlighted)',
        missingText: 'Translation missing', print: '🖨️ Print / Save as PDF', close: '✕ Close'
      }
      : {
        title: 'Shrnutí změn – ' + label + ' (stará → nová verze)', document: 'Dokument', language: 'Jazyk',
        reference: 'reference', refCols: 'Referenční sloupce', date: 'Datum', kpiTitle: '1. Statistika změn',
        metric: 'Kategorie', count: 'Počet', share: 'Podíl', total: 'Celkem odstavců', unchanged: 'Beze změny',
        modified: 'Upravené', added: 'Přidané', removed: 'Zrušené', scope: 'Rozsah revize',
        scopeText: 'odstavců bylo upraveno, přidáno nebo zrušeno.', statusTitle: 'Stav překladu',
        verified: 'Ověřeno', review: 'Čeká na kontrolu', missing: 'Chybí překlad',
        modTitle: '2. Upravené odstavce', addTitle: '3. Přidané odstavce', remTitle: '4. Zrušené odstavce',
        row: 'Řádek', similarity: 'Míra shody', addedTerms: 'Přidáno', removedTerms: 'Odebráno',
        oldVersion: 'Stará verze', newVersion: 'Nová verze', newVersionDiff: 'Nová verze (vyznačené změny)',
        missingText: 'Překlad chybí', print: '🖨️ Vytisknout / Uložit PDF', close: '✕ Zavřít'
      };

    let total = 0;
    let unchanged = 0;
    const status = { verified: 0, review: 0, missing: 0 };
    const modifiedItems = [];
    const addedItems = [];
    const removedItems = [];

    rows.forEach(function (row, idx) {
      const refOld = row[mappings.refOldCol] !== undefined ? String(row[mappings.refOldCol]) : '';
      const refNew = row[mappings.refNewCol] !== undefined ? String(row[mappings.refNewCol]) : '';
      const classification = classifyRow(refOld, refNew);
      if (classification === 'EMPTY') return;
      total++;

      const stub = stubs ? stubs[idx] : null;
      const oldText = isReference ? refOld : String(cellValue(row[langCol]));
      const newText = isReference ? refNew : String(cellValue(stub && stub.targetNew));
      const rowNumber = rowNumberOf(params.rowNumbers, idx);
      const missing = !isReference && classification !== 'REMOVED' && !newText.trim();
      if (stub && classification !== 'REMOVED') {
        const st = getStubState(stub);
        if (status[st] !== undefined) status[st]++;
      }

      if (classification === 'UNCHANGED') {
        unchanged++;
      } else if (classification === 'MODIFIED') {
        const diffRes = computeDiff(oldText, newText);
        modifiedItems.push({
          index: idx,
          rowNumber: rowNumber,
          oldText: oldText,
          newText: newText,
          missing: missing,
          diffResult: diffRes,
          addedTerms: diffRes.summary.addedTerms.filter(function (t) { return t.length > 1; }),
          removedTerms: diffRes.summary.removedTerms.filter(function (t) { return t.length > 1; }),
          similarity: diffRes.stats.similarity
        });
      } else if (classification === 'ADDED') {
        addedItems.push({ index: idx, rowNumber: rowNumber, text: newText, missing: missing });
      } else if (classification === 'REMOVED') {
        removedItems.push({ index: idx, rowNumber: rowNumber, text: oldText });
      }
    });

    const modified = modifiedItems.length;
    const added = addedItems.length;
    const removed = removedItems.length;
    const pct = function (n) { return total > 0 ? Math.round((n / total) * 100) : 0; };
    const pctChanged = pct(modified + added + removed);
    const today = new Date().toISOString().split('T')[0];
    const languageText = label + (isReference ? ' (' + L.reference + ')' : '');

    // 1. Markdown
    const md = [];
    md.push('# ' + L.title);
    md.push('');
    if (documentName) md.push('**' + L.document + ':** ' + documentName + '  ');
    md.push('**' + L.language + ':** ' + languageText + '  ');
    md.push('**' + L.refCols + ':** ' + mappings.refOldCol + ' → ' + mappings.refNewCol + '  ');
    md.push('**' + L.date + ':** ' + today);
    md.push('');
    md.push('## ' + L.kpiTitle);
    md.push('');
    md.push('| ' + L.metric + ' | ' + L.count + ' | ' + L.share + ' |');
    md.push('| :--- | :--- | :--- |');
    md.push('| **' + L.total + '** | ' + total + ' | 100% |');
    md.push('| ' + L.unchanged + ' | ' + unchanged + ' | ' + pct(unchanged) + '% |');
    md.push('| ' + L.modified + ' | ' + modified + ' | ' + pct(modified) + '% |');
    md.push('| ' + L.added + ' | ' + added + ' | ' + pct(added) + '% |');
    md.push('| ' + L.removed + ' | ' + removed + ' | ' + pct(removed) + '% |');
    md.push('');
    md.push('> **' + L.scope + ':** ' + pctChanged + ' % ' + L.scopeText);
    md.push('');
    if (!isReference) {
      md.push('### ' + L.statusTitle);
      md.push('');
      md.push('| ' + L.metric + ' | ' + L.count + ' |');
      md.push('| :--- | :--- |');
      md.push('| ' + L.verified + ' | ' + status.verified + ' |');
      md.push('| ' + L.review + ' | ' + status.review + ' |');
      md.push('| ' + L.missing + ' | ' + status.missing + ' |');
      md.push('');
    }
    if (modifiedItems.length > 0) {
      md.push('## ' + L.modTitle);
      md.push('');
      modifiedItems.forEach(function (m) {
        md.push('### ' + L.row + ' ' + m.rowNumber);
        if (m.missing) {
          md.push('- **' + L.missingText + '**');
        } else {
          md.push('- **' + L.similarity + ':** ' + Math.round(m.similarity * 100) + ' %');
          if (m.addedTerms.length > 0) md.push('- **' + L.addedTerms + ':** `' + m.addedTerms.join('`, `') + '`');
          if (m.removedTerms.length > 0) md.push('- **' + L.removedTerms + ':** `' + m.removedTerms.join('`, `') + '`');
        }
        md.push('');
        md.push('**' + L.oldVersion + ':**');
        md.push(mdQuote(m.oldText));
        md.push('');
        if (!m.missing) {
          md.push('**' + L.newVersion + ':**');
          md.push(mdQuote(m.newText));
          md.push('');
        }
      });
    }
    if (addedItems.length > 0) {
      md.push('## ' + L.addTitle);
      md.push('');
      addedItems.forEach(function (a) {
        md.push('### ➕ ' + L.row + ' ' + a.rowNumber);
        md.push(a.missing ? '*' + L.missingText + '*' : mdQuote(a.text));
        md.push('');
      });
    }
    if (removedItems.length > 0) {
      md.push('## ' + L.remTitle);
      md.push('');
      removedItems.forEach(function (r) {
        md.push('### 🗑️ ' + L.row + ' ' + r.rowNumber);
        md.push(mdQuote(r.text));
        md.push('');
      });
    }
    const markdownDoc = md.join('\n');

    // 2. HTML (Word document; with a print toolbar when params.printable is set)
    const esc = escapeHtml;
    const missingHtml = '<em class="missing">' + esc(L.missingText) + '</em>';
    const kpiRow = function (name, n, cls) {
      return '<tr><td>' + (cls ? '<span class="badge ' + cls + '">' + esc(name) + '</span>' : esc(name)) + '</td><td>' + n + '</td><td>' + pct(n) + '%</td></tr>';
    };
    const htmlDoc = `<!DOCTYPE html>
<html lang="${isEn ? 'en' : 'cs'}" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns="http://www.w3.org/TR/REC-html40">
<head>
  <meta charset="UTF-8">
  <title>${esc(L.title)}</title>
  <!--[if gte mso 9]>
  <xml>
    <w:WordDocument>
      <w:View>Print</w:View>
      <w:Zoom>100</w:Zoom>
      <w:DoNotOptimizeForBrowser/>
    </w:WordDocument>
  </xml>
  <![endif]-->
  <style>
    @page { mso-page-orientation: portrait; size: 21cm 29.7cm; margin: 2cm; }
    body { font-family: Calibri, Arial, sans-serif; font-size: 11pt; line-height: 1.5; color: #1e293b; margin: 30px; }
    h1 { font-size: 20pt; color: #0f172a; border-bottom: 2pt solid #2563eb; padding-bottom: 6pt; margin-bottom: 12pt; }
    h2 { font-size: 14pt; color: #1e40af; margin-top: 18pt; margin-bottom: 8pt; border-bottom: 1pt solid #cbd5e1; padding-bottom: 4pt; }
    h3 { font-size: 12pt; color: #0f172a; margin-top: 12pt; margin-bottom: 4pt; }
    table { width: 100%; border-collapse: collapse; margin-top: 8pt; margin-bottom: 14pt; }
    th { background-color: #f1f5f9; color: #334155; font-weight: bold; border: 1pt solid #cbd5e1; padding: 6pt 8pt; text-align: left; }
    td { border: 1pt solid #e2e8f0; padding: 6pt 8pt; vertical-align: top; }
    .kpi-table th { background-color: #2563eb; color: #ffffff; }
    .badge { font-weight: bold; padding: 2pt 5pt; border-radius: 3pt; font-size: 9pt; }
    .badge-mod { background-color: #fef3c7; color: #b45309; }
    .badge-add { background-color: #dcfce7; color: #15803d; }
    .badge-rem { background-color: #fee2e2; color: #b91c1c; }
    blockquote { border-left: 3pt solid #94a3b8; margin: 4pt 0 10pt 0; padding-left: 10pt; color: #334155; background: #f8fafc; }
    ins { background-color: #dcfce7; color: #15803d; text-decoration: none; font-weight: bold; }
    del { background-color: #fee2e2; color: #b91c1c; text-decoration: line-through; }
    .missing { color: #b91c1c; }
    .no-print-toolbar { position: sticky; top: 0; background: #ffffff; border-bottom: 1px solid #e2e8f0; padding: 0.75rem 1.5rem; margin: -30px -30px 20px; display: flex; justify-content: space-between; align-items: center; }
    .no-print-toolbar button { border-radius: 6px; padding: 0.45rem 1rem; cursor: pointer; font-weight: 700; border: 1px solid #cbd5e1; background: #f1f5f9; }
    .no-print-toolbar .btn-print { background: #2563eb; color: #fff; border-color: #2563eb; }
    @media print { .no-print-toolbar { display: none !important; } body { margin: 0; } h2, h3, tr { break-inside: avoid; page-break-inside: avoid; } }
  </style>
</head>
<body>
  ${params.printable ? `<div class="no-print-toolbar"><strong>${esc(L.title)}</strong><span><button class="btn-print" onclick="window.print();">${L.print}</button> <button onclick="window.close();">${L.close}</button></span></div>` : ''}
  <h1>${esc(L.title)}</h1>
  <p>
    ${documentName ? `<strong>${esc(L.document)}:</strong> ${esc(documentName)}<br>` : ''}
    <strong>${esc(L.language)}:</strong> ${esc(languageText)}<br>
    <strong>${esc(L.refCols)}:</strong> ${esc(mappings.refOldCol)} &rarr; ${esc(mappings.refNewCol)}<br>
    <strong>${esc(L.date)}:</strong> ${today}
  </p>

  <h2>${esc(L.kpiTitle)}</h2>
  <table class="kpi-table">
    <tr><th>${esc(L.metric)}</th><th>${esc(L.count)}</th><th>${esc(L.share)}</th></tr>
    <tr><td><strong>${esc(L.total)}</strong></td><td><strong>${total}</strong></td><td><strong>100%</strong></td></tr>
    ${kpiRow(L.unchanged, unchanged, '')}
    ${kpiRow(L.modified, modified, 'badge-mod')}
    ${kpiRow(L.added, added, 'badge-add')}
    ${kpiRow(L.removed, removed, 'badge-rem')}
  </table>
  ${isReference ? '' : `
  <h3>${esc(L.statusTitle)}</h3>
  <table>
    <tr><th>${esc(L.metric)}</th><th>${esc(L.count)}</th></tr>
    <tr><td>${esc(L.verified)}</td><td>${status.verified}</td></tr>
    <tr><td>${esc(L.review)}</td><td>${status.review}</td></tr>
    <tr><td>${esc(L.missing)}</td><td>${status.missing}</td></tr>
  </table>`}

  ${modifiedItems.length > 0 ? `<h2>${esc(L.modTitle)}</h2>` : ''}
  ${modifiedItems.map(function (m) {
    return `
      <h3>${esc(L.row)} ${m.rowNumber} <span class="badge badge-mod">${esc(L.modified)}</span></h3>
      <table>
        <tr><th style="width:50%;">${esc(L.oldVersion)}</th><th style="width:50%;">${esc(L.newVersionDiff)}</th></tr>
        <tr><td>${escapeHtmlMultiline(m.oldText)}</td><td>${m.missing ? missingHtml : renderDiffHtml(m.diffResult, 'inline').replace(/\r?\n/g, '<br>')}</td></tr>
      </table>`;
  }).join('')}

  ${addedItems.length > 0 ? `<h2>${esc(L.addTitle)}</h2>` + addedItems.map(function (a) {
    return `<h3>➕ ${esc(L.row)} ${a.rowNumber} <span class="badge badge-add">${esc(L.added)}</span></h3>
      <blockquote>${a.missing ? missingHtml : escapeHtmlMultiline(a.text)}</blockquote>`;
  }).join('') : ''}

  ${removedItems.length > 0 ? `<h2>${esc(L.remTitle)}</h2>` + removedItems.map(function (r) {
    return `<h3>🗑️ ${esc(L.row)} ${r.rowNumber} <span class="badge badge-rem">${esc(L.removed)}</span></h3>
      <blockquote>${escapeHtmlMultiline(r.text)}</blockquote>`;
  }).join('') : ''}
</body>
</html>`;

    const stats = {
      totalRows: total,
      unchangedCount: unchanged,
      modifiedCount: modified,
      addedCount: added,
      removedCount: removed,
      pctChanged: pctChanged
    };
    return {
      language: label,
      isReference: isReference,
      stats: stats,
      kpis: { total: total, unchanged: unchanged, modified: modified, added: added, removed: removed, pctChanged: pctChanged },
      translationStatus: isReference ? null : status,
      modifiedItems: modifiedItems,
      addedItems: addedItems,
      removedItems: removedItems,
      markdown: markdownDoc,
      htmlDoc: htmlDoc
    };
  }

  return {
    escapeHtml: escapeHtml,
    escapeHtmlMultiline: escapeHtmlMultiline,
    cleanText: cleanText,
    normalizeForComparison: normalizeForComparison,
    classifyRow: classifyRow,
    computeDiff: computeDiff,
    renderDiffHtml: renderDiffHtml,
    rowNumberOf: rowNumberOf,
    rowTag: rowTag,
    defaultNewColName: defaultNewColName,
    defaultLanguageLabel: defaultLanguageLabel,
    getLanguageInfo: getLanguageInfo,
    getLanguageLabel: getLanguageLabel,
    generateTargetStub: generateTargetStub,
    rehydrateStubs: rehydrateStubs,
    resolveStubs: resolveStubs,
    getStubState: getStubState,
    idsMatch: idsMatch,
    buildCopilotPrompt: buildCopilotPrompt,
    buildBatchCopilotPrompt: buildBatchCopilotPrompt,
    parseBatchTranslationResponse: parseBatchTranslationResponse,
    generateManagerialChangeSummary: generateManagerialChangeSummary,
    prepareExportData: prepareExportData,
    prepareLanguageExportData: prepareLanguageExportData,
    buildExportWorkbook: buildExportWorkbook
  };
}));
