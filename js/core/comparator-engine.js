/**
 * ComparatorEngine - Core logic for text diffing, row classification, 
 * target stub assembly, and prompt generation.
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
   * Generate initial translation stub for a target language.
   * 
   * @param {Array<Object>} rows - All table rows
   * @param {string} targetCol - Column name of the target language (e.g. 'CS 2024')
   * @param {string} refOldCol - Reference 2024 column (e.g. 'DE 2024')
   * @param {string} refNewCol - Reference 2026 column (e.g. 'DE 2026')
   * @returns {Array<Object>} Stubs for each row
   */
  function generateTargetStub(rows, targetCol, refOldCol, refNewCol) {
    if (!rows || !Array.isArray(rows)) return [];

    return rows.map(function (row, idx) {
      const refOld = row[refOldCol];
      const refNew = row[refNewCol];
      const targetOld = row[targetCol] !== undefined && row[targetCol] !== null ? String(row[targetCol]) : '';
      const classification = classifyRow(refOld, refNew);

      let initialNewText = '';
      let status = 'pending';
      let isVerified = false;

      switch (classification) {
        case 'UNCHANGED':
          // Auto-copy verified translation from 2024
          initialNewText = targetOld;
          status = 'auto_copied';
          isVerified = true;
          break;

        case 'REMOVED':
          // Removed paragraph in 2026
          initialNewText = '';
          status = 'removed';
          isVerified = true;
          break;

        case 'ADDED':
          // Completely new paragraph, requires fresh translation
          initialNewText = '';
          status = 'pending_translation';
          isVerified = false;
          break;

        case 'MODIFIED':
          // Modified paragraph, requires review & adjustment
          // By default, leave empty or provide targetOld as draft
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

      return {
        rowIndex: idx,
        targetCol: targetCol,
        classification: classification,
        refOld: refOld || '',
        refNew: refNew || '',
        targetOld: targetOld,
        targetNew: initialNewText,
        status: status,
        isVerified: isVerified,
        notes: ''
      };
    });
  }

  /**
   * Build a structured prompt for Copilot / LLM (supports Czech and English).
   */
  function buildCopilotPrompt(params) {
    const lang = params.promptLang || 'cs';
    const targetLang = params.targetLangName || (lang === 'en' ? 'target language' : 'cílového jazyka');
    const classification = params.classification || 'MODIFIED';
    const rowId = params.rowId || ((lang === 'en' ? 'Row ' : 'Řádek ') + (params.rowIndex !== undefined ? params.rowIndex + 1 : ''));
    const refOld = cleanText(params.refOld);
    const refNew = cleanText(params.refNew);
    const targetOld = cleanText(params.targetOld);
    const diffResult = params.diffResult;

    // Build changes summary
    let changesSummary = '';
    if (diffResult && diffResult.summary) {
      const added = diffResult.summary.addedTerms.filter(function(t){ return t.length > 1; });
      const removed = diffResult.summary.removedTerms.filter(function(t){ return t.length > 1; });
      if (lang === 'en') {
        if (added.length > 0) changesSummary += '\n[Key Added German Words/Phrases (2026)]: ' + added.slice(0, 10).join(', ');
        if (removed.length > 0) changesSummary += '\n[Key Removed German Words/Phrases (2024)]: ' + removed.slice(0, 10).join(', ');
      } else {
        if (added.length > 0) changesSummary += '\n[Klíčová přidaná slova/fráze v němčině 2026]: ' + added.slice(0, 10).join(', ');
        if (removed.length > 0) changesSummary += '\n[Klíčová odstraněná slova/fráze z němčiny 2024]: ' + removed.slice(0, 10).join(', ');
      }
    }

    // Check if user provided a custom prompt template
    if (params.customTemplate && typeof params.customTemplate === 'string' && params.customTemplate.trim()) {
      let templ = params.customTemplate;
      templ = templ.replace(/\{TARGET_LANG\}/g, targetLang);
      templ = templ.replace(/\{ROW_ID\}/g, rowId);
      templ = templ.replace(/\{REF_OLD\}/g, refOld);
      templ = templ.replace(/\{REF_NEW\}/g, refNew);
      templ = templ.replace(/\{TARGET_OLD\}/g, targetOld);
      templ = templ.replace(/\{CHANGES_SUMMARY\}/g, changesSummary ? changesSummary.trim() : '');
      templ = templ.replace(/\{CLASSIFICATION\}/g, classification);
      return templ.trim();
    }

    if (lang === 'en') {
      if (classification === 'ADDED') {
        return [
          'You are an expert legal and corporate translator specializing in Codes of Ethics.',
          'TASK: Translate the new paragraph from the 2026 Code of Conduct into language: ' + targetLang + '.',
          'Paragraph Identifier: ' + rowId,
          '',
          'TRANSLATION INSTRUCTIONS:',
          '1. Maintain the formal, legally precise, and ethical tone of the corporate Code of Conduct in ' + targetLang + '.',
          '2. ONLY output the clean translated text ready for direct spreadsheet insertion (no preamble or quotes).',
          '',
          '=== NEW GERMAN SOURCE TEXT (2026 Reference Version) ===',
          refNew
        ].join('\n');
      }

      // MODIFIED English
      return [
        'You are an expert legal and corporate translator specializing in Codes of Ethics.',
        'TASK: Update the existing certified legacy translation for the 2026 version in language: ' + targetLang + '.',
        'Paragraph Identifier: ' + rowId,
        '',
        '=== CRITICAL INSTRUCTIONS ===',
        '1. Make ONLY the minimal necessary alterations strictly required to reflect the changes between the 2024 and 2026 German reference versions.',
        '2. Keep all unchanged sentences, established terminology, and stylistic phrasing untouched to preserve continuity.',
        '3. Output ONLY the raw updated text ready for direct spreadsheet insertion (no preamble, commentary, or quotes).',
        '',
        '=== 1. ORIGINAL GERMAN TEXT (2024 Reference) ===',
        refOld,
        '',
        '=== 2. UPDATED GERMAN TEXT (2026 Reference) ===',
        refNew + (changesSummary ? '\n' + changesSummary.trim() : ''),
        '',
        '=== 3. CERTIFIED LEGACY TRANSLATION (2024 Version in ' + targetLang + ') ===',
        targetOld
      ].join('\n');
    }

    // Czech default
    if (classification === 'ADDED') {
      return [
        'Jste špičkový odborný překladatel firemních a právních dokumentů (Etický kodex).',
        'ÚKOL: Přeložte nový odstavec z Etického kodexu 2026 do jazyka: ' + targetLang + '.',
        'Identifikátor odstavce: ' + rowId,
        '',
        'POKYNY PRO PŘEKLAD:',
        '1. Dodržte formální, právně přesný a etický styl firemního kodexu v jazyce ' + targetLang + '.',
        '2. POUZE přeložte text – nevkládejte žádný úvod, vysvětlení, komentáře ani uvozovky.',
        '3. Výstupem musí být výhradně čistý přeložený text připravený pro vložení do buňky tabulky.',
        '',
        '=== NOVÝ TEXT V NĚMČINĚ (Referenční verze 2026) ===',
        refNew
      ].join('\n');
    }

    // MODIFIED case (Czech)
    return [
      'Jste špičkový odborný překladatel firemních a právních dokumentů (Etický kodex).',
      'ÚKOL: Aktualizujte stávající schválený překlad odstavce pro rok 2026 v jazyce: ' + targetLang + '.',
      'Identifikátor odstavce: ' + rowId,
      '',
      '=== DŮLEŽITÉ POKYNY PRO PŘEKLAD ===',
      '1. Proveďte v dosavadním překladu POUZE minimální nutné změny, které přesně odpovídají rozdílům mezi německou verzí 2024 a 2026.',
      '2. Ponechte všechny nezměněné věty, formulace a zavedenou terminologii beze změny pro zachování kontinuity.',
      '3. POUZE vraťte výsledný text – nevkládejte žádný úvod, vysvětlivky, komentáře ani markdown bloky.',
      '',
      '=== 1. PŮVODNÍ NĚMECKÝ TEXT (Referenční verze 2024) ===',
      refOld,
      '',
      '=== 2. NOVÝ NĚMECKÝ TEXT (Referenční verze 2026) ===',
      refNew + (changesSummary ? '\n' + changesSummary.trim() : ''),
      '',
      '=== 3. DOSAVADNÍ OVĚŘENÝ PŘEKLAD (Verze 2024 v jazyce ' + targetLang + ') ===',
      targetOld
    ].join('\n');
  }

  /**
   * Helper to format a column header for the 2026 version.
   */
  function format2026ColName(col) {
    if (!col) return '2026';
    if (/2024/i.test(col)) {
      return col.replace(/2024/gi, '2026');
    }
    if (/2026/.test(col)) {
      return col;
    }
    return col + ' (2026)';
  }

  /**
   * Returns `name`, or a suffixed variant of it, that is not yet in `taken`
   * (so a generated 2026 column never overwrites an existing column of the same name).
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
   * generated stubs, so reference texts and the 2024 target text always come from the current rows.
   * Rows whose reference classification changed since the work was saved lose their verification.
   */
  function rehydrateStubs(rows, savedStubs, targetCol, refOldCol, refNewCol) {
    const fresh = generateTargetStub(rows, targetCol, refOldCol, refNewCol);
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
   * Returns the translation stubs for a target column, generating defaults when the language
   * was never opened in the translation studio (unchanged rows then carry the 2024 text).
   */
  function resolveStubs(rows, translations, targetCol, mappings) {
    const existing = translations && translations[targetCol];
    if (Array.isArray(existing) && existing.length === rows.length) {
      return existing;
    }
    return rehydrateStubs(rows, existing, targetCol, mappings.refOldCol, mappings.refNewCol);
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

  /**
   * Prepares workbook data for export.
   */
  function prepareExportData(options) {
    const rawRows = options.rows || [];
    const headers = options.headers || [];
    const mappings = options.mappings || {};
    const translations = options.translations || {}; // { [targetCol]: [ { targetNew, status, ... } ] }
    const exportType = options.exportType || 'comprehensive'; // 'comprehensive' | 'clean_2026'
    const targetCols = (mappings.targetCols || []).filter(function (col) { return headers.indexOf(col) !== -1; });

    const stubsByCol = {};
    targetCols.forEach(function (col) {
      stubsByCol[col] = resolveStubs(rawRows, translations, col, mappings);
    });

    function translatedText(col, idx) {
      const item = stubsByCol[col][idx];
      return item ? cellValue(item.targetNew) : '';
    }

    if (exportType === 'clean_2026') {
      // Export only valid 2026 rows, omitting removed ones
      const taken = [];
      const sourceCols = []; // [{ from, to }] metadata + new reference columns
      headers.forEach(function (h) {
        if (h === mappings.refOldCol || targetCols.indexOf(h) !== -1) return;
        sourceCols.push({ from: h, to: null });
      });
      // Reserve original names first so generated names never collide with them
      sourceCols.forEach(function (c) {
        if (c.from !== mappings.refNewCol) c.to = uniqueColName(c.from, taken);
      });
      sourceCols.forEach(function (c) {
        if (c.from === mappings.refNewCol) c.to = uniqueColName(format2026ColName(c.from), taken);
      });
      const targetNames = targetCols.map(function (col) {
        return { col: col, to: uniqueColName(format2026ColName(col), taken) };
      });

      const cleanHeaders = sourceCols.map(function (c) { return c.to; })
        .concat(targetNames.map(function (t) { return t.to; }));

      const cleanRows = [];
      rawRows.forEach(function (row, idx) {
        const classification = classifyRow(row[mappings.refOldCol], row[mappings.refNewCol]);
        if (classification === 'REMOVED' || classification === 'EMPTY') {
          return; // Skip removed or empty rows in clean 2026 version
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
        sheetName: 'Kodex_2026_Final',
        headers: cleanHeaders,
        data: cleanRows,
        stubsByCol: stubsByCol
      };
    }

    // Comprehensive mode: keeps all original columns, inserts new 2026 column next to each target language
    const taken = headers.slice();
    const newColNames = {};
    targetCols.forEach(function (col) {
      newColNames[col] = uniqueColName(format2026ColName(col), taken);
    });

    const compHeaders = [];
    headers.forEach(function (h) {
      compHeaders.push(h);
      if (newColNames[h]) compHeaders.push(newColNames[h]);
    });

    const comprehensiveRows = rawRows.map(function (row, idx) {
      const newRow = {};
      headers.forEach(function (h) {
        newRow[h] = cellValue(row[h]);
        // If this column is a target language, append the 2026 column right after it
        if (newColNames[h]) {
          newRow[newColNames[h]] = translatedText(h, idx);
        }
      });
      return newRow;
    });

    return {
      sheetName: 'Srovnani_2024_2026_Komplet',
      headers: compHeaders,
      data: comprehensiveRows,
      newColNames: newColNames,
      stubsByCol: stubsByCol
    };
  }

  /**
   * Builds the per-row verification status table (one column per target language) used
   * as the second sheet of the comprehensive workbook.
   */
  function prepareStatusSheetData(options, exportData) {
    const rows = options.rows || [];
    const headers = options.headers || [];
    const mappings = options.mappings || {};
    const isEn = options.uiLang === 'en';
    const idCol = headers[0];
    const sectionCol = headers[1];
    const labels = isEn
      ? { verified: 'Verified', review: 'Awaiting review', missing: 'Translation missing', removed: 'Removed in 2026',
          id: 'ID', section: 'Section', change: 'Reference change',
          UNCHANGED: 'Unchanged', MODIFIED: 'Modified', ADDED: 'Added', REMOVED: 'Removed', EMPTY: 'Empty' }
      : { verified: 'Ověřeno', review: 'Čeká na kontrolu', missing: 'Chybí překlad', removed: 'Zrušeno v 2026',
          id: 'ID', section: 'Oblast', change: 'Změna reference',
          UNCHANGED: 'Beze změny', MODIFIED: 'Upraveno', ADDED: 'Přidáno', REMOVED: 'Odstraněno', EMPTY: 'Prázdné' };

    const cols = Object.keys(exportData.newColNames);
    const statusHeaders = [labels.id, labels.section, labels.change].concat(cols.map(function (col) {
      return exportData.newColNames[col];
    }));
    const data = rows.map(function (row, idx) {
      const classification = classifyRow(row[mappings.refOldCol], row[mappings.refNewCol]);
      const out = {};
      out[labels.id] = idCol !== undefined ? cellValue(row[idCol]) : idx + 1;
      out[labels.section] = sectionCol !== undefined ? cellValue(row[sectionCol]) : '';
      out[labels.change] = labels[classification] || classification;
      cols.forEach(function (col) {
        out[exportData.newColNames[col]] = labels[getStubState(exportData.stubsByCol[col][idx])];
      });
      return out;
    });

    return {
      sheetName: isEn ? 'Translation_Status' : 'Stav_prekladu',
      headers: statusHeaders,
      data: data
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
      return { wch: Math.min(Math.max(maxLen + 3, 14), 63) };
    });
    if (ws['!ref']) ws['!autofilter'] = { ref: ws['!ref'] };
    return ws;
  }

  /**
   * Builds the export workbook with SheetJS (passed in so this works in Node and the browser).
   * The comprehensive workbook gets a second sheet with the verification status of every row.
   */
  function buildExportWorkbook(XLSX, options) {
    const exportData = prepareExportData(options);
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
   * Build a consolidated batch prompt for multiple modified/added articles.
   */
  function buildBatchCopilotPrompt(params) {
    const lang = params.promptLang || 'cs';
    const targetLang = params.targetLangName || (lang === 'en' ? 'target language' : 'cílového jazyka');
    let items = params.items;
    if (!items || items.length === 0) {
      items = [];
      const rows = params.rows || [];
      const headers = params.headers || [];
      const mappings = params.mappings || {};
      const translations = params.translations || {};
      const tLang = params.targetLangName;
      const stubs = (translations && tLang && translations[tLang]) ? translations[tLang] : [];
      const idCol = headers[0] || 'Artikel-Nr.';
      const sectionCol = headers[1] || 'Bereich';

      rows.forEach(function (row, idx) {
        const stub = stubs[idx];
        const refOld = row[mappings.refOldCol] !== undefined ? String(row[mappings.refOldCol]) : '';
        const refNew = row[mappings.refNewCol] !== undefined ? String(row[mappings.refNewCol]) : '';
        const classification = stub ? stub.classification : classifyRow(refOld, refNew);

        // Include MODIFIED and ADDED articles
        if (classification === 'MODIFIED' || classification === 'ADDED') {
          const rowId = row[idCol] || ('Art. ' + (idx + 1));
          const section = row[sectionCol] || '';
          items.push({
            index: idx,
            rowId: rowId + (section ? ' (' + section + ')' : ''),
            classification: classification,
            refOld: refOld,
            refNew: refNew,
            targetOld: stub ? stub.targetOld : (row[tLang] || ''),
            diffResult: computeDiff(refOld, refNew)
          });
        }
      });
    }

    if (items.length === 0) {
      return lang === 'en' 
        ? 'No modified or added articles found to translate.'
        : 'Nebyly nalezeny žádné změněné ani nově přidané odstavce k překladu.';
    }

    if (lang === 'en') {
      const header = [
        '=== BATCH TRANSLATION & UPDATE – CODE OF ETHICS 2026 ===',
        'You are an expert legal and corporate translator specializing in Codes of Ethics and Compliance.',
        'TASK: Provide updated/new translations for the following ' + items.length + ' articles in language: ' + targetLang + '.',
        '',
        '=== CRITICAL INSTRUCTIONS ===',
        '1. For EACH article, output the exact tag [IDENTIFIER] on its own line, followed immediately by the pure translated text.',
        '2. For MODIFIED articles, make ONLY minimal necessary alterations strictly reflecting the German diff. Keep unchanged phrases identical to the 2024 translation.',
        '3. For ADDED articles, translate the new German text faithfully.',
        '4. Do NOT output any intro, commentary, or conversational filler. Only output the tagged articles.',
        ''
      ].join('\n');

      const body = items.map(function (it, idx) {
        const idTag = it.rowId || ('Art. ' + (idx + 1));
        const diffRes = it.diffResult || computeDiff(it.refOld, it.refNew);
        let changesSummary = '';
        if (diffRes && diffRes.summary) {
          const added = diffRes.summary.addedTerms.filter(function(t){ return t.length > 1; });
          const removed = diffRes.summary.removedTerms.filter(function(t){ return t.length > 1; });
          if (added.length > 0) changesSummary += ' [Added in German: ' + added.slice(0, 8).join(', ') + ']';
          if (removed.length > 0) changesSummary += ' [Removed in German: ' + removed.slice(0, 8).join(', ') + ']';
        }

        if (it.classification === 'ADDED') {
          return [
            '----------------------------------------',
            'ARTICLE: [' + idTag + '] (Type: NEWLY ADDED)',
            'NEW GERMAN TEXT (2026):',
            cleanText(it.refNew),
            ''
          ].join('\n');
        }

        return [
          '----------------------------------------',
          'ARTICLE: [' + idTag + '] (Type: MODIFIED' + changesSummary + ')',
          'ORIGINAL GERMAN (2024):',
          cleanText(it.refOld),
          'UPDATED GERMAN (2026):',
          cleanText(it.refNew),
          'APPROVED LEGACY TRANSLATION (' + targetLang + ' 2024):',
          cleanText(it.targetOld),
          ''
        ].join('\n');
      }).join('\n');

      return header + '\n' + body;
    }

    // Czech batch prompt
    const header = [
      '=== DÁVKOVÝ PŘEKLAD A AKTUALIZACE ETICKÉHO KODEXU 2026 ===',
      'Jste špičkový odborný překladatel firemních a právních dokumentů (Etický kodex a Compliance).',
      'ÚKOL: Připravte aktualizované / nové překlady pro následujících ' + items.length + ' odstavců v jazyce: ' + targetLang + '.',
      '',
      '=== PŘÍSNÉ POKYNY PRO FORMÁT ODPOVĚDI ===',
      '1. Pro KAŽDÝ článek uveďte na samostatném řádku přesnou značku ve tvaru [IDENTIFIKÁTOR] a pod ni čistý přeložený text.',
      '2. U položek typu UPRAVENO proveďte POUZE minimální nutné změny odpovídající německému rozdílu. Všechny nezměněné věty a zavedenou terminologii z roku 2024 ponechte beze změny.',
      '3. U položek typu PŘIDÁNO přeložte novou německou verzi do spisovného a právně přesného jazyka.',
      '4. Nevypisujte žádné úvody, komentáře ani zdvořilostní fráze. Pouze označené odstavce.',
      ''
    ].join('\n');

    const body = items.map(function (it, idx) {
      const idTag = it.rowId || ('Art. ' + (idx + 1));
      const diffRes = it.diffResult || computeDiff(it.refOld, it.refNew);
      let changesSummary = '';
      if (diffRes && diffRes.summary) {
        const added = diffRes.summary.addedTerms.filter(function(t){ return t.length > 1; });
        const removed = diffRes.summary.removedTerms.filter(function(t){ return t.length > 1; });
        if (added.length > 0) changesSummary += ' [Přidáno v němčině: ' + added.slice(0, 8).join(', ') + ']';
        if (removed.length > 0) changesSummary += ' [Odebráno z němčiny: ' + removed.slice(0, 8).join(', ') + ']';
      }

      if (it.classification === 'ADDED') {
        return [
          '----------------------------------------',
          'ODSTAVEC: [' + idTag + '] (Typ: NOVĚ PŘIDÁNO)',
          'NOVÝ NĚMECKÝ TEXT (2026):',
          cleanText(it.refNew),
          ''
        ].join('\n');
      }

      return [
        '----------------------------------------',
        'ODSTAVEC: [' + idTag + '] (Typ: UPRAVENO' + changesSummary + ')',
        'PŮVODNÍ NĚMČINA (2024):',
        cleanText(it.refOld),
        'NOVÁ NĚMČINA (2026):',
        cleanText(it.refNew),
        'SCHVÁLENÝ PŘEKLAD (' + targetLang + ' 2024):',
        cleanText(it.targetOld),
        ''
      ].join('\n');
    }).join('\n');

    return header + '\n' + body;
  }

  /**
   * Helper to check if two article identifiers match (e.g. 'Art. 04' and '04')
   */
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
   * Generates a comprehensive Managerial Change Summary (2024 -> 2026).
   * Returns KPIs, detailed item lists, clean Markdown, and Word-ready HTML document.
   */
  function generateManagerialChangeSummary(params) {
    const rows = params.rows || [];
    const headers = params.headers || [];
    const mappings = params.mappings || {};
    const uiLang = params.uiLang || 'cs';
    const companyName = params.companyName || 'Corporate Group';

    const refOldCol = mappings.refOldCol;
    const refNewCol = mappings.refNewCol;
    const idCol = headers[0] || 'Artikel-Nr.';
    const sectionCol = headers[1] || 'Bereich';

    let total = 0;
    let unchanged = 0;
    let modified = 0;
    let added = 0;
    let removed = 0;

    const modifiedItems = [];
    const addedItems = [];
    const removedItems = [];

    rows.forEach(function (row, idx) {
      const refOld = row[refOldCol] !== undefined ? String(row[refOldCol]) : '';
      const refNew = row[refNewCol] !== undefined ? String(row[refNewCol]) : '';
      const classification = classifyRow(refOld, refNew);
      const rowId = row[idCol] || ('Art. ' + (idx + 1));
      const section = row[sectionCol] || '';

      if (classification === 'EMPTY') return;
      total++;

      if (classification === 'UNCHANGED') {
        unchanged++;
      } else if (classification === 'MODIFIED') {
        modified++;
        const diffRes = computeDiff(refOld, refNew);
        modifiedItems.push({
          index: idx + 1,
          id: rowId,
          section: section,
          refOld: refOld,
          refNew: refNew,
          diffResult: diffRes,
          addedTerms: diffRes.summary.addedTerms.filter(function(t){ return t.length > 1; }),
          removedTerms: diffRes.summary.removedTerms.filter(function(t){ return t.length > 1; }),
          similarity: diffRes.stats.similarity
        });
      } else if (classification === 'ADDED') {
        added++;
        addedItems.push({
          index: idx + 1,
          id: rowId,
          section: section,
          refNew: refNew
        });
      } else if (classification === 'REMOVED') {
        removed++;
        removedItems.push({
          index: idx + 1,
          id: rowId,
          section: section,
          refOld: refOld
        });
      }
    });

    const pctChanged = total > 0 ? Math.round(((modified + added + removed) / total) * 100) : 0;
    const isEn = uiLang === 'en';

    // 1. Build Markdown Document
    const mdLines = [];
    mdLines.push('# ' + (isEn ? 'Code of Conduct & Ethics – Management Change Summary (2024 → 2026)' : 'Etický kodex – Manažerské shrnutí změn (2024 → 2026)'));
    mdLines.push('');
    mdLines.push('**' + (isEn ? 'Organization' : 'Organizace') + ':** ' + companyName);
    mdLines.push('**' + (isEn ? 'Date' : 'Datum') + ':** ' + new Date().toISOString().split('T')[0]);
    mdLines.push('**' + (isEn ? 'Reference Versions' : 'Referenční verze') + ':** ' + refOldCol + ' → ' + refNewCol);
    mdLines.push('');
    mdLines.push('## ' + (isEn ? '1. Executive KPI Summary' : '1. Manažerské shrnutí a statistika změn'));
    mdLines.push('');
    mdLines.push('| ' + (isEn ? 'Metric' : 'Metrika') + ' | ' + (isEn ? 'Count' : 'Počet') + ' | ' + (isEn ? 'Share' : 'Podíl') + ' |');
    mdLines.push('| :--- | :--- | :--- |');
    mdLines.push('| **' + (isEn ? 'Total Evaluated Articles' : 'Celkem hodnocených odstavců') + '** | ' + total + ' | 100% |');
    mdLines.push('| **' + (isEn ? 'Unchanged (Valid & Certified)' : 'Beze změny (Platné a převzaté)') + '** | ' + unchanged + ' | ' + (total > 0 ? Math.round((unchanged/total)*100) : 0) + '% |');
    mdLines.push('| **' + (isEn ? 'Modified Articles' : 'Upravené odstavce') + '** | ' + modified + ' | ' + (total > 0 ? Math.round((modified/total)*100) : 0) + '% |');
    mdLines.push('| **' + (isEn ? 'Newly Added Articles' : 'Nově zavedené odstavce') + '** | ' + added + ' | ' + (total > 0 ? Math.round((added/total)*100) : 0) + '% |');
    mdLines.push('| **' + (isEn ? 'Removed / Repealed Articles' : 'Zrušené / Odstraněné odstavce') + '** | ' + removed + ' | ' + (total > 0 ? Math.round((removed/total)*100) : 0) + '% |');
    mdLines.push('');
    mdLines.push('> **' + (isEn ? 'Overall Revision Scope' : 'Celkový rozsah revize') + ':** ' + pctChanged + '% ' + (isEn ? 'of articles were subject to textual modification or restructuring.' : 'článků prošlo věcnou či formální úpravou.'));
    mdLines.push('');

    if (modifiedItems.length > 0) {
      mdLines.push('## ' + (isEn ? '2. Modified Articles (Detailed Analysis)' : '2. Detailně upravené články'));
      mdLines.push('');
      modifiedItems.forEach(function (m) {
        mdLines.push('### ' + m.id + (m.section ? ' – ' + m.section : ''));
        mdLines.push('- **' + (isEn ? 'Similarity' : 'Míra shody') + ':** ' + Math.round(m.similarity * 100) + '%');
        if (m.addedTerms.length > 0) {
          mdLines.push('- **' + (isEn ? 'Key Added Elements (2026)' : 'Klíčové přidané prvky (2026)') + ':** `' + m.addedTerms.join('`, `') + '`');
        }
        if (m.removedTerms.length > 0) {
          mdLines.push('- **' + (isEn ? 'Key Removed Elements (2024)' : 'Klíčové odstraněné prvky (2024)') + ':** `' + m.removedTerms.join('`, `') + '`');
        }
        mdLines.push('');
        mdLines.push('**' + (isEn ? 'German Source 2024' : 'Německý text 2024') + ':**');
        mdLines.push(mdQuote(m.refOld));
        mdLines.push('');
        mdLines.push('**' + (isEn ? 'German Source 2026' : 'Německý text 2026') + ':**');
        mdLines.push(mdQuote(m.refNew));
        mdLines.push('');
      });
    }

    if (addedItems.length > 0) {
      mdLines.push('## ' + (isEn ? '3. Newly Introduced Articles (2026)' : '3. Nově zavedené články (2026)'));
      mdLines.push('');
      addedItems.forEach(function (a) {
        mdLines.push('### ➕ ' + a.id + (a.section ? ' – ' + a.section : ''));
        mdLines.push('**' + (isEn ? 'New Text (2026)' : 'Nové znění (2026)') + ':**');
        mdLines.push(mdQuote(a.refNew));
        mdLines.push('');
      });
    }

    if (removedItems.length > 0) {
      mdLines.push('## ' + (isEn ? '4. Removed / Retired Articles' : '4. Zrušené / Odstraněné články'));
      mdLines.push('');
      removedItems.forEach(function (r) {
        mdLines.push('### 🗑️ ' + r.id + (r.section ? ' – ' + r.section : ''));
        mdLines.push('**' + (isEn ? 'Archived Text (2024)' : 'Archivované znění (2024)') + ':**');
        mdLines.push(mdQuote(r.refOld));
        mdLines.push('');
      });
    }

    const markdownDoc = mdLines.join('\n');

    // 2. Build Word-ready HTML document
    const htmlWord = `<!DOCTYPE html>
<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns="http://www.w3.org/TR/REC-html40">
<head>
  <meta charset="UTF-8">
  <title>${escapeHtml(isEn ? 'Management Change Summary' : 'Manažerské shrnutí změn')} - ${escapeHtml(companyName)}</title>
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
    blockquote { border-left: 3pt solid #94a3b8; margin: 4pt 0 10pt 0; padding-left: 10pt; color: #334155; font-style: italic; background: #f8fafc; }
    ins { background-color: #dcfce7; color: #15803d; text-decoration: none; font-weight: bold; }
    del { background-color: #fee2e2; color: #b91c1c; text-decoration: line-through; }
  </style>
</head>
<body>
  <h1>${escapeHtml(isEn ? 'Code of Conduct & Ethics – Management Change Summary (2024 → 2026)' : 'Etický kodex – Manažerské shrnutí změn (2024 → 2026)')}</h1>
  <p>
    <strong>${escapeHtml(isEn ? 'Organization' : 'Organizace')}:</strong> ${escapeHtml(companyName)} | 
    <strong>${escapeHtml(isEn ? 'Date' : 'Datum')}:</strong> ${new Date().toLocaleDateString()}<br>
    <strong>${escapeHtml(isEn ? 'Reference Columns' : 'Referenční sloupce')}:</strong> ${escapeHtml(refOldCol)} &rarr; ${escapeHtml(refNewCol)}
  </p>

  <h2>${escapeHtml(isEn ? '1. Executive KPI Summary' : '1. Manažerské shrnutí a statistika změn')}</h2>
  <table class="kpi-table">
    <tr>
      <th>${escapeHtml(isEn ? 'Revision Category' : 'Kategorie revize')}</th>
      <th>${escapeHtml(isEn ? 'Article Count' : 'Počet odstavců')}</th>
      <th>${escapeHtml(isEn ? 'Percentage' : 'Procentuální podíl')}</th>
    </tr>
    <tr>
      <td><strong>${escapeHtml(isEn ? 'Total Evaluated Articles' : 'Celkem hodnocených odstavců')}</strong></td>
      <td><strong>${total}</strong></td>
      <td><strong>100%</strong></td>
    </tr>
    <tr>
      <td>${escapeHtml(isEn ? 'Unchanged (Directly Certified & Adopted)' : 'Beze změny (Platné a přímo převzaté)')}</td>
      <td>${unchanged}</td>
      <td>${total > 0 ? Math.round((unchanged/total)*100) : 0}%</td>
    </tr>
    <tr>
      <td><span class="badge badge-mod">${escapeHtml(isEn ? 'Modified Articles' : 'Upravené odstavce')}</span></td>
      <td>${modified}</td>
      <td>${total > 0 ? Math.round((modified/total)*100) : 0}%</td>
    </tr>
    <tr>
      <td><span class="badge badge-add">${escapeHtml(isEn ? 'Newly Added Articles' : 'Nově přidané odstavce')}</span></td>
      <td>${added}</td>
      <td>${total > 0 ? Math.round((added/total)*100) : 0}%</td>
    </tr>
    <tr>
      <td><span class="badge badge-rem">${escapeHtml(isEn ? 'Removed / Retired Articles' : 'Zrušené / Odstraněné odstavce')}</span></td>
      <td>${removed}</td>
      <td>${total > 0 ? Math.round((removed/total)*100) : 0}%</td>
    </tr>
  </table>

  ${modifiedItems.length > 0 ? `<h2>${escapeHtml(isEn ? '2. Modified Articles Detail' : '2. Detailně upravené články')}</h2>` : ''}
  ${modifiedItems.map(function(m){
    const diffHtml = renderDiffHtml(m.diffResult, 'inline');
    return `
      <h3>${escapeHtml(m.id)} ${m.section ? '– ' + escapeHtml(m.section) : ''} <span class="badge badge-mod">${escapeHtml(isEn ? 'Modified' : 'Upraveno')}</span></h3>
      <table>
        <tr>
          <th style="width:50%;">${escapeHtml(isEn ? 'Original German (2024)' : 'Původní němčina (2024)')}</th>
          <th style="width:50%;">${escapeHtml(isEn ? 'Updated German (2026) with Diff' : 'Nová němčina (2026) s vyznačeným rozdílem')}</th>
        </tr>
        <tr>
          <td>${escapeHtmlMultiline(m.refOld)}</td>
          <td>${diffHtml}</td>
        </tr>
      </table>
    `;
  }).join('')}

  ${addedItems.length > 0 ? `
    <h2>${escapeHtml(isEn ? '3. Newly Introduced Articles' : '3. Nově zavedené články')}</h2>
    ${addedItems.map(function(a){
      return `
        <h3>➕ ${escapeHtml(a.id)} ${a.section ? '– ' + escapeHtml(a.section) : ''} <span class="badge badge-add">${escapeHtml(isEn ? 'Added' : 'Přidáno')}</span></h3>
        <blockquote>${escapeHtmlMultiline(a.refNew)}</blockquote>
      `;
    }).join('')}
  ` : ''}

  ${removedItems.length > 0 ? `
    <h2>${escapeHtml(isEn ? '4. Removed / Retired Articles' : '4. Zrušené / Odstraněné články')}</h2>
    ${removedItems.map(function(r){
      return `
        <h3>🗑️ ${escapeHtml(r.id)} ${r.section ? '– ' + escapeHtml(r.section) : ''} <span class="badge badge-rem">${escapeHtml(isEn ? 'Removed' : 'Odstraněno')}</span></h3>
        <blockquote>${escapeHtmlMultiline(r.refOld)}</blockquote>
      `;
    }).join('')}
  ` : ''}
</body>
</html>`;

    return {
      stats: {
        totalRows: total,
        unchangedCount: unchanged,
        modifiedCount: modified,
        addedCount: added,
        removedCount: removed,
        pctChanged: pctChanged
      },
      kpis: {
        total: total,
        unchanged: unchanged,
        modified: modified,
        added: added,
        removed: removed,
        pctChanged: pctChanged
      },
      modifiedItems: modifiedItems,
      addedItems: addedItems,
      removedItems: removedItems,
      markdown: markdownDoc,
      htmlDoc: htmlWord
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
    generateTargetStub: generateTargetStub,
    rehydrateStubs: rehydrateStubs,
    resolveStubs: resolveStubs,
    getStubState: getStubState,
    idsMatch: idsMatch,
    format2026ColName: format2026ColName,
    buildCopilotPrompt: buildCopilotPrompt,
    buildBatchCopilotPrompt: buildBatchCopilotPrompt,
    parseBatchTranslationResponse: parseBatchTranslationResponse,
    generateManagerialChangeSummary: generateManagerialChangeSummary,
    prepareExportData: prepareExportData,
    buildExportWorkbook: buildExportWorkbook
  };
}));
