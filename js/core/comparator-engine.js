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
   * Helper to check if a language was added newly and has no old version column.
   */
  function isNewLanguage(mappings, col) {
    if (!col) return false;
    const languages = (mappings && mappings.languages) || {};
    const info = languages[col];
    if (info && info.isNewLanguage) return true;
    if (typeof col === 'string' && col.indexOf('__new__') === 0) return true;
    return false;
  }

  /**
   * Settings of one language (keyed by its old-version column).
   * newCol: an existing column with the new version ('' = a new column is created on export).
   */
  function getLanguageInfo(mappings, col) {
    const languages = (mappings && mappings.languages) || {};
    const info = languages[col] || {};
    const isNew = isNewLanguage(mappings, col);
    return {
      col: col,
      label: info.label || (isNew ? (info.newColName || String(col).replace(/^__new__/, '')) : col),
      newCol: info.newCol || '',
      newColName: info.newColName || defaultNewColName(col, mappings && mappings.refOldCol, mappings && mappings.refNewCol),
      isNewLanguage: isNew
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
   * @param {string} targetCol - Column with the old version of the language (or ID of new language)
   * @param {string} refOldCol - Reference column, old version
   * @param {string} refNewCol - Reference column, new version
   * @param {string} [newCol] - Existing column that already holds (part of) the new version
   * @param {boolean} [isNewLang] - True if this is a newly created language with no old version
   * @returns {Array<Object>} Stubs for each row
   */
  function generateTargetStub(rows, targetCol, refOldCol, refNewCol, newCol, isNewLang) {
    if (!rows || !Array.isArray(rows)) return [];

    const isPureNew = isNewLang !== undefined ? !!isNewLang : (typeof targetCol === 'string' && targetCol.indexOf('__new__') === 0);

    return rows.map(function (row, idx) {
      const refOld = row[refOldCol];
      const refNew = row[refNewCol];
      const targetOld = (!isPureNew && row[targetCol] !== undefined && row[targetCol] !== null) ? String(row[targetCol]) : '';
      const existingNew = newCol && row[newCol] !== undefined && row[newCol] !== null ? String(row[newCol]) : '';
      const classification = classifyRow(refOld, refNew);

      let initialNewText = '';
      let status = 'pending';
      let isVerified = false;

      if (isPureNew) {
        // Purely new languages have no old translation to carry over
        switch (classification) {
          case 'REMOVED':
          case 'EMPTY':
            initialNewText = '';
            status = 'removed';
            isVerified = true;
            break;

          case 'UNCHANGED':
          case 'MODIFIED':
          case 'ADDED':
          default:
            initialNewText = '';
            status = 'pending_translation';
            isVerified = false;
            break;
        }
      } else {
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
      }

      // Text already present in the new-version column of the sheet wins, but has to be reviewed
      // unless it is simply the unchanged old translation
      if (existingNew.trim()) {
        const sameAsAuto = !isPureNew && classification === 'UNCHANGED' && existingNew === targetOld;
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
        isNewLanguage: isPureNew,
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

    const auxText = cleanText(params.auxiliaryText);
    const auxLang = params.auxiliaryLangName || '';

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
        CLASSIFICATION: classification,
        AUX_LANG: auxLang,
        AUX_TEXT: auxText
      };
      return params.customTemplate.replace(/\{([A-Z_]+)\}/g, function (all, key) {
        return Object.prototype.hasOwnProperty.call(values, key) ? values[key] : all;
      }).trim();
    }

    if (lang === 'en') {
      if (params.isNewLanguage) {
        const lines = [
          'You are an experienced professional translator.',
          'TASK: Translate the text into language: ' + targetLang + '.',
          'Paragraph: ' + rowId,
          '',
          'INSTRUCTIONS:',
          '1. Translate faithfully, maintaining a natural style and professional terminology in ' + targetLang + '.',
          '2. Output ONLY the translated text (no introduction, comments, quotes or markdown).',
          '',
          '=== REFERENCE TEXT (' + refLang + ') ===',
          refNew
        ];
        if (auxLang && auxText) {
          lines.push('', '=== AUXILIARY TRANSLATION FOR REFERENCE (' + auxLang + ') ===', auxText);
        }
        return lines.join('\n');
      }

      if (classification === 'ADDED') {
        const lines = [
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
        ];
        if (auxLang && auxText) {
          lines.push('', '=== AUXILIARY TRANSLATION FOR REFERENCE (' + auxLang + ') ===', auxText);
        }
        return lines.join('\n');
      }

      const lines = [
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
      ];
      if (auxLang && auxText) {
        lines.push('', '=== 4. AUXILIARY TRANSLATION FOR REFERENCE (' + auxLang + ') ===', auxText);
      }
      return lines.join('\n');
    }

    // Czech default
    if (params.isNewLanguage) {
      const lines = [
        'Jste zkušený odborný překladatel.',
        'ÚKOL: Přeložte text do jazyka: ' + targetLang + '.',
        'Odstavec: ' + rowId,
        '',
        'POKYNY:',
        '1. Přeložte text věrně a dodržte přirozený styl a odbornou terminologii v jazyce ' + targetLang + '.',
        '2. POUZE vraťte přeložený text – bez úvodu, komentářů, uvozovek a markdownu.',
        '',
        '=== REFERENČNÍ TEXT (' + refLang + ') ===',
        refNew
      ];
      if (auxLang && auxText) {
        lines.push('', '=== POMOCNÝ PŘEKLAD PRO POROVNÁNÍ (' + auxLang + ') ===', auxText);
      }
      return lines.join('\n');
    }

    if (classification === 'ADDED') {
      const lines = [
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
      ];
      if (auxLang && auxText) {
        lines.push('', '=== POMOCNÝ PŘEKLAD PRO POROVNÁNÍ (' + auxLang + ') ===', auxText);
      }
      return lines.join('\n');
    }

    const lines = [
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
    ];
    if (auxLang && auxText) {
      lines.push('', '=== 4. POMOCNÝ PŘEKLAD PRO POROVNÁNÍ (' + auxLang + ') ===', auxText);
    }
    return lines.join('\n');
  }

  /**
   * Generates a Copilot prompt for auditing and verifying translation fidelity and meaning equivalence.
   *
   * @param {Object} params
   * @param {string} [params.promptLang='cs'] - 'cs' | 'en'
   * @param {string} [params.targetLangName] - Name of the target language
   * @param {string} [params.refLangName] - Name of the reference language
   * @param {string} [params.auxiliaryLangName] - Name of the auxiliary language if available
   * @param {string} [params.auxiliaryText] - Auxiliary reference translation text
   * @param {string} [params.rowId] - Formatted row identifier
   * @param {string} [params.refNew] - New reference text
   * @param {string} [params.targetNew] - New translation text under review
   * @param {string} [params.customTemplate] - Optional custom template
   * @returns {string} Verification prompt
   */
  function buildVerificationPrompt(params) {
    if (!params) return '';
    const lang = params.promptLang || 'cs';
    const targetLang = params.targetLangName || (lang === 'en' ? 'target language' : 'cílový jazyk');
    const refLang = params.refLangName || (lang === 'en' ? 'reference' : 'reference');
    const rowId = params.rowId || ((lang === 'en' ? 'Row ' : 'Řádek ') + (params.rowIndex !== undefined ? params.rowIndex + 2 : ''));
    const refNew = cleanText(params.refNew);
    const targetNew = cleanText(params.targetNew);
    const auxText = cleanText(params.auxiliaryText);
    const auxLang = params.auxiliaryLangName || '';

    if (params.customTemplate && typeof params.customTemplate === 'string' && params.customTemplate.trim()) {
      const values = {
        TARGET_LANG: targetLang,
        REF_LANG: refLang,
        ROW_ID: rowId,
        REF_NEW: refNew,
        TARGET_NEW: targetNew,
        AUX_LANG: auxLang,
        AUX_TEXT: auxText
      };
      return params.customTemplate.replace(/\{([A-Z_]+)\}/g, function (all, key) {
        return Object.prototype.hasOwnProperty.call(values, key) ? values[key] : all;
      }).trim();
    }

    if (lang === 'en') {
      const lines = [
        'You are an experienced translation proofreader and language quality auditor.',
        'TASK: Verify the semantic equivalence and accuracy of the new translation in ' + targetLang + ' against the reference text (' + refLang + ').',
        'Paragraph: ' + rowId,
        '',
        '=== 1. REFERENCE TEXT (' + refLang + ') ===',
        refNew,
        '',
        '=== 2. NEW TRANSLATION UNDER REVIEW (' + targetLang + ') ===',
        targetNew
      ];

      if (auxLang && auxText) {
        lines.push('', '=== 3. AUXILIARY REFERENCE TRANSLATION (' + auxLang + ') ===', auxText);
      }

      lines.push(
        '',
        '=== EVALUATION INSTRUCTIONS ===',
        '1. Check whether the translation accurately and completely reflects the meaning of the reference text without omissions, unwanted additions or meaning distortion.',
        '2. Verify consistent terminology, grammar, professional tone and natural phrasing in ' + targetLang + '.',
        '3. Structure your response clearly as follows:',
        '   - VERDICT: [ACCURATE / MINOR REVISIONS RECOMMENDED / SIGNIFICANT DISCREPANCIES]',
        '   - DISCREPANCIES & DEFECTS: (Concise list of specific nuances missed, shifts in meaning or inaccuracies; if none, state "None")',
        '   - SUGGESTED REVISION: (If any flaws or improvements are identified, provide the recommended complete revised paragraph)'
      );
      return lines.join('\n');
    }

    // Czech default
    const lines = [
      'Jste zkušený jazykový korektor a revizor překladů.',
      'ÚKOL: Ověřte věcnou a významovou shodnost nového překladu v jazyce ' + targetLang + ' vůči referenčnímu textu (' + refLang + ').',
      'Odstavec: ' + rowId,
      '',
      '=== 1. REFERENČNÍ TEXT (' + refLang + ') ===',
      refNew,
      '',
      '=== 2. KONTROLOVANÝ NOVÝ PŘEKLAD (' + targetLang + ') ===',
      targetNew
    ];

    if (auxLang && auxText) {
      lines.push('', '=== 3. POMOCNÝ REFERENČNÍ PŘEKLAD (' + auxLang + ') ===', auxText);
    }

    lines.push(
      '',
      '=== POKYNY K VYHODNOCENÍ ===',
      '1. Posuďte, zda kontrolovaný překlad přesně a úplně zachovává význam referenčního textu (žádné opomenuté informace, nepatřičné přídavky ani významové posuny).',
      '2. Zkontrolujte odbornou terminologii, gramatiku, styl a přirozenost formulace v jazyce ' + targetLang + '.',
      '3. Odpověď strukturujte přehledně do těchto bodů:',
      '   - ZÁVĚR: [PŘEKLAD JE V POŘÁDKU / DOPORUČENY ÚPRAVY / ZÁVAŽNÉ ROZDÍLY]',
      '   - ROZDÍLY A NEDOSTATKY: (Stručný výčet konkrétních odchylek, chybějících nuancí či nepřesností; pokud žádné nejsou, uveďte „Žádné“)',
      '   - NÁVRH ÚPRAVY: (Pokud byly zjištěny nedostatky, uveďte doporučené opravené znění celého odstavce)'
    );
    return lines.join('\n');
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
  function rehydrateStubs(rows, savedStubs, targetCol, refOldCol, refNewCol, newCol, isNewLang) {
    const isPureNew = isNewLang !== undefined ? !!isNewLang : (typeof targetCol === 'string' && targetCol.indexOf('__new__') === 0);
    const fresh = generateTargetStub(rows, targetCol, refOldCol, refNewCol, newCol, isPureNew);
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
    if (Array.isArray(existing) && existing.length === rows.length &&
      existing.every(function (s) { return s && s.targetCol !== undefined; })) {
      return existing;
    }
    const info = getLanguageInfo(mappings, targetCol);
    return rehydrateStubs(rows, existing, targetCol, mappings.refOldCol, mappings.refNewCol,
      info.newCol, info.isNewLanguage);
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
          skipped: 'Skipped', comprehensive: 'Comparison', clean: 'New_version', statusSheet: 'Translation_Status' }
      : { row: 'Řádek', change: 'Změna reference', status: 'Stav',
          verified: 'Ověřeno', review: 'Čeká na kontrolu', missing: 'Chybí překlad', removed: 'Zrušeno v nové verzi',
          UNCHANGED: 'Beze změny', MODIFIED: 'Upraveno', ADDED: 'Přidáno', REMOVED: 'Odstraněno', EMPTY: 'Prázdné',
          skipped: 'Přeskočeno', comprehensive: 'Srovnani', clean: 'Nova_verze', statusSheet: 'Stav_prekladu' };
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
    const targetCols = (mappings.targetCols || []).filter(function (col) {
      return headers.indexOf(col) !== -1 || isNewLanguage(mappings, col);
    });

    const updatingCols = targetCols.filter(function (col) { return headers.indexOf(col) !== -1; });
    const newLangCols = targetCols.filter(function (col) { return isNewLanguage(mappings, col) && headers.indexOf(col) === -1; });

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
        if (h === mappings.refOldCol || updatingCols.indexOf(h) !== -1 || existingNewCols[h]) return;
        sourceCols.push({ from: h, to: uniqueColName(h, taken) });
      });
      const targetNames = targetCols.map(function (col) {
        return { col: col, to: uniqueColName(infos[col].newCol || infos[col].newColName || infos[col].label, taken) };
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
      newColNames[col] = infos[col].newCol || uniqueColName(infos[col].newColName || infos[col].label, taken);
    });

    const compHeaders = [];
    headers.forEach(function (h) {
      compHeaders.push(h);
      if (updatingCols.indexOf(h) !== -1 && !infos[h].newCol) compHeaders.push(newColNames[h]);
    });
    // Newly added languages get ONLY ONE column, placed at the end
    newLangCols.forEach(function (col) {
      compHeaders.push(newColNames[col]);
    });

    function processedRow(row, idx) {
      const newRow = {};
      headers.forEach(function (h) {
        newRow[h] = existingNewCols[h] ? translatedText(existingNewCols[h], idx) : cellValue(row[h]);
        if (updatingCols.indexOf(h) !== -1 && !infos[h].newCol) {
          newRow[newColNames[h]] = translatedText(h, idx);
        }
      });
      newLangCols.forEach(function (col) {
        newRow[newColNames[col]] = translatedText(col, idx);
      });
      return newRow;
    }

    // Rows skipped by the user stay in the full workbook unchanged (new columns left empty)
    function skippedRow(row) {
      const newRow = {};
      headers.forEach(function (h) {
        newRow[h] = cellValue(row[h]);
        if (updatingCols.indexOf(h) !== -1 && !infos[h].newCol) newRow[newColNames[h]] = '';
      });
      newLangCols.forEach(function (col) {
        newRow[newColNames[col]] = '';
      });
      return newRow;
    }

    const order = exportRowOrder(rawRows, options.rowNumbers, options.skippedRows);
    const comprehensiveRows = order.map(function (entry) {
      return entry.skipped ? skippedRow(entry.row) : processedRow(rawRows[entry.idx], entry.idx);
    });

    return {
      sheetName: labels.comprehensive,
      headers: compHeaders,
      data: comprehensiveRows,
      newColNames: newColNames,
      stubsByCol: stubsByCol,
      order: order
    };
  }

  /**
   * Processed rows and skipped rows ([{ rowNumber, row }]) in the order of the original sheet.
   * Returns [{ rowNumber, idx }] for processed rows and [{ rowNumber, row, skipped: true }].
   */
  function exportRowOrder(rows, rowNumbers, skippedRows) {
    const entries = rows.map(function (row, idx) {
      return { rowNumber: rowNumberOf(rowNumbers, idx), idx: idx };
    });
    (Array.isArray(skippedRows) ? skippedRows : []).forEach(function (s) {
      if (s && s.row) entries.push({ rowNumber: s.rowNumber, row: s.row, skipped: true });
    });
    // Stable sort by row number (processed rows keep their order for equal numbers)
    return entries
      .map(function (e, i) { return { e: e, i: i }; })
      .sort(function (a, b) { return (a.e.rowNumber - b.e.rowNumber) || (a.i - b.i); })
      .map(function (x) { return x.e; });
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

    const order = exportData.order || exportRowOrder(rows, options.rowNumbers, null);
    const data = order.map(function (entry) {
      const out = {};
      out[labels.row] = entry.rowNumber;
      if (entry.skipped) {
        out[labels.change] = labels.skipped;
        cols.forEach(function (col, k) { out[colTitles[k]] = labels.skipped; });
        return out;
      }
      const row = rows[entry.idx];
      const classification = classifyRow(row[mappings.refOldCol], row[mappings.refNewCol]);
      out[labels.change] = labels[classification] || classification;
      cols.forEach(function (col, k) {
        out[colTitles[k]] = labels[getStubState(exportData.stubsByCol[col][entry.idx])];
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
   * Partial workbook for one language: row number, chosen context columns, old and new version of
   * the language and its status. For the reference itself (langCol === refNewCol) the old and new
   * reference columns are exported.
   * contextCols entries: a column name of the sheet, or { newOf: <old-version column> } for the
   * new version (current translation work) of another language. A sheet column that serves as a
   * language's new-version column is exported with the current translation, not the stale cell.
   */
  function prepareLanguageExportData(options) {
    const rows = options.rows || [];
    const headers = options.headers || [];
    const mappings = options.mappings || {};
    const labels = sheetLabels(options.uiLang);
    const langCol = options.langCol || options.targetCol;
    const isReference = langCol === mappings.refNewCol;
    const isPureNew = !isReference && isNewLanguage(mappings, langCol);
    const info = isReference ? null : getLanguageInfo(mappings, langCol);
    const oldCol = isReference ? mappings.refOldCol : (isPureNew ? '' : langCol);
    const newColExisting = isReference ? mappings.refNewCol : (info.newCol && headers.indexOf(info.newCol) !== -1 ? info.newCol : '');
    const stubs = isReference ? null : resolveStubs(rows, options.translations || {}, langCol, mappings);

    const targetCols = (mappings.targetCols || []).filter(function (c) {
      return headers.indexOf(c) !== -1 || isNewLanguage(mappings, c);
    });
    const newVersionOwner = {};
    targetCols.forEach(function (c) {
      const inf = getLanguageInfo(mappings, c);
      if (inf.newCol && headers.indexOf(inf.newCol) !== -1) newVersionOwner[inf.newCol] = c;
    });

    const context = [];
    const seen = {};
    (options.contextCols || []).forEach(function (c) {
      let entry = null;
      if (c && typeof c === 'object' && typeof c.newOf === 'string') {
        if (c.newOf !== langCol && targetCols.indexOf(c.newOf) !== -1) entry = { newOf: c.newOf };
      } else if (typeof c === 'string' && headers.indexOf(c) !== -1 && c !== oldCol && c !== newColExisting) {
        entry = newVersionOwner[c] && newVersionOwner[c] !== langCol ? { newOf: newVersionOwner[c] } : { col: c };
      }
      if (!entry) return;
      const key = entry.newOf ? 'new:' + entry.newOf : 'col:' + entry.col;
      if (seen[key]) return;
      seen[key] = true;
      if (entry.newOf) {
        const inf = getLanguageInfo(mappings, entry.newOf);
        entry.name = inf.newCol && headers.indexOf(inf.newCol) !== -1 ? inf.newCol : (inf.newColName || inf.label);
        entry.stubs = resolveStubs(rows, options.translations || {}, entry.newOf, mappings);
      } else {
        entry.name = entry.col;
      }
      context.push(entry);
    });
    const taken = [labels.row];
    const contextNames = context.map(function (e) { return uniqueColName(e.name, taken); });
    const oldName = (!isPureNew && oldCol) ? uniqueColName(oldCol, taken) : '';
    const newName = uniqueColName(newColExisting || (info && (info.newColName || info.label)) || 'New', taken);
    const statusName = isReference ? null : uniqueColName(labels.status, taken);

    const partialHeaders = [labels.row].concat(contextNames);
    if (oldName) partialHeaders.push(oldName);
    partialHeaders.push(newName);
    if (statusName) partialHeaders.push(statusName);

    const data = rows.map(function (row, idx) {
      const out = {};
      out[labels.row] = rowNumberOf(options.rowNumbers, idx);
      context.forEach(function (e, k) {
        out[contextNames[k]] = e.stubs ? cellValue(e.stubs[idx] && e.stubs[idx].targetNew) : cellValue(row[e.col]);
      });
      if (oldName) out[oldName] = cellValue(row[oldCol]);
      out[newName] = isReference ? cellValue(row[mappings.refNewCol]) : cellValue(stubs[idx] && stubs[idx].targetNew);
      if (statusName) out[statusName] = labels[getStubState(stubs[idx])];
      return out;
    });

    const label = isReference ? String(langCol) : info.label;
    return {
      sheetName: String(label).replace(/[\\/?*[\]:]/g, '_').slice(0, 31) || 'Export',
      headers: partialHeaders,
      data: data,
      stubsByCol: stubs ? { [langCol]: stubs } : {}
    };
  }

  /**
   * Format-independent export of the work as tables (rendered by a 'table' export adapter).
   * exportType: 'comprehensive' (+ status table), 'clean', or 'language' (partial, one language).
   * Returns { tables: [{ name, headers, rows }], unverifiedCount, exportData }.
   */
  function prepareExportTables(options) {
    const exportData = options.exportType === 'language'
      ? prepareLanguageExportData(options)
      : prepareExportData(options);
    const tables = [{ name: exportData.sheetName, headers: exportData.headers, rows: exportData.data }];

    let unverifiedCount = 0;
    Object.keys(exportData.stubsByCol).forEach(function (col) {
      exportData.stubsByCol[col].forEach(function (s) {
        const st = getStubState(s);
        if (st === 'review' || st === 'missing') unverifiedCount++;
      });
    });

    if (exportData.newColNames) {
      const statusData = prepareStatusSheetData(options, exportData);
      tables.push({ name: statusData.sheetName, headers: statusData.headers, rows: statusData.data });
    }

    return { tables: tables, unverifiedCount: unverifiedCount, exportData: exportData };
  }

  /**
   * Build a consolidated batch prompt for all modified/added paragraphs of one language.
   */
  function buildBatchCopilotPrompt(params) {
    const lang = params.promptLang || 'cs';
    const targetLang = params.targetLangName || (lang === 'en' ? 'target language' : 'cílový jazyk');
    const refLang = params.refLangName || 'reference';
    const isNewLang = !!params.isNewLanguage;
    let items = params.items;
    if (!items || items.length === 0) {
      items = [];
      const rows = params.rows || [];
      const mappings = params.mappings || {};
      const translations = params.translations || {};
      const targetCol = params.targetCol || params.targetLangName;
      const stubs = (translations && targetCol && translations[targetCol]) ? translations[targetCol] : [];
      const auxCol = params.auxiliaryCol;
      const auxStubs = (translations && auxCol && translations[auxCol]) ? translations[auxCol] : [];

      rows.forEach(function (row, idx) {
        const stub = stubs[idx];
        const refOld = row[mappings.refOldCol] !== undefined ? String(row[mappings.refOldCol]) : '';
        const refNew = row[mappings.refNewCol] !== undefined ? String(row[mappings.refNewCol]) : '';
        const classification = stub ? stub.classification : classifyRow(refOld, refNew);

        const shouldInclude = isNewLang
          ? (classification !== 'REMOVED' && classification !== 'EMPTY' && refNew.trim() !== '')
          : (classification === 'MODIFIED' || classification === 'ADDED');

        if (shouldInclude) {
          const item = {
            index: idx,
            rowId: rowTag(params.rowNumbers, idx),
            classification: classification,
            refOld: refOld,
            refNew: refNew,
            targetOld: stub ? stub.targetOld : (row[targetCol] || ''),
            diffResult: computeDiff(refOld, refNew)
          };
          if (auxCol) {
            item.auxText = auxStubs[idx] ? (auxStubs[idx].targetNew || auxStubs[idx].targetOld) : (row[auxCol] || '');
          }
          items.push(item);
        }
      });
    }

    if (items.length === 0) {
      return lang === 'en'
        ? (isNewLang ? 'No paragraphs found to translate.' : 'No modified or added paragraphs found to translate.')
        : (isNewLang ? 'Nebyly nalezeny žádné odstavce k překladu.' : 'Nebyly nalezeny žádné změněné ani nově přidané odstavce k překladu.');
    }

    const auxLangName = params.auxiliaryLangName || '';

    const t = lang === 'en'
      ? {
        title: isNewLang ? '=== BATCH TRANSLATION ===' : '=== BATCH TRANSLATION UPDATE ===',
        role: 'You are an experienced professional translator.',
        task: isNewLang
          ? 'TASK: Provide translations of the following ' + items.length + ' paragraphs in language: ' + targetLang + '.'
          : 'TASK: Provide updated or new translations of the following ' + items.length + ' paragraphs in language: ' + targetLang + '.',
        rulesTitle: '=== CRITICAL INSTRUCTIONS ===',
        rules: isNewLang
          ? [
            '1. For EACH paragraph output its exact tag (e.g. [#12]) on its own line, followed immediately by the translated text.',
            '2. Translate faithfully, maintaining a natural style and professional terminology in ' + targetLang + '.',
            '3. Do NOT output any introduction, commentary or filler – only the tagged paragraphs.'
          ]
          : [
            '1. For EACH paragraph output its exact tag (e.g. [#12]) on its own line, followed immediately by the translated text.',
            '2. For MODIFIED paragraphs make ONLY the minimal changes reflecting the reference difference; keep unchanged phrases identical to the current translation.',
            '3. For ADDED paragraphs translate the new reference text faithfully.',
            '4. Do NOT output any introduction, commentary or filler – only the tagged paragraphs.'
          ],
        para: 'PARAGRAPH', added: 'NEWLY ADDED', modified: 'MODIFIED', typ: 'Type',
        newRef: 'NEW REFERENCE TEXT (' + refLang + '):', oldRef: 'OLD REFERENCE TEXT (' + refLang + '):',
        updRef: 'NEW REFERENCE TEXT:', approved: 'CURRENT APPROVED TRANSLATION (' + targetLang + '):',
        aux: 'AUXILIARY REFERENCE (' + auxLangName + '):'
      }
      : {
        title: isNewLang ? '=== HROMADNÝ PŘEKLAD ===' : '=== HROMADNÁ AKTUALIZACE PŘEKLADU ===',
        role: 'Jste zkušený odborný překladatel.',
        task: isNewLang
          ? 'ÚKOL: Připravte překlady následujících ' + items.length + ' odstavců v jazyce: ' + targetLang + '.'
          : 'ÚKOL: Připravte aktualizované nebo nové překlady následujících ' + items.length + ' odstavců v jazyce: ' + targetLang + '.',
        rulesTitle: '=== PŘÍSNÉ POKYNY PRO FORMÁT ODPOVĚDI ===',
        rules: isNewLang
          ? [
            '1. U KAŽDÉHO odstavce uveďte na samostatném řádku jeho přesnou značku (např. [#12]) a pod ni přeložený text.',
            '2. Přeložte text věrně a zachovejte přirozený styl a odbornou terminologii v jazyce ' + targetLang + '.',
            '3. Nevypisujte žádné úvody, komentáře ani zdvořilostní fráze – pouze označené odstavce.'
          ]
          : [
            '1. U KAŽDÉHO odstavce uveďte na samostatném řádku jeho přesnou značku (např. [#12]) a pod ni přeložený text.',
            '2. U odstavců typu UPRAVENO proveďte POUZE minimální změny odpovídající rozdílu v referenci; nezměněné formulace ponechte shodné s dosavadním překladem.',
            '3. U odstavců typu PŘIDÁNO přeložte nový referenční text věrně.',
            '4. Nevypisujte žádné úvody, komentáře ani zdvořilostní fráze – pouze označené odstavce.'
          ],
        para: 'ODSTAVEC', added: 'NOVĚ PŘIDÁNO', modified: 'UPRAVENO', typ: 'Typ',
        newRef: 'NOVÝ REFERENČNÍ TEXT (' + refLang + '):', oldRef: 'PŮVODNÍ REFERENČNÍ TEXT (' + refLang + '):',
        updRef: 'NOVÝ REFERENČNÍ TEXT:', approved: 'DOSAVADNÍ SCHVÁLENÝ PŘEKLAD (' + targetLang + '):',
        aux: 'POMOCNÝ PŘEKLAD (' + auxLangName + '):'
      };

    const header = [t.title, t.role, t.task, '', t.rulesTitle].concat(t.rules, ['']).join('\n');

    const body = items.map(function (it, idx) {
      const idTag = it.rowId || ('#' + (idx + 1));
      if (isNewLang) {
        const lines = [
          '----------------------------------------',
          t.para + ': [' + idTag + ']',
          t.newRef,
          cleanText(it.refNew)
        ];
        if (auxLangName && it.auxText && cleanText(it.auxText)) {
          lines.push(t.aux, cleanText(it.auxText));
        }
        lines.push('');
        return lines.join('\n');
      }

      if (it.classification === 'ADDED') {
        const lines = [
          '----------------------------------------',
          t.para + ': [' + idTag + '] (' + t.typ + ': ' + t.added + ')',
          t.newRef,
          cleanText(it.refNew)
        ];
        if (auxLangName && it.auxText && cleanText(it.auxText)) {
          lines.push(t.aux, cleanText(it.auxText));
        }
        lines.push('');
        return lines.join('\n');
      }

      const summary = buildChangesSummary(it.diffResult || computeDiff(it.refOld, it.refNew), lang, 8).replace(/\n/g, ' ');
      const lines = [
        '----------------------------------------',
        t.para + ': [' + idTag + '] (' + t.typ + ': ' + t.modified + ')' + (summary ? ' ' + summary : ''),
        t.oldRef,
        cleanText(it.refOld),
        t.updRef,
        cleanText(it.refNew),
        t.approved,
        cleanText(it.targetOld)
      ];
      if (auxLangName && it.auxText && cleanText(it.auxText)) {
        lines.push(t.aux, cleanText(it.auxText));
      }
      lines.push('');
      return lines.join('\n');
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
   * Change summary (old → new version) of ONE language: the reference (langCol === refNewCol or
   * omitted) or a translated language (langCol = its old-version column, new text = translation).
   * Returns a format-independent model rendered by 'summary' export adapters.
   */
  function buildChangeSummary(params) {
    const rows = params.rows || [];
    const mappings = params.mappings || {};
    const langCol = params.langCol || mappings.refNewCol;
    const isReference = langCol === mappings.refNewCol;
    const stubs = isReference ? null : resolveStubs(rows, params.translations || {}, langCol, mappings);

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
    const pctChanged = total > 0 ? Math.round(((modified + added + removed) / total) * 100) : 0;

    return {
      uiLang: params.uiLang || 'cs',
      documentName: params.documentName || '',
      language: getLanguageLabel(mappings, langCol),
      isReference: isReference,
      refOldCol: mappings.refOldCol,
      refNewCol: mappings.refNewCol,
      date: new Date().toISOString().split('T')[0],
      stats: {
        totalRows: total,
        unchangedCount: unchanged,
        modifiedCount: modified,
        addedCount: added,
        removedCount: removed,
        pctChanged: pctChanged
      },
      translationStatus: isReference ? null : status,
      modifiedItems: modifiedItems,
      addedItems: addedItems,
      removedItems: removedItems
    };
  }

  /**
   * Format-independent model of the new version of one language (targetCol = its old-version
   * column, or the reference new column), rendered by 'document' export adapters.
   * Paragraphs carry no headings; untranslated ones are marked as missing.
   */
  function buildDocumentModel(options) {
    const rows = options.rows || [];
    const mappings = options.mappings || {};
    const targetCol = options.targetCol;
    const tpl = options.template || {};
    const isReference = targetCol === mappings.refNewCol;
    const label = getLanguageLabel(mappings, targetCol);
    // Stubs for languages never opened in the studio are generated on the fly, so unchanged
    // paragraphs carry the approved old text and changed ones are reported as missing.
    const stubs = isReference ? [] : resolveStubs(rows, options.translations || {}, targetCol, mappings);

    const paragraphs = [];
    let missingCount = 0;
    let unverifiedCount = 0;
    rows.forEach(function (row, idx) {
      const classification = classifyRow(row[mappings.refOldCol], row[mappings.refNewCol]);
      if (classification === 'REMOVED' || classification === 'EMPTY') return;

      let text;
      let missing = false;
      if (isReference) {
        text = String(cellValue(row[mappings.refNewCol]));
      } else {
        const item = stubs[idx];
        text = item && item.targetNew ? String(item.targetNew) : '';
        missing = !text.trim();
        if (missing) missingCount++;
        else if (!item.isVerified) unverifiedCount++;
      }
      paragraphs.push({ rowNumber: rowNumberOf(options.rowNumbers, idx), text: text.trim(), missing: missing });
    });

    return {
      uiLang: options.uiLang || 'cs',
      language: label,
      isReference: isReference,
      title: tpl.docTitle || options.documentName || '',
      subtitle: tpl.docSubtitle ? tpl.docSubtitle.replace(/\{LANG\}/g, label) : label,
      organization: tpl.companyName || '',
      versionLabel: tpl.versionLabel || '',
      dateLabel: tpl.effectiveDate || '',
      introduction: tpl.introPreamble || '',
      accentColor: /^#[0-9a-f]{3,8}$/i.test(tpl.accentColor || '') ? tpl.accentColor : '#2563eb',
      signatures: tpl.showSignatures === true ? [tpl.signDept1 || '', tpl.signDept2 || ''] : [],
      paragraphs: paragraphs,
      missingCount: missingCount,
      unverifiedCount: unverifiedCount,
      isDraft: missingCount > 0 || unverifiedCount > 0
    };
  }

  return {
    escapeHtml: escapeHtml,
    escapeHtmlMultiline: escapeHtmlMultiline,
    mdQuote: mdQuote,
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
    isNewLanguage: isNewLanguage,
    generateTargetStub: generateTargetStub,
    rehydrateStubs: rehydrateStubs,
    resolveStubs: resolveStubs,
    getStubState: getStubState,
    idsMatch: idsMatch,
    buildCopilotPrompt: buildCopilotPrompt,
    buildVerificationPrompt: buildVerificationPrompt,
    buildBatchCopilotPrompt: buildBatchCopilotPrompt,
    parseBatchTranslationResponse: parseBatchTranslationResponse,
    buildChangeSummary: buildChangeSummary,
    buildDocumentModel: buildDocumentModel,
    prepareExportData: prepareExportData,
    prepareLanguageExportData: prepareLanguageExportData,
    prepareExportTables: prepareExportTables
  };
}));
