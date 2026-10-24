/**
 * Change summary export adapters (default summary exporters). They render the summary model
 * (ComparatorEngine.buildChangeSummary) as
 *   - 'summary-markdown': Markdown file,
 *   - 'summary-word':     Word-compatible HTML document (.doc),
 *   - 'summary-print':    print window (print / save as PDF).
 * Registers itself in the Exporters registry. Universal module: works in browser and Node.js.
 */
(function (root, factory) {
  if (typeof define === 'function' && define.amd) {
    define(['comparator-engine', 'exporters'], factory);
  } else if (typeof module === 'object' && module.exports) {
    module.exports = factory(require('../core/comparator-engine'), require('./exporters'));
  } else {
    root.SummaryExporters = factory(root.ComparatorEngine, root.Exporters);
  }
}(typeof self !== 'undefined' ? self : this, function (ComparatorEngine, Exporters) {
  'use strict';

  const esc = ComparatorEngine.escapeHtml;

  function labels(model) {
    const isEn = model.uiLang === 'en';
    const label = model.language;
    return isEn
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
  }

  function percent(model, n) {
    return model.stats.totalRows > 0 ? Math.round((n / model.stats.totalRows) * 100) : 0;
  }

  function languageText(model, L) {
    return model.language + (model.isReference ? ' (' + L.reference + ')' : '');
  }

  function renderSummaryMarkdown(model) {
    const L = labels(model);
    const s = model.stats;
    const md = [];
    md.push('# ' + L.title);
    md.push('');
    if (model.documentName) md.push('**' + L.document + ':** ' + model.documentName + '  ');
    md.push('**' + L.language + ':** ' + languageText(model, L) + '  ');
    md.push('**' + L.refCols + ':** ' + (model.refOldCol ? model.refOldCol + ' → ' : '') + model.refNewCol + '  ');
    md.push('**' + L.date + ':** ' + model.date);
    md.push('');
    md.push('## ' + L.kpiTitle);
    md.push('');
    md.push('| ' + L.metric + ' | ' + L.count + ' | ' + L.share + ' |');
    md.push('| :--- | :--- | :--- |');
    md.push('| **' + L.total + '** | ' + s.totalRows + ' | 100% |');
    md.push('| ' + L.unchanged + ' | ' + s.unchangedCount + ' | ' + percent(model, s.unchangedCount) + '% |');
    md.push('| ' + L.modified + ' | ' + s.modifiedCount + ' | ' + percent(model, s.modifiedCount) + '% |');
    md.push('| ' + L.added + ' | ' + s.addedCount + ' | ' + percent(model, s.addedCount) + '% |');
    md.push('| ' + L.removed + ' | ' + s.removedCount + ' | ' + percent(model, s.removedCount) + '% |');
    md.push('');
    md.push('> **' + L.scope + ':** ' + s.pctChanged + ' % ' + L.scopeText);
    md.push('');
    if (model.translationStatus) {
      md.push('### ' + L.statusTitle);
      md.push('');
      md.push('| ' + L.metric + ' | ' + L.count + ' |');
      md.push('| :--- | :--- |');
      md.push('| ' + L.verified + ' | ' + model.translationStatus.verified + ' |');
      md.push('| ' + L.review + ' | ' + model.translationStatus.review + ' |');
      md.push('| ' + L.missing + ' | ' + model.translationStatus.missing + ' |');
      md.push('');
    }
    if (model.modifiedItems.length > 0) {
      md.push('## ' + L.modTitle);
      md.push('');
      model.modifiedItems.forEach(function (m) {
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
        md.push(ComparatorEngine.mdQuote(m.oldText));
        md.push('');
        if (!m.missing) {
          md.push('**' + L.newVersion + ':**');
          md.push(ComparatorEngine.mdQuote(m.newText));
          md.push('');
        }
      });
    }
    if (model.addedItems.length > 0) {
      md.push('## ' + L.addTitle);
      md.push('');
      model.addedItems.forEach(function (a) {
        md.push('### ➕ ' + L.row + ' ' + a.rowNumber);
        md.push(a.missing ? '*' + L.missingText + '*' : ComparatorEngine.mdQuote(a.text));
        md.push('');
      });
    }
    if (model.removedItems.length > 0) {
      md.push('## ' + L.remTitle);
      md.push('');
      model.removedItems.forEach(function (r) {
        md.push('### 🗑️ ' + L.row + ' ' + r.rowNumber);
        md.push(ComparatorEngine.mdQuote(r.text));
        md.push('');
      });
    }
    return md.join('\n');
  }

  /**
   * Word-compatible HTML; with options.printable a print toolbar (hidden in print) is added.
   */
  function renderSummaryHtml(model, options) {
    const L = labels(model);
    const s = model.stats;
    const printable = options && options.printable;
    const multiline = ComparatorEngine.escapeHtmlMultiline;
    const missingHtml = '<em class="missing">' + esc(L.missingText) + '</em>';
    const kpiRow = function (name, n, cls) {
      return '<tr><td>' + (cls ? '<span class="badge ' + cls + '">' + esc(name) + '</span>' : esc(name)) + '</td><td>' + n + '</td><td>' + percent(model, n) + '%</td></tr>';
    };

    return `<!DOCTYPE html>
<html lang="${model.uiLang === 'en' ? 'en' : 'cs'}" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns="http://www.w3.org/TR/REC-html40">
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
  ${printable ? `<div class="no-print-toolbar"><strong>${esc(L.title)}</strong><span><button class="btn-print" onclick="window.print();">${L.print}</button> <button onclick="window.close();">${L.close}</button></span></div>` : ''}
  <h1>${esc(L.title)}</h1>
  <p>
    ${model.documentName ? `<strong>${esc(L.document)}:</strong> ${esc(model.documentName)}<br>` : ''}
    <strong>${esc(L.language)}:</strong> ${esc(languageText(model, L))}<br>
    <strong>${esc(L.refCols)}:</strong> ${model.refOldCol ? esc(model.refOldCol) + ' &rarr; ' : ''}${esc(model.refNewCol)}<br>
    <strong>${esc(L.date)}:</strong> ${model.date}
  </p>

  <h2>${esc(L.kpiTitle)}</h2>
  <table class="kpi-table">
    <tr><th>${esc(L.metric)}</th><th>${esc(L.count)}</th><th>${esc(L.share)}</th></tr>
    <tr><td><strong>${esc(L.total)}</strong></td><td><strong>${s.totalRows}</strong></td><td><strong>100%</strong></td></tr>
    ${kpiRow(L.unchanged, s.unchangedCount, '')}
    ${kpiRow(L.modified, s.modifiedCount, 'badge-mod')}
    ${kpiRow(L.added, s.addedCount, 'badge-add')}
    ${kpiRow(L.removed, s.removedCount, 'badge-rem')}
  </table>
  ${model.translationStatus ? `
  <h3>${esc(L.statusTitle)}</h3>
  <table>
    <tr><th>${esc(L.metric)}</th><th>${esc(L.count)}</th></tr>
    <tr><td>${esc(L.verified)}</td><td>${model.translationStatus.verified}</td></tr>
    <tr><td>${esc(L.review)}</td><td>${model.translationStatus.review}</td></tr>
    <tr><td>${esc(L.missing)}</td><td>${model.translationStatus.missing}</td></tr>
  </table>` : ''}

  ${model.modifiedItems.length > 0 ? `<h2>${esc(L.modTitle)}</h2>` : ''}
  ${model.modifiedItems.map(function (m) {
    return `
      <h3>${esc(L.row)} ${m.rowNumber} <span class="badge badge-mod">${esc(L.modified)}</span></h3>
      <table>
        <tr><th style="width:50%;">${esc(L.oldVersion)}</th><th style="width:50%;">${esc(L.newVersionDiff)}</th></tr>
        <tr><td>${multiline(m.oldText)}</td><td>${m.missing ? missingHtml : ComparatorEngine.renderDiffHtml(m.diffResult, 'inline').replace(/\r?\n/g, '<br>')}</td></tr>
      </table>`;
  }).join('')}

  ${model.addedItems.length > 0 ? `<h2>${esc(L.addTitle)}</h2>` + model.addedItems.map(function (a) {
    return `<h3>➕ ${esc(L.row)} ${a.rowNumber} <span class="badge badge-add">${esc(L.added)}</span></h3>
      <blockquote>${a.missing ? missingHtml : multiline(a.text)}</blockquote>`;
  }).join('') : ''}

  ${model.removedItems.length > 0 ? `<h2>${esc(L.remTitle)}</h2>` + model.removedItems.map(function (r) {
    return `<h3>🗑️ ${esc(L.row)} ${r.rowNumber} <span class="badge badge-rem">${esc(L.removed)}</span></h3>
      <blockquote>${multiline(r.text)}</blockquote>`;
  }).join('') : ''}
</body>
</html>`;
  }

  const markdownAdapter = {
    id: 'summary-markdown',
    label: 'Markdown (.md)',
    group: 'summary',
    isDefault: true,
    export: function (model) {
      return { content: renderSummaryMarkdown(model), mimeType: 'text/markdown;charset=utf-8', extension: '.md' };
    }
  };

  const wordAdapter = {
    id: 'summary-word',
    label: 'Word (.doc)',
    group: 'summary',
    export: function (model) {
      // BOM so Word reliably detects UTF-8 (diacritics)
      return { content: '﻿' + renderSummaryHtml(model), mimeType: 'application/msword;charset=utf-8', extension: '.doc' };
    }
  };

  const printAdapter = {
    id: 'summary-print',
    label: 'Tisk / PDF',
    group: 'summary',
    export: function (model) {
      return { action: 'print', window: Exporters.openPrintWindow(renderSummaryHtml(model, { printable: true })) };
    }
  };

  if (Exporters) {
    Exporters.register(markdownAdapter);
    Exporters.register(wordAdapter);
    Exporters.register(printAdapter);
  }

  return {
    renderSummaryMarkdown: renderSummaryMarkdown,
    renderSummaryHtml: renderSummaryHtml,
    markdownAdapter: markdownAdapter,
    wordAdapter: wordAdapter,
    printAdapter: printAdapter
  };
}));
