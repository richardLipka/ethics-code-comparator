/**
 * HTML document export adapters (default document exporters): render the document model
 * (ComparatorEngine.buildDocumentModel) as a formatted A4 HTML page and
 *   - 'html-print': open it in a print window (print / save as PDF),
 *   - 'html-file':  download it as a standalone HTML file.
 * Registers itself in the Exporters registry. Universal module: works in browser and Node.js.
 */
(function (root, factory) {
  if (typeof define === 'function' && define.amd) {
    define(['comparator-engine', 'exporters'], factory);
  } else if (typeof module === 'object' && module.exports) {
    module.exports = factory(require('../core/comparator-engine'), require('./exporters'));
  } else {
    root.HtmlDocumentExporter = factory(root.ComparatorEngine, root.Exporters);
  }
}(typeof self !== 'undefined' ? self : this, function (ComparatorEngine, Exporters) {
  'use strict';

  const esc = ComparatorEngine.escapeHtml;

  function labels(isEn) {
    return isEn
      ? { untitled: 'Document – new version', paragraphs: 'Paragraphs', date: 'Date', missing: '[Translation missing]',
          print: '🖨️ Print / Save as PDF', close: '✕ Close',
          draft: function (m, u) { return `DRAFT – ${m} paragraph(s) without translation, ${u} not yet verified.`; } }
      : { untitled: 'Dokument – nová verze', paragraphs: 'Počet odstavců', date: 'Datum', missing: '[Překlad chybí]',
          print: '🖨️ Vytisknout / Uložit jako PDF', close: '✕ Zavřít',
          draft: function (m, u) { return `KONCEPT – ${m} odstavců bez překladu, ${u} dosud neověřeno.`; } };
  }

  /**
   * Renders the document model as a complete HTML page (with a print toolbar hidden in print).
   */
  function renderDocumentHtml(model) {
    const isEn = model.uiLang === 'en';
    const L = labels(isEn);
    const title = model.title || L.untitled;
    const accent = esc(model.accentColor || '#2563eb');
    const draftNotice = model.isDraft ? L.draft(model.missingCount, model.unverifiedCount) : '';

    return `<!DOCTYPE html>
<html lang="${isEn ? 'en' : 'cs'}">
<head>
  <meta charset="UTF-8">
  <title>${esc(title)} - ${esc(model.language)}</title>
  <style>
    @page { size: A4; margin: 20mm 15mm 20mm 15mm; }
    *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
    body {
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
      color: #0f172a; background-color: #f8fafc; line-height: 1.6; -webkit-font-smoothing: antialiased;
    }
    .no-print-toolbar {
      position: sticky; top: 0; z-index: 1000; background: #ffffff; border-bottom: 1px solid #e2e8f0;
      padding: 0.75rem 1.5rem; display: flex; align-items: center; justify-content: space-between; gap: 1rem;
      box-shadow: 0 2px 4px rgba(0,0,0,0.05);
    }
    .btn-print { background: #2563eb; color: #ffffff; border: none; border-radius: 6px; padding: 0.5rem 1.25rem; font-size: 0.95rem; font-weight: 700; cursor: pointer; }
    .btn-print:hover { background: #1d4ed8; }
    .btn-close { background: #f1f5f9; color: #334155; border: 1px solid #cbd5e1; border-radius: 6px; padding: 0.5rem 1rem; font-size: 0.85rem; font-weight: 600; cursor: pointer; }
    .document-page { max-width: 800px; margin: 1.5rem auto; background: #ffffff; border: 1px solid #e2e8f0; border-radius: 8px; padding: 40px 50px; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05); }
    .doc-header { border-bottom: 3px solid #0f172a; padding-bottom: 1.25rem; margin-bottom: 1.75rem; display: flex; justify-content: space-between; align-items: flex-end; gap: 1rem; }
    .company-name { font-size: 0.85rem; font-weight: 700; text-transform: uppercase; letter-spacing: 0.08em; color: #64748b; margin-bottom: 0.35rem; }
    .doc-title { font-size: 1.65rem; font-weight: 800; color: #0f172a; letter-spacing: -0.02em; line-height: 1.2; }
    .doc-subtitle { font-size: 1rem; font-weight: 600; color: ${accent}; margin-top: 0.25rem; }
    .header-meta { text-align: right; font-size: 0.8rem; color: #64748b; line-height: 1.4; white-space: nowrap; }
    .version-badge { display: inline-block; padding: 0.2rem 0.6rem; background-color: #ecfdf5; color: #047857; border: 1px solid #a7f3d0; border-radius: 9999px; font-weight: 700; font-size: 0.75rem; margin-bottom: 0.35rem; }
    .doc-intro { background-color: #f8fafc; border-left: 4px solid ${accent}; border-radius: 4px; padding: 0.85rem 1.15rem; margin-bottom: 2rem; font-size: 0.875rem; color: #334155; white-space: pre-wrap; }
    .draft-notice { border: 2px solid #b91c1c; color: #b91c1c; background: #fef2f2; font-weight: 700; padding: 0.6rem 1rem; border-radius: 4px; margin-bottom: 1.5rem; text-align: center; }
    .paragraph { font-size: 0.95rem; line-height: 1.65; color: #1e293b; text-align: justify; white-space: pre-wrap; margin-bottom: 1rem; break-inside: avoid; page-break-inside: avoid; }
    .paragraph.missing { color: #b91c1c; font-style: italic; }
    .sign-block { margin-top: 3rem; padding-top: 1.5rem; border-top: 2px dashed #cbd5e1; display: flex; justify-content: space-between; break-inside: avoid; page-break-inside: avoid; }
    .sign-box { width: 45%; border-top: 1px solid #0f172a; padding-top: 0.5rem; font-size: 0.8rem; color: #475569; }
    @media print {
      body { background: #ffffff !important; color: #000000 !important; }
      .no-print-toolbar { display: none !important; }
      .document-page { max-width: 100% !important; margin: 0 !important; padding: 0 !important; border: none !important; box-shadow: none !important; }
    }
  </style>
</head>
<body>
  <div class="no-print-toolbar">
    <div><strong>${esc(model.language)}</strong> – ${esc(title)}</div>
    <div style="display:flex; gap:0.5rem; align-items:center;">
      <button class="btn-print" onclick="window.print();">${L.print}</button>
      <button class="btn-close" onclick="window.close();">${L.close}</button>
    </div>
  </div>

  <div class="document-page">
    <div class="doc-header">
      <div>
        ${model.organization ? `<div class="company-name">${esc(model.organization)}</div>` : ''}
        <div class="doc-title">${esc(title)}</div>
        <div class="doc-subtitle">${esc(model.subtitle)}</div>
      </div>
      <div class="header-meta">
        ${model.versionLabel ? `<div><span class="version-badge">${esc(model.versionLabel)}</span></div>` : ''}
        ${model.dateLabel ? `<div>${esc(model.dateLabel)}</div>` : ''}
        <div>${L.paragraphs}: ${model.paragraphs.length}</div>
      </div>
    </div>

    ${draftNotice ? `<div class="draft-notice">${esc(draftNotice)}</div>` : ''}
    ${model.introduction ? `<div class="doc-intro">${esc(model.introduction)}</div>` : ''}

    <div class="paragraphs">
      ${model.paragraphs.map(function (p) {
        return `<p class="paragraph${p.missing ? ' missing' : ''}" data-row="${p.rowNumber}">${p.missing ? esc(L.missing) : esc(p.text)}</p>`;
      }).join('\n      ')}
    </div>

    ${model.signatures && model.signatures.length > 0 ? `
      <div class="sign-block">
        ${model.signatures.map(function (s) {
          return `<div class="sign-box"><strong>${esc(s)}</strong><br>${L.date}: ____________________</div>`;
        }).join('\n        ')}
      </div>
    ` : ''}
  </div>
</body>
</html>`;
  }

  const printAdapter = {
    id: 'html-print',
    label: 'Tisk / PDF',
    group: 'document',
    isDefault: true,
    export: function (model) {
      return { action: 'print', window: Exporters.openPrintWindow(renderDocumentHtml(model)) };
    }
  };

  const fileAdapter = {
    id: 'html-file',
    label: 'HTML',
    group: 'document',
    export: function (model) {
      return { content: renderDocumentHtml(model), mimeType: 'text/html;charset=utf-8', extension: '.html' };
    }
  };

  if (Exporters) {
    Exporters.register(printAdapter);
    Exporters.register(fileAdapter);
  }

  return {
    renderDocumentHtml: renderDocumentHtml,
    printAdapter: printAdapter,
    fileAdapter: fileAdapter
  };
}));
