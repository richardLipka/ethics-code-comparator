/**
 * Printable document exporter.
 * Generates a formatted document with the new version of one language (or of the reference)
 * and the printable change summary. Supports PDF via the browser print dialog and a standalone
 * HTML download.
 * Universal module: works in browser and Node.js.
 */
(function (root, factory) {
  if (typeof define === 'function' && define.amd) {
    define(['comparator-engine', 'i18n'], factory);
  } else if (typeof module === 'object' && module.exports) {
    const ComparatorEngine = require('./comparator-engine');
    module.exports = factory(ComparatorEngine, null);
  } else {
    root.PdfExporter = factory(root.ComparatorEngine, root.I18n);
  }
}(typeof self !== 'undefined' ? self : this, function (ComparatorEngine, I18n) {
  'use strict';

  function uiLanguage(options) {
    return (I18n && typeof I18n.getLanguage === 'function') ? I18n.getLanguage() : (options.uiLang || 'cs');
  }

  /**
   * Generates the printable HTML of the new version of one language. targetCol is the
   * language's old-version column, or the reference new column for the reference itself.
   * Paragraphs are printed as they are – no headings are assumed in the data.
   */
  function generateDocumentHtml(options) {
    const rows = options.rows || [];
    const mappings = options.mappings || {};
    const translations = options.translations || {};
    const targetCol = options.targetCol;
    const lang = uiLanguage(options);
    const isEn = lang === 'en';
    const tpl = options.template || {};
    const esc = ComparatorEngine.escapeHtml;

    const accentColor = /^#[0-9a-f]{3,8}$/i.test(tpl.accentColor || '') ? tpl.accentColor : '#2563eb';
    const isReference = targetCol === mappings.refNewCol;
    const langLabel = ComparatorEngine.getLanguageLabel(mappings, targetCol);

    // Stubs for languages never opened in the studio are generated on the fly, so unchanged
    // paragraphs carry the approved old text and changed ones are reported as missing.
    const stubs = isReference ? [] : ComparatorEngine.resolveStubs(rows, translations, targetCol, mappings);

    // Paragraphs valid in the new version (omit REMOVED and EMPTY)
    const paragraphs = [];
    let missingCount = 0;
    let unverifiedCount = 0;
    rows.forEach(function (row, idx) {
      const classification = ComparatorEngine.classifyRow(row[mappings.refOldCol], row[mappings.refNewCol]);
      if (classification === 'REMOVED' || classification === 'EMPTY') return;

      let text = '';
      let missing = false;
      if (isReference) {
        text = row[mappings.refNewCol] !== undefined && row[mappings.refNewCol] !== null ? String(row[mappings.refNewCol]) : '';
      } else {
        const item = stubs[idx];
        text = item && item.targetNew ? String(item.targetNew) : '';
        missing = !text.trim();
        if (missing) missingCount++;
        else if (!item.isVerified) unverifiedCount++;
      }
      paragraphs.push({ rowNumber: ComparatorEngine.rowNumberOf(options.rowNumbers, idx), text: text.trim(), missing: missing });
    });

    const docTitle = tpl.docTitle || options.documentName || (isEn ? 'Document – new version' : 'Dokument – nová verze');
    const docSubtitle = tpl.docSubtitle
      ? tpl.docSubtitle.replace(/\{LANG\}/g, langLabel)
      : langLabel;
    const companyName = tpl.companyName || '';
    const versionLabel = tpl.versionLabel || '';
    const dateLabel = tpl.effectiveDate || '';
    const introPreamble = tpl.introPreamble || '';
    const showSignatures = tpl.showSignatures === true;
    const signDept1 = tpl.signDept1 || '';
    const signDept2 = tpl.signDept2 || '';
    const totalLabel = isEn ? 'Paragraphs' : 'Počet odstavců';
    const dateWord = isEn ? 'Date' : 'Datum';
    const missingLabel = isEn ? '[Translation missing]' : '[Překlad chybí]';
    const draftNotice = (missingCount > 0 || unverifiedCount > 0)
      ? (isEn
        ? `DRAFT – ${missingCount} paragraph(s) without translation, ${unverifiedCount} not yet verified.`
        : `KONCEPT – ${missingCount} odstavců bez překladu, ${unverifiedCount} dosud neověřeno.`)
      : '';

    return `<!DOCTYPE html>
<html lang="${isEn ? 'en' : 'cs'}">
<head>
  <meta charset="UTF-8">
  <title>${esc(docTitle)} - ${esc(langLabel)}</title>
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
    .doc-subtitle { font-size: 1rem; font-weight: 600; color: ${esc(accentColor)}; margin-top: 0.25rem; }
    .header-meta { text-align: right; font-size: 0.8rem; color: #64748b; line-height: 1.4; white-space: nowrap; }
    .version-badge { display: inline-block; padding: 0.2rem 0.6rem; background-color: #ecfdf5; color: #047857; border: 1px solid #a7f3d0; border-radius: 9999px; font-weight: 700; font-size: 0.75rem; margin-bottom: 0.35rem; }
    .doc-intro { background-color: #f8fafc; border-left: 4px solid ${esc(accentColor)}; border-radius: 4px; padding: 0.85rem 1.15rem; margin-bottom: 2rem; font-size: 0.875rem; color: #334155; white-space: pre-wrap; }
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
    <div><strong>${esc(langLabel)}</strong> – ${esc(docTitle)}</div>
    <div style="display:flex; gap:0.5rem; align-items:center;">
      <button class="btn-print" onclick="window.print();">${isEn ? '🖨️ Print / Save as PDF' : '🖨️ Vytisknout / Uložit jako PDF'}</button>
      <button class="btn-close" onclick="window.close();">${isEn ? '✕ Close' : '✕ Zavřít'}</button>
    </div>
  </div>

  <div class="document-page">
    <div class="doc-header">
      <div>
        ${companyName ? `<div class="company-name">${esc(companyName)}</div>` : ''}
        <div class="doc-title">${esc(docTitle)}</div>
        <div class="doc-subtitle">${esc(docSubtitle)}</div>
      </div>
      <div class="header-meta">
        ${versionLabel ? `<div><span class="version-badge">${esc(versionLabel)}</span></div>` : ''}
        ${dateLabel ? `<div>${esc(dateLabel)}</div>` : ''}
        <div>${totalLabel}: ${paragraphs.length}</div>
      </div>
    </div>

    ${draftNotice ? `<div class="draft-notice">${esc(draftNotice)}</div>` : ''}
    ${introPreamble ? `<div class="doc-intro">${esc(introPreamble)}</div>` : ''}

    <div class="paragraphs">
      ${paragraphs.map(function (p) {
        return `<p class="paragraph${p.missing ? ' missing' : ''}" data-row="${p.rowNumber}">${p.missing ? esc(missingLabel) : esc(p.text)}</p>`;
      }).join('\n      ')}
    </div>

    ${showSignatures ? `
      <div class="sign-block">
        <div class="sign-box"><strong>${esc(signDept1)}</strong><br>${dateWord}: ____________________</div>
        <div class="sign-box"><strong>${esc(signDept2)}</strong><br>${dateWord}: ____________________</div>
      </div>
    ` : ''}
  </div>
</body>
</html>`;
  }

  /**
   * Printable change summary of one language (see ComparatorEngine.generateManagerialChangeSummary).
   */
  function generateManagerialReportHtml(options) {
    const params = Object.assign({}, options, { uiLang: uiLanguage(options), printable: true });
    return ComparatorEngine.generateManagerialChangeSummary(params).htmlDoc;
  }

  function downloadBlob(blob, fileName) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = fileName;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    // Revoking synchronously can cancel the download in some browsers
    setTimeout(function () { URL.revokeObjectURL(url); }, 1000);
  }

  function safeFilePart(text) {
    return String(text || '').replace(/[^\p{L}\p{N}_-]+/gu, '_').replace(/^_+|_+$/g, '') || 'export';
  }

  function summaryFileName(options, ext) {
    const base = safeFilePart(options.documentName || 'dokument');
    const label = safeFilePart(ComparatorEngine.getLanguageLabel(options.mappings || {}, options.langCol || (options.mappings || {}).refNewCol));
    return `${base}_zmeny_${label}.${ext}`;
  }

  /**
   * Writes HTML into a new window and opens its print dialog. A document written with
   * document.write is often already complete, in which case `onload` never fires.
   */
  function openPrintWindow(html) {
    const printWindow = window.open('', '_blank', 'width=900,height=950');
    if (!printWindow) {
      alert('Vyskakovací okno bylo zablokováno prohlížečem. Prosím povolte vyskakovací okna pro tuto stránku. / Pop-up window was blocked by the browser.');
      return null;
    }
    printWindow.document.open();
    printWindow.document.write(html);
    printWindow.document.close();

    let printed = false;
    function triggerPrint() {
      if (printed) return;
      printed = true;
      setTimeout(function () {
        try { printWindow.focus(); printWindow.print(); } catch (e) { /* window closed by user */ }
      }, 300);
    }
    if (printWindow.document.readyState === 'complete') {
      triggerPrint();
    } else {
      printWindow.addEventListener('load', triggerPrint);
    }
    return printWindow;
  }

  function printManagerialReport(options) {
    return openPrintWindow(generateManagerialReportHtml(options));
  }

  /**
   * Downloads the change summary as an MS Word (.doc) document.
   */
  function downloadManagerialWordDoc(options) {
    const summary = ComparatorEngine.generateManagerialChangeSummary(Object.assign({}, options, { uiLang: uiLanguage(options) }));
    // BOM so Word reliably detects UTF-8 (diacritics)
    const blob = new Blob(['﻿' + summary.htmlDoc], { type: 'application/msword;charset=utf-8' });
    const fileName = summaryFileName(options, 'doc');
    downloadBlob(blob, fileName);
    return fileName;
  }

  /**
   * Downloads the change summary as a Markdown (.md) document.
   */
  function downloadManagerialMarkdown(options) {
    const summary = ComparatorEngine.generateManagerialChangeSummary(Object.assign({}, options, { uiLang: uiLanguage(options) }));
    const blob = new Blob([summary.markdown], { type: 'text/markdown;charset=utf-8' });
    const fileName = summaryFileName(options, 'md');
    downloadBlob(blob, fileName);
    return fileName;
  }

  /**
   * Opens the printable document in a new window and automatically triggers the print dialog.
   */
  function printLanguageMutation(options) {
    return openPrintWindow(generateDocumentHtml(options));
  }

  /**
   * Downloads the printable document as a standalone offline HTML file.
   */
  function downloadPrintableHtml(options) {
    const html = generateDocumentHtml(options);
    const label = ComparatorEngine.getLanguageLabel(options.mappings || {}, options.targetCol);
    const fileName = `${safeFilePart(options.documentName || 'dokument')}_${safeFilePart(label)}.html`;
    downloadBlob(new Blob([html], { type: 'text/html;charset=utf-8' }), fileName);
    return fileName;
  }

  return {
    generateDocumentHtml: generateDocumentHtml,
    printLanguageMutation: printLanguageMutation,
    downloadPrintableHtml: downloadPrintableHtml,
    generateManagerialReportHtml: generateManagerialReportHtml,
    printManagerialReport: printManagerialReport,
    downloadManagerialWordDoc: downloadManagerialWordDoc,
    downloadManagerialMarkdown: downloadManagerialMarkdown
  };
}));
