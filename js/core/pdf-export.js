/**
 * PDF & Printable Document Exporter - Code of Ethics 2026
 * Generates beautifully formatted, high-fidelity printable documents for each language mutation.
 * Supports direct vector PDF export via browser print dialog and standalone HTML download.
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

  /**
   * Generates the printable HTML string for a specific language mutation of the 2026 Code.
   */
  function generateDocumentHtml(options) {
    const rows = options.rows || [];
    const headers = options.headers || [];
    const mappings = options.mappings || {};
    const translations = options.translations || {};
    const targetCol = options.targetCol; // e.g. 'CS 2024' or 'DE 2026 (Referenz)'
    const lang = (I18n && typeof I18n.getLanguage === 'function') ? I18n.getLanguage() : (options.uiLang || 'cs');
    const tpl = options.template || {};

    const companyName = tpl.companyName || options.companyName || 'Corporate Group';
    const accentColor = /^#[0-9a-f]{3,8}$/i.test(tpl.accentColor || '') ? tpl.accentColor : '#2563eb';

    const isGermanRef = targetCol === mappings.refNewCol;
    const cleanLangName = isGermanRef 
      ? 'Deutsch (DE - Referenz)' 
      : targetCol.replace(/2024/i, '2026').replace(/\(.*\)/, '').trim();

    const idCol = headers[0] || 'Artikel-Nr.';
    const sectionCol = headers[1] || 'Bereich';

    // Stubs for languages never opened in the studio are generated on the fly, so unchanged
    // articles carry the approved 2024 text and changed ones are reported as missing.
    const stubs = isGermanRef ? [] : ComparatorEngine.resolveStubs(rows, translations, targetCol, mappings);

    // Filter valid 2026 articles (omit REMOVED and EMPTY)
    const activeArticles = [];
    let missingCount = 0;
    let unverifiedCount = 0;
    rows.forEach(function (row, idx) {
      const classification = ComparatorEngine.classifyRow(row[mappings.refOldCol], row[mappings.refNewCol]);
      if (classification === 'REMOVED' || classification === 'EMPTY') {
        return; // Exclude deleted articles in 2026 version
      }

      let text = '';
      let missing = false;
      if (isGermanRef) {
        text = row[mappings.refNewCol] !== undefined && row[mappings.refNewCol] !== null ? String(row[mappings.refNewCol]) : '';
      } else {
        const item = stubs[idx];
        text = item && item.targetNew ? String(item.targetNew) : '';
        missing = !text.trim();
        if (missing) missingCount++;
        else if (!item.isVerified) unverifiedCount++;
      }

      activeArticles.push({
        index: idx + 1,
        id: row[idCol] || ('Art. ' + (idx + 1)),
        section: row[sectionCol] || '',
        classification: classification,
        text: text,
        missing: missing
      });
    });

    const docTitle = tpl.docTitle || (lang === 'en' ? 'CODE OF CONDUCT & ETHICS 2026' : 'ETICKÝ KODEX A PRAVIDLA CHOVÁNÍ 2026');
    const docSubtitle = tpl.docSubtitle 
      ? tpl.docSubtitle.replace(/\{LANG\}/g, cleanLangName)
      : (lang === 'en' ? `Official Certified Edition – ${cleanLangName}` : `Oficiální certifikované znění – ${cleanLangName}`);
    const dateLabel = tpl.effectiveDate || (lang === 'en' ? 'Effective Date: January 1, 2026' : 'Účinnost od: 1. ledna 2026');
    const versionLabel = tpl.versionLabel || (lang === 'en' ? 'Version 2.6 – Compliance Approved' : 'Verze 2.6 – Schváleno compliance');
    const introPreamble = tpl.introPreamble || (lang === 'en' 
      ? 'This document represents the official and binding Code of Conduct and Ethics for the 2026 cycle. All provisions herein supersede previous editions and apply to all operational entities and employees.'
      : 'Tento dokument představuje oficiální a závazný Etický kodex pro období od roku 2026. Veškerá zde uvedená ustanovení nahrazují dřívější znění a jsou závazná pro všechny zaměstnance i vedoucí pracovníky.');
    const showSignatures = tpl.showSignatures !== false;
    const signDept1 = tpl.signDept1 || (lang === 'en' ? 'Corporate Compliance & Legal' : 'Oddělení Compliance & Právní');
    const signDept2 = tpl.signDept2 || (lang === 'en' ? 'Executive Board' : 'Vedení společnosti');
    const printPromptLabel = lang === 'en' ? '🖨️ Print / Save as PDF' : '🖨️ Vytisknout / Uložit jako PDF';
    const closeLabel = lang === 'en' ? '✕ Close' : '✕ Zavřít';
    const totalLabel = lang === 'en' ? 'Total articles' : 'Počet článků';
    const dateWord = lang === 'en' ? 'Date' : 'Datum';
    const missingLabel = lang === 'en' ? '[Translation missing]' : '[Překlad chybí]';
    const draftNotice = (missingCount > 0 || unverifiedCount > 0)
      ? (lang === 'en'
        ? `DRAFT – ${missingCount} article(s) without translation, ${unverifiedCount} not yet verified.`
        : `KONCEPT – ${missingCount} článků bez překladu, ${unverifiedCount} dosud neověřeno.`)
      : '';

    return `<!DOCTYPE html>
<html lang="${lang === 'en' ? 'en' : 'cs'}">
<head>
  <meta charset="UTF-8">
  <title>${ComparatorEngine.escapeHtml(docTitle)} - ${ComparatorEngine.escapeHtml(cleanLangName)}</title>
  <style>
    /* ==========================================================================
       Print & High-Fidelity PDF Stylesheet
       ========================================================================== */
    @page {
      size: A4;
      margin: 20mm 15mm 20mm 15mm;
      @bottom-right {
        content: counter(page);
      }
    }

    *, *::before, *::after {
      box-sizing: border-box;
      margin: 0;
      padding: 0;
    }

    body {
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
      color: #0f172a;
      background-color: #f8fafc;
      line-height: 1.6;
      -webkit-font-smoothing: antialiased;
    }

    /* Floating Action Bar (Hidden in Print) */
    .no-print-toolbar {
      position: sticky;
      top: 0;
      z-index: 1000;
      background: #ffffff;
      border-bottom: 1px solid #e2e8f0;
      padding: 0.75rem 1.5rem;
      display: flex;
      align-items: center;
      justify-content: space-between;
      box-shadow: 0 2px 4px rgba(0,0,0,0.05);
    }

    .btn-print {
      background: #2563eb;
      color: #ffffff;
      border: none;
      border-radius: 6px;
      padding: 0.5rem 1.25rem;
      font-size: 0.95rem;
      font-weight: 700;
      cursor: pointer;
      display: inline-flex;
      align-items: center;
      gap: 0.5rem;
      transition: background 0.15s ease;
    }
    .btn-print:hover { background: #1d4ed8; }

    .btn-close {
      background: #f1f5f9;
      color: #334155;
      border: 1px solid #cbd5e1;
      border-radius: 6px;
      padding: 0.5rem 1rem;
      font-size: 0.85rem;
      font-weight: 600;
      cursor: pointer;
    }

    /* Document Sheet Container (A4 preview) */
    .document-page {
      max-width: 800px;
      margin: 1.5rem auto;
      background: #ffffff;
      border: 1px solid #e2e8f0;
      border-radius: 8px;
      padding: 40px 50px;
      box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05);
    }

    /* Document Header */
    .doc-header {
      border-bottom: 3px solid #0f172a;
      padding-bottom: 1.5rem;
      margin-bottom: 2rem;
      display: flex;
      justify-content: space-between;
      align-items: flex-end;
    }

    .company-name {
      font-size: 0.85rem;
      font-weight: 700;
      text-transform: uppercase;
      letter-spacing: 0.08em;
      color: #64748b;
      margin-bottom: 0.35rem;
    }

    .doc-title {
      font-size: 1.65rem;
      font-weight: 800;
      color: #0f172a;
      letter-spacing: -0.02em;
      line-height: 1.2;
    }

    .doc-subtitle {
      font-size: 1rem;
      font-weight: 600;
      color: ${ComparatorEngine.escapeHtml(accentColor)};
      margin-top: 0.25rem;
    }

    .header-meta {
      text-align: right;
      font-size: 0.8rem;
      color: #64748b;
      line-height: 1.4;
    }

    .badge-certified {
      display: inline-block;
      padding: 0.2rem 0.6rem;
      background-color: #ecfdf5;
      color: #047857;
      border: 1px solid #a7f3d0;
      border-radius: 9999px;
      font-weight: 700;
      font-size: 0.75rem;
      text-transform: uppercase;
      margin-bottom: 0.35rem;
    }

    /* Preamble Callout */
    .doc-intro {
      background-color: #f8fafc;
      border-left: 4px solid ${ComparatorEngine.escapeHtml(accentColor)};
      border-radius: 4px;
      padding: 0.85rem 1.15rem;
      margin-bottom: 2rem;
      font-size: 0.875rem;
      color: #334155;
    }

    /* Article Cards */
    .article-item {
      margin-bottom: 1.5rem;
      padding-bottom: 1.5rem;
      border-bottom: 1px solid #f1f5f9;
      page-break-inside: avoid;
      break-inside: avoid;
    }

    .article-item:last-child {
      border-bottom: none;
    }

    .article-item-header {
      display: flex;
      align-items: baseline;
      gap: 0.75rem;
      margin-bottom: 0.45rem;
    }

    .article-id-tag {
      font-family: ui-monospace, monospace;
      font-size: 0.95rem;
      font-weight: 800;
      color: #0f172a;
      background: #f1f5f9;
      padding: 0.15rem 0.5rem;
      border-radius: 4px;
    }

    .article-section-title {
      font-size: 0.95rem;
      font-weight: 700;
      color: #334155;
    }

    .article-text {
      font-size: 0.95rem;
      line-height: 1.65;
      color: #1e293b;
      text-align: justify;
      white-space: pre-wrap;
    }

    .article-text.missing {
      color: #b91c1c;
      font-style: italic;
    }

    .draft-notice {
      border: 2px solid #b91c1c;
      color: #b91c1c;
      background: #fef2f2;
      font-weight: 700;
      padding: 0.6rem 1rem;
      border-radius: 4px;
      margin-bottom: 1.5rem;
      text-align: center;
    }

    /* Sign-off Block */
    .sign-block {
      margin-top: 3rem;
      padding-top: 1.5rem;
      border-top: 2px dashed #cbd5e1;
      display: flex;
      justify-content: space-between;
      page-break-inside: avoid;
      break-inside: avoid;
    }

    .sign-box {
      width: 45%;
      border-top: 1px solid #0f172a;
      padding-top: 0.5rem;
      font-size: 0.8rem;
      color: #475569;
    }

    /* Print Overrides */
    @media print {
      body {
        background: #ffffff !important;
        color: #000000 !important;
      }
      .no-print-toolbar {
        display: none !important;
      }
      .document-page {
        max-width: 100% !important;
        margin: 0 !important;
        padding: 0 !important;
        border: none !important;
        box-shadow: none !important;
      }
      .article-item {
        break-inside: avoid;
        page-break-inside: avoid;
      }
    }
  </style>
</head>
<body>

  <!-- Toolbar with print button -->
  <div class="no-print-toolbar">
    <div>
      <strong>${ComparatorEngine.escapeHtml(cleanLangName)}</strong> – ${ComparatorEngine.escapeHtml(docTitle)}
    </div>
    <div style="display:flex; gap:0.5rem; align-items:center;">
      <button class="btn-print" onclick="window.print();">${printPromptLabel}</button>
      <button class="btn-close" onclick="window.close();">${closeLabel}</button>
    </div>
  </div>

  <!-- Document Page Container -->
  <div class="document-page">
    <div class="doc-header">
      <div>
        <div class="company-name">${ComparatorEngine.escapeHtml(companyName)}</div>
        <div class="doc-title">${ComparatorEngine.escapeHtml(docTitle)}</div>
        <div class="doc-subtitle">${ComparatorEngine.escapeHtml(docSubtitle)}</div>
      </div>
      <div class="header-meta">
        <div><span class="badge-certified">${ComparatorEngine.escapeHtml(versionLabel)}</span></div>
        <div>${ComparatorEngine.escapeHtml(dateLabel)}</div>
        <div>${totalLabel}: ${activeArticles.length}</div>
      </div>
    </div>

    ${draftNotice ? `<div class="draft-notice">${ComparatorEngine.escapeHtml(draftNotice)}</div>` : ''}

    <div class="doc-intro">
      ${ComparatorEngine.escapeHtml(introPreamble)}
    </div>

    <div class="articles-list">
      ${activeArticles.map(function (art) {
        return `
          <div class="article-item">
            <div class="article-item-header">
              <span class="article-id-tag">${ComparatorEngine.escapeHtml(art.id)}</span>
              ${art.section ? `<span class="article-section-title">${ComparatorEngine.escapeHtml(art.section)}</span>` : ''}
            </div>
            <div class="article-text${art.missing ? ' missing' : ''}">${art.missing ? missingLabel : ComparatorEngine.escapeHtml(art.text)}</div>
          </div>
        `;
      }).join('')}
    </div>

    ${showSignatures ? `
      <div class="sign-block">
        <div class="sign-box">
          <strong>${ComparatorEngine.escapeHtml(signDept1)}</strong><br>
          ${lang === 'en' ? 'Certified Translation & Review' : 'Certifikace a schválení překladu'}<br>
          ${dateWord}: ____________________
        </div>
        <div class="sign-box">
          <strong>${ComparatorEngine.escapeHtml(signDept2)}</strong><br>
          ${lang === 'en' ? 'Approval & Implementation' : 'Schválení a implementace'}<br>
          ${dateWord}: ____________________
        </div>
      </div>
    ` : ''}
  </div>

</body>
</html>`;
  }

  /**
   * Generates printable HTML for Managerial Change Summary.
   */
  function generateManagerialReportHtml(options) {
    const summary = ComparatorEngine.generateManagerialChangeSummary(options);
    const lang = (I18n && typeof I18n.getLanguage === 'function') ? I18n.getLanguage() : (options.uiLang || 'cs');
    const isEn = lang === 'en';
    const companyName = options.companyName || 'Corporate Group';

    return `<!DOCTYPE html>
<html lang="${isEn ? 'en' : 'cs'}">
<head>
  <meta charset="UTF-8">
  <title>${isEn ? 'Management Change Summary' : 'Manažerské shrnutí změn'} - ${ComparatorEngine.escapeHtml(companyName)}</title>
  <style>
    @page { size: A4; margin: 20mm 15mm; }
    body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Arial, sans-serif; color: #0f172a; background: #f8fafc; line-height: 1.5; margin: 0; }
    .no-print-toolbar { position: sticky; top: 0; background: #ffffff; border-bottom: 1px solid #e2e8f0; padding: 0.75rem 1.5rem; display: flex; justify-content: space-between; align-items: center; box-shadow: 0 2px 4px rgba(0,0,0,0.05); z-index: 100; }
    .btn-print { background: #2563eb; color: #ffffff; border: none; border-radius: 6px; padding: 0.5rem 1.25rem; font-weight: 700; cursor: pointer; }
    .btn-close { background: #f1f5f9; color: #334155; border: 1px solid #cbd5e1; border-radius: 6px; padding: 0.5rem 1rem; cursor: pointer; }
    .page-container { max-width: 820px; margin: 1.5rem auto; background: #ffffff; border: 1px solid #e2e8f0; border-radius: 8px; padding: 40px 50px; box-shadow: 0 4px 6px rgba(0,0,0,0.05); }
    h1 { font-size: 1.6rem; color: #0f172a; border-bottom: 3px solid #2563eb; padding-bottom: 0.5rem; margin-bottom: 1rem; }
    h2 { font-size: 1.2rem; color: #1e40af; margin-top: 1.75rem; margin-bottom: 0.75rem; border-bottom: 1px solid #e2e8f0; padding-bottom: 0.25rem; }
    h3 { font-size: 1rem; color: #0f172a; margin-top: 1rem; margin-bottom: 0.35rem; }
    table { width: 100%; border-collapse: collapse; margin-bottom: 1rem; font-size: 0.9rem; }
    th { background: #f1f5f9; color: #334155; font-weight: 700; border: 1px solid #cbd5e1; padding: 8px 10px; text-align: left; }
    td { border: 1px solid #e2e8f0; padding: 8px 10px; vertical-align: top; }
    .kpi-table th { background: #2563eb; color: #ffffff; }
    .badge { font-weight: 700; padding: 2px 6px; border-radius: 4px; font-size: 0.75rem; }
    .badge-mod { background: #fef3c7; color: #b45309; }
    .badge-add { background: #dcfce7; color: #15803d; }
    .badge-rem { background: #fee2e2; color: #b91c1c; }
    ins { background: #dcfce7; color: #15803d; text-decoration: none; font-weight: 700; }
    del { background: #fee2e2; color: #b91c1c; text-decoration: line-through; }
    blockquote { border-left: 3px solid #94a3b8; margin: 0.5rem 0; padding-left: 0.75rem; color: #334155; font-style: italic; background: #f8fafc; padding: 0.5rem; }
    @media print {
      body { background: #ffffff !important; }
      .no-print-toolbar { display: none !important; }
      .page-container { max-width: 100% !important; margin: 0 !important; padding: 0 !important; border: none !important; box-shadow: none !important; }
      h2, h3, tr { break-inside: avoid; page-break-inside: avoid; }
    }
  </style>
</head>
<body>
  <div class="no-print-toolbar">
    <div><strong>${isEn ? 'Management Change Log' : 'Manažerské shrnutí změn'}</strong> – ${ComparatorEngine.escapeHtml(companyName)}</div>
    <div>
      <button class="btn-print" onclick="window.print();">${isEn ? '🖨️ Print / Save as PDF' : '🖨️ Vytisknout / Uložit PDF'}</button>
      <button class="btn-close" onclick="window.close();">${isEn ? '✕ Close' : '✕ Zavřít'}</button>
    </div>
  </div>
  <div class="page-container">
    <h1>${isEn ? 'Code of Conduct & Ethics – Management Change Summary (2024 → 2026)' : 'Etický kodex – Manažerské shrnutí změn (2024 → 2026)'}</h1>
    <p style="color:#64748b; font-size:0.875rem; margin-bottom:1.5rem;">
      <strong>${isEn ? 'Organization' : 'Organizace'}:</strong> ${ComparatorEngine.escapeHtml(companyName)} | 
      <strong>${isEn ? 'Date' : 'Datum'}:</strong> ${new Date().toLocaleDateString()}<br>
      <strong>${isEn ? 'Reference Columns' : 'Referenční sloupce'}:</strong> ${ComparatorEngine.escapeHtml(options.mappings.refOldCol)} &rarr; ${ComparatorEngine.escapeHtml(options.mappings.refNewCol)}
    </p>

    <h2>${isEn ? '1. Executive KPI Summary' : '1. Manažerské shrnutí a statistika změn'}</h2>
    <table class="kpi-table">
      <tr>
        <th>${isEn ? 'Revision Category' : 'Kategorie revize'}</th>
        <th>${isEn ? 'Article Count' : 'Počet odstavců'}</th>
        <th>${isEn ? 'Percentage' : 'Procentuální podíl'}</th>
      </tr>
      <tr>
        <td><strong>${isEn ? 'Total Evaluated Articles' : 'Celkem hodnocených odstavců'}</strong></td>
        <td><strong>${summary.kpis.total}</strong></td>
        <td><strong>100%</strong></td>
      </tr>
      <tr>
        <td>${isEn ? 'Unchanged (Directly Certified & Adopted)' : 'Beze změny (Platné a přímo převzaté)'}</td>
        <td>${summary.kpis.unchanged}</td>
        <td>${summary.kpis.total > 0 ? Math.round((summary.kpis.unchanged/summary.kpis.total)*100) : 0}%</td>
      </tr>
      <tr>
        <td><span class="badge badge-mod">${isEn ? 'Modified Articles' : 'Upravené odstavce'}</span></td>
        <td>${summary.kpis.modified}</td>
        <td>${summary.kpis.total > 0 ? Math.round((summary.kpis.modified/summary.kpis.total)*100) : 0}%</td>
      </tr>
      <tr>
        <td><span class="badge badge-add">${isEn ? 'Newly Added Articles' : 'Nově přidané odstavce'}</span></td>
        <td>${summary.kpis.added}</td>
        <td>${summary.kpis.total > 0 ? Math.round((summary.kpis.added/summary.kpis.total)*100) : 0}%</td>
      </tr>
      <tr>
        <td><span class="badge badge-rem">${isEn ? 'Removed / Retired Articles' : 'Zrušené / Odstraněné odstavce'}</span></td>
        <td>${summary.kpis.removed}</td>
        <td>${summary.kpis.total > 0 ? Math.round((summary.kpis.removed/summary.kpis.total)*100) : 0}%</td>
      </tr>
    </table>

    <h2>${isEn ? '2. Modified Articles Detail' : '2. Detailně upravené články'}</h2>
    ${summary.modifiedItems.map(function(m){
      const diffHtml = ComparatorEngine.renderDiffHtml(m.diffResult, 'inline');
      return `
        <h3>${ComparatorEngine.escapeHtml(m.id)} ${m.section ? '– ' + ComparatorEngine.escapeHtml(m.section) : ''} <span class="badge badge-mod">${isEn ? 'Modified' : 'Upraveno'}</span></h3>
        <table>
          <tr>
            <th style="width:50%;">${isEn ? 'Original German (2024)' : 'Původní němčina (2024)'}</th>
            <th style="width:50%;">${isEn ? 'Updated German (2026) with Diff' : 'Nová němčina (2026) s vyznačeným rozdílem'}</th>
          </tr>
          <tr>
            <td>${ComparatorEngine.escapeHtmlMultiline(m.refOld)}</td>
            <td>${diffHtml}</td>
          </tr>
        </table>
      `;
    }).join('')}

    ${summary.addedItems.length > 0 ? `
      <h2>${isEn ? '3. Newly Introduced Articles' : '3. Nově zavedené články'}</h2>
      ${summary.addedItems.map(function(a){
        return `
          <h3>➕ ${ComparatorEngine.escapeHtml(a.id)} ${a.section ? '– ' + ComparatorEngine.escapeHtml(a.section) : ''} <span class="badge badge-add">${isEn ? 'Added' : 'Přidáno'}</span></h3>
          <blockquote>${ComparatorEngine.escapeHtmlMultiline(a.refNew)}</blockquote>
        `;
      }).join('')}
    ` : ''}

    ${summary.removedItems.length > 0 ? `
      <h2>${isEn ? '4. Removed / Retired Articles' : '4. Zrušené / Odstraněné články'}</h2>
      ${summary.removedItems.map(function(r){
        return `
          <h3>🗑️ ${ComparatorEngine.escapeHtml(r.id)} ${r.section ? '– ' + ComparatorEngine.escapeHtml(r.section) : ''} <span class="badge badge-rem">${isEn ? 'Removed' : 'Odstraněno'}</span></h3>
          <blockquote>${ComparatorEngine.escapeHtmlMultiline(r.refOld)}</blockquote>
        `;
      }).join('')}
    ` : ''}
  </div>
</body>
</html>`;
  }

  /**
   * Opens printable Managerial Report in a print dialog window.
   */
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
   * Downloads Managerial Report as an MS Word (.doc) document.
   */
  function downloadManagerialWordDoc(options) {
    const summary = ComparatorEngine.generateManagerialChangeSummary(options);
    // BOM so Word reliably detects UTF-8 (diacritics)
    const blob = new Blob(['﻿' + summary.htmlDoc], { type: 'application/msword;charset=utf-8' });
    const fileName = 'Eticky_Kodex_Manazerske_Shrnuti_Zmen_2024_2026.doc';
    downloadBlob(blob, fileName);
    return fileName;
  }

  /**
   * Downloads Managerial Report as a Markdown (.md) document.
   */
  function downloadManagerialMarkdown(options) {
    const summary = ComparatorEngine.generateManagerialChangeSummary(options);
    const blob = new Blob([summary.markdown], { type: 'text/markdown;charset=utf-8' });
    const fileName = 'Eticky_Kodex_Manazerske_Shrnuti_Zmen_2024_2026.md';
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
    const targetCol = options.targetCol || 'Kodex';
    const cleanName = targetCol.replace(/[^a-zA-Z0-9_-]/g, '_');
    const blob = new Blob([html], { type: 'text/html;charset=utf-8' });
    const fileName = `Eticky_Kodex_2026_${cleanName}.html`;
    downloadBlob(blob, fileName);
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
