/**
 * Step 2: overview of the changes between the old and new reference text.
 * Part of the application controller; the parts share the `App` namespace and are loaded in the
 * order given in index.html (js/app/core.js first, js/app/main.js last).
 */
(function (App) {
  'use strict';

  const { elements, rowLabelOf, state, t } = App;

  /**
   * Step 2: Render Reference Changes Overview
   */
  function renderReferenceOverview() {
    const refOldCol = state.mappings.refOldCol;
    const refNewCol = state.mappings.refNewCol;

    let unchangedCount = 0;
    let modifiedCount = 0;
    let addedCount = 0;
    let removedCount = 0;

    elements.refDiffTableBody.innerHTML = '';

    state.rows.forEach(function (row, idx) {
      const refOld = refOldCol ? row[refOldCol] : '';
      const refNew = row[refNewCol];
      const classification = ComparatorEngine.classifyRow(refOld, refNew);

      if (classification === 'UNCHANGED') unchangedCount++;
      else if (classification === 'MODIFIED') modifiedCount++;
      else if (classification === 'ADDED') addedCount++;
      else if (classification === 'REMOVED') removedCount++;

      const tr = document.createElement('tr');
      tr.style.borderBottom = '1px solid var(--border-color)';

      // Paragraphs have no headings of their own; the Excel row number identifies them
      const rowLabel = rowLabelOf(idx);

      // Diff preview
      let diffHtml = '';
      if (!refOldCol) {
        diffHtml = '<div class="diff-box" style="color:var(--text-primary);">' + ComparatorEngine.escapeHtml(refNew) + '</div>';
      } else if (classification === 'UNCHANGED') {
        diffHtml = '<div class="diff-unchanged" style="max-height:80px; overflow-y:auto;">' + ComparatorEngine.escapeHtml(refOld) + '</div>';
      } else if (classification === 'ADDED') {
        diffHtml = '<div class="diff-box"><ins class="diff-ins">' + ComparatorEngine.escapeHtml(refNew) + '</ins></div>';
      } else if (classification === 'REMOVED') {
        diffHtml = '<div class="diff-box"><del class="diff-del">' + ComparatorEngine.escapeHtml(refOld) + '</del></div>';
      } else {
        const diffRes = ComparatorEngine.computeDiff(refOld, refNew);
        diffHtml = '<div class="diff-box">' + ComparatorEngine.renderDiffHtml(diffRes, 'inline') + '</div>';
      }

      tr.innerHTML = `
        <td style="padding:0.75rem; font-weight:700; font-family:var(--font-mono); white-space:nowrap; vertical-align:top;">${ComparatorEngine.escapeHtml(rowLabel)}</td>
        <td style="padding:0.75rem; vertical-align:top;">${renderBadgeHtml(classification)}</td>
        <td style="padding:0.75rem; vertical-align:top;">${diffHtml}</td>
      `;

      elements.refDiffTableBody.appendChild(tr);
    });

    elements.kpiTotalRows.textContent = state.rows.length;
    elements.kpiUnchanged.textContent = unchangedCount;
    elements.kpiModified.textContent = modifiedCount;
    elements.kpiAdded.textContent = addedCount;
    elements.kpiRemoved.textContent = removedCount;
  }

  function renderBadgeHtml(classification) {
    switch (classification) {
      case 'UNCHANGED':
        return `<span class="badge badge-unchanged">${t('badgeUnchanged')}</span>`;
      case 'MODIFIED':
        return `<span class="badge badge-modified">${t('badgeModified')}</span>`;
      case 'ADDED':
        return `<span class="badge badge-added">${t('badgeAdded')}</span>`;
      case 'REMOVED':
        return `<span class="badge badge-removed">${t('badgeRemoved')}</span>`;
      default:
        return `<span class="badge">${classification}</span>`;
    }
  }

  Object.assign(App, { renderReferenceOverview: renderReferenceOverview, renderBadgeHtml: renderBadgeHtml });
}(window.App = window.App || {}));
