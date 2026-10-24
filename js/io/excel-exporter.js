/**
 * Excel (.xlsx) export adapter for table payloads (default table exporter), based on SheetJS.
 * Registers itself in the Exporters registry. Universal module: works in browser and Node.js.
 */
(function (root, factory) {
  if (typeof define === 'function' && define.amd) {
    define(['xlsx', 'exporters'], factory);
  } else if (typeof module === 'object' && module.exports) {
    module.exports = factory(require('xlsx'), require('./exporters'));
  } else {
    root.ExcelExporter = factory(root.XLSX, root.Exporters);
  }
}(typeof self !== 'undefined' ? self : this, function (XLSX, Exporters) {
  'use strict';

  const MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

  // Excel sheet names: max. 31 characters, no []:*?/\ and unique within the workbook
  function sheetName(name, taken) {
    const base = (String(name || 'Sheet').replace(/[\\/?*[\]:]/g, '_').slice(0, 31)) || 'Sheet';
    let candidate = base;
    let n = 2;
    while (taken.indexOf(candidate) !== -1) {
      const suffix = ' (' + n + ')';
      candidate = base.slice(0, 31 - suffix.length) + suffix;
      n++;
    }
    taken.push(candidate);
    return candidate;
  }

  function buildSheet(table) {
    const ws = XLSX.utils.json_to_sheet(table.rows, { header: table.headers });
    // Columns imported without a header name (SheetJS keys "__EMPTY", "__EMPTY_1", …) keep an
    // empty header cell instead of the internal key
    table.headers.forEach(function (h, c) {
      if (/^__EMPTY(_\d+)?$/.test(h)) ws[XLSX.utils.encode_cell({ r: 0, c: c })] = { t: 's', v: '' };
    });
    ws['!cols'] = table.headers.map(function (h) {
      let maxLen = String(h).length;
      table.rows.forEach(function (r) {
        const len = r[h] !== undefined && r[h] !== null ? String(r[h]).length : 0;
        if (len > maxLen) maxLen = len;
      });
      return { wch: Math.min(Math.max(maxLen + 3, 10), 63) };
    });
    if (ws['!ref']) ws['!autofilter'] = { ref: ws['!ref'] };
    return ws;
  }

  /**
   * Builds a SheetJS workbook with one sheet per table (column widths, autofilter).
   */
  function buildWorkbook(payload) {
    const workbook = XLSX.utils.book_new();
    const taken = [];
    payload.tables.forEach(function (table) {
      XLSX.utils.book_append_sheet(workbook, buildSheet(table), sheetName(table.name, taken));
    });
    return workbook;
  }

  const adapter = {
    id: 'xlsx',
    label: 'Excel (.xlsx)',
    group: 'table',
    isDefault: true,
    export: function (payload) {
      return {
        content: XLSX.write(buildWorkbook(payload), { type: 'array', bookType: 'xlsx' }),
        mimeType: MIME,
        extension: '.xlsx'
      };
    }
  };

  if (Exporters) Exporters.register(adapter);
  adapter.buildWorkbook = buildWorkbook;
  return adapter;
}));
