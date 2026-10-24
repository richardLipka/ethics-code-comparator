/**
 * Excel / CSV import adapter (default importer), based on SheetJS.
 * Registers itself in the Importers registry. Universal module: works in browser and Node.js.
 */
(function (root, factory) {
  if (typeof define === 'function' && define.amd) {
    define(['xlsx', 'importers'], factory);
  } else if (typeof module === 'object' && module.exports) {
    module.exports = factory(require('xlsx'), require('./importers'));
  } else {
    root.ExcelImporter = factory(root.XLSX, root.Importers);
  }
}(typeof self !== 'undefined' ? self : this, function (XLSX, Importers) {
  'use strict';

  const EXTENSIONS = ['.xlsx', '.xlsm', '.xls', '.ods', '.csv'];

  function isCsv(fileName, mimeType) {
    return /\.csv$/i.test(fileName || '') || mimeType === 'text/csv';
  }

  function pad2(n) {
    return (n < 10 ? '0' : '') + n;
  }

  /**
   * Text of a date cell as an unambiguous ISO date ("2024-02-01", with " 12:30" when the format
   * shows a time), or null for other cells. SheetJS does not render every Excel date format
   * (e.g. "d.m.yyyy") correctly, and the date is the same in every time zone this way.
   */
  function dateCellText(cell) {
    if (!cell) return null;
    let parts = null;
    if (cell.t === 'd' && cell.v instanceof Date && !isNaN(cell.v)) {
      const d = cell.v;
      parts = { y: d.getFullYear(), m: d.getMonth() + 1, d: d.getDate(), H: d.getHours(), M: d.getMinutes() };
    } else if (cell.t === 'n' && typeof cell.z === 'string' && XLSX.SSF.is_date(cell.z)) {
      parts = XLSX.SSF.parse_date_code(cell.v);
    }
    if (!parts) return null;
    const hasTime = cell.t === 'n' ? /[hs]/i.test(cell.z.replace(/"[^"]*"|\[[^\]]*\]/g, '')) : (parts.H || parts.M);
    const date = parts.y + '-' + pad2(parts.m) + '-' + pad2(parts.d);
    return hasTime ? date + ' ' + pad2(parts.H) + ':' + pad2(parts.M) : date;
  }

  /**
   * Reads one sheet into the table model. With hasHeaderRow the first row holds the column names
   * (in the sheet's own order – Object.keys would move numeric names first); without it the
   * columns are named by their letters and the first row is data.
   */
  function readWorksheet(sheet, hasHeaderRow) {
    const range = sheet && sheet['!ref'] ? XLSX.utils.decode_range(sheet['!ref']) : null;
    if (!range) return null;
    // Date cells become text (a serial number would be meaningless in the texts, the project file
    // and the exported workbook)
    Object.keys(sheet).forEach(function (addr) {
      if (addr.charAt(0) === '!') return;
      const text = dateCellText(sheet[addr]);
      if (text !== null) sheet[addr] = { t: 's', v: text };
    });
    const rawRows = XLSX.utils.sheet_to_json(sheet, hasHeaderRow ? { defval: '' } : { header: 'A', defval: '' });
    if (!rawRows || rawRows.length === 0) return null;

    const letters = [];
    for (let c = range.s.c; c <= range.e.c; c++) letters.push(XLSX.utils.encode_col(c));

    let headers;
    let rows = rawRows;
    const columnLetters = {};
    if (hasHeaderRow) {
      const headerRow = XLSX.utils.sheet_to_json(sheet, { header: 1, range: range.s.r, defval: '' })[0] || [];
      const keys = Object.keys(rawRows[0]);
      const ordered = [];
      headerRow.forEach(function (h, i) {
        let key = keys.indexOf(String(h)) !== -1 && ordered.indexOf(String(h)) === -1 ? String(h) : undefined;
        if (key === undefined) {
          key = keys.find(function (k) { return /^__EMPTY/.test(k) && ordered.indexOf(k) === -1; });
        }
        if (key !== undefined && ordered.indexOf(key) === -1) {
          ordered.push(key);
          columnLetters[key] = letters[i];
        }
      });
      headers = ordered.length === keys.length ? ordered : keys;

      // A header cell holding an empty string gives the key "" – rename it like SheetJS names
      // missing header cells, so the column is shown as "Column C"
      if (headers.indexOf('') !== -1) {
        let n = 0;
        let emptyKey = '__EMPTY';
        while (headers.indexOf(emptyKey) !== -1) emptyKey = '__EMPTY_' + (++n);
        headers = headers.map(function (h) { return h === '' ? emptyKey : h; });
        columnLetters[emptyKey] = columnLetters[''];
        delete columnLetters[''];
        rows = rawRows.map(function (r) {
          const copy = Object.assign({}, r);
          copy[emptyKey] = r[''];
          delete copy[''];
          return copy;
        });
      }
    } else {
      headers = letters;
      letters.forEach(function (l) { columnLetters[l] = l; });
    }

    return {
      headers: headers,
      rows: rows,
      rowNumbers: rawRows.map(function (r, i) {
        return typeof r.__rowNum__ === 'number' ? r.__rowNum__ + 1 : i + (hasHeaderRow ? 2 : 1);
      }),
      columnLetters: columnLetters,
      // Sheet row that holds the column names (null without a header row)
      headerRowNumber: hasHeaderRow ? range.s.r + 1 : null
    };
  }

  const adapter = {
    id: 'excel',
    label: 'Excel / CSV',
    extensions: EXTENSIONS,
    canImport: function (fileName, mimeType) {
      return EXTENSIONS.indexOf(Importers.extensionOf(fileName)) !== -1 || isCsv(fileName, mimeType);
    },
    // CSV is read as bytes and decoded here (UTF-8, or a legacy code page such as Windows-1250
    // that Excel uses for "CSV (separated by semicolons)"); SheetJS alone would garble diacritics
    readAs: 'arraybuffer',
    parse: function (content, options) {
      const fileName = options && options.fileName;
      const csv = typeof content === 'string' || isCsv(fileName, options && options.mimeType);
      // CSV cells stay text: "0012" must not become 12 and "1.2.2024" not a date serial number.
      // Date cells of a spreadsheet are read as dates and shown as formatted in the sheet.
      const workbook = csv
        ? XLSX.read(Importers.decodeText(content), { type: 'string', raw: true })
        : XLSX.read(new Uint8Array(content), { type: 'array', cellNF: true });
      return {
        fileName: options && options.fileName,
        sheetNames: workbook.SheetNames.slice(),
        readSheet: function (name, readOptions) {
          const hasHeaderRow = !readOptions || readOptions.hasHeaderRow !== false;
          return readWorksheet(workbook.Sheets[name], hasHeaderRow);
        }
      };
    }
  };

  if (Importers) Importers.register(adapter);
  return adapter;
}));
