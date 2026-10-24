/**
 * Excel export adapter that writes into a copy of the original workbook (.xlsx / .xlsm), so its
 * formatting, formulas, merged cells, column widths, other sheets, charts and macros are kept.
 * Only the changed cells of one sheet are rewritten; every other part of the file is copied as is.
 *
 * Group 'workbook-update', payload (ComparatorEngine.prepareWorkbookUpdate plus the file):
 *   {
 *     original: ArrayBuffer | Uint8Array,   // the workbook the table was imported from
 *     sheetName: 'Sheet1',                  // the sheet to update
 *     headerRow: 1 | null,                  // row with the column names (null = no header row)
 *     extension: '.xlsx' | '.xlsm',
 *     columns: [{ letter | null, header, styleFrom, cells: [{ row, value }] }],
 *                                           // letter null = new column after the last used one
 *     statusTable: { name, headers, rows } | null   // written to its own sheet
 *   }
 * New columns copy the cell style and width of their `styleFrom` column. Text is written as
 * inline strings, so the shared string table is not touched.
 *
 * Registers itself in the Exporters registry. Universal module: works in browser and Node.js.
 */
(function (root, factory) {
  if (typeof define === 'function' && define.amd) {
    define(['ooxml', 'exporters'], factory);
  } else if (typeof module === 'object' && module.exports) {
    module.exports = factory(require('./ooxml'), require('./exporters'));
  } else {
    root.XlsxUpdateExporter = factory(root.Ooxml, root.Exporters);
  }
}(typeof self !== 'undefined' ? self : this, function (Ooxml, Exporters) {
  'use strict';

  const esc = Ooxml.escapeXml;
  const MAX_CELL_TEXT = 32767; // Excel refuses longer cell texts
  const WORKSHEET_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml';
  const MIME = {
    '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    '.xlsm': 'application/vnd.ms-excel.sheet.macroEnabled.12'
  };

  function updateError(code, message, params) {
    const err = new Error(message);
    err.code = code;
    err.params = params || {};
    return err;
  }

  function colIndex(letters) {
    let n = 0;
    for (let i = 0; i < letters.length; i++) n = n * 26 + (letters.charCodeAt(i) - 64);
    return n;
  }

  function colLetters(index) {
    let s = '';
    while (index > 0) {
      const r = (index - 1) % 26;
      s = String.fromCharCode(65 + r) + s;
      index = Math.floor((index - 1) / 26);
    }
    return s;
  }

  function splitRef(ref) {
    const m = /^\$?([A-Z]+)\$?(\d+)$/.exec(ref);
    return m ? { col: colIndex(m[1]), row: parseInt(m[2], 10) } : null;
  }

  function attr(tag, name) {
    const m = new RegExp('\\s' + name.replace(':', '\\:') + '="([^"]*)"').exec(tag);
    return m ? Ooxml.decodeEntities(m[1]) : null;
  }

  function setAttr(tag, name, value) {
    const re = new RegExp('(\\s' + name + '=")[^"]*(")');
    if (re.test(tag)) return tag.replace(re, '$1' + value + '$2');
    return tag.replace(/(\/?>)$/, ' ' + name + '="' + value + '"$1');
  }

  function removeAttr(tag, name) {
    return tag.replace(new RegExp('\\s' + name + '="[^"]*"'), '');
  }

  // ---------------------------------------------------------------- workbook structure

  function resolveTarget(target) {
    if (target.charAt(0) === '/') return target.slice(1);
    const parts = ('xl/' + target).split('/');
    const out = [];
    parts.forEach(function (p) { if (p === '..') out.pop(); else if (p !== '.') out.push(p); });
    return out.join('/');
  }

  function readWorkbook(pkg) {
    const workbook = pkg.getText('xl/workbook.xml');
    const rels = pkg.getText('xl/_rels/workbook.xml.rels');
    if (workbook === null || rels === null) throw updateError('NOT_XLSX', 'The file is not an Excel workbook.');
    const targets = {};
    (rels.match(/<Relationship\b[^>]*>/g) || []).forEach(function (tag) {
      targets[attr(tag, 'Id')] = attr(tag, 'Target');
    });
    const sheets = (workbook.match(/<sheet\b[^>]*>/g) || []).map(function (tag) {
      const idMatch = /\s([\w]+):id="([^"]*)"/.exec(tag);
      return {
        tag: tag,
        name: attr(tag, 'name'),
        sheetId: parseInt(attr(tag, 'sheetId'), 10) || 0,
        relPrefix: idMatch ? idMatch[1] : 'r',
        path: idMatch && targets[idMatch[2]] ? resolveTarget(targets[idMatch[2]]) : null
      };
    });
    return { workbook: workbook, rels: rels, sheets: sheets };
  }

  // ---------------------------------------------------------------- worksheet editing

  function parseRows(sheetXml) {
    const open = /<sheetData\b[^>]*?(\/?)>/.exec(sheetXml);
    if (!open) throw updateError('NOT_XLSX', 'The sheet has no cell data.');
    const dataStart = open.index + open[0].length;
    const dataEnd = open[1] === '/' ? dataStart : sheetXml.indexOf('</sheetData>', dataStart);
    const inner = sheetXml.slice(dataStart, dataEnd);
    const rows = [];
    let lastRow = 0;
    const re = /<row\b[^>]*?(?:\/>|>[\s\S]*?<\/row>)/g;
    let m;
    while ((m = re.exec(inner)) !== null) {
      const openTag = /^<row\b[^>]*?\/?>/.exec(m[0])[0];
      const r = parseInt(attr(openTag, 'r'), 10) || lastRow + 1;
      lastRow = r;
      const selfClosing = /\/>$/.test(openTag);
      const body = selfClosing ? '' : m[0].slice(openTag.length, m[0].length - '</row>'.length);
      const cells = [];
      let lastCol = 0;
      const cre = /<c\b[^>]*?(?:\/>|>[\s\S]*?<\/c>)/g;
      let c;
      while ((c = cre.exec(body)) !== null) {
        const ctag = /^<c\b[^>]*?\/?>/.exec(c[0])[0];
        const ref = splitRef(attr(ctag, 'r') || '');
        const col = ref ? ref.col : lastCol + 1;
        lastCol = col;
        cells.push({ col: col, xml: c[0], style: attr(ctag, 's') });
      }
      rows.push({ r: r, openTag: selfClosing ? openTag.replace(/\/>$/, '>') : openTag, cells: cells, changed: false });
    }
    return {
      rows: rows,
      before: sheetXml.slice(0, open.index) + '<sheetData>',
      after: '</sheetData>' + sheetXml.slice(open[1] === '/' ? dataStart : dataEnd + '</sheetData>'.length)
    };
  }

  function maxUsedColumn(sheetXml, rows) {
    let max = 0;
    rows.forEach(function (row) { row.cells.forEach(function (c) { if (c.col > max) max = c.col; }); });
    (sheetXml.match(/<mergeCell\b[^>]*>/g) || []).forEach(function (tag) {
      const end = (attr(tag, 'ref') || '').split(':')[1];
      const ref = end && splitRef(end);
      if (ref && ref.col > max) max = ref.col;
    });
    return max;
  }

  function cellXml(ref, style, value) {
    const s = style !== null && style !== undefined ? ' s="' + style + '"' : '';
    if (value === '') return '<c r="' + ref + '"' + s + '/>';
    const text = String(value).slice(0, MAX_CELL_TEXT);
    return '<c r="' + ref + '"' + s + ' t="inlineStr"><is><t xml:space="preserve">' + esc(text) + '</t></is></c>';
  }

  function getRow(rows, r) {
    let i = 0;
    while (i < rows.length && rows[i].r < r) i++;
    if (rows[i] && rows[i].r === r) return rows[i];
    const row = { r: r, openTag: '<row r="' + r + '">', cells: [], changed: true };
    rows.splice(i, 0, row);
    return row;
  }

  function setCell(rows, r, col, value, styleFromCol, keepEmpty) {
    if (value === '' && !keepEmpty) return;
    const row = getRow(rows, r);
    const existing = row.cells.find(function (c) { return c.col === col; });
    const source = row.cells.find(function (c) { return c.col === styleFromCol; });
    const style = existing && existing.style !== null ? existing.style : (source ? source.style : null);
    const xml = cellXml(colLetters(col) + r, style, value);
    if (existing) {
      existing.xml = xml;
    } else {
      let i = 0;
      while (i < row.cells.length && row.cells[i].col < col) i++;
      row.cells.splice(i, 0, { col: col, xml: xml, style: style });
    }
    row.changed = true;
  }

  function serializeRows(rows) {
    return rows.map(function (row) {
      // "spans" is only an optimisation hint and would be wrong after adding cells
      const open = row.changed ? removeAttr(row.openTag, 'spans') : row.openTag;
      return open + row.cells.map(function (c) { return c.xml; }).join('') + '</row>';
    }).join('');
  }

  /**
   * New columns get the width (and column style) of the column they copy the cell style from.
   */
  function addColumnWidths(sheetXml, widths) {
    if (widths.length === 0) return sheetXml;
    const colsMatch = /<cols>([\s\S]*?)<\/cols>/.exec(sheetXml);
    let defs = colsMatch ? (colsMatch[1].match(/<col\b[^>]*\/>/g) || []).map(function (tag) {
      return { min: parseInt(attr(tag, 'min'), 10), max: parseInt(attr(tag, 'max'), 10), tag: tag };
    }) : [];
    widths.forEach(function (w) {
      const source = defs.find(function (d) { return d.min <= w.from && w.from <= d.max; });
      // A range that covers the new column is split around it
      const covering = defs.find(function (d) { return d.min <= w.col && w.col <= d.max; });
      if (covering) {
        defs = defs.filter(function (d) { return d !== covering; });
        if (covering.min < w.col) defs.push({ min: covering.min, max: w.col - 1, tag: setAttr(setAttr(covering.tag, 'min', covering.min), 'max', w.col - 1) });
        if (covering.max > w.col) defs.push({ min: w.col + 1, max: covering.max, tag: setAttr(setAttr(covering.tag, 'min', w.col + 1), 'max', covering.max) });
      }
      let tag = source ? removeAttr(removeAttr(source.tag, 'hidden'), 'outlineLevel') : '<col min="1" max="1" width="50" customWidth="1"/>';
      tag = setAttr(setAttr(tag, 'min', w.col), 'max', w.col);
      defs.push({ min: w.col, max: w.col, tag: tag });
    });
    defs.sort(function (a, b) { return a.min - b.min; });
    const cols = '<cols>' + defs.map(function (d) { return d.tag; }).join('') + '</cols>';
    if (colsMatch) return sheetXml.replace(colsMatch[0], cols);
    return sheetXml.replace(/<sheetData\b/, cols + '<sheetData');
  }

  function updateDimension(sheetXml, rows, maxCol) {
    const maxRow = rows.length ? rows[rows.length - 1].r : 1;
    return sheetXml.replace(/<dimension\b[^>]*ref="([^"]*)"[^>]*\/>/, function (tag, ref) {
      const start = ref.split(':')[0];
      return tag.replace(ref, start + ':' + colLetters(maxCol) + maxRow);
    });
  }

  function updateSheet(sheetXml, payload) {
    const parsed = parseRows(sheetXml);
    const rows = parsed.rows;
    let nextCol = maxUsedColumn(sheetXml, rows) + 1;
    let maxCol = nextCol - 1;
    const widths = [];
    payload.columns.forEach(function (column) {
      const isNew = !column.letter;
      const col = isNew ? nextCol++ : colIndex(column.letter);
      const from = colIndex(column.styleFrom);
      if (col > maxCol) maxCol = col;
      if (isNew) widths.push({ col: col, from: from });
      if (isNew && payload.headerRow) setCell(rows, payload.headerRow, col, column.header, from, false);
      column.cells.forEach(function (cell) {
        // An existing column is overwritten (also emptied); a new column only gets texts
        setCell(rows, cell.row, col, cell.value, from, !isNew);
      });
    });
    let xml = parsed.before + serializeRows(rows) + parsed.after;
    xml = addColumnWidths(xml, widths);
    return updateDimension(xml, rows, maxCol);
  }

  // ---------------------------------------------------------------- status sheet

  function statusSheetXml(table) {
    const headers = table.headers;
    const widths = headers.map(function (h) {
      let max = String(h).length;
      table.rows.forEach(function (r) { max = Math.max(max, String(r[h] === undefined ? '' : r[h]).length); });
      return Math.min(Math.max(max + 3, 10), 60);
    });
    const line = function (r, values) {
      return '<row r="' + r + '">' + values.map(function (v, i) {
        const ref = colLetters(i + 1) + r;
        return typeof v === 'number' ? '<c r="' + ref + '"><v>' + v + '</v></c>' : cellXml(ref, null, v === undefined ? '' : v);
      }).join('') + '</row>';
    };
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      '<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>' +
      '<cols>' + widths.map(function (w, i) { return '<col min="' + (i + 1) + '" max="' + (i + 1) + '" width="' + w + '" customWidth="1"/>'; }).join('') + '</cols>' +
      '<sheetData>' + line(1, headers) + table.rows.map(function (r, i) {
        return line(i + 2, headers.map(function (h) { return r[h]; }));
      }).join('') + '</sheetData>' +
      '<autoFilter ref="A1:' + colLetters(headers.length) + (table.rows.length + 1) + '"/>' +
      '</worksheet>';
  }

  function sheetName(name) {
    const base = String(name || 'Sheet').replace(/[\\/?*[\]:]/g, '_').slice(0, 31) || 'Sheet';
    return base;
  }

  /**
   * Writes the status table to its own sheet: an existing sheet of that name (from an earlier
   * update) is replaced, otherwise a new sheet is added at the end.
   */
  function writeStatusSheet(pkg, info, table) {
    const name = sheetName(table.name);
    const xml = statusSheetXml(table);
    const existing = info.sheets.find(function (s) { return s.name && s.name.toLowerCase() === name.toLowerCase(); });
    if (existing && existing.path) {
      pkg.setText(existing.path, xml);
      return;
    }
    let n = 1;
    while (pkg.has('xl/worksheets/sheet' + n + '.xml')) n++;
    const path = 'xl/worksheets/sheet' + n + '.xml';
    pkg.setText(path, xml);

    const ids = (info.rels.match(/\sId="([^"]*)"/g) || []).map(function (s) { return s.slice(5, -1); });
    let rid = 1;
    while (ids.indexOf('rId' + rid) !== -1) rid++;
    pkg.setText('xl/_rels/workbook.xml.rels', info.rels.replace('</Relationships>',
      '<Relationship Id="rId' + rid + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet' + n + '.xml"/></Relationships>'));

    const sheetId = info.sheets.reduce(function (max, s) { return Math.max(max, s.sheetId); }, 0) + 1;
    const prefix = info.sheets.length ? info.sheets[0].relPrefix : 'r';
    pkg.setText('xl/workbook.xml', info.workbook.replace('</sheets>',
      '<sheet name="' + esc(name) + '" sheetId="' + sheetId + '" ' + prefix + ':id="rId' + rid + '"/></sheets>'));

    const types = pkg.getText('[Content_Types].xml');
    pkg.setText('[Content_Types].xml', types.replace('</Types>',
      '<Override PartName="/' + path + '" ContentType="' + WORKSHEET_TYPE + '"/></Types>'));
  }

  /**
   * Returns the bytes of the updated workbook.
   */
  function updateWorkbook(payload) {
    const pkg = Ooxml.openPackage(payload.original);
    const info = readWorkbook(pkg);
    const sheet = info.sheets.find(function (s) { return s.name === payload.sheetName; }) || (payload.sheetName ? null : info.sheets[0]);
    if (!sheet || !sheet.path || !pkg.has(sheet.path)) {
      throw updateError('MISSING_SHEET', 'Sheet "' + payload.sheetName + '" was not found in the workbook.', { SHEET: payload.sheetName });
    }
    pkg.setText(sheet.path, updateSheet(pkg.getText(sheet.path), payload));
    if (payload.statusTable) writeStatusSheet(pkg, info, payload.statusTable);
    return pkg.toBytes();
  }

  const adapter = {
    id: 'xlsx-update',
    label: 'Excel – original workbook',
    group: 'workbook-update',
    export: function (payload) {
      const extension = payload.extension === '.xlsm' ? '.xlsm' : '.xlsx';
      return { content: updateWorkbook(payload), mimeType: MIME[extension], extension: extension };
    }
  };

  if (Exporters) Exporters.register(adapter);
  adapter.updateWorkbook = updateWorkbook;
  adapter.colLetters = colLetters;
  return adapter;
}));
