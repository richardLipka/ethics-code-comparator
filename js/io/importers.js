/**
 * Importers - registry of import adapters.
 *
 * An import adapter turns a file into one or more tables in the common table model used by the
 * application, so new formats (e.g. JSON resource bundles, XLIFF, .po) can be added without
 * touching the application controller. See docs/ADAPTERS.md.
 *
 * Adapter interface:
 *   {
 *     id: 'excel',                         // unique id
 *     label: 'Excel / CSV',                // shown to the user
 *     extensions: ['.xlsx', '.csv'],       // used for the file picker and for matching
 *     canImport(fileName, mimeType) -> boolean            (optional, defaults to extension match)
 *     readAs(fileName, mimeType) -> 'arraybuffer' | 'text' (optional, default 'arraybuffer')
 *     parse(content, { fileName }) -> source
 *   }
 *
 * Source (one imported file):
 *   {
 *     sheetNames: ['Sheet1', ...],         // tables in the file (at least one)
 *     readSheet(name, { hasHeaderRow }) -> table | null   (null = empty table)
 *   }
 *
 * Table model:
 *   {
 *     headers: ['Column name', ...],       // column keys in display order
 *     rows: [{ [header]: value }, ...],    // one row per paragraph / text unit
 *     rowNumbers: [2, 3, ...],             // position of each row in the source (shown to the user)
 *     columnLetters: { [header]: 'C' }     // optional, for columns without a name
 *   }
 *
 * Universal module: works in browser and Node.js.
 */
(function (root, factory) {
  if (typeof define === 'function' && define.amd) {
    define([], factory);
  } else if (typeof module === 'object' && module.exports) {
    module.exports = factory();
  } else {
    root.Importers = factory();
  }
}(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  const adapters = [];

  function extensionOf(fileName) {
    const m = String(fileName || '').toLowerCase().match(/\.[^./\\]+$/);
    return m ? m[0] : '';
  }

  /**
   * Registers an adapter (an adapter with the same id is replaced). Adapters registered later
   * are tried first, so a specialised adapter can take over an extension from a default one.
   */
  function register(adapter) {
    if (!adapter || !adapter.id || typeof adapter.parse !== 'function') {
      throw new Error('Import adapter needs an id and a parse() function.');
    }
    unregister(adapter.id);
    adapters.unshift(adapter);
    return adapter;
  }

  function unregister(id) {
    const pos = adapters.findIndex(function (a) { return a.id === id; });
    if (pos !== -1) adapters.splice(pos, 1);
  }

  function list() {
    return adapters.slice();
  }

  function get(id) {
    return adapters.find(function (a) { return a.id === id; }) || null;
  }

  function canImport(adapter, fileName, mimeType) {
    if (typeof adapter.canImport === 'function') return !!adapter.canImport(fileName, mimeType);
    return (adapter.extensions || []).indexOf(extensionOf(fileName)) !== -1;
  }

  function findForFile(fileName, mimeType) {
    return adapters.find(function (a) { return canImport(a, fileName, mimeType); }) || null;
  }

  /**
   * Value for the `accept` attribute of a file input: all extensions of all adapters.
   */
  function acceptAttribute() {
    const all = [];
    adapters.forEach(function (a) {
      (a.extensions || []).forEach(function (ext) {
        if (all.indexOf(ext) === -1) all.push(ext);
      });
    });
    return all.join(',');
  }

  // Cells often carry stray spaces or line breaks around the text (a cell may even look empty in
  // a narrow preview), so text is trimmed for every importer
  function trimCell(value) {
    return typeof value === 'string' ? value.trim() : (value === undefined || value === null ? '' : value);
  }

  /**
   * Brings a table returned by an adapter into the common model: every row has every header,
   * strings are trimmed and every row has a row number.
   */
  function normalizeTable(table) {
    if (!table || !Array.isArray(table.rows) || table.rows.length === 0) return null;
    const headers = Array.isArray(table.headers) && table.headers.length > 0
      ? table.headers.map(String)
      : Object.keys(table.rows[0]);
    const rowNumbers = Array.isArray(table.rowNumbers) && table.rowNumbers.length === table.rows.length
      ? table.rowNumbers.slice()
      : table.rows.map(function (r, idx) { return idx + 2; });
    return {
      headers: headers,
      rows: table.rows.map(function (r) {
        const row = {};
        headers.forEach(function (h) { row[h] = trimCell(r[h]); });
        return row;
      }),
      rowNumbers: rowNumbers,
      columnLetters: table.columnLetters || {},
      headerRowNumber: typeof table.headerRowNumber === 'number' ? table.headerRowNumber : null
    };
  }

  /**
   * Reads one table of an imported source and normalises it (null when the table is empty).
   */
  function readSheet(source, name, options) {
    return normalizeTable(source.readSheet(name, options || {}));
  }

  /**
   * Decodes the bytes of a text file (CSV, TXT). A UTF-8 / UTF-16 byte order mark wins; otherwise
   * strict UTF-8 is tried and, when the bytes are not valid UTF-8 (files saved by Excel or
   * Notepad in a legacy code page), the text is decoded as Windows-1250 (Central European),
   * falling back to Windows-1252. A string is returned unchanged (without a BOM).
   */
  function decodeText(content) {
    if (typeof content === 'string') return content.replace(/^\uFEFF/, '');
    const bytes = content instanceof Uint8Array ? content : new Uint8Array(content || []);
    if (bytes.length >= 2 && bytes[0] === 0xFF && bytes[1] === 0xFE) return decodeWith('utf-16le', bytes.subarray(2));
    if (bytes.length >= 2 && bytes[0] === 0xFE && bytes[1] === 0xFF) return decodeWith('utf-16be', bytes.subarray(2));
    const body = bytes.length >= 3 && bytes[0] === 0xEF && bytes[1] === 0xBB && bytes[2] === 0xBF ? bytes.subarray(3) : bytes;
    try {
      return new TextDecoder('utf-8', { fatal: true }).decode(body);
    } catch (e) {
      return decodeWith('windows-1250', body) || decodeWith('windows-1252', body);
    }
  }

  function decodeWith(encoding, bytes) {
    try {
      return new TextDecoder(encoding).decode(bytes);
    } catch (e) {
      return '';
    }
  }

  function readFileContent(file, mode) {
    return new Promise(function (resolve, reject) {
      const reader = new FileReader();
      reader.onload = function () { resolve(reader.result); };
      reader.onerror = function () { reject(reader.error || new Error(file.name)); };
      if (mode === 'text') {
        reader.readAsText(file, 'UTF-8');
      } else {
        reader.readAsArrayBuffer(file);
      }
    });
  }

  /**
   * Imports a browser File with the matching adapter.
   * Resolves to { adapter, source, content } (content = what was read from the file, so an
   * exporter can later write into a copy of it); rejects when no adapter accepts the file or
   * parsing fails.
   */
  function importFile(file) {
    const adapter = findForFile(file.name, file.type);
    if (!adapter) {
      const err = new Error('Unsupported file type: ' + (extensionOf(file.name) || file.name));
      err.code = 'UNSUPPORTED';
      return Promise.reject(err);
    }
    const mode = typeof adapter.readAs === 'function' ? adapter.readAs(file.name, file.type) : (adapter.readAs || 'arraybuffer');
    return readFileContent(file, mode).then(function (content) {
      const source = adapter.parse(content, { fileName: file.name, mimeType: file.type });
      if (!source || !Array.isArray(source.sheetNames) || source.sheetNames.length === 0) {
        const err = new Error('The file contains no tables.');
        err.code = 'NO_SHEETS';
        throw err;
      }
      return { adapter: adapter, source: source, content: content };
    });
  }

  return {
    register: register,
    unregister: unregister,
    list: list,
    get: get,
    findForFile: findForFile,
    acceptAttribute: acceptAttribute,
    normalizeTable: normalizeTable,
    readSheet: readSheet,
    importFile: importFile,
    trimCell: trimCell,
    decodeText: decodeText,
    extensionOf: extensionOf
  };
}));
