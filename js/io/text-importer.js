/**
 * Text file import adapter (.txt).
 * Extracts paragraphs (separated by blank lines or newlines) as rows of the reference column.
 * Registers itself in the Importers registry. Universal module: works in browser and Node.js.
 */
(function (root, factory) {
  if (typeof define === 'function' && define.amd) {
    define(['importers'], factory);
  } else if (typeof module === 'object' && module.exports) {
    module.exports = factory(require('./importers'));
  } else {
    root.TextImporter = factory(root.Importers);
  }
}(typeof self !== 'undefined' ? self : this, function (Importers) {
  'use strict';

  const EXTENSIONS = ['.txt'];

  function parseText(content) {
    const str = String(content || '')
      .replace(/^\uFEFF/, '')
      .replace(/\r\n/g, '\n')
      .replace(/\r/g, '\n')
      .trim();

    if (!str) return [];

    let paragraphs;
    if (/\n\s*\n/.test(str)) {
      paragraphs = str.split(/\n\s*(?:\n\s*)+/).map(function (p) {
        return p.trim();
      }).filter(Boolean);
    } else {
      paragraphs = str.split(/\n+/).map(function (p) {
        return p.trim();
      }).filter(Boolean);
    }
    return paragraphs;
  }

  const adapter = {
    id: 'text',
    label: 'Text (.txt)',
    extensions: EXTENSIONS,
    readAs: 'text',
    canImport: function (fileName, mimeType) {
      return EXTENSIONS.indexOf(Importers.extensionOf(fileName)) !== -1 || mimeType === 'text/plain';
    },
    parse: function (content, options) {
      const paragraphs = parseText(content);
      if (!paragraphs || paragraphs.length === 0) {
        const err = new Error('The text file is empty.');
        err.code = 'EMPTY_FILE';
        throw err;
      }

      return {
        fileName: options && options.fileName,
        adapterId: 'text',
        sheetNames: ['Text'],
        readSheet: function () {
          return {
            headers: ['Reference'],
            rows: paragraphs.map(function (p) { return { 'Reference': p }; }),
            rowNumbers: paragraphs.map(function (_, idx) { return idx + 1; }),
            columnLetters: { 'Reference': 'A' }
          };
        }
      };
    }
  };

  if (Importers) Importers.register(adapter);
  return adapter;
}));
