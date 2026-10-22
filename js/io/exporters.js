/**
 * Exporters - registry of export adapters.
 *
 * The application prepares format-independent payloads and an export adapter turns a payload
 * into a file (or opens a print window). New output formats are added by registering another
 * adapter; the application picks adapters by group. See docs/ADAPTERS.md.
 *
 * Groups and their payloads:
 *   'table'    { tables: [{ name, headers: [...], rows: [{ [header]: value }] }] }
 *              – full workbook, new version only, partial workbook of one language
 *   'document' document model (ComparatorEngine.buildDocumentModel) – new version of one language
 *   'summary'  change summary model (ComparatorEngine.buildChangeSummary) – one language
 *
 * Adapter interface:
 *   {
 *     id: 'xlsx', label: 'Excel (.xlsx)', group: 'table',
 *     isDefault: true,                                  (optional; else the first registered)
 *     export(payload, { baseName, uiLang }) ->
 *        { content: string | ArrayBuffer | Uint8Array, mimeType, extension }   – a file to download
 *        | { action: 'print', window }                                         – handled by the adapter
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
    root.Exporters = factory();
  }
}(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  const adapters = [];

  function register(adapter) {
    if (!adapter || !adapter.id || !adapter.group || typeof adapter.export !== 'function') {
      throw new Error('Export adapter needs an id, a group and an export() function.');
    }
    unregister(adapter.id);
    adapters.push(adapter);
    return adapter;
  }

  function unregister(id) {
    const pos = adapters.findIndex(function (a) { return a.id === id; });
    if (pos !== -1) adapters.splice(pos, 1);
  }

  function list(group) {
    return adapters.filter(function (a) { return !group || a.group === group; });
  }

  function get(id) {
    return adapters.find(function (a) { return a.id === id; }) || null;
  }

  function getDefault(group) {
    const inGroup = list(group);
    return inGroup.find(function (a) { return a.isDefault; }) || inGroup[0] || null;
  }

  /**
   * File-name-safe version of a text (letters of any script, digits, "_" and "-").
   */
  function safeFilePart(text) {
    return String(text || '').replace(/[^\p{L}\p{N}_-]+/gu, '_').replace(/^_+|_+$/g, '') || 'export';
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

  /**
   * Runs an export adapter. A returned file is downloaded as `<baseName><extension>`.
   * Returns { adapter, fileName } for downloads, or { adapter, action } for other results.
   */
  function run(id, payload, options) {
    const adapter = get(id);
    if (!adapter) throw new Error('Unknown export format: ' + id);
    const opts = options || {};
    const result = adapter.export(payload, opts) || {};
    if (result.content !== undefined) {
      const fileName = safeFilePart(opts.baseName || 'export') + (result.extension || '');
      downloadBlob(new Blob([result.content], { type: result.mimeType || 'application/octet-stream' }), fileName);
      return { adapter: adapter, fileName: fileName };
    }
    return { adapter: adapter, action: result.action || null };
  }

  return {
    register: register,
    unregister: unregister,
    list: list,
    get: get,
    getDefault: getDefault,
    run: run,
    safeFilePart: safeFilePart,
    downloadBlob: downloadBlob,
    openPrintWindow: openPrintWindow
  };
}));
