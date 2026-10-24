# Import and export adapters

Internationalisation Manager reads and writes files only through **adapters**. The application
(`js/app/*.js`) and the engine (`js/core/comparator-engine.js`) never call a file-format library
directly, so a new format is added by writing one adapter file and listing it in `index.html` –
no other code has to change.

```
            file ──► Importers ──► table model ──► app / engine ──► payload ──► Exporters ──► file / print
                     (registry)                                                (registry)
                        ▲                                                         ▲
                 excel-importer.js                                  excel-exporter.js (table)
                 (your adapter)                                     html-document-exporter.js (document)
                                                                    summary-exporters.js (summary)
                                                                    (your adapter)
```

All modules are plain browser scripts (no build step) that also work as Node modules for the
tests. Each adapter file registers itself when it is loaded; load it after `js/io/importers.js` or
`js/io/exporters.js` and before `js/app/core.js`.

## Import adapters

Registry: `js/io/importers.js` (global `Importers`). Built-in importers:
- `js/io/excel-importer.js` (`.xlsx`, `.xlsm`, `.xls`, `.ods`, `.csv`, via SheetJS)
- `js/io/word-importer.js` (`.docx`, `.doc` – Word 97-2003, RTF or HTML saved as .doc –, `.rtf`)
- `js/io/text-importer.js` (`.txt`)

```js
Importers.register({
  id: 'json-bundle',                 // unique id
  label: 'JSON resource bundle',
  extensions: ['.json'],             // file picker filter + matching by extension
  // canImport(fileName, mimeType) { return ...; }   optional, default: extension match
  readAs: 'text',                    // 'text' (UTF-8) or 'arraybuffer' (default); may be a function
                                     // for text in any encoding read bytes and use Importers.decodeText()
  parse(content, { fileName }) {
    const data = JSON.parse(content);
    return {
      sheetNames: ['Strings'],       // one or more tables in the file
      readSheet(name, { hasHeaderRow }) {
        return {
          headers: ['key', 'en-old', 'en-new', 'cs'],
          rows: data.map(d => ({ key: d.key, 'en-old': d.old, 'en-new': d.new, cs: d.cs })),
          rowNumbers: data.map((d, i) => i + 1)   // shown to the user as "Row N"; optional
          // columnLetters: { header: 'C' }       optional, names for columns without a header
        };
      }
    };
  }
});
```

Rules:

- One row = one text unit (paragraph, string). Columns hold the versions/languages; the user
  then chooses the reference columns, the languages and their new-version columns.
- `Importers.readSheet()` normalises every table: all rows get all headers, text is trimmed and
  missing `rowNumbers` are generated – adapters do not need to do this themselves.
- Adapters registered later take precedence for the same extension.
- Throwing from `parse()` shows the error message to the user.

## Export adapters

Registry: `js/io/exporters.js` (global `Exporters`). An adapter belongs to a **group**; the
engine prepares a format-independent payload for each group:

| Group | Payload | Built by | Default adapters |
| --- | --- | --- | --- |
| `table` | `{ tables: [{ name, headers, rows }] }` | `ComparatorEngine.prepareExportTables()` | `xlsx` (`excel-exporter.js`) |
| `document` | document model | `ComparatorEngine.buildDocumentModel()` | `html-print`, `html-file` (`html-document-exporter.js`), `docx` (`docx-exporter.js`) |
| `workbook-update` | `{ original, sheetName, headerRow, extension, columns, statusTable }` | `ComparatorEngine.prepareWorkbookUpdate()` + the original file | `xlsx-update` (`xlsx-update-exporter.js`) |
| `summary` | change summary model | `ComparatorEngine.buildChangeSummary()` | `summary-markdown`, `summary-word`, `summary-print` (`summary-exporters.js`) |

```js
Exporters.register({
  id: 'csv',
  label: 'CSV (.csv)',
  group: 'table',
  // isDefault: true,               // optional – otherwise the first registered adapter is the default
  export(payload, { baseName, uiLang }) {
    const t = payload.tables[0];
    const lines = [t.headers.join(';')].concat(t.rows.map(r => t.headers.map(h => r[h]).join(';')));
    return { content: '﻿' + lines.join('\r\n'), mimeType: 'text/csv;charset=utf-8', extension: '.csv' };
  }
});
```

- Return `{ content, mimeType, extension }` and the registry downloads it as
  `<baseName><extension>`; or perform the action yourself (e.g. open a print window with
  `Exporters.openPrintWindow(html)`) and return `{ action: 'print' }`.
- When more than one `table` adapter is registered, Step 4 shows a **table format** selector;
  the default adapter stays pre-selected.
- Exporters that write into the **original file** receive its bytes: `docx` as `options.original`,
  `xlsx-update` as `payload.original`. `Importers.importFile()` resolves with the file `content`,
  which the application keeps in memory (and asks for again after a reload). `js/io/ooxml.js`
  opens and rewrites Office packages and defines the paragraphs of a Word document, shared by
  the Word importer and the `docx` exporter so that paragraph N is the same in both.
- Document and summary models contain plain data (texts, numbers, row numbers, diff results) and
  `uiLang` for labels – no HTML. See `buildDocumentModel` / `buildChangeSummary` in the engine
  for their fields, and `html-document-exporter.js` / `summary-exporters.js` as examples.

## Tests

`npm test` runs the engine/adapter tests in Node and a jsdom GUI battery. The GUI battery loads the
scripts in the order `index.html` lists them, so a new adapter file is included automatically.
`test/gui-battery.test.js` contains a test that registers a TSV importer and a CSV exporter at
runtime – a template for testing your own adapter.

## Ideas for further generalisation

- **Key-based resource formats** (JSON/YAML bundles, Gettext `.po`, XLIFF, Android/iOS strings,
  `.resx`, `.properties`): import adapters producing one row per key; the key could be shown
  instead of the row number (an optional *identifier column* in the table model).
- **Round trip**: export adapters writing the same resource format back (with the new version as
  the target), so the tool can sit inside a software localisation workflow.
- **Diff granularity per language**: word diff does not work for scripts without spaces
  (Chinese, Japanese, Thai) – character-level diff and right-to-left text direction per language.
- **Quality checks**: placeholders (`{0}`, `%s`, ICU messages), HTML tags, numbers and length
  limits compared between reference and translation before a unit can be approved.
- **Terminology**: a glossary of approved terms per language used in prompts and checks.
- **Translation providers** as optional adapters (an LLM or MT API) – off by default to keep the
  application fully offline.
- **Storage**: IndexedDB instead of LocalStorage for large projects; the project file itself could
  become an import/export adapter.
