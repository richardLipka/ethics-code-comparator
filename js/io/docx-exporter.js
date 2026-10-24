/**
 * Word (.docx) export adapter for the document model (ComparatorEngine.buildDocumentModel).
 *
 *  - With `options.original` (the bytes of the .docx the text was imported from) the translated
 *    text is written into a copy of that document: every paragraph keeps its style, numbering,
 *    tables, pictures, headers and footers. Paragraphs are matched by their position in the
 *    document (row N = text paragraph N, see Ooxml.docxParagraphs). Paragraphs without a
 *    translation and paragraphs skipped in Step 1 keep their original text.
 *  - Without it a new, plainly formatted document is created (heading from the document
 *    template, one paragraph per row, missing translations marked).
 *
 * Registers itself in the Exporters registry. Universal module: works in browser and Node.js.
 */
(function (root, factory) {
  if (typeof define === 'function' && define.amd) {
    define(['xlsx', 'ooxml', 'exporters'], factory);
  } else if (typeof module === 'object' && module.exports) {
    module.exports = factory(require('xlsx'), require('./ooxml'), require('./exporters'));
  } else {
    root.DocxExporter = factory(root.XLSX, root.Ooxml, root.Exporters);
  }
}(typeof self !== 'undefined' ? self : this, function (XLSX, Ooxml, Exporters) {
  'use strict';

  const MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
  const esc = Ooxml.escapeXml;

  function labels(isEn) {
    return isEn
      ? { untitled: 'Document – new version', missing: 'Translation missing', row: 'row',
          draft: function (m, u) { return 'DRAFT – ' + m + ' paragraph(s) without translation, ' + u + ' not yet verified.'; } }
      : { untitled: 'Dokument – nová verze', missing: 'Překlad chybí', row: 'řádek',
          draft: function (m, u) { return 'KONCEPT – ' + m + ' odstavců bez překladu, ' + u + ' dosud neověřeno.'; } };
  }

  /**
   * Writes the translated paragraphs into a copy of the original document.
   * Returns { bytes, replaced, kept } (kept = paragraphs left in the original language).
   */
  function updateDocument(model, original) {
    const pkg = Ooxml.openPackage(original);
    const xml = pkg.getText('word/document.xml');
    if (xml === null) {
      const err = new Error('The original file has no word/document.xml.');
      err.code = 'NOT_DOCX';
      throw err;
    }
    const paragraphs = Ooxml.docxParagraphs(xml);
    const byRow = {};
    model.paragraphs.forEach(function (p) { byRow[p.rowNumber] = p; });

    let edits = [];
    let replaced = 0;
    let kept = 0;
    paragraphs.forEach(function (para, i) {
      const item = byRow[i + 1];
      if (!item) return;
      if (item.missing || !item.text) {
        kept++;
        return;
      }
      if (item.text === para.text) return;
      edits = edits.concat(Ooxml.docxTextEdits(xml, para, item.text));
      replaced++;
    });
    pkg.setText('word/document.xml', Ooxml.applyEdits(xml, edits));
    return { bytes: pkg.toBytes(), replaced: replaced, kept: kept };
  }

  function run(text, props) {
    return '<w:r>' + (props ? '<w:rPr>' + props + '</w:rPr>' : '') + String(text).split('\n').map(function (line) {
      return '<w:t xml:space="preserve">' + esc(line) + '</w:t>';
    }).join('<w:br/>') + '</w:r>';
  }

  function paragraph(content, style, align) {
    const pPr = (style ? '<w:pStyle w:val="' + style + '"/>' : '') + (align ? '<w:jc w:val="' + align + '"/>' : '');
    return '<w:p>' + (pPr ? '<w:pPr>' + pPr + '</w:pPr>' : '') + content + '</w:p>';
  }

  /**
   * A new document with the heading of the document template and one paragraph per row.
   */
  function createDocument(model) {
    const L = labels(model.uiLang === 'en');
    const accent = (model.accentColor || '#2563eb').replace('#', '').slice(0, 6);
    const body = [];
    if (model.isDraft) body.push(paragraph(run(L.draft(model.missingCount, model.unverifiedCount), '<w:b/><w:color w:val="B91C1C"/>')));
    if (model.organization) body.push(paragraph(run(model.organization, '<w:b/><w:color w:val="' + accent + '"/>')));
    body.push(paragraph(run(model.title || L.untitled), 'Title'));
    if (model.subtitle) body.push(paragraph(run(model.subtitle), 'Subtitle'));
    const meta = [model.versionLabel, model.dateLabel].filter(Boolean).join(' · ');
    if (meta) body.push(paragraph(run(meta, '<w:color w:val="64748B"/>')));
    if (model.introduction) body.push(paragraph(run(model.introduction, '<w:i/>')));
    model.paragraphs.forEach(function (p) {
      body.push(p.missing
        ? paragraph(run('[' + L.missing + ' – ' + L.row + ' ' + p.rowNumber + ']', '<w:i/><w:color w:val="B91C1C"/><w:highlight w:val="yellow"/>'))
        : paragraph(run(p.text)));
    });
    if (model.signatures && model.signatures.length) {
      const cells = model.signatures.map(function (s) {
        return '<w:tc><w:tcPr><w:tcW w:w="4500" w:type="dxa"/></w:tcPr>' +
          paragraph(run('______________________________')) + paragraph(run(s || ''), null, 'left') + '</w:tc>';
      }).join('');
      body.push('<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid>' +
        model.signatures.map(function () { return '<w:gridCol w:w="4500"/>'; }).join('') +
        '</w:tblGrid><w:tr>' + cells + '</w:tr></w:tbl>');
      body.push(paragraph(''));
    }

    const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
    const files = {
      '[Content_Types].xml': '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        '<Default Extension="xml" ContentType="application/xml"/>' +
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
        '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' +
        '</Types>',
      '_rels/.rels': '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
        '</Relationships>',
      'word/_rels/document.xml.rels': '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
        '</Relationships>',
      'word/styles.xml': '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ' + W + '>' +
        '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri" w:eastAsia="Calibri"/>' +
        '<w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="' + (model.uiLang === 'en' ? 'en-GB' : 'cs-CZ') + '"/></w:rPr></w:rPrDefault>' +
        '<w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="276" w:lineRule="auto"/><w:jc w:val="both"/></w:pPr></w:pPrDefault></w:docDefaults>' +
        '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>' +
        '<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>' +
        '<w:pPr><w:spacing w:after="120"/><w:jc w:val="left"/></w:pPr><w:rPr><w:b/><w:color w:val="' + accent + '"/><w:sz w:val="40"/><w:szCs w:val="40"/></w:rPr></w:style>' +
        '<w:style w:type="paragraph" w:styleId="Subtitle"><w:name w:val="Subtitle"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>' +
        '<w:pPr><w:spacing w:after="240"/><w:jc w:val="left"/></w:pPr><w:rPr><w:color w:val="475569"/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr></w:style>' +
        '</w:styles>',
      'word/document.xml': '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ' + W + '><w:body>' +
        body.join('') +
        '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="851" w:bottom="1134" w:left="851" w:header="567" w:footer="567" w:gutter="0"/></w:sectPr>' +
        '</w:body></w:document>'
    };
    const cfb = XLSX.CFB.utils.cfb_new();
    Object.keys(files).forEach(function (path) {
      XLSX.CFB.utils.cfb_add(cfb, '/' + path, Ooxml.utf8Encode(files[path]));
    });
    return XLSX.CFB.write(cfb, { type: 'array', fileType: 'zip', compression: true });
  }

  const adapter = {
    id: 'docx',
    label: 'Word (.docx)',
    group: 'document',
    // options.original: bytes of the .docx the text was imported from (keeps its formatting)
    export: function (model, options) {
      const opts = options || {};
      const content = opts.original ? updateDocument(model, opts.original).bytes : createDocument(model);
      return { content: content, mimeType: MIME, extension: '.docx' };
    }
  };

  if (Exporters) Exporters.register(adapter);
  adapter.updateDocument = updateDocument;
  adapter.createDocument = createDocument;
  return adapter;
}));
