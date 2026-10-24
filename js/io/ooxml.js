/**
 * Ooxml - helpers for Office Open XML packages (.docx, .xlsx), shared by the Word importer and
 * the exporters that write into an original file.
 *
 *   openPackage(bytes)        – zip package: get / getText / set / setText / has / paths / toBytes
 *   escapeXml(text)           – text for XML content and attributes (invalid characters removed)
 *   docxParagraphs(xml)       – the text paragraphs of word/document.xml in document order
 *   docxTextEdits(...)        – edits replacing the text of one paragraph, keeping its formatting
 *   applyEdits(xml, edits)    – applies edits of several paragraphs at once
 *
 * The paragraph scan is the single definition of "paragraph N" for both import and export, so a
 * translated text is always written back to the paragraph it was read from.
 *
 * Universal module: works in browser and Node.js.
 */
(function (root, factory) {
  if (typeof define === 'function' && define.amd) {
    define(['xlsx'], factory);
  } else if (typeof module === 'object' && module.exports) {
    module.exports = factory(require('xlsx'));
  } else {
    root.Ooxml = factory(root.XLSX);
  }
}(typeof self !== 'undefined' ? self : this, function (XLSX) {
  'use strict';

  function utf8Decode(bytes) {
    return new TextDecoder('utf-8').decode(bytes);
  }

  function utf8Encode(text) {
    return new TextEncoder().encode(text);
  }

  /**
   * Opens a zip package (ArrayBuffer / Uint8Array). Paths are given without a leading slash,
   * e.g. 'xl/workbook.xml'. Parts that are not changed are written back unchanged.
   */
  function openPackage(bytes) {
    if (!XLSX || !XLSX.CFB) throw new Error('The zip reader (SheetJS CFB) is not loaded.');
    const data = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
    if (data.length < 4 || data[0] !== 0x50 || data[1] !== 0x4B) {
      const err = new Error('The file is not an Office Open XML package.');
      err.code = 'NOT_OOXML';
      throw err;
    }
    const cfb = XLSX.CFB.read(data, { type: 'array' });
    const entry = function (path) { return XLSX.CFB.find(cfb, '/' + path); };
    return {
      has: function (path) { return !!entry(path); },
      get: function (path) {
        const e = entry(path);
        return e && e.content ? new Uint8Array(e.content) : null;
      },
      getText: function (path) {
        const e = entry(path);
        return e && e.content ? utf8Decode(new Uint8Array(e.content)) : null;
      },
      set: function (path, content) {
        const e = entry(path);
        if (e) e.content = content;
        else XLSX.CFB.utils.cfb_add(cfb, '/' + path, content);
      },
      setText: function (path, text) {
        this.set(path, utf8Encode(text));
      },
      paths: function () {
        return cfb.FullPaths.map(function (p) { return p.replace(/^[^/]*\//, ''); })
          .filter(function (p) { return p && !/\/$/.test(p) && p.charAt(0) !== '\u0001'; });
      },
      toBytes: function () {
        return XLSX.CFB.write(cfb, { type: 'array', fileType: 'zip', compression: true });
      }
    };
  }

  // Characters that XML 1.0 does not allow are dropped (they cannot be stored in a document)
  function escapeXml(text) {
    return String(text === undefined || text === null ? '' : text)
      // eslint-disable-next-line no-control-regex -- removing control characters is the point
      .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]/g, '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }

  // Single pass, so an escaped entity ("&amp;lt;") is not decoded twice
  function decodeEntities(str) {
    if (!str) return '';
    const named = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" };
    return str.replace(/&(#x[0-9a-fA-F]+|#\d+|[a-zA-Z]+);/g, function (m, name) {
      if (name.charAt(0) === '#') {
        const code = name.charAt(1) === 'x' || name.charAt(1) === 'X' ? parseInt(name.slice(2), 16) : parseInt(name.slice(1), 10);
        try { return String.fromCodePoint(code); } catch (e) { return m; }
      }
      return Object.prototype.hasOwnProperty.call(named, name) ? named[name] : m;
    });
  }

  // Text-bearing elements of a paragraph (page and column breaks are layout, not text)
  const TEXT_TOKEN = /<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>|<w:t(?:\s[^>]*)?\/>|<w:tab(?:\s[^>]*)?\/>|<w:(?:br|cr)(?:\s[^>]*)?\/>|<w:noBreakHyphen\/>/g;

  function isLayoutBreak(token) {
    return /^<w:br\b/.test(token) && /w:type="(?:page|column)"/.test(token);
  }

  // Ranges [start, end) that are not document content: the paragraph properties (they hold tab
  // stop definitions) and the fallback copies of text boxes (the same text is in mc:Choice)
  function maskedRanges(xml) {
    const ranges = [];
    const re = /<mc:Fallback\b[\s\S]*?<\/mc:Fallback>|<w:pPr\b[\s\S]*?<\/w:pPr>/g;
    let m;
    while ((m = re.exec(xml)) !== null) ranges.push([m.index, m.index + m[0].length]);
    return ranges;
  }

  function inRanges(pos, ranges) {
    return ranges.some(function (r) { return pos >= r[0] && pos < r[1]; });
  }

  /**
   * Paragraphs of word/document.xml in document order. Paragraphs nested in text boxes are
   * separate paragraphs; the text of a paragraph never includes nested paragraphs.
   * Returns [{ start, end, segments: [[from, to]], text }] – only paragraphs with text.
   */
  function docxParagraphs(xml) {
    const masked = maskedRanges(xml);
    const fallback = masked.filter(function (r) { return /^<mc:Fallback/.test(xml.slice(r[0], r[0] + 12)); });
    const all = [];
    const stack = [];
    const re = /<w:p(?=[\s>/])[^>]*?(\/?)>|<\/w:p>/g;
    let m;
    while ((m = re.exec(xml)) !== null) {
      if (m[0] === '</w:p>') {
        const p = stack.pop();
        if (p) { p.innerEnd = m.index; p.end = m.index + m[0].length; }
        continue;
      }
      const p = { start: m.index, innerStart: m.index + m[0].length, children: [] };
      if (stack.length > 0) stack[stack.length - 1].children.push(p);
      all.push(p);
      if (m[1] === '/') { p.innerEnd = p.innerStart; p.end = p.innerStart; } else stack.push(p);
    }

    const result = [];
    all.forEach(function (p) {
      if (p.end === undefined || inRanges(p.start, fallback)) return;
      // Own content: the inner XML without nested paragraphs
      const segments = [];
      let from = p.innerStart;
      p.children.forEach(function (c) {
        if (c.start > from) segments.push([from, c.start]);
        from = Math.max(from, c.end || c.start);
      });
      if (p.innerEnd > from) segments.push([from, p.innerEnd]);

      let text = '';
      segments.forEach(function (seg) {
        const part = xml.slice(seg[0], seg[1]);
        TEXT_TOKEN.lastIndex = 0;
        let t;
        while ((t = TEXT_TOKEN.exec(part)) !== null) {
          if (inRanges(seg[0] + t.index, masked)) continue;
          if (t[1] !== undefined) text += decodeEntities(t[1]);
          else if (/^<w:tab/.test(t[0])) text += '\t';
          else if (/^<w:noBreakHyphen/.test(t[0])) text += '-';
          else if (/^<w:t\b/.test(t[0])) continue;
          else if (!isLayoutBreak(t[0])) text += '\n';
        }
      });
      text = text.trim();
      if (text) result.push({ start: p.start, end: p.end, segments: segments, text: text });
    });
    return result;
  }

  function textRunXml(text) {
    // Line breaks and tabs of the new text become <w:br/> and <w:tab/> inside the same run
    return String(text).split('\n').map(function (line) {
      return line.split('\t').map(function (part) {
        return '<w:t xml:space="preserve">' + escapeXml(part) + '</w:t>';
      }).join('<w:tab/>');
    }).join('<w:br/>');
  }

  /**
   * Edits that replace the text of a paragraph found by docxParagraphs. The new text goes into
   * the first text element outside a field result, so the paragraph keeps its style, numbering
   * and the formatting of its first run; the other text of the paragraph is removed. Pictures,
   * page breaks, field codes and nested text boxes are kept. Returns [[start, end, replacement]].
   */
  function docxTextEdits(xml, paragraph, text) {
    const masked = maskedRanges(xml);
    const tokens = [];
    paragraph.segments.forEach(function (seg) {
      const re = /<w:t(?:\s[^>]*)?>[\s\S]*?<\/w:t>|<w:t(?:\s[^>]*)?\/>|<w:tab(?:\s[^>]*)?\/>|<w:(?:br|cr)(?:\s[^>]*)?\/>|<w:noBreakHyphen\/>|<w:fldChar\b[^>]*>|<w:fldSimple\b[^>]*?(\/?)>|<\/w:fldSimple>/g;
      const part = xml.slice(seg[0], seg[1]);
      let t;
      while ((t = re.exec(part)) !== null) {
        const pos = seg[0] + t.index;
        if (!inRanges(pos, masked)) tokens.push({ pos: pos, end: pos + t[0].length, xml: t[0] });
      }
    });

    // Field results (between "separate" and "end", or inside fldSimple) are generated by Word
    const fieldState = [];
    let simpleDepth = 0;
    let target = null;
    let firstText = null;
    tokens.forEach(function (tok) {
      const fld = /w:fldCharType="(begin|separate|end)"/.exec(tok.xml);
      if (/^<w:fldChar/.test(tok.xml) && fld) {
        if (fld[1] === 'begin') fieldState.push('code');
        else if (fld[1] === 'separate' && fieldState.length) fieldState[fieldState.length - 1] = 'result';
        else if (fld[1] === 'end') fieldState.pop();
        return;
      }
      if (/^<w:fldSimple/.test(tok.xml)) { if (!/\/>$/.test(tok.xml)) simpleDepth++; return; }
      if (tok.xml === '</w:fldSimple>') { simpleDepth = Math.max(0, simpleDepth - 1); return; }
      tok.isText = /^<w:t\b/.test(tok.xml);
      tok.removable = tok.isText || !isLayoutBreak(tok.xml);
      const inResult = simpleDepth > 0 || fieldState.indexOf('result') !== -1;
      const inCode = fieldState.length > 0 && fieldState[fieldState.length - 1] === 'code';
      if (inCode) tok.removable = false;
      if (tok.isText && !inCode) {
        if (!firstText) firstText = tok;
        if (!target && !inResult) target = tok;
      }
    });
    target = target || firstText;
    if (!target) return [];

    const edits = [];
    tokens.forEach(function (tok) {
      if (tok === target) edits.push([tok.pos, tok.end, textRunXml(text)]);
      else if (tok.removable && tok.isText !== undefined) edits.push([tok.pos, tok.end, '']);
    });
    return edits;
  }

  /**
   * Applies edits [[start, end, replacement]] computed on the same XML (they must not overlap).
   */
  function applyEdits(xml, edits) {
    const sorted = edits.slice().sort(function (a, b) { return b[0] - a[0]; });
    let out = xml;
    sorted.forEach(function (e) { out = out.slice(0, e[0]) + e[2] + out.slice(e[1]); });
    return out;
  }

  function setDocxParagraphText(xml, paragraph, text) {
    return applyEdits(xml, docxTextEdits(xml, paragraph, text));
  }

  return {
    openPackage: openPackage,
    escapeXml: escapeXml,
    decodeEntities: decodeEntities,
    docxParagraphs: docxParagraphs,
    docxTextEdits: docxTextEdits,
    applyEdits: applyEdits,
    setDocxParagraphText: setDocxParagraphText,
    utf8Decode: utf8Decode,
    utf8Encode: utf8Encode
  };
}));
