/**
 * Word document import adapter (.docx, .doc, .rtf; a .doc may also be RTF or HTML saved by Word).
 * Extracts text paragraphs as rows of the reference column.
 * Registers itself in the Importers registry. Universal module: works in browser and Node.js.
 */
(function (root, factory) {
  if (typeof define === 'function' && define.amd) {
    define(['xlsx', 'importers'], factory);
  } else if (typeof module === 'object' && module.exports) {
    module.exports = factory(require('xlsx'), require('./importers'));
  } else {
    root.WordImporter = factory(root.XLSX, root.Importers);
  }
}(typeof self !== 'undefined' ? self : this, function (XLSX, Importers) {
  'use strict';

  const EXTENSIONS = ['.docx', '.doc', '.rtf'];

  function decodeUtf8(bytes) {
    if (!bytes || bytes.length === 0) return '';
    return new TextDecoder('utf-8').decode(bytes);
  }

  // Single pass, so an escaped entity ("&amp;lt;") is not decoded twice
  function decodeEntities(str) {
    if (!str) return '';
    const named = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'", nbsp: '\u00A0' };
    return str.replace(/&(#x[0-9a-fA-F]+|#\d+|[a-zA-Z]+);/g, function (m, name) {
      if (name.charAt(0) === '#') {
        const code = name.charAt(1) === 'x' || name.charAt(1) === 'X' ? parseInt(name.slice(2), 16) : parseInt(name.slice(1), 10);
        try { return String.fromCodePoint(code); } catch (e) { return m; }
      }
      return Object.prototype.hasOwnProperty.call(named, name.toLowerCase()) ? named[name.toLowerCase()] : m;
    });
  }

  function parseDocxXml(xml) {
    // Paragraph properties hold tab stop definitions (<w:tab .../>), which are not text
    const pRegex = /<w:p(?:\s[^>]*)?>([\s\S]*?)<\/w:p>/g;
    const tRegex = /<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>|<w:tab(?:\s[^>]*)?\/>|<w:(?:br|cr)(?:\s[^>]*)?\/>|<w:noBreakHyphen\/>/g;
    const paragraphs = [];
    let pMatch;
    while ((pMatch = pRegex.exec(xml)) !== null) {
      const pXml = pMatch[1].replace(/<w:pPr>[\s\S]*?<\/w:pPr>/g, '');
      let pText = '';
      let tMatch;
      while ((tMatch = tRegex.exec(pXml)) !== null) {
        if (tMatch[1] !== undefined) pText += decodeEntities(tMatch[1]);
        else if (/^<w:tab/.test(tMatch[0])) pText += '\t';
        else if (/^<w:noBreakHyphen/.test(tMatch[0])) pText += '-';
        else pText += '\n';
      }
      pText = pText.trim();
      if (pText) paragraphs.push(pText);
    }
    return paragraphs;
  }

  function parseHtmlDoc(html) {
    const cleanHtml = html
      .replace(/<!--[\s\S]*?-->/g, '')
      .replace(/<style[\s\S]*?<\/style>/gi, '')
      .replace(/<script[\s\S]*?<\/script>/gi, '')
      .replace(/<xml[\s\S]*?<\/xml>/gi, '')
      .replace(/<head[\s\S]*?<\/head>/gi, '');

    // Source line breaks are only whitespace in HTML; <br> is a real line break
    function toText(fragment) {
      return decodeEntities(fragment.replace(/<br\s*\/?>/gi, '\u2028').replace(/<[^>]+>/g, ''))
        .replace(/\u00A0/g, ' ')
        .replace(/[ \t\r\n]+/g, ' ')
        .split('\u2028').map(function (line) { return line.trim(); }).join('\n')
        .trim();
    }

    const pRegex = /<(p|li|h[1-6]|td|th)(?:\s[^>]*)?>([\s\S]*?)<\/\1>/gi;
    const paragraphs = [];
    let match;
    while ((match = pRegex.exec(cleanHtml)) !== null) {
      const text = toText(match[2]);
      if (text) paragraphs.push(text);
    }
    if (paragraphs.length === 0) {
      decodeEntities(cleanHtml.replace(/<[^>]+>/g, '\n')).split(/\n+/).forEach(function (line) {
        const t = line.replace(/\u00A0/g, ' ').trim();
        if (t) paragraphs.push(t);
      });
    }
    return paragraphs;
  }

  // Groups whose content is not document text (font table, styles, metadata, pictures, …)
  const RTF_SKIP_DESTINATIONS = ['fonttbl', 'colortbl', 'stylesheet', 'info', 'pict', 'object', 'header', 'footer',
    'headerl', 'headerr', 'headerf', 'footerl', 'footerr', 'footerf', 'listtable', 'listoverridetable', 'rsidtbl',
    'generator', 'xmlnstbl', 'themedata', 'colorschememapping', 'datastore', 'latentstyles', 'filetbl', 'revtbl',
    'pgdsctbl', 'footnote', 'fldinst', 'bkmkstart', 'bkmkend', 'mmathPr', 'listtext', 'pntext'];

  /**
   * Extracts paragraphs from RTF: control words are interpreted (\par, \line, \tab, \'xx in the
   * document's code page, \uN with \ucN fallback skipping) and non-text groups are skipped.
   */
  function parseRtfDoc(rtf) {
    const cpMatch = /\\ansicpg(\d+)/.exec(rtf);
    let decoder = null;
    try { decoder = new TextDecoder('windows-' + (cpMatch ? cpMatch[1] : '1252')); } catch (e) { decoder = null; }
    const paragraphs = [];
    let current = '';
    let pendingBytes = [];
    const stack = [];
    let group = { skip: false, uc: 1 };
    let skipChars = 0;

    function flushBytes() {
      if (pendingBytes.length === 0) return;
      const arr = new Uint8Array(pendingBytes);
      current += decoder ? decoder.decode(arr) : String.fromCharCode.apply(null, pendingBytes);
      pendingBytes = [];
    }
    function emit(text) {
      if (group.skip) return;
      if (skipChars > 0) { skipChars--; return; }
      flushBytes();
      current += text;
    }
    function endParagraph() {
      flushBytes();
      const t = current.replace(/[ \t]+/g, ' ').trim();
      if (t) paragraphs.push(t);
      current = '';
    }

    let i = 0;
    while (i < rtf.length) {
      const ch = rtf.charAt(i);
      if (ch === '{') {
        flushBytes();
        stack.push(group);
        group = { skip: group.skip, uc: group.uc };
        i++;
        if (rtf.substr(i, 2) === '\\*') group.skip = true;
      } else if (ch === '}') {
        flushBytes();
        group = stack.pop() || { skip: false, uc: 1 };
        i++;
      } else if (ch === '\\') {
        const next = rtf.charAt(i + 1);
        if (next === "'") {
          const hex = parseInt(rtf.substr(i + 2, 2), 16);
          i += 4;
          if (group.skip) continue;
          if (skipChars > 0) { skipChars--; continue; }
          if (!isNaN(hex)) pendingBytes.push(hex);
        } else if (/[a-zA-Z]/.test(next)) {
          const m = /^\\([a-zA-Z]+)(-?\d+)? ?/.exec(rtf.substr(i, 40));
          i += m[0].length;
          const word = m[1];
          const param = m[2] !== undefined ? parseInt(m[2], 10) : null;
          if (RTF_SKIP_DESTINATIONS.indexOf(word) !== -1) {
            group.skip = true;
          } else if (word === 'uc' && param !== null) {
            group.uc = param;
          } else if (word === 'u' && param !== null) {
            if (!group.skip) {
              flushBytes();
              current += String.fromCharCode(param < 0 ? param + 65536 : param);
              skipChars = group.uc;
            }
          } else if (!group.skip) {
            if (word === 'par' || word === 'sect' || word === 'page' || word === 'row' || word === 'cell' || word === 'nestcell') endParagraph();
            else if (word === 'line') emit('\n');
            else if (word === 'tab') emit('\t');
            else if (word === 'emdash') emit('—');
            else if (word === 'endash') emit('–');
            else if (word === 'lquote') emit('‘');
            else if (word === 'rquote') emit('’');
            else if (word === 'ldblquote') emit('“');
            else if (word === 'rdblquote') emit('”');
            else if (word === 'bullet') emit('•');
          }
        } else {
          // Escaped character: \\ \{ \} or a control symbol (\~ non-breaking space, \- soft hyphen)
          i += 2;
          if (next === '\\' || next === '{' || next === '}') emit(next);
          else if (next === '~') emit(' ');
          else if (next === '\n' || next === '\r') endParagraph();
        }
      } else {
        i++;
        if (ch === '\r' || ch === '\n') continue;
        emit(ch);
      }
    }
    endParagraph();
    return paragraphs;
  }

  function readUInt16(bytes, pos) { return bytes[pos] | (bytes[pos + 1] << 8); }
  function readUInt32(bytes, pos) { return (bytes[pos] | (bytes[pos + 1] << 8) | (bytes[pos + 2] << 16) | (bytes[pos + 3] << 24)) >>> 0; }

  /**
   * Main document text of a Word 97-2003 binary file, read through its piece table (the text is
   * stored in pieces, each either UTF-16LE or 8-bit Windows-1252).
   */
  function readBinaryDocText(wordStream, tableStream) {
    const fib = wordStream;
    if (readUInt16(fib, 0) !== 0xA5EC) throw new Error('Invalid Word document (FIB)');
    const ccpText = readUInt32(fib, 0x4C);
    const fcClx = readUInt32(fib, 0x1A2);
    const lcbClx = readUInt32(fib, 0x1A6);
    if (!tableStream || lcbClx === 0 || fcClx + lcbClx > tableStream.length) throw new Error('Invalid Word document (piece table)');

    // Clx: Prc entries (0x01) followed by the Pcdt (0x02)
    let pos = fcClx;
    const end = fcClx + lcbClx;
    while (pos < end && tableStream[pos] === 0x01) pos += 3 + readUInt16(tableStream, pos + 1);
    if (tableStream[pos] !== 0x02) throw new Error('Invalid Word document (piece table)');
    const lcb = readUInt32(tableStream, pos + 1);
    const plc = pos + 5;
    const n = Math.floor((lcb - 4) / 12);
    const cp1252 = new TextDecoder('windows-1252');
    let text = '';
    for (let k = 0; k < n && text.length < ccpText; k++) {
      const cpStart = readUInt32(tableStream, plc + k * 4);
      const cpEnd = readUInt32(tableStream, plc + (k + 1) * 4);
      const pcd = plc + (n + 1) * 4 + k * 8;
      const fcRaw = readUInt32(tableStream, pcd + 2);
      const compressed = (fcRaw & 0x40000000) !== 0;
      const count = Math.min(cpEnd - cpStart, ccpText - text.length);
      if (compressed) {
        const fc = (fcRaw & 0x3FFFFFFF) / 2;
        text += cp1252.decode(wordStream.subarray(fc, fc + count));
      } else {
        const fc = fcRaw & 0x3FFFFFFF;
        const slice = wordStream.subarray(fc, fc + count * 2);
        let s = '';
        for (let j = 0; j + 1 < slice.length; j += 2) s += String.fromCharCode(slice[j] | (slice[j + 1] << 8));
        text += s;
      }
    }
    return text;
  }

  /**
   * Splits Word binary text into paragraphs: fields keep only their result (code between 0x13
   * and 0x14 is dropped), 0x0D ends a paragraph, 0x07 ends a table cell.
   */
  function binaryDocParagraphs(text) {
    let out = '';
    let depth = 0;
    const inCode = [];
    for (let i = 0; i < text.length; i++) {
      const c = text.charCodeAt(i);
      if (c === 0x13) { depth++; inCode[depth] = true; continue; }
      if (c === 0x14) { if (depth > 0) inCode[depth] = false; continue; }
      if (c === 0x15) { if (depth > 0) depth--; continue; }
      if (depth > 0 && inCode[depth]) continue;
      if (c === 0x0D || c === 0x07 || c === 0x0C) out += '\n';
      else if (c === 0x0B) out += '\u2028';
      else if (c === 0x09) out += '\t';
      else if (c === 0x1E) out += '-';
      else if (c === 0xA0) out += ' ';
      else if (c >= 0x20 && c !== 0x1F) out += text.charAt(i);
    }
    return out.split('\n').map(function (p) { return p.replace(/\u2028/g, '\n').trim(); }).filter(Boolean);
  }

  function findCfbEntry(cfb, name) {
    let entry = XLSX.CFB.find(cfb, '/' + name) || XLSX.CFB.find(cfb, name);
    if (!entry && cfb.FullPaths) {
      const p = cfb.FullPaths.find(function (fp) { return new RegExp('/' + name + '$', 'i').test(fp); });
      if (p) entry = XLSX.CFB.find(cfb, p);
    }
    return entry && entry.content ? new Uint8Array(entry.content) : null;
  }

  function extractParagraphs(content) {
    const bytes = content instanceof Uint8Array ? content : new Uint8Array(content);

    // 1. .docx (ZIP archive)
    if (bytes.length >= 4 && bytes[0] === 0x50 && bytes[1] === 0x4B) {
      if (!XLSX || !XLSX.CFB) {
        throw new Error('CFB parser is required for docx.');
      }
      const cfb = XLSX.CFB.read(bytes, { type: 'array' });
      let entry = XLSX.CFB.find(cfb, '/word/document.xml');
      if (!entry && cfb.FullPaths) {
        const p = cfb.FullPaths.find(function (fp) {
          return /\/word\/document\.xml$/i.test(fp) || /document\.xml$/i.test(fp);
        });
        if (p) entry = XLSX.CFB.find(cfb, p);
      }
      if (!entry || !entry.content) {
        throw new Error('Invalid Word document: missing word/document.xml');
      }
      const xml = decodeUtf8(entry.content);
      return parseDocxXml(xml);
    }

    // 2. Binary Word 97-2003 OLE/CFB (.doc)
    if (bytes.length >= 8 && bytes[0] === 0xD0 && bytes[1] === 0xCF && bytes[2] === 0x11 && bytes[3] === 0xE0) {
      if (!XLSX || !XLSX.CFB) throw new Error('CFB parser is required for doc.');
      const cfb = XLSX.CFB.read(bytes, { type: 'array' });
      const wordStream = findCfbEntry(cfb, 'WordDocument');
      if (!wordStream) throw new Error('Invalid Word document: missing WordDocument stream');
      if (readUInt16(wordStream, 0x0A) & 0x0100) {
        const err = new Error('The Word document is encrypted.');
        err.code = 'ENCRYPTED';
        throw err;
      }
      // fWhichTblStm selects the table stream holding the piece table
      const tableStream = findCfbEntry(cfb, (readUInt16(wordStream, 0x0A) & 0x0200) ? '1Table' : '0Table');
      return binaryDocParagraphs(readBinaryDocText(wordStream, tableStream));
    }

    // 3. HTML-based .doc or RTF or plain text (UTF-8 or a legacy code page)
    const text = Importers.decodeText(bytes);
    const trimmed = text.trim();
    if (trimmed.indexOf('{\\rtf') === 0) {
      return parseRtfDoc(text);
    }
    if (trimmed.charAt(0) === '<') {
      return parseHtmlDoc(text);
    }

    // Fallback: plain text
    return text.split(/\r?\n+/).map(function (s) { return s.trim(); }).filter(Boolean);
  }

  const adapter = {
    id: 'word',
    label: 'Word (.docx, .doc, .rtf)',
    extensions: EXTENSIONS,
    readAs: 'arraybuffer',
    canImport: function (fileName) {
      const ext = Importers.extensionOf(fileName);
      return EXTENSIONS.indexOf(ext) !== -1;
    },
    parse: function (content, options) {
      const paragraphs = extractParagraphs(content);
      if (!paragraphs || paragraphs.length === 0) {
        const err = new Error('The Word document contains no text.');
        err.code = 'NO_TEXT';
        throw err;
      }

      return {
        fileName: options && options.fileName,
        adapterId: 'word',
        sheetNames: ['Dokument'],
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
