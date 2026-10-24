/**
 * Word document import adapter (.docx, .doc).
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

  const EXTENSIONS = ['.docx', '.doc'];

  function decodeUtf8(bytes) {
    if (!bytes || bytes.length === 0) return '';
    if (typeof TextDecoder !== 'undefined') {
      return new TextDecoder('utf-8').decode(bytes);
    }
    if (typeof Buffer !== 'undefined') {
      return Buffer.from(bytes).toString('utf8');
    }
    let str = '';
    for (let i = 0; i < bytes.length; i++) {
      str += String.fromCharCode(bytes[i]);
    }
    try {
      return decodeURIComponent(escape(str));
    } catch (_) {
      return str;
    }
  }

  function unescapeXml(str) {
    if (!str) return '';
    return str
      .replace(/&amp;/g, '&')
      .replace(/&lt;/g, '<')
      .replace(/&gt;/g, '>')
      .replace(/&quot;/g, '"')
      .replace(/&apos;/g, "'")
      .replace(/&#(\d+);/g, function (_, n) { return String.fromCharCode(parseInt(n, 10)); })
      .replace(/&#x([0-9a-fA-F]+);/g, function (_, h) { return String.fromCharCode(parseInt(h, 16)); });
  }

  function parseDocxXml(xml) {
    const pRegex = /<w:p(?:\s+[^>]*)?>([\s\S]*?)<\/w:p>/gi;
    const tRegex = /<w:t(?:\s+[^>]*)?>([\s\S]*?)<\/w:t>|<w:tab\/>|<w:br\/>/gi;
    const paragraphs = [];
    let pMatch;
    while ((pMatch = pRegex.exec(xml)) !== null) {
      const pXml = pMatch[1];
      let pText = '';
      let tMatch;
      while ((tMatch = tRegex.exec(pXml)) !== null) {
        if (tMatch[0] === '<w:tab/>') pText += '\t';
        else if (tMatch[0] === '<w:br/>') pText += '\n';
        else if (tMatch[1] !== undefined) pText += tMatch[1];
      }
      pText = unescapeXml(pText).trim();
      if (pText) {
        paragraphs.push(pText);
      }
    }
    return paragraphs;
  }

  function parseHtmlDoc(html) {
    const cleanHtml = html
      .replace(/<style[\s\S]*?<\/style>/gi, '')
      .replace(/<script[\s\S]*?<\/script>/gi, '')
      .replace(/<xml[\s\S]*?<\/xml>/gi, '')
      .replace(/<head[\s\S]*?<\/head>/gi, '');

    const pRegex = /<(?:p|div|tr|li|h[1-6])(?:\s+[^>]*)?>([\s\S]*?)<\/(?:p|div|tr|li|h[1-6])>/gi;
    const paragraphs = [];
    let match;
    while ((match = pRegex.exec(cleanHtml)) !== null) {
      let text = match[1]
        .replace(/<br\s*\/?>/gi, '\n')
        .replace(/<[^>]+>/g, '')
        .replace(/&nbsp;/g, ' ')
        .replace(/&amp;/g, '&')
        .replace(/&lt;/g, '<')
        .replace(/&gt;/g, '>')
        .replace(/&quot;/g, '"')
        .replace(/&#39;/g, "'")
        .replace(/&#(\d+);/g, function (_, n) { return String.fromCharCode(parseInt(n, 10)); })
        .replace(/&#x([0-9a-fA-F]+);/g, function (_, h) { return String.fromCharCode(parseInt(h, 16)); })
        .trim();
      if (text) {
        paragraphs.push(text);
      }
    }
    if (paragraphs.length === 0) {
      const plain = cleanHtml.replace(/<[^>]+>/g, '\n').replace(/&nbsp;/g, ' ');
      plain.split(/\n+/).forEach(function (line) {
        const t = line.trim();
        if (t) paragraphs.push(t);
      });
    }
    return paragraphs;
  }

  function parseRtfDoc(rtf) {
    const paragraphs = [];
    const parts = rtf.split(/\\par(?:\r?\n|\s+|$)/i);
    parts.forEach(function (part) {
      let text = part
        .replace(/\\'[0-9a-fA-F]{2}/g, function (m) {
          return String.fromCharCode(parseInt(m.slice(2), 16));
        })
        .replace(/\\u(-?\d+)\??/g, function (_, n) {
          const code = parseInt(n, 10);
          return String.fromCharCode(code < 0 ? code + 65536 : code);
        })
        .replace(/\\[a-zA-Z]+-?\d*[\s]?/g, '')
        .replace(/[{}]/g, '')
        .trim();
      if (text) paragraphs.push(text);
    });
    return paragraphs;
  }

  function parseBinaryDocStream(bytes) {
    // Extract readable text sequences from WordDocument stream
    // Words are terminated by \r (0x0D)
    const paragraphs = [];
    let current = '';

    // First attempt: check for UTF-16LE text
    let utf16Count = 0;
    for (let i = 0; i < bytes.length - 1; i += 2) {
      if (bytes[i + 1] === 0x00 && bytes[i] >= 0x20 && bytes[i] <= 0x7E) utf16Count++;
    }

    if (utf16Count > 20) {
      for (let i = 0; i < bytes.length - 1; i += 2) {
        const code = bytes[i] | (bytes[i + 1] << 8);
        if (code === 0x0D || code === 0x0A) {
          const t = current.trim();
          if (t.length >= 2) paragraphs.push(t);
          current = '';
        } else if (code >= 0x20 && code < 0xFFFE) {
          current += String.fromCharCode(code);
        }
      }
      if (current.trim().length >= 2) paragraphs.push(current.trim());
    }

    if (paragraphs.length === 0) {
      // 8-bit ASCII / ANSI run extraction
      current = '';
      for (let i = 0; i < bytes.length; i++) {
        const b = bytes[i];
        if (b === 0x0D || b === 0x0A) {
          const t = current.trim();
          if (t.length >= 2) paragraphs.push(t);
          current = '';
        } else if (b >= 0x20 && b <= 0xFF) {
          current += String.fromCharCode(b);
        }
      }
      if (current.trim().length >= 2) paragraphs.push(current.trim());
    }

    return paragraphs;
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
      if (XLSX && XLSX.CFB) {
        const cfb = XLSX.CFB.read(bytes, { type: 'array' });
        let entry = XLSX.CFB.find(cfb, '/WordDocument') || XLSX.CFB.find(cfb, 'WordDocument');
        if (!entry && cfb.FullPaths) {
          const p = cfb.FullPaths.find(function (fp) { return /WordDocument/i.test(fp); });
          if (p) entry = XLSX.CFB.find(cfb, p);
        }
        if (entry && entry.content) {
          const paras = parseBinaryDocStream(entry.content);
          if (paras.length > 0) return paras;
        }
      }
    }

    // 3. HTML-based .doc or RTF or plain text
    const text = decodeUtf8(bytes);
    const trimmed = text.trim();
    if (trimmed.charAt(0) === '<') {
      return parseHtmlDoc(text);
    }
    if (trimmed.indexOf('{\\rtf') === 0) {
      return parseRtfDoc(text);
    }

    // Fallback: plain text
    return text.split(/\r?\n+/).map(function (s) { return s.trim(); }).filter(Boolean);
  }

  const adapter = {
    id: 'word',
    label: 'Word (.docx, .doc)',
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
