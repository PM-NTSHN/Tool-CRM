/* xlsx-io.js – bộ đọc/ghi file Excel (.xlsx) tối giản, không phụ thuộc thư viện ngoài.
 * Chạy được trên trình duyệt hiện đại (Chrome/Edge/Firefox/Safari) và Node 18+.
 * Dùng DecompressionStream / CompressionStream có sẵn của trình duyệt để giải nén / nén zip.
 *
 *   XLSXIO.read(arrayBuffer)  -> Promise<{ sheets: [{ name, rows: any[][] }] }>
 *   XLSXIO.write(spec)        -> Promise<Uint8Array>
 *
 * Ngày tháng: mọi giá trị ngày được biểu diễn bằng Date theo giờ UTC "ngây thơ"
 * (giờ hiển thị trong Excel = giờ UTC của Date), tránh lệch múi giờ.
 */
(function (root) {
  'use strict';

  // ---------------------------------------------------------------- tiện ích chung
  const te = new TextEncoder();
  const td = new TextDecoder('utf-8');

  const CRC_TABLE = (() => {
    const t = new Uint32Array(256);
    for (let n = 0; n < 256; n++) {
      let c = n;
      for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
      t[n] = c >>> 0;
    }
    return t;
  })();
  function crc32(buf) {
    let c = 0xffffffff;
    for (let i = 0; i < buf.length; i++) c = CRC_TABLE[(c ^ buf[i]) & 0xff] ^ (c >>> 8);
    return (c ^ 0xffffffff) >>> 0;
  }

  async function streamThrough(data, stream) {
    const out = new Response(new Blob([data]).stream().pipeThrough(stream));
    return new Uint8Array(await out.arrayBuffer());
  }

  // ---------------------------------------------------------------- ĐỌC ZIP
  async function unzip(u8) {
    const dv = new DataView(u8.buffer, u8.byteOffset, u8.byteLength);
    let eocd = -1;
    for (let i = u8.length - 22; i >= Math.max(0, u8.length - 65557); i--) {
      if (dv.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
    }
    if (eocd < 0) throw new Error('File không phải định dạng .xlsx hợp lệ (không đọc được cấu trúc zip).');
    const count = dv.getUint16(eocd + 10, true);
    let p = dv.getUint32(eocd + 16, true);
    const files = {};
    for (let k = 0; k < count; k++) {
      if (dv.getUint32(p, true) !== 0x02014b50) break;
      const method = dv.getUint16(p + 10, true);
      const csize = dv.getUint32(p + 20, true);
      const nlen = dv.getUint16(p + 28, true), elen = dv.getUint16(p + 30, true), clen = dv.getUint16(p + 32, true);
      const off = dv.getUint32(p + 42, true);
      const name = td.decode(u8.subarray(p + 46, p + 46 + nlen));
      const lnlen = dv.getUint16(off + 26, true), lelen = dv.getUint16(off + 28, true);
      const start = off + 30 + lnlen + lelen;
      files[name] = { method, data: u8.subarray(start, start + csize) };
      p += 46 + nlen + elen + clen;
    }
    return {
      has: (n) => n in files,
      async text(n) {
        const f = files[n];
        if (!f) return null;
        if (f.method === 0) return td.decode(f.data);
        if (f.method === 8) return td.decode(await streamThrough(f.data, new DecompressionStream('deflate-raw')));
        throw new Error('Kiểu nén zip không hỗ trợ: ' + f.method);
      },
    };
  }

  // ---------------------------------------------------------------- ĐỌC XML
  function xmlDecode(s) {
    if (s == null) return s;
    return s
      .replace(/&(lt|gt|amp|quot|apos|#\d+|#x[0-9a-fA-F]+);/g, (m, e) => {
        switch (e) {
          case 'lt': return '<'; case 'gt': return '>'; case 'amp': return '&';
          case 'quot': return '"'; case 'apos': return "'";
          default: return String.fromCodePoint(e[1] === 'x' ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10));
        }
      })
      .replace(/_x([0-9A-Fa-f]{4})_/g, (m, h) => String.fromCharCode(parseInt(h, 16)));
  }
  function attrs(s) {
    const o = {};
    s.replace(/([\w:]+)\s*=\s*"([^"]*)"/g, (m, k, v) => { o[k] = xmlDecode(v); return ''; });
    return o;
  }
  function richText(x) {
    x = x.replace(/<rPh\b[\s\S]*?<\/rPh>/g, '');
    let out = '';
    x.replace(/<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g, (m, t) => { out += t; return ''; });
    return xmlDecode(out);
  }

  const DATE_IDS = new Set([14, 15, 16, 17, 18, 19, 20, 21, 22, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 45, 46, 47, 50, 51, 52, 53, 54, 55, 56, 57, 58]);
  function isDateFormat(code) {
    const c = code.replace(/"[^"]*"/g, '').replace(/\[[^\]]*\]/g, '').replace(/\\./g, '');
    return /[dmyhs]/i.test(c) && !/^general$/i.test(c.trim());
  }
  function serialToDate(n) {
    return new Date(Math.round((n - 25569) * 86400000));
  }
  function colIndex(ref) {
    let n = 0;
    for (let i = 0; i < ref.length; i++) {
      const c = ref.charCodeAt(i);
      if (c < 65 || c > 90) break;
      n = n * 26 + (c - 64);
    }
    return n - 1;
  }

  async function read(buffer) {
    const u8 = buffer instanceof Uint8Array ? buffer : new Uint8Array(buffer);
    const z = await unzip(u8);
    const wb = await z.text('xl/workbook.xml');
    if (!wb) throw new Error('Không tìm thấy xl/workbook.xml – file không phải Excel .xlsx.');
    const rels = (await z.text('xl/_rels/workbook.xml.rels')) || '';
    const relMap = {};
    rels.replace(/<Relationship\b([^>]*)\/?>/g, (m, a) => { const o = attrs(a); relMap[o.Id] = o.Target; return ''; });

    const sst = [];
    const sstXml = await z.text('xl/sharedStrings.xml');
    if (sstXml) sstXml.replace(/<si>([\s\S]*?)<\/si>|<si\/>/g, (m, x) => { sst.push(x ? richText(x) : ''); return ''; });

    const dateXf = [];
    const stXml = await z.text('xl/styles.xml');
    if (stXml) {
      const fmts = {};
      stXml.replace(/<numFmt\b([^>]*)\/?>/g, (m, a) => { const o = attrs(a); fmts[+o.numFmtId] = o.formatCode || ''; return ''; });
      const xfs = (stXml.match(/<cellXfs\b[\s\S]*?<\/cellXfs>/) || [''])[0];
      xfs.replace(/<xf\b([^>]*)\/?>/g, (m, a) => {
        const id = +(attrs(a).numFmtId || 0);
        dateXf.push(DATE_IDS.has(id) || (fmts[id] != null && isDateFormat(fmts[id])));
        return '';
      });
    }

    const sheets = [];
    const sheetDefs = [];
    wb.replace(/<sheet\b([^>]*)\/?>/g, (m, a) => { sheetDefs.push(attrs(a)); return ''; });
    for (const sd of sheetDefs) {
      let target = relMap[sd['r:id']] || '';
      target = target.startsWith('/') ? target.slice(1) : 'xl/' + target.replace(/^\.\//, '');
      const xml = await z.text(target);
      if (!xml) continue;
      const rows = [];
      let rowSeq = 0;
      const sheetData = (xml.match(/<sheetData\b[\s\S]*?<\/sheetData>/) || [''])[0];
      sheetData.replace(/<row\b([^>]*?)(?:\/>|>([\s\S]*?)<\/row>)/g, (m, ra, body) => {
        const rAttr = attrs(ra);
        const r = rAttr.r ? +rAttr.r - 1 : rowSeq;
        rowSeq = r + 1;
        const row = [];
        let colSeq = 0;
        (body || '').replace(/<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g, (m2, ca, cb) => {
          const a = attrs(ca);
          const c = a.r ? colIndex(a.r) : colSeq;
          colSeq = c + 1;
          if (!cb) return '';
          const t = a.t || 'n';
          let v = null;
          const vm = cb.match(/<v>([\s\S]*?)<\/v>/);
          const raw = vm ? xmlDecode(vm[1]) : null;
          if (t === 's') v = raw == null ? null : sst[+raw];
          else if (t === 'inlineStr') v = richText((cb.match(/<is>([\s\S]*?)<\/is>/) || ['', ''])[1]);
          else if (t === 'str') v = raw;
          else if (t === 'b') v = raw === '1';
          else if (t === 'e') v = null;
          else if (t === 'd') v = raw ? new Date(raw.endsWith('Z') ? raw : raw + 'Z') : null;
          else if (raw != null && raw !== '') {
            v = Number(raw);
            if (a.s != null && dateXf[+a.s]) v = serialToDate(v);
          }
          row[c] = v;
          return '';
        });
        rows[r] = row;
        return '';
      });
      for (let i = 0; i < rows.length; i++) if (!rows[i]) rows[i] = [];
      sheets.push({ name: sd.name, rows });
    }
    return { sheets };
  }

  // ---------------------------------------------------------------- GHI XLSX
  function esc(s) {
    return String(s)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')
      .replace(/[\x00-\x08\x0b\x0c\x0e-\x1f]/g, '');
  }
  function colName(i) {
    let s = '';
    i++;
    while (i > 0) { const m = (i - 1) % 26; s = String.fromCharCode(65 + m) + s; i = Math.floor((i - 1) / 26); }
    return s;
  }
  function dateToSerial(d) {
    return d.getTime() / 86400000 + 25569;
  }

  /* Style: { font:{name,sz,b,i,u,color}, fill:'RRGGBB', border:'RRGGBB'|{color,style}, numFmt:'#,##0',
   *          align:{h,v,wrap,indent,rotate} } */
  function StyleBook() {
    const fonts = ['<font><sz val="10"/><name val="Arial"/></font>'];
    const fills = ['<fill><patternFill patternType="none"/></fill>', '<fill><patternFill patternType="gray125"/></fill>'];
    const borders = ['<border><left/><right/><top/><bottom/><diagonal/></border>'];
    const numFmts = [];
    const xfs = ['<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'];
    const idx = (arr, x) => { let i = arr.indexOf(x); if (i < 0) { arr.push(x); i = arr.length - 1; } return i; };
    const cache = new Map();
    const BUILTIN = { 'General': 0, '0': 1, '0.00': 2, '#,##0': 3, '#,##0.00': 4, '0%': 9, '0.00%': 10 };
    return {
      id(st) {
        if (!st) return 0;
        const key = JSON.stringify(st);
        if (cache.has(key)) return cache.get(key);
        const f = st.font || {};
        const font = '<font>' + (f.b ? '<b/>' : '') + (f.i ? '<i/>' : '') + (f.u ? '<u/>' : '') +
          `<sz val="${f.sz || 10}"/>` + (f.color ? `<color rgb="FF${f.color}"/>` : '') +
          `<name val="${esc(f.name || 'Arial')}"/></font>`;
        const fontId = idx(fonts, font);
        const fillId = st.fill ? idx(fills, `<fill><patternFill patternType="solid"><fgColor rgb="FF${st.fill}"/><bgColor indexed="64"/></patternFill></fill>`) : 0;
        let borderId = 0;
        if (st.border) {
          const b = typeof st.border === 'string' ? { color: st.border } : st.border;
          const side = (n) => `<${n} style="${b.style || 'thin'}"><color rgb="FF${b.color || 'D9D9D9'}"/></${n}>`;
          const sides = b.sides || ['left', 'right', 'top', 'bottom'];
          borderId = idx(borders, '<border>' + ['left', 'right', 'top', 'bottom'].map((n) => (sides.includes(n) ? side(n) : `<${n}/>`)).join('') + '<diagonal/></border>');
        }
        let numFmtId = 0;
        if (st.numFmt) {
          if (st.numFmt in BUILTIN) numFmtId = BUILTIN[st.numFmt];
          else numFmtId = 164 + idx(numFmts, st.numFmt);
        }
        const a = st.align;
        const al = a ? '<alignment' + (a.h ? ` horizontal="${a.h}"` : '') + (a.v ? ` vertical="${a.v}"` : '') +
          (a.wrap ? ' wrapText="1"' : '') + (a.indent ? ` indent="${a.indent}"` : '') + '/>' : '';
        const xf = `<xf numFmtId="${numFmtId}" fontId="${fontId}" fillId="${fillId}" borderId="${borderId}" xfId="0"` +
          (numFmtId ? ' applyNumberFormat="1"' : '') + (fontId ? ' applyFont="1"' : '') + (fillId ? ' applyFill="1"' : '') +
          (borderId ? ' applyBorder="1"' : '') + (al ? ' applyAlignment="1">' + al + '</xf>' : '/>');
        const id = idx(xfs, xf);
        cache.set(key, id);
        return id;
      },
      xml() {
        return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
          '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
          (numFmts.length ? `<numFmts count="${numFmts.length}">` + numFmts.map((c, i) => `<numFmt numFmtId="${164 + i}" formatCode="${esc(c)}"/>`).join('') + '</numFmts>' : '') +
          `<fonts count="${fonts.length}">${fonts.join('')}</fonts>` +
          `<fills count="${fills.length}">${fills.join('')}</fills>` +
          `<borders count="${borders.length}">${borders.join('')}</borders>` +
          '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
          `<cellXfs count="${xfs.length}">${xfs.join('')}</cellXfs>` +
          '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
          '</styleSheet>';
      },
    };
  }

  function sheetXml(sh, styles, sst) {
    const rows = sh.rows || [];
    let maxC = 0;
    const parts = [];
    rows.forEach((row, r) => {
      if (!row) return;
      const cells = [];
      row.forEach((cell, c) => {
        if (cell === undefined) return;
        let v = cell, st = null;
        if (cell && typeof cell === 'object' && !(cell instanceof Date)) { v = cell.v; st = cell.s; }
        if (v instanceof Date && st && !st.numFmt) st = Object.assign({}, st, { numFmt: 'dd/mm/yyyy' });
        if (v instanceof Date && !st) st = { numFmt: 'dd/mm/yyyy' };
        const s = styles.id(st);
        const ref = colName(c) + (r + 1);
        const sa = s ? ` s="${s}"` : '';
        if (c + 1 > maxC) maxC = c + 1;
        if (v === null || v === undefined || v === '' || (typeof v === 'number' && !isFinite(v))) {
          if (s) cells.push(`<c r="${ref}"${sa}/>`);
        } else if (v instanceof Date) {
          if (isNaN(v)) { if (s) cells.push(`<c r="${ref}"${sa}/>`); } else cells.push(`<c r="${ref}"${sa}><v>${dateToSerial(v)}</v></c>`);
        } else if (typeof v === 'number') {
          cells.push(`<c r="${ref}"${sa}><v>${v}</v></c>`);
        } else if (typeof v === 'boolean') {
          cells.push(`<c r="${ref}"${sa} t="b"><v>${v ? 1 : 0}</v></c>`);
        } else if (typeof v === 'object' && v.f) {
          cells.push(`<c r="${ref}"${sa}><f>${esc(v.f)}</f></c>`);
        } else {
          let str = String(v);
          if (str.length > 32767) str = str.slice(0, 32767);
          let i = sst.map.get(str);
          if (i === undefined) { i = sst.list.length; sst.list.push(str); sst.map.set(str, i); }
          sst.count++;
          cells.push(`<c r="${ref}"${sa} t="s"><v>${i}</v></c>`);
        }
      });
      const ht = sh.rowHeights && sh.rowHeights[r] ? ` ht="${sh.rowHeights[r]}" customHeight="1"` : '';
      parts.push(`<row r="${r + 1}"${ht}>${cells.join('')}</row>`);
    });
    const fr = sh.freeze;
    let view = '<sheetView workbookViewId="0"' + (sh.showGrid === false ? ' showGridLines="0"' : '') + (sh.selected ? ' tabSelected="1"' : '') + '>';
    if (fr && (fr.row || fr.col)) {
      const tl = colName(fr.col || 0) + ((fr.row || 0) + 1);
      const pane = fr.row && fr.col ? 'bottomRight' : fr.row ? 'bottomLeft' : 'topRight';
      view += `<pane${fr.col ? ` xSplit="${fr.col}"` : ''}${fr.row ? ` ySplit="${fr.row}"` : ''} topLeftCell="${tl}" activePane="${pane}" state="frozen"/>` +
        `<selection pane="${pane}" activeCell="${tl}" sqref="${tl}"/>`;
    }
    view += '</sheetView>';
    const cols = (sh.cols || []).map((w, i) => (w ? `<col min="${i + 1}" max="${i + 1}" width="${w}" customWidth="1"/>` : '')).join('');
    const dim = rows.length && maxC ? `A1:${colName(maxC - 1)}${rows.length}` : 'A1';
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      (sh.tabColor ? `<sheetPr><tabColor rgb="FF${sh.tabColor}"/></sheetPr>` : '') +
      `<dimension ref="${dim}"/><sheetViews>${view}</sheetViews><sheetFormatPr defaultRowHeight="15"/>` +
      (cols ? `<cols>${cols}</cols>` : '') +
      `<sheetData>${parts.join('')}</sheetData>` +
      (sh.autoFilter ? `<autoFilter ref="${sh.autoFilter}"/>` : '') +
      (sh.merges && sh.merges.length ? `<mergeCells count="${sh.merges.length}">${sh.merges.map((m) => `<mergeCell ref="${m}"/>`).join('')}</mergeCells>` : '') +
      '<pageMargins left="0.5" right="0.5" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>' +
      '</worksheet>';
  }

  async function zip(entries) {
    const canDeflate = typeof CompressionStream !== 'undefined';
    const chunks = [], central = [];
    let offset = 0;
    for (const [name, content] of entries) {
      const raw = typeof content === 'string' ? te.encode(content) : content;
      const crc = crc32(raw);
      let data = raw, method = 0;
      if (canDeflate && raw.length > 256) {
        try { data = await streamThrough(raw, new CompressionStream('deflate-raw')); method = 8; } catch (e) { data = raw; method = 0; }
      }
      const nb = te.encode(name);
      const lh = new DataView(new ArrayBuffer(30));
      lh.setUint32(0, 0x04034b50, true); lh.setUint16(4, 20, true); lh.setUint16(6, 0x0800, true);
      lh.setUint16(8, method, true); lh.setUint16(10, 0, true); lh.setUint16(12, 0x21, true);
      lh.setUint32(14, crc, true); lh.setUint32(18, data.length, true); lh.setUint32(22, raw.length, true);
      lh.setUint16(26, nb.length, true); lh.setUint16(28, 0, true);
      chunks.push(new Uint8Array(lh.buffer), nb, data);
      const ch = new DataView(new ArrayBuffer(46));
      ch.setUint32(0, 0x02014b50, true); ch.setUint16(4, 20, true); ch.setUint16(6, 20, true); ch.setUint16(8, 0x0800, true);
      ch.setUint16(10, method, true); ch.setUint16(12, 0, true); ch.setUint16(14, 0x21, true);
      ch.setUint32(16, crc, true); ch.setUint32(20, data.length, true); ch.setUint32(24, raw.length, true);
      ch.setUint16(28, nb.length, true); ch.setUint32(42, offset, true);
      central.push(new Uint8Array(ch.buffer), nb);
      offset += 30 + nb.length + data.length;
    }
    const cdSize = central.reduce((s, c) => s + c.length, 0);
    const end = new DataView(new ArrayBuffer(22));
    end.setUint32(0, 0x06054b50, true); end.setUint16(8, entries.length, true); end.setUint16(10, entries.length, true);
    end.setUint32(12, cdSize, true); end.setUint32(16, offset, true);
    const all = [...chunks, ...central, new Uint8Array(end.buffer)];
    const out = new Uint8Array(all.reduce((s, c) => s + c.length, 0));
    let p = 0;
    for (const c of all) { out.set(c, p); p += c.length; }
    return out;
  }

  async function write(spec) {
    const styles = StyleBook();
    const sst = { list: [], map: new Map(), count: 0 };
    const sheetFiles = spec.sheets.map((sh, i) => [`xl/worksheets/sheet${i + 1}.xml`, sheetXml(Object.assign({ selected: i === 0 }, sh), styles, sst)]);
    const names = spec.sheets.map((s) => esc(String(s.name).replace(/[\\/?*[\]:]/g, ' ').slice(0, 31)));
    const definedNames = spec.sheets.map((s, i) => (s.autoFilter ? `<definedName name="_xlnm._FilterDatabase" localSheetId="${i}" hidden="1">'${names[i].replace(/'/g, "''")}'!${s.autoFilter.split(':').map((x) => x.replace(/([A-Z]+)(\d+)/, '$$$1$$$2')).join(':')}</definedName>` : '')).join('');
    const entries = [
      ['[Content_Types].xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>' +
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
        spec.sheets.map((s, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('') +
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
        '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>' +
        '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>' +
        '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>'],
      ['_rels/.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>' +
        '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>'],
      ['docProps/core.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
        `<dc:title>${esc(spec.title || '')}</dc:title><dc:creator>${esc(spec.creator || 'NTS Hanoi')}</dc:creator>` +
        `<dcterms:created xsi:type="dcterms:W3CDTF">${new Date().toISOString().replace(/\.\d+Z$/, 'Z')}</dcterms:created></cp:coreProperties>`],
      ['docProps/app.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Microsoft Excel</Application></Properties>'],
      ['xl/workbook.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
        '<bookViews><workbookView activeTab="0"/></bookViews><sheets>' +
        names.map((n, i) => `<sheet name="${n}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('') + '</sheets>' +
        (definedNames ? `<definedNames>${definedNames}</definedNames>` : '') + '</workbook>'],
      ['xl/_rels/workbook.xml.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        spec.sheets.map((s, i) => `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join('') +
        `<Relationship Id="rId${spec.sheets.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>` +
        `<Relationship Id="rId${spec.sheets.length + 2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/></Relationships>`],
      ...sheetFiles,
    ];
    entries.push(['xl/styles.xml', styles.xml()]);
    entries.push(['xl/sharedStrings.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${sst.count}" uniqueCount="${sst.list.length}">` +
      sst.list.map((s) => `<si><t xml:space="preserve">${esc(s)}</t></si>`).join('') + '</sst>']);
    return zip(entries);
  }

  root.XLSXIO = { read, write, colName };
})(typeof globalThis !== 'undefined' ? globalThis : this);
