/* Sun Mill VAT template exporter.
 * Rebuilds sheet 'Q-1' of the template from template.json plus the app's
 * entries, recomputes every formula's cached value, and zips the workbook.
 * Works in the browser (window.SunMillExport) and in Node (module.exports). */
(function (root) {
  const SALES_FIRST = 23, SALES_LAST = 122, EXP_FIRST = 139, EXP_LAST = 270;
  const EDIT_COLS = ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I'];

  const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '');
  const num = (n) => String(+(+n).toPrecision(15));

  // Excel serial date from 'YYYY-MM-DD'
  function serial(iso) {
    const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(iso || '');
    if (!m) return null;
    const t = Date.UTC(+m[1], +m[2] - 1, +m[3]);
    return Math.round(t / 86400000) + 25569;
  }

  // ---------- tiny formula evaluator (arithmetic, refs, SUM, string refs) ----------
  function colNum(c) { let n = 0; for (const ch of c) n = n * 26 + ch.charCodeAt(0) - 64; return n; }
  function colName(n) { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = (n - m - 1) / 26; } return s; }

  function tokenize(f) {
    const re = /\s*(\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?|\d+(?:\.\d+)?|SUM|[-+*/(),])/gy;
    const out = []; let m; re.lastIndex = 0;
    while (re.lastIndex < f.length && (m = re.exec(f))) out.push(m[1]);
    if (re.lastIndex < f.length && f.slice(re.lastIndex).trim()) throw new Error('Cannot parse formula: ' + f);
    return out;
  }

  function evaluator(get) {
    const cache = new Map();
    function evalFormula(f) {
      const toks = tokenize(f); let i = 0;
      const peek = () => toks[i];
      const val = (x) => (typeof x === 'number' ? x : (x === '' || x == null ? 0 : (isNaN(+x) ? 0 : +x)));
      function expr() { let v = term(); while (peek() === '+' || peek() === '-') { const op = toks[i++]; const r = term(); v = op === '+' ? val(v) + val(r) : val(v) - val(r); } return v; }
      function term() { let v = unary(); while (peek() === '*' || peek() === '/') { const op = toks[i++]; const r = unary(); v = op === '*' ? val(v) * val(r) : (val(r) === 0 ? 0 : val(v) / val(r)); } return v; }
      function unary() { if (peek() === '-') { i++; return -val(unary()); } if (peek() === '+') { i++; return val(unary()); } return atom(); }
      function atom() {
        const t = toks[i++];
        if (t === '(') { const v = expr(); i++; return v; }
        if (t === 'SUM') {
          i++; let total = 0;
          for (;;) {
            const a = toks[i];
            if (a.includes(':')) { i++; total += sumRange(a); } else total += val(expr());
            if (toks[i] === ',') { i++; continue; }
            i++; break;
          }
          return total;
        }
        if (/^\d/.test(t)) return +t;
        if (/^\$?[A-Z]/.test(t)) return cell(t.replace(/\$/g, ''));
        throw new Error('Bad token ' + t);
      }
      return expr();
    }
    function sumRange(rng) {
      const [a, b] = rng.replace(/\$/g, '').split(':');
      const [, ca, ra] = /^([A-Z]+)(\d+)$/.exec(a); const [, cb, rb] = /^([A-Z]+)(\d+)$/.exec(b);
      let t = 0;
      for (let c = colNum(ca); c <= colNum(cb); c++) for (let r = +ra; r <= +rb; r++) {
        const v = cell(colName(c) + r); if (typeof v === 'number') t += v;
      }
      return t;
    }
    function cell(ref) {
      if (cache.has(ref)) return cache.get(ref);
      const c = get(ref);
      let v = '';
      if (c) {
        if (c.f != null) { cache.set(ref, 0); v = evalFormula(c.f); }
        else if (c.num != null) v = c.num;
        else if (c.str != null) v = c.str;
      }
      cache.set(ref, v);
      return v;
    }
    return cell;
  }

  // ---------- zip ----------
  const CRC_TABLE = (() => { const t = new Uint32Array(256); for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; } return t; })();
  function crc32(buf) { let c = 0xFFFFFFFF; for (let i = 0; i < buf.length; i++) c = CRC_TABLE[(c ^ buf[i]) & 0xFF] ^ (c >>> 8); return (c ^ 0xFFFFFFFF) >>> 0; }
  function b64ToBytes(b64) {
    if (typeof atob === 'function') { const s = atob(b64); const u = new Uint8Array(s.length); for (let i = 0; i < s.length; i++) u[i] = s.charCodeAt(i); return u; }
    return new Uint8Array(Buffer.from(b64, 'base64'));
  }
  function zip(entries) {
    const enc = new TextEncoder(); const chunks = []; const central = []; let offset = 0;
    const DOS_TIME = 0, DOS_DATE = ((2026 - 1980) << 9) | (1 << 5) | 1;
    for (const e of entries) {
      const name = enc.encode(e.name);
      const h = new DataView(new ArrayBuffer(30));
      h.setUint32(0, 0x04034b50, true); h.setUint16(4, 20, true); h.setUint16(6, 0x0800, true);
      h.setUint16(8, e.method, true); h.setUint16(10, DOS_TIME, true); h.setUint16(12, DOS_DATE, true);
      h.setUint32(14, e.crc, true); h.setUint32(18, e.data.length, true); h.setUint32(22, e.size, true);
      h.setUint16(26, name.length, true); h.setUint16(28, 0, true);
      chunks.push(new Uint8Array(h.buffer), name, e.data);
      const c = new DataView(new ArrayBuffer(46));
      c.setUint32(0, 0x02014b50, true); c.setUint16(4, 20, true); c.setUint16(6, 20, true); c.setUint16(8, 0x0800, true);
      c.setUint16(10, e.method, true); c.setUint16(12, DOS_TIME, true); c.setUint16(14, DOS_DATE, true);
      c.setUint32(16, e.crc, true); c.setUint32(20, e.data.length, true); c.setUint32(24, e.size, true);
      c.setUint16(28, name.length, true); c.setUint32(42, offset, true);
      central.push(new Uint8Array(c.buffer), name);
      offset += 30 + name.length + e.data.length;
    }
    const cdSize = central.reduce((n, a) => n + a.length, 0);
    const end = new DataView(new ArrayBuffer(22));
    end.setUint32(0, 0x06054b50, true); end.setUint16(8, entries.length, true); end.setUint16(10, entries.length, true);
    end.setUint32(12, cdSize, true); end.setUint32(16, offset, true);
    const all = [...chunks, ...central, new Uint8Array(end.buffer)];
    const out = new Uint8Array(all.reduce((n, a) => n + a.length, 0)); let p = 0;
    for (const a of all) { out.set(a, p); p += a.length; }
    return out;
  }

  // ---------- build ----------
  /* data = {
   *   period: '01.08.2026 to 31.10.2026',
   *   sales:    [{date, invoiceNo, name, desc, category, payment, gross, vat, vatMode:'num'|'sixth', orange}],
   *   expenses: [{date, invoiceNo, name, desc, category, payment, gross, vat, vatMode:'num'|'sixth', orange}]
   * } date: 'YYYY-MM-DD' (written as a real date) or any other text (written as text). */
  /* Make room for more expense lines: insert `extra` rows after the last
   * expense row (270), shifting the total row and everything below, and
   * stretching every range that ends on row 270 (sums, dropdowns, checks). */
  function makeShifter(extra) {
    const last = EXP_LAST;
    const one = (row) => (row > last ? row + extra : row);
    const refRe = /(\$?)([A-Z]{1,3})(\$?)(\d+)(?::(\$?)([A-Z]{1,3})(\$?)(\d+))?/g;
    function text(s) {
      if (!extra) return s;
      return s.replace(refRe, (m, d1, c1, d2, r1, d3, c2, d4, r2) => {
        const a = +r1;
        if (r2 === undefined) return `${d1}${c1}${d2}${one(a)}`;
        const b = +r2;
        const end = b === last && a >= EXP_FIRST && a <= last ? last + extra : one(b);
        return `${d1}${c1}${d2}${one(a)}:${d3}${c2}${d4}${end}`;
      });
    }
    // only touch cell references inside attribute values / formula text, never names like "Q-1"
    function xml(s) {
      if (!extra) return s;
      return s.replace(/( (?:ref|sqref)=")([^"]*)(")/g, (m, a, v, b) => a + text(v) + b)
        .replace(/(<formula1?>)([^<]*)(<\/formula1?>)/g, (m, a, v, b) => a + text(v) + b);
    }
    return { one, text, xml };
  }

  function build(tpl, data) {
    const sales = data.sales || [], expenses = data.expenses || [];
    if (sales.length > SALES_LAST - SALES_FIRST + 1) throw new Error('Too many income lines (max 100).');
    const extra = Math.max(0, expenses.length - (EXP_LAST - EXP_FIRST + 1));
    const sh = makeShifter(extra);
    const expLast = EXP_LAST + extra;
    const R = (ref) => ref.replace(/\d+$/, (n) => sh.one(+n));

    // clone the row model (shifted below the expense block), adding rows if needed
    const rows = [];
    for (const r of tpl.sheet.rows) {
      const n = +r.a.r;
      rows.push({ a: { ...r.a, r: String(sh.one(n)) }, c: r.c.map((c) => {
        const x = { ...c, r: R(c.r) };
        if (x.f != null) x.f = sh.text(x.f);
        return x;
      }) });
      if (n === EXP_LAST) {
        for (let k = 1; k <= extra; k++) {
          const nr = EXP_LAST + k;
          rows.push({ a: { ...r.a, r: String(nr) }, c: r.c.map((c) => {
            const x = { ...c, r: c.r.replace(/\d+$/, String(nr)) };
            if (x.r[0] === 'A' && typeof x.v === 'number') x.v = x.v + k;
            if (x.f != null) x.f = x.f.replace(/([A-Z]{1,3})270\b/g, `$1${nr}`);
            return x;
          }) });
        }
      }
    }
    const byRef = new Map();
    for (const r of rows) for (const c of r.c) byRef.set(c.r, c);
    const rowByNum = new Map(rows.map((r) => [+r.a.r, r]));

    function setCell(ref, patch) {
      let c = byRef.get(ref);
      if (!c) {
        const rn = +ref.replace(/\D/g, ''); const row = rowByNum.get(rn);
        c = { r: ref }; row.c.push(c); row.c.sort((x, y) => colNum(x.r.replace(/\d/g, '')) - colNum(y.r.replace(/\d/g, '')));
        byRef.set(ref, c);
      }
      delete c.f; delete c.v; delete c.t; delete c.is;
      Object.assign(c, patch);
    }
    const text = (ref, s) => (s == null || s === '' ? setCell(ref, {}) : setCell(ref, { t: 'inlineStr', is: String(s) }));
    const number = (ref, n) => (n == null || n === '' || isNaN(+n) ? setCell(ref, {}) : setCell(ref, { v: +n }));

    function fillRow(rowNum, e, kind) {
      if (!e) {
        for (const col of EDIT_COLS) setCell(col + rowNum, {});
        if (kind === 'exp') setCell('I' + rowNum, { f: 'H' + rowNum + '/6' });
        return;
      }
      const d = serial(e.date);
      if (d != null) setCell('B' + rowNum, { v: d, s: kind === 'exp' ? 109 : 81 });
      else text('B' + rowNum, e.date);
      text('C' + rowNum, e.invoiceNo);
      text('D' + rowNum, e.name);
      text('E' + rowNum, e.desc);
      text('F' + rowNum, e.category);
      text('G' + rowNum, e.payment);
      number('H' + rowNum, e.gross);
      if (e.vatMode === 'sixth') setCell('I' + rowNum, { f: 'H' + rowNum + '/6' });
      else number('I' + rowNum, e.vat == null && e.gross != null ? 0 : e.vat);
      const ic = byRef.get('I' + rowNum);
      const srcRow = Math.min(rowNum, EXP_LAST);
      const baseStyle = String(tpl.sheet.rows.find((r) => +r.a.r === srcRow).c.find((c) => c.r === 'I' + srcRow)?.s ?? 76);
      ic.s = e.orange ? (tpl.orange[baseStyle] ?? tpl.orange['76']) : +baseStyle;
    }

    for (let r = SALES_FIRST; r <= SALES_LAST; r++) fillRow(r, sales[r - SALES_FIRST], 'sales');
    for (let r = EXP_FIRST; r <= expLast; r++) fillRow(r, expenses[r - EXP_FIRST], 'exp');
    if (data.period) setCell('D6', { t: 'inlineStr', is: data.period, s: byRef.get('D6').s });

    // recompute cached values
    const evalCell = evaluator((ref) => {
      const c = byRef.get(ref); if (!c) return null;
      if (c.f != null) return { f: c.f };
      if (c.t === 's') return { str: tpl.sst[c.v] };
      if (c.t === 'inlineStr') return { str: c.is };
      if (c.t === 'str') return { str: c.v };
      if (typeof c.v === 'number') return { num: c.v };
      return null;
    });

    // serialise
    const out = [tpl.sheet.prefix, '<sheetData>'];
    for (const r of rows) {
      const attrs = Object.entries(r.a).map(([k, v]) => `${k}="${esc(v)}"`).join(' ');
      if (!r.c.length) { out.push(`<row ${attrs}/>`); continue; }
      out.push(`<row ${attrs}>`);
      for (const c of r.c) {
        const s = c.s != null ? ` s="${c.s}"` : '';
        if (c.f != null) {
          const v = evalCell(c.r);
          if (typeof v === 'string') out.push(`<c r="${c.r}"${s} t="str"><f>${esc(c.f)}</f><v>${esc(v)}</v></c>`);
          else out.push(`<c r="${c.r}"${s}><f>${esc(c.f)}</f><v>${num(v)}</v></c>`);
        } else if (c.t === 'inlineStr') out.push(`<c r="${c.r}"${s} t="inlineStr"><is><t xml:space="preserve">${esc(c.is)}</t></is></c>`);
        else if (c.t === 's') out.push(`<c r="${c.r}"${s} t="s"><v>${c.v}</v></c>`);
        else if (c.t) out.push(`<c r="${c.r}"${s} t="${c.t}"><v>${esc(c.v)}</v></c>`);
        else if (typeof c.v === 'number') out.push(`<c r="${c.r}"${s}><v>${num(c.v)}</v></c>`);
        else out.push(`<c r="${c.r}"${s}/>`);
      }
      out.push('</row>');
    }
    out.push('</sheetData>', sh.xml(tpl.sheet.suffix));
    const sheetBytes = new TextEncoder().encode(out.join(''));

    const entries = tpl.parts.map((p) => p.dynamic
      ? { name: p.name, method: 0, crc: crc32(sheetBytes), size: sheetBytes.length, data: sheetBytes }
      : { name: p.name, method: 8, crc: p.crc, size: p.size, data: b64ToBytes(p.b64) });
    const totals = {
      salesGross: evalCell('H124'), salesVat: evalCell('I124'),
      expGross: evalCell('H' + (expLast + 1)), expVat: evalCell('I' + (expLast + 1)), vatDue: evalCell('I9'),
      extraRows: extra,
    };
    return { bytes: zip(entries), totals };
  }

  // Uncompressed zip of {name, data: Uint8Array} files (for check files with photos)
  function zipStore(files) {
    const enc = new TextEncoder();
    return zip(files.map((f) => {
      const data = typeof f.data === 'string' ? enc.encode(f.data) : f.data;
      return { name: f.name, method: 0, crc: crc32(data), size: data.length, data };
    }));
  }
  // Read an uncompressed (stored) zip back into {name: Uint8Array}
  function unzipStore(buf) {
    const u8 = buf instanceof Uint8Array ? buf : new Uint8Array(buf);
    const dv = new DataView(u8.buffer, u8.byteOffset, u8.byteLength);
    let eocd = -1;
    for (let i = u8.length - 22; i >= Math.max(0, u8.length - 65557); i--) if (dv.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
    if (eocd < 0) throw new Error('Not a zip file');
    const count = dv.getUint16(eocd + 10, true); let p = dv.getUint32(eocd + 16, true);
    const out = {}; const dec = new TextDecoder();
    for (let n = 0; n < count; n++) {
      if (dv.getUint32(p, true) !== 0x02014b50) throw new Error('Damaged zip file');
      const method = dv.getUint16(p + 10, true), csize = dv.getUint32(p + 20, true);
      const nlen = dv.getUint16(p + 28, true), xlen = dv.getUint16(p + 30, true), clen = dv.getUint16(p + 32, true);
      const off = dv.getUint32(p + 42, true);
      const name = dec.decode(u8.subarray(p + 46, p + 46 + nlen));
      if (method !== 0) throw new Error('This zip was re-compressed; use the original check file');
      const lnlen = dv.getUint16(off + 26, true), lxlen = dv.getUint16(off + 28, true);
      const start = off + 30 + lnlen + lxlen;
      out[name] = u8.subarray(start, start + csize);
      p += 46 + nlen + xlen + clen;
    }
    return out;
  }

  const api = { build, serial, zipStore, unzipStore };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.SunMillExport = api;
})(typeof window !== 'undefined' ? window : globalThis);
