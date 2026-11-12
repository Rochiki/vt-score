/* =====================================================================
   xlsx-stream.js — pembaca file Excel (.xlsx) besar secara streaming.
   Tidak memuat seluruh file ke memori: membaca daftar isi ZIP, lalu
   meng-inflate sheet yang dibutuhkan sepotong demi sepotong
   (DecompressionStream "deflate-raw", bawaan browser modern).
   Dipakai di Web Worker maupun Node (untuk pengujian).
   ===================================================================== */
(function (root) {
  'use strict';
  const td = new TextDecoder('utf-8');

  async function readBytes(blob, start, end) {
    return new Uint8Array(await blob.slice(start, end).arrayBuffer());
  }
  const u16 = (b, o) => b[o] | (b[o + 1] << 8);
  const u32 = (b, o) => (b[o] | (b[o + 1] << 8) | (b[o + 2] << 16) | (b[o + 3] << 24)) >>> 0;

  /** Daftar isi ZIP: { name: {method, csize, usize, lho} } */
  async function zipEntries(blob) {
    const size = blob.size;
    const tailLen = Math.min(size, 65536 + 22);
    const tail = await readBytes(blob, size - tailLen, size);
    let eocd = -1;
    for (let i = tail.length - 22; i >= 0; i--) {
      if (tail[i] === 0x50 && tail[i + 1] === 0x4b && tail[i + 2] === 0x05 && tail[i + 3] === 0x06) { eocd = i; break; }
    }
    if (eocd < 0) throw new Error('Bukan file .xlsx yang valid (struktur ZIP tidak ditemukan).');
    let cdCount = u16(tail, eocd + 10), cdSize = u32(tail, eocd + 12), cdOff = u32(tail, eocd + 16);
    // ZIP64
    if (cdOff === 0xffffffff || cdCount === 0xffff) {
      const loc = eocd - 20;
      if (loc >= 0 && u32(tail, loc) === 0x07064b50) {
        const z64off = Number(new DataView(tail.buffer, tail.byteOffset + loc + 8, 8).getBigUint64(0, true));
        const z = await readBytes(blob, z64off, z64off + 56);
        const dv = new DataView(z.buffer, z.byteOffset, z.byteLength);
        cdCount = Number(dv.getBigUint64(32, true)); cdSize = Number(dv.getBigUint64(40, true)); cdOff = Number(dv.getBigUint64(48, true));
      }
    }
    const cd = await readBytes(blob, cdOff, cdOff + cdSize);
    const out = {};
    let p = 0;
    for (let i = 0; i < cdCount && p + 46 <= cd.length; i++) {
      if (u32(cd, p) !== 0x02014b50) break;
      const method = u16(cd, p + 10);
      let csize = u32(cd, p + 20), usize = u32(cd, p + 24);
      const nlen = u16(cd, p + 28), xlen = u16(cd, p + 30), clen = u16(cd, p + 32);
      let lho = u32(cd, p + 42);
      const name = td.decode(cd.subarray(p + 46, p + 46 + nlen));
      if (csize === 0xffffffff || usize === 0xffffffff || lho === 0xffffffff) {
        let x = p + 46 + nlen; const xe = x + xlen;
        while (x + 4 <= xe) {
          const id = u16(cd, x), sz = u16(cd, x + 2);
          if (id === 1) {
            const dv = new DataView(cd.buffer, cd.byteOffset + x + 4, sz); let q = 0;
            if (usize === 0xffffffff) { usize = Number(dv.getBigUint64(q, true)); q += 8; }
            if (csize === 0xffffffff) { csize = Number(dv.getBigUint64(q, true)); q += 8; }
            if (lho === 0xffffffff) { lho = Number(dv.getBigUint64(q, true)); q += 8; }
          }
          x += 4 + sz;
        }
      }
      out[name] = { method, csize, usize, lho };
      p += 46 + nlen + xlen + clen;
    }
    return out;
  }

  async function entryStream(blob, e) {
    const h = await readBytes(blob, e.lho, e.lho + 30);
    if (u32(h, 0) !== 0x04034b50) throw new Error('Header ZIP rusak');
    const start = e.lho + 30 + u16(h, 26) + u16(h, 28);
    const raw = blob.slice(start, start + e.csize).stream();
    if (e.method === 0) return raw;
    if (e.method !== 8) throw new Error('Metode kompresi ZIP tidak didukung: ' + e.method);
    if (typeof DecompressionStream !== 'undefined' && !api._forceJS) {
      try { return raw.pipeThrough(new DecompressionStream('deflate-raw')); } catch (e) { /* lanjut ke cadangan */ }
    }
    // Cadangan untuk browser lama: inflate JavaScript murni (lib/fflate.min.js)
    const FF = (typeof fflate !== 'undefined' && fflate) || (typeof self !== 'undefined' && self.fflate) || (typeof require === 'function' ? (() => { try { return require('../lib/fflate.min.js'); } catch (e) { return null; } })() : null);
    if (!FF) throw new Error('Browser ini belum mendukung pembacaan file Excel dan pustaka cadangan (lib/fflate.min.js) tidak termuat. Perbarui Chrome/Edge ke versi 103+ atau muat ulang halaman.');
    const rd = raw.getReader(); const infl = new FF.Inflate();
    let q = [], done = false;
    infl.ondata = (d, fin) => { if (d && d.length) q.push(d); if (fin) done = true; };
    return new ReadableStream({
      async pull(ctrl) {
        while (!q.length && !done) {
          const r = await rd.read();
          if (r.done) { try { infl.push(new Uint8Array(0), true); } catch (e) { /* sudah selesai */ } done = true; break; }
          infl.push(r.value);
        }
        if (q.length) ctrl.enqueue(q.shift()); else ctrl.close();
      },
      cancel() { try { rd.cancel(); } catch (e) { /* */ } }
    });
  }

  /** Baca isi entri sebagai teks; onText(chunkText) dipanggil berurutan bila diberikan. */
  async function entryText(blob, e, onText, onProgress) {
    const reader = (await entryStream(blob, e)).getReader();
    const dec = new TextDecoder('utf-8');
    let all = onText ? null : [];
    let got = 0;
    for (;;) {
      const { value, done } = await reader.read();
      if (done) break;
      got += value.length;
      const t = dec.decode(value, { stream: true });
      if (onText) onText(t); else all.push(t);
      if (onProgress) onProgress(got, e.usize);
    }
    const last = dec.decode();
    if (onText) { if (last) onText(last); return null; }
    all.push(last);
    return all.join('');
  }

  const ENT = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" };
  function unesc(s) {
    if (s.indexOf('&') < 0) return s;
    return s.replace(/&(#x[0-9a-fA-F]+|#[0-9]+|amp|lt|gt|quot|apos);/g, (m, g) =>
      g[0] === '#' ? String.fromCodePoint(g[1] === 'x' ? parseInt(g.slice(2), 16) : parseInt(g.slice(1), 10)) : ENT[g]);
  }
  function colIndex(ref) {
    let n = 0;
    for (let i = 0; i < ref.length; i++) {
      const c = ref.charCodeAt(i);
      if (c >= 65 && c <= 90) n = n * 26 + (c - 64); else if (c >= 97 && c <= 122) n = n * 26 + (c - 96); else break;
    }
    return n - 1;
  }
  function textOf(xml) {
    // gabungkan semua <t>…</t> (rich text) — abaikan teks fonetik <rPh>
    xml = xml.replace(/<rPh\b[\s\S]*?<\/rPh>/g, '');
    let out = '', m; const re = /<t(?:\s[^>]*)?>([\s\S]*?)<\/t>|<t(?:\s[^>]*)?\/>/g;
    while ((m = re.exec(xml))) out += m[1] ? unesc(m[1]) : '';
    return out;
  }

  /** Buka workbook: daftar sheet (nama → path) dan shared strings. */
  async function openWorkbook(blob) {
    const ents = await zipEntries(blob);
    const get = async n => (ents[n] ? entryText(blob, ents[n]) : null);
    const wbx = await get('xl/workbook.xml');
    if (!wbx) throw new Error('File ini bukan workbook Excel (.xlsx).');
    const rels = (await get('xl/_rels/workbook.xml.rels')) || '';
    const rmap = {}; let m;
    const rre = /<Relationship\b([^>]*)\/?>/g;
    while ((m = rre.exec(rels))) {
      const id = /\bId="([^"]+)"/.exec(m[1]), tg = /\bTarget="([^"]+)"/.exec(m[1]);
      if (id && tg) { let t = tg[1]; t = t.startsWith('/') ? t.slice(1) : 'xl/' + t.replace(/^\.\//, ''); rmap[id[1]] = t; }
    }
    const sheets = []; const sre = /<sheet\b([^>]*)\/?>/g;
    while ((m = sre.exec(wbx))) {
      const nm = /\bname="([^"]*)"/.exec(m[1]), rid = /\br:id="([^"]+)"/.exec(m[1]) || /\bid="([^"]+)"/.exec(m[1]);
      if (nm && rid && rmap[rid[1]]) sheets.push({ name: unesc(nm[1]), path: rmap[rid[1]] });
    }
    const date1904 = /date1904="(1|true)"/.test(wbx);
    let sst = [];
    const sstx = ents['xl/sharedStrings.xml'] ? await get('xl/sharedStrings.xml') : null;
    if (sstx) {
      const re = /<si>([\s\S]*?)<\/si>|<si\/>/g;
      while ((m = re.exec(sstx))) sst.push(m[1] ? textOf(m[1]) : '');
    }
    return { blob, ents, sheets, sst, date1904 };
  }

  /**
   * Stream baris sheet. onRow(rowNumber1, values[]) — values berisi angka (number),
   * teks (string), boolean, atau undefined. Kembalikan false dari onRow untuk berhenti.
   */
  async function readRows(wb, sheet, onRow, onProgress) {
    const e = wb.ents[sheet.path];
    if (!e) throw new Error('Sheet tidak ditemukan: ' + sheet.name);
    let buf = '', stop = false, auto = 0;
    const sst = wb.sst;
    const cre = /<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g;
    const handleRow = (attrs, inner) => {
      const rm = /\br="(\d+)"/.exec(attrs);
      const rn = rm ? +rm[1] : auto + 1; auto = rn;
      const vals = [];
      if (inner) {
        let m, seq = 0; cre.lastIndex = 0;
        while ((m = cre.exec(inner))) {
          const a = m[1], body = m[2];
          const r = /\br="([A-Z]+)\d*"/.exec(a);
          const ci = r ? colIndex(r[1]) : seq; seq = ci + 1;
          if (!body) continue;
          const tm = /\bt="(\w+)"/.exec(a); const t = tm ? tm[1] : 'n';
          let v;
          if (t === 'inlineStr') v = textOf(body);
          else {
            const vm = /<v>([\s\S]*?)<\/v>/.exec(body);
            if (!vm) continue;
            const raw = vm[1];
            if (t === 's') v = sst[+raw];
            else if (t === 'str' || t === 'e') v = unesc(raw);
            else if (t === 'b') v = raw === '1';
            else { v = +raw; if (Number.isNaN(v)) v = unesc(raw); }
          }
          vals[ci] = v;
        }
      }
      if (onRow(rn, vals) === false) stop = true;
    };
    const rre = /<row\b([^>]*?)(?:\/>|>([\s\S]*?)<\/row>)/g;
    await entryText(wb.blob, e, chunk => {
      if (stop) return;
      buf += chunk;
      const lastEnd = buf.lastIndexOf('</row>');
      const lastSelf = buf.lastIndexOf('/>');
      let cut = lastEnd >= 0 ? lastEnd + 6 : -1;
      if (cut < 0) { if (buf.length > 8e6 && lastSelf > 0) cut = 0; else return; }
      const part = buf.slice(0, cut); buf = buf.slice(cut);
      let m; rre.lastIndex = 0;
      while (!stop && (m = rre.exec(part))) handleRow(m[1], m[2]);
    }, onProgress);
    if (!stop && buf) { let m; rre.lastIndex = 0; while (!stop && (m = rre.exec(buf))) handleRow(m[1], m[2]); }
  }

  /** Baca sheet kecil sekaligus jadi array of arrays (maks maxRows). */
  async function readSheetAll(wb, sheet, maxRows) {
    const rows = [];
    await readRows(wb, sheet, (rn, v) => { rows[rn - 1] = v; if (maxRows && rn >= maxRows) return false; });
    for (let i = 0; i < rows.length; i++) if (!rows[i]) rows[i] = [];
    return rows;
  }

  /** Konversi serial tanggal Excel → {y,m,d,hh,mi,ss, day(serial hari)} */
  function serialToParts(s, date1904) {
    if (date1904) s += 1462;
    const day = Math.floor(s + 1e-9);
    let secs = Math.round((s - day) * 86400);
    let dd = day;
    if (secs >= 86400) { secs -= 86400; dd += 1; }
    const ms = (dd - 25569) * 86400000;
    const dt = new Date(ms);
    return { y: dt.getUTCFullYear(), m: dt.getUTCMonth() + 1, d: dt.getUTCDate(), hh: Math.floor(secs / 3600), mi: Math.floor(secs % 3600 / 60), ss: secs % 60, day: dd, serial: dd + secs / 86400 };
  }

  const api = { zipEntries, entryText, openWorkbook, readRows, readSheetAll, serialToParts, colIndex, unesc };
  root.XLSXStream = api;
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
})(typeof self !== 'undefined' ? self : globalThis);
