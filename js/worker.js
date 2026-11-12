/* worker.js — membaca file unggahan di latar belakang agar layar tidak macet. */
importScripts('../lib/fflate.min.js', 'xlsx-stream.js', 'engine.js', 'db.js');

function post(type, data) { self.postMessage(Object.assign({ type }, data)); }

async function mergeCartrack(trips, stats) {
  const byLbl = {};
  trips.forEach(t => { (byLbl[t.bw] = byLbl[t.bw] || []).push(t); });
  const labels = Object.keys(stats.byPeriod);
  const out = []; let added = 0, dup = 0;
  for (const lbl of labels) {
    const ex = await DB.getPeriod(lbl);
    const nw = byLbl[lbl] || []; const st = stats.byPeriod[lbl];
    if (!ex) { out.push({ lbl, stats: { rows: st.rows, idle: st.idle, drive: nw.length, minDay: st.minDay, maxDay: st.maxDay }, trips: nw }); added += nw.length; continue; }
    const seen = new Set(ex.trips.map(t => t.k)); const merged = ex.trips.slice();
    for (const t of nw) { if (seen.has(t.k)) { dup++; continue; } seen.add(t.k); merged.push(t); added++; }
    out.push({ lbl, stats: { rows: Math.max(ex.stats.rows, st.rows), idle: Math.max(ex.stats.idle, st.idle), drive: merged.length, minDay: Math.min(ex.stats.minDay, st.minDay), maxDay: Math.max(ex.stats.maxDay, st.maxDay) }, trips: merged });
  }
  await DB.putPeriods(out);
  return { added, dup, labels };
}

async function handleCartrack(id, name, r) {
  try {
    const m = await mergeCartrack(r.trips, r.stats);
    post('done', { id, name, kind: 'cartrack', stats: { rows: r.stats.rows, drive: r.stats.drive, idle: r.stats.idle, periods: m.labels, added: m.added, dup: m.dup, sheet: r.stats.sheet, headerRow: r.stats.headerRow } });
  } catch (e) {
    // Penyimpanan browser gagal (mode privat / penuh): kirim trip ke halaman agar tetap bisa dipakai sesi ini.
    post('done', { id, name, kind: 'cartrack-inline', trips: r.trips, stats: { rows: r.stats.rows, drive: r.stats.drive, idle: r.stats.idle, byPeriod: r.stats.byPeriod, dbError: e && e.message ? e.message : String(e) } });
  }
}

self.onmessage = async (ev) => {
  const { id, cmd, file } = ev.data;
  try {
    if (cmd !== 'import') throw new Error('Perintah tidak dikenal');
    const name = file.name || 'file';
    const prog = (a, b) => post('progress', { id, name, pct: b ? Math.min(99, Math.round(a / b * 100)) : 0 });
    if (/\.csv$/i.test(name) || /\.txt$/i.test(name)) {
      const head = (await file.slice(0, 65536).text()).toLowerCase();
      if (head.indexOf('nama did') >= 0) return post('done', { id, name, kind: 'master', master: VT.importMasterRows(VT.csvRows(await file.text())) });
      if (head.indexOf('mileage') < 0 && head.indexOf('remark') < 0 && (head.indexOf('speed') >= 0 || /speed|overspeed|speeding|violat|pelanggaran/i.test(name)) && (head.indexOf('duration') >= 0 || head.indexOf('durasi') >= 0 || head.indexOf('end') >= 0 || head.indexOf('selesai') >= 0)) {
        const r = VT.importOsDetailRows(VT.csvRows(await file.text()), name);
        return post('done', { id, name, kind: 'osdetail', os: r });
      }
      const r = await VT.importCartrackCsv(file, prog);
      return handleCartrack(id, name, r);
    }
    if (/\.json$/i.test(name)) {
      const txt = await file.text();
      return post('done', { id, name, kind: 'backup', json: JSON.parse(txt) });
    }
    if (!/\.(xlsx|xlsm)$/i.test(name)) throw new Error('Jenis file "' + (name.split('.').pop() || '?') + '" tidak didukung. Gunakan .xlsx, .csv, .txt, atau cadangan .json. File .xls (Excel lama) harap disimpan ulang sebagai .xlsx.');
    const det = await VT.detectXlsx(file);
    if (det.type === 'cartrack') {
      const r = await VT.importCartrackXlsx(det, prog);
      return handleCartrack(id, name, r);
    }
    if (det.type === 'ga') {
      const r = await VT.importGA(det);
      return post('done', { id, name, kind: 'ga', ga: r });
    }
    if (det.type === 'master') {
      const rows = await VT.importMasterXlsx(det);
      return post('done', { id, name, kind: 'master', master: rows });
    }
    if (det.type === 'osdetail') {
      const r = await VT.importOsDetailXlsx(det);
      return post('done', { id, name, kind: 'osdetail', os: r });
    }
    throw new Error('Jenis file tidak dikenali. Yang didukung: data mentah Cartrack (Data Score Behavior), lampiran GA "Data Case", master karyawan, detail kejadian overspeed (laporan speeding Cartrack), atau file cadangan .json.');
  } catch (e) {
    post('error', { id, name: file && file.name, message: e && e.message ? e.message : String(e) });
  }
};
