/* export-all.js — "Unduh semua (Excel)": satu workbook berisi seluruh laporan aplikasi.
   Angka disimpan sebagai angka asli; Excel berlokal Indonesia otomatis menampilkan koma desimal. */
'use strict';
async function exportAll(pid) {
  if (!S.res) { toast('Belum ada data.'); return; }
  busy('Menyusun workbook Excel lengkap…');
  await new Promise(r => setTimeout(r, 30));
  try {
    await loadLib('lib/exceljs.min.js');
    const res = S.res, Pp = res.P, p = per(pid), master = res.master;
    const wb = new ExcelJS.Workbook(); wb.creator = 'QSHE Department – PT Hexindo Adiperkasa Tbk'; wb.created = new Date();
    const BLACK = 'FF111111', OR = 'FFF7941D', LIGHT = 'FFF4F4F4';
    const titleRows = (ws, title, sub, ncol) => {
      ws.addRow([title]); ws.addRow([sub || '']);
      ws.mergeCells(1, 1, 1, Math.max(2, ncol)); ws.mergeCells(2, 1, 2, Math.max(2, ncol));
      const t = ws.getCell(1, 1); t.font = { bold: true, size: 13, color: { argb: 'FFFFFFFF' } }; t.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: BLACK } }; t.alignment = { vertical: 'middle', indent: 1 };
      ws.getRow(1).height = 26;
      const s = ws.getCell(2, 1); s.font = { italic: true, size: 9, color: { argb: 'FF595959' } }; s.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: LIGHT } };
      ws.getCell(2, 1).border = { left: { style: 'thick', color: { argb: OR } } };
    };
    /** cols: [label, getter(row), numFmt?, width?] */
    const sheet = (name, title, sub, cols, rows) => {
      const ws = wb.addWorksheet(name.slice(0, 31), { views: [{ state: 'frozen', ySplit: 3 }] });
      titleRows(ws, title, sub, cols.length);
      const hr = ws.addRow(cols.map(c => c[0]));
      hr.eachCell(c => { c.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 9 }; c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1A1A1A' } }; c.alignment = { wrapText: true, vertical: 'middle' }; c.border = { bottom: { style: 'medium', color: { argb: OR } } }; });
      hr.height = 32;
      for (const r of rows) {
        ws.addRow(cols.map(c => { let v; try { v = c[1](r); } catch (e) { v = ''; } if (typeof v === 'number' && !isFinite(v)) v = null; return v == null ? '' : v; }));
      }
      cols.forEach((c, i) => { const col = ws.getColumn(i + 1); col.width = c[3] || Math.min(40, Math.max(9, String(c[0]).length + 3)); if (c[2]) col.numFmt = c[2]; });
      if (rows.length) ws.autoFilter = { from: { row: 3, column: 1 }, to: { row: 3, column: cols.length } };
      return ws;
    };
    const N0 = '#,##0', N1 = '0.0', N2 = '0.00', PC = '0%', PC1 = '0.0%';
    const nameP = x => (per(x) || {}).lbl || '';
    const k = VT.kpis(res, pid, 'Semua cabang');
    const nts = VT.buildNotifications(res, pid, Pp, master, res.tripsByKey);
    const stamp = 'Periode ' + p.lbl + ' (' + VT.dtxt(p.start) + ' – ' + VT.dtxt(p.end) + ') · dibuat ' + new Date().toLocaleString('id-ID') + ' · Vehicle Telematic Score and Report v' + VT.VERSION;

    // 1. DASHBOARD (halaman depan) — sama dengan dashboard di aplikasi, termasuk 4 grafik
    const exI = VTR.exec(res, pid, 'Semua cabang', 'id', S), exE = VTR.exec(res, pid, 'Semua cabang', 'en', S);
    const dash = wb.addWorksheet('Dashboard', { properties: { tabColor: { argb: OR } }, views: [{ showGridLines: false }] });
    const dashFail = [];
    {
      const NC = 12; for (let i = 1; i <= NC; i++) dash.getColumn(i).width = 12.5;
      const solid = argb => ({ type: 'pattern', pattern: 'solid', fgColor: { argb } });
      const WH = 'FFFFFFFF', INK = 'FF111111';
      const band = (row, c1, c2, val, o) => {
        o = o || {};
        if (c2 > c1) dash.mergeCells(row, c1, row, c2);
        for (let c = c1; c <= c2; c++) { const x = dash.getCell(row, c); if (o.fill) x.fill = solid(o.fill); if (o.border) x.border = o.border; }
        const c = dash.getCell(row, c1); c.value = val === undefined || val === null ? '' : val;
        c.font = Object.assign({ size: 10, color: { argb: INK } }, o.font || {});
        c.alignment = Object.assign({ vertical: 'middle', wrapText: true }, o.al || {});
        if (o.fmt) c.numFmt = o.fmt;
        if (o.h) dash.getRow(row).height = o.h;
        return c;
      };
      const LEFTOR = { left: { style: 'thick', color: { argb: OR } } };
      const secRow = (row, text) => band(row, 1, NC, text, { font: { bold: true, size: 11 }, fill: LIGHT, h: 22, border: LEFTOR, al: { indent: 1 } });
      const est = (t, per) => Math.max(18, 14.5 * Math.ceil(String(t).length / per) + 4);
      band(1, 1, NC, '', { fill: 'FF0B0B0B', h: 252 });
      band(2, 1, NC, '', { fill: OR, h: 4 });
      band(3, 1, NC, stamp, { font: { size: 9, italic: true, color: { argb: 'FF595959' } }, fill: LIGHT, h: 18, border: LEFTOR, al: { indent: 1 } });
      secRow(5, 'RINGKASAN · ' + p.lbl.toUpperCase() + ' (' + VT.dtxt(p.start) + ' – ' + VT.dtxt(p.end) + ') · NASIONAL');
      const WHB = { right: { style: 'thick', color: { argb: WH } } };
      const card = (row, c, label, val, sub, fmt, acc) => {
        const fg = acc ? INK : WH, bg = acc ? OR : 'FF1A1A1A';
        band(row, c, c + 2, label, { font: { size: 8, bold: true, color: { argb: acc ? INK : OR } }, fill: bg, h: 20, border: WHB, al: { indent: 1 } });
        band(row + 1, c, c + 2, val, { font: { size: 24, bold: true, color: { argb: fg } }, fill: bg, h: 40, fmt, border: WHB, al: { indent: 1, horizontal: 'left' } });
        band(row + 2, c, c + 2, sub, { font: { size: 8, color: { argb: fg } }, fill: bg, h: 32, border: WHB, al: { indent: 1, vertical: 'top' } });
      };
      const dl = k.finPrev != null && k.fin != null ? k.fin - k.finPrev : null;
      const cards = [
        ['SKOR RATA-RATA (METODE BARU)', k.fin, 'metode lama ' + VT.fx(k.old, 1) + (dl != null ? ' · ' + (dl >= 0 ? '▲ ' : '▼ ') + VT.fx(Math.abs(dl), 1) + ' vs periode lalu' : ''), '0.0', true],
        ['PENGEMUDI AKTIF', k.n, VT.fx(k.km, 0) + ' km · rata-rata ' + VT.fx(k.kmPer, 0) + ' km/orang', '#,##0'],
        ['GRADE A–B', k.ab, k.de + ' orang grade D/E perlu pembinaan', '0%'],
        ['NOTIFIKASI KE ATASAN', k.notif, k.esk + ' eskalasi · lihat sheet Notifikasi', '0'],
        ['PELANGGARAN BERULANG (2 MINGGU)', k.rep, (k.n ? VT.fx(k.rep / k.n * 100, 0) : 0) + '% dari pengemudi aktif', '0'],
        ['OVERSPEED STATUS MERAH', k.merah, 'KUNING ' + k.kuning + ' · count dihitung ' + VT.fx(k.osv, 0), '0'],
        ['NIGHT DRIVING (JAM)', k.nh, VT.fx(k.nhPct, 1) + '% jam mengemudi · mulai ' + String(Pp.NSTART).padStart(2, '0') + ':00', '#,##0'],
        ['HARSH BRAKING / 1.000 KM', k.brk, 'periode lalu ' + (k.brkPrev == null ? '–' : VT.fx(k.brkPrev, 2)), '0.00']];
      cards.forEach((c, i) => card(6 + Math.floor(i / 4) * 4, 1 + (i % 4) * 3, c[0], c[1], c[2], c[3], c[4]));
      let r = 13;
      r += 2; secRow(r, 'GRAFIK');
      r++; band(r, 1, 6, 'Tren skor rata-rata per periode (garis oranye = baru, abu putus-putus = lama)', { font: { bold: true, size: 9 }, h: 18, al: { indent: 1 } });
      band(r, 7, 12, 'Sebaran grade · periode terpilih', { font: { bold: true, size: 9 }, al: { indent: 1 } });
      const gy = { grid: { color: '#e5e5e5' }, border: { display: false } }, gx = { grid: { display: false } };
      const tr = VT.trend(res, res.lastPid, 'Semua cabang', 11);
      const png = async cfg => {
        const cv = document.createElement('canvas'); cv.width = 545; cv.height = 262;
        cfg.options = Object.assign({ maintainAspectRatio: false }, cfg.options, { responsive: false, animation: false, devicePixelRatio: 2 });
        cfg.plugins = [{ id: 'bg', beforeDraw(c) { const x = c.ctx; x.save(); x.globalCompositeOperation = 'destination-over'; x.fillStyle = '#ffffff'; x.fillRect(0, 0, c.width, c.height); x.restore(); } }];
        const c = new Chart(cv, cfg); c.draw(); const u = cv.toDataURL('image/png'); c.destroy(); return u;
      };
      const C1 = { type: 'line', data: { labels: tr.map(x => x.lbl), datasets: [
        { label: 'Skor metode lama', data: tr.map(x => x.old), borderColor: '#9e9e9e', borderDash: [6, 4], borderWidth: 1.75, pointRadius: 2, tension: 0 },
        { label: 'Skor baru', data: tr.map(x => x.fin), borderColor: '#f7941d', backgroundColor: '#f7941d', borderWidth: 3, pointRadius: 3, tension: 0 }] },
        options: { scales: { y: Object.assign({ min: 60, max: 100 }, gy), x: gx }, plugins: { legend: { position: 'bottom' } } } };
      const C2 = { type: 'bar', data: { labels: ['A', 'B', 'C', 'D', 'E'], datasets: [{ label: 'Pengemudi', data: k.grades, backgroundColor: ['#111111', '#6b6b6b', '#f9c38a', '#f7941d', '#c2410c'], borderRadius: 3, maxBarThickness: 64 }] },
        options: { plugins: { legend: { display: false } }, scales: { y: Object.assign({ beginAtZero: true, ticks: { precision: 0 } }, gy), x: gx } } };
      const C3 = { type: 'line', data: { labels: tr.map(x => x.lbl), datasets: [
        { label: 'Overspeed dihitung /1.000 km', data: tr.map(x => x.osv), borderColor: '#f7941d', borderWidth: 3, pointRadius: 2, tension: 0 },
        { label: 'Harsh braking /1.000 km', data: tr.map(x => x.brk), borderColor: '#111111', borderWidth: 2.25, pointRadius: 2, tension: 0 },
        { label: 'Harsh accel /1.000 km', data: tr.map(x => x.acc), borderColor: '#9e9e9e', borderWidth: 2, pointRadius: 2, tension: 0 }] },
        options: { scales: { y: Object.assign({ beginAtZero: true }, gy), x: gx }, plugins: { legend: { position: 'bottom' } } } };
      const C4 = { type: 'bar', data: { labels: k.bands.map(b => b.band), datasets: [
        { label: 'Skor metode lama', data: k.bands.map(b => b.old || null), backgroundColor: '#bdbdbd', borderRadius: 3, maxBarThickness: 40 },
        { label: 'Skor baru', data: k.bands.map(b => b.fin || null), backgroundColor: '#f7941d', borderRadius: 3, maxBarThickness: 40 }] },
        options: { scales: { y: Object.assign({ min: 60, max: 100 }, gy), x: gx }, plugins: { legend: { position: 'bottom' } } } };
      const heroPng = async () => {
        const W = 1110, H = 335, DPR = 2; const cv = document.createElement('canvas'); cv.width = W * DPR; cv.height = H * DPR; const x = cv.getContext('2d'); x.scale(DPR, DPR);
        try { await Promise.all(['700 46px Poppins', '600 12px Barlow', '400 16px Barlow'].map(f => document.fonts.load(f))); } catch (e) {}
        const img = await new Promise((ok, no) => { const i = new Image(); i.onload = () => ok(i); i.onerror = () => no(new Error('assets/hero.jpg tidak termuat')); i.src = 'assets/hero.jpg'; });
        x.fillStyle = '#0b0b0b'; x.fillRect(0, 0, W, H);
        const sc = Math.max(W / img.width, H / img.height), dw = img.width * sc, dh = img.height * sc; x.drawImage(img, W - dw, (H - dh) / 2, dw, dh);
        const g = x.createLinearGradient(0, 0, W, 0); g.addColorStop(0, 'rgba(10,10,10,.96)'); g.addColorStop(.36, 'rgba(10,10,10,.82)'); g.addColorStop(.62, 'rgba(10,10,10,0)'); x.fillStyle = g; x.fillRect(0, 0, W, H);
        const rr = (X, Y, w, h, r_) => { x.beginPath(); x.moveTo(X + r_, Y); x.arcTo(X + w, Y, X + w, Y + h, r_); x.arcTo(X + w, Y + h, X, Y + h, r_); x.arcTo(X, Y + h, X, Y, r_); x.arcTo(X, Y, X + w, Y, r_); x.closePath(); };
        const dot = (X, Y) => { x.fillStyle = '#f7941d'; x.beginPath(); x.arc(X, Y, 3.5, 0, 7); x.fill(); };
        const SANS = "'Segoe UI', Barlow, system-ui, sans-serif";
        x.textBaseline = 'middle';
        const pill = 'QSHE DEPARTMENT · PT HEXINDO ADIPERKASA TBK' + (p.fy ? ' · FY' + p.fy : '');
        x.font = '600 12px ' + SANS; const pw = x.measureText(pill).width + 14 * 2 + 15;
        rr(40, 30, pw, 26, 13); x.strokeStyle = '#f7941d'; x.lineWidth = 1.5; x.stroke(); dot(40 + 14 + 3.5, 43); x.fillStyle = '#fff'; x.fillText(pill, 40 + 14 + 15, 43.5);
        x.textBaseline = 'alphabetic'; x.fillStyle = '#fff'; x.font = '700 46px Poppins, ' + SANS; x.fillText('Vehicle Telematic', 40, 118);
        x.fillText('Score and Report', 40, 172); const w1 = x.measureText('Score and Report').width; x.fillStyle = '#f7941d'; x.fillText('.', 40 + w1, 172);
        x.fillStyle = '#d8d8d8'; x.font = '400 16px ' + SANS; x.fillText('Adil untuk travel load tinggi. Tegas untuk pelanggaran berulang.', 40, 204);
        let cx = 40; ['Travel load 5 band', 'Berulang 2 minggu', 'Insiden & pemimpin', 'Notifikasi atasan'].forEach(t => {
          x.font = '400 13px ' + SANS; const w = x.measureText(t).width + 13 * 2 + 15; rr(cx, 218, w, 28, 14); x.fillStyle = 'rgba(255,255,255,.07)'; x.fill(); x.strokeStyle = 'rgba(255,255,255,.14)'; x.lineWidth = 1; x.stroke();
          dot(cx + 13 + 3.5, 232); x.fillStyle = '#fff'; x.textBaseline = 'middle'; x.fillText(t, cx + 13 + 15, 232.5); x.textBaseline = 'alphabetic'; cx += w + 8; });
        const first = res.periods[0], mon = d => VT.MON[VT.dayToDate(d).getUTCMonth()];
        x.fillStyle = '#8a8a8a'; x.font = '400 12px ' + SANS; x.fillText('Data Vehicle Telematics ' + mon(first.start) + ' ' + VT.dayToDate(first.start).getUTCFullYear() + '–' + mon(p.end) + ' ' + VT.dayToDate(p.end).getUTCFullYear() + ' · skoring berimbang travel load', 40, 296);
        const night = String(Pp.NSTART).padStart(2, '0') + ':00 – ' + String((Pp.NEND + 23) % 24).padStart(2, '0') + ':59';
        const co = (t1, t2, left, right, topPct) => {
          x.font = '700 12.5px ' + SANS; const w1_ = x.measureText(t1).width; x.font = '400 11.5px ' + SANS; const w2_ = x.measureText(t2).width, w = Math.max(w1_, w2_);
          const X = left != null ? W * left : W - W * right - w, Y = H * topPct;
          x.fillStyle = '#fff'; x.font = '700 12.5px ' + SANS; x.fillText(t1, X, Y + 12); x.fillStyle = '#cfcfcf'; x.font = '400 11.5px ' + SANS; x.fillText(t2, X, Y + 28);
          x.strokeStyle = '#f7941d'; x.lineWidth = 1; x.beginPath(); x.moveTo(X, Y + 34); x.lineTo(X + w, Y + 34); x.stroke(); };
        co('TRAVEL LOAD', 'dibanding band km sejenis', .484, null, .10); co('NIGHT DRIVING', 'mulai ' + night, .697, null, .07);
        co('OVERSPEED', 'toleransi ≤' + Pp.OSTOL + ' count/trip', null, .012, .21); co('HARSH BRAKING', 'laju per 1.000 km', null, .012, .82);
        return cv.toDataURL('image/png');
      };
      try { const hid = wb.addImage({ base64: await heroPng(), extension: 'png' }); dash.addImage(hid, { tl: { col: 0, row: 0 }, ext: { width: 1110, height: 335 } }); }
      catch (e) { console.error('[VT] hero Excel', e); dashFail.push('hero'); band(1, 1, NC, 'VEHICLE TELEMATIC SCORE AND REPORT', { font: { size: 22, bold: true, color: { argb: WH } }, fill: 'FF0B0B0B', h: 252, al: { indent: 2 } }); }
      const place = async (cfg, col, row, name) => {
        try { const id = wb.addImage({ base64: await png(cfg), extension: 'png' }); dash.addImage(id, { tl: { col, row }, ext: { width: 545, height: 262 } }); }
        catch (e) { console.error('[VT] grafik Excel ' + name, e); dashFail.push(name); band(row + 1, col + 1, col + 6, '(grafik "' + name + '" tidak dapat disisipkan: ' + e.message + ')', { font: { italic: true, color: { argb: 'FFC2410C' } } }); }
      };
      await place(C1, 0, r, 'Tren skor'); await place(C2, 6, r, 'Sebaran grade');
      r += 14; band(r, 1, 6, 'Laju pelanggaran per 1.000 km', { font: { bold: true, size: 9 }, h: 18, al: { indent: 1 } }); band(r, 7, 12, 'Keadilan travel load · skor per band', { font: { bold: true, size: 9 }, al: { indent: 1 } });
      await place(C3, 0, r, 'Laju pelanggaran'); await place(C4, 6, r, 'Skor per band');
      r += 14;
      const hdr = (row, spec) => spec.forEach(([c1, c2, t]) => band(row, c1, c2, t, { font: { bold: true, size: 9, color: { argb: WH } }, fill: 'FF1A1A1A', h: 26, border: { bottom: { style: 'medium', color: { argb: OR } } }, al: { horizontal: 'center' } }));
      const bot = { bottom: { style: 'thin', color: { argb: 'FFE5E5E5' } } };
      const cell = (row, c1, c2, v, o) => band(row, c1, c2, v, Object.assign({ border: bot, font: { size: 9 } }, o || {}));
      r++; secRow(r, '10 PENGEMUDI PRIORITAS COACHING · skor terendah periode ini');
      r++; hdr(r, [[1, 1, '#'], [2, 3, 'Nama pengemudi'], [4, 5, 'Cabang'], [6, 6, 'Band'], [7, 7, 'KM'], [8, 8, 'Skor'], [9, 9, 'Grade'], [10, 10, 'Status overspeed'], [11, 11, 'Berulang'], [12, 12, 'Fokus / tindak lanjut']]);
      VT.sel(res, pid, 'Semua cabang').slice().sort((a, b) => a.FIN - b.FIN || (a.drv < b.drv ? -1 : 1)).slice(0, 10).forEach((x, i) => {
        r++; const ctr = { al: { horizontal: 'center' } };
        cell(r, 1, 1, i + 1, ctr); cell(r, 2, 3, x.drv, { font: { bold: true, size: 9 } }); cell(r, 4, 5, x.cab); cell(r, 6, 6, x.band, ctr); cell(r, 7, 7, x.km, { fmt: '#,##0' });
        cell(r, 8, 8, x.FIN, { fmt: '0.0', font: { bold: true, size: 9 } }); cell(r, 9, 9, x.GR, ctr); cell(r, 10, 10, x.OSS, ctr); cell(r, 11, 11, x.NREP > 0 ? x.NREP + ' komponen' : '–', ctr);
        cell(r, 12, 12, (x.NOTIF > 0 ? 'NOTIF ATASAN · ' : '') + x.WK, { font: { size: 8 }, h: 26 });
      });
      r += 2; const br = VT.branches(res, pid, Pp).filter(b => b.n >= Pp.MIN_BRANCH);
      band(r, 1, 6, 'CABANG TERBAIK · periode ini', { font: { bold: true, size: 11 }, fill: LIGHT, h: 22, border: LEFTOR, al: { indent: 1 } }); band(r, 7, 12, 'CABANG PERLU PERHATIAN · periode ini', { font: { bold: true, size: 11 }, fill: LIGHT, border: LEFTOR, al: { indent: 1 } });
      r++; hdr(r, [[1, 1, '#'], [2, 3, 'Cabang'], [4, 4, 'Pengemudi'], [5, 5, 'Skor'], [6, 6, '% D/E'], [7, 7, '#'], [8, 9, 'Cabang'], [10, 10, 'Pengemudi'], [11, 11, 'Skor'], [12, 12, '% D/E']]);
      const bestB = br.slice().sort((a, b) => b.fin - a.fin).slice(0, 5), worstB = br.slice().sort((a, b) => a.fin - b.fin).slice(0, 5);
      for (let i = 0; i < 5; i++) {
        r++; const ctr = { al: { horizontal: 'center' } };
        [[bestB[i], 0], [worstB[i], 6]].forEach(([b, o]) => {
          cell(r, o + 1, o + 1, b ? i + 1 : '', ctr); cell(r, o + 2, o + 3, b ? b.cab : '', { font: { bold: true, size: 9 } }); cell(r, o + 4, o + 4, b ? b.n : '', ctr);
          cell(r, o + 5, o + 5, b ? b.fin : '', { fmt: '0.0' }); cell(r, o + 6, o + 6, b ? b.de : '', { fmt: '0%' });
        });
      }
      r += 2; secRow(r, 'SEBARAN GRADE & BAND TRAVEL LOAD');
      r++; hdr(r, [[1, 2, 'Grade'], [3, 4, 'Jumlah pengemudi'], [5, 6, '% dari aktif'], [7, 8, 'Band travel load'], [9, 9, 'Pengemudi'], [10, 10, 'Skor lama'], [11, 11, 'Skor baru'], [12, 12, 'Selisih']]);
      const gtot = k.grades.reduce((a, b) => a + b, 0) || 1;
      for (let i = 0; i < 5; i++) {
        r++; const b = k.bands[i]; const ctr = { al: { horizontal: 'center' } };
        cell(r, 1, 2, 'Grade ' + 'ABCDE'[i], ctr); cell(r, 3, 4, k.grades[i], Object.assign({ fmt: '0' }, ctr)); cell(r, 5, 6, k.grades[i] / gtot, Object.assign({ fmt: '0%' }, ctr));
        cell(r, 7, 8, b ? b.band : '', ctr); cell(r, 9, 9, b ? b.n : '', ctr); cell(r, 10, 10, b && b.old ? b.old : '', { fmt: '0.0' }); cell(r, 11, 11, b && b.fin ? b.fin : '', { fmt: '0.0' });
        cell(r, 12, 12, b && b.fin && b.old ? b.fin - b.old : '', { fmt: '+0.0;-0.0;0.0' });
      }
      r += 2; secRow(r, 'EXECUTIVE REPORT · ' + p.lbl);
      const line = (t, o) => { r++; band(r, 1, NC, t, Object.assign({ h: est(t, 150), al: { indent: 1, vertical: 'top' } }, o || {})); };
      const block = (ex, lang) => {
        line(lang === 'id' ? 'BAHASA INDONESIA' : 'ENGLISH', { font: { bold: true, size: 9, color: { argb: 'FFC2410C' } }, h: 18 });
        line(ex.headline, { font: { bold: true, size: 11 }, h: est(ex.headline, 125) });
        ex.insights.forEach(x => line('• ' + x));
        ex.risks.forEach(x => line((lang === 'id' ? '⚠ Risiko: ' : '⚠ Risk: ') + x));
        ex.actions.forEach(a => line('→ ' + a[0] + '   [' + (lang === 'id' ? 'PIC: ' : 'Owner: ') + a[1] + ' · ' + a[2] + ']'));
        line(ex.target, { font: { italic: true, size: 10, color: { argb: 'FF3A3A3A' } } });
      };
      block(exI, 'id'); block(exE, 'en');
      // status & tindak lanjut (dengan tautan ke sheet)
      r += 2; secRow(r, 'STATUS DATA & TINDAK LANJUT');
      const ntsLast = VT.buildNotifications(res, res.lastPid, Pp, master, res.tripsByKey);
      const ms = S.inputs.master || [], okM = ms.filter(m => /@/.test(m.atasanEmail || '')).length, corr = S.inputs.corrections || [], F = S.inputs.factors || {};
      const lastRows = res.rows.filter(x => x.PID === res.lastPid);
      const stat = [
        ['Notifikasi periode laporan terakhir: ' + ntsLast.filter(x => x.kirim === 'YA').length + ' siap kirim · ' + ntsLast.filter(x => x.kirim !== 'YA').length + ' tertahan karena email atasan belum ada', 'Notifikasi'],
        ['Master karyawan: ' + okM + ' lengkap · ' + (ms.length - okM) + ' belum ada email atasan langsung · ' + (new Set(lastRows.map(x => x.drv)).size - lastRows.filter(x => master.has(x.drv)).length) + ' pengemudi aktif belum ada di master', 'Master karyawan'],
        ['Faktor koreksi: ' + corr.filter(c => c.status === 'Disetujui').length + ' disetujui · ' + corr.filter(c => c.status === 'Diajukan').length + ' menunggu persetujuan QSHE', 'Faktor koreksi'],
        ['Faktor lain: ' + (F.insiden || []).length + ' insiden · ' + (F.promosi || []).length + ' promosi · ' + (F.mutasi || []).length + ' mutasi · ' + (F.keluar || []).length + ' keluar · ' + (F.pelatihan || []).length + ' pelatihan', 'Faktor lain'],
        ['Pemimpin terdeteksi (FR/CSS/Unit/Section Head): ' + VT.recapLeaders(res, res.lastPid, Pp).length + ' orang', 'Rekap pemimpin'],
        ['Kualitas data: ' + VT.fx(res.dq.generic, 0) + ' trip ID generik (DRIVER…/MASTER…, ' + res.dq.genericIds + ' ID) · ' + VT.fx(res.dq.v120, 0) + ' trip kecepatan rata-rata >120 km/jam', 'Data & cadangan'],
        ['Laporan bulanan GA: ' + (S.ga.months.length ? S.ga.months.map(m => VT.monthLabel(m)).join(', ') : 'belum diunggah'), S.ga.months.length ? 'Laporan bulanan GA' : 'Panduan']];
      stat.forEach(([t, sh]) => { r++; cell(r, 1, 10, t, { h: est(t, 110), al: { indent: 1 } }); const c = cell(r, 11, 12, { text: 'Buka sheet ›', hyperlink: "#'" + sh + "'!A1" }, { font: { size: 9, bold: true, underline: true, color: { argb: 'FFC2410C' } }, al: { horizontal: 'center' } }); });
      r += 2; band(r, 1, NC, 'Vehicle Telematic Score and Report · QSHE Department PT Hexindo Adiperkasa Tbk · model v2.0 (sama dengan workbook Excel) · ilustrasi kendaraan generik · aplikasi v' + VT.VERSION + ' · lanjutkan ke sheet "Daftar isi" untuk seluruh menu.', { font: { size: 8, italic: true, color: { argb: 'FF595959' } }, h: 28, al: { indent: 1, vertical: 'top' } });
      dash.pageSetup = { orientation: 'portrait', fitToPage: true, fitToWidth: 1, fitToHeight: 0, paperSize: 9 };
    }

    // 2. Skor 2-mingguan (semua periode)
    const C = VT.COMP;
    sheet('Detail pengemudi', 'SKOR PER PENGEMUDI PER PERIODE 2-MINGGUAN', 'Semua periode tersimpan · trip Driving, DID personal · ' + stamp, [
      ['PID', r => r.PID, '0', 6], ['Periode', r => nameP(r.PID), null, 22], ['Nama pengemudi (DID)', r => r.drv, null, 30], ['Cabang', r => r.cab, null, 20], ['Area', r => r.area, null, 12], ['Jabatan', r => r.JAB, null, 14],
      ['Band', r => r.band, null, 6], ['KM', r => r.km, N1, 9], ['Jam mengemudi', r => r.h, N2, 9], ['Trip', r => r.trips, '0', 6],
      ['Overspeed mentah', r => r.spd_raw, '0', 9], ['Overspeed dihitung', r => r.spd, '0', 9], ['Trip ngebut beruntun', r => r.sus, '0', 9], ['Fatigue', r => r.fat, '0', 8], ['Night driving (jam)', r => r.nh, N2, 9],
      ['Harsh braking', r => r.brk, '0', 8], ['Harsh accel', r => r.acc, '0', 8], ['Harsh cornering', r => r.cor, '0', 8],
      ...C.map(c => ['Skor ' + c.name, r => r['S_' + c.k], N1, 9]), ['Skor perilaku terbobot', r => r.VB, N1, 10], ['Kredit beban', r => r.LCR, N1, 8], ['Pengurang berulang', r => r.REP, N1, 9],
      ['Komponen berulang', r => r.NREP, '0', 8], ['Pengurang insiden', r => r.INS, N1, 9], ['Pengurang pemimpin', r => r.LDR, N1, 9], ['SKOR AKHIR', r => r.FIN, N1, 9], ['Grade', r => r.GR, null, 6],
      ['Skor metode lama', r => r.S_old, N1, 9], ['Selisih vs lama', r => r.DIF, '+0.0;-0.0;0.0', 9], ['Status overspeed', r => r.OSS, null, 14], ['Notifikasi (komponen)', r => r.NOTIF, '0', 9],
      ['Level (1 perhatian, 2 eskalasi)', r => r.ESK, '0', 9], ['Komponen terlemah', r => r.WK, null, 16], ['Kode fokus', r => r.FOC, null, 7],
      ['Aksi pekerja', r => VT.REK[r.FOC][1], null, 40], ['Aksi atasan', r => VT.REK[r.FOC][2], null, 40]], res.rows.slice().sort((a, b) => a.PID - b.PID || (a.drv < b.drv ? -1 : 1)));

    // 3. Tren per periode (nasional)
    const allTr = VT.trend(res, res.periods[res.periods.length - 1].pid, 'Semua cabang', res.periods.length);
    sheet('Tren periode', 'TREN NASIONAL PER PERIODE 2-MINGGUAN', stamp, [['PID', r => r.pid, '0', 6], ['Periode', r => r.label, null, 24], ['Skor metode lama', r => r.old, N1], ['Skor baru', r => r.fin, N1],
      ['Overspeed dihitung /1.000 km', r => r.osv, N2, 14], ['Harsh braking /1.000 km', r => r.brk, N2, 14], ['Harsh accel /1.000 km', r => r.acc, N2, 14], ['Pengemudi', r => r.n, '0']], allTr);

    // 4. Notifikasi periode terpilih
    sheet('Notifikasi', 'NOTIFIKASI EMAIL KE ATASAN · ' + p.lbl, 'Tabel siap untuk Power Automate (Kirim = YA)', [['No', r => r.no, '0', 5], ['Kirim', r => r.kirim, null, 14], ['Level', r => r.level, null, 11], ['Nama', r => r.nama, null, 28],
      ['Nama_Lengkap', r => r.namaLengkap, null, 28], ['Cabang', r => r.cab, null, 20], ['Periode', r => r.periode, null, 30], ['Atasan_Langsung', r => r.atasan, null, 22], ['Email_To', r => r.to, null, 28], ['Email_CC', r => r.cc, null, 28],
      ['Subjek', r => r.subject, null, 60], ['Isi_Email', r => (S.settings.aiEmail || {})[r.nama + '|' + pid] || r.body, null, 100]], nts);

    // 5. Detail trip pelanggaran periode terpilih
    const dt = []; for (const [key, arr] of res.tripsByKey) { if (+key.split('|').pop() === pid) arr.forEach(t => dt.push(t)); }
    dt.sort((a, b) => (a.d < b.d ? -1 : a.d > b.d ? 1 : a.t0 - b.t0));
    const xd = s => VT.dayToDate(Math.floor(s)); const tdt = s => (s == null ? '' : new Date(xd(s).getTime() + Math.round((s - Math.floor(s)) * 86400) * 1000));
    sheet('Detail trip', 'DETAIL TRIP PELANGGARAN · ' + p.lbl, 'Isi bagian "Kapan & di mana" pada email', [['Nama DID', t => t.d, null, 28], ['Mulai', t => tdt(t.t0), 'dd-mmm-yy hh:mm', 16], ['Selesai', t => tdt(t.t1), 'dd-mmm-yy hh:mm', 16],
      ['Cabang', t => t.cab, null, 20], ['Plat', t => t.pl, null, 12], ['KM', t => t.km, N1, 8], ['Jam', t => t.h, N2, 7], ['Overspeed', t => t.spd, '0', 8], ['Fatigue', t => t.fat, '0', 7], ['Braking', t => t.brk, '0', 7], ['Accel', t => t.acc, '0', 7],
      ['Pelanggaran', t => VT.tripViol(t, Pp).text, null, 40], ['Baris email', t => VT.tripLine(t, Pp), null, 90]], dt);

    // 6. Rekap
    const fyTxt = 'FY' + p.fy + ' s.d. ' + p.lbl;
    sheet('Rekap pengemudi', 'REKAP PENGEMUDI ' + fyTxt, 'Peringkat 1 = terbaik', [['Nama pengemudi (DID)', r => r.drv, null, 30], ['Cabang (terakhir)', r => r.cab, null, 20], ['Jabatan', r => r.jab, null, 12], ['Periode aktif', r => r.periods, '0', 8],
      ['Total KM', r => r.km, N0, 10], ['KM rata-rata / periode', r => r.kmPer, N0, 10], ['Band periode terakhir', r => r.band, null, 9], ['Skor LAMA rata-rata', r => r.old, N1, 10], ['Skor BARU rata-rata', r => r.fin, N1, 10], ['Selisih', r => r.dif, '+0.0;-0.0;0.0', 8],
      ['Skor BARU periode terakhir', r => r.lastFin, N1, 10], ['Grade periode terakhir', r => r.lastGr, null, 8], ['Tren', r => r.trend, '+0.0;-0.0;0.0', 8], ['Periode berulang', r => r.repP, '0', 8], ['Total pengurang berulang', r => r.rep, N1, 10],
      ['Overspeed dihitung', r => r.osv, '0', 9], ['Notifikasi atasan', r => r.notif, '0', 9], ['Pengurang insiden (terakhir)', r => r.ins, N1, 10], ['Peringkat BARU', r => r.rank_fin, '0', 8], ['Peringkat LAMA', r => r.rank_old, '0', 8],
      ['Fokus', r => r.foc, null, 7], ['Atasan langsung', r => r.atasan, null, 22], ['Email atasan', r => r.atasanEmail, null, 26]], VT.recapDrivers(res, pid, master, Pp).sort((a, b) => (a.rank_fin || 1e9) - (b.rank_fin || 1e9)));
    sheet('Rekap cabang', 'REKAP CABANG ' + fyTxt, '', [['Cabang', r => r.cab, null, 24], ['Area', r => r.area, null, 12], ['Pengemudi aktif', r => r.drivers, '0'], ['Total KM', r => r.km, N0, 11], ['Skor LAMA', r => r.old, N1], ['Skor BARU', r => r.fin, N1],
      ['Selisih', r => r.dif, '+0.0;-0.0;0.0'], ['% grade D/E', r => r.de, PC1], ['Driver-periode berulang', r => r.rep, '0'], ['Overspeed dihitung', r => r.osv, '0'], ['Notifikasi', r => r.notif, '0'], ['Insiden', r => r.ins, '0'], ['Peringkat BARU', r => r.rank, '0'], ['Peringkat LAMA', r => r.rankOld, '0']],
      VT.recapBranches(res, pid, S.inputs.factors.insiden).sort((a, b) => a.rank - b.rank));
    sheet('Rekap pemimpin', 'REKAP PEMIMPIN · ' + p.lbl, 'FR = Foreman · CSS = CS Supervisor · Unit/Section Head', [['Nama', r => r.drv, null, 28], ['Jabatan', r => r.jab, null, 14], ['Cabang', r => r.cab, null, 20], ['Faktor peran', r => r.rf, N1], ['Skor mengemudi sendiri', r => r.own, N1, 12],
      ['Rata-rata tim cabang', r => r.team, N1, 12], ['Pengurang pemimpin', r => r.ldr, N1, 12], ['SKOR PEMIMPIN', r => r.score, N1, 12], ['Grade', r => r.gr, null, 7], ['Dasar skor', r => r.src, null, 24]], VT.recapLeaders(res, pid, Pp));
    sheet('Uji keadilan', 'UJI KEADILAN TRAVEL LOAD · ' + fyTxt, 'Adil = rata-rata skor antar band hampir sama', [['Band', r => r.band, null, 7], ['Driver-periode', r => r.n, '0'], ['KM rata-rata', r => r.km, N0], ['Skor LAMA', r => r.old, N1], ['Skor perilaku', r => r.vb, N1],
      ['Pengurang berulang', r => r.rep, N1], ['Kredit beban', r => r.lcr, N1], ['Skor BARU akhir', r => r.fin, N1], ['% grade A', r => r.a, PC], ['% grade D/E', r => r.de, PC]], VT.fairness(res, pid));
    sheet('Analisa tren', 'TREN BULANAN (dari trip Driving, ID umum dikecualikan)', '', [['Bulan', r => VT.monthLabel(r.mo), null, 16], ['Pengemudi', r => r.drivers, '0'], ['Trip', r => r.trips, N0], ['KM', r => r.km, N0, 11], ['Jam', r => r.h, N0],
      ['Overspeed mentah /1.000 km', r => (r.km ? r.spd / r.km * 1000 : null), N2, 13], ['Overspeed dihitung /1.000 km', r => (r.km ? r.spdc / r.km * 1000 : null), N2, 13], ['Fatigue /100 jam', r => (r.h ? r.fat / r.h * 100 : null), N2, 11],
      ['Braking /1.000 km', r => (r.km ? r.brk / r.km * 1000 : null), N2, 11], ['Accel /1.000 km', r => (r.km ? r.acc / r.km * 1000 : null), N2, 11], ['Trip malam (aturan lama ≥19:00)', r => r.nOld, N0, 13], ['Trip malam (aturan baru)', r => r.nNew, N0, 13]], VT.monthly(res, 'Semua cabang'));

    // 7. Input: koreksi, faktor lain, master
    sheet('Faktor koreksi', 'FAKTOR KOREKSI (audit trail)', 'Hanya status Disetujui yang mengubah skor', [['Status', r => r.status, null, 10], ['Nama DID', r => r.nama, null, 28], ['Alasan', r => (VT.CORR_REASONS.find(x => x.id === r.alasan) || {}).label || r.alasan, null, 44],
      ['Aksi', r => VT.CORR_ACTIONS[r.aksi], null, 26], ['Komponen dianulir', r => (r.komponen || []).join(', '), null, 16], ['Pemakai sebenarnya', r => r.pemakai, null, 24], ['Periode', r => (r.pid ? nameP(+r.pid) : ''), null, 22],
      ['Dari', r => r.dari, null, 11], ['Sampai', r => r.sampai, null, 11], ['Plat', r => r.plat, null, 11], ['Trip dipilih', r => (r.trips || []).length, '0', 8], ['Trip terdampak', r => (res.corAudit[r.id] || {}).trips || 0, '0', 9],
      ['KM terdampak', r => (res.corAudit[r.id] || {}).km || 0, N0, 9], ['No. bukti', r => r.bukti, null, 18], ['Diajukan oleh', r => r.oleh, null, 16], ['Disetujui oleh', r => r.approver, null, 16], ['Keterangan', r => r.keterangan, null, 40]], S.inputs.corrections || []);
    const fl = []; VT.FACTOR_TYPES.forEach(f => (S.inputs.factors[f.id] || []).forEach(x => fl.push(Object.assign({}, x, { tipe: f.label }))));
    sheet('Faktor lain', 'FAKTOR LAIN (insiden, promosi, mutasi, cuti, keluar, pelatihan, ID bersama)', '', [['Faktor', r => r.tipe, null, 30], ['Nama DID', r => r.nama, null, 28], ['Tanggal', r => r.tanggal || r.dari, null, 11], ['Sampai', r => r.sampai, null, 11],
      ['Cabang', r => r.cabang, null, 18], ['Jabatan / jenis', r => r.jabatan || r.jenis || '', null, 22], ['Cabang lama / jabatan lama', r => r.cabangLama || r.jabLama || '', null, 18], ['Keparahan', r => r.keparahan, null, 22], ['Kesalahan', r => r.kesalahan, null, 16], ['CAPA', r => r.capa, null, 7],
      ['No. laporan', r => r.nolap, null, 18], ['Alasan', r => r.alasan, null, 20], ['Keterangan', r => r.ket, null, 40]], fl);
    sheet('Master karyawan', 'MASTER KARYAWAN, EMAIL & ATASAN', '', VT.MASTER_FIELDS.map(([key, l]) => [l, r => r[key], null, key.indexOf('mail') >= 0 ? 28 : 20]), (S.inputs.master || []).slice().sort((a, b) => (a.did < b.did ? -1 : 1)));

    // 8. Data GA bulanan (bulan terbaru)
    if (S.ga.months.length) {
      const m = S.ga.months[S.ga.months.length - 1]; const A = VT.gaAnalysis(S.ga, m, Pp, master, S.ga.log); const ml = VT.monthLabel(m);
      sheet('GA per kategori', 'LAPORAN BULANAN GA · ' + ml + ' vs ' + VT.monthLabel(A.prevMonth), '', [['Kategori', r => r.kat, null, 14], ['Bulan lalu', r => r.prv, N0], ['Bulan ini', r => r.cur, N0], ['Perubahan', r => r.chg, '+0%;-0%;0%'], ['Orang tercatat', r => r.org, '0'],
        ['Dari pengemudi berulang', r => r.rep, PC, 12], ['Berulang, tidak turun', r => r.tt, '0', 11], ['Berulang, sudah turun', r => r.tu, '0', 11], ['Baru', r => r.bar, '0']], A.cat);
      sheet('GA prioritas', 'PRIORITAS TINDAK LANJUT · ' + ml, 'P1/P2 sesuai parameter laporan bulanan', [['Prioritas', r => r.pr, null, 9], ['Nama / DID', r => r.nama, null, 28], ['Unit', r => r.unit, null, 22], ['OS lalu', r => r.os_prv, '0'], ['OS ini', r => r.os_cur, '0'],
        ['Fatigue lalu', r => r.ft_prv, '0'], ['Fatigue ini', r => r.ft_cur, '0'], ['Malam lalu', r => r.nd_prv, '0'], ['Malam ini', r => r.nd_cur, '0'], ['Alasan', r => r.al, null, 40], ['Atasan langsung', r => r.atasan, null, 24], ['TL bulan lalu', r => r.tlPrev, null, 11], ['TL bulan ini', r => r.tlCur, null, 12]], A.list);
      sheet('GA per unit', 'PER UNIT · ' + ml, '', [['Unit', r => r.unit, null, 24], ['Tipe', r => r.tipe, null, 12], ['OS lalu', r => r.os_prv, '0'], ['OS ini', r => r.os_cur, '0'], ['Fatigue lalu', r => r.ft_prv, '0'], ['Fatigue ini', r => r.ft_cur, '0'],
        ['Malam lalu', r => r.nd_prv, '0'], ['Malam ini', r => r.nd_cur, '0'], ['Total lalu', r => r.tot_prv, '0'], ['Total ini', r => r.tot_cur, '0'], ['Selisih', r => r.sel, '+0;-0;0'], ['Orang tercatat', r => r.org, '0'], ['Orang P1', r => r.p1, '0']], A.units);
      sheet('Log tindak lanjut', 'LOG TINDAK LANJUT PRIORITAS', '', [['Bulan', r => VT.monthLabel(r.bulan), null, 14], ['Nama', r => r.nama, null, 28], ['Prioritas', r => r.prioritas, null, 9], ['PIC', r => r.pic, null, 20], ['Tanggal', r => r.tanggal, null, 11], ['Tindakan', r => r.tindakan, null, 40], ['Hasil', r => r.hasil, null, 40], ['Status', r => r.status, null, 10]], S.ga.log || []);
    }

    // 9. Referensi
    const pr = [];
    [['NSTART', 'Jam mulai night driving'], ['NEND', 'Night driving sampai sebelum jam'], ['OSTOL', 'Toleransi overspeed (count/trip)'], ['OSKUN', 'Batas overspeed KUNING (/1.000 km)'], ['OSSUS', 'Count ngebut beruntun'], ['BASE', 'Basis skor komponen'], ['KKM', 'Km bayangan'], ['KH', 'Jam bayangan'],
      ['TOL', 'Toleransi tidak berubah'], ['CRIT', 'Level kritis (× rata-rata)'], ['PTET', 'Pengurang berulang tetap'], ['PBUR', 'Pengurang berulang makin buruk'], ['OSMULT', 'Pengali overspeed'], ['STREAK', 'Pengali beruntun'], ['REPCAP', 'Maks pengurang berulang'],
      ['DECAY', 'Masa berlaku insiden (hari)'], ['CAPA', 'Faktor CAPA tepat waktu'], ['LCAP', 'Maks pengurang pemimpin'], ['NOTIFR', 'Batas MERAH komponen (rasio)'], ['NOTIFMIN', 'Minimal kejadian notifikasi'], ['DEADLINE', 'Tenggat tindak lanjut (hari kerja)'], ['GAP1OS', 'GA Prioritas 1: overspeed ≥'], ['GAP2', 'GA Prioritas 2: kasus ≥']].forEach(([key, l]) => pr.push([l, Pp[key], key]));
    VT.LAB.forEach((b, i) => pr.push(['Band ' + b + ' batas bawah km / kredit beban', Pp.BAND_LO[i] + ' / ' + Pp.LC[i], 'BAND']));
    C.forEach(c => pr.push(['Bobot ' + c.name, Pp.W[c.k], 'W']));
    C.forEach(c => VT.LAB.forEach(b => pr.push(['Rata-rata kelompok ' + c.name + ' ' + b + ' (' + c.unit + ')', Pp.BM[b][c.k], 'BM'])));
    pr.push(['Grade A/B/C/D minimum', Pp.GRADE.join(' / '), 'GRADE']);
    sheet('Parameter', 'PARAMETER & STANDAR YANG DIPAKAI', Object.keys(S.inputs.params || {}).length ? 'Ada parameter yang diubah dari bawaan' : 'Semua parameter = bawaan workbook Excel', [['Parameter', r => r[0], null, 54], ['Nilai', r => r[1], null, 14], ['Kode', r => r[2], null, 10]], pr);
    sheet('Periode 2-mingguan', 'DAFTAR PERIODE 2-MINGGUAN', '', [['PID', r => r.pid, '0', 6], ['Periode (label Beweekly)', r => r.lbl, null, 26], ['Mulai', r => VT.dayToDate(r.start), 'dd-mmm-yy', 11], ['Selesai', r => VT.dayToDate(r.end), 'dd-mmm-yy', 11],
      ['Hari data', r => r.days, '0'], ['FY', r => 'FY' + r.fy, null, 8], ['Trip Driving', r => r.drive, N0], ['Trip Idling', r => r.idle, N0], ['Pengemudi dinilai', r => res.rows.filter(x => x.PID === r.pid).length, '0']], res.periods);
    sheet('Matriks tindak lanjut', 'MATRIKS TINDAK LANJUT', '', [['Kode', r => r[0], null, 7], ['Pemicu', r => VT.REK[r[0]][0], null, 30], ['Aksi pekerja', r => VT.REK[r[0]][1], null, 60], ['Aksi atasan / safety manager', r => VT.REK[r[0]][2], null, 60], ['Tenggat', r => VT.REK[r[0]][3], null, 16]], Object.keys(VT.REK).map(x => [x]));

    // ---- helper: sheet bersusun (beberapa tabel kecil dalam satu sheet)
    const multi = (name, title, sub, widths, sections) => {
      const ws = wb.addWorksheet(name.slice(0, 31), { views: [{ state: 'frozen', ySplit: 2 }] });
      titleRows(ws, title, sub, widths.length);
      widths.forEach((w, i) => { ws.getColumn(i + 1).width = w; });
      let r = 2;
      sections.forEach(sc => {
        r += 2; const t = ws.getRow(r); t.getCell(1).value = sc.t; t.getCell(1).font = { bold: true, size: 11 }; t.getCell(1).border = { left: { style: 'thick', color: { argb: OR } } }; t.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: LIGHT } };
        if (sc.cols) { r++; const h = ws.getRow(r); sc.cols.forEach((c, i) => { const x = h.getCell(i + 1); x.value = c; x.font = { bold: true, size: 9, color: { argb: 'FFFFFFFF' } }; x.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1A1A1A' } }; x.alignment = { wrapText: true, vertical: 'middle' }; x.border = { bottom: { style: 'medium', color: { argb: OR } } }; }); h.height = 26; }
        sc.rows.forEach(row => {
          r++; const rr = ws.getRow(r);
          row.forEach((v, i) => { const x = rr.getCell(i + 1); if (typeof v === 'number' && !isFinite(v)) v = null; x.value = v == null ? '' : v; x.alignment = { wrapText: true, vertical: 'top' }; x.font = { size: 10 }; if (sc.fmts && sc.fmts[i] && typeof v === 'number') x.numFmt = sc.fmts[i]; });
          if (sc.bold) rr.getCell(1).font = { size: 10, bold: true };
          const longest = Math.max(...row.map((v, i) => Math.ceil(String(v == null ? '' : v).length / Math.max(8, (widths[i] || 12) * 1.05))));
          if (longest > 1) rr.height = Math.min(220, 14.5 * longest + 2);
        });
      });
      return ws;
    };
    const nt0 = String(Pp.NSTART).padStart(2, '0') + ':00–' + String((Pp.NEND + 23) % 24).padStart(2, '0') + ':59';

    // PANDUAN
    multi('Panduan', 'PANDUAN PEMAKAIAN', 'Untuk Tim VT, atasan, dan QSHE. Semua angka & aturan sama dengan workbook Excel "Vehicle Telematic Score and Report".', [34, 120], [
      { t: 'Ringkasnya dalam 1 menit', bold: true, rows: [
        ['Sumber data', 'File mentah Cartrack ("Data Score Behavior", .xlsx atau .csv). Sheet "Data" dipakai; hanya trip Driving dengan DID personal yang dinilai.'],
        ['Travel load berimbang', 'Pengemudi dibandingkan dengan rata-rata pengemudi di band km yang sama (L1–L5), jadi yang jaraknya jauh tidak otomatis kalah.'],
        ['Overspeed', 'Overspeed ≤ ' + Pp.OSTOL + ' count per trip ditoleransi. ≥ ' + Pp.OSSUS + ' count dalam satu trip = ngebut beruntun → status MERAH.'],
        ['Night driving', 'Trip yang mulai ' + nt0 + '.'],
        ['Pelanggaran berulang', 'Pelanggaran yang sama 2 periode berturut-turut → pengurang skor. Bila di luar standar (MERAH), atasan menerima email berisi kapan, di mana, dan apa pelanggarannya.'],
        ['Penyimpanan', 'Data diolah dan disimpan di perangkat (browser) saja. Gunakan menu Data & cadangan untuk memindahkan data.']] },
      { t: 'Langkah kerja tiap 2 minggu', bold: true, rows: [
        ['1. Unggah data mentah', 'Menu Unggah data. Boleh file FY penuh atau 2-mingguan; trip ganda otomatis dilewati.'],
        ['2. Faktor koreksi', 'DID dipinjam, error GPS, darurat, perjalanan malam ber-JRM, dll. Ajukan, lalu QSHE menyetujui — hanya yang Disetujui mengubah skor.'],
        ['3. Faktor lain', 'Insiden/kecelakaan, promosi, mutasi, cuti, resign, pelatihan, ID bersama.'],
        ['4. Master karyawan', 'Email pengemudi, atasan langsung, atasan tidak langsung. Tanpa email atasan, notifikasi tertahan.'],
        ['5. Dashboard', 'Baca executive report, KPI, dan 10 prioritas coaching.'],
        ['6. Notifikasi email atasan', 'Salin isi email / buka di aplikasi email, atau ekspor tabel untuk Power Automate.'],
        ['7. Laporan bulanan GA', 'Unggah lampiran "Data Case" GA, isi angka infografis, catat tindak lanjut Prioritas 1–2.'],
        ['8. Laporan bilingual', 'Unduh Word/PDF (ID/EN) untuk manajemen.']] },
      { t: 'Angka penting (bisa diubah QSHE di menu Parameter)', bold: true, rows: [
        ['Toleransi overspeed', '≤ ' + Pp.OSTOL + ' count per trip tidak dihitung'], ['Batas overspeed KUNING', '≤ ' + VT.fx(Pp.OSKUN, 0) + ' per 1.000 km; lebih dari itu MERAH'], ['Jam malam', nt0],
        ['Grade', 'A ≥ ' + Pp.GRADE[0] + ' · B ≥ ' + Pp.GRADE[1] + ' · C ≥ ' + Pp.GRADE[2] + ' · D ≥ ' + Pp.GRADE[3] + ' · E di bawahnya'],
        ['Catatan sheet ini', 'Workbook ini dibuat oleh tombol "Unduh semua (Excel)". Sheet pertama = Dashboard; sheet "Daftar isi" memuat tautan ke semua menu.']] }]);

    // UNGGAH DATA — riwayat data yang sudah diunggah
    sheet('Unggah data', 'UNGGAH DATA · RIWAYAT DATA YANG TERSIMPAN', 'Setiap baris = satu periode 2-mingguan hasil unggahan file mentah Cartrack', [['PID', r => r.pid, '0', 6], ['Periode (label Beweekly)', r => r.lbl, null, 26], ['Mulai', r => VT.dayToDate(r.start), 'dd-mmm-yy', 11], ['Selesai', r => VT.dayToDate(r.end), 'dd-mmm-yy', 11],
      ['Baris data', r => r.rows, N0, 10], ['Trip Driving', r => r.drive, N0, 10], ['Trip Idling (tidak dinilai)', r => r.idle, N0, 12], ['% Idling', r => (r.rows ? r.idle / r.rows : null), PC, 9], ['Hari data', r => r.days, '0', 9], ['Catatan', r => (r.days < 7 ? 'data <7 hari' : ''), null, 16]], res.periods);

    // LAPORAN BILINGUAL (ID/EN) — isi laporan yang sama dengan Word/PDF
    {
      const R = VTR.buildReport(res, pid, S);
      const ws = wb.addWorksheet('Laporan bilingual', { views: [{ showGridLines: false }] });
      for (let i = 1; i <= 8; i++) ws.getColumn(i).width = 22;
      const sol = a => ({ type: 'pattern', pattern: 'solid', fgColor: { argb: a } });
      const bi = c => (c && typeof c === 'object' ? (c.id === c.en ? String(c.id) : c.id + '\n' + c.en) : String(c == null ? '' : c));
      const lines = (t, w) => String(t).split('\n').reduce((a, s) => a + Math.max(1, Math.ceil(s.length / w)), 0);
      let r = 0; const row = () => ++r;
      const mg = (c1, c2, v, o) => { const rr = r; if (c2 > c1) ws.mergeCells(rr, c1, rr, c2); const c = ws.getCell(rr, c1); c.value = v; c.alignment = { wrapText: true, vertical: 'top' }; c.font = Object.assign({ size: 10 }, (o && o.font) || {}); if (o && o.fill) { for (let i = c1; i <= c2; i++) ws.getCell(rr, i).fill = sol(o.fill); } return c; };
      row(); mg(1, 8, 'QSHE DEPARTMENT · PT HEXINDO ADIPERKASA TBK', { font: { size: 9, bold: true, color: { argb: 'FF595959' } } });
      row(); mg(1, 8, R.title.id, { font: { size: 18, bold: true, color: { argb: 'FFFFFFFF' } }, fill: BLACK }); ws.getRow(r).height = 32;
      row(); mg(1, 8, R.title.en, { font: { size: 13, color: { argb: OR } }, fill: BLACK }); ws.getRow(r).height = 22;
      row(); mg(1, 8, R.sub.id + '   |   ' + R.sub.en, { font: { size: 9, italic: true }, fill: LIGHT });
      R.sec.forEach((s, i) => {
        row(); row(); mg(1, 8, (i + 1) + '. ' + s.h.id + '  /  ' + s.h.en, { font: { bold: true, size: 12 }, fill: LIGHT }).border = { left: { style: 'thick', color: { argb: OR } } }; ws.getRow(r).height = 22;
        const pair = (a, b, pre) => { row(); mg(1, 4, (pre || '') + a); mg(5, 8, (pre || '') + b, { font: { italic: true, color: { argb: 'FF555555' } } }); ws.getRow(r).height = Math.min(400, 14.5 * Math.max(lines(a, 88), lines(b, 88)) + 4); };
        (s.p || []).forEach(x => pair(x.id, x.en)); (s.list || []).forEach(x => pair(x.id, x.en, '• '));
        if (s.table) {
          const n = Math.min(8, s.table.head.length);
          row(); s.table.head.slice(0, n).forEach((h, j) => { const c = ws.getCell(r, j + 1); c.value = bi(h); c.font = { bold: true, size: 9, color: { argb: 'FFFFFFFF' } }; c.fill = sol('FF1A1A1A'); c.alignment = { wrapText: true, vertical: 'middle' }; c.border = { bottom: { style: 'medium', color: { argb: OR } } }; }); ws.getRow(r).height = 30;
          s.table.rows.forEach(rw => { row(); let ml = 1; rw.slice(0, n).forEach((v, j) => { const t = bi(v); const c = ws.getCell(r, j + 1); c.value = t; c.alignment = { wrapText: true, vertical: 'top' }; c.font = { size: 9 }; c.border = { bottom: { style: 'thin', color: { argb: 'FFE5E5E5' } } }; ml = Math.max(ml, lines(t, 23)); }); ws.getRow(r).height = Math.min(240, 13 * ml + 3); });
        }
      });
      row(); row(); mg(1, 8, 'Disusun otomatis oleh aplikasi Vehicle Telematic Score and Report v' + VT.VERSION + ' pada ' + new Date().toLocaleString('id-ID') + '.', { font: { size: 8, italic: true, color: { argb: 'FF595959' } } });
    }

    // LAPORAN BULANAN GA — ringkasan KPI (bulan terbaru)
    if (S.ga.months.length) {
      const m = S.ga.months[S.ga.months.length - 1]; const A = VT.gaAnalysis(S.ga, m, Pp, master, S.ga.log); const info = S.ga.info[m] || {};
      const kend = +info.kend || 0, kct = +info.kct || 0, peng = +info.peng || 0;
      multi('Laporan bulanan GA', 'LAPORAN BULANAN & CONCERN (DATA GA) · ' + VT.monthLabel(m) + ' vs ' + VT.monthLabel(A.prevMonth), 'Semua angka otomatis dari lampiran "Data Case" GA + angka infografis yang diisi', [46, 18, 80], [
        { t: 'Ringkasan', cols: ['Indikator', 'Nilai', 'Catatan'], fmts: [null, PC, null], rows: [
          ['Pengguna tercatat ≥1 pelanggaran', peng ? A.orgCur / peng : 'n/a', A.orgCur + ' dari ' + (peng || '? (isi infografis)') + ' pengguna'],
          ['Kasus dari pengemudi berulang', A.repAll, 'tercatat di kategori yang sama bulan lalu'],
          ['Prioritas 1 (dipanggil minggu ini)', A.p1, 'Prioritas 2: ' + A.p2 + ' orang'],
          ['Membaik (bersih bulan ini)', A.mb, 'tercatat bulan lalu, kini nol kasus'],
          ['Overspeed dari 1 orang', A.cat[0].cur ? (A.topOs ? A.topOs.os_cur : 0) / A.cat[0].cur : 0, A.topOs ? A.topOs.nama + ' · ' + A.topOs.os_cur + ' kasus' : ''],
          ['Night driving: pengguna tercatat', peng ? A.orgNd / peng : 'n/a', VT.fx(A.cat[2].cur ? A.ndProj / A.cat[2].cur * 100 : 0, 0) + '% kasus dari project / dept'],
          ['Kendaraan tanpa Cartrack', kend ? kend - kct : 'n/a', kend ? VT.fx((kend - kct) / kend * 100, 0) + '% dari ' + kend + ' unit' : 'isi angka infografis'],
          ['Kasus atas ID umum MASTER/DRIVER', A.gen, 'termasuk ' + A.genOs + ' overspeed · tak bisa ditelusuri']] }]);
    }

    // PERBANDINGAN — periode terpilih vs periode sebelumnya (nasional) + bulan
    {
      const idx = res.periods.findIndex(x => x.pid === pid); const pa = idx > 0 ? res.periods[idx - 1] : null;
      const secs = [];
      if (pa) {
        const ka = VT.kpis(res, pa.pid, 'Semua cabang'), kb = k;
        const K = [['Pengemudi aktif', 'n', N0], ['Total km', 'km', N0], ['Skor rata-rata (baru)', 'fin', N1], ['Skor rata-rata (lama)', 'old', N1], ['% grade A–B', 'ab', PC], ['Notifikasi atasan', 'notif', '0'], ['Pelanggaran berulang', 'rep', '0'], ['Overspeed MERAH', 'merah', '0'],
          ['Overspeed dihitung', 'osv', N0], ['Night driving (jam)', 'nh', N0], ['Night % jam', 'nhPct', N1], ['Harsh braking /1.000 km', 'brk', N2]];
        secs.push({ t: 'KPI nasional · A (pembanding) = ' + pa.lbl + ' · B (dianalisa) = ' + p.lbl, cols: ['Indikator', 'A · ' + pa.lbl, 'B · ' + p.lbl, 'Selisih', '% perubahan'], rows: K.map(([l, f]) => [l, ka[f], kb[f], ka[f] != null && kb[f] != null ? kb[f] - ka[f] : '', ka[f] ? (kb[f] - ka[f]) / Math.abs(ka[f]) : '']), fmts: [null, N2, N2, '+0.00;-0.00;0.00', '+0%;-0%;0%'], bold: true });
        const ba = new Map(VT.branches(res, pa.pid, Pp).map(b => [b.cab, b])), bb = VT.branches(res, pid, Pp);
        secs.push({ t: 'Per cabang', cols: ['Cabang', 'Skor A', 'Skor B', 'Selisih', 'Notifikasi A', 'Notifikasi B'], rows: bb.map(b => { const a = ba.get(b.cab); return [b.cab, a ? a.fin : '', b.fin, a ? b.fin - a.fin : '', a ? a.notif : '', b.notif]; }).sort((x, y) => (x[3] === '' ? 1 : y[3] === '' ? -1 : y[3] - x[3])), fmts: [null, N1, N1, '+0.0;-0.0;0.0', '0', '0'] });
        const da = new Map(), db = new Map(); res.rows.filter(x => x.PID === pa.pid).forEach(x => da.set(x.drv, x)); res.rows.filter(x => x.PID === pid).forEach(x => db.set(x.drv, x));
        const mv = [...db].filter(([d]) => da.has(d)).map(([d, x]) => [d, x.cab, da.get(d).FIN, x.FIN, x.FIN - da.get(d).FIN]);
        const mc = ['Pengemudi', 'Cabang', 'Skor A', 'Skor B', 'Selisih'], mf = [null, null, N1, N1, '+0.0;-0.0;0.0'];
        secs.push({ t: 'Pengemudi paling membaik (10)', cols: mc, rows: mv.slice().sort((x, y) => y[4] - x[4]).slice(0, 10), fmts: mf });
        secs.push({ t: 'Pengemudi paling memburuk (10)', cols: mc, rows: mv.slice().sort((x, y) => x[4] - y[4]).slice(0, 10), fmts: mf });
      } else secs.push({ t: 'Belum ada periode pembanding', rows: [['Hanya ada satu periode tersimpan.']] });
      const mo = VT.monthly(res, 'Semua cabang');
      secs.push({ t: 'Bulanan (semua bulan tersimpan)', cols: ['Bulan', 'Pengemudi', 'Trip', 'KM', 'Overspeed /1.000 km', 'Braking /1.000 km', 'Fatigue /100 jam'], rows: mo.map(x => [VT.monthLabel(x.mo), x.drivers, x.trips, x.km, x.km ? x.spdc / x.km * 1000 : '', x.km ? x.brk / x.km * 1000 : '', x.h ? x.fat / x.h * 100 : '']), fmts: [null, '0', N0, N0, N2, N2, N2] });
      multi('Perbandingan', 'PERBANDINGAN PERIODE / BULAN', 'Default: periode terpilih (B) dibanding periode sebelumnya (A). Untuk pasangan lain gunakan menu Perbandingan di aplikasi lalu tombol unduh di sana.', [34, 18, 18, 16, 16, 16, 16], secs);
    }

    // DATA & CADANGAN — ringkasan penyimpanan + kualitas data
    {
      const ps = res.periods, idle = VT.sum(ps, x => x.idle), rws = VT.sum(ps, x => x.rows), dq = res.dq;
      const notInMaster = [...new Set(res.rows.filter(x => x.PID === res.lastPid).map(x => x.drv))].filter(d => !master.has(d)).length;
      multi('Data & cadangan', 'DATA & CADANGAN · RINGKASAN PENYIMPANAN DAN KUALITAS DATA', 'Cadangan lengkap (JSON) tetap diunduh dari menu Data & cadangan di aplikasi.', [58, 22, 60], [
        { t: 'Ringkasan penyimpanan', cols: ['Item', 'Jumlah', 'Catatan'], fmts: [null, N0, null], rows: [['Trip Driving tersimpan', S.trips.length, ''], ['Periode 2-mingguan', ps.length, ps.length ? ps[0].lbl + ' s.d. ' + ps[ps.length - 1].lbl : ''], ['Master karyawan', (S.inputs.master || []).length, ''],
          ['Faktor koreksi', (S.inputs.corrections || []).length, ''], ['Bulan data GA', S.ga.months.length, S.ga.months.map(m => VT.monthLabel(m)).join(', ')], ['Entri log tindak lanjut', (S.ga.log || []).length, '']] },
        { t: 'Kualitas data', cols: ['Pemeriksaan', 'Jumlah', 'Keterangan'], fmts: [null, N0, null], rows: [
          ['Trip Idling (tidak dipakai untuk skor)', idle, VT.fx(rws ? idle / rws * 100 : 0, 0) + '% dari ' + VT.fx(rws, 0) + ' baris'],
          ['Trip Driving berID umum (DRIVER…/MASTER…)', dq.generic, dq.genericIds + ' ID · ' + VT.fx(dq.genericKm, 0) + ' km — tidak masuk skor personal'],
          ['Trip kecepatan rata-rata >120 km/jam', dq.v120, dq.v120os + ' ber-overspeed — cek di Faktor koreksi → Saran'],
          ['ID ditandai bukan pengemudi personal', dq.nondriver, 'trip'],
          ['Periode dengan data <7 hari', ps.filter(x => x.days < 7).length, ps.filter(x => x.days < 7).map(x => x.lbl).join(', ') || 'tidak ada'],
          ['Pengemudi aktif periode terakhir belum ada di master', notInMaster, '']] }]);
    }

    // ---- urutan sheet = urutan menu di sidebar aplikasi; sheet pertama = Dashboard
    const ORDER = ['Dashboard', 'Daftar isi', 'Panduan', 'Unggah data', 'Notifikasi', 'Rekap pengemudi', 'Rekap cabang', 'Rekap pemimpin', 'Laporan bilingual', 'Laporan bulanan GA', 'GA per kategori', 'GA prioritas', 'GA per unit', 'Log tindak lanjut',
      'Faktor koreksi', 'Faktor lain', 'Master karyawan', 'Perbandingan', 'Uji keadilan', 'Analisa tren', 'Tren periode', 'Detail pengemudi', 'Detail trip', 'Parameter', 'Matriks tindak lanjut', 'Periode 2-mingguan', 'Data & cadangan'];
    const MENU = { 'Dashboard': ['Dashboard & executive report', 'Halaman depan: KPI, executive report ID/EN, 4 grafik, 10 prioritas coaching, cabang terbaik/perlu perhatian, status data'], 'Panduan': ['Panduan', 'Cara pakai & aturan skoring'], 'Unggah data': ['Unggah data', 'Riwayat periode/data yang sudah diunggah'],
      'Notifikasi': ['Notifikasi email atasan', 'Tabel siap Power Automate periode terpilih'], 'Rekap pengemudi': ['Rekap pengemudi', 'Rekap FY per pengemudi'], 'Rekap cabang': ['Rekap cabang', 'Rekap FY per cabang'], 'Rekap pemimpin': ['Rekap pemimpin', 'Skor FR/CSS/Unit/Section Head'],
      'Laporan bilingual': ['Laporan bilingual (ID/EN)', 'Isi laporan profesional yang sama dengan Word/PDF'], 'Laporan bulanan GA': ['Laporan bulanan & concern', 'Ringkasan KPI bulanan data GA'], 'GA per kategori': ['Laporan bulanan & concern', 'Per kategori, bulan lalu vs bulan ini'],
      'GA prioritas': ['Laporan bulanan & concern', 'Daftar prioritas tindak lanjut P1/P2'], 'GA per unit': ['Laporan bulanan & concern', 'Per unit (cabang/project/dept)'], 'Log tindak lanjut': ['Log tindak lanjut', 'PIC, tindakan, hasil, status'],
      'Faktor koreksi': ['Faktor koreksi', 'Audit trail koreksi (DID dipinjam, dll.)'], 'Faktor lain': ['Faktor lain (insiden, promosi…)', 'Insiden, promosi, mutasi, cuti, keluar, pelatihan, ID bersama'], 'Master karyawan': ['Master karyawan & email', 'Email pengemudi & atasan'],
      'Perbandingan': ['Perbandingan periode / bulan', 'KPI, cabang, pengemudi membaik/memburuk, bulanan'], 'Uji keadilan': ['Uji keadilan travel load', 'Skor per band km'], 'Analisa tren': ['Analisa tren', 'Tren bulanan dari trip Driving'], 'Tren periode': ['Analisa tren', 'Tren nasional per periode 2-mingguan'],
      'Detail pengemudi': ['Detail pengemudi', 'Skor & seluruh komponen per pengemudi per periode (semua periode)'], 'Detail trip': ['Notifikasi / Detail pengemudi', 'Trip pelanggaran periode terpilih'], 'Parameter': ['Parameter & standar', 'Semua parameter yang dipakai'],
      'Matriks tindak lanjut': ['Matriks tindak lanjut', 'Aksi pekerja & atasan per kode fokus'], 'Periode 2-mingguan': ['Periode 2-mingguan', 'Daftar periode'], 'Data & cadangan': ['Data & cadangan', 'Ringkasan penyimpanan & kualitas data'] };
    const present = new Set(wb.worksheets.map(w => w.name));
    {
      const ws = wb.addWorksheet('Daftar isi', { properties: { tabColor: { argb: BLACK } } });
      const have = ORDER.filter(n => n === 'Daftar isi' || present.has(n));
      titleRows(ws, 'DAFTAR ISI · SELURUH MENU APLIKASI DALAM SATU FILE', 'Urutan sheet sama dengan urutan menu di sidebar. Klik nama sheet untuk membuka. Menu AI (Claude) dan Instal aplikasi adalah pengaturan perangkat — tidak ada data untuk diekspor.', 4);
      const hr = ws.addRow(['No', 'Sheet', 'Menu di aplikasi', 'Isi']); hr.eachCell(c => { c.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 9 }; c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1A1A1A' } }; c.border = { bottom: { style: 'medium', color: { argb: OR } } }; });
      have.forEach((n, i) => { const m = MENU[n] || ['Daftar isi', 'Halaman ini']; const row = ws.addRow([i + 1, n === 'Daftar isi' ? n : { text: n, hyperlink: "#'" + n + "'!A1" }, m[0], m[1]]); if (n !== 'Daftar isi') row.getCell(2).font = { underline: true, color: { argb: 'FFC2410C' }, bold: true }; });
      [6, 28, 36, 90].forEach((w, i) => { ws.getColumn(i + 1).width = w; });
      present.add('Daftar isi');
    }
    wb.worksheets.forEach(w => { const i = ORDER.indexOf(w.name); w.orderNo = i >= 0 ? i : 100 + w.orderNo; });
    wb.views = [{ activeTab: 0, firstSheet: 0, x: 0, y: 0, width: 20000, height: 20000 }];

    const buf = await wb.xlsx.writeBuffer();
    download(new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), 'VT_Score_and_Report_' + p.lbl.replace(/\W+/g, '_') + '_semua.xlsx');
    toast(wb.worksheets.length + ' sheet diunduh dalam satu file Excel.');
  } catch (e) { console.error(e); toast('Gagal menyusun Excel: ' + e.message, 8000); }
  finally { busy(false); }
}
