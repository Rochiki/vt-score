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

    // 1. Ringkasan + executive report
    const ex = VTR.exec(res, pid, 'Semua cabang', 'id', S), exE = VTR.exec(res, pid, 'Semua cabang', 'en', S);
    const R1 = [
        ['Skor rata-rata (metode baru)', k.fin, 'metode lama ' + VT.fx(k.old, 1), N1], ['Perubahan vs periode lalu', k.finPrev != null ? k.fin - k.finPrev : '', 'poin', '+0.0;-0.0;0.0'],
        ['Pengemudi aktif', k.n, VT.fx(k.km, 0) + ' km · rata-rata ' + VT.fx(k.kmPer, 0) + ' km/orang', '0'], ['Grade A–B', k.ab, k.de + ' orang grade D/E', PC],
        ['Notifikasi ke atasan', k.notif, k.esk + ' eskalasi', '0'], ['Pelanggaran berulang (2 minggu)', k.rep, 'pengemudi', '0'], ['Overspeed status MERAH', k.merah, 'KUNING ' + k.kuning + ' · count dihitung ' + k.osv, '0'],
        ['Night driving (jam)', k.nh, VT.fx(k.nhPct, 1) + '% jam mengemudi', N0], ['Harsh braking / 1.000 km', k.brk, 'periode lalu ' + (k.brkPrev == null ? '–' : VT.fx(k.brkPrev, 2)), N2],
        ...'ABCDE'.split('').map((g, i) => ['Jumlah pengemudi grade ' + g, k.grades[i], '', '0']), ...k.bands.map(b => ['Skor baru band ' + b.band, b.fin, 'lama ' + VT.fx(b.old, 1) + ' · ' + b.n + ' pengemudi', N1])];
    const ws1 = sheet('Ringkasan', 'VEHICLE TELEMATIC SCORE AND REPORT · RINGKASAN NASIONAL', stamp,
      [['Indikator', r => r[0], null, 40], ['Nilai', r => r[1], null, 14], ['Satuan / catatan', r => r[2], null, 60]], R1);
    R1.forEach((r, i) => { ws1.getCell(4 + i, 2).numFmt = r[3]; });
    const exRows = [['EXECUTIVE REPORT (ID)', ''], [ex.headline, exE.headline]];
    ex.insights.forEach((x, i) => exRows.push(['Kondisi: ' + x, 'Situation: ' + exE.insights[i]]));
    ex.risks.forEach((x, i) => exRows.push(['Risiko: ' + x, 'Risk: ' + exE.risks[i]]));
    ex.actions.forEach((a, i) => exRows.push(['Aksi: ' + a[0] + ' (PIC: ' + a[1] + ' · ' + a[2] + ')', 'Action: ' + exE.actions[i][0] + ' (Owner: ' + exE.actions[i][1] + ' · ' + exE.actions[i][2] + ')']));
    exRows.push([ex.target, exE.target]);
    const ws2 = sheet('Executive report', 'EXECUTIVE REPORT · ' + p.lbl, 'Bahasa Indonesia | English', [['Bahasa Indonesia', r => r[0], null, 90], ['English', r => r[1], null, 90]], exRows.slice(1));
    ws2.eachRow((row, i) => { if (i > 3) row.alignment = { wrapText: true, vertical: 'top' }; });

    // 2. Skor 2-mingguan (semua periode)
    const C = VT.COMP;
    sheet('Skor 2-mingguan', 'SKOR PER PENGEMUDI PER PERIODE 2-MINGGUAN', 'Semua periode tersimpan · trip Driving, DID personal · ' + stamp, [
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
    sheet('Tren bulanan', 'TREN BULANAN (dari trip Driving, ID umum dikecualikan)', '', [['Bulan', r => VT.monthLabel(r.mo), null, 16], ['Pengemudi', r => r.drivers, '0'], ['Trip', r => r.trips, N0], ['KM', r => r.km, N0, 11], ['Jam', r => r.h, N0],
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
    sheet('Periode', 'DAFTAR PERIODE 2-MINGGUAN', '', [['PID', r => r.pid, '0', 6], ['Periode (label Beweekly)', r => r.lbl, null, 26], ['Mulai', r => VT.dayToDate(r.start), 'dd-mmm-yy', 11], ['Selesai', r => VT.dayToDate(r.end), 'dd-mmm-yy', 11],
      ['Hari data', r => r.days, '0'], ['FY', r => 'FY' + r.fy, null, 8], ['Trip Driving', r => r.drive, N0], ['Trip Idling', r => r.idle, N0], ['Pengemudi dinilai', r => res.rows.filter(x => x.PID === r.pid).length, '0']], res.periods);
    sheet('Matriks tindak lanjut', 'MATRIKS TINDAK LANJUT', '', [['Kode', r => r[0], null, 7], ['Pemicu', r => VT.REK[r[0]][0], null, 30], ['Aksi pekerja', r => VT.REK[r[0]][1], null, 60], ['Aksi atasan / safety manager', r => VT.REK[r[0]][2], null, 60], ['Tenggat', r => VT.REK[r[0]][3], null, 16]], Object.keys(VT.REK).map(x => [x]));

    const buf = await wb.xlsx.writeBuffer();
    download(new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), 'VT_Score_and_Report_' + p.lbl.replace(/\W+/g, '_') + '_semua.xlsx');
    toast(wb.worksheets.length + ' sheet diunduh dalam satu file Excel.');
  } catch (e) { console.error(e); toast('Gagal menyusun Excel: ' + e.message, 8000); }
  finally { busy(false); }
}
