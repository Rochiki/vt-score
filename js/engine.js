/* =====================================================================
   engine.js — mesin Vehicle Telematic Score and Report.
   Port 1:1 dari rumus workbook Excel "Vehicle_Telematic_Score_and_Report_Hexindo.xlsx"
   (sheet 02 PARAMETER, 05 SKOR_2MINGGUAN, 11 NOTIFIKASI, 12–16 rekap & analisa,
   18–21 laporan bulanan GA) + faktor koreksi & faktor lain (fitur aplikasi).
   Fungsi murni: dipakai di Web Worker, di halaman, dan di Node untuk pengujian.
   ===================================================================== */
(function (root) {
  'use strict';
  const X = root.XLSXStream || (typeof require !== 'undefined' ? require('./xlsx-stream.js') : null);
  const VT = {};
  VT.VERSION = '1.2.0';

  /* ------------------------------------------------------------ parameter */
  VT.LAB = ['L1', 'L2', 'L3', 'L4', 'L5'];
  VT.COMP = [
    { k: 'spd', name: 'Overspeed', en: 'Overspeed', code: 'OSV', unit: 'per 1.000 km', cnt: 'spd', ex: 'km', m: 1000 },
    { k: 'fat', name: 'Fatigue', en: 'Fatigue', code: 'FAT', unit: 'per 100 jam', cnt: 'fat', ex: 'h', m: 100 },
    { k: 'night', name: 'Night driving', en: 'Night driving', code: 'NGT', unit: 'jam malam per 100 jam', cnt: 'nh', ex: 'h', m: 100 },
    { k: 'brk', name: 'Harsh braking', en: 'Harsh braking', code: 'BRK', unit: 'per 1.000 km', cnt: 'brk', ex: 'km', m: 1000 },
    { k: 'acc', name: 'Harsh acceleration', en: 'Harsh acceleration', code: 'ACC', unit: 'per 1.000 km', cnt: 'acc', ex: 'km', m: 1000 },
    { k: 'cor', name: 'Harsh cornering', en: 'Harsh cornering', code: 'COR', unit: 'per 1.000 km', cnt: 'cor', ex: 'km', m: 1000 }];
  VT.REPC = ['spd', 'fat', 'night', 'brk', 'acc'];

  VT.DEFAULTS = {
    NSTART: 20, NEND: 5, OSTOL: 2, OSKUN: 10, OSSUS: 10,
    BAND_LO: [0, 100, 250, 450, 750], LC: [0, 0, 1, 2, 3],
    BM: {
      L1: { spd: 2.02, fat: 2.13, night: 3.1, brk: 2.46, acc: 2.51, cor: 97.24 },
      L2: { spd: 2.02, fat: 2.13, night: 3.1, brk: 3.16, acc: 2.73, cor: 139.64 },
      L3: { spd: 2.02, fat: 2.47, night: 3.1, brk: 3.65, acc: 3.38, cor: 163.39 },
      L4: { spd: 2.02, fat: 4.15, night: 3.1, brk: 4.8, acc: 3.79, cor: 200.68 },
      L5: { spd: 2.02, fat: 5.7, night: 3.1, brk: 5.67, acc: 3.53, cor: 212.94 }
    },
    W: { spd: 0.25, fat: 0.20, night: 0.10, brk: 0.20, acc: 0.15, cor: 0.10 },
    BASE: 0.85, KKM: 300, KH: 20,
    TOL: 0.2, CRIT: 3, PTET: 3, PBUR: 5, OSMULT: 2, STREAK: 1.5, REPCAP: 20,
    SEV: [['Near miss', 0, 0], ['Kerusakan properti ringan', 5, 3], ['Cedera ringan (FAC/MTC)', 10, 5], ['LTI / kerusakan berat', 15, 8], ['Fatal / berat', 25, 12]],
    FAULT: [['Penyebab utama', 1], ['Berkontribusi', 0.5], ['Tidak bersalah', 0]],
    DECAY: 180,
    ROLE: [['Foreman', 1, 'FR'], ['CS Supervisor', 0.8, 'CSS'], ['Unit Head', 0.8, ''], ['Section Head', 0.6, '']],
    CAPA: 0.5, LCAP: 20,
    GRADE: [90, 80, 70, 60],
    NOTIFR: 2.0, NOTIFMIN: 2,
    CCDRV: 'Tidak', CCQSHE: '', SENDER: 'QSHE Department – PT Hexindo Adiperkasa Tbk', DEADLINE: 3,
    MIN_BRANCH: 3,
    // laporan bulanan GA
    GAP1OS: 10, GAP2: 5
  };
  VT.REK = {
    INS: ['Ada pengurang insiden aktif', 'Ikuti hasil investigasi. Jalani refreshment defensive driving & briefing ulang rute sebelum kembali membawa unit.', 'Pastikan CAPA investigasi ditutup; review kelayakan izin mengemudi (KIMHAP); dampingi 2 periode.', 'Sebelum kembali mengemudi'],
    LDR: ['Pemimpin: ada insiden kendaraan di cabang', 'Pimpin toolbox talk pembelajaran insiden di cabang dalam 3 hari; pastikan tim paham penyebab & pencegahannya.', 'Pastikan investigasi & CAPA ditutup tepat waktu (memotong 50% pengurang); verifikasi CAPA di lapangan.', 'Sesuai tenggat CAPA'],
    OSR: ['Overspeed berulang 2 periode', 'Patuhi batas kecepatan rute – tidak ada toleransi. Berangkat lebih awal supaya tidak kejar waktu.', 'Coaching 1-on-1 + teguran tertulis sesuai SK; cek apakah target waktu kerja memaksa ngebut; pertimbangkan pembatas kecepatan.', '≤ 3 hari kerja'],
    OSV: ['Ada overspeed dihitung (di luar toleransi)', 'Kurangi kecepatan di ruas yang tercatat; minta rute & batas kecepatan terbaru ke atasan.', 'Cek status HIJAU/KUNING/MERAH; bila MERAH lakukan coaching; anulir hanya lewat menu Faktor koreksi dengan bukti.', '≤ 1 minggu'],
    REP: ['Pelanggaran berulang (selain overspeed)', 'Perhatikan kebiasaan yang muncul 2 periode berturut. Target periode depan: turun ≥20%.', 'Diskusikan penyebab (rute, beban kerja, kondisi unit); jadwalkan ride-along/review data VT; pantau periode berikutnya.', '≤ 1 minggu'],
    FAT: ['Fatigue', 'Istirahat 15 menit tiap 2 jam atau sebelum 4 jam nyetir. Jangan nyetir sendirian perjalanan jauh setelah lembur.', 'Atur ulang penugasan perjalanan jauh; terapkan journey management (titik istirahat, 2 pengemudi untuk rute jauh); cek jam kerja & lembur 2 minggu terakhir.', 'Penugasan berikutnya'],
    NGT: ['Night driving', 'Hindari berangkat ≥20:00 atau sebelum 05:00 kecuali darurat/berizin; bila terpaksa lapor atasan & pastikan unit layak.', 'Evaluasi kebutuhan perjalanan malam; wajibkan izin & komunikasi berkala; jadwalkan ulang pekerjaan non-darurat.', '≤ 1 minggu'],
    BRK: ['Harsh braking', 'Jaga jarak aman (min. 3 detik, 4–5 detik di jalan hauling/basah); antisipasi lebih awal.', 'Cek kondisi rem & ban unit; identifikasi ruas rawan dari data VT; ride-along singkat untuk melihat gaya mengemudi.', '≤ 2 minggu'],
    ACC: ['Harsh acceleration', 'Injak gas bertahap, terutama bermuatan dan di tanjakan.', 'Coaching smooth driving; cek apakah ada tekanan waktu dari jadwal pekerjaan.', '≤ 2 minggu'],
    COR: ['Harsh cornering', 'Kurangi kecepatan sebelum masuk tikungan, bukan saat di tikungan.', 'Petakan tikungan rawan; cek sensitivitas sensor dengan Tim VT sebelum menilai perilaku.', '≤ 2 minggu'],
    OK: ['Grade A/B tanpa temuan', 'Pertahankan. Bagikan tips aman ke rekan di safety talk.', 'Apresiasi terbuka; jadikan mentor rute untuk pengemudi baru.', '–']
  };
  /** Alasan faktor koreksi (menu Faktor koreksi): aksi bawaan + komponen */
  VT.CORR_REASONS = [
    { id: 'did_pinjam', label: 'DID tag dipinjam / dipakai orang lain (pemakai diketahui)', en: 'DID tag lent to / used by another person (actual driver known)', action: 'reassign' },
    { id: 'did_tertukar', label: 'DID tag tertukar / terbawa orang lain (pemakai tidak diketahui)', en: 'DID tag swapped or carried by someone else (actual driver unknown)', action: 'exclude' },
    { id: 'gps_error', label: 'Error GPS / sensor terverifikasi Tim VT atau vendor', en: 'GPS / sensor error verified by VT team or vendor', action: 'annul', comps: ['spd', 'fat', 'brk', 'acc', 'cor'] },
    { id: 'gps_spike', label: 'Lonjakan GPS (kecepatan rata-rata trip > 120 km/jam)', en: 'GPS spike (trip average speed > 120 km/h)', action: 'exclude' },
    { id: 'darurat', label: 'Keadaan darurat medis / evakuasi', en: 'Medical emergency / evacuation', action: 'annul', comps: ['spd', 'night', 'fat'] },
    { id: 'kawal', label: 'Pengawalan resmi / konvoi berizin', en: 'Official escort / permitted convoy', action: 'annul', comps: ['spd'] },
    { id: 'jrm_malam', label: 'Perjalanan malam disetujui (JRM / izin atasan)', en: 'Night trip approved (JRM / supervisor permit)', action: 'annul', comps: ['night'] },
    { id: 'test_drive', label: 'Unit diuji mekanik / test drive setelah perbaikan', en: 'Vehicle road-tested by mechanic after repair', action: 'exclude' },
    { id: 'pihak3', label: 'Unit diderek / dioperasikan pihak ketiga (vendor, towing, rental)', en: 'Vehicle towed / operated by a third party (vendor, towing, rental)', action: 'exclude' },
    { id: 'duplikat', label: 'Data ganda / duplikasi trip', en: 'Duplicate data / duplicated trip', action: 'exclude' },
    { id: 'lain', label: 'Alasan lain (wajib keterangan & bukti)', en: 'Other reason (description & evidence required)', action: 'annul', comps: ['spd'] }];
  VT.CORR_ACTIONS = { reassign: 'Pindahkan ke pemakai sebenarnya', exclude: 'Keluarkan trip dari skor', annul: 'Anulir komponen tertentu' };
  /** Faktor lain (menu Faktor lain) */
  VT.FACTOR_TYPES = [
    { id: 'insiden', label: 'Insiden / kecelakaan kendaraan', en: 'Vehicle incident / accident' },
    { id: 'promosi', label: 'Promosi / perubahan jabatan', en: 'Promotion / role change' },
    { id: 'mutasi', label: 'Mutasi / pindah cabang', en: 'Branch transfer' },
    { id: 'cuti', label: 'Cuti / sakit / tidak bertugas', en: 'Leave / sick / off duty' },
    { id: 'keluar', label: 'Resign / keluar / habis kontrak', en: 'Resigned / left / contract ended' },
    { id: 'pelatihan', label: 'Pelatihan / coaching selesai', en: 'Training / coaching completed' },
    { id: 'nondriver', label: 'Bukan pengemudi personal (ID bersama / pool)', en: 'Not a personal driver (shared / pool ID)' }];

  VT.params = function (over) {
    const p = JSON.parse(JSON.stringify(VT.DEFAULTS));
    if (over) for (const k in over) if (over[k] !== undefined && over[k] !== null && over[k] !== '') p[k] = over[k];
    return p;
  };

  /* ------------------------------------------------------------ util */
  const norm = s => String(s == null ? '' : s).replace(/ /g, ' ').replace(/\s+/g, ' ').trim().toUpperCase();
  const normLbl = s => String(s == null ? '' : s).replace(/\s+/g, ' ').trim();
  VT.norm = norm; VT.normLbl = normLbl;
  const isGeneric = n => /^(DRIVER|MASTER)\b/.test(n);
  VT.isGeneric = isGeneric;
  const MON = ['Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun', 'Jul', 'Agu', 'Sep', 'Okt', 'Nov', 'Des'];
  const MONL = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];
  const MONEN = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
  const HARI = ['Minggu', 'Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu'];
  VT.MON = MON; VT.MONL = MONL; VT.MONEN = MONEN;
  const dayToDate = d => new Date((d - 25569) * 86400000);
  VT.dayToDate = dayToDate;
  VT.dateToDay = dt => Math.round(Date.UTC(dt.getFullYear(), dt.getMonth(), dt.getDate()) / 86400000) + 25569;
  VT.isoToDay = s => { if (!s) return null; const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(s); return m ? Date.UTC(+m[1], +m[2] - 1, +m[3]) / 86400000 + 25569 : null; };
  VT.dayToIso = d => { const t = dayToDate(d); return t.toISOString().slice(0, 10); };
  const dtxt = d => { const t = dayToDate(Math.floor(d)); return t.getUTCDate() + ' ' + MON[t.getUTCMonth()] + ' ' + t.getUTCFullYear(); };
  VT.dtxt = dtxt;
  VT.dtxtEN = d => { const t = dayToDate(Math.floor(d)); return t.getUTCDate() + ' ' + MONEN[t.getUTCMonth()].slice(0, 3) + ' ' + t.getUTCFullYear(); };
  const hm = s => { const secs = Math.round((s - Math.floor(s)) * 86400) % 86400; return String(Math.floor(secs / 3600)).padStart(2, '0') + '.' + String(Math.floor(secs % 3600 / 60)).padStart(2, '0'); };
  const hourOf = s => Math.floor((Math.round((s - Math.floor(s)) * 86400) % 86400) / 3600);
  VT.hourOf = hourOf;
  // format angka gaya Indonesia (FIXED di Excel lokal ID)
  const fx = (v, dp) => { if (v == null || !isFinite(v)) return '–'; const s = Math.abs(v).toFixed(dp); const [a, b] = s.split('.'); const g = a.replace(/\B(?=(\d{3})+(?!\d))/g, '.'); return (v < 0 && +s !== 0 ? '-' : '') + g + (b ? ',' + b : ''); };
  VT.fx = fx;
  const fy = day => { const t = dayToDate(day); return t.getUTCMonth() >= 3 ? t.getUTCFullYear() : t.getUTCFullYear() - 1; };
  VT.fyOf = fy;
  const monthKey = day => { const t = dayToDate(day); return t.getUTCFullYear() + '-' + String(t.getUTCMonth() + 1).padStart(2, '0'); };
  VT.monthKey = monthKey;
  VT.monthLabel = (mk, en) => { const [y, m] = mk.split('-'); return (en ? MONEN : MONL)[+m - 1] + ' ' + y; };
  function leaderCode(name, code) { if (!code) return false; return name.startsWith(code + ' ') || name.indexOf('(' + code + ')') >= 0; }
  VT.leaderCode = leaderCode;

  /* ------------------------------------------------------------ parsing nilai */
  function parseDateStr(s) {
    s = String(s).trim(); if (!s) return null; let m;
    if ((m = /^(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?/.exec(s))) return mk(+m[1], +m[2], +m[3], +(m[4] || 0), +(m[5] || 0), +(m[6] || 0));
    if ((m = /^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{4})(?:[ T](\d{1,2})[:.](\d{2})(?:[:.](\d{2}))?\s*(AM|PM)?)?/i.exec(s))) {
      let d = +m[1], mo = +m[2]; if (mo > 12 && d <= 12) { const t = d; d = mo; mo = t; }
      let h = +(m[4] || 0); if (m[7]) { if (/PM/i.test(m[7]) && h < 12) h += 12; if (/AM/i.test(m[7]) && h === 12) h = 0; }
      return mk(+m[3], mo, d, h, +(m[5] || 0), +(m[6] || 0));
    }
    const t = Date.parse(s); if (!isNaN(t)) { const dt = new Date(t); return mk(dt.getFullYear(), dt.getMonth() + 1, dt.getDate(), dt.getHours(), dt.getMinutes(), dt.getSeconds()); }
    return null;
    function mk(y, mo, d, h, mi, se) { return Date.UTC(y, mo - 1, d) / 86400000 + 25569 + (h * 3600 + mi * 60 + se) / 86400; }
  }
  VT.parseDateStr = parseDateStr;
  const toSerial = v => (typeof v === 'number' ? v : v instanceof Date ? VT.dateToDay(v) + (v.getHours() * 3600 + v.getMinutes() * 60 + v.getSeconds()) / 86400 : v == null || v === '' ? null : parseDateStr(v));
  function toHours(v) {
    if (v == null || v === '') return 0;
    if (typeof v === 'number') return v < 2 ? v * 24 : v; // fraksi hari (format waktu Excel)
    const s = String(v).trim(); const p = s.split(':');
    if (p.length >= 2) return (+p[0] || 0) + (+p[1] || 0) / 60 + (+(p[2] || 0) || 0) / 3600;
    const n = parseFloat(s.replace(',', '.')); return isNaN(n) ? 0 : n;
  }
  const toNum = v => { if (typeof v === 'number') return v; if (v == null || v === '') return 0; const n = parseFloat(String(v).replace(/\s/g, '').replace(',', '.')); return isNaN(n) ? 0 : n; };

  /* ------------------------------------------------------------ deteksi & impor file */
  const H = s => String(s == null ? '' : s).replace(/^\uFEFF/, '').replace(/\s+/g, ' ').trim().toLowerCase();
  const CT_COLS = {
    cabang: ['cabang'], cabnew: ['cabang new'], area: ['area'], t0: ['tanggal id/on', 'tanggal on', 'start time'], t1: ['tanggal engine off', 'end time'],
    dtl: ['driving time'], bw: ['beweekly', 'biweekly', 'bi-weekly'], month: ['month', 'bulan'], km: ['mileage (km)', 'mileage', 'km'], remark: ['remark'],
    vid: ['vehicle id'], plate2: ['plate number2'], plate: ['plate number'], h: ['driving time2'], acc: ['harsh acceleration count'], brk: ['harsh braking count'],
    cor: ['harsh turning count', 'harsh cornering count'], spd: ['over speed driving count', 'overspeed driving count'], fat: ['fatigue driving count'],
    drv: ['driver'], did: ['serial did'], fy: ['fy']
  };
  function mapHeader(row) {
    const idx = {}; const hs = row.map(H);
    for (const k in CT_COLS) { for (const cand of CT_COLS[k]) { const i = hs.indexOf(cand); if (i >= 0) { idx[k] = i; break; } } }
    return idx;
  }
  const isCartrackHeader = row => { const hs = row.map(H); return hs.includes('driver') && (hs.includes('mileage (km)') || hs.includes('mileage')) && hs.includes('remark'); };

  /** Deteksi jenis file: 'cartrack' | 'ga' | 'master' | 'unknown' (+ info sheet/header) */
  VT.detectXlsx = async function (blob) {
    const wb = await X.openWorkbook(blob);
    const order = [...wb.sheets].sort((a, b) => (b.name === 'Data') - (a.name === 'Data'));
    const names = wb.sheets.map(s => s.name.toLowerCase());
    if (names.some(n => n.includes('overspeed')) && names.some(n => n.includes('night'))) {
      return { type: 'ga', wb };
    }
    for (const sh of order) {
      if (/^(pivot|case)$/i.test(sh.name)) continue;
      let hdr = null, hr = 0;
      await X.readRows(wb, sh, (rn, v) => {
        if (isCartrackHeader(v)) { hdr = v; hr = rn; return false; }
        const hs = v.map(H);
        if (hs.some(x => x && x.startsWith('nama did')) && hs.some(x => x && x.indexOf('atasan langsung') >= 0)) { hdr = v; hr = -rn; return false; }
        if (rn > 80) return false;
      });
      if (hdr && hr > 0) return { type: 'cartrack', wb, sheet: sh, headerRow: hr, header: hdr };
      if (hdr && hr < 0) return { type: 'master', wb, sheet: sh, headerRow: -hr, header: hdr };
    }
    return { type: 'unknown', wb };
  };

  function tripFrom(v, ix, stats) {
    const remark = String(v[ix.remark] == null ? '' : v[ix.remark]).trim();
    const drv = norm(v[ix.drv]);
    const bw = normLbl(v[ix.bw]);
    const t0 = toSerial(v[ix.t0]);
    if (!bw || t0 == null) { stats.skipped++; return null; }
    stats.rows++;
    stats.byPeriod[bw] = stats.byPeriod[bw] || { rows: 0, idle: 0, drive: 0, minDay: 1e9, maxDay: 0 };
    const bp = stats.byPeriod[bw]; bp.rows++;
    const dday = Math.floor(t0 + 1e-9); if (dday < bp.minDay) bp.minDay = dday; if (dday > bp.maxDay) bp.maxDay = dday;
    if (remark !== 'Driving') { if (remark === 'Idling') { stats.idle++; bp.idle++; } else stats.other++; return null; }
    stats.drive++; bp.drive++;
    const km = toNum(v[ix.km]);
    const mo = ix.month != null && v[ix.month] != null && v[ix.month] !== '' ? toSerial(v[ix.month]) : null;
    const t = {
      d: drv, cab: String(v[ix.cabnew != null ? ix.cabnew : ix.cabang] || '').trim(), ar: String(v[ix.area] || '').trim(),
      t0, t1: ix.t1 != null ? toSerial(v[ix.t1]) : null, dl: String(v[ix.dtl] || '').trim() === 'Diatas Jam 7 Malam' ? 1 : 0,
      bw, mo: mo != null ? monthKey(Math.floor(mo + 1e-9)) : monthKey(dday), km,
      pl: String(v[ix.plate2 != null ? ix.plate2 : ix.plate] || '').trim(), vid: String(v[ix.vid] == null ? '' : v[ix.vid]),
      h: toHours(v[ix.h]), acc: toNum(v[ix.acc]), brk: toNum(v[ix.brk]), cor: toNum(v[ix.cor]), spd: toNum(v[ix.spd]), fat: toNum(v[ix.fat])
    };
    t.k = t.vid + '|' + Math.round(t0 * 86400) + '|' + drv;
    return t;
  }

  /** Impor data mentah Cartrack (.xlsx) → { trips (hanya Driving), stats } */
  VT.importCartrackXlsx = async function (det, onProgress) {
    const ix = mapHeader(det.header);
    const need = ['drv', 'km', 'remark', 't0', 'bw'];
    const miss = need.filter(k => ix[k] == null);
    if (miss.length) throw new Error('Kolom wajib tidak ditemukan: ' + miss.join(', ') + '. Pastikan memakai format laporan "Data Score Behavior" Cartrack.');
    const trips = []; const stats = { rows: 0, drive: 0, idle: 0, other: 0, skipped: 0, byPeriod: {}, sheet: det.sheet.name, headerRow: det.headerRow };
    await X.readRows(det.wb, det.sheet, (rn, v) => {
      if (rn <= det.headerRow) return;
      if (v[ix.drv] == null && v[ix.t0] == null) return;
      const t = tripFrom(v, ix, stats); if (t) trips.push(t);
    }, onProgress);
    return { trips, stats };
  };

  /** Impor CSV (dipisah koma/titik koma/tab) dengan kolom yang sama */
  VT.importCartrackCsv = async function (blob, onProgress) {
    const reader = blob.stream().getReader(); const dec = new TextDecoder('utf-8');
    let buf = '', delim = null, ix = null, hdrDone = false, got = 0, row = [], field = '', inq = false;
    const trips = []; const stats = { rows: 0, drive: 0, idle: 0, other: 0, skipped: 0, byPeriod: {}, sheet: 'CSV', headerRow: 0 };
    let lines = 0;
    const endRow = () => {
      row.push(field); field = ''; lines++;
      if (!hdrDone) { if (isCartrackHeader(row)) { ix = mapHeader(row); hdrDone = true; stats.headerRow = lines; } else if (lines > 80) throw new Error('Header data Cartrack tidak ditemukan di CSV.'); }
      else { const t = tripFrom(row, ix, stats); if (t) trips.push(t); }
      row = [];
    };
    const feed = (s, final) => {
      if (delim === null) { const first = s.split(/\r?\n/).slice(0, 40).join('\n'); const c = (first.match(/;/g) || []).length, k = (first.match(/,/g) || []).length, tb = (first.match(/\t/g) || []).length; delim = tb > c && tb > k ? '\t' : c > k ? ';' : ','; }
      for (let i = 0; i < s.length; i++) {
        const ch = s[i];
        if (inq) { if (ch === '"') { if (s[i + 1] === '"') { field += '"'; i++; } else inq = false; } else field += ch; continue; }
        if (ch === '"') inq = true; else if (ch === delim) { row.push(field); field = ''; } else if (ch === '\n') endRow(); else if (ch !== '\r') field += ch;
      }
      if (final && (field || row.length)) endRow();
    };
    for (;;) { const { value, done } = await reader.read(); if (done) break; got += value.length; feed(dec.decode(value, { stream: true }), false); if (onProgress) onProgress(got, blob.size); }
    feed(dec.decode(), true);
    if (!hdrDone) throw new Error('Header data Cartrack tidak ditemukan di CSV.');
    return { trips, stats };
  };

  /** Impor master karyawan (format sheet 10 workbook Excel atau template CSV aplikasi) */
  VT.MASTER_FIELDS = [['did', 'Nama DID'], ['cab', 'Cabang'], ['nama', 'Nama lengkap'], ['nik', 'NIK'], ['email', 'Email pengemudi'], ['jab', 'Jabatan'],
    ['atasan', 'Atasan langsung – nama'], ['atasanJab', 'Atasan langsung – jabatan'], ['atasanEmail', 'Atasan langsung – email'],
    ['atasan2', 'Atasan tidak langsung – nama'], ['atasan2Jab', 'Atasan tidak langsung – jabatan'], ['atasan2Email', 'Atasan tidak langsung – email']];
  function mapMasterHeader(row) {
    const hs = row.map(H); const f = (...keys) => hs.findIndex(h => h && keys.every(k => h.indexOf(k) >= 0));
    return {
      did: f('nama did'), cab: f('cabang'), nama: f('nama lengkap'), nik: f('nik'), email: f('email pengemudi'),
      jab: (() => { const i = f('jabatan (isi'); return i >= 0 ? i : f('jabatan'); })(),
      atasan: f('atasan langsung', 'nama'), atasanJab: f('atasan langsung', 'jabatan'), atasanEmail: f('atasan langsung', 'email'),
      atasan2: f('atasan tidak langsung', 'nama'), atasan2Jab: f('atasan tidak langsung', 'jabatan'), atasan2Email: f('atasan tidak langsung', 'email')
    };
  }
  VT.importMasterRows = function (rows) {
    // baris header = baris yang punya sel "Nama DID" DAN sel lain tentang atasan (bukan baris judul/keterangan)
    let hi = rows.findIndex(r => { if (!r) return false; const hs = r.map(H); return hs.some(x => x && x.startsWith('nama did')) && hs.some(x => x && x.indexOf('atasan') >= 0); });
    if (hi < 0) hi = rows.findIndex(r => r && r.map(H).some(x => x && x.startsWith('nama did')));
    if (hi < 0) throw new Error('Kolom "Nama DID" tidak ditemukan.');
    const ix = mapMasterHeader(rows[hi]); const out = [];
    for (let i = hi + 1; i < rows.length; i++) {
      const r = rows[i] || []; const did = norm(r[ix.did]); if (!did) continue;
      const o = { did };
      for (const [k] of VT.MASTER_FIELDS) if (k !== 'did') o[k] = ix[k] >= 0 && r[ix[k]] != null ? String(r[ix[k]]).trim() : '';
      out.push(o);
    }
    return out;
  };
  VT.importMasterXlsx = async function (det) {
    const rows = await X.readSheetAll(det.wb, det.sheet, 5000);
    return VT.importMasterRows(rows);
  };

  /** Impor lampiran GA "Data Case" (sheet Overspeed / Fatigue Driving / Night Driving, blok bulan berdampingan) */
  VT.importGA = async function (det) {
    const out = []; const months = new Set();
    const katOf = n => /over/i.test(n) ? 'Overspeed' : /fatig/i.test(n) ? 'Fatigue' : /night/i.test(n) ? 'Night driving' : null;
    for (const sh of det.wb.sheets) {
      const kat = katOf(sh.name); if (!kat) continue;
      const rows = await X.readSheetAll(det.wb, sh, 3000);
      let hr = rows.findIndex(r => r && r.some(v => /branch\/project/i.test(String(v || ''))));
      if (hr < 0) continue;
      const blocks = [];
      rows[hr].forEach((v, ci) => { if (/branch\/project/i.test(String(v || ''))) blocks.push(ci); });
      for (const c0 of blocks) {
        // label bulan: cari "Period : ..." di atas header pada kolom blok
        let mk = null;
        for (let r = hr - 1; r >= Math.max(0, hr - 4) && !mk; r--) { const s = String((rows[r] || [])[c0] || ''); if (/period/i.test(s)) mk = parseMonthLabel(s); }
        if (!mk) { const s = String(rows[hr][c0 + 2] || ''); mk = parseMonthLabel(s); }
        if (!mk) continue; months.add(mk);
        for (let r = hr + 1; r < rows.length; r++) {
          const row = rows[r] || []; const unit = row[c0], nama = row[c0 + 1], n = row[c0 + 2];
          if (unit == null || nama == null || typeof n !== 'number') continue;
          out.push({ bulan: mk, kat, unit: String(unit).trim(), nama: norm(nama), kasus: n });
        }
      }
    }
    if (!out.length) throw new Error('Format lampiran GA tidak dikenali (butuh sheet Overspeed / Fatigue Driving / Night Driving dengan kolom Branch/Project/Department).');
    return { rows: out, months: [...months].sort() };
  };
  function parseMonthLabel(s) {
    const t = s.toLowerCase();
    const names = [['jan'], ['feb', 'peb'], ['mar'], ['apr'], ['mei', 'may'], ['jun'], ['jul'], ['agu', 'aug'], ['sep'], ['okt', 'oct'], ['nov'], ['des', 'dec']];
    const y = /(20\d{2})/.exec(t); if (!y) return null;
    for (let i = 0; i < 12; i++) if (names[i].some(n => new RegExp('\\b' + n).test(t))) return y[1] + '-' + String(i + 1).padStart(2, '0');
    return null;
  }
  VT.parseMonthLabel = parseMonthLabel;

  /** CSV kecil → array baris (untuk master karyawan / log) */
  VT.csvRows = function (text) {
    const first = text.split(/\r?\n/).slice(0, 20).join('\n');
    const c = (first.match(/;/g) || []).length, k = (first.match(/,/g) || []).length, tb = (first.match(/\t/g) || []).length;
    const delim = tb > c && tb > k ? '\t' : c > k ? ';' : ',';
    const rows = []; let row = [], f = '', q = false;
    for (let i = 0; i < text.length; i++) {
      const ch = text[i];
      if (q) { if (ch === '"') { if (text[i + 1] === '"') { f += '"'; i++; } else q = false; } else f += ch; continue; }
      if (ch === '"') q = true; else if (ch === delim) { row.push(f); f = ''; } else if (ch === '\n') { row.push(f); rows.push(row); row = []; f = ''; } else if (ch !== '\r') f += ch;
    }
    if (f || row.length) { row.push(f); rows.push(row); }
    return rows;
  };

  /* ------------------------------------------------------------ periode */
  /** periods: dari statistik per label (minDay/maxDay). PID urut tanggal mulai. */
  VT.buildPeriods = function (pstats) {
    const arr = Object.keys(pstats).map(lbl => ({ lbl, start: pstats[lbl].minDay, end: pstats[lbl].maxDay, rows: pstats[lbl].rows, idle: pstats[lbl].idle, drive: pstats[lbl].drive }));
    arr.sort((a, b) => a.start - b.start || (a.lbl < b.lbl ? -1 : 1));
    arr.forEach((p, i) => { p.pid = i + 1; p.fy = fy(p.start); p.days = p.end - p.start + 1; });
    return arr;
  };

  /* ------------------------------------------------------------ faktor koreksi (tingkat trip) */
  function corrMatches(c, t, pidOf) {
    if (norm(c.nama) !== t.d) return false;
    if (c.trips && c.trips.length) return c.trips.indexOf(t.k) >= 0;
    if (c.pid && pidOf[t.bw] !== +c.pid) return false;
    const day = Math.floor(t.t0 + 1e-9);
    const df = VT.isoToDay(c.dari), ds = VT.isoToDay(c.sampai);
    if (df != null && day < df) return false;
    if (ds != null && day > ds) return false;
    if (c.plat && norm(c.plat).replace(/\s/g, '') !== norm(t.pl).replace(/\s/g, '')) return false;
    if (!c.pid && df == null && ds == null && !c.plat) return false; // lingkup wajib dibatasi
    return true;
  }
  VT.corrMatches = corrMatches;
  /** Terapkan koreksi yang DISETUJUI. Hasil: array trip baru + jejak audit per koreksi. */
  VT.applyCorrections = function (trips, corrections, pidOf, P) {
    const act = (corrections || []).filter(c => c.status === 'Disetujui');
    const audit = {};
    for (const c of act) audit[c.id] = { trips: 0, spd: 0, fat: 0, brk: 0, acc: 0, cor: 0, nightTrips: 0, km: 0 };
    if (!act.length) return { trips, audit };
    const out = [];
    for (const t0 of trips) {
      let t = t0, excluded = false;
      for (const c of act) {
        if (excluded) break;
        if (!corrMatches(c, t, pidOf)) continue;
        const a = audit[c.id]; a.trips++; a.km += t.km;
        if (t === t0) t = Object.assign({}, t0);
        const isNight = (() => { const hr = hourOf(t.t0); return hr >= P.NSTART || hr < P.NEND; })();
        if (c.aksi === 'exclude') { a.spd += t.spd; a.fat += t.fat; a.brk += t.brk; a.acc += t.acc; a.cor += t.cor; if (isNight && !t.nn) a.nightTrips++; excluded = true; break; }
        if (c.aksi === 'reassign') { const nd = norm(c.pemakai); if (nd) { t.d = nd; t.reassigned = c.id; } else excluded = true; continue; }
        if (c.aksi === 'annul') {
          for (const k of (c.komponen || [])) {
            if (k === 'night') { if (isNight && !t.nn) a.nightTrips++; t.nn = 1; }
            else if (t[k] != null) { a[k] += t[k]; t[k] = 0; }
          }
          t.annulled = (t.annulled || []).concat(c.id);
        }
      }
      if (!excluded) out.push(t);
    }
    return { trips: out, audit };
  };

  /* ------------------------------------------------------------ agregasi driver × periode */
  function modeOf(cnt) { let best = null, bn = -1; for (const k in cnt) { if (k === '' && Object.keys(cnt).length > 1) continue; if (cnt[k] > bn || (cnt[k] === bn && k < best)) { best = k; bn = cnt[k]; } } return best || ''; }
  VT.aggregate = function (trips, pidOf, P, excludeNames) {
    const map = new Map();
    const dq = { generic: 0, genericIds: new Set(), genericKm: 0, nondriver: 0, v120: 0, v120os: 0, tripsDrive: 0 };
    for (const t of trips) {
      const pid = pidOf[t.bw]; if (!pid) continue;
      dq.tripsDrive++;
      if (t.h > 0 && t.km / t.h > 120) { dq.v120++; if (t.spd > 0) dq.v120os++; }
      if (isGeneric(t.d)) { dq.generic++; dq.genericIds.add(t.d); dq.genericKm += t.km; continue; }
      if (excludeNames && excludeNames.has(t.d)) { dq.nondriver++; continue; }
      const key = t.d + '|' + pid;
      let r = map.get(key);
      if (!r) { r = { drv: t.d, PID: pid, km: 0, h: 0, trips: 0, so_sum: 0, so_n: 0, spd_raw: 0, fat: 0, brk: 0, acc: 0, cor: 0, b1: 0, b2: 0, b3: 0, b4: 0, b5: 0, sus: 0, nh: 0, nt: 0, nt_old: 0, _cab: {}, _ar: {}, osRaw: 0 }; map.set(key, r); }
      r.km += t.km; r.h += t.h; r.trips++;
      if (t.km > 0) {
        const c = x => Math.max(100 - x * 1000 / t.km, 0);
        r.so_sum += c(t.fat) * .25 + c(t.spd) * .25 + c(t.brk) * .2 + c(t.acc) * .2 + c(t.cor) * .1; r.so_n++;
      }
      const s = t.spd; r.spd_raw += s; r.fat += t.fat; r.brk += t.brk; r.acc += t.acc; r.cor += t.cor;
      if (s >= 1 && s <= 5 && s === Math.floor(s)) r['b' + s] += s;
      if (s >= P.OSSUS) r.sus++;
      const hr = hourOf(t.t0);
      if ((hr >= P.NSTART || hr < P.NEND) && !t.nn) { r.nh += t.h; r.nt++; }
      if (t.dl) r.nt_old++;
      r._cab[t.cab] = (r._cab[t.cab] || 0) + 1; r._ar[t.ar] = (r._ar[t.ar] || 0) + 1;
    }
    const rows = [];
    for (const r of map.values()) { r.cab = modeOf(r._cab); r.area = modeOf(r._ar); r.S_old = r.so_n ? r.so_sum / r.so_n : null; delete r._cab; delete r._ar; rows.push(r); }
    dq.genericIds = dq.genericIds.size;
    return { rows, dq };
  };

  /* ------------------------------------------------------------ skor */
  const bandOf = (km, P) => { const lo = P.BAND_LO; return km < lo[1] ? 'L1' : km < lo[2] ? 'L2' : km < lo[3] ? 'L3' : km < lo[4] ? 'L4' : 'L5'; };
  VT.bandOf = bandOf;
  VT.gradeOf = (v, P) => v >= P.GRADE[0] ? 'A' : v >= P.GRADE[1] ? 'B' : v >= P.GRADE[2] ? 'C' : v >= P.GRADE[3] ? 'D' : 'E';

  /** ctx: {periods, master(Map did→obj), factors{incidents,promotions,transfers,...}} */
  VT.score = function (rows, P, ctx) {
    const per = ctx.periods; const pend = {}; per.forEach(p => { pend[p.pid] = p.end; });
    rows.sort((a, b) => (a.drv < b.drv ? -1 : a.drv > b.drv ? 1 : a.PID - b.PID));
    const tolMask = k => k <= P.OSTOL;
    const roleF = {}; P.ROLE.forEach(r => { roleF[r[0]] = r[1]; });
    const sevD = {}, sevL = {}; P.SEV.forEach(s => { sevD[s[0]] = s[1]; sevL[s[0]] = s[2]; });
    const faultV = {}; P.FAULT.forEach(f => { faultV[f[0]] = f[1]; });
    const F = ctx.factors || {};
    const incidents = (F.insiden || []).filter(x => x.nama || x.cabang).map(x => ({ nm: norm(x.nama), cab: String(x.cabang || '').trim(), day: VT.isoToDay(x.tanggal), sevD: sevD[x.keparahan] || 0, sevL: sevL[x.keparahan] || 0, fault: faultV[x.kesalahan] != null ? faultV[x.kesalahan] : 1, capa: x.capa === 'Ya' ? P.CAPA : 1 }));
    const promos = (F.promosi || []).map(x => ({ nm: norm(x.nama), day: VT.isoToDay(x.tanggal), jab: x.jabatan })).filter(x => x.nm && x.day != null).sort((a, b) => a.day - b.day);
    const transfers = (F.mutasi || []).map(x => ({ nm: norm(x.nama), day: VT.isoToDay(x.tanggal), cab: String(x.cabang || '').trim() })).filter(x => x.nm && x.day != null && x.cab).sort((a, b) => a.day - b.day);
    const leaves = (F.keluar || []).map(x => ({ nm: norm(x.nama), day: VT.isoToDay(x.tanggal) })).filter(x => x.nm && x.day != null);
    const master = ctx.master || new Map();
    const codeFR = (P.ROLE.find(r => r[0] === 'Foreman') || [])[2] || 'FR', codeCSS = (P.ROLE.find(r => r[0] === 'CS Supervisor') || [])[2] || 'CSS';
    const dec = (pe, d) => (d == null || d <= 0 || d > pe || pe - d > P.DECAY) ? 0 : (1 - (pe - d) / P.DECAY);
    const kkm = P.KKM / 1000, kh = P.KH / 100;
    let prev = null;
    for (const r of rows) {
      const pe = pend[r.PID];
      // mutasi: cabang efektif
      for (const t of transfers) if (t.nm === r.drv && t.day <= pe) r.cab = t.cab;
      r.band = bandOf(r.km, P);
      let tol = 0; for (let k = 1; k <= 5; k++) if (tolMask(k)) tol += r['b' + k];
      r.spd = r.spd_raw - tol;
      const C = { spd: r.spd, fat: r.fat, night: r.nh, brk: r.brk, acc: r.acc, cor: r.cor };
      let vb = 0;
      for (const c of VT.COMP) {
        const b = P.BM[r.band][c.k]; const ex = c.ex === 'km' ? r.km : r.h; const k = c.ex === 'km' ? kkm : kh;
        r['b_' + c.k] = b;
        r['r_' + c.k] = ex > 0 ? C[c.k] / ex * c.m : 0;
        r['R_' + c.k] = ((C[c.k] + k * b) / (ex / c.m + k)) / b;
        r['S_' + c.k] = 100 * Math.pow(P.BASE, r['R_' + c.k]);
        vb += P.W[c.k] * r['S_' + c.k];
      }
      r.VB = vb;
      const pr = prev && prev.drv === r.drv ? prev : null;
      const cons = !!(pr && r.PID - pr.PID === 1);
      r.cons = cons; r.prevRow = pr && cons ? pr : null;
      let pen = 0, nrep = 0;
      for (const c of VT.REPC) {
        const cnt = c === 'night' ? 'nh' : c;
        const rc = r['r_' + c];
        let F_ = 0;
        if (cons) {
          const prc = pr[cnt], prr = pr['r_' + c];
          const above = c === 'spd' ? true : rc > r['b_' + c];
          const both = r[cnt] > 0 && prc > 0 && above;
          const rep = both && (rc >= (1 - P.TOL) * prr || rc >= P.CRIT * r['b_' + c]);
          if (rep) F_ = rc > (1 + P.TOL) * prr ? 2 : 1;
        }
        r['F_' + c] = F_;
        const base = (F_ === 2 ? P.PBUR : F_ === 1 ? P.PTET : 0) * (c === 'spd' ? P.OSMULT : 1);
        const streak = F_ > 0 && cons && pr['F_' + c] > 0 ? P.STREAK : 1;
        r['P_' + c] = base * streak; pen += r['P_' + c]; if (F_ > 0) nrep++;
      }
      r.REP = Math.min(pen, P.REPCAP); r.NREP = nrep;
      r.LCR = r.spd === 0 && nrep === 0 ? P.LC[VT.LAB.indexOf(r.band)] : 0;
      // jabatan: promosi > master > kode nama
      const m = master.get(r.drv);
      let jab = '';
      if (leaderCode(r.drv, codeFR)) jab = 'Foreman'; else if (leaderCode(r.drv, codeCSS)) jab = 'CS Supervisor';
      if (m && m.jab && !jab) jab = m.jab;
      for (const p of promos) if (p.nm === r.drv && p.day <= pe) jab = p.jab;
      r.JAB = jab; r.RF = roleF[jab] || 0;
      // insiden
      let ins = 0, ldr = 0;
      for (const x of incidents) {
        const d = dec(pe, x.day); if (!d) continue;
        if (x.nm === r.drv) ins += d * x.sevD * x.fault;
        if (r.RF && x.cab === r.cab && x.nm !== r.drv) ldr += d * x.sevL * x.capa;
      }
      r.INS = ins; r.LDR = r.RF ? Math.min(P.LCAP, r.RF * ldr) : 0;
      r.FIN = Math.max(0, Math.min(100, r.VB + r.LCR - r.REP - r.INS - r.LDR));
      r.GR = VT.gradeOf(r.FIN, P);
      r.DIF = r.S_old == null ? null : r.FIN - r.S_old;
      // status overspeed
      r.OSS = r.spd === 0 ? (r.spd_raw > 0 ? 'HIJAU (toleransi)' : 'HIJAU') : ((r.sus > 0 || r.F_spd > 0 || r.spd / Math.max(r.km, 1) * 1000 > P.OSKUN) ? 'MERAH' : 'KUNING');
      // komponen terlemah & fokus
      let wi = 0, wv = Infinity; VT.COMP.forEach((c, i) => { if (r['S_' + c.k] < wv) { wv = r['S_' + c.k]; wi = i; } });
      r.WK = VT.COMP[wi].name;
      r.FOC = r.INS > 0 ? 'INS' : r.LDR > 0 ? 'LDR' : r.F_spd > 0 ? 'OSR' : r.spd > 0 ? 'OSV' : r.REP > 0 ? 'REP' : r.FIN >= P.GRADE[1] ? 'OK' : VT.COMP[wi].code;
      // notifikasi
      let nt = 0;
      for (const c of VT.REPC) {
        const cnt = c === 'night' ? 'nh' : c;
        const out = c === 'spd' ? r.OSS === 'MERAH' : (r['R_' + c] >= P.NOTIFR && r[cnt] >= P.NOTIFMIN);
        r['N_' + c] = r['F_' + c] > 0 && out ? 1 : 0; nt += r['N_' + c];
      }
      r.NOTIF = nt;
      r.ESK = nt > 0 ? (cons && pr.NOTIF > 0 ? 2 : 1) : 0;
      // status keluar (resign)
      r.KELUAR = leaves.some(l => l.nm === r.drv && per.find(p => p.pid === r.PID).start > l.day) ? 1 : 0;
      prev = r;
    }
    return rows;
  };

  /** Kalibrasi ulang rata-rata kelompok dari data (port calibrate() Python) */
  VT.calibrate = function (rows, P, pidFrom, pidTo) {
    const cal = rows.filter(r => r.PID >= pidFrom && r.PID <= pidTo);
    const tolSum = r => { let t = 0; for (let k = 1; k <= 5; k++) if (k <= P.OSTOL) t += r['b' + k]; return r.spd_raw - t; };
    const s = { km: 0, h: 0, spd: 0, fat: 0, brk: 0, acc: 0, cor: 0, nh: 0 };
    cal.forEach(r => { s.km += r.km; s.h += r.h; s.spd += tolSum(r); s.fat += r.fat; s.brk += r.brk; s.acc += r.acc; s.cor += r.cor; s.nh += r.nh; });
    const NAT = { spd: s.spd / s.km * 1000, night: s.nh / s.h, fat: s.fat / s.h * 100, brk: s.brk / s.km * 1000, acc: s.acc / s.km * 1000, cor: s.cor / s.km * 1000 };
    const BM = {};
    for (const b of VT.LAB) {
      const g = cal.filter(r => bandOf(r.km, P) === b); const t = { km: 0, h: 0, fat: 0, brk: 0, acc: 0, cor: 0 };
      g.forEach(r => { t.km += r.km; t.h += r.h; t.fat += r.fat; t.brk += r.brk; t.acc += r.acc; t.cor += r.cor; });
      const raw = { fat: t.fat / t.h * 100, brk: t.brk / t.km * 1000, acc: t.acc / t.km * 1000, cor: t.cor / t.km * 1000 };
      BM[b] = {};
      for (const k in raw) BM[b][k] = Math.round(Math.max(isFinite(raw[k]) ? raw[k] : 0, NAT[k] * 0.5) * 100) / 100;
      BM[b].spd = Math.round(NAT.spd * 100) / 100;
      BM[b].night = Math.round(Math.round(NAT.night * 1e4) / 1e4 * 100 * 100) / 100;
    }
    return { BM, NAT, n: cal.length };
  };

  /* ------------------------------------------------------------ detail trip (isi email) */
  VT.tripViol = function (t, P) {
    const hr = hourOf(t.t0); const night = (hr >= P.NSTART || hr < P.NEND) && !t.nn;
    const parts = [];
    if (t.spd > P.OSTOL) parts.push('overspeed ' + t.spd + 'x'); else if (t.spd > 0) parts.push('overspeed singkat ' + t.spd + 'x (ditoleransi)');
    if (t.fat > 0) parts.push('fatigue ' + t.fat + 'x');
    if (t.brk > 0) parts.push('rem mendadak ' + t.brk + 'x');
    if (t.acc > 0) parts.push('akselerasi kasar ' + t.acc + 'x');
    if (night) parts.push('mulai mengemudi malam');
    return { night, text: parts.join(', '), any: parts.length > 0 };
  };
  VT.tripLine = function (t, P) {
    const v = VT.tripViol(t, P); const d = Math.floor(t.t0 + 1e-9);
    const dt = dayToDate(d);
    return HARI[dt.getUTCDay()] + ', ' + dtxt(d) + ' pukul ' + hm(t.t0) + (t.t1 != null ? '–' + hm(t.t1) : '') + ' · ' + t.pl + ' · ' + fx(t.km, 0) + ' km · ' + (v.text || '—') + (t.lok ? ' · lokasi: ' + t.lok : '') + ' (cabang ' + t.cab + ')';
  };

  /* ------------------------------------------------------------ notifikasi email */
  VT.buildNotifications = function (res, pid, P, master, tripsByKey) {
    const per = res.periods.find(p => p.pid === pid); if (!per) return [];
    const rows = res.rows.filter(r => r.PID === pid && r.NOTIF > 0 && !r.KELUAR);
    const out = [];
    const nightTxt = String(P.NSTART).padStart(2, '0') + ':00–' + String((P.NEND + 23) % 24).padStart(2, '0') + ':59';
    rows.forEach((r, i) => {
      const m = master.get(r.drv) || {};
      const esk = r.ESK === 2;
      const nama = m.nama || r.drv;
      const periode = per.lbl + ' (' + dtxt(per.start) + ' – ' + dtxt(per.end) + ')';
      const cc = [];
      if (esk && m.atasan2Email) cc.push(m.atasan2Email);
      if (P.CCDRV === 'Ya' && m.email) cc.push(m.email);
      if (P.CCQSHE) cc.push(P.CCQSHE);
      const to = (m.atasanEmail || '').trim();
      const pr = r.prevRow;
      const jab = r.JAB || 'pengemudi';
      const trend = c => (r['F_' + c] === 2 ? 'makin sering' : 'belum berkurang');
      let body = 'Yth. Bapak/Ibu ' + (m.atasan || 'Atasan') + ',\n\n' +
        'Kami informasikan bahwa ' + nama + ' (' + jab + ', ' + r.cab + ') kembali tercatat melakukan pelanggaran berkendara yang sama dalam dua periode 2-mingguan berturut-turut, dan nilainya sudah di luar batas yang dapat diterima.' +
        (esk ? ' Pola ini sudah muncul juga di notifikasi periode sebelumnya, sehingga kami tembuskan ke ' + (m.atasan2 || 'atasan tidak langsung') + '.' : '') +
        '\n\nRINGKASAN PERIODE ' + periode.toUpperCase() + '\n' +
        '• Jarak tempuh: ' + fx(r.km, 0) + ' km, ' + fx(r.h, 1) + ' jam mengemudi, ' + r.trips + ' trip (band beban ' + r.band + ')\n' +
        '• Skor periode ini: ' + fx(r.FIN, 1) + ' (grade ' + r.GR + ')' + (pr ? '; periode lalu ' + fx(pr.FIN, 1) : '');
      body += '\n\nPELANGGARAN YANG BERULANG & DI LUAR STANDAR';
      if (r.N_spd) body += '\n• Overspeed: ' + r.spd + ' kali dihitung (periode lalu ' + (pr ? pr.spd : 0) + '), ' + trend('spd') + (r.sus > 0 ? '; termasuk ' + r.sus + ' trip ngebut beruntun (≥10 count dalam satu trip)' : '');
      if (r.N_fat) body += '\n• Fatigue (indikasi lelah/mengantuk): ' + r.fat + ' kali = ' + fx(r.r_fat, 1) + ' per 100 jam; rata-rata kelompok beban yang sama ' + fx(r.b_fat, 1) + ' (periode lalu ' + (pr ? pr.fat : 0) + ' kali), ' + trend('fat');
      if (r.N_night) body += '\n• Mengemudi malam (mulai ' + nightTxt + '): ' + fx(r.nh, 1) + ' jam = ' + fx(r.r_night, 1) + '% dari waktu mengemudi; rata-rata ' + fx(r.b_night, 1) + '% (periode lalu ' + fx(pr ? pr.nh : 0, 1) + ' jam), ' + trend('night');
      if (r.N_brk) body += '\n• Rem mendadak (harsh braking): ' + r.brk + ' kali = ' + fx(r.r_brk, 1) + ' per 1.000 km; rata-rata kelompok ' + fx(r.b_brk, 1) + ' (periode lalu ' + (pr ? pr.brk : 0) + ' kali), ' + trend('brk');
      if (r.N_acc) body += '\n• Akselerasi kasar (harsh acceleration): ' + r.acc + ' kali = ' + fx(r.r_acc, 1) + ' per 1.000 km; rata-rata kelompok ' + fx(r.b_acc, 1) + ' (periode lalu ' + (pr ? pr.acc : 0) + ' kali), ' + trend('acc');
      const tl = (tripsByKey.get(r.drv + '|' + pid) || []);
      body += '\n\nKAPAN & DI MANA (trip tercatat, maks. 10, urut waktu)';
      tl.slice(0, 10).forEach((t, j) => { body += '\n' + (j + 1) + '. ' + VT.tripLine(t, P); });
      if (tl.length > 10) body += '\n… dan ' + (tl.length - 10) + ' trip lain (lihat menu Detail pengemudi).';
      if (!tl.length) body += '\n(detail trip tidak tersedia)';
      const code = r.N_spd ? 'OSR' : r.N_fat ? 'FAT' : r.N_night ? 'NGT' : r.N_brk ? 'BRK' : 'ACC';
      body += '\n\nYANG KAMI MOHON DARI BAPAK/IBU (paling lambat ' + P.DEADLINE + ' hari kerja)\n' +
        '1. Ajak ' + nama + ' bicara empat mata. Tanyakan apa yang membuat pola ini terulang: jadwal terlalu padat, rute, kondisi unit, atau kelelahan.\n' +
        '2. ' + VT.REK[code][2] + '\n' +
        '3. Sepakati satu target sederhana untuk 2 minggu ke depan, lalu balas email ini dengan ringkasan hasil pembicaraan.\n\n' +
        'Pesan ini bertujuan mencegah kecelakaan, bukan mencari kesalahan. Terima kasih atas dukungan Bapak/Ibu menjaga tim selamat sampai di rumah.\n\n' +
        'Salam keselamatan,\n' + P.SENDER + '\n\n(Pesan otomatis dari Vehicle Telematic Score and Report – data Vehicle Telematics/Cartrack.)';
      out.push({
        no: i + 1, kirim: to.indexOf('@') >= 0 ? 'YA' : 'TIDAK – email atasan belum ada', level: esk ? 'ESKALASI' : 'PERHATIAN', nama: r.drv, namaLengkap: nama, cab: r.cab,
        periode, atasan: m.atasan || '', to, cc: cc.join(';'), subject: (esk ? '[ESKALASI] ' : '[Perlu tindak lanjut] ') + 'Pola berkendara berulang – ' + nama + ' (' + r.cab + ') – periode ' + per.lbl,
        body, code, row: r, trips: tl.length
      });
    });
    return out;
  };

  /* ------------------------------------------------------------ pipeline penuh */
  /**
   * compute({trips, pstats, inputs}) → hasil lengkap untuk UI.
   * inputs: {params, corrections, factors, master:[...], scoreFromPid}
   */
  VT.compute = function (data) {
    const inp = data.inputs || {};
    const P = VT.params(inp.params);
    const periods = VT.buildPeriods(data.pstats);
    const pidOf = {}; periods.forEach(p => { pidOf[p.lbl] = p.pid; });
    const master = new Map(); (inp.master || []).forEach(m => { if (m.did) master.set(norm(m.did), m); });
    const F = inp.factors || {};
    const nondriver = new Set((F.nondriver || []).map(x => norm(x.nama)).filter(Boolean));
    const from = inp.scoreFromPid || 1;
    let trips = data.trips.filter(t => (pidOf[t.bw] || 0) >= from);
    const cor = VT.applyCorrections(trips, inp.corrections, pidOf, P);
    trips = cor.trips;
    const ag = VT.aggregate(trips, pidOf, P, nondriver);
    const rows = VT.score(ag.rows, P, { periods, master, factors: F });
    // detail trip per driver|PID (urut waktu)
    const tripsByKey = new Map();
    for (const t of trips) {
      if (isGeneric(t.d) || nondriver.has(t.d)) continue;
      const pid = pidOf[t.bw]; if (!pid) continue;
      if (!VT.tripViol(t, P).any) continue;
      const k = t.d + '|' + pid; let a = tripsByKey.get(k); if (!a) { a = []; tripsByKey.set(k, a); } a.push(t);
    }
    for (const a of tripsByKey.values()) a.sort((x, y) => x.t0 - y.t0);
    const lastPid = inp.lastPid && periods.find(p => p.pid === +inp.lastPid) ? +inp.lastPid : (() => {
      // periode laporan terakhir = periode terbaru dengan ≥ 7 hari data
      for (let i = periods.length - 1; i >= 0; i--) if (periods[i].days >= 7 && rows.some(r => r.PID === periods[i].pid)) return periods[i].pid;
      return periods.length ? periods[periods.length - 1].pid : 0;
    })();
    return { P, periods, rows, tripsByKey, corAudit: cor.audit, dq: ag.dq, lastPid, master, pidOf, tripsUsed: trips };
  };

  /* ------------------------------------------------------------ analitik (dashboard, rekap) */
  const avg = a => (a.length ? a.reduce((s, x) => s + x, 0) / a.length : null);
  const sum = (a, f) => a.reduce((s, x) => s + (f ? f(x) : x), 0);
  VT.avg = avg; VT.sum = sum;
  VT.sel = (res, pid, cab) => res.rows.filter(r => r.PID === pid && (!cab || cab === 'Semua cabang' || r.cab === cab));
  VT.kpis = function (res, pid, cab) {
    const s = VT.sel(res, pid, cab), sp = VT.sel(res, pid - 1, cab);
    const km = sum(s, r => r.km), h = sum(s, r => r.h);
    const o = {
      n: s.length, fin: avg(s.map(r => r.FIN)), finPrev: avg(sp.map(r => r.FIN)), old: avg(s.filter(r => r.S_old != null).map(r => r.S_old)),
      km, kmPer: s.length ? km / s.length : 0, ab: s.length ? s.filter(r => r.GR === 'A' || r.GR === 'B').length / s.length : 0,
      de: s.filter(r => r.GR === 'D' || r.GR === 'E').length, notif: s.filter(r => r.NOTIF > 0).length, esk: s.filter(r => r.ESK === 2).length,
      rep: s.filter(r => r.NREP > 0).length, merah: s.filter(r => r.OSS === 'MERAH').length, kuning: s.filter(r => r.OSS === 'KUNING').length,
      osv: sum(s, r => r.spd), nh: sum(s, r => r.nh), h, brk: km ? sum(s, r => r.brk) / km * 1000 : 0,
      brkPrev: sum(sp, r => r.km) ? sum(sp, r => r.brk) / sum(sp, r => r.km) * 1000 : null
    };
    o.nhPct = h ? o.nh / h * 100 : 0;
    o.grades = 'ABCDE'.split('').map(g => s.filter(r => r.GR === g).length);
    o.bands = VT.LAB.map(b => { const g = s.filter(r => r.band === b); return { band: b, n: g.length, old: avg(g.filter(r => r.S_old != null).map(r => r.S_old)) || 0, fin: avg(g.map(r => r.FIN)) || 0 }; });
    return o;
  };
  VT.trend = function (res, endPid, cab, n) {
    n = n || 11; const out = [];
    for (let i = 0; i < n; i++) {
      const pid = endPid - n + 1 + i; const p = res.periods.find(x => x.pid === pid); if (!p) continue;
      const s = VT.sel(res, pid, cab); const km = sum(s, r => r.km);
      const t = dayToDate(p.start);
      out.push({ pid, lbl: t.getUTCDate() + ' ' + MON[t.getUTCMonth()], label: p.lbl, old: avg(s.filter(r => r.S_old != null).map(r => r.S_old)), fin: avg(s.map(r => r.FIN)),
        osv: km ? sum(s, r => r.spd) / km * 1000 : null, brk: km ? sum(s, r => r.brk) / km * 1000 : null, acc: km ? sum(s, r => r.acc) / km * 1000 : null, n: s.length });
    }
    return out;
  };
  VT.branches = function (res, pid, P) {
    const s = res.rows.filter(r => r.PID === pid); const g = {};
    s.forEach(r => { (g[r.cab] = g[r.cab] || []).push(r); });
    return Object.keys(g).map(c => { const a = g[c]; return { cab: c, n: a.length, fin: avg(a.map(r => r.FIN)), de: a.filter(r => r.GR === 'D' || r.GR === 'E').length / a.length, notif: a.filter(r => r.NOTIF > 0).length, rep: a.filter(r => r.NREP > 0).length }; })
      .sort((a, b) => (a.cab < b.cab ? -1 : 1));
  };
  VT.fyPids = (res, pid) => { const p = res.periods.find(x => x.pid === pid); if (!p) return []; return res.periods.filter(x => x.fy === p.fy && x.pid <= pid).map(x => x.pid); };
  VT.recapDrivers = function (res, pid, master, P) {
    const pids = new Set(VT.fyPids(res, pid)); const g = new Map();
    res.rows.forEach(r => { if (pids.has(r.PID)) { let a = g.get(r.drv); if (!a) { a = []; g.set(r.drv, a); } a.push(r); } });
    const out = [];
    for (const [d, a] of g) {
      const last = a.find(r => r.PID === pid); const m = master.get(d) || {};
      const fin = avg(a.map(r => r.FIN)), old = avg(a.filter(r => r.S_old != null).map(r => r.S_old));
      out.push({ drv: d, cab: a[a.length - 1].cab, jab: a[a.length - 1].JAB || m.jab || '', periods: a.length, km: sum(a, r => r.km), kmPer: sum(a, r => r.km) / a.length,
        band: last ? last.band : 'tidak aktif', old, fin, dif: old == null ? null : fin - old, lastFin: last ? last.FIN : null, lastGr: last ? last.GR : '–', trend: last ? last.FIN - fin : null,
        repP: a.filter(r => r.NREP > 0).length, rep: sum(a, r => r.REP), osv: sum(a, r => r.spd), notif: a.filter(r => r.NOTIF > 0).length, ins: last ? last.INS : 0,
        foc: last ? last.FOC : '–', atasan: m.atasan || '', atasanEmail: m.atasanEmail || '', series: a.map(r => [r.PID, r.FIN]) });
    }
    const rk = (arr, k) => { const v = arr.map(x => x[k]).filter(x => x != null); arr.forEach(x => { x['rank_' + k] = x[k] == null ? null : v.filter(y => y > x[k]).length + 1; }); };
    rk(out, 'fin'); rk(out, 'old');
    return out.sort((a, b) => (a.drv < b.drv ? -1 : 1));
  };
  VT.recapBranches = function (res, pid, incidents) {
    const pids = new Set(VT.fyPids(res, pid)); const g = new Map();
    res.rows.forEach(r => { if (pids.has(r.PID)) { let a = g.get(r.cab); if (!a) { a = []; g.set(r.cab, a); } a.push(r); } });
    const out = [];
    for (const [c, a] of g) {
      const ar = {}; a.forEach(r => { ar[r.area] = (ar[r.area] || 0) + 1; });
      const fin = avg(a.map(r => r.FIN)), old = avg(a.filter(r => r.S_old != null).map(r => r.S_old));
      out.push({ cab: c, area: modeOf(ar), drivers: new Set(a.map(r => r.drv)).size, km: sum(a, r => r.km), old, fin, dif: old == null ? null : fin - old,
        de: a.filter(r => r.GR === 'D' || r.GR === 'E').length / a.length, rep: a.filter(r => r.NREP > 0).length, osv: sum(a, r => r.spd), notif: a.filter(r => r.NOTIF > 0).length,
        ins: (incidents || []).filter(x => String(x.cabang || '').trim() === c).length });
    }
    const v = out.map(x => x.fin); out.forEach(x => { x.rank = v.filter(y => y > x.fin).length + 1; });
    const vo = out.map(x => x.old).filter(x => x != null); out.forEach(x => { x.rankOld = x.old == null ? null : vo.filter(y => y > x.old).length + 1; });
    return out.sort((a, b) => (a.cab < b.cab ? -1 : 1));
  };
  VT.recapLeaders = function (res, pid, P) {
    const s = res.rows.filter(r => r.PID === pid);
    const leaders = new Map();
    const roleF = {}; P.ROLE.forEach(r => { roleF[r[0]] = r[1]; });
    s.forEach(r => { if (r.RF > 0) leaders.set(r.drv, { drv: r.drv, jab: r.JAB, cab: r.cab, rf: r.RF, own: r.FIN, ldr: r.LDR, src: 'Skor mengemudi sendiri' }); });
    // pemimpin dari master tanpa data VT di periode ini
    for (const [d, m] of res.master) {
      if (leaders.has(d)) continue;
      const jab = leaderCode(d, 'FR') ? 'Foreman' : leaderCode(d, 'CSS') ? 'CS Supervisor' : m.jab;
      if (!roleF[jab]) continue;
      const team = s.filter(r => r.cab === m.cab);
      const teamAvg = avg(team.map(r => r.FIN));
      leaders.set(d, { drv: d, jab, cab: m.cab, rf: roleF[jab], own: null, team: teamAvg, ldr: 0, src: 'Rata-rata tim − pengurang' });
    }
    const out = [...leaders.values()];
    out.forEach(l => { l.team = l.team != null ? l.team : avg(s.filter(r => r.cab === l.cab).map(r => r.FIN)); l.score = l.own != null ? l.own : (l.team != null ? Math.max(0, l.team - l.ldr) : null); l.gr = l.score != null ? VT.gradeOf(l.score, P) : ''; });
    return out.sort((a, b) => (a.drv < b.drv ? -1 : 1));
  };
  VT.fairness = function (res, pid) {
    const pids = new Set(VT.fyPids(res, pid)); const s = res.rows.filter(r => pids.has(r.PID));
    return VT.LAB.map(b => { const g = s.filter(r => r.band === b); const n = g.length || 1; return { band: b, n: g.length, km: avg(g.map(r => r.km)) || 0, old: avg(g.filter(r => r.S_old != null).map(r => r.S_old)) || 0, vb: avg(g.map(r => r.VB)) || 0, rep: avg(g.map(r => r.REP)) || 0, lcr: avg(g.map(r => r.LCR)) || 0, fin: avg(g.map(r => r.FIN)) || 0, a: g.filter(r => r.GR === 'A').length / n, de: g.filter(r => r.GR === 'D' || r.GR === 'E').length / n }; });
  };
  /** Statistik bulanan langsung dari trip (untuk tren & perbandingan bulanan) */
  VT.monthly = function (res, cab) {
    const P = res.P; const g = {};
    for (const t of res.tripsUsed) {
      if (isGeneric(t.d)) continue;
      const pid = res.pidOf[t.bw]; if (!pid) continue;
      if (cab && cab !== 'Semua cabang' && t.cab !== cab) continue;
      const m = g[t.mo] = g[t.mo] || { mo: t.mo, km: 0, h: 0, spd: 0, spdc: 0, fat: 0, brk: 0, acc: 0, cor: 0, nOld: 0, nNew: 0, n20: 0, nh: 0, trips: 0, drv: new Set() };
      m.km += t.km; m.h += t.h; m.spd += t.spd; if (t.spd > P.OSTOL) m.spdc += t.spd; m.fat += t.fat; m.brk += t.brk; m.acc += t.acc; m.cor += t.cor; m.trips++; m.drv.add(t.d);
      const hr = hourOf(t.t0); if (t.dl) m.nOld++; if ((hr >= P.NSTART || hr < P.NEND) && !t.nn) { m.nNew++; m.nh += t.h; } if (hr >= 20) m.n20++;
    }
    return Object.values(g).sort((a, b) => (a.mo < b.mo ? -1 : 1)).map(m => Object.assign(m, { drivers: m.drv.size, drv: undefined }));
  };

  /* ------------------------------------------------------------ laporan bulanan GA */
  VT.gaAnalysis = function (ga, month, P, master, log) {
    const prevMonth = (() => { const [y, m] = month.split('-').map(Number); const d = new Date(Date.UTC(y, m - 2, 1)); return d.getUTCFullYear() + '-' + String(d.getUTCMonth() + 1).padStart(2, '0'); })();
    const rel = ga.rows.filter(r => r.bulan === month || r.bulan === prevMonth);
    const people = new Map();
    const K = { Overspeed: 'os', Fatigue: 'ft', 'Night driving': 'nd' };
    for (const r of rel) {
      let o = people.get(r.nama);
      if (!o) { o = { nama: r.nama, units: {}, os_prv: 0, os_cur: 0, ft_prv: 0, ft_cur: 0, nd_prv: 0, nd_cur: 0 }; people.set(r.nama, o); }
      const b = r.bulan === month ? 'cur' : 'prv'; o[K[r.kat] + '_' + b] += r.kasus;
      o.units[r.unit] = (o.units[r.unit] || 0) + r.kasus * (b === 'cur' ? 1000 : 1);
    }
    const st = (p, c) => (p > 0 && c > 0 ? (c >= p ? 'Tidak turun' : 'Turun') : c > 0 ? 'Baru' : p > 0 ? 'Bersih' : '');
    const nm = { os: 'overspeed', ft: 'fatigue', nd: 'night-driving' };
    const list = [...people.values()].map(o => {
      const S = { os: st(o.os_prv, o.os_cur), ft: st(o.ft_prv, o.ft_cur), nd: st(o.nd_prv, o.nd_cur) };
      const naik = ['os', 'ft', 'nd'].filter(k => S[k] === 'Tidak turun');
      const nk = ['os', 'ft', 'nd'].filter(k => o[k + '_cur'] > 0).length;
      const tot = o.os_cur + o.ft_cur + o.nd_cur, totp = o.os_prv + o.ft_prv + o.nd_prv;
      const generik = isGeneric(o.nama);
      let pr, al;
      const lst = naik.map(k => nm[k]).join(' dan ');
      if (o.os_cur >= P.GAP1OS) { pr = 'P1'; al = 'Overspeed ' + o.os_cur + ' kasus dalam sebulan'; }
      else if (naik.length >= 2) { pr = 'P1'; al = 'Tidak turun di ' + lst; }
      else if (nk === 3) { pr = 'P1'; al = 'Tercatat di ketiga kategori'; }
      else if (naik.length === 1 && o[naik[0] + '_cur'] >= P.GAP2) { pr = 'P2'; al = 'Tidak turun di ' + lst; }
      else if (tot > 0) { pr = 'P3'; al = naik.length === 1 ? 'Tidak turun di ' + lst + ', kasus sedikit' : 'Baru atau sudah turun'; }
      else { pr = 'Membaik'; al = 'Tercatat bulan lalu, bersih bulan ini'; }
      if (generik) al += '; ID umum, telusuri pemakai';
      const unit = Object.entries(o.units).sort((a, b) => b[1] - a[1]).map(x => x[0]).join(' / ');
      const m = master.get(o.nama);
      const lg = k => { const e = (log || []).find(x => norm(x.nama) === o.nama && x.bulan === k); return e ? e.status : ''; };
      return { nama: o.nama, unit, tipe: /branch/i.test(unit) ? 'Cabang' : 'Project/Dept', os_prv: o.os_prv, os_cur: o.os_cur, ft_prv: o.ft_prv, ft_cur: o.ft_cur, nd_prv: o.nd_prv, nd_cur: o.nd_cur,
        st_os: S.os, st_ft: S.ft, st_nd: S.nd, naik: naik.length, nk, tot, totp, pr, al, generik, pemimpin: /^(FR|CSS|PSR) /.test(o.nama),
        atasan: m ? (m.atasan || '(belum diisi di master)') : '(nama belum ada di master)', atasanEmail: m ? m.atasanEmail || '' : '',
        tlPrev: lg(prevMonth) || '–', tlCur: lg(month) || 'Belum dicatat' };
    });
    const ord = { P1: 4, P2: 3, P3: 2, Membaik: 1 };
    list.sort((a, b) => ord[b.pr] - ord[a.pr] || b.tot - a.tot || b.totp - a.totp || (a.nama < b.nama ? -1 : 1));
    const S = (kat, b, f) => ga.rows.filter(r => r.kat === kat && r.bulan === b && (!f || f(r))).reduce((s, r) => s + r.kasus, 0);
    const cat = ['Overspeed', 'Fatigue', 'Night driving'].map(kat => {
      const k = K[kat]; const prv = S(kat, prevMonth), cur = S(kat, month);
      const tt = list.filter(x => x['st_' + k] === 'Tidak turun').reduce((s, x) => s + x[k + '_cur'], 0);
      const tu = list.filter(x => x['st_' + k] === 'Turun').reduce((s, x) => s + x[k + '_cur'], 0);
      const bar = list.filter(x => x['st_' + k] === 'Baru').reduce((s, x) => s + x[k + '_cur'], 0);
      return { kat, k, prv, cur, chg: prv ? cur / prv - 1 : null, org: list.filter(x => x[k + '_cur'] > 0).length, rep: cur ? (tt + tu) / cur : 0, tt, tu, bar };
    });
    const units = {};
    ga.rows.filter(r => r.bulan === month || r.bulan === prevMonth).forEach(r => {
      const u = units[r.unit] = units[r.unit] || { unit: r.unit, tipe: /branch/i.test(r.unit) ? 'Cabang' : 'Project/Dept', os_prv: 0, os_cur: 0, ft_prv: 0, ft_cur: 0, nd_prv: 0, nd_cur: 0 };
      u[K[r.kat] + '_' + (r.bulan === month ? 'cur' : 'prv')] += r.kasus;
    });
    const ul = Object.values(units).map(u => Object.assign(u, { tot_cur: u.os_cur + u.ft_cur + u.nd_cur, tot_prv: u.os_prv + u.ft_prv + u.nd_prv }))
      .map(u => Object.assign(u, { sel: u.tot_cur - u.tot_prv, org: list.filter(x => x.unit === u.unit && x.tot > 0).length, p1: list.filter(x => x.unit === u.unit && x.pr === 'P1').length }))
      .sort((a, b) => b.tot_cur - a.tot_cur || b.tot_prv - a.tot_prv);
    const topOs = list.reduce((b, x) => (x.os_cur > (b ? b.os_cur : -1) ? x : b), null);
    const k = {
      month, prevMonth, cat, list, units: ul,
      orgCur: list.filter(x => x.tot > 0).length, orgNd: list.filter(x => x.nd_cur > 0).length, orgOs: list.filter(x => x.os_cur > 0).length,
      p1: list.filter(x => x.pr === 'P1').length, p2: list.filter(x => x.pr === 'P2').length, p3: list.filter(x => x.pr === 'P3').length, mb: list.filter(x => x.pr === 'Membaik').length,
      topOs, gen: S('Overspeed', month, r => isGeneric(r.nama)) + S('Fatigue', month, r => isGeneric(r.nama)) + S('Night driving', month, r => isGeneric(r.nama)),
      genOs: S('Overspeed', month, r => isGeneric(r.nama)), lead: list.filter(x => x.pemimpin && x.tot > 0).length,
      ndProj: S('Night driving', month, r => !/branch/i.test(r.unit)),
      tlN: (log || []).filter(x => x.bulan === prevMonth).length, tlOk: (log || []).filter(x => x.bulan === prevMonth && x.status === 'Selesai').length
    };
    const all = cat.reduce((s, c) => s + c.cur, 0); k.repAll = all ? cat.reduce((s, c) => s + c.rep * c.cur, 0) / all : 0; k.totCur = all;
    const topUnitOs = [...ul].sort((a, b) => b.os_cur - a.os_cur)[0]; k.topUnitOs = topUnitOs;
    k.topUnitOsOrg = topUnitOs ? new Set(ga.rows.filter(r => r.bulan === month && r.kat === 'Overspeed' && r.unit === topUnitOs.unit).map(r => r.nama)).size : 0;
    k.topUnitNd = [...ul].sort((a, b) => b.nd_cur - a.nd_cur)[0];
    return k;
  };

  root.VT = VT;
  if (typeof module !== 'undefined' && module.exports) module.exports = VT;
})(typeof self !== 'undefined' ? self : globalThis);
