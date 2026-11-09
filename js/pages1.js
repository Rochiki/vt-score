/* pages1.js — Panduan, Unggah, Notifikasi, Rekap, Laporan bilingual, Referensi */
'use strict';

/* ------------------------------------------------------------ PANDUAN */
PAGES.panduan = {
  title: 'Panduan', render(v) {
    const Pp = P(); const nt = String(Pp.NSTART).padStart(2, '0') + ':00–' + String((Pp.NEND + 23) % 24).padStart(2, '0') + ':59';
    v.innerHTML = head('Panduan pemakaian', 'Untuk Tim VT, atasan, dan QSHE. Semua angka & aturan sama dengan workbook Excel "Vehicle Telematic Score and Report".') + `<div class="pad guide">
    <h3 class="sec">Ringkasnya dalam 1 menit</h3>
    <ul><li>Unggah file mentah Cartrack ("Data Score Behavior", .xlsx atau .csv). Aplikasi membaca sheet <b>Data</b>, memakai trip <b>Driving</b> saja, lalu menghitung skor per pengemudi per periode 2 minggu (label Beweekly).</li>
    <li>Setiap pengemudi dibandingkan dengan rata-rata pengemudi di <b>band km yang sama</b> (L1–L5), jadi yang travel load-nya tinggi tidak otomatis kalah.</li>
    <li>Overspeed ≤${Pp.OSTOL} count per trip ditoleransi. Night driving = trip yang <b>mulai</b> ${nt}.</li>
    <li>Pelanggaran yang sama 2 periode berturut-turut → pengurang. Bila di luar standar (MERAH), atasan menerima email berisi kapan, di mana, dan apa pelanggarannya.</li>
    <li>Data diolah dan disimpan <b>di perangkat ini saja</b> (tidak dikirim ke server). Gunakan menu <a href="#/data">Data &amp; cadangan</a> untuk memindahkan data ke perangkat lain.</li></ul>
    <h3 class="sec">Langkah kerja tiap 2 minggu</h3>
    <ol class="steps">
      <li><b>Unggah data mentah Cartrack</b> di menu <a href="#/unggah">Unggah data</a>. Boleh file FY penuh atau file 2-mingguan; trip ganda otomatis dilewati.</li>
      <li><b>Cek Faktor koreksi</b>: DID dipinjam, error GPS, darurat, perjalanan malam ber-JRM, dll. Ajukan, lalu QSHE menyetujui — hanya yang <i>Disetujui</i> mengubah skor.</li>
      <li><b>Catat Faktor lain</b>: insiden/kecelakaan, promosi, mutasi, cuti, resign, pelatihan, ID bersama.</li>
      <li><b>Lengkapi Master karyawan</b> (email pengemudi, atasan langsung, atasan tidak langsung). Tanpa email atasan, notifikasi tertahan.</li>
      <li><b>Buka Dashboard</b>: baca executive report, KPI, dan 10 prioritas coaching.</li>
      <li><b>Kirim Notifikasi email atasan</b>: salin isi email / buka di aplikasi email, atau ekspor tabel untuk Power Automate.</li>
      <li><b>Laporan bulanan GA</b>: unggah lampiran "Data Case" dari GA, isi angka infografis, lalu catat tindak lanjut Prioritas 1–2 di Log tindak lanjut.</li>
      <li><b>Unduh Laporan bilingual (Word/PDF)</b> untuk manajemen.</li></ol>
    <h3 class="sec">Angka penting (bisa diubah QSHE di menu Parameter)</h3>
    <div class="kv">
      <div>Toleransi overspeed</div><div>≤ ${Pp.OSTOL} count per trip tidak dihitung (overspeed singkat, mis. menyalip)</div>
      <div>Ngebut beruntun</div><div>≥ ${Pp.OSSUS} count dalam satu trip → status overspeed MERAH</div>
      <div>Batas overspeed KUNING</div><div>≤ ${fx(Pp.OSKUN, 0)} per 1.000 km; lebih dari itu MERAH</div>
      <div>Jam malam</div><div>trip mulai ${nt} (19:00–19:59 tidak dihitung)</div>
      <div>Band travel load (km/2 minggu)</div><div>L1 &lt;${Pp.BAND_LO[1]} · L2 ${Pp.BAND_LO[1]}–${Pp.BAND_LO[2]} · L3 ${Pp.BAND_LO[2]}–${Pp.BAND_LO[3]} · L4 ${Pp.BAND_LO[3]}–${Pp.BAND_LO[4]} · L5 ≥${Pp.BAND_LO[4]}</div>
      <div>Skor komponen</div><div>100 × ${fx(Pp.BASE, 2)} ^ rasio · 0 kejadian = 100 · sama dengan rata-rata = 85 · 2× = 72</div>
      <div>Pengurang berulang</div><div>${Pp.PTET} (tetap) · ${Pp.PBUR} (makin buruk) · overspeed ×${Pp.OSMULT} · beruntun ×${fx(Pp.STREAK, 1)} · maks ${Pp.REPCAP}</div>
      <div>Kredit beban</div><div>L3 +${Pp.LC[2]} · L4 +${Pp.LC[3]} · L5 +${Pp.LC[4]} (bila overspeed 0 dan tidak berulang)</div>
      <div>Notifikasi atasan</div><div>berulang 2 periode + rasio ≥ ${fx(Pp.NOTIFR, 1)} dan ≥ ${Pp.NOTIFMIN} kejadian (overspeed: status MERAH); periode sebelumnya juga kena → ESKALASI</div>
      <div>Grade</div><div>A ≥${Pp.GRADE[0]} · B ≥${Pp.GRADE[1]} · C ≥${Pp.GRADE[2]} · D ≥${Pp.GRADE[3]} · E di bawahnya</div></div>
    <h3 class="sec">Faktor koreksi yang sah</h3>
    <ul>${VT.CORR_REASONS.map(r => `<li><b>${esc(r.label)}</b> → ${esc(VT.CORR_ACTIONS[r.action])}${r.comps ? ' (' + r.comps.map(c => VT.COMP.find(x => x.k === c).name).join(', ') + ')' : ''}</li>`).join('')}</ul>
    <p class="note">Tidak bisa dikoreksi: kejar target/jadwal, jalan sepi/lurus, "ikut arus", tidak tahu batas kecepatan, unit tanpa speed limiter. Setiap koreksi wajib punya nomor bukti dan persetujuan QSHE.</p>
    <h3 class="sec">Hal krusial</h3>
    <ul><li>Nama DID harus persis sama di data VT dan master karyawan (huruf besar/kecil & spasi ganda disamakan otomatis).</li>
    <li>Periode harus berurutan; aturan berulang membandingkan periode n dengan n−1.</li>
    <li>Jangan ubah count overspeed jadi 1/0 di data mentah — toleransi & status MERAH bergantung pada count per trip.</li>
    <li>Data di aplikasi tersimpan di browser. Membersihkan data situs akan menghapusnya: rutin unduh <b>cadangan</b> (menu Data &amp; cadangan) dan aktifkan "penyimpanan permanen".</li></ul>
    <h3 class="sec">Instal di perangkat</h3><p>Lihat menu <a href="#/instal">Instal aplikasi</a>: Android/Chrome/Edge → tombol <b>Pasang aplikasi</b>; iPhone/iPad → Safari → Bagikan → <b>Tambah ke Layar Utama</b>. Setelah terpasang, aplikasi jalan tanpa internet.</p></div>`;
  }
};

/* ------------------------------------------------------------ UNGGAH */
PAGES.unggah = {
  title: 'Unggah data', render(v) {
    v.innerHTML = head('Unggah data', 'Data mentah Cartrack, lampiran GA "Data Case", master karyawan, atau file cadangan. Bisa beberapa file sekaligus.') + `<div class="pad">
      <label class="drop" id="drop" tabindex="0"><input type="file" id="files" multiple accept=".xlsx,.csv,.txt,.json" hidden>
        <b>Klik atau seret file ke sini</b><div class="note">Didukung: <b>Data Score Behavior</b> Cartrack (.xlsx / .csv, bisa ratusan MB) · <b>Data Case</b> bulanan GA (.xlsx) · <b>Master karyawan</b> (.xlsx sheet 10 dari workbook Excel atau template .csv) · <b>Cadangan</b> aplikasi (.json)</div></label>
      <div id="jobs" style="margin-top:12px;display:grid;gap:8px"></div>
      <h3 class="sec">Riwayat unggah</h3><div id="hist"></div>
      <h3 class="sec">Template</h3><div class="tools"><button class="btn" id="tplMaster">Unduh template master karyawan (.csv)</button><a class="btn" href="#/periode">Lihat periode tersimpan</a></div></div>`;
    const inp = $('#files'), drop = $('#drop');
    inp.onchange = () => handleFiles([...inp.files]);
    drop.addEventListener('keydown', e => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); inp.click(); } });
    ['dragenter', 'dragover'].forEach(t => drop.addEventListener(t, e => { e.preventDefault(); drop.classList.add('over'); }));
    ['dragleave', 'drop'].forEach(t => drop.addEventListener(t, e => { e.preventDefault(); drop.classList.remove('over'); }));
    drop.addEventListener('drop', e => handleFiles([...e.dataTransfer.files]));
    $('#tplMaster').onclick = () => {
      const h = VT.MASTER_FIELDS.map(f => f[1]).join(';');
      const rows = S.res ? [...new Set(S.res.rows.filter(r => r.PID === S.res.lastPid).map(r => r.drv + ';' + r.cab))].map(x => x + ';;;;;;;;;;') : [];
      download(new Blob(['﻿' + h + '\n' + rows.join('\n')], { type: 'text/csv' }), 'template_master_karyawan.csv');
    };
    drawHist();
  }
};
function drawHist() {
  const h = $('#hist'); if (!h) return;
  const a = (S.inputs.imports || []).slice().reverse();
  h.innerHTML = a.length ? `<div class="tbox free"><table class="t"><thead><tr><th>Waktu</th><th>File</th><th>Jenis</th><th>Hasil</th></tr></thead><tbody>${a.map(x => `<tr><td>${esc(new Date(x.at).toLocaleString('id-ID'))}</td><td>${esc(x.name)}</td><td>${esc(x.kind)}</td><td>${esc(x.info)}</td></tr>`).join('')}</tbody></table></div>` : '<p class="note">Belum ada.</p>';
}
async function handleFiles(files) {
  const jobs = $('#jobs');
  let changed = false;
  for (const f of files) {
    const box = document.createElement('div'); box.className = 'card';
    box.innerHTML = `<b>${esc(f.name)}</b> <span class="muted small">${f0(f.size / 1048576)} MB</span><div class="prog"><i></i></div><div class="small muted st">Membaca…</div>`;
    jobs.prepend(box);
    const bar = $('.prog i', box), st = $('.st', box);
    try {
      const m = await importFile(f, pct => { bar.style.width = pct + '%'; st.textContent = 'Membaca… ' + pct + '%'; });
      bar.style.width = '100%';
      let info = '';
      if (m.kind === 'cartrack') {
        info = `${f0(m.stats.rows)} baris · ${f0(m.stats.drive)} trip Driving · ${f0(m.stats.idle)} Idling · ${m.stats.periods.length} periode · ${f0(m.stats.added)} trip baru · ${f0(m.stats.dup)} ganda dilewati`;
        changed = true;
      } else if (m.kind === 'ga') {
        const keep = S.ga.rows.filter(r => !m.ga.months.includes(r.bulan));
        S.ga.rows = keep.concat(m.ga.rows); S.ga.months = [...new Set(S.ga.rows.map(r => r.bulan))].sort();
        await saveGA();
        info = `${f0(m.ga.rows.length)} baris kasus · bulan ${m.ga.months.map(x => VT.monthLabel(x)).join(', ')}`;
      } else if (m.kind === 'master') {
        const map = new Map((S.inputs.master || []).map(x => [VT.norm(x.did), x]));
        let nw = 0, up = 0;
        m.master.forEach(x => { const o = map.get(x.did); if (o) { for (const k in x) if (x[k]) o[k] = x[k]; up++; } else { map.set(x.did, x); nw++; } });
        S.inputs.master = [...map.values()]; saveInputs(true); changed = true;
        info = `${nw} karyawan baru · ${up} diperbarui`;
      } else if (m.kind === 'backup') {
        info = await restoreBackup(m.json); changed = true;
      }
      st.innerHTML = '<b style="color:var(--ok)">Berhasil.</b> ' + esc(info);
      S.inputs.imports = (S.inputs.imports || []).concat({ at: Date.now(), name: f.name, kind: m.kind, info }).slice(-60); saveInputs();
    } catch (e) { bar.style.background = 'var(--bad)'; st.innerHTML = '<b style="color:var(--bad)">Gagal:</b> ' + esc(e.message); }
  }
  if (changed) {
    busy('Memuat ulang data & menghitung skor…');
    try { setTrips(await DB.periods()); compute(); } catch (e) { toast('Gagal menghitung: ' + e.message, 8000); } finally { busy(false); }
    renderSide(); toast('Data diperbarui. Buka Dashboard untuk melihat hasil.');
  }
  drawHist();
}

/* ------------------------------------------------------------ NOTIFIKASI */
PAGES.notifikasi = {
  title: 'Notifikasi email atasan', render(v, q) {
    if (needData(v)) return;
    let pid = +(q.get('pid') || S.res.lastPid);
    v.innerHTML = head('Notifikasi email ke atasan', 'Pemicu: pelanggaran yang sama berulang 2 periode berturut-turut DAN di luar standar keberterimaan. Eskalasi = periode sebelumnya juga kena notifikasi (atasan tidak langsung di-CC).') + `<div class="pad">
      <div class="tools"><label for="nPer" class="small"><b>Periode</b></label><select id="nPer">${periodOptions(pid)}</select><span class="sp"></span>
      <button class="btn" id="nCsv">Ekspor CSV (Power Automate)</button><button class="btn" id="nXls">Ekspor Excel</button></div>
      <div id="nSum" class="note"></div><div class="grid2" style="grid-template-columns:minmax(0,1.15fr) minmax(0,1fr)"><div id="nTab"></div><div><h3 class="sec" style="margin-top:8px">Pratinjau email</h3><div id="nPv"><p class="note">Klik satu baris untuk melihat email.</p></div></div></div></div>`;
    $('#nPer').onchange = e => { location.hash = '#/notifikasi?pid=' + e.target.value; };
    const nts = VT.buildNotifications(S.res, pid, P(), S.res.master, S.res.tripsByKey);
    $('#nSum').innerHTML = `<b>${nts.filter(x => x.kirim === 'YA').length}</b> siap kirim · <b>${nts.filter(x => x.kirim !== 'YA').length}</b> tertahan (email atasan belum ada di <a href="#/master">Master karyawan</a>) · <b>${nts.filter(x => x.level === 'ESKALASI').length}</b> eskalasi · pengemudi berstatus keluar tidak dikirimi.`;
    const show = x => {
      const ai = (S.settings.aiEmail || {})[x.nama + '|' + pid];
      const body = ai || x.body;
      $('#nPv').innerHTML = `<div class="mailpv"><div class="mh"><span class="subj">${esc(x.subject)}</span><div><b>Kepada</b><span>${esc(x.to || '(email atasan belum ada)')}</span></div><div><b>Cc</b><span>${esc(x.cc || '—')}</span></div><div><b>Status</b><span><span class="st ${x.kirim === 'YA' ? 'YA' : 'TIDAK'}">${esc(x.kirim)}</span> <span class="st ${x.level === 'ESKALASI' ? 'ESK' : 'PER'}">${x.level}</span></span></div></div><pre>${esc(body)}</pre></div>
      <div class="tools"><button class="btn pri" data-a="copy">Salin isi email</button><button class="btn" data-a="mail" ${x.to ? '' : 'disabled'}>Buka di aplikasi email</button><button class="btn" data-a="ai">Rapikan dengan AI</button>${ai ? '<button class="btn" data-a="orig">Kembali ke versi asli</button>' : ''}<a class="btn" href="#/pengemudi?d=${encodeURIComponent(x.nama)}&pid=${pid}">Detail pengemudi</a></div>`;
      $$('[data-a]', $('#nPv')).forEach(b => b.onclick = async () => {
        const a = b.dataset.a;
        if (a === 'copy') copyText('Subjek: ' + x.subject + '\n\n' + body);
        if (a === 'mail') { const u = 'mailto:' + encodeURIComponent(x.to) + '?' + (x.cc ? 'cc=' + encodeURIComponent(x.cc) + '&' : '') + 'subject=' + encodeURIComponent(x.subject) + '&body=' + encodeURIComponent(body.length > 1800 ? body.slice(0, 1800) + '\n…(lanjutan: tempel dari tombol Salin)' : body); location.href = u; }
        if (a === 'ai') { const t = await AI.tidy(x.body, 'email'); if (t) { S.settings.aiEmail = S.settings.aiEmail || {}; S.settings.aiEmail[x.nama + '|' + pid] = t; saveSettings(); show(x); } }
        if (a === 'orig') { delete S.settings.aiEmail[x.nama + '|' + pid]; saveSettings(); show(x); }
      });
    };
    table($('#nTab'), [
      { k: 'no', l: 'No', cls: 'c' }, { k: 'kirim', l: 'Kirim', f: x => `<span class="st ${x.kirim === 'YA' ? 'YA' : 'TIDAK'}">${x.kirim === 'YA' ? 'YA' : 'TIDAK'}</span>` },
      { k: 'level', l: 'Level', f: x => `<span class="st ${x.level === 'ESKALASI' ? 'ESK' : 'PER'}">${x.level}</span>` }, { k: 'namaLengkap', l: 'Nama', f: x => `<b>${esc(x.namaLengkap)}</b>` },
      { k: 'cab', l: 'Cabang' }, { k: 'atasan', l: 'Atasan langsung', f: x => esc(x.atasan || '—') }, { k: 'komp', l: 'Komponen', v: x => VT.REPC.filter(c => x.row['N_' + c]).map(c => VT.COMP.find(z => z.k === c).name).join(', ') }
    ], nts, { onRow: show, exportName: 'notifikasi_' + perLabel(pid) });
    if (nts.length) show(nts[0]);
    const cols = [['No', 'no'], ['Kirim', 'kirim'], ['Level', 'level'], ['Nama', 'nama'], ['Nama_Lengkap', 'namaLengkap'], ['Cabang', 'cab'], ['Periode', 'periode'], ['Atasan_Langsung', 'atasan'], ['Email_To', 'to'], ['Email_CC', 'cc'], ['Subjek', 'subject'], ['Isi_Email', 'body']];
    const bodyOf = x => (S.settings.aiEmail || {})[x.nama + '|' + pid] || x.body;
    $('#nCsv').onclick = () => {
      const q2 = s => '"' + String(s == null ? '' : s).replace(/"/g, '""') + '"';
      const csv = cols.map(c => c[0]).join(',') + '\n' + nts.map(x => cols.map(c => q2(c[1] === 'body' ? bodyOf(x) : x[c[1]])).join(',')).join('\n');
      download(new Blob(['﻿' + csv], { type: 'text/csv' }), 'tblNotifikasi_' + perLabel(pid).replace(/\W+/g, '_') + '.csv');
    };
    $('#nXls').onclick = () => exportXlsx('tblNotifikasi_' + perLabel(pid), [{ name: 'tblNotifikasi', cols: cols.map(c => ({ l: c[0], x: x => (c[1] === 'body' ? bodyOf(x) : x[c[1]]), w: c[1] === 'body' ? 80 : 22 })), rows: nts }]);
  }
};

/* ------------------------------------------------------------ REKAP */
function pidSelect(id, pid) { return `<label for="${id}" class="small"><b>Sampai periode</b></label><select id="${id}">${periodOptions(pid)}</select>`; }
PAGES['rekap-pengemudi'] = {
  title: 'Rekap pengemudi', render(v, q) {
    if (needData(v)) return;
    const pid = +(q.get('pid') || S.res.lastPid); const p = per(pid);
    v.innerHTML = head('Rekap pengemudi FY' + p.fy, 'Periode FY berjalan s.d. periode terpilih. Peringkat 1 = terbaik. Klik baris untuk detail perhitungan.') + `<div class="pad"><div class="tools">${pidSelect('rp', pid)}</div><div id="t"></div></div>`;
    $('#rp').onchange = e => { location.hash = '#/rekap-pengemudi?pid=' + e.target.value; };
    const rows = VT.recapDrivers(S.res, pid, S.res.master, P());
    table($('#t'), [
      { k: 'drv', l: 'Nama pengemudi (DID)', f: r => `<b>${esc(r.drv)}</b>` }, { k: 'cab', l: 'Cabang (terakhir)' }, { k: 'jab', l: 'Jabatan' }, { k: 'periods', l: 'Periode aktif', cls: 'c' },
      { k: 'km', l: 'Total KM', f: r => f0(r.km), cls: 'r', nf: '#,##0' }, { k: 'kmPer', l: 'KM/periode', f: r => f0(r.kmPer), cls: 'r' }, { k: 'band', l: 'Band terakhir', cls: 'c' },
      { k: 'old', l: 'Skor LAMA rata-rata', f: r => f1(r.old), cls: 'r' }, { k: 'fin', l: 'Skor BARU rata-rata', f: r => `<b>${f1(r.fin)}</b>`, cls: 'r' }, { k: 'dif', l: 'Selisih', f: r => sg(r.dif), cls: 'r' },
      { k: 'lastFin', l: 'Skor periode terakhir', f: r => f1(r.lastFin), cls: 'r' }, { k: 'lastGr', l: 'Grade', f: r => gradeChip(r.lastGr), cls: 'c' }, { k: 'trend', l: 'Tren', f: r => `<span class="${r.trend > 0 ? 'dn' : r.trend < 0 ? 'up' : ''}">${sg(r.trend)}</span>`, cls: 'r' },
      { k: 'repP', l: 'Periode berulang', cls: 'c' }, { k: 'osv', l: 'Overspeed dihitung', cls: 'r' }, { k: 'notif', l: 'Notifikasi atasan', cls: 'c' }, { k: 'rank_fin', l: 'Peringkat BARU', cls: 'c' }, { k: 'rank_old', l: 'Peringkat LAMA', cls: 'c' },
      { k: 'foc', l: 'Fokus' }, { k: 'atasan', l: 'Atasan langsung' }
    ], rows, { sort: 16, dir: 1, exportName: 'rekap_pengemudi_FY' + p.fy, onRow: r => { location.hash = '#/pengemudi?d=' + encodeURIComponent(r.drv) + '&pid=' + pid; } });
  }
};
PAGES['rekap-cabang'] = {
  title: 'Rekap cabang', render(v, q) {
    if (needData(v)) return;
    const pid = +(q.get('pid') || S.res.lastPid); const p = per(pid);
    v.innerHTML = head('Rekap cabang FY' + p.fy, 'Skor rata-rata = rata-rata semua driver-periode di cabang. % D/E = porsi driver-periode grade D atau E.') + `<div class="pad"><div class="tools">${pidSelect('rc', pid)}</div><div class="panel"><h4>Skor baru per cabang</h4><div class="chartbox" style="height:${Math.max(260, 0)}px"><canvas id="cCab"></canvas></div></div><div id="t" style="margin-top:12px"></div></div>`;
    $('#rc').onchange = e => { location.hash = '#/rekap-cabang?pid=' + e.target.value; };
    const rows = VT.recapBranches(S.res, pid, S.inputs.factors.insiden);
    const srt = rows.filter(r => r.drivers >= 1).slice().sort((a, b) => b.fin - a.fin);
    $('#cCab').parentElement.style.height = Math.max(260, srt.length * 18 + 40) + 'px';
    chart('cCab', { type: 'bar', data: { labels: srt.map(r => r.cab), datasets: [{ label: 'Skor baru', data: srt.map(r => r.fin), backgroundColor: srt.map(r => (r.fin < 80 ? '#c2410c' : '#f7941d')), borderRadius: 2, barThickness: 12 }, { label: 'Skor lama', data: srt.map(r => r.old), backgroundColor: '#cfcfcf', borderRadius: 2, barThickness: 12 }] },
      options: { indexAxis: 'y', scales: { x: Object.assign({ min: 50, max: 100 }, gridY), y: { grid: { display: false }, ticks: { autoSkip: false, font: { size: 11 } } } }, plugins: { legend: { position: 'bottom' } } } });
    table($('#t'), [
      { k: 'cab', l: 'Cabang', f: r => `<b>${esc(r.cab)}</b>` }, { k: 'area', l: 'Area' }, { k: 'drivers', l: 'Pengemudi aktif', cls: 'c' }, { k: 'km', l: 'Total KM', f: r => f0(r.km), cls: 'r' },
      { k: 'old', l: 'Skor LAMA', f: r => f1(r.old), cls: 'r' }, { k: 'fin', l: 'Skor BARU', f: r => `<b>${f1(r.fin)}</b>`, cls: 'r' }, { k: 'dif', l: 'Selisih', f: r => sg(r.dif), cls: 'r' },
      { k: 'de', l: '% D/E', f: r => pc(r.de, 1), cls: 'r' }, { k: 'rep', l: 'Driver-periode berulang', cls: 'c' }, { k: 'osv', l: 'Overspeed dihitung', cls: 'r' }, { k: 'notif', l: 'Notifikasi', cls: 'c' },
      { k: 'ins', l: 'Insiden (input)', cls: 'c' }, { k: 'rank', l: 'Peringkat BARU', cls: 'c' }, { k: 'rankOld', l: 'Peringkat LAMA', cls: 'c' }
    ], rows, { sort: 12, dir: 1, exportName: 'rekap_cabang_FY' + p.fy });
  }
};
PAGES['rekap-pemimpin'] = {
  title: 'Rekap pemimpin', render(v, q) {
    if (needData(v)) return;
    const pid = +(q.get('pid') || S.res.lastPid);
    v.innerHTML = head('Rekap pemimpin', 'FR = Foreman · CSS = CS Supervisor · Unit/Section Head dari master atau menu Promosi. Pemimpin tanpa data VT di periode ini memakai rata-rata skor tim cabangnya.') + `<div class="pad"><div class="tools"><label class="small" for="rl"><b>Periode</b></label><select id="rl">${periodOptions(pid)}</select></div><div id="t"></div></div>`;
    $('#rl').onchange = e => { location.hash = '#/rekap-pemimpin?pid=' + e.target.value; };
    table($('#t'), [
      { k: 'drv', l: 'Nama', f: r => `<b>${esc(r.drv)}</b>` }, { k: 'jab', l: 'Jabatan' }, { k: 'cab', l: 'Cabang' }, { k: 'rf', l: 'Faktor peran', f: r => fx(r.rf, 1), cls: 'c' },
      { k: 'own', l: 'Skor mengemudi sendiri', f: r => (r.own == null ? 'tidak ada VT' : f1(r.own)), cls: 'r' }, { k: 'team', l: 'Rata-rata tim cabang', f: r => f1(r.team), cls: 'r' },
      { k: 'ldr', l: 'Pengurang pemimpin', f: r => f1(r.ldr), cls: 'r' }, { k: 'score', l: 'SKOR PEMIMPIN', f: r => `<b>${f1(r.score)}</b>`, cls: 'r' }, { k: 'gr', l: 'Grade', f: r => gradeChip(r.gr), cls: 'c' }, { k: 'src', l: 'Dasar skor' }
    ], VT.recapLeaders(S.res, pid, P()), { exportName: 'rekap_pemimpin' });
  }
};

/* ------------------------------------------------------------ LAPORAN BILINGUAL */
PAGES.laporan = {
  title: 'Laporan bilingual', render(v, q) {
    if (needData(v)) return;
    const pid = +(q.get('pid') || S.res.lastPid); const mode = q.get('m') || 'bi';
    v.innerHTML = head('Laporan profesional bilingual', 'Catatan detail terpisah dari dashboard: ringkasan, metodologi, hasil, notifikasi, cabang, pemimpin, data GA, faktor koreksi, kualitas data, rekomendasi. Bahasa Indonesia & English.') + `<div class="pad">
      <div class="tools noprint"><label class="small" for="lp"><b>Periode</b></label><select id="lp">${periodOptions(pid)}</select>
      <label class="small" for="lm"><b>Bahasa</b></label><select id="lm"><option value="bi">Indonesia + English</option><option value="id">Indonesia</option><option value="en">English</option></select><span class="sp"></span>
      <button class="btn pri" id="lDoc">Unduh Word (.docx)</button><button class="btn" id="lPrint">Cetak / simpan PDF</button><button class="btn" id="lAi">Rapikan ringkasan dengan AI</button></div>
      <div class="card" id="rep" style="max-width:1000px"></div></div>`;
    $('#lm').value = mode;
    const go = () => { location.hash = '#/laporan?pid=' + $('#lp').value + '&m=' + $('#lm').value; };
    $('#lp').onchange = go; $('#lm').onchange = go;
    const R = VTR.buildReport(S.res, pid, S);
    const ai = (S.settings.aiReport || {})[pid];
    if (ai) R.sec[0].p = [{ id: ai, en: ai }];
    $('#rep').innerHTML = VTR.reportHtml(R, mode);
    $('#lPrint').onclick = () => window.print();
    $('#lDoc').onclick = async () => {
      busy('Menyusun dokumen Word…');
      try { await loadLib('lib/docx.iife.js'); const b = await VTR.reportDocx(R, mode); download(b, 'Laporan_VT_' + perLabel(pid).replace(/\W+/g, '_') + (mode === 'bi' ? '_ID-EN' : '_' + mode.toUpperCase()) + '.docx'); }
      catch (e) { console.error(e); toast('Gagal membuat .docx: ' + e.message, 8000); } finally { busy(false); }
    };
    $('#lAi').onclick = async () => {
      const src = R.sec[0].p.map(p => p.id + '\n' + p.en).join('\n\n') + '\n\n' + (R.sec[0].list || []).map(x => '• ' + x.id + '\n  ' + x.en).join('\n');
      const t = await AI.tidy(src, 'report'); if (t) { S.settings.aiReport = S.settings.aiReport || {}; S.settings.aiReport[pid] = t; saveSettings(); route(); }
    };
  }
};

/* ------------------------------------------------------------ REFERENSI */
PAGES.matriks = {
  title: 'Matriks tindak lanjut', render(v) {
    v.innerHTML = head('Matriks tindak lanjut', 'Dipanggil otomatis oleh kode fokus pengemudi dan isi email ke atasan.') + '<div class="pad"><div id="t"></div></div>';
    table($('#t'), [{ k: 'code', l: 'Kode', f: r => `<b>${r.code}</b>` }, { k: 'p', l: 'Pemicu' }, { k: 'w', l: 'Aksi pekerja (bahasa lapangan)' }, { k: 'm', l: 'Aksi atasan / safety manager' }, { k: 't', l: 'Tenggat' }],
      Object.keys(VT.REK).map(k => ({ code: k, p: VT.REK[k][0], w: VT.REK[k][1], m: VT.REK[k][2], t: VT.REK[k][3] })), { search: false, free: true, exportName: 'matriks_tindak_lanjut' });
  }
};
PAGES.periode = {
  title: 'Periode 2-mingguan', render(v) {
    if (needData(v)) return;
    v.innerHTML = head('Periode 2-mingguan', 'Label = kolom "Beweekly" data VT. PID = urutan tanggal mulai. Periode dengan data &lt;7 hari tidak otomatis jadi periode laporan terakhir.') + `<div class="pad">
      <div class="form card" style="margin-bottom:12px"><div class="fld"><label for="lastP">Periode laporan terakhir</label><select id="lastP"><option value="">Otomatis (periode terbaru ≥7 hari)</option>${periodOptions(S.inputs.lastPid ? +S.inputs.lastPid : null)}</select><div class="help">Dipakai dashboard, notifikasi, dan rekap.</div></div>
      <div class="fld"><label for="fromP">Hitung skor mulai periode</label><select id="fromP"><option value="">Semua periode tersimpan</option>${periodOptions(S.inputs.scoreFromPid ? +S.inputs.scoreFromPid : null)}</select><div class="help">Periode sebelumnya diabaikan (rantai "berulang" dimulai dari sini). Workbook Excel FY2026 memakai W3–W4 Maret 2026 sebagai awal.</div></div></div>
      <div id="t"></div></div>`;
    $('#lastP').value = S.inputs.lastPid || ''; $('#fromP').value = S.inputs.scoreFromPid || '';
    $('#lastP').onchange = e => { S.inputs.lastPid = e.target.value ? +e.target.value : null; saveInputs(true); route(); };
    $('#fromP').onchange = e => { S.inputs.scoreFromPid = e.target.value ? +e.target.value : null; saveInputs(true); route(); };
    const cnt = {}; S.res.rows.forEach(r => { cnt[r.PID] = (cnt[r.PID] || 0) + 1; });
    table($('#t'), [
      { k: 'pid', l: 'PID', cls: 'c' }, { k: 'lbl', l: 'Periode (label Beweekly)', f: r => `<b>${esc(r.lbl)}</b>` }, { k: 'start', l: 'Mulai', f: r => VT.dtxt(r.start) }, { k: 'end', l: 'Selesai', f: r => VT.dtxt(r.end) },
      { k: 'days', l: 'Hari data', cls: 'c', f: r => (r.days < 7 ? `<span class="up">${r.days}</span>` : r.days) }, { k: 'fy', l: 'FY', f: r => 'FY' + r.fy }, { k: 'drive', l: 'Trip Driving', f: r => f0(r.drive), cls: 'r' }, { k: 'idle', l: 'Trip Idling', f: r => f0(r.idle), cls: 'r' },
      { k: 'n', l: 'Pengemudi dinilai', v: r => cnt[r.pid] || 0, cls: 'r' }, { k: 'x', l: '', nosort: true, f: r => `<button class="btn sm danger" data-del="${esc(r.lbl)}">Hapus</button>` }
    ], S.res.periods, { search: false, exportName: 'periode' });
    $('#t').addEventListener('click', async e => {
      const b = e.target.closest('[data-del]'); if (!b) return;
      if (!(await confirmBox('Hapus semua trip periode "' + b.dataset.del + '" dari perangkat ini?', 'Hapus'))) return;
      await DB.delPeriod(b.dataset.del); setTrips(await DB.periods()); await recompute(); route();
    });
  }
};
PAGES.parameter = {
  title: 'Parameter & standar', render(v) {
    const Pp = P(); const o = S.inputs.params || {};
    const num = (k, l, help, step) => `<div class="fld"><label for="p_${k}">${l}</label><input class="inp" type="number" step="${step || 'any'}" id="p_${k}" data-k="${k}" value="${Pp[k]}">${help ? `<div class="help">${help}</div>` : ''}</div>`;
    const txt = (k, l, help) => `<div class="fld"><label for="p_${k}">${l}</label><input class="inp" type="text" id="p_${k}" data-k="${k}" data-t="1" value="${esc(Pp[k])}">${help ? `<div class="help">${help}</div>` : ''}</div>`;
    v.innerHTML = head('Parameter & standar keberterimaan', 'Kotak oranye boleh diubah QSHE. Nilai bawaan = workbook Excel. Setelah disimpan, semua skor dihitung ulang.') + `<div class="pad">
    <div class="tools"><button class="btn pri" id="pSave">Simpan & hitung ulang</button><button class="btn" id="pReset">Kembalikan ke nilai bawaan</button><span class="sp"></span><span class="note">${Object.keys(o).length ? Object.keys(o).length + ' parameter diubah dari bawaan' : 'Semua parameter = bawaan'}</span></div>
    <h3 class="sec">A. Aturan data</h3><div class="form card">${num('NSTART', 'Jam mulai night driving (0–23)', 'Trip yang MULAI pada jam ini atau lebih = night driving.', 1)}${num('NEND', 'Night driving berlaku sampai sebelum jam', 'Trip mulai 00:00 s.d. sebelum jam ini juga malam. 0 = hanya ≥ jam mulai.', 1)}${num('OSTOL', 'Toleransi overspeed (count per trip, 0–5)', 'Trip dengan count ≤ nilai ini tidak dihitung.', 1)}${num('OSKUN', 'Batas laju overspeed KUNING (per 1.000 km)', 'Lebih dari ini = MERAH.')}${num('OSSUS', 'Count per trip disebut "ngebut beruntun"', 'Trip dengan count ≥ ini → MERAH.', 1)}</div>
    <h3 class="sec">B. Band travel load & kredit beban</h3><div class="card"><table class="t"><thead><tr><th>Band</th><th>Batas bawah km</th><th>Kredit beban (poin)</th></tr></thead><tbody>${VT.LAB.map((b, i) => `<tr><td><b>${b}</b></td><td><input class="cell" type="number" data-arr="BAND_LO" data-i="${i}" value="${Pp.BAND_LO[i]}" ${i === 0 ? 'disabled' : ''}></td><td><input class="cell" type="number" data-arr="LC" data-i="${i}" value="${Pp.LC[i]}"></td></tr>`).join('')}</tbody></table></div>
    <h3 class="sec">C. Rata-rata kelompok (benchmark)</h3><div class="card"><div class="tbox free"><table class="t"><thead><tr><th>Komponen</th>${VT.LAB.map(b => `<th>${b}</th>`).join('')}<th>Satuan</th></tr></thead><tbody>${VT.COMP.map(c => `<tr><td><b>${c.name}</b></td>${VT.LAB.map(b => `<td><input class="cell" type="number" step="0.01" data-bm="${b}" data-c="${c.k}" value="${Pp.BM[b][c.k]}"></td>`).join('')}<td>${c.unit}</td></tr>`).join('')}</tbody></table></div>
      <div class="tools"><button class="btn" id="pCal">Kalibrasi ulang dari data tersimpan…</button><span class="note">Hitung ulang tiap awal FY: total kejadian ÷ total eksposur per band, lantai 50% nasional.</span></div></div>
    <h3 class="sec">D. Bobot & rumus skor</h3><div class="form card">${VT.COMP.map(c => `<div class="fld"><label for="w_${c.k}">Bobot ${c.name}</label><input class="inp" type="number" step="0.01" id="w_${c.k}" data-w="${c.k}" value="${Pp.W[c.k]}"></div>`).join('')}${num('BASE', 'Basis skor komponen', 'skor = 100 × basis ^ rasio')}${num('KKM', 'Km bayangan (kredibilitas)')}${num('KH', 'Jam bayangan (kredibilitas)')}</div>
    <h3 class="sec">E. Pelanggaran berulang</h3><div class="form card">${num('TOL', 'Toleransi "tidak ada perubahan" (0,2 = ±20%)')}${num('CRIT', 'Level kritis (kelipatan rata-rata)')}${num('PTET', 'Pengurang: berulang & tetap')}${num('PBUR', 'Pengurang: berulang & makin buruk')}${num('OSMULT', 'Pengali khusus overspeed')}${num('STREAK', 'Pengali beruntun (≥3 periode)')}${num('REPCAP', 'Batas maksimum pengurang berulang')}</div>
    <h3 class="sec">F. Insiden & pemimpin</h3><div class="card"><table class="t"><thead><tr><th>Tingkat keparahan</th><th>Pengurang pengemudi</th><th>Pengurang pemimpin</th></tr></thead><tbody>${Pp.SEV.map((s, i) => `<tr><td><b>${esc(s[0])}</b></td><td><input class="cell" type="number" data-sev="${i}" data-j="1" value="${s[1]}"></td><td><input class="cell" type="number" data-sev="${i}" data-j="2" value="${s[2]}"></td></tr>`).join('')}</tbody></table>
      <table class="t" style="margin-top:10px"><thead><tr><th>Jabatan pemimpin</th><th>Faktor peran</th><th>Kode di nama DID</th></tr></thead><tbody>${Pp.ROLE.map((r, i) => `<tr><td><b>${esc(r[0])}</b></td><td><input class="cell" type="number" step="0.1" data-role="${i}" data-j="1" value="${r[1]}"></td><td>${r[2] ? `<input class="cell" type="text" data-role="${i}" data-j="2" value="${esc(r[2])}">` : '<span class="muted">isi di master / promosi</span>'}</td></tr>`).join('')}</tbody></table>
      <div class="form" style="margin-top:10px">${num('DECAY', 'Masa berlaku pengurang insiden (hari)')}${num('CAPA', 'Faktor bila CAPA selesai tepat waktu')}${num('LCAP', 'Batas maksimum pengurang pemimpin')}</div></div>
    <h3 class="sec">G. Grade & notifikasi</h3><div class="form card">${['A', 'B', 'C', 'D'].map((g, i) => `<div class="fld"><label for="g_${g}">Skor minimum grade ${g}</label><input class="inp" type="number" id="g_${g}" data-g="${i}" value="${Pp.GRADE[i]}"></div>`).join('')}${num('NOTIFR', 'Batas MERAH komponen (rasio terkoreksi)')}${num('NOTIFMIN', 'Minimal kejadian agar notifikasi dikirim', '', 1)}${num('DEADLINE', 'Tenggat tindak lanjut atasan (hari kerja)', '', 1)}${num('MIN_BRANCH', 'Minimal pengemudi agar cabang diperingkat', '', 1)}
      <div class="fld"><label for="p_CCDRV">CC ke pengemudi?</label><select class="inp" id="p_CCDRV" data-k="CCDRV" data-t="1"><option ${Pp.CCDRV === 'Tidak' ? 'selected' : ''}>Tidak</option><option ${Pp.CCDRV === 'Ya' ? 'selected' : ''}>Ya</option></select></div>${txt('CCQSHE', 'Email QSHE yang selalu di-CC', 'Kosongkan bila tidak perlu.')}${txt('SENDER', 'Nama pengirim di akhir email')}</div>
    <h3 class="sec">H. Laporan bulanan GA</h3><div class="form card">${num('GAP1OS', 'Prioritas 1: overspeed ≥ … kasus sebulan', '', 1)}${num('GAP2', 'Prioritas 2: tidak turun di 1 kategori dengan ≥ … kasus', '', 1)}</div></div>`;
    $('#pSave').onclick = async () => {
      const np = {}; const d = VT.DEFAULTS;
      $$('[data-k]', v).forEach(i => { const k = i.dataset.k; const val = i.dataset.t ? i.value : +i.value; if (JSON.stringify(val) !== JSON.stringify(d[k])) np[k] = val; });
      const arr = k => { const a = d[k].slice(); $$(`[data-arr="${k}"]`, v).forEach(i => { a[+i.dataset.i] = +i.value; }); return a; };
      ['BAND_LO', 'LC'].forEach(k => { const a = arr(k); if (JSON.stringify(a) !== JSON.stringify(d[k])) np[k] = a; });
      const bm = JSON.parse(JSON.stringify(d.BM)); $$('[data-bm]', v).forEach(i => { bm[i.dataset.bm][i.dataset.c] = +i.value; }); if (JSON.stringify(bm) !== JSON.stringify(d.BM)) np.BM = bm;
      const w = Object.assign({}, d.W); $$('[data-w]', v).forEach(i => { w[i.dataset.w] = +i.value; });
      const tw = Object.values(w).reduce((a, b) => a + b, 0); if (Math.abs(tw - 1) > 0.001) { toast('Total bobot harus 100% (sekarang ' + fx(tw * 100, 1) + '%).', 6000); return; }
      if (JSON.stringify(w) !== JSON.stringify(d.W)) np.W = w;
      const sev = JSON.parse(JSON.stringify(d.SEV)); $$('[data-sev]', v).forEach(i => { sev[+i.dataset.sev][+i.dataset.j] = +i.value; }); if (JSON.stringify(sev) !== JSON.stringify(d.SEV)) np.SEV = sev;
      const role = JSON.parse(JSON.stringify(d.ROLE)); $$('[data-role]', v).forEach(i => { role[+i.dataset.role][+i.dataset.j] = +i.dataset.j === 1 ? +i.value : i.value.trim().toUpperCase(); }); if (JSON.stringify(role) !== JSON.stringify(d.ROLE)) np.ROLE = role;
      const g = d.GRADE.slice(); $$('[data-g]', v).forEach(i => { g[+i.dataset.g] = +i.value; }); if (JSON.stringify(g) !== JSON.stringify(d.GRADE)) np.GRADE = g;
      const bl = np.BAND_LO || d.BAND_LO; if (bl.some((x, i) => i && x <= bl[i - 1])) { toast('Batas band harus naik berurutan.', 6000); return; }
      S.inputs.params = np; saveInputs(true); await recompute(); toast('Parameter disimpan. Skor dihitung ulang.'); route();
    };
    $('#pReset').onclick = async () => { if (!(await confirmBox('Kembalikan semua parameter ke nilai bawaan workbook Excel?'))) return; S.inputs.params = {}; saveInputs(true); await recompute(); route(); };
    $('#pCal').onclick = () => {
      if (!S.res) { toast('Unggah data dulu.'); return; }
      const ps = S.res.periods;
      const md = modal('Kalibrasi ulang rata-rata kelompok', `<p class="note">Pilih rentang periode acuan. Hasil ditampilkan dulu; klik Terapkan untuk memakai.</p><div class="form"><div class="fld"><label for="cf">Dari periode</label><select id="cf">${ps.map(p => `<option value="${p.pid}">${esc(p.lbl)}</option>`).join('')}</select></div><div class="fld"><label for="ct">Sampai periode</label><select id="ct">${ps.map(p => `<option value="${p.pid}">${esc(p.lbl)}</option>`).join('')}</select></div></div><div id="calOut" style="margin-top:10px"></div><div class="tools"><span class="sp"></span><button class="btn" id="cRun">Hitung</button><button class="btn pri" id="cApply" disabled>Terapkan</button></div>`, { wide: true });
      $('#cf', md.body).value = ps[Math.min(1, ps.length - 1)].pid; $('#ct', md.body).value = ps.filter(p => p.days >= 7).slice(-1)[0].pid;
      let out = null;
      $('#cRun', md.body).onclick = () => {
        const ag = VT.aggregate(S.res.tripsUsed, S.res.pidOf, P(), new Set());
        out = VT.calibrate(ag.rows, P(), +$('#cf', md.body).value, +$('#ct', md.body).value);
        $('#calOut', md.body).innerHTML = `<p class="note">${f0(out.n)} driver-periode dipakai.</p><div class="tbox free"><table class="t"><thead><tr><th>Komponen</th>${VT.LAB.map(b => `<th>${b} (lama → baru)</th>`).join('')}</tr></thead><tbody>${VT.COMP.map(c => `<tr><td>${c.name}</td>${VT.LAB.map(b => `<td>${f2(P().BM[b][c.k])} → <b>${f2(out.BM[b][c.k])}</b></td>`).join('')}</tr>`).join('')}</tbody></table></div>`;
        $('#cApply', md.body).disabled = false;
      };
      $('#cApply', md.body).onclick = async () => { S.inputs.params = Object.assign({}, S.inputs.params, { BM: out.BM }); saveInputs(true); md.close(); await recompute(); route(); toast('Benchmark baru diterapkan.'); };
    };
  }
};
