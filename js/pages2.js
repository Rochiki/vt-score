/* pages2.js — Bulanan GA, Faktor koreksi, Faktor lain, Master, Analisa, Pengaturan */
'use strict';
const drvList = () => (S.res ? [...new Set(S.res.rows.map(r => r.drv))].sort() : []);
const drvDatalist = id => `<datalist id="${id}">${drvList().map(d => `<option value="${esc(d)}">`).join('')}${(S.inputs.master || []).filter(m => !S.res || !S.res.rows.some(r => r.drv === m.did)).map(m => `<option value="${esc(m.did)}">`).join('')}</datalist>`;
const cabDatalist = id => `<datalist id="${id}">${(S.res ? cabList() : []).map(c => `<option value="${esc(c)}">`).join('')}</datalist>`;

/* ------------------------------------------------------------ LAPORAN BULANAN GA */
PAGES.ga = {
  title: 'Laporan bulanan GA', render(v, q) {
    if (!S.ga.months.length) {
      v.innerHTML = head('Laporan bulanan & concern (data GA)', '') + `<div class="pad"><div class="empty"><b>Belum ada data bulanan GA</b>Unggah lampiran email GA "Data Case … .xlsx" (sheet Overspeed, Fatigue Driving, Night Driving) di menu <a href="#/unggah">Unggah data</a>.</div></div>`; return;
    }
    const month = q.get('m') || S.ga.months[S.ga.months.length - 1];
    const Pp = P(); const master = S.res ? S.res.master : new Map((S.inputs.master || []).map(m => [VT.norm(m.did), m]));
    const A = VT.gaAnalysis(S.ga, month, Pp, master, S.ga.log);
    const info = S.ga.info[month] || {}; const pinfo = S.ga.info[A.prevMonth] || {};
    const kend = +info.kend || 0, kct = +info.kct || 0, peng = +info.peng || 0;
    const ml = m => VT.monthLabel(m);
    v.innerHTML = head('Laporan bulanan VT (data GA) – sebelum vs sesudah & concern', 'Semua angka otomatis dari lampiran "Data Case" GA + angka infografis yang Anda isi. Bulan dianalisa dibanding bulan sebelumnya.') + `<div class="pad">
      <div class="tools"><label class="small" for="gm"><b>Bulan dianalisa</b></label><select id="gm">${S.ga.months.slice().reverse().map(m => `<option value="${m}" ${m === month ? 'selected' : ''}>${ml(m)}</option>`).join('')}</select><span class="muted small">pembanding: ${ml(A.prevMonth)}</span><span class="sp"></span><button class="btn" id="gInfo">Isi angka infografis GA</button><button class="btn" id="gXls">Ekspor Excel</button></div>
      <div class="kpis">
        <div class="kpi"><span class="l">PENGGUNA TERCATAT ≥1 PELANGGARAN</span><span class="v">${peng ? pc(A.orgCur / peng) : 'n/a'}</span><span class="s">${A.orgCur} dari ${peng || '? (isi infografis)'} pengguna</span></div>
        <div class="kpi acc"><span class="l">KASUS DARI PENGEMUDI BERULANG</span><span class="v">${pc(A.repAll)}</span><span class="s">tercatat di kategori yang sama bulan lalu</span></div>
        <div class="kpi"><span class="l">PRIORITAS 1 (DIPANGGIL MINGGU INI)</span><span class="v">${A.p1}</span><span class="s">Prioritas 2: ${A.p2} orang</span></div>
        <div class="kpi"><span class="l">MEMBAIK (BERSIH BULAN INI)</span><span class="v">${A.mb}</span><span class="s">tercatat bulan lalu, kini nol kasus</span></div>
        <div class="kpi"><span class="l">OVERSPEED DARI 1 ORANG</span><span class="v">${pc(A.cat[0].cur ? (A.topOs ? A.topOs.os_cur : 0) / A.cat[0].cur : 0)}</span><span class="s">${A.topOs ? A.topOs.os_cur + ' dari ' + A.cat[0].cur + ' kasus · ' + esc(A.topOs.nama) : '–'}</span></div>
        <div class="kpi"><span class="l">NIGHT DRIVING: PENGGUNA TERCATAT</span><span class="v">${peng ? pc(A.orgNd / peng) : 'n/a'}</span><span class="s">${pc(A.cat[2].cur ? A.ndProj / A.cat[2].cur : 0)} kasus dari project / dept</span></div>
        <div class="kpi"><span class="l">KENDARAAN TANPA CARTRACK</span><span class="v">${kend ? f0(kend - kct) : 'n/a'}</span><span class="s">${kend ? pc((kend - kct) / kend) + ' dari ' + f0(kend) + ' unit' : 'isi angka infografis'}</span></div>
        <div class="kpi"><span class="l">KASUS ATAS ID UMUM MASTER/DRIVER</span><span class="v">${A.gen}</span><span class="s">termasuk ${A.genOs} overspeed · tak bisa ditelusuri</span></div></div>
      <h3 class="sec">Per kategori – bulan lalu vs bulan ini</h3><div id="gCat"></div>
      <div class="grid2"><div class="panel"><h4>Kasus bulan lalu vs bulan ini</h4><div class="chartbox sm"><canvas id="cGa"></canvas></div></div><div class="panel"><h4>Kasus bulan ini datang dari siapa</h4><div class="sub">Berulang tidak turun · berulang sudah turun · baru</div><div class="chartbox sm"><canvas id="cGa2"></canvas></div></div></div>
      <h3 class="sec">Sebelum vs sesudah – cara baca lama vs yang sebenarnya ditunjukkan data</h3><div id="gBA"></div>
      <h3 class="sec">Daftar prioritas tindak lanjut</h3><p class="note">P1: overspeed ≥ ${Pp.GAP1OS} kasus sebulan, atau tidak turun di dua kategori, atau tercatat di ketiga kategori. P2: tidak turun di satu kategori dengan ≥ ${Pp.GAP2} kasus. Tidak turun = tercatat bulan lalu dan bulan ini dengan kasus sama/lebih banyak.</p><div id="gPr"></div>
      <h3 class="sec">Per unit (cabang / project / dept)</h3><div id="gUn"></div>
      <h3 class="sec">Concern & kesepakatan dengan GA / Tim VT</h3><div id="gCon"></div>
      <h3 class="sec">Prompt flyer (ChatGPT / Gemini / Claude)</h3><div class="card"><pre id="gPrompt" style="white-space:pre-wrap;margin:0;font-size:12.5px"></pre><div class="tools"><button class="btn pri" id="gCopy">Salin prompt</button></div></div></div>`;
    $('#gm').onchange = e => { location.hash = '#/ga?m=' + e.target.value; };
    const chg = c => (c.chg == null ? '–' : `<span class="${c.chg > 0 ? 'up' : c.chg < 0 ? 'dn' : ''}">${sg(c.chg * 100, 0)}%</span>`);
    const infoTot = { Overspeed: info.os, Fatigue: info.ft, 'Night driving': info.nd };
    $('#gCat').innerHTML = `<div class="tbox free"><table class="t"><thead><tr><th>Kategori</th><th class="r">Bulan lalu</th><th class="r">Bulan ini</th><th class="r">Perubahan</th><th class="r">Orang tercatat</th><th class="r">Dari pengemudi berulang</th><th class="r">Berulang, tidak turun</th><th class="r">Berulang, sudah turun</th><th class="r">Baru</th><th class="r">Total infografis GA</th><th>Cek</th></tr></thead><tbody>${A.cat.map(c => { const it = infoTot[c.kat]; const ok = it == null || it === '' ? 'belum diisi' : +it === c.cur ? 'cocok' : 'beda – tanyakan ke GA'; return `<tr><td><b>${c.kat}</b></td><td class="r">${f0(c.prv)}</td><td class="r">${f0(c.cur)}</td><td class="r">${chg(c)}</td><td class="r">${c.org}</td><td class="r"><b>${pc(c.rep)}</b></td><td class="r">${c.tt}</td><td class="r">${c.tu}</td><td class="r">${c.bar}</td><td class="r">${it == null || it === '' ? '–' : f0(+it)}</td><td>${ok.startsWith('beda') ? `<span class="st MERAH">${ok}</span>` : ok === 'cocok' ? `<span class="st HIJAU">${ok}</span>` : ok}</td></tr>`; }).join('')}</tbody></table></div>`;
    chart('cGa', { type: 'bar', data: { labels: A.cat.map(c => c.kat), datasets: [{ label: ml(A.prevMonth), data: A.cat.map(c => c.prv), backgroundColor: '#bdbdbd', borderRadius: 3 }, { label: ml(month), data: A.cat.map(c => c.cur), backgroundColor: '#f7941d', borderRadius: 3 }] }, options: { plugins: { legend: { position: 'bottom' } }, scales: { y: Object.assign({ beginAtZero: true }, gridY), x: { grid: { display: false } } } } });
    chart('cGa2', { type: 'bar', data: { labels: A.cat.map(c => c.kat), datasets: [{ label: 'Berulang, tidak turun', data: A.cat.map(c => c.tt), backgroundColor: '#c00000' }, { label: 'Berulang, sudah turun', data: A.cat.map(c => c.tu), backgroundColor: '#f7941d' }, { label: 'Baru', data: A.cat.map(c => c.bar), backgroundColor: '#cfcfcf' }] }, options: { indexAxis: 'y', plugins: { legend: { position: 'bottom' } }, scales: { x: Object.assign({ stacked: true, beginAtZero: true }, gridY), y: { stacked: true, grid: { display: false } } } } });
    // sebelum vs sesudah
    const c0 = A.cat[0], c1 = A.cat[1], c2 = A.cat[2];
    const nt = (a, b) => (a >= b ? 'naik' : 'turun');
    const BA = [
      ['Overspeed', `Overspeed ${nt(c0.cur, c0.prv)} dari ${c0.prv} menjadi ${c0.cur} kasus, ditampilkan dengan top 3 area.`, A.topOs && A.topOs.os_cur ? `Satu orang (${A.topOs.nama}, ${A.topOs.unit}) menyumbang ${A.topOs.os_cur} kasus (${pc(c0.cur ? A.topOs.os_cur / c0.cur : 0)}). Tanpa dia: ${c0.cur - A.topOs.os_cur} kasus dari ${Math.max(0, A.orgOs - 1)} orang.` : 'Tidak ada overspeed bulan ini.', 'Coaching per orang dan cek perangkat / setelan batas kecepatan untuk kasus ekstrem. Bukan himbauan umum ke semua cabang.'],
      ['Peringkat area', 'Top 3 area diurutkan dari jumlah kasus.', A.topUnitOs ? `Unit overspeed tertinggi (${A.topUnitOs.unit}, ${A.topUnitOs.os_cur} kasus) hanya berisi ${A.topUnitOsOrg} orang. Jumlah kasus tanpa km dan jam mengemudi tidak adil untuk membandingkan cabang.` : '–', 'Bandingkan laju per 1.000 km / per 100 jam (Rekap cabang). Minta GA menyertakan km dan jam per pengemudi.'],
      ['Night driving', `Night driving ${nt(c2.cur, c2.prv)} sekitar ${c2.prv ? f0(Math.abs(c2.cur / c2.prv - 1) * 100) : '–'}%.`, `${peng ? pc(A.orgNd / peng) : A.orgNd + ' orang'} pengguna tercatat night driving. ${pc(c2.cur ? A.ndProj / c2.cur : 0)} kasus dari project / dept; tertinggi ${A.topUnitNd ? A.topUnitNd.unit : '–'}. ${pc(c2.rep)} kasus dari pengemudi berulang.`, 'Pisahkan perjalanan malam yang disetujui lewat JRM dan yang tidak. Yang tanpa JRM → atasan.'],
      ['Fatigue', `Fatigue ${nt(c1.cur, c1.prv)} sekitar ${c1.prv ? f0(Math.abs(c1.cur / c1.prv - 1) * 100) : '–'}%.`, `${pc(c1.rep)} kasus dari orang yang sama dengan bulan lalu.${info.ftjam ? ' Rekor ' + info.ftjam + ' jam / ' + f0(+info.ftkm || 0) + ' km tanpa berhenti, tanpa nama pengemudi.' : ''}`, 'Kasus terberat wajib punya nama, atasan, dan tindakan. Cek jadwal & rencana perjalanan, wajib istirahat.'],
      ['Required action', 'Pengawasan ditingkatkan; QSHE & OPSYS diminta analisa lebih lanjut.', `${A.p1} orang Prioritas 1 dan ${A.p2} orang Prioritas 2 (nama, unit, atasan di bawah). ${A.mb} orang yang tercatat bulan lalu sudah bersih.${A.tlN ? ' Tindak lanjut bulan lalu: ' + A.tlOk + ' dari ' + A.tlN + ' selesai.' : ' Tindak lanjut bulan lalu belum tercatat.'}`, 'Setiap tindakan punya PIC atasan, tenggat, dan status di Log tindak lanjut. Yang membaik diapresiasi.'],
      ['Cakupan & ID umum', 'Total kendaraan, GPS, Cartrack, dan pengguna ditampilkan sebagai angka saja.', `${kend ? f0(kend - kct) + ' kendaraan (' + pc((kend - kct) / kend) + ') tidak terpantau Cartrack. ' : ''}${A.gen} kasus tercatat atas ID umum MASTER/DRIVER (termasuk ${A.genOs} overspeed). ${A.lead} pemimpin berkode FR/CSS ikut tercatat.`, 'Target bulanan: nol ID umum (tap DID pribadi) dan jadwal pemasangan Cartrack untuk unit sisa.'],
      ['Definisi', `Night driving dihitung ${info.gaNight || '19.00–07.00'}. Batas overspeed & fatigue tidak dituliskan.`, `Skor VT 2-mingguan memakai night driving ${String(Pp.NSTART).padStart(2, '0')}.00–${String((Pp.NEND + 23) % 24).padStart(2, '0')}.59. Batas overspeed GA: ${info.oslim ? info.oslim + ' km/jam' : 'belum ditetapkan'}. Batas fatigue GA: ${info.ftlim ? info.ftlim + ' jam' : 'belum ditetapkan'}.`, 'Sepakati satu definisi tertulis GA – Tim VT – QSHE agar angka bulan ke bulan bisa dibandingkan.']];
    $('#gBA').innerHTML = `<div class="tbox free"><table class="t"><thead><tr><th>Topik</th><th>Laporan GA (sebelum)</th><th>Yang ditunjukkan data (sesudah)</th><th>Tindakan yang tepat</th></tr></thead><tbody>${BA.map(r => `<tr><td><b>${r[0]}</b></td><td class="muted">${esc(r[1])}</td><td style="border-left:3px solid var(--or)">${esc(r[2])}</td><td style="color:var(--ord);font-weight:600">${esc(r[3])}</td></tr>`).join('')}</tbody></table></div>`;
    let fSel = 'P1P2';
    const redrawPr = () => {
      const f = fSel; const rows = A.list.filter(r => r.pr !== 'Membaik' && (!f || (f === 'P1P2' ? r.pr === 'P1' || r.pr === 'P2' : r.pr === f)));
      table($('#gPr'), [
        { k: 'pr', l: 'Prioritas', f: r => `<span class="st ${r.pr}">${r.pr}</span>` }, { k: 'nama', l: 'Nama / DID', f: r => `<b>${esc(r.nama)}</b>` }, { k: 'unit', l: 'Unit' },
        { k: 'os', l: 'OS lalu → ini', v: r => r.os_cur, f: r => r.os_prv + ' → ' + `<span class="${r.os_cur > r.os_prv ? 'up' : ''}">${r.os_cur}</span>`, x: r => r.os_prv + ' → ' + r.os_cur },
        { k: 'ft', l: 'Fatigue lalu → ini', v: r => r.ft_cur, f: r => r.ft_prv + ' → ' + `<span class="${r.ft_prv && r.ft_cur >= r.ft_prv ? 'up' : ''}">${r.ft_cur}</span>`, x: r => r.ft_prv + ' → ' + r.ft_cur },
        { k: 'nd', l: 'Malam lalu → ini', v: r => r.nd_cur, f: r => r.nd_prv + ' → ' + `<span class="${r.nd_prv && r.nd_cur >= r.nd_prv ? 'up' : ''}">${r.nd_cur}</span>`, x: r => r.nd_prv + ' → ' + r.nd_cur },
        { k: 'al', l: 'Alasan' }, { k: 'atasan', l: 'Atasan langsung' }, { k: 'tlPrev', l: 'TL bulan lalu' }, { k: 'tlCur', l: 'TL bulan ini' },
        { k: 'b', l: '', nosort: true, f: r => `<button class="btn sm" data-log="${esc(r.nama)}">Catat TL</button>` }
      ], rows, { exportName: 'prioritas_GA_' + month, tools: '<select id="gF" aria-label="Filter prioritas"><option value="P1P2">Prioritas 1 dan 2</option><option value="">Semua yang tercatat</option><option value="P1">P1 saja</option><option value="P2">P2 saja</option></select>' });
      $('#gF').value = fSel; $('#gF').onchange = e => { fSel = e.target.value; redrawPr(); };
    };
    redrawPr();
    $('#gPr').addEventListener('click', e => { const b = e.target.closest('[data-log]'); if (!b) return; const r = A.list.find(x => x.nama === b.dataset.log); logForm({ bulan: month, nama: r.nama, prioritas: r.pr === 'Membaik' ? '' : r.pr, pic: r.atasan.startsWith('(') ? '' : r.atasan }); });
    table($('#gUn'), [{ k: 'unit', l: 'Unit', f: r => `<b>${esc(r.unit)}</b>` }, { k: 'tipe', l: 'Tipe' }, { k: 'os_prv', l: 'OS lalu', cls: 'r' }, { k: 'os_cur', l: 'OS ini', cls: 'r', f: r => `<span class="${r.os_cur > r.os_prv ? 'up' : r.os_cur < r.os_prv ? 'dn' : ''}">${r.os_cur}</span>` },
      { k: 'ft_prv', l: 'Fatigue lalu', cls: 'r' }, { k: 'ft_cur', l: 'Fatigue ini', cls: 'r', f: r => `<span class="${r.ft_cur > r.ft_prv ? 'up' : r.ft_cur < r.ft_prv ? 'dn' : ''}">${r.ft_cur}</span>` }, { k: 'nd_prv', l: 'Malam lalu', cls: 'r' },
      { k: 'nd_cur', l: 'Malam ini', cls: 'r', f: r => `<span class="${r.nd_cur > r.nd_prv ? 'up' : r.nd_cur < r.nd_prv ? 'dn' : ''}">${r.nd_cur}</span>` }, { k: 'tot_prv', l: 'Total lalu', cls: 'r' }, { k: 'tot_cur', l: 'Total ini', cls: 'r', f: r => `<b>${r.tot_cur}</b>` },
      { k: 'sel', l: 'Selisih', cls: 'r', f: r => `<span class="${r.sel > 0 ? 'up' : r.sel < 0 ? 'dn' : ''}">${sg(r.sel, 0)}</span>` }, { k: 'org', l: 'Orang tercatat', cls: 'c' }, { k: 'p1', l: 'Orang P1', cls: 'c' }], A.units, { exportName: 'unit_GA_' + month });
    // concerns
    const CON = [['Km & jam per pengemudi', 'Tanpa paparan, cabang dengan perjalanan jauh selalu terlihat paling buruk.', 'GA melampirkan km dan jam mengemudi per nama tiap bulan.', 'GA / Tim VT'],
      ['Definisi satu "kasus"', 'Satu kasus = satu trip, event, atau count? Menentukan arti angka besar.', 'Tulis definisi per kategori di laporan; samakan dengan rekap 2-mingguan.', 'Tim VT'],
      ['Batas overspeed & fatigue', 'Tanpa batas tertulis, perubahan setelan bisa terbaca sebagai "kenaikan".', 'Isi batas di angka infografis; beri catatan bila setelan berubah.', 'GA / QSHE'],
      ['Jam night driving', 'GA memakai 19.00–07.00, skor VT memakai 20.00–04.59.', 'Sepakati satu definisi untuk semua laporan.', 'QSHE'],
      ['ID umum MASTER/DRIVER', 'Kasus tidak punya penanggung jawab dan tidak masuk skor personal.', 'Daftar ID umum per cabang + target nol bulan depan.', 'GA + Branch Head'],
      ['Kendaraan belum Cartrack', 'Pelanggaran di unit ini tidak terlihat sama sekali.', 'Daftar unit + jadwal pasang / alasan.', 'GA'],
      ['Nama sama di beberapa unit', 'Pengemudi mobile tercatat di 2 cabang; tanggung jawab atasan kabur.', 'Nama DID sama persis dengan master; unit = homebase.', 'Tim VT / HR'],
      ['Status tindak lanjut bulan lalu', 'Tanpa close-loop, laporan berulang tanpa perubahan.', 'Atasan mengisi Log tindak lanjut; laporan berikut menampilkan statusnya.', 'Atasan + QSHE'],
      ['Kalimat laporan bertentangan', 'Contoh: rute "Jalan Tol" vs "Jalan Provinsi"; fatigue "after 7 PM" padahal mayoritas siang.', 'Cek silang teks highlight dengan data sebelum dikirim.', 'GA']];
    const cs = S.ga.concerns || {};
    $('#gCon').innerHTML = `<div class="tbox free"><table class="t"><thead><tr><th>Concern</th><th>Kenapa penting</th><th>Usulan</th><th>PIC</th><th>Target</th><th>Status</th></tr></thead><tbody>${CON.map((c, i) => { const s = cs[i] || {}; return `<tr><td><b>${esc(c[0])}</b></td><td class="muted">${esc(c[1])}</td><td>${esc(c[2])}</td><td><input class="cell" data-cc="${i}" data-f="pic" value="${esc(s.pic || c[3])}"></td><td><input class="cell" type="date" data-cc="${i}" data-f="tgl" value="${esc(s.tgl || '')}"></td><td><select class="cell" data-cc="${i}" data-f="st">${['Belum', 'Proses', 'Sepakat', 'Selesai'].map(o => `<option ${(s.st || 'Belum') === o ? 'selected' : ''}>${o}</option>`).join('')}</select></td></tr>`; }).join('')}</tbody></table></div>`;
    $$('[data-cc]', v).forEach(i => i.onchange = () => { S.ga.concerns = S.ga.concerns || {}; const o = S.ga.concerns[i.dataset.cc] = S.ga.concerns[i.dataset.cc] || {}; o[i.dataset.f] = i.value; saveGA(); });
    const prompt = `Buat infografis vertikal 4:5 (1080x1350 px) berjudul "VT DRIVING BEHAVIOR · ${ml(month).toUpperCase()}" untuk PT Hexindo Adiperkasa Tbk. Gaya korporat industri alat berat: latar putih, panel hitam, aksen oranye #F7941D hanya untuk angka penting, huruf sans-serif tebal kondensed. Ilustrasi pickup double cabin putih generik dengan stripe oranye, tanpa merek.
Isi dari atas:
(1) Angka utama besar "${pc(A.repAll)}" dengan teks "kasus datang dari pengemudi yang sama dengan bulan lalu".
(2) Tiga kartu: OVERSPEED ${c0.prv} → ${c0.cur} kasus ("${pc(c0.cur ? (A.topOs ? A.topOs.os_cur : 0) / c0.cur : 0)} dari 1 orang"); NIGHT DRIVING ${c2.prv} → ${c2.cur} kasus ("${peng ? pc(A.orgNd / peng) : A.orgNd + ' orang'} pengguna tercatat"); FATIGUE ${c1.prv} → ${c1.cur} kasus.
(3) Pita status: Prioritas 1: ${A.p1} orang · Prioritas 2: ${A.p2} orang · Membaik: ${A.mb} orang.
(4) Kotak titik buta: ${kend ? (kend - kct) + ' kendaraan belum Cartrack · ' : ''}${A.gen} kasus atas ID umum.
(5) Footer hitam dengan 3 aksi: "Atasan panggil Prioritas 1 minggu ini", "Perjalanan malam wajib JRM", "Tap DID tag pribadi". Baris kecil: "Sumber: VT Cartrack, diolah QSHE & GA".
Aturan: teks bahasa Indonesia persis seperti ini, angka jangan diubah, tanpa wajah orang nyata, kontras tinggi untuk layar HP.`;
    $('#gPrompt').textContent = prompt; $('#gCopy').onclick = () => copyText(prompt);
    $('#gInfo').onclick = () => {
      const f = [['kend', 'Total kendaraan'], ['gps', 'Kendaraan ber-GPS'], ['kct', 'Kendaraan Cartrack'], ['peng', 'Pengguna (pekerja)'], ['os', 'Overspeed (infografis)'], ['ft', 'Fatigue (infografis)'], ['nd', 'Night driving (infografis)'], ['vmax', 'Kecepatan tertinggi (km/jam)'], ['osdur', 'Rata-rata durasi overspeed (detik)'], ['ftjam', 'Mengemudi terlama tanpa henti (jam)'], ['ftkm', 'Jarak terlama tanpa henti (km)'], ['oslim', 'Batas kecepatan overspeed GA (km/jam)'], ['ftlim', 'Batas fatigue GA (jam)'], ['gaNight', 'Jam malam versi GA (teks)']];
      const md = modal('Angka infografis GA · ' + ml(month), `<div class="form">${f.map(([k, l]) => `<div class="fld"><label for="gi_${k}">${l}</label><input class="inp" id="gi_${k}" data-k="${k}" ${k === 'gaNight' ? 'type="text" placeholder="19.00–07.00"' : 'type="number" step="any"'} value="${esc(info[k] == null ? '' : info[k])}"></div>`).join('')}</div><div class="tools"><span class="sp"></span><button class="btn pri" id="giSave">Simpan</button></div>`, { wide: true });
      $('#giSave', md.body).onclick = () => { const o = {}; $$('[data-k]', md.body).forEach(i => { if (i.value !== '') o[i.dataset.k] = i.type === 'number' ? +i.value : i.value; }); S.ga.info[month] = o; saveGA(); md.close(); route(); };
    };
    $('#gXls').onclick = () => exportXlsx('Laporan_bulanan_GA_' + month, [
      { name: 'Per kategori', cols: [{ l: 'Kategori', k: 'kat' }, { l: 'Bulan lalu', k: 'prv' }, { l: 'Bulan ini', k: 'cur' }, { l: 'Perubahan', x: c => c.chg, nf: '0%' }, { l: 'Orang tercatat', k: 'org' }, { l: 'Dari pengemudi berulang', k: 'rep', nf: '0%' }, { l: 'Berulang tidak turun', k: 'tt' }, { l: 'Berulang sudah turun', k: 'tu' }, { l: 'Baru', k: 'bar' }], rows: A.cat },
      { name: 'Sebelum vs sesudah', cols: [{ l: 'Topik', x: r => r[0], w: 18 }, { l: 'Laporan GA (sebelum)', x: r => r[1], w: 50 }, { l: 'Yang ditunjukkan data (sesudah)', x: r => r[2], w: 70 }, { l: 'Tindakan', x: r => r[3], w: 50 }], rows: BA },
      { name: 'Prioritas', cols: [{ l: 'Prioritas', k: 'pr' }, { l: 'Nama', k: 'nama', w: 30 }, { l: 'Unit', k: 'unit', w: 24 }, { l: 'OS lalu', k: 'os_prv' }, { l: 'OS ini', k: 'os_cur' }, { l: 'Fatigue lalu', k: 'ft_prv' }, { l: 'Fatigue ini', k: 'ft_cur' }, { l: 'Malam lalu', k: 'nd_prv' }, { l: 'Malam ini', k: 'nd_cur' }, { l: 'Alasan', k: 'al', w: 40 }, { l: 'Atasan langsung', k: 'atasan', w: 26 }, { l: 'TL bulan lalu', k: 'tlPrev' }, { l: 'TL bulan ini', k: 'tlCur' }], rows: A.list },
      { name: 'Per unit', cols: [{ l: 'Unit', k: 'unit', w: 26 }, { l: 'Tipe', k: 'tipe' }, { l: 'OS lalu', k: 'os_prv' }, { l: 'OS ini', k: 'os_cur' }, { l: 'Fatigue lalu', k: 'ft_prv' }, { l: 'Fatigue ini', k: 'ft_cur' }, { l: 'Malam lalu', k: 'nd_prv' }, { l: 'Malam ini', k: 'nd_cur' }, { l: 'Total lalu', k: 'tot_prv' }, { l: 'Total ini', k: 'tot_cur' }, { l: 'Selisih', k: 'sel' }, { l: 'Orang tercatat', k: 'org' }, { l: 'Orang P1', k: 'p1' }], rows: A.units }]);
  }
};
function logForm(pre, idx) {
  const o = Object.assign({ bulan: S.ga.months.slice(-1)[0] || '', nama: '', prioritas: '', pic: '', tanggal: today(), tindakan: '', hasil: '', status: 'Proses' }, pre || {});
  const md = modal(idx != null ? 'Ubah tindak lanjut' : 'Catat tindak lanjut', `<div class="form">
    <div class="fld"><label for="lgB">Bulan laporan</label><select class="inp" id="lgB">${S.ga.months.slice().reverse().map(m => `<option value="${m}" ${m === o.bulan ? 'selected' : ''}>${VT.monthLabel(m)}</option>`).join('')}</select></div>
    <div class="fld"><label for="lgN">Nama / DID</label><input class="inp" id="lgN" list="dlN" value="${esc(o.nama)}">${drvDatalist('dlN')}</div>
    <div class="fld"><label for="lgP">Prioritas saat itu</label><select class="inp" id="lgP">${['', 'P1', 'P2', 'P3'].map(p => `<option ${o.prioritas === p ? 'selected' : ''}>${p}</option>`).join('')}</select></div>
    <div class="fld"><label for="lgC">Atasan / PIC</label><input class="inp" id="lgC" value="${esc(o.pic)}"></div>
    <div class="fld"><label for="lgT">Tanggal tindakan</label><input class="inp" type="date" id="lgT" value="${esc(o.tanggal)}"></div>
    <div class="fld"><label for="lgS">Status</label><select class="inp" id="lgS">${['Selesai', 'Proses', 'Belum'].map(p => `<option ${o.status === p ? 'selected' : ''}>${p}</option>`).join('')}</select></div>
    <div class="fld full"><label for="lgA">Tindakan (coaching, teguran, ride-along, cek unit, jadwal ulang)</label><textarea class="inp" id="lgA">${esc(o.tindakan)}</textarea></div>
    <div class="fld full"><label for="lgH">Hasil / kesepakatan</label><textarea class="inp" id="lgH">${esc(o.hasil)}</textarea></div></div>
    <div class="tools"><span class="sp"></span><button class="btn pri" id="lgSave">Simpan</button></div>`, { wide: true });
  $('#lgSave', md.body).onclick = () => {
    const n = { id: o.id || uid(), bulan: $('#lgB', md.body).value, nama: VT.norm($('#lgN', md.body).value), prioritas: $('#lgP', md.body).value, pic: $('#lgC', md.body).value.trim(), tanggal: $('#lgT', md.body).value, tindakan: $('#lgA', md.body).value.trim(), hasil: $('#lgH', md.body).value.trim(), status: $('#lgS', md.body).value };
    if (!n.nama) { toast('Isi nama.'); return; }
    S.ga.log = S.ga.log || [];
    const i = S.ga.log.findIndex(x => x.id === n.id); if (i >= 0) S.ga.log[i] = n; else S.ga.log.push(n);
    saveGA(); md.close(); toast('Tindak lanjut tersimpan.'); route();
  };
}
PAGES['ga-log'] = {
  title: 'Log tindak lanjut', render(v) {
    v.innerHTML = head('Log tindak lanjut prioritas bulanan', 'Atasan/QSHE mencatat tindakan untuk pengemudi Prioritas 1–2. Status muncul otomatis di laporan bulan berikutnya ("TL bulan lalu").') + `<div class="pad"><div class="tools"><button class="btn pri" id="lgAdd" ${S.ga.months.length ? '' : 'disabled'}>+ Catat tindak lanjut</button></div><div id="t"></div></div>`;
    $('#lgAdd').onclick = () => logForm();
    const rows = (S.ga.log || []).slice().sort((a, b) => (a.bulan < b.bulan ? 1 : -1));
    table($('#t'), [{ k: 'bulan', l: 'Bulan', f: r => VT.monthLabel(r.bulan) }, { k: 'nama', l: 'Nama / DID', f: r => `<b>${esc(r.nama)}</b>` }, { k: 'prioritas', l: 'Prioritas', f: r => (r.prioritas ? `<span class="st ${r.prioritas}">${r.prioritas}</span>` : '') }, { k: 'pic', l: 'PIC' }, { k: 'tanggal', l: 'Tanggal' }, { k: 'tindakan', l: 'Tindakan' }, { k: 'hasil', l: 'Hasil' }, { k: 'status', l: 'Status', f: r => `<span class="st ${r.status === 'Selesai' ? 'HIJAU' : r.status === 'Proses' ? 'KUNING' : 'MERAH'}">${r.status}</span>` },
      { k: 'x', l: '', nosort: true, f: r => `<button class="btn sm" data-e="${r.id}">Ubah</button> <button class="btn sm danger" data-d="${r.id}">Hapus</button>` }], rows, { exportName: 'log_tindak_lanjut' });
    $('#t').addEventListener('click', async e => {
      const ed = e.target.closest('[data-e]'), dl = e.target.closest('[data-d]');
      if (ed) logForm(S.ga.log.find(x => x.id === ed.dataset.e), 1);
      if (dl && await confirmBox('Hapus catatan ini?', 'Hapus')) { S.ga.log = S.ga.log.filter(x => x.id !== dl.dataset.d); saveGA(); route(); }
    });
  }
};

/* ------------------------------------------------------------ FAKTOR KOREKSI */
PAGES.koreksi = {
  title: 'Faktor koreksi', render(v, q) {
    if (needData(v)) return;
    v.innerHTML = head('Faktor koreksi', 'Untuk kasus yang sah: DID dipinjam/tertukar, error GPS, darurat, pengawalan, perjalanan malam ber-JRM, test drive mekanik, unit pihak ketiga, data ganda. Hanya status <b>Disetujui</b> (QSHE) yang mengubah skor.') + `<div class="pad">
      <div class="tools"><button class="btn pri" id="kAdd">+ Ajukan faktor koreksi</button><span class="sp"></span><span class="note">Setiap koreksi tercatat (audit trail) dan dilaporkan di Laporan bilingual.</span></div>
      <div id="kList"></div><h3 class="sec">Saran koreksi otomatis (perlu dicek Tim VT)</h3><p class="note">Ditemukan dari pola data: trip dengan kecepatan rata-rata &gt;120 km/jam (indikasi lonjakan GPS), DID dipakai saat tercatat cuti, atau setelah tanggal keluar.</p><div id="kSug"></div></div>`;
    $('#kAdd').onclick = () => corrForm(q.get('d') ? { nama: q.get('d'), pid: q.get('pid') ? +q.get('pid') : '' } : null);
    const rows = S.inputs.corrections || [];
    table($('#kList'), [
      { k: 'status', l: 'Status', f: r => `<span class="st ${r.status}">${r.status}</span>` }, { k: 'nama', l: 'Nama DID', f: r => `<b>${esc(r.nama)}</b>` },
      { k: 'alasan', l: 'Alasan', v: r => (VT.CORR_REASONS.find(x => x.id === r.alasan) || {}).label || r.alasan },
      { k: 'aksi', l: 'Aksi', v: r => VT.CORR_ACTIONS[r.aksi] + (r.aksi === 'annul' ? ' (' + (r.komponen || []).map(c => VT.COMP.find(z => z.k === c).name).join(', ') + ')' : r.aksi === 'reassign' ? ' → ' + r.pemakai : '') },
      { k: 'lingkup', l: 'Lingkup', v: r => corrScope(r) }, { k: 'efek', l: 'Efek (trip)', v: r => (S.res.corAudit[r.id] ? S.res.corAudit[r.id].trips : r.status === 'Disetujui' ? 0 : '—'), cls: 'c' },
      { k: 'bukti', l: 'No. bukti' }, { k: 'oleh', l: 'Diajukan oleh' }, { k: 'approver', l: 'Disetujui oleh' },
      { k: 'x', l: '', nosort: true, f: r => `${r.status !== 'Disetujui' ? `<button class="btn sm pri" data-ok="${r.id}">Setujui</button> ` : ''}${r.status !== 'Ditolak' ? `<button class="btn sm" data-no="${r.id}">Tolak</button> ` : ''}<button class="btn sm" data-e="${r.id}">Ubah</button> <button class="btn sm danger" data-d="${r.id}">Hapus</button>` }
    ], rows.slice().reverse(), { exportName: 'faktor_koreksi' });
    $('#kList').addEventListener('click', async e => {
      const g = a => e.target.closest('[' + a + ']'); const id = (g('data-ok') || g('data-no') || g('data-e') || g('data-d') || {}).dataset;
      if (!id) return;
      const c = rows.find(x => x.id === (id.ok || id.no || id.e || id.d)); if (!c) return;
      if (id.e) return corrForm(c);
      if (id.d) { if (!(await confirmBox('Hapus faktor koreksi ini?', 'Hapus'))) return; S.inputs.corrections = rows.filter(x => x !== c); }
      if (id.ok) { if (!c.bukti) { toast('Isi nomor bukti dulu (tombol Ubah) sebelum menyetujui.', 6000); return; } c.status = 'Disetujui'; c.approver = c.approver || 'QSHE'; c.tglSetuju = today(); }
      if (id.no) { c.status = 'Ditolak'; }
      saveInputs(true); await recompute(); route();
    });
    // saran
    const sug = [];
    const F = S.inputs.factors;
    const leaves = (F.cuti || []).map(x => ({ nm: VT.norm(x.nama), a: VT.isoToDay(x.dari), b: VT.isoToDay(x.sampai), t: 'cuti' }));
    const outs = (F.keluar || []).map(x => ({ nm: VT.norm(x.nama), a: VT.isoToDay(x.tanggal) }));
    for (const t of S.trips) {
      const day = Math.floor(t.t0 + 1e-9);
      if (t.h > 0 && t.km / t.h > 120) sug.push({ t, why: 'Kecepatan rata-rata ' + f0(t.km / t.h) + ' km/jam', alasan: 'gps_spike' });
      for (const l of leaves) if (l.nm === t.d && l.a != null && day >= l.a && (l.b == null || day <= l.b)) sug.push({ t, why: 'DID dipakai saat tercatat cuti/tidak bertugas', alasan: 'did_tertukar' });
      for (const l of outs) if (l.nm === t.d && l.a != null && day > l.a) sug.push({ t, why: 'DID dipakai setelah tanggal keluar', alasan: 'did_tertukar' });
    }
    const existing = new Set(rows.flatMap(c => c.trips || []));
    const sg2 = sug.filter(s => !existing.has(s.t.k) && !VT.isGeneric(s.t.d));
    table($('#kSug'), [{ k: 'd', l: 'Nama DID', v: s => s.t.d }, { k: 'w', l: 'Waktu', v: s => s.t.t0, f: s => VT.dtxt(s.t.t0) + ' ' + String(VT.hourOf(s.t.t0)).padStart(2, '0') + ':xx' }, { k: 'pl', l: 'Plat', v: s => s.t.pl }, { k: 'km', l: 'KM', v: s => s.t.km, f: s => f0(s.t.km), cls: 'r' }, { k: 'h', l: 'Jam', v: s => s.t.h, f: s => f2(s.t.h), cls: 'r' },
      { k: 'os', l: 'Overspeed', v: s => s.t.spd, cls: 'c' }, { k: 'why', l: 'Temuan' }, { k: 'x', l: '', nosort: true, f: (s) => `<button class="btn sm" data-s="${esc(s.t.k)}|${s.alasan}">Ajukan koreksi</button>` }], sg2, { limit: 50 });
    $('#kSug').addEventListener('click', e => { const b = e.target.closest('[data-s]'); if (!b) return; const [k, al] = [b.dataset.s.slice(0, b.dataset.s.lastIndexOf('|')), b.dataset.s.slice(b.dataset.s.lastIndexOf('|') + 1)]; const t = S.trips.find(x => x.k === k); corrForm({ nama: t.d, alasan: al, pid: S.res.pidOf[t.bw], trips: [k], keterangan: sg2.find(s => s.t.k === k).why }); });
  }
};
function corrScope(c) {
  if (c.trips && c.trips.length) return c.trips.length + ' trip dipilih';
  const a = [];
  if (c.pid) a.push(perLabel(+c.pid));
  if (c.dari || c.sampai) a.push((c.dari || '…') + ' s.d. ' + (c.sampai || '…'));
  if (c.plat) a.push('plat ' + c.plat);
  return a.join(' · ') || '—';
}
function corrForm(pre) {
  const o = Object.assign({ id: null, nama: '', alasan: 'did_pinjam', aksi: 'reassign', komponen: [], pemakai: '', pid: S.res.lastPid, dari: '', sampai: '', plat: '', trips: [], bukti: '', keterangan: '', oleh: '', status: 'Diajukan', approver: '' }, pre || {});
  if (!pre || !pre.aksi) { const r = VT.CORR_REASONS.find(x => x.id === o.alasan); if (r) { o.aksi = r.action; o.komponen = (r.comps || []).slice(); } }
  const md = modal(o.id ? 'Ubah faktor koreksi' : 'Ajukan faktor koreksi', `<div class="form">
    <div class="fld"><label for="cN">Nama DID (pemilik tag yang tercatat)</label><input class="inp" id="cN" list="dlC" value="${esc(o.nama)}" required>${drvDatalist('dlC')}</div>
    <div class="fld"><label for="cA">Alasan</label><select class="inp" id="cA">${VT.CORR_REASONS.map(r => `<option value="${r.id}" ${o.alasan === r.id ? 'selected' : ''}>${esc(r.label)}</option>`).join('')}</select></div>
    <div class="fld"><label for="cX">Aksi</label><select class="inp" id="cX">${Object.keys(VT.CORR_ACTIONS).map(k => `<option value="${k}" ${o.aksi === k ? 'selected' : ''}>${VT.CORR_ACTIONS[k]}</option>`).join('')}</select></div>
    <div class="fld" id="fPem"><label for="cU">Pemakai sebenarnya (nama DID)</label><input class="inp" id="cU" list="dlC" value="${esc(o.pemakai)}"><div class="help">Pelanggaran dipindah ke nama ini. Kosongkan = keluarkan dari skor.</div></div>
    <div class="fld full" id="fKom"><label>Komponen yang dianulir</label><div class="checks">${VT.COMP.map(c => `<label><input type="checkbox" value="${c.k}" ${o.komponen.includes(c.k) ? 'checked' : ''}> ${c.name}</label>`).join('')}</div></div>
    <div class="fld"><label for="cP">Periode</label><select class="inp" id="cP"><option value="">(pakai tanggal)</option>${periodOptions(o.pid ? +o.pid : null)}</select></div>
    <div class="fld"><label for="cD1">Atau tanggal dari</label><input class="inp" type="date" id="cD1" value="${esc(o.dari)}"></div>
    <div class="fld"><label for="cD2">sampai</label><input class="inp" type="date" id="cD2" value="${esc(o.sampai)}"></div>
    <div class="fld"><label for="cPl">Plat nomor (opsional)</label><input class="inp" id="cPl" value="${esc(o.plat)}"></div>
    <div class="fld full"><label>Trip tertentu (opsional)</label><div class="tools" style="margin:0"><button class="btn sm" id="cLoad" type="button">Tampilkan trip sesuai lingkup</button><span class="muted small" id="cSel">${o.trips.length ? o.trips.length + ' trip dipilih' : 'semua trip dalam lingkup'}</span></div><div id="cTrips"></div></div>
    <div class="fld"><label for="cB">No. bukti (surat tugas / laporan VT / tiket vendor / JRM)</label><input class="inp" id="cB" value="${esc(o.bukti)}"></div>
    <div class="fld"><label for="cO">Diajukan oleh</label><input class="inp" id="cO" value="${esc(o.oleh)}"></div>
    <div class="fld"><label for="cS">Status</label><select class="inp" id="cS">${['Diajukan', 'Disetujui', 'Ditolak'].map(s => `<option ${o.status === s ? 'selected' : ''}>${s}</option>`).join('')}</select></div>
    <div class="fld"><label for="cAp">Disetujui oleh (QSHE)</label><input class="inp" id="cAp" value="${esc(o.approver)}"></div>
    <div class="fld full"><label for="cK">Keterangan / kronologi</label><textarea class="inp" id="cK">${esc(o.keterangan)}</textarea></div></div>
    <div class="note" id="cPrev"></div><div class="tools"><span class="sp"></span><button class="btn" id="cTest">Cek efek</button><button class="btn pri" id="cSave">Simpan</button></div>`, { wide: true });
  const B = md.body; let sel = new Set(o.trips);
  const sync = () => { const a = $('#cX', B).value; $('#fPem', B).hidden = a !== 'reassign'; $('#fKom', B).hidden = a !== 'annul'; };
  $('#cA', B).onchange = () => { const r = VT.CORR_REASONS.find(x => x.id === $('#cA', B).value); $('#cX', B).value = r.action; $$('#fKom input', B).forEach(i => { i.checked = (r.comps || []).includes(i.value); }); sync(); };
  $('#cX', B).onchange = sync; sync();
  const read = () => ({ id: o.id || uid(), nama: VT.norm($('#cN', B).value), alasan: $('#cA', B).value, aksi: $('#cX', B).value, komponen: $$('#fKom input:checked', B).map(i => i.value), pemakai: VT.norm($('#cU', B).value),
    pid: $('#cP', B).value ? +$('#cP', B).value : '', dari: $('#cD1', B).value, sampai: $('#cD2', B).value, plat: $('#cPl', B).value.trim(), trips: [...sel], bukti: $('#cB', B).value.trim(), oleh: $('#cO', B).value.trim(), status: $('#cS', B).value, approver: $('#cAp', B).value.trim(), keterangan: $('#cK', B).value.trim(), dibuat: o.dibuat || today() });
  const scopeTrips = c => { const t2 = Object.assign({}, c, { trips: [] }); return S.trips.filter(t => VT.corrMatches(t2, t, S.res.pidOf)); };
  $('#cLoad', B).onclick = () => {
    const c = read(); const ts = scopeTrips(c).sort((a, b) => a.t0 - b.t0);
    if (!c.nama) { toast('Isi nama DID dulu.'); return; }
    if (!ts.length) { $('#cTrips', B).innerHTML = '<p class="note">Tidak ada trip untuk nama & lingkup ini. Pilih periode atau tanggal.</p>'; return; }
    $('#cTrips', B).innerHTML = `<div class="tbox" style="max-height:260px"><table class="t"><thead><tr><th><input type="checkbox" id="cAll" aria-label="Pilih semua"></th><th>Waktu</th><th>Plat</th><th class="r">KM</th><th class="r">Jam</th><th>Pelanggaran</th></tr></thead><tbody>${ts.map(t => `<tr><td><input type="checkbox" data-k="${esc(t.k)}" ${sel.has(t.k) ? 'checked' : ''}></td><td>${esc(VT.dtxt(t.t0))} ${String(VT.hourOf(t.t0)).padStart(2, '0')}:xx</td><td>${esc(t.pl)}</td><td class="r">${f0(t.km)}</td><td class="r">${f2(t.h)}</td><td>${esc(VT.tripViol(t, P()).text || '—')}</td></tr>`).join('')}</tbody></table></div>`;
    const upd = () => { $('#cSel', B).textContent = sel.size ? sel.size + ' trip dipilih' : 'semua trip dalam lingkup'; };
    $$('[data-k]', B).forEach(i => i.onchange = () => { i.checked ? sel.add(i.dataset.k) : sel.delete(i.dataset.k); upd(); });
    $('#cAll', B).onchange = e => { $$('[data-k]', B).forEach(i => { i.checked = e.target.checked; e.target.checked ? sel.add(i.dataset.k) : sel.delete(i.dataset.k); }); upd(); };
  };
  const preview = () => {
    const c = read(); const ts = c.trips.length ? S.trips.filter(t => c.trips.includes(t.k)) : scopeTrips(c);
    const s = k => ts.reduce((a, t) => a + t[k], 0);
    $('#cPrev', B).innerHTML = `<b>Efek bila disetujui:</b> ${ts.length} trip · ${f0(s('km'))} km · overspeed ${s('spd')} · fatigue ${s('fat')} · braking ${s('brk')} · accel ${s('acc')}.`;
    return { c, ts };
  };
  $('#cTest', B).onclick = preview;
  $('#cSave', B).onclick = async () => {
    const { c, ts } = preview();
    if (!c.nama) { toast('Isi nama DID.'); return; }
    if (!c.pid && !c.dari && !c.sampai && !c.plat && !c.trips.length) { toast('Batasi lingkup: pilih periode, tanggal, plat, atau trip.'); return; }
    if (!ts.length) { toast('Tidak ada trip yang cocok dengan lingkup ini.'); return; }
    if (c.aksi === 'annul' && !c.komponen.length) { toast('Pilih minimal satu komponen yang dianulir.'); return; }
    if (c.aksi === 'reassign' && c.pemakai === c.nama) { toast('Pemakai sebenarnya harus berbeda dari pemilik DID.'); return; }
    if (c.status === 'Disetujui' && !c.bukti) { toast('Koreksi yang disetujui wajib punya nomor bukti.'); return; }
    if (c.alasan === 'lain' && !c.keterangan) { toast('Alasan lain wajib diberi keterangan.'); return; }
    if (c.status === 'Disetujui' && !c.approver) c.approver = 'QSHE';
    const arr = S.inputs.corrections || (S.inputs.corrections = []);
    const i = arr.findIndex(x => x.id === c.id); if (i >= 0) arr[i] = c; else arr.push(c);
    saveInputs(true); md.close(); await recompute(); toast(c.status === 'Disetujui' ? 'Koreksi disetujui & skor dihitung ulang.' : 'Koreksi tersimpan (menunggu persetujuan).'); route();
  };
}

/* ------------------------------------------------------------ FAKTOR LAIN */
const FDEF = {
  insiden: { cols: [['tanggal', 'Tanggal kejadian', 'date'], ['nama', 'Nama pengemudi (DID)', 'drv'], ['cabang', 'Cabang', 'cab'], ['nolap', 'No. laporan', 'text'], ['keparahan', 'Tingkat keparahan', 'sev'], ['kesalahan', 'Tingkat kesalahan pengemudi', 'fault'], ['capa', 'CAPA selesai tepat waktu?', 'yn'], ['ket', 'Kronologi singkat', 'area']],
    effect: 'Mengurangi skor pengemudi (keparahan × tingkat kesalahan) dan pemimpin cabang (FR/CSS/Unit/Section Head), memudar linier selama masa berlaku insiden.' },
  promosi: { cols: [['tanggal', 'Tanggal efektif', 'date'], ['nama', 'Nama DID', 'drv'], ['jabatan', 'Jabatan baru', 'role'], ['jabLama', 'Jabatan lama', 'text'], ['ket', 'Keterangan', 'area']],
    effect: 'Mulai periode yang berakhir setelah tanggal efektif, jabatan baru dipakai: pemimpin (Foreman/CSS/Unit/Section Head) ikut menanggung insiden di cabangnya dan muncul di Rekap pemimpin.' },
  mutasi: { cols: [['tanggal', 'Tanggal efektif', 'date'], ['nama', 'Nama DID', 'drv'], ['cabangLama', 'Cabang lama', 'cab'], ['cabang', 'Cabang baru', 'cab'], ['ket', 'Keterangan', 'area']],
    effect: 'Mulai periode yang berakhir setelah tanggal efektif, pengemudi dihitung di cabang baru (rekap cabang, pemimpin yang bertanggung jawab, notifikasi).' },
  cuti: { cols: [['dari', 'Dari tanggal', 'date'], ['sampai', 'Sampai tanggal', 'date'], ['nama', 'Nama DID', 'drv'], ['jenis', 'Jenis', 'leave'], ['ket', 'Keterangan', 'area']],
    effect: 'Trip atas nama ini selama cuti ditandai di Saran koreksi (kemungkinan DID dipinjam/tertukar) dan dijelaskan di laporan.' },
  keluar: { cols: [['tanggal', 'Tanggal keluar', 'date'], ['nama', 'Nama DID', 'drv'], ['alasan', 'Alasan', 'text'], ['ket', 'Keterangan', 'area']],
    effect: 'Setelah tanggal ini tidak ada notifikasi ke atasan; trip yang masih memakai DID ini ditandai di Saran koreksi.' },
  pelatihan: { cols: [['tanggal', 'Tanggal selesai', 'date'], ['nama', 'Nama DID', 'drv'], ['jenis', 'Jenis pelatihan / coaching', 'train'], ['penyelenggara', 'Penyelenggara / coach', 'text'], ['ket', 'Keterangan', 'area']],
    effect: 'Tidak mengubah skor. Dipakai untuk analisa efektivitas: rata-rata skor 3 periode sebelum vs sesudah pelatihan.' },
  nondriver: { cols: [['nama', 'Nama DID / ID', 'drv'], ['alasan', 'Alasan (ID pool, kendaraan bersama, dll)', 'text'], ['ket', 'Keterangan', 'area']],
    effect: 'ID ini dikeluarkan dari skor personal (seperti ID umum DRIVER…/MASTER…).' }
};
function factorInput(t, k, val) {
  const Pp = P();
  const id = 'f_' + k;
  if (t === 'date') return `<input class="inp" type="date" id="${id}" value="${esc(val || '')}">`;
  if (t === 'drv') return `<input class="inp" id="${id}" list="dlF" value="${esc(val || '')}">`;
  if (t === 'cab') return `<input class="inp" id="${id}" list="dlCab" value="${esc(val || '')}">`;
  if (t === 'area') return `<textarea class="inp" id="${id}">${esc(val || '')}</textarea>`;
  const opt = { sev: Pp.SEV.map(s => s[0]), fault: Pp.FAULT.map(f => f[0]), yn: ['Tidak', 'Ya'], role: ['Foreman', 'CS Supervisor', 'Unit Head', 'Section Head', 'Branch Head', 'Teknisi', 'Lainnya'], leave: ['Cuti tahunan', 'Sakit', 'Dinas non-mengemudi', 'Lainnya'], train: ['Defensive driving', 'Coaching 1-on-1', 'Refreshment KIMHAP', 'Ride-along', 'Lainnya'] }[t];
  if (opt) return `<select class="inp" id="${id}">${opt.map(o => `<option ${o === val ? 'selected' : ''}>${esc(o)}</option>`).join('')}</select>`;
  return `<input class="inp" type="text" id="${id}" value="${esc(val || '')}">`;
}
PAGES.faktor = {
  title: 'Faktor lain', render(v, q) {
    const tab = q.get('t') || 'insiden';
    const def = FDEF[tab]; const ft = VT.FACTOR_TYPES.find(f => f.id === tab);
    v.innerHTML = head('Faktor lain', 'Kejadian di luar data VT yang memengaruhi penilaian: insiden/kecelakaan, promosi, mutasi, cuti, resign, pelatihan, ID bersama.') + `<div class="pad">
      <div class="tabs" role="tablist">${VT.FACTOR_TYPES.map(f => `<button role="tab" aria-selected="${f.id === tab}" data-t="${f.id}">${esc(f.label)} <span class="muted small">(${(S.inputs.factors[f.id] || []).length})</span></button>`).join('')}</div>
      <p class="note"><b>Efek:</b> ${esc(def.effect)}</p>
      <div class="card"><div class="form">${def.cols.map(([k, l, t]) => `<div class="fld ${t === 'area' ? 'full' : ''}"><label for="f_${k}">${esc(l)}</label>${factorInput(t, k, '')}</div>`).join('')}</div>${drvDatalist('dlF')}${cabDatalist('dlCab')}
      <div class="tools"><span class="sp"></span><button class="btn pri" id="fSave">Simpan ${esc(ft.label.toLowerCase())}</button></div></div>
      <div id="fList" style="margin-top:12px"></div><div id="fExtra"></div></div>`;
    $$('[data-t]', v).forEach(b => b.onclick = () => { location.hash = '#/faktor?t=' + b.dataset.t; });
    let editing = null;
    $('#fSave').onclick = async () => {
      const o = { id: editing || uid() };
      def.cols.forEach(([k, , t]) => { const el = $('#f_' + k); o[k] = t === 'drv' ? VT.norm(el.value) : el.value.trim(); });
      if (!o.nama) { toast('Nama wajib diisi.'); return; }
      if (def.cols.some(c => c[2] === 'date' && c[0] !== 'sampai') && !def.cols.filter(c => c[2] === 'date' && c[0] !== 'sampai').every(c => o[c[0]])) { toast('Tanggal wajib diisi.'); return; }
      if (tab === 'insiden' && !o.cabang && S.res) { const r = S.res.rows.filter(x => x.drv === o.nama).slice(-1)[0]; if (r) o.cabang = r.cab; }
      const arr = S.inputs.factors[tab]; const i = arr.findIndex(x => x.id === o.id); if (i >= 0) arr[i] = o; else arr.push(o);
      saveInputs(true); await recompute(); toast('Tersimpan. Skor dihitung ulang.'); route();
    };
    const rows = (S.inputs.factors[tab] || []).slice().reverse();
    table($('#fList'), def.cols.map(([k, l]) => ({ k, l })).concat([{ k: 'x', l: '', nosort: true, f: r => `<button class="btn sm" data-e="${r.id}">Ubah</button> <button class="btn sm danger" data-d="${r.id}">Hapus</button>` }]), rows, { exportName: 'faktor_' + tab });
    $('#fList').addEventListener('click', async e => {
      const ed = e.target.closest('[data-e]'), dl = e.target.closest('[data-d]');
      if (ed) { const r = rows.find(x => x.id === ed.dataset.e); editing = r.id; def.cols.forEach(([k]) => { const el = $('#f_' + k); el.value = r[k] || ''; }); window.scrollTo({ top: 0, behavior: 'smooth' }); toast('Ubah isian lalu klik Simpan.'); }
      if (dl && await confirmBox('Hapus data ini?', 'Hapus')) { S.inputs.factors[tab] = S.inputs.factors[tab].filter(x => x.id !== dl.dataset.d); saveInputs(true); await recompute(); route(); }
    });
    if (tab === 'pelatihan' && S.res && rows.length) {
      const out = rows.map(r => {
        const d = VT.isoToDay(r.tanggal); const a = S.res.rows.filter(x => x.drv === r.nama);
        const pe = p => per(p).end;
        const before = a.filter(x => pe(x.PID) < d).slice(-3), after = a.filter(x => per(x.PID).start > d).slice(0, 3);
        const b = VT.avg(before.map(x => x.FIN)), af = VT.avg(after.map(x => x.FIN));
        return { nama: r.nama, jenis: r.jenis, tanggal: r.tanggal, b, af, d: b != null && af != null ? af - b : null, nb: before.length, na: after.length, rb: before.filter(x => x.NREP > 0).length, ra: after.filter(x => x.NREP > 0).length };
      });
      $('#fExtra').innerHTML = '<h3 class="sec">Efektivitas pelatihan (3 periode sebelum vs sesudah)</h3><div id="fEf"></div>';
      table($('#fEf'), [{ k: 'nama', l: 'Nama' }, { k: 'jenis', l: 'Jenis' }, { k: 'tanggal', l: 'Tanggal' }, { k: 'b', l: 'Skor sebelum', f: r => f1(r.b) + ` <span class="muted small">(${r.nb} per.)</span>`, cls: 'r' }, { k: 'af', l: 'Skor sesudah', f: r => f1(r.af) + ` <span class="muted small">(${r.na} per.)</span>`, cls: 'r' }, { k: 'd', l: 'Perubahan', f: r => `<span class="${r.d > 0 ? 'dn' : r.d < 0 ? 'up' : ''}">${sg(r.d)}</span>`, cls: 'r' }, { k: 'rb', l: 'Periode berulang sebelum → sesudah', f: r => r.rb + ' → ' + r.ra, cls: 'c' }], out, { search: false, free: true });
    }
  }
};

/* ------------------------------------------------------------ MASTER KARYAWAN */
PAGES.master = {
  title: 'Master karyawan', render(v) {
    const ms = S.inputs.master || (S.inputs.master = []);
    const active = S.res ? new Set(S.res.rows.filter(r => r.PID === S.res.lastPid).map(r => r.drv)) : new Set();
    const lastCab = {}; if (S.res) S.res.rows.forEach(r => { lastCab[r.drv] = r.cab; });
    v.innerHTML = head('Master karyawan, email & atasan', 'Dasar notifikasi email. Nama DID harus sama dengan data VT. Jabatan FR/CSS terbaca otomatis dari nama DID. Impor dari sheet 10 workbook Excel atau template CSV di menu Unggah data.') + `<div class="pad">
      <div class="tools"><button class="btn pri" id="mAdd">+ Tambah karyawan</button><button class="btn" id="mFill" ${S.res ? '' : 'disabled'}>Tambahkan semua pengemudi aktif yang belum ada</button><a class="btn" href="#/unggah">Impor file master</a><span class="sp"></span>
      <span class="note">${ms.length} karyawan · ${ms.filter(m => /@/.test(m.atasanEmail || '')).length} punya email atasan</span></div><div id="t"></div></div>`;
    const sugg = d => { if (!S.res) return ''; const c = lastCab[d]; if (!c) return ''; const l = [...new Set(S.res.rows.filter(r => r.cab === c && (VT.leaderCode(r.drv, 'FR') || VT.leaderCode(r.drv, 'CSS'))).map(r => r.drv))].filter(x => x !== d); return l.join(' / '); };
    const cols = [
      { k: 'did', l: 'Nama DID', f: r => `<b>${esc(r.did)}</b>${active.has(r.did) ? ' <span class="pill-o">aktif</span>' : ''}` },
      ...VT.MASTER_FIELDS.slice(1).map(([k, l]) => ({ k, l, f: r => `<input class="cell" data-did="${esc(r.did)}" data-k="${k}" value="${esc(r[k] || '')}" ${k === 'cab' && !r.cab && lastCab[r.did] ? `placeholder="${esc(lastCab[r.did])}"` : ''}>` })),
      { k: 'sar', l: 'Saran atasan (FR/CSS cabang sama)', v: r => sugg(r.did) },
      { k: 'ok', l: 'Kelengkapan', v: r => (!/@/.test(r.atasanEmail || '') ? 'Lengkapi email atasan langsung' : !/@/.test(r.atasan2Email || '') ? 'Email atasan tidak langsung kosong' : 'Lengkap'), f: r => { const t = !/@/.test(r.atasanEmail || '') ? ['MERAH', 'Lengkapi email atasan langsung'] : !/@/.test(r.atasan2Email || '') ? ['KUNING', 'Email atasan tidak langsung kosong'] : ['HIJAU', 'Lengkap']; return `<span class="st ${t[0]}">${t[1]}</span>`; } },
      { k: 'x', l: '', nosort: true, f: r => `<button class="btn sm danger" data-d="${esc(r.did)}">Hapus</button>` }];
    table($('#t'), cols, ms.slice().sort((a, b) => (a.did < b.did ? -1 : 1)), { exportName: 'master_karyawan', limit: 300 });
    $('#t').addEventListener('change', e => { const i = e.target.closest('[data-did]'); if (!i) return; const m = ms.find(x => x.did === i.dataset.did); if (m) { m[i.dataset.k] = i.value.trim(); saveInputs(true); } });
    $('#t').addEventListener('click', async e => { const b = e.target.closest('[data-d]'); if (b && await confirmBox('Hapus ' + b.dataset.d + ' dari master?', 'Hapus')) { S.inputs.master = ms.filter(x => x.did !== b.dataset.d); saveInputs(true); route(); } });
    $('#mAdd').onclick = () => {
      const md = modal('Tambah karyawan', `<div class="form">${VT.MASTER_FIELDS.map(([k, l]) => `<div class="fld"><label for="m_${k}">${l}</label><input class="inp" id="m_${k}" ${k === 'did' ? 'list="dlM"' : ''}></div>`).join('')}</div>${drvDatalist('dlM')}<div class="tools"><span class="sp"></span><button class="btn pri" id="mS">Simpan</button></div>`, { wide: true });
      $('#mS', md.body).onclick = () => { const o = {}; VT.MASTER_FIELDS.forEach(([k]) => { o[k] = $('#m_' + k, md.body).value.trim(); }); o.did = VT.norm(o.did); if (!o.did) { toast('Nama DID wajib.'); return; } if (ms.some(x => x.did === o.did)) { toast('Nama DID sudah ada.'); return; } if (!o.cab) o.cab = lastCab[o.did] || ''; ms.push(o); saveInputs(true); md.close(); route(); };
    };
    $('#mFill').onclick = () => { let n = 0; const have = new Set(ms.map(m => m.did)); [...active].sort().forEach(d => { if (!have.has(d)) { ms.push({ did: d, cab: lastCab[d] || '' }); n++; } }); saveInputs(true); toast(n + ' pengemudi aktif ditambahkan. Lengkapi kolom atasan & email.'); route(); };
  }
};

/* ------------------------------------------------------------ PERBANDINGAN */
PAGES.banding = {
  title: 'Perbandingan', render(v, q) {
    if (needData(v)) return;
    const mode = q.get('mode') || 'per';
    const cab = q.get('cab') || 'Semua cabang';
    const months = [...new Set(S.res.tripsUsed.map(t => t.mo))].sort();
    let A, B;
    if (mode === 'per') { B = +(q.get('b') || S.res.lastPid); A = +(q.get('a') || Math.max(1, B - 1)); }
    else { B = q.get('b') || months[months.length - 1]; A = q.get('a') || months[Math.max(0, months.length - 2)]; }
    const opts = sel => (mode === 'per' ? periodOptions(sel) : months.slice().reverse().map(m => `<option value="${m}" ${m === sel ? 'selected' : ''}>${VT.monthLabel(m)}</option>`).join(''));
    v.innerHTML = head('Perbandingan periode / bulan', 'Bandingkan dua periode 2-mingguan (mis. W1–W2 vs W3–W4) atau dua bulan kalender: KPI, cabang, dan pengemudi yang paling membaik / memburuk.') + `<div class="pad">
      <div class="tools"><label class="small" for="bm"><b>Mode</b></label><select id="bm"><option value="per">Periode 2-mingguan</option><option value="mon">Bulanan</option></select>
      <label class="small" for="ba"><b>A (pembanding)</b></label><select id="ba">${opts(A)}</select><label class="small" for="bb"><b>B (dianalisa)</b></label><select id="bb">${opts(B)}</select>
      <label class="small" for="bc"><b>Cabang</b></label><select id="bc"><option>Semua cabang</option>${cabList().map(c => `<option ${c === cab ? 'selected' : ''}>${esc(c)}</option>`).join('')}</select><span class="sp"></span><button class="btn" id="bX">Ekspor Excel</button></div>
      <div class="note" id="bNar"></div><div id="bK"></div><div class="grid2" style="margin-top:12px"><div class="panel"><h4>Laju pelanggaran A vs B</h4><div class="chartbox"><canvas id="cB1"></canvas></div></div><div class="panel"><h4>Sebaran grade A vs B</h4><div class="chartbox"><canvas id="cB2"></canvas></div></div></div>
      <h3 class="sec">Per cabang</h3><div id="bC"></div><div class="grid2"><div><h3 class="sec">Paling membaik</h3><div id="bUp"></div></div><div><h3 class="sec">Paling memburuk</h3><div id="bDn"></div></div></div></div>`;
    $('#bm').value = mode;
    const go = () => { const m = $('#bm').value; location.hash = '#/banding?mode=' + m + (m === mode ? '&a=' + $('#ba').value + '&b=' + $('#bb').value : '') + '&cab=' + encodeURIComponent($('#bc').value); };
    ['bm', 'ba', 'bb', 'bc'].forEach(i => { $('#' + i).onchange = go; });
    // kumpulan baris skor per sisi
    const pidsOf = x => (mode === 'per' ? [x] : S.res.periods.filter(p => VT.monthKey(p.start) === x).map(p => p.pid));
    const rowsOf = x => { const ps = new Set(pidsOf(x)); return S.res.rows.filter(r => ps.has(r.PID) && (cab === 'Semua cabang' || r.cab === cab)); };
    const mon = mode === 'mon' ? VT.monthly(S.res, cab) : null;
    const stat = x => {
      const rs = rowsOf(x); const km = VT.sum(rs, r => r.km), h = VT.sum(rs, r => r.h);
      const o = { n: new Set(rs.map(r => r.drv)).size, km, h, fin: VT.avg(rs.map(r => r.FIN)), old: VT.avg(rs.filter(r => r.S_old != null).map(r => r.S_old)), ab: rs.length ? rs.filter(r => r.GR <= 'B').length / rs.length : null,
        osv: km ? VT.sum(rs, r => r.spd) / km * 1000 : null, fat: h ? VT.sum(rs, r => r.fat) / h * 100 : null, night: h ? VT.sum(rs, r => r.nh) / h * 100 : null, brk: km ? VT.sum(rs, r => r.brk) / km * 1000 : null, acc: km ? VT.sum(rs, r => r.acc) / km * 1000 : null,
        rep: rs.filter(r => r.NREP > 0).length, notif: rs.filter(r => r.NOTIF > 0).length, merah: rs.filter(r => r.OSS === 'MERAH').length, grades: 'ABCDE'.split('').map(g => rs.filter(r => r.GR === g).length), rs };
      if (mon) { const m = mon.find(z => z.mo === x); if (m) { o.km = m.km; o.h = m.h; o.n = m.drivers; o.osv = m.km ? m.spdc / m.km * 1000 : null; o.brk = m.km ? m.brk / m.km * 1000 : null; o.acc = m.km ? m.acc / m.km * 1000 : null; o.fat = m.h ? m.fat / m.h * 100 : null; o.night = m.h ? m.nh / m.h * 100 : null; } }
      return o;
    };
    const a = stat(A), b = stat(B);
    const lbl = x => (mode === 'per' ? perLabel(x) : VT.monthLabel(x));
    const K = [['Pengemudi aktif', 'n', 0, 1], ['Total km', 'km', 0, 1], ['Jam mengemudi', 'h', 0, 1], ['Skor rata-rata (baru)', 'fin', 1, 1], ['Skor rata-rata (lama)', 'old', 1, 1], ['% grade A–B', 'ab', 'p', 1], ['Overspeed dihitung /1.000 km', 'osv', 2, -1], ['Fatigue /100 jam', 'fat', 2, -1], ['Night driving % jam', 'night', 1, -1], ['Harsh braking /1.000 km', 'brk', 2, -1], ['Harsh acceleration /1.000 km', 'acc', 2, -1], ['Driver-periode berulang', 'rep', 0, -1], ['Notifikasi atasan', 'notif', 0, -1], ['Overspeed MERAH', 'merah', 0, -1]];
    const fm = (x, d) => (d === 'p' ? pc(x) : fx(x, d));
    $('#bK').innerHTML = `<div class="tbox free"><table class="t"><thead><tr><th>Indikator</th><th class="r">A · ${esc(lbl(A))}</th><th class="r">B · ${esc(lbl(B))}</th><th class="r">Selisih</th><th class="r">%</th><th>Arah</th></tr></thead><tbody>${K.map(([l, k, d, good]) => { const x = a[k], y = b[k]; const df = x != null && y != null ? y - x : null; const better = df == null || df === 0 || ['n', 'km', 'h'].includes(k) ? '' : (df * good > 0 ? 'dn' : 'up'); return `<tr><td>${l}</td><td class="r">${fm(x, d)}</td><td class="r"><b>${fm(y, d)}</b></td><td class="r ${better}">${df == null ? '–' : d === 'p' ? sg(df * 100, 0) + ' pp' : sg(df, d === 0 ? 0 : d)}</td><td class="r">${x ? sg(df / x * 100, 0) + '%' : '–'}</td><td>${better === 'dn' ? '<span class="dn">membaik</span>' : better === 'up' ? '<span class="up">memburuk</span>' : '–'}</td></tr>`; }).join('')}</tbody></table></div>`;
    const dfin = b.fin != null && a.fin != null ? b.fin - a.fin : null;
    $('#bNar').innerHTML = `<b>Ringkas:</b> skor rata-rata ${esc(lbl(B))} ${f1(b.fin)} vs ${esc(lbl(A))} ${f1(a.fin)} (${sg(dfin)} poin). Overspeed dihitung ${fx(b.osv, 2)} vs ${fx(a.osv, 2)} per 1.000 km; harsh braking ${fx(b.brk, 2)} vs ${fx(a.brk, 2)}; notifikasi atasan ${b.notif} vs ${a.notif}.${mode === 'mon' ? ' Mode bulanan: skor = rata-rata driver-periode yang dimulai pada bulan tersebut; laju dihitung langsung dari trip bulan kalender.' : ''}`;
    chart('cB1', { type: 'bar', data: { labels: ['Overspeed/1.000 km', 'Fatigue/100 jam', 'Night % jam', 'Braking/1.000 km', 'Accel/1.000 km'], datasets: [{ label: 'A · ' + lbl(A), data: [a.osv, a.fat, a.night, a.brk, a.acc], backgroundColor: '#bdbdbd', borderRadius: 3 }, { label: 'B · ' + lbl(B), data: [b.osv, b.fat, b.night, b.brk, b.acc], backgroundColor: '#f7941d', borderRadius: 3 }] }, options: { plugins: { legend: { position: 'bottom' } }, scales: { y: Object.assign({ beginAtZero: true }, gridY), x: { grid: { display: false } } } } });
    chart('cB2', { type: 'bar', data: { labels: ['A', 'B', 'C', 'D', 'E'], datasets: [{ label: 'A · ' + lbl(A), data: a.grades, backgroundColor: '#bdbdbd', borderRadius: 3 }, { label: 'B · ' + lbl(B), data: b.grades, backgroundColor: '#f7941d', borderRadius: 3 }] }, options: { plugins: { legend: { position: 'bottom' } }, scales: { y: Object.assign({ beginAtZero: true, ticks: { precision: 0 } }, gridY), x: { grid: { display: false } } } } });
    const byCab = rs => { const g = {}; rs.forEach(r => { (g[r.cab] = g[r.cab] || []).push(r); }); return g; };
    const ga = byCab(a.rs), gb = byCab(b.rs);
    const cabs = [...new Set(Object.keys(ga).concat(Object.keys(gb)))].sort();
    const cr = cabs.map(c => { const x = ga[c] || [], y = gb[c] || []; const fa = VT.avg(x.map(r => r.FIN)), fb = VT.avg(y.map(r => r.FIN)); return { cab: c, na: x.length, nb: y.length, fa, fb, d: fa != null && fb != null ? fb - fa : null, notA: x.filter(r => r.NOTIF > 0).length, notB: y.filter(r => r.NOTIF > 0).length, repA: x.filter(r => r.NREP > 0).length, repB: y.filter(r => r.NREP > 0).length }; });
    table($('#bC'), [{ k: 'cab', l: 'Cabang', f: r => `<b>${esc(r.cab)}</b>` }, { k: 'na', l: 'Driver-periode A', cls: 'c' }, { k: 'nb', l: 'Driver-periode B', cls: 'c' }, { k: 'fa', l: 'Skor A', f: r => f1(r.fa), cls: 'r' }, { k: 'fb', l: 'Skor B', f: r => `<b>${f1(r.fb)}</b>`, cls: 'r' }, { k: 'd', l: 'Selisih', f: r => `<span class="${r.d > 0 ? 'dn' : r.d < 0 ? 'up' : ''}">${sg(r.d)}</span>`, cls: 'r' }, { k: 'repA', l: 'Berulang A → B', f: r => r.repA + ' → ' + r.repB, cls: 'c' }, { k: 'notA', l: 'Notifikasi A → B', f: r => r.notA + ' → ' + r.notB, cls: 'c' }], cr, { sort: 5, dir: 1 });
    const dr = rs => { const g = {}; rs.forEach(r => { (g[r.drv] = g[r.drv] || []).push(r.FIN); }); const o = {}; for (const k in g) o[k] = VT.avg(g[k]); return o; };
    const da = dr(a.rs), db = dr(b.rs);
    const mv = Object.keys(db).filter(k => da[k] != null).map(k => ({ drv: k, a: da[k], b: db[k], d: db[k] - da[k], cab: (b.rs.find(r => r.drv === k) || {}).cab }));
    const mcols = [{ k: 'drv', l: 'Pengemudi', f: r => `<b>${esc(r.drv)}</b>` }, { k: 'cab', l: 'Cabang' }, { k: 'a', l: 'Skor A', f: r => f1(r.a), cls: 'r' }, { k: 'b', l: 'Skor B', f: r => f1(r.b), cls: 'r' }, { k: 'd', l: 'Selisih', f: r => `<span class="${r.d > 0 ? 'dn' : 'up'}">${sg(r.d)}</span>`, cls: 'r' }];
    table($('#bUp'), mcols, mv.slice().sort((x, y) => y.d - x.d).slice(0, 10), { search: false, free: true, onRow: r => { location.hash = '#/pengemudi?d=' + encodeURIComponent(r.drv); } });
    table($('#bDn'), mcols, mv.slice().sort((x, y) => x.d - y.d).slice(0, 10), { search: false, free: true, onRow: r => { location.hash = '#/pengemudi?d=' + encodeURIComponent(r.drv); } });
    $('#bX').onclick = () => exportXlsx('perbandingan_' + lbl(A) + '_vs_' + lbl(B), [{ name: 'KPI', cols: [{ l: 'Indikator', x: k => k[0], w: 32 }, { l: 'A ' + lbl(A), x: k => a[k[1]] }, { l: 'B ' + lbl(B), x: k => b[k[1]] }, { l: 'Selisih', x: k => (a[k[1]] != null && b[k[1]] != null ? b[k[1]] - a[k[1]] : '') }], rows: K },
      { name: 'Per cabang', cols: [{ l: 'Cabang', k: 'cab', w: 26 }, { l: 'Skor A', k: 'fa' }, { l: 'Skor B', k: 'fb' }, { l: 'Selisih', k: 'd' }, { l: 'Notifikasi A', k: 'notA' }, { l: 'Notifikasi B', k: 'notB' }], rows: cr },
      { name: 'Pengemudi', cols: [{ l: 'Pengemudi', k: 'drv', w: 30 }, { l: 'Cabang', k: 'cab', w: 22 }, { l: 'Skor A', k: 'a' }, { l: 'Skor B', k: 'b' }, { l: 'Selisih', k: 'd' }], rows: mv }]);
  }
};

/* ------------------------------------------------------------ KEADILAN & TREN */
PAGES.keadilan = {
  title: 'Uji keadilan', render(v, q) {
    if (needData(v)) return;
    const pid = +(q.get('pid') || S.res.lastPid); const fr = VT.fairness(S.res, pid); const p = per(pid);
    v.innerHTML = head('Uji keadilan: apakah skor masih menghukum travel load tinggi?', 'FY' + p.fy + ' s.d. periode terpilih. Adil = rata-rata skor antar band hampir sama; perbedaan hanya karena perilaku.') + `<div class="pad"><div class="tools">${pidSelect('kp', pid)}</div><div id="t"></div>
      <p class="note"><b>Selisih L5 − L1:</b> skor lama ${sg(fr[4].old - fr[0].old)} · skor baru ${sg(fr[4].fin - fr[0].fin)}. Metode lama: makin jauh jalan makin turun. Metode baru: rata antar band.</p>
      <div class="panel"><h4>Skor per band</h4><div class="chartbox"><canvas id="cK"></canvas></div></div></div>`;
    $('#kp').onchange = e => { location.hash = '#/keadilan?pid=' + e.target.value; };
    table($('#t'), [{ k: 'band', l: 'Band', f: r => `<b>${r.band}</b>` }, { k: 'n', l: 'Driver-periode', cls: 'r' }, { k: 'km', l: 'KM rata-rata', f: r => f0(r.km), cls: 'r' }, { k: 'old', l: 'Skor LAMA', f: r => f1(r.old), cls: 'r' }, { k: 'vb', l: 'Skor perilaku (terkoreksi beban)', f: r => f1(r.vb), cls: 'r' }, { k: 'rep', l: 'Pengurang berulang', f: r => f1(r.rep), cls: 'r' }, { k: 'lcr', l: 'Kredit beban', f: r => f1(r.lcr), cls: 'r' }, { k: 'fin', l: 'Skor BARU akhir', f: r => `<b>${f1(r.fin)}</b>`, cls: 'r' }, { k: 'a', l: '% grade A', f: r => pc(r.a), cls: 'r' }, { k: 'de', l: '% grade D/E', f: r => pc(r.de), cls: 'r' }], fr, { search: false, free: true, exportName: 'uji_keadilan' });
    chart('cK', { type: 'bar', data: { labels: fr.map(b => b.band), datasets: [{ label: 'Skor lama', data: fr.map(b => b.old), backgroundColor: '#bdbdbd', borderRadius: 3 }, { label: 'Skor baru', data: fr.map(b => b.fin), backgroundColor: '#f7941d', borderRadius: 3 }] }, options: { scales: { y: Object.assign({ min: 60, max: 100 }, gridY), x: { grid: { display: false } } }, plugins: { legend: { position: 'bottom' } } } });
  }
};
PAGES.tren = {
  title: 'Analisa tren', render(v, q) {
    if (needData(v)) return;
    const cab = q.get('cab') || 'Semua cabang';
    const m = VT.monthly(S.res, cab);
    v.innerHTML = head('Analisa tren bulanan', 'Langsung dari trip Driving per bulan kalender (ID umum dikecualikan). Night driving aturan lama (≥19:00) vs aturan baru (mulai 20:00–04:59).') + `<div class="pad"><div class="tools"><label class="small" for="tc"><b>Cabang</b></label><select id="tc"><option>Semua cabang</option>${cabList().map(c => `<option ${c === cab ? 'selected' : ''}>${esc(c)}</option>`).join('')}</select></div>
      <div class="grid2"><div class="panel"><h4>Laju per 1.000 km</h4><div class="chartbox"><canvas id="cT1"></canvas></div></div><div class="panel"><h4>Night driving: aturan lama vs baru (trip)</h4><div class="chartbox"><canvas id="cT2"></canvas></div></div></div><div id="t" style="margin-top:12px"></div></div>`;
    $('#tc').onchange = e => { location.hash = '#/tren?cab=' + encodeURIComponent(e.target.value); };
    const L = m.map(x => VT.monthLabel(x.mo).replace(/(\w{3})\w* /, '$1 '));
    chart('cT1', { type: 'line', data: { labels: L, datasets: [{ label: 'Overspeed (count mentah)', data: m.map(x => x.km ? x.spd / x.km * 1000 : null), borderColor: '#f7941d', borderWidth: 3, tension: 0 }, { label: 'Harsh braking', data: m.map(x => x.km ? x.brk / x.km * 1000 : null), borderColor: '#111', borderWidth: 2, tension: 0 }, { label: 'Harsh accel', data: m.map(x => x.km ? x.acc / x.km * 1000 : null), borderColor: '#9e9e9e', borderWidth: 2, tension: 0 }] }, options: { scales: { y: Object.assign({ beginAtZero: true }, gridY), x: { grid: { display: false } } }, plugins: { legend: { position: 'bottom' } } } });
    chart('cT2', { type: 'bar', data: { labels: L, datasets: [{ label: 'Aturan lama (≥19:00)', data: m.map(x => x.nOld), backgroundColor: '#bdbdbd' }, { label: 'Aturan baru', data: m.map(x => x.nNew), backgroundColor: '#f7941d' }] }, options: { scales: { y: Object.assign({ beginAtZero: true }, gridY), x: { grid: { display: false } } }, plugins: { legend: { position: 'bottom' } } } });
    table($('#t'), [{ k: 'mo', l: 'Bulan', f: r => VT.monthLabel(r.mo), v: r => r.mo }, { k: 'drivers', l: 'Pengemudi', cls: 'r' }, { k: 'trips', l: 'Trip', f: r => f0(r.trips), cls: 'r' }, { k: 'km', l: 'KM', f: r => f0(r.km), cls: 'r' }, { k: 'h', l: 'Jam', f: r => f0(r.h), cls: 'r' },
      { k: 'os', l: 'Overspeed/1.000 km', v: r => (r.km ? r.spd / r.km * 1000 : null), f: r => fx(r.km ? r.spd / r.km * 1000 : null, 2), cls: 'r' }, { k: 'fat', l: 'Fatigue/100 jam', v: r => (r.h ? r.fat / r.h * 100 : null), f: r => fx(r.h ? r.fat / r.h * 100 : null, 2), cls: 'r' },
      { k: 'brk', l: 'Braking/1.000 km', v: r => (r.km ? r.brk / r.km * 1000 : null), f: r => fx(r.km ? r.brk / r.km * 1000 : null, 2), cls: 'r' }, { k: 'acc', l: 'Accel/1.000 km', v: r => (r.km ? r.acc / r.km * 1000 : null), f: r => fx(r.km ? r.acc / r.km * 1000 : null, 2), cls: 'r' },
      { k: 'nOld', l: 'Trip malam (lama)', f: r => f0(r.nOld), cls: 'r' }, { k: 'nNew', l: 'Trip malam (baru)', f: r => f0(r.nNew), cls: 'r' }], m, { search: false, exportName: 'tren_bulanan' });
  }
};

/* ------------------------------------------------------------ DETAIL PENGEMUDI */
PAGES.pengemudi = {
  title: 'Detail pengemudi', render(v, q) {
    if (needData(v)) return;
    const names = drvList();
    const d = VT.norm(q.get('d') || '');
    const rows = S.res.rows.filter(r => r.drv === d);
    const pid = +(q.get('pid') || (rows.some(x => x.PID === S.res.lastPid) ? S.res.lastPid : rows.length ? rows[rows.length - 1].PID : 0));
    v.innerHTML = head('Detail pengemudi (simulasi kasus)', 'Setiap langkah perhitungan diuraikan — sama dengan sheet 17 SIMULASI_KASUS di Excel.') + `<div class="pad"><div class="tools"><label class="small" for="pd"><b>Nama</b></label><input id="pd" list="dlP" value="${esc(d)}" placeholder="Ketik nama DID…" style="min-width:260px">${drvDatalist('dlP')}
      <label class="small" for="pp"><b>Periode</b></label><select id="pp">${rows.slice().reverse().map(r => `<option value="${r.PID}" ${r.PID === pid ? 'selected' : ''}>${esc(perLabel(r.PID))}</option>`).join('')}</select><span class="sp"></span>${d ? `<a class="btn" href="#/koreksi?d=${encodeURIComponent(d)}&pid=${pid}">Ajukan faktor koreksi</a>` : ''}</div><div id="pBody"></div></div>`;
    $('#pd').onchange = e => { const n = VT.norm(e.target.value); if (names.includes(n)) location.hash = '#/pengemudi?d=' + encodeURIComponent(n); else toast('Nama tidak ditemukan di data.'); };
    $('#pp').onchange = e => { location.hash = '#/pengemudi?d=' + encodeURIComponent(d) + '&pid=' + e.target.value; };
    if (!d) { $('#pBody').innerHTML = '<div class="empty"><b>Pilih pengemudi</b>Ketik nama DID, atau klik baris di dashboard / rekap.</div>'; return; }
    if (!rows.length) { $('#pBody').innerHTML = '<div class="empty"><b>Tidak ada data skor untuk nama ini</b>Mungkin ID umum, ditandai bukan pengemudi, atau tidak ada trip Driving.</div>'; return; }
    const r = rows.find(x => x.PID === pid) || rows[rows.length - 1]; const Pp = P(); const pr = r.prevRow; const m = S.res.master.get(d) || {};
    const F = S.inputs.factors; const facts = [];
    VT.FACTOR_TYPES.forEach(ft => (F[ft.id] || []).forEach(x => { if (VT.norm(x.nama) === d) facts.push(ft.label + ': ' + [x.tanggal || x.dari, x.jabatan || x.cabang || x.jenis || x.keparahan || x.alasan || ''].filter(Boolean).join(' · ')); }));
    const cors = (S.inputs.corrections || []).filter(c => c.nama === d || c.pemakai === d);
    const ex = c => (c.ex === 'km' ? r.km : r.h);
    $('#pBody').innerHTML = `<div class="kpis">
      <div class="kpi acc"><span class="l">SKOR AKHIR · ${esc(perLabel(r.PID))}</span><span class="v">${f1(r.FIN)}</span><span class="s">grade ${r.GR} · metode lama ${f1(r.S_old)}${pr ? ' · periode lalu ' + f1(pr.FIN) : ''}</span></div>
      <div class="kpi"><span class="l">TRAVEL LOAD</span><span class="v">${f0(r.km)} km</span><span class="s">${f1(r.h)} jam · ${r.trips} trip · band ${r.band}</span></div>
      <div class="kpi"><span class="l">STATUS OVERSPEED</span><span class="v" style="font-size:24px">${esc(r.OSS)}</span><span class="s">${r.spd_raw} count mentah · ${r.spd} dihitung · ${r.sus} trip ≥${Pp.OSSUS}</span></div>
      <div class="kpi"><span class="l">CABANG · JABATAN</span><span class="v" style="font-size:20px">${esc(r.cab)}</span><span class="s">${esc(r.JAB || 'pengemudi')} · atasan: ${esc(m.atasan || '(belum di master)')}</span></div></div>
      <h3 class="sec">Langkah perhitungan</h3><div class="tbox free"><table class="t"><thead><tr><th>Komponen</th><th class="r">Kejadian</th><th class="r">Eksposur</th><th class="r">Laju</th><th class="r">Rata-rata kelompok (${r.band})</th><th class="r">Rasio terkoreksi</th><th class="r">Skor komponen</th><th class="r">Bobot</th><th class="r">Kontribusi</th><th>Berulang</th><th class="r">Pengurang</th></tr></thead><tbody>
      ${VT.COMP.map(c => { const cnt = c.k === 'night' ? r.nh : r[c.k]; const F_ = r['F_' + c.k]; return `<tr><td><b>${c.name}</b></td><td class="r">${c.k === 'night' ? f1(cnt) + ' jam' : f0(cnt)}</td><td class="r">${c.ex === 'km' ? f0(ex(c)) + ' km' : f1(ex(c)) + ' jam'}</td><td class="r">${f2(r['r_' + c.k])}</td><td class="r">${f2(r['b_' + c.k])}</td><td class="r">${f2(r['R_' + c.k])}</td><td class="r"><b>${f1(r['S_' + c.k])}</b></td><td class="r">${pc(Pp.W[c.k])}</td><td class="r">${f1(Pp.W[c.k] * r['S_' + c.k])}</td><td>${c.k === 'cor' ? '<span class="muted">dikecualikan</span>' : F_ === 2 ? '<span class="up">makin buruk</span>' : F_ === 1 ? '<span class="up">tetap</span>' : '–'}${r['N_' + c.k] ? ' · <b>NOTIF</b>' : ''}</td><td class="r">${c.k === 'cor' ? '' : f1(r['P_' + c.k])}</td></tr>`; }).join('')}
      </tbody></table></div>
      <div class="kv" style="margin-top:12px"><div>Skor perilaku terbobot</div><div>${f1(r.VB)}</div><div>+ Kredit beban (${r.band})</div><div>${f1(r.LCR)} ${r.LCR ? '' : '<span class="muted">(hanya bila overspeed 0 & tidak berulang)</span>'}</div><div>− Pengurang berulang</div><div>${f1(r.REP)} (${r.NREP} komponen, maks ${Pp.REPCAP})</div><div>− Pengurang insiden</div><div>${f1(r.INS)}</div><div>− Pengurang pemimpin</div><div>${f1(r.LDR)}${r.RF ? ' (faktor peran ' + fx(r.RF, 1) + ')' : ''}</div><div><b>= Skor akhir</b></div><div><b>${f1(r.FIN)}</b> → grade ${gradeChip(r.GR)} · notifikasi: ${r.NOTIF ? (r.ESK === 2 ? 'ESKALASI' : 'PERHATIAN') : 'tidak'}</div>
      <div>Fokus tindak lanjut</div><div><b>${r.FOC}</b> — pekerja: ${esc(VT.REK[r.FOC][1])}<br>atasan: ${esc(VT.REK[r.FOC][2])}</div></div>
      ${facts.length || cors.length ? `<h3 class="sec">Faktor koreksi & faktor lain</h3><ul>${cors.map(c => `<li>Koreksi (${esc(c.status)}): ${esc((VT.CORR_REASONS.find(x => x.id === c.alasan) || {}).label || c.alasan)} · ${esc(corrScope(c))}</li>`).join('')}${facts.map(f => `<li>${esc(f)}</li>`).join('')}</ul>` : ''}
      <div class="panel" style="margin-top:12px"><h4>Riwayat skor</h4><div class="chartbox sm"><canvas id="cP"></canvas></div></div>
      <h3 class="sec">Trip periode ini</h3><div id="pT"></div>`;
    chart('cP', { type: 'line', data: { labels: rows.map(x => VT.trend(S.res, x.PID, 'Semua cabang', 1)[0].lbl), datasets: [{ label: 'Skor baru', data: rows.map(x => x.FIN), borderColor: '#f7941d', borderWidth: 3, tension: 0, pointRadius: rows.map(x => (x.PID === r.PID ? 6 : 3)) }, { label: 'Skor lama', data: rows.map(x => x.S_old), borderColor: '#9e9e9e', borderDash: [6, 4], borderWidth: 1.75, tension: 0 }] }, options: { scales: { y: Object.assign({ suggestedMin: 50, max: 100 }, gridY), x: { grid: { display: false } } }, plugins: { legend: { position: 'bottom' } } } });
    const lbl = perLabel(r.PID);
    const ts = S.res.tripsUsed.filter(t => t.d === d && t.bw === lbl).sort((a, b) => a.t0 - b.t0);
    table($('#pT'), [{ k: 't0', l: 'Waktu', f: t => esc(VT.tripLine(t, Pp).split(' · ')[0]) }, { k: 'pl', l: 'Plat' }, { k: 'km', l: 'KM', f: t => f0(t.km), cls: 'r' }, { k: 'h', l: 'Jam', f: t => f2(t.h), cls: 'r' }, { k: 'spd', l: 'Overspeed', cls: 'c' }, { k: 'fat', l: 'Fatigue', cls: 'c' }, { k: 'brk', l: 'Braking', cls: 'c' }, { k: 'acc', l: 'Accel', cls: 'c' }, { k: 'cor', l: 'Cornering', cls: 'c' },
      { k: 'v', l: 'Catatan', v: t => [VT.tripViol(t, Pp).text, t.reassigned ? 'dipindah dari pemilik DID lain' : '', t.annulled ? 'ada komponen dianulir' : ''].filter(Boolean).join(' · ') }], ts, { exportName: 'trip_' + d });
  }
};

/* ------------------------------------------------------------ AI */
PAGES.ai = {
  title: 'AI', render(v) {
    const a = S.settings.ai;
    v.innerHTML = head('Opsi AI (opsional) – Claude', 'Aplikasi tetap lengkap tanpa AI. AI hanya merapikan bahasa executive report, email, dan laporan. Angka tidak diubah.') + `<div class="pad"><div class="card" style="max-width:860px"><div class="form">
      <div class="fld"><label for="aiP">Layanan AI</label><select class="inp" id="aiP">${Object.keys(AI.PROVIDERS).map(k => `<option value="${k}" ${a.provider === k ? 'selected' : ''}>${AI.PROVIDERS[k].name}${k === 'claude' ? ' (disarankan)' : ''}</option>`).join('')}</select></div>
      <div class="fld"><label for="aiM">Cara terhubung</label><select class="inp" id="aiM"><option value="web" ${a.mode === 'web' ? 'selected' : ''}>Buka situs AI & login dengan akun sendiri (gratis sesuai paket Anda)</option><option value="api" ${a.mode === 'api' ? 'selected' : ''}>API key Anthropic (langsung di aplikasi, hanya Claude)</option></select></div>
      <div class="fld full" id="aiK"><label for="aiKey">API key Anthropic</label><input class="inp" type="password" id="aiKey" value="${esc(a.apiKey)}" autocomplete="off" placeholder="sk-ant-…"><div class="help">Disimpan hanya di perangkat ini dan tidak ikut di file cadangan. Biaya pemakaian API mengikuti akun Anthropic Anda.</div></div>
      <div class="fld" id="aiMo"><label for="aiModel">Model Claude</label><input class="inp" id="aiModel" value="${esc(a.model)}"><div class="help">Contoh: claude-sonnet-5-5, claude-opus-5-5, claude-haiku-5-5.</div></div>
      <div class="fld full"><label><input type="checkbox" id="aiAn" ${a.anonymize ? 'checked' : ''}> Samarkan nama orang sebelum dikirim ke AI (disarankan)</label></div></div>
      <div class="tools"><button class="btn pri" id="aiSave">Simpan</button><button class="btn dark" id="aiLogin">Masuk / login ke ${esc((AI.PROVIDERS[a.provider] || AI.PROVIDERS.claude).name)}</button><button class="btn" id="aiTest">Uji koneksi</button><span class="note" id="aiMsg"></span></div></div>
      <h3 class="sec">Cara pakai</h3><ol class="steps"><li>Pilih Claude dan cara terhubung. Mode situs: klik <b>Masuk / login</b> sekali agar browser menyimpan login Anda.</li><li>Di Dashboard (executive report), Notifikasi email, atau Laporan bilingual, klik <b>Rapikan dengan AI</b>.</li><li>Mode situs: prompt disalin, Claude.ai terbuka — tempel, kirim, salin jawabannya kembali ke aplikasi. Mode API: hasil langsung muncul.</li></ol></div>`;
    const sync = () => { const api = $('#aiM').value === 'api' && $('#aiP').value === 'claude'; $('#aiK').hidden = !api; $('#aiMo').hidden = !api; };
    $('#aiM').onchange = sync; $('#aiP').onchange = () => { if ($('#aiP').value !== 'claude') $('#aiM').value = 'web'; sync(); }; sync();
    const save = () => { Object.assign(S.settings.ai, { provider: $('#aiP').value, mode: $('#aiM').value, apiKey: $('#aiKey').value.trim(), model: $('#aiModel').value.trim() || 'claude-sonnet-5-5', anonymize: $('#aiAn').checked }); saveSettings(); };
    $('#aiSave').onclick = () => { save(); toast('Pengaturan AI disimpan.'); };
    $('#aiLogin').onclick = () => { save(); window.open((AI.PROVIDERS[S.settings.ai.provider] || AI.PROVIDERS.claude).login, '_blank', 'noopener'); };
    $('#aiTest').onclick = async () => { save(); $('#aiMsg').textContent = 'Menguji…'; try { $('#aiMsg').textContent = await AI.test(); } catch (e) { $('#aiMsg').textContent = 'Gagal: ' + e.message; } };
  }
};

/* ------------------------------------------------------------ INSTAL */
let installEvt = null;
window.addEventListener('beforeinstallprompt', e => { e.preventDefault(); installEvt = e; const b = $('#btnInst'); if (b) b.disabled = false; });
window.addEventListener('appinstalled', () => { installEvt = null; toast('Aplikasi terpasang di perangkat.'); });
PAGES.instal = {
  title: 'Instal aplikasi', render(v) {
    const standalone = window.matchMedia('(display-mode: standalone)').matches || navigator.standalone;
    const ios = /iphone|ipad|ipod/i.test(navigator.userAgent) || (navigator.platform === 'MacIntel' && navigator.maxTouchPoints > 1);
    v.innerHTML = head('Instal aplikasi di perangkat', 'Seperti aplikasi Temuan Lapangan QSHE: sekali pasang, bisa dibuka dari layar utama dan tetap jalan tanpa internet.') + `<div class="pad guide">
      <div class="card"><b>Status:</b> ${standalone ? '<span class="st HIJAU">Sudah terpasang & berjalan sebagai aplikasi</span>' : '<span class="st KUNING">Dibuka di browser</span>'} · Service worker: <span id="swSt">memeriksa…</span> · Penyimpanan: <span id="perSt">memeriksa…</span>
      <div class="tools"><button class="btn pri" id="btnInst" ${installEvt ? '' : 'disabled'}>Pasang aplikasi</button><button class="btn" id="btnPersist">Aktifkan penyimpanan permanen</button><button class="btn" id="btnUpd">Periksa pembaruan</button></div>
      ${!installEvt && !standalone ? `<p class="note">${ios ? 'Di iPhone/iPad tombol tidak tersedia — ikuti langkah Safari di bawah.' : 'Bila tombol tidak aktif: buka dengan Chrome/Edge versi terbaru, atau pakai menu ⋮ browser → "Instal aplikasi" / "Tambahkan ke layar utama".'}</p>` : ''}</div>
      <h3 class="sec">Android (Chrome)</h3><ol class="steps"><li>Buka alamat aplikasi di Chrome.</li><li>Klik <b>Pasang aplikasi</b> di atas, atau menu ⋮ → <b>Instal aplikasi</b>.</li><li>Ikon "VT Score" muncul di layar utama.</li></ol>
      <h3 class="sec">iPhone / iPad (Safari)</h3><ol class="steps"><li>Buka alamat aplikasi di <b>Safari</b> (bukan Chrome).</li><li>Ketuk tombol <b>Bagikan</b> (kotak dengan panah ke atas).</li><li>Pilih <b>Tambah ke Layar Utama</b> → Tambah.</li></ol>
      <h3 class="sec">Laptop / PC (Chrome atau Edge)</h3><ol class="steps"><li>Buka alamat aplikasi.</li><li>Klik ikon instal di kolom alamat (monitor dengan panah) atau tombol <b>Pasang aplikasi</b>.</li><li>Aplikasi terbuka di jendela sendiri dan ada di Start menu.</li></ol>
      <h3 class="sec">Supaya tidak ada masalah lanjutan</h3><ul>
        <li><b>Data tersimpan di perangkat</b> (browser). Klik <b>Aktifkan penyimpanan permanen</b> agar tidak dihapus otomatis saat ruang penuh.</li>
        <li>Rutin <b>unduh cadangan</b> (menu Data &amp; cadangan) — terutama sebelum ganti perangkat, reset browser, atau hapus cache.</li>
        <li>File mentah FY penuh (±100 MB) paling nyaman diunggah dari laptop. HP cocok untuk file 2-mingguan dan untuk membaca dashboard.</li>
        <li>Saat ada versi baru, aplikasi memberi tahu; muat ulang halaman untuk memakainya. Data tidak hilang.</li></ul></div>`;
    $('#btnInst').onclick = async () => { if (!installEvt) return; installEvt.prompt(); await installEvt.userChoice.catch(() => {}); installEvt = null; route(); };
    $('#btnPersist').onclick = async () => { if (!navigator.storage || !navigator.storage.persist) { toast('Browser ini tidak mendukung penyimpanan permanen.'); return; } const ok = await navigator.storage.persist(); toast(ok ? 'Penyimpanan permanen aktif.' : 'Browser menolak; pasang aplikasi dulu lalu coba lagi.'); route(); };
    $('#btnUpd').onclick = async () => { if (!navigator.serviceWorker) return; const r = await navigator.serviceWorker.getRegistration(); if (r) { await r.update(); toast('Pemeriksaan selesai. Bila ada versi baru, muat ulang halaman.'); } };
    (async () => {
      const r = navigator.serviceWorker ? await navigator.serviceWorker.getRegistration() : null; $('#swSt').textContent = r && r.active ? 'aktif (offline siap)' : 'belum aktif';
      const p = navigator.storage && navigator.storage.persisted ? await navigator.storage.persisted() : false; const est = navigator.storage && navigator.storage.estimate ? await navigator.storage.estimate() : null;
      $('#perSt').textContent = (p ? 'permanen' : 'sementara') + (est ? ' · terpakai ' + f0(est.usage / 1048576) + ' MB dari ' + f0(est.quota / 1048576) + ' MB' : '');
    })();
  }
};

/* ------------------------------------------------------------ DATA & CADANGAN */
async function restoreBackup(j) {
  if (!j || j.app !== 'vt-score') throw new Error('Bukan file cadangan aplikasi ini.');
  if (j.inputs) { S.inputs = Object.assign(S.inputs, j.inputs); VT.FACTOR_TYPES.forEach(f => { S.inputs.factors[f.id] = S.inputs.factors[f.id] || []; }); await DB.set('inputs', S.inputs); }
  if (j.ga) { S.ga = Object.assign(S.ga, j.ga); await DB.set('ga', S.ga); }
  if (j.settings) { const key = S.settings.ai.apiKey; S.settings = Object.assign(S.settings, j.settings); S.settings.ai.apiKey = key; await DB.set('settings', S.settings); }
  let n = 0;
  if (j.periods && j.periods.length) {
    const ex = new Map((await DB.periods()).map(p => [p.lbl, p]));
    const out = j.periods.map(p => { const e = ex.get(p.lbl); if (!e) { n += p.trips.length; return p; } const seen = new Set(e.trips.map(t => t.k)); const add = p.trips.filter(t => !seen.has(t.k)); n += add.length; return { lbl: p.lbl, stats: { rows: Math.max(e.stats.rows, p.stats.rows), idle: Math.max(e.stats.idle, p.stats.idle), drive: e.trips.length + add.length, minDay: Math.min(e.stats.minDay, p.stats.minDay), maxDay: Math.max(e.stats.maxDay, p.stats.maxDay) }, trips: e.trips.concat(add) }; });
    await DB.putPeriods(out);
  }
  return 'Cadangan dipulihkan' + (j.periods ? ' · ' + f0(n) + ' trip ditambahkan' : '') + ' · dibuat ' + new Date(j.at).toLocaleString('id-ID');
}
PAGES.data = {
  title: 'Data & cadangan', render(v) {
    v.innerHTML = head('Data & cadangan', 'Semua data ada di browser perangkat ini. Cadangkan untuk berpindah perangkat atau berbagi dengan rekan QSHE.') + `<div class="pad">
      <div class="grid2"><div class="card"><h4 style="margin-top:0">Cadangan</h4>
        <div class="tools"><button class="btn pri" id="bkIn">Unduh cadangan input (kecil)</button><button class="btn" id="bkAll">Unduh cadangan lengkap + data trip</button></div>
        <p class="note">Cadangan input = faktor koreksi, faktor lain, master karyawan, parameter, data GA, log. Cadangan lengkap juga memuat trip mentah (besar). API key AI tidak ikut. Pulihkan lewat menu Unggah data (file .json).</p></div>
      <div class="card"><h4 style="margin-top:0">Ringkasan penyimpanan</h4><div id="stg" class="small">memeriksa…</div>
        <div class="tools"><button class="btn danger" id="clr">Hapus semua data di perangkat ini</button></div></div></div>
      <h3 class="sec">Kualitas data</h3><div id="dq"></div></div>`;
    const pack = async withTrips => ({ app: 'vt-score', ver: VT.VERSION, at: Date.now(), inputs: S.inputs, ga: S.ga, settings: Object.assign({}, S.settings, { ai: Object.assign({}, S.settings.ai, { apiKey: '' }) }), periods: withTrips ? await DB.periods() : undefined });
    const stamp = () => new Date().toISOString().slice(0, 16).replace(/[:T]/g, '-');
    $('#bkIn').onclick = async () => download(new Blob([JSON.stringify(await pack(false))], { type: 'application/json' }), 'vt-score_cadangan_input_' + stamp() + '.json');
    $('#bkAll').onclick = async () => { busy('Menyiapkan cadangan lengkap…'); try { download(new Blob([JSON.stringify(await pack(true))], { type: 'application/json' }), 'vt-score_cadangan_lengkap_' + stamp() + '.json'); } finally { busy(false); } };
    $('#clr').onclick = async () => { if (!(await confirmBox('Hapus SEMUA data (trip, input, master, GA, pengaturan) dari perangkat ini? Pastikan sudah mengunduh cadangan.', 'Hapus semua'))) return; await DB.clearAll(); location.reload(); };
    (async () => {
      const est = navigator.storage && navigator.storage.estimate ? await navigator.storage.estimate() : null;
      $('#stg').innerHTML = `<div class="kv"><div>Trip Driving tersimpan</div><div>${f0(S.trips.length)}</div><div>Periode</div><div>${Object.keys(S.pstats).length}</div><div>Master karyawan</div><div>${(S.inputs.master || []).length}</div><div>Faktor koreksi</div><div>${(S.inputs.corrections || []).length}</div><div>Baris data GA</div><div>${S.ga.rows.length}</div><div>Ruang terpakai</div><div>${est ? f0(est.usage / 1048576) + ' MB dari ' + f0(est.quota / 1048576) + ' MB' : '–'}</div></div>`;
    })();
    if (S.res) {
      const dq = S.res.dq; const ps = S.res.periods;
      const idle = VT.sum(ps, p => p.idle), rows = VT.sum(ps, p => p.rows);
      const notInMaster = [...new Set(S.res.rows.filter(r => r.PID === S.res.lastPid).map(r => r.drv))].filter(d => !S.res.master.has(d)).length;
      $('#dq').innerHTML = `<ul class="status">${[
        `Trip Idling: ${f0(idle)} dari ${f0(rows)} baris (${pc(rows ? idle / rows : 0)}) — tidak dipakai untuk skor.`,
        `ID umum (DRIVER…/MASTER…): ${f0(dq.generic)} trip Driving, ${dq.genericIds} ID, ${f0(dq.genericKm)} km — tidak masuk skor personal.`,
        `Trip dengan kecepatan rata-rata >120 km/jam: ${f0(dq.v120)} (${f0(dq.v120os)} ber-overspeed) — cek di Faktor koreksi → Saran.`,
        `ID ditandai bukan pengemudi personal: ${f0(dq.nondriver)} trip.`,
        `Periode dengan data <7 hari: ${ps.filter(p => p.days < 7).map(p => p.lbl).join(', ') || 'tidak ada'}.`,
        `Pengemudi aktif periode laporan terakhir yang belum ada di master: ${notInMaster}.`].map(t => `<li><span>${esc(t)}</span></li>`).join('')}</ul>`;
    } else $('#dq').innerHTML = '<p class="note">Belum ada data.</p>';
  }
};
