/* =====================================================================
   report.js — executive report strategis (dashboard) dan laporan
   profesional bilingual Indonesia / English (.docx + cetak/PDF).
   Semua kalimat disusun dari angka hasil perhitungan (tanpa AI).
   ===================================================================== */
'use strict';
const VTR = (function () {
  const T = (lang, id, en) => (lang === 'en' ? en : id);
  const COMPN = { spd: ['overspeed', 'overspeed'], fat: ['fatigue', 'fatigue'], night: ['night driving', 'night driving'], brk: ['harsh braking', 'harsh braking'], acc: ['harsh acceleration', 'harsh acceleration'] };
  const PROG = {
    spd: ['Tegakkan disiplin kecepatan: coaching 1-on-1 untuk pengemudi overspeed MERAH, evaluasi speed limiter, dan cek target waktu kerja yang memicu ngebut', 'Enforce speed discipline: 1-on-1 coaching for RED overspeed drivers, evaluate speed limiters, and review job time targets that push speeding'],
    fat: ['Terapkan journey management untuk perjalanan jauh: titik istirahat tiap 2 jam, batas 4 jam nyetir, dan 2 pengemudi untuk rute panjang', 'Apply journey management on long trips: rest stops every 2 hours, 4-hour driving cap, and two drivers on long routes'],
    night: ['Wajibkan izin JRM untuk perjalanan mulai 20:00–04:59 dan jadwalkan ulang pekerjaan non-darurat ke siang hari', 'Require JRM approval for trips starting 20:00–04:59 and reschedule non-urgent work to daytime'],
    brk: ['Program defensive driving (jaga jarak 3–5 detik) dan pemeriksaan rem & ban di unit dengan harsh braking berulang', 'Run defensive driving (3–5 second following distance) and brake & tyre checks on units with repeated harsh braking'],
    acc: ['Coaching smooth driving dan tinjau tekanan jadwal kerja yang memicu akselerasi kasar', 'Coach smooth driving and review schedule pressure that triggers harsh acceleration']
  };
  const avg = VT.avg, sum = VT.sum, fx = VT.fx;
  const f1 = v => fx(v, 1), f0 = v => fx(v, 0), pc = v => (v == null || !isFinite(v) ? '–' : fx(v * 100, 0) + '%');
  const en1 = v => (v == null || !isFinite(v) ? '–' : v.toLocaleString('en-US', { maximumFractionDigits: 1, minimumFractionDigits: 1 }));
  const en0 = v => (v == null || !isFinite(v) ? '–' : Math.round(v).toLocaleString('en-US'));

  /** Ringkasan strategis untuk periode & cabang terpilih */
  function exec(res, pid, cab, lang, S) {
    const P = res.P; const n1 = lang === 'en' ? en1 : f1; const n0 = lang === 'en' ? en0 : f0;
    const k = VT.kpis(res, pid, cab);
    const s = VT.sel(res, pid, cab);
    const tr = VT.trend(res, pid, cab, 3).filter(x => x.fin != null);
    const p = res.periods.find(x => x.pid === pid);
    const delta = k.finPrev != null ? k.fin - k.finPrev : null;
    // komponen risiko utama
    const comp = VT.REPC.map(c => ({ c, rep: s.filter(r => r['F_' + c] > 0).length, notif: s.filter(r => r['N_' + c] > 0).length })).sort((a, b) => b.rep - a.rep || b.notif - a.notif);
    const top = comp[0] && comp[0].rep > 0 ? comp[0] : null, sec = comp[1] && comp[1].rep > 0 ? comp[1] : null;
    const br = VT.branches(res, pid, P).filter(b => b.n >= P.MIN_BRANCH && (cab === 'Semua cabang' || b.cab === cab));
    const worst = br.slice().sort((a, b) => a.fin - b.fin).slice(0, 3), best = br.slice().sort((a, b) => b.fin - a.fin).slice(0, 2);
    const notifTot = s.filter(r => r.NOTIF > 0).length;
    const notifTop3 = br.slice().sort((a, b) => b.notif - a.notif).slice(0, 3);
    const conc = notifTot ? sum(notifTop3, b => b.notif) / notifTot : 0;
    const esk = s.filter(r => r.ESK === 2).sort((a, b) => a.FIN - b.FIN).slice(0, 3);
    const ldr = s.filter(r => r.LDR > 0).length, ins = s.filter(r => r.INS > 0).length;
    const nts = VT.buildNotifications(res, pid, P, res.master, res.tripsByKey).filter(x => cab === 'Semua cabang' || x.cab === cab);
    const held = nts.filter(x => x.kirim !== 'YA').length;
    const gaM = S && S.ga && S.ga.months && S.ga.months.length ? S.ga.months[S.ga.months.length - 1] : null;
    const ga = gaM ? VT.gaAnalysis(S.ga, gaM, P, res.master, S.ga.log) : null;
    const trendTxt = tr.length >= 2 ? (tr[tr.length - 1].fin >= tr[0].fin ? T(lang, 'naik', 'up') : T(lang, 'turun', 'down')) + ' ' + n1(Math.abs(tr[tr.length - 1].fin - tr[0].fin)) + T(lang, ' poin dalam ', ' points over ') + tr.length + T(lang, ' periode terakhir', ' recent periods') : '';
    const scope = cab === 'Semua cabang' ? T(lang, 'nasional', 'nationwide') : cab;
    const headline = T(lang,
      `Skor rata-rata ${scope} ${n1(k.fin)}${delta != null ? ' (' + (delta >= 0 ? '+' : '−') + n1(Math.abs(delta)) + ' vs periode lalu)' : ''} dengan ${pc(k.ab)} pengemudi grade A–B. ${top ? 'Risiko utama: ' + COMPN[top.c][0] + ' berulang pada ' + top.rep + ' pengemudi. ' : 'Tidak ada pelanggaran berulang yang menonjol. '}${notifTot} pengemudi perlu ditindak atasan (${k.esk} eskalasi).`,
      `Average ${scope} score ${n1(k.fin)}${delta != null ? ' (' + (delta >= 0 ? '+' : '−') + n1(Math.abs(delta)) + ' vs previous period)' : ''} with ${pc(k.ab)} of drivers at grade A–B. ${top ? 'Key risk: repeated ' + COMPN[top.c][1] + ' by ' + top.rep + ' drivers. ' : 'No significant repeated violations. '}${notifTot} drivers require supervisor follow-up (${k.esk} escalations).`);
    const insights = [];
    insights.push(T(lang, `Kinerja: ${n0(k.n)} pengemudi aktif menempuh ${n0(k.km)} km; skor ${trendTxt || 'belum punya tren'}. ${k.de} orang grade D/E perlu pembinaan terstruktur.`,
      `Performance: ${n0(k.n)} active drivers covered ${n0(k.km)} km; score ${trendTxt || 'has no trend yet'}. ${k.de} drivers at grade D/E need structured coaching.`));
    if (top) insights.push(T(lang, `Pola berulang: ${COMPN[top.c][0]} (${top.rep} pengemudi, ${top.notif} di luar standar)${sec ? ' dan ' + COMPN[sec.c][0] + ' (' + sec.rep + ' pengemudi)' : ''}. Total ${k.rep} pengemudi (${pc(k.n ? k.rep / k.n : 0)}) mengulang pelanggaran yang sama dari periode lalu.`,
      `Repeat patterns: ${COMPN[top.c][1]} (${top.rep} drivers, ${top.notif} outside standard)${sec ? ' and ' + COMPN[sec.c][1] + ' (' + sec.rep + ' drivers)' : ''}. In total ${k.rep} drivers (${pc(k.n ? k.rep / k.n : 0)}) repeated the same violation as last period.`));
    insights.push(T(lang, `Overspeed: ${k.merah} pengemudi berstatus MERAH, ${k.kuning} KUNING; ${n0(k.osv)} count dihitung setelah toleransi. Night driving ${n1(k.nhPct)}% dari jam mengemudi.`,
      `Overspeed: ${k.merah} drivers RED, ${k.kuning} YELLOW; ${n0(k.osv)} counts after tolerance. Night driving is ${n1(k.nhPct)}% of driving hours.`));
    if (worst.length && cab === 'Semua cabang') insights.push(T(lang, `Cabang perlu perhatian: ${worst.map(b => b.cab + ' (' + n1(b.fin) + ')').join(', ')}. Terbaik: ${best.map(b => b.cab + ' (' + n1(b.fin) + ')').join(', ')}.`,
      `Branches needing attention: ${worst.map(b => b.cab + ' (' + n1(b.fin) + ')').join(', ')}. Best: ${best.map(b => b.cab + ' (' + n1(b.fin) + ')').join(', ')}.`));
    if (ga) insights.push(T(lang, `Data bulanan GA (${VT.monthLabel(ga.month)}): ${pc(ga.repAll)} kasus dari pengemudi yang sama dengan bulan lalu; ${ga.p1} orang Prioritas 1, ${ga.mb} orang membaik.`,
      `GA monthly data (${VT.monthLabel(ga.month, true)}): ${pc(ga.repAll)} of cases came from the same drivers as last month; ${ga.p1} drivers Priority 1, ${ga.mb} improved.`));
    const risks = [];
    if (notifTot && cab === 'Semua cabang') risks.push(T(lang, `Konsentrasi: ${pc(conc)} notifikasi berasal dari 3 cabang (${notifTop3.map(b => b.cab).join(', ')}) — masalah pengawasan lokal, bukan nasional.`,
      `Concentration: ${pc(conc)} of notifications come from 3 branches (${notifTop3.map(b => b.cab).join(', ')}) — a local supervision issue, not a national one.`));
    if (esk.length) risks.push(T(lang, `Eskalasi: ${k.esk} pengemudi tidak berubah setelah notifikasi sebelumnya${esk.length ? ', antara lain ' + esk.map(r => r.drv + ' (' + r.cab + ')').join(', ') : ''}.`,
      `Escalation: ${k.esk} drivers did not change after the previous notification${esk.length ? ', including ' + esk.map(r => r.drv + ' (' + r.cab + ')').join(', ') : ''}.`));
    if (ins || ldr) risks.push(T(lang, `Insiden: ${ins} pengemudi menanggung pengurang insiden aktif dan ${ldr} pemimpin terkena pengurang insiden cabang.`, `Incidents: ${ins} drivers carry active incident deductions and ${ldr} leaders carry branch incident deductions.`));
    if (held) risks.push(T(lang, `Notifikasi tertahan: ${held} email belum terkirim karena email atasan belum ada di master karyawan.`, `Held notifications: ${held} emails cannot be sent because the supervisor email is missing in the employee master.`));
    if (res.dq.generic) risks.push(T(lang, `Akuntabilitas data: ${n0(res.dq.generic)} trip memakai ID umum (DRIVER…/MASTER…, ${res.dq.genericIds} ID) sehingga tidak masuk skor personal.`, `Data accountability: ${n0(res.dq.generic)} trips use shared IDs (DRIVER…/MASTER…, ${res.dq.genericIds} IDs) and cannot be scored personally.`));
    const actions = [];
    if (notifTot) actions.push([T(lang, `Tindak lanjuti ${notifTot} notifikasi atasan (coaching empat mata + target 2 minggu); ${k.esk} eskalasi ditinjau langsung oleh Branch Head`, `Follow up ${notifTot} supervisor notifications (one-to-one coaching + 2-week target); ${k.esk} escalations reviewed directly by the Branch Head`), T(lang, 'Atasan langsung / Branch Head', 'Direct supervisor / Branch Head'), T(lang, `≤ ${P.DEADLINE} hari kerja`, `≤ ${P.DEADLINE} working days`)]);
    if (top) actions.push([T(lang, PROG[top.c][0], PROG[top.c][1]), T(lang, 'QSHE + Branch Head', 'QSHE + Branch Head'), T(lang, 'Periode berikut', 'Next period')]);
    if (worst.length && cab === 'Semua cabang') actions.push([T(lang, `Kunjungan & coaching terfokus QSHE ke ${worst.slice(0, 2).map(b => b.cab).join(' dan ')}; tinjau beban rute dan jadwal`, `QSHE focused visit & coaching at ${worst.slice(0, 2).map(b => b.cab).join(' and ')}; review route load and scheduling`), 'QSHE HO', T(lang, '2 minggu', '2 weeks')]);
    if (held || res.dq.generic) actions.push([T(lang, `Tata kelola data: lengkapi email atasan (${held} tertahan) dan tutup ID umum (${res.dq.genericIds} ID) dengan DID tag pribadi`, `Data governance: complete supervisor emails (${held} held) and close shared IDs (${res.dq.genericIds} IDs) with personal DID tags`), T(lang, 'HR + Tim VT GA', 'HR + GA VT team'), T(lang, 'Sebelum periode berikut', 'Before next period')]);
    if (best.length && cab === 'Semua cabang') actions.push([T(lang, `Apresiasi terbuka untuk ${best.map(b => b.cab).join(' dan ')} serta pengemudi grade A tanpa pelanggaran berulang`, `Publicly recognise ${best.map(b => b.cab).join(' and ')} and grade A drivers without repeated violations`), T(lang, 'Manajemen', 'Management'), T(lang, 'Rapat berikut', 'Next meeting')]);
    const target = T(lang, `Target periode berikut: skor rata-rata ≥ ${n1(Math.min(100, (k.fin || 0) + 1))}, notifikasi ≤ ${Math.max(0, Math.floor(notifTot * 0.8))}, eskalasi turun ke ${Math.max(0, Math.floor(k.esk / 2))}.`,
      `Next period targets: average score ≥ ${n1(Math.min(100, (k.fin || 0) + 1))}, notifications ≤ ${Math.max(0, Math.floor(notifTot * 0.8))}, escalations down to ${Math.max(0, Math.floor(k.esk / 2))}.`);
    return { lang, title: T(lang, 'Executive report', 'Executive report'), period: p ? p.lbl : '', headline, insights, risks, actions, target,
      labels: { ins: T(lang, 'Kondisi', 'Situation'), risk: T(lang, 'Risiko strategis', 'Strategic risks'), act: T(lang, 'Keputusan & aksi strategis', 'Strategic decisions & actions'), who: T(lang, 'PIC', 'Owner'), when: T(lang, 'Tenggat', 'Timing') } };
  }
  const e = s => String(s == null ? '' : s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  function execHtml(x) {
    return `<div class="head">${e(x.headline)}</div><div class="cols">
      <div><h5>${e(x.labels.ins)}</h5><ul>${x.insights.map(i => `<li>${e(i)}</li>`).join('')}</ul></div>
      <div><h5>${e(x.labels.risk)}</h5><ul>${x.risks.map(i => `<li>${e(i)}</li>`).join('') || '<li>–</li>'}</ul></div>
      <div><h5>${e(x.labels.act)}</h5><ul>${x.actions.map(a => `<li>${e(a[0])} <span class="muted small">(${e(x.labels.who)}: ${e(a[1])} · ${e(a[2])})</span></li>`).join('')}</ul></div></div>
      <p class="note" style="margin-top:10px"><b>${e(x.target)}</b></p>`;
  }
  function execText(x) {
    return [x.title.toUpperCase() + ' – ' + x.period, '', x.headline, '', x.labels.ins.toUpperCase(), ...x.insights.map(i => '• ' + i), '', x.labels.risk.toUpperCase(), ...x.risks.map(i => '• ' + i), '',
      x.labels.act.toUpperCase(), ...x.actions.map((a, i) => (i + 1) + '. ' + a[0] + ' (' + x.labels.who + ': ' + a[1] + ' · ' + a[2] + ')'), '', x.target].join('\n');
  }

  /* ------------------------------------------------------------ laporan bilingual */
  /** Susun isi laporan (struktur netral) → dipakai untuk HTML & DOCX */
  function buildReport(res, pid, S) {
    const P = res.P; const p = res.periods.find(x => x.pid === pid);
    const k = VT.kpis(res, pid, 'Semua cabang');
    const exI = exec(res, pid, 'Semua cabang', 'id', S), exE = exec(res, pid, 'Semua cabang', 'en', S);
    const fyP = VT.fyPids(res, pid);
    const first = res.periods[0];
    const fair = VT.fairness(res, pid);
    const br = VT.recapBranches(res, pid, (S.inputs.factors || {}).insiden);
    const brS = br.filter(b => b.drivers >= P.MIN_BRANCH).sort((a, b) => b.fin - a.fin);
    const leaders = VT.recapLeaders(res, pid, P);
    const nts = VT.buildNotifications(res, pid, P, res.master, res.tripsByKey);
    const tr = VT.trend(res, pid, 'Semua cabang', 6);
    const corr = (S.inputs.corrections || []);
    const F = S.inputs.factors || {};
    const gaM = S.ga.months && S.ga.months.length ? S.ga.months[S.ga.months.length - 1] : null;
    const ga = gaM ? VT.gaAnalysis(S.ga, gaM, P, res.master, S.ga.log) : null;
    const sec = [];
    const B = (id, en) => ({ id, en });
    sec.push({ h: B('Ringkasan eksekutif', 'Executive summary'), p: [B(exI.headline, exE.headline)], list: exI.insights.map((x, i) => B(x, exE.insights[i])) });
    sec.push({ h: B('Ruang lingkup & sumber data', 'Scope & data sources'), p: [
      B(`Laporan ini mencakup periode 2-mingguan ${p.lbl} (${VT.dtxt(p.start)} – ${VT.dtxt(p.end)}) dan rekap FY${p.fy} sebanyak ${fyP.length} periode. Data yang tersimpan: ${f0(res.tripsUsed.length)} trip Driving dari ${res.periods.length} periode (${VT.dtxt(first.start)} s.d. ${VT.dtxt(res.periods[res.periods.length - 1].end)}), sumber laporan "Data Score Behavior" Cartrack yang diolah di perangkat pengguna.`,
        `This report covers the bi-weekly period ${p.lbl} (${VT.dtxtEN(p.start)} – ${VT.dtxtEN(p.end)}) and the FY${p.fy} recap of ${fyP.length} periods. Stored data: ${en0(res.tripsUsed.length)} driving trips across ${res.periods.length} periods (${VT.dtxtEN(first.start)} to ${VT.dtxtEN(res.periods[res.periods.length - 1].end)}), sourced from the Cartrack "Data Score Behavior" report and processed on the user's device.`),
      B('Hanya trip berstatus Driving dengan DID tag personal yang dinilai. Trip Idling, ID umum (DRIVER…/MASTER…), dan ID yang ditandai bukan pengemudi personal tidak masuk skor perorangan.',
        'Only Driving trips with a personal DID tag are scored. Idling trips, shared IDs (DRIVER…/MASTER…), and IDs flagged as non-personal are excluded from individual scores.')] });
    sec.push({ h: B('Metodologi skoring', 'Scoring methodology'), p: [
      B(`Skor dihitung per pengemudi per periode 2 minggu. Setiap pengemudi dibandingkan dengan rata-rata kelompok pada band travel load yang sama (L1 <${P.BAND_LO[1]} km, L2 ${P.BAND_LO[1]}–${P.BAND_LO[2]}, L3 ${P.BAND_LO[2]}–${P.BAND_LO[3]}, L4 ${P.BAND_LO[3]}–${P.BAND_LO[4]}, L5 ≥${P.BAND_LO[4]} km). Skor komponen = 100 × ${fx(P.BASE, 2)}^rasio, dengan km/jam bayangan ${P.KKM} km / ${P.KH} jam agar adil bagi yang jarang mengemudi.`,
        `Scores are computed per driver per two-week period. Each driver is compared with the peer average of the same travel-load band (L1 <${P.BAND_LO[1]} km, L2 ${P.BAND_LO[1]}–${P.BAND_LO[2]}, L3 ${P.BAND_LO[2]}–${P.BAND_LO[3]}, L4 ${P.BAND_LO[3]}–${P.BAND_LO[4]}, L5 ≥${P.BAND_LO[4]} km). Component score = 100 × ${P.BASE}^ratio, with credibility exposure of ${P.KKM} km / ${P.KH} hours so infrequent drivers are treated fairly.`),
      B(`Bobot: overspeed ${P.W.spd * 100}%, fatigue ${P.W.fat * 100}%, night driving ${P.W.night * 100}%, harsh braking ${P.W.brk * 100}%, harsh acceleration ${P.W.acc * 100}%, harsh cornering ${P.W.cor * 100}%. Overspeed ≤${P.OSTOL} count per trip ditoleransi; trip ≥${P.OSSUS} count = ngebut beruntun (MERAH). Night driving = trip yang mulai ${String(P.NSTART).padStart(2, '0')}:00–${String((P.NEND + 23) % 24).padStart(2, '0')}:59.`,
        `Weights: overspeed ${P.W.spd * 100}%, fatigue ${P.W.fat * 100}%, night driving ${P.W.night * 100}%, harsh braking ${P.W.brk * 100}%, harsh acceleration ${P.W.acc * 100}%, harsh cornering ${P.W.cor * 100}%. Overspeed ≤${P.OSTOL} counts per trip is tolerated; a trip with ≥${P.OSSUS} counts is sustained speeding (RED). Night driving = trips starting ${String(P.NSTART).padStart(2, '0')}:00–${String((P.NEND + 23) % 24).padStart(2, '0')}:59.`),
      B(`Pelanggaran yang sama dalam 2 periode berturut-turut mengurangi skor ${P.PTET} poin (tetap) atau ${P.PBUR} poin (makin buruk); overspeed dikali ${P.OSMULT}, berulang ≥3 periode dikali ${P.STREAK}, maksimum ${P.REPCAP} poin. Insiden mengurangi skor pengemudi dan pemimpin cabang (memudar linier ${P.DECAY} hari). Grade: A ≥${P.GRADE[0]}, B ≥${P.GRADE[1]}, C ≥${P.GRADE[2]}, D ≥${P.GRADE[3]}, E di bawahnya.`,
        `Repeating the same violation in 2 consecutive periods deducts ${P.PTET} points (unchanged) or ${P.PBUR} points (worse); overspeed ×${P.OSMULT}, repeating ≥3 periods ×${P.STREAK}, capped at ${P.REPCAP} points. Incidents deduct from the driver and branch leaders (linear fade over ${P.DECAY} days). Grades: A ≥${P.GRADE[0]}, B ≥${P.GRADE[1]}, C ≥${P.GRADE[2]}, D ≥${P.GRADE[3]}, E below.`)] });
    sec.push({ h: B('Hasil utama periode', 'Key results of the period'), table: { head: [B('Indikator', 'Indicator'), B('Nilai', 'Value'), B('Catatan', 'Note')], rows: [
      [B('Skor rata-rata (metode baru)', 'Average score (new method)'), f1(k.fin), B(`metode lama ${f1(k.old)}; periode lalu ${f1(k.finPrev)}`, `old method ${en1(k.old)}; previous period ${en1(k.finPrev)}`)],
      [B('Pengemudi aktif', 'Active drivers'), f0(k.n), B(`${f0(k.km)} km`, `${en0(k.km)} km`)],
      [B('Grade A–B', 'Grade A–B'), pc(k.ab), B(`${k.de} orang grade D/E`, `${k.de} drivers at grade D/E`)],
      [B('Notifikasi atasan', 'Supervisor notifications'), f0(k.notif), B(`${k.esk} eskalasi`, `${k.esk} escalations`)],
      [B('Pelanggaran berulang', 'Repeated violations'), f0(k.rep), B('pengemudi', 'drivers')],
      [B('Overspeed MERAH / KUNING', 'Overspeed RED / YELLOW'), k.merah + ' / ' + k.kuning, B(`${f0(k.osv)} count dihitung`, `${en0(k.osv)} counted`)],
      [B('Night driving', 'Night driving'), f0(k.nh) + ' jam', B(`${f1(k.nhPct)}% jam mengemudi`, `${en1(k.nhPct)}% of driving hours`)],
      [B('Harsh braking / 1.000 km', 'Harsh braking / 1,000 km'), fx(k.brk, 2), B(`periode lalu ${k.brkPrev == null ? '–' : fx(k.brkPrev, 2)}`, `previous ${k.brkPrev == null ? '–' : k.brkPrev.toFixed(2)}`)]] } });
    sec.push({ h: B('Tren 6 periode', 'Six-period trend'), table: { head: [B('Periode', 'Period'), B('Skor lama', 'Old score'), B('Skor baru', 'New score'), B('Overspeed/1.000 km', 'Overspeed/1,000 km'), B('Braking/1.000 km', 'Braking/1,000 km'), B('Pengemudi', 'Drivers')],
      rows: tr.map(x => [x.label, f1(x.old), f1(x.fin), fx(x.osv, 2), fx(x.brk, 2), f0(x.n)]) } });
    sec.push({ h: B('Keadilan travel load', 'Travel-load fairness'), p: [B(`Selisih skor baru antara band L5 dan L1 = ${fx(fair[4].fin - fair[0].fin, 1)} poin (metode lama ${fx(fair[4].old - fair[0].old, 1)}). Skor yang adil membuat pengemudi beban tinggi tidak dirugikan hanya karena jaraknya.`,
      `New-score gap between band L5 and L1 = ${en1(fair[4].fin - fair[0].fin)} points (old method ${en1(fair[4].old - fair[0].old)}). A fair score means high-load drivers are not penalised merely for distance.`)],
      table: { head: [B('Band', 'Band'), B('Driver-periode', 'Driver-periods'), B('KM rata-rata', 'Avg km'), B('Skor lama', 'Old score'), B('Skor baru', 'New score'), B('% D/E', '% D/E')], rows: fair.map(b => [b.band, f0(b.n), f0(b.km), f1(b.old), f1(b.fin), pc(b.de)]) } });
    sec.push({ h: B('Pelanggaran berulang & notifikasi atasan', 'Repeated violations & supervisor notifications'), p: [B(`${nts.length} pengemudi memenuhi syarat notifikasi (berulang 2 periode dan di luar standar), ${nts.filter(x => x.level === 'ESKALASI').length} di antaranya eskalasi. ${nts.filter(x => x.kirim === 'YA').length} siap dikirim; ${nts.filter(x => x.kirim !== 'YA').length} tertahan karena email atasan belum ada.`,
      `${nts.length} drivers meet the notification rule (repeated for 2 periods and outside standard), ${nts.filter(x => x.level === 'ESKALASI').length} of them escalations. ${nts.filter(x => x.kirim === 'YA').length} are ready to send; ${nts.filter(x => x.kirim !== 'YA').length} are held because the supervisor email is missing.`)],
      table: { head: [B('Nama', 'Name'), B('Cabang', 'Branch'), B('Level', 'Level'), B('Skor', 'Score'), B('Komponen', 'Component')], rows: nts.slice(0, 40).map(x => [x.nama, x.cab, x.level === 'ESKALASI' ? 'ESKALASI / ESCALATION' : 'PERHATIAN / ATTENTION', f1(x.row.FIN), VT.REPC.filter(c => x.row['N_' + c]).map(c => COMPN[c][0]).join(', ')]) } });
    sec.push({ h: B('Kinerja cabang (FY berjalan)', 'Branch performance (current FY)'), table: { head: [B('Cabang', 'Branch'), B('Pengemudi', 'Drivers'), B('Skor lama', 'Old score'), B('Skor baru', 'New score'), B('% D/E', '% D/E'), B('Notifikasi', 'Notifications')],
      rows: brS.map(b => [b.cab, f0(b.drivers), f1(b.old), f1(b.fin), pc(b.de), f0(b.notif)]) } });
    if (leaders.length) sec.push({ h: B('Akuntabilitas pemimpin', 'Leader accountability'), table: { head: [B('Nama', 'Name'), B('Jabatan', 'Role'), B('Cabang', 'Branch'), B('Skor pemimpin', 'Leader score'), B('Pengurang insiden cabang', 'Branch incident deduction')], rows: leaders.map(l => [l.drv, l.jab, l.cab, f1(l.score), f1(l.ldr)]) } });
    if (ga) sec.push({ h: B('Laporan bulanan GA & concern', 'GA monthly report & concerns'), p: [
      B(`Bulan ${VT.monthLabel(ga.month)} dibanding ${VT.monthLabel(ga.prevMonth)}: overspeed ${ga.cat[0].prv} → ${ga.cat[0].cur}, fatigue ${ga.cat[1].prv} → ${ga.cat[1].cur}, night driving ${ga.cat[2].prv} → ${ga.cat[2].cur} kasus. ${pc(ga.repAll)} kasus datang dari pengemudi yang sama dengan bulan lalu. ${ga.p1} orang Prioritas 1, ${ga.p2} Prioritas 2, ${ga.mb} membaik.${ga.topOs && ga.topOs.os_cur ? ' Overspeed tertinggi dari satu orang: ' + ga.topOs.os_cur + ' kasus (' + pc(ga.cat[0].cur ? ga.topOs.os_cur / ga.cat[0].cur : 0) + ').' : ''}`,
        `${VT.monthLabel(ga.month, true)} versus ${VT.monthLabel(ga.prevMonth, true)}: overspeed ${ga.cat[0].prv} → ${ga.cat[0].cur}, fatigue ${ga.cat[1].prv} → ${ga.cat[1].cur}, night driving ${ga.cat[2].prv} → ${ga.cat[2].cur} cases. ${pc(ga.repAll)} of cases came from the same drivers as last month. ${ga.p1} drivers Priority 1, ${ga.p2} Priority 2, ${ga.mb} improved.${ga.topOs && ga.topOs.os_cur ? ' Highest overspeed from a single driver: ' + ga.topOs.os_cur + ' cases (' + pc(ga.cat[0].cur ? ga.topOs.os_cur / ga.cat[0].cur : 0) + ').' : ''}`),
      B('Data GA belum memuat km dan jam mengemudi per orang, sehingga perbandingan antar cabang belum adil; definisi kasus, batas kecepatan, dan jam malam perlu disepakati bersama GA, Tim VT, dan QSHE.',
        'GA data does not yet include per-driver km and driving hours, so branch comparisons are not yet fair; the case definition, speed limit, and night-hour window need to be agreed between GA, the VT team, and QSHE.')] });
    const ca = corr.filter(c => c.status === 'Disetujui');
    sec.push({ h: B('Faktor koreksi & faktor lain yang diterapkan', 'Corrections & other factors applied'), p: [B(`${ca.length} faktor koreksi disetujui QSHE diterapkan pada perhitungan (${corr.filter(c => c.status === 'Diajukan').length} masih menunggu, ${corr.filter(c => c.status === 'Ditolak').length} ditolak). Faktor lain tercatat: ${(F.insiden || []).length} insiden, ${(F.promosi || []).length} promosi, ${(F.mutasi || []).length} mutasi, ${(F.cuti || []).length} cuti, ${(F.keluar || []).length} keluar, ${(F.pelatihan || []).length} pelatihan, ${(F.nondriver || []).length} ID bukan pengemudi personal.`,
      `${ca.length} QSHE-approved corrections are applied (${corr.filter(c => c.status === 'Diajukan').length} pending, ${corr.filter(c => c.status === 'Ditolak').length} rejected). Other factors recorded: ${(F.insiden || []).length} incidents, ${(F.promosi || []).length} promotions, ${(F.mutasi || []).length} transfers, ${(F.cuti || []).length} leaves, ${(F.keluar || []).length} leavers, ${(F.pelatihan || []).length} trainings, ${(F.nondriver || []).length} non-personal IDs.`)],
      table: ca.length ? { head: [B('Nama', 'Name'), B('Alasan', 'Reason'), B('Aksi', 'Action'), B('Trip terdampak', 'Trips affected'), B('Bukti', 'Evidence')], rows: ca.map(c => { const r = VT.CORR_REASONS.find(x => x.id === c.alasan) || {}; const a = res.corAudit[c.id] || {}; return [c.nama, B(r.label || c.alasan, r.en || c.alasan), VT.CORR_ACTIONS[c.aksi] || c.aksi, f0(a.trips || 0), c.bukti || '–']; }) } : null });
    sec.push({ h: B('Kualitas data & keterbatasan', 'Data quality & limitations'), list: [
      B(`${f0(res.dq.generic)} trip Driving memakai ID umum (${res.dq.genericIds} ID) dengan ${f0(res.dq.genericKm)} km — tidak bisa ditelusuri ke orang.`, `${en0(res.dq.generic)} driving trips use shared IDs (${res.dq.genericIds} IDs) covering ${en0(res.dq.genericKm)} km — not traceable to a person.`),
      B(`${f0(res.dq.v120)} trip punya kecepatan rata-rata >120 km/jam (indikasi lonjakan GPS); ${f0(res.dq.v120os)} di antaranya ber-overspeed. Ajukan faktor koreksi bila terverifikasi.`, `${en0(res.dq.v120)} trips have an average speed >120 km/h (possible GPS spikes); ${en0(res.dq.v120os)} of them include overspeed. Submit a correction when verified.`),
      B('Data VT tidak memuat durasi overspeed; toleransi count per trip dipakai sebagai kriteria keberterimaan pengganti aturan "<10 detik".', 'VT data has no overspeed duration; the per-trip count tolerance is used as the acceptance criterion in place of the "<10 seconds" rule.'),
      B('Harsh cornering tidak dihitung sebagai pelanggaran berulang karena sangat sensitif terhadap sensor dan kondisi jalan.', 'Harsh cornering is excluded from the repeat rule because it is highly sensitive to sensors and road conditions.')] });
    sec.push({ h: B('Rekomendasi strategis & rencana tindak lanjut', 'Strategic recommendations & action plan'), table: { head: [B('Tindakan', 'Action'), B('PIC', 'Owner'), B('Tenggat', 'Timing')], rows: exI.actions.map((a, i) => [B(a[0], exE.actions[i][0]), B(a[1], exE.actions[i][1]), B(a[2], exE.actions[i][2])]) }, p: [B(exI.target, exE.target)] });
    return { p, sec, title: B('Laporan Vehicle Telematic Score and Report', 'Vehicle Telematic Score and Report'), sub: B(`Periode ${p.lbl} · ${VT.dtxt(p.start)} – ${VT.dtxt(p.end)}`, `Period ${p.lbl} · ${VT.dtxtEN(p.start)} – ${VT.dtxtEN(p.end)}`) };
  }
  const cellTxt = (c, lang) => (c && typeof c === 'object' ? (lang === 'en' ? c.en : c.id) : c);
  function reportHtml(R, mode) {
    // mode: 'bi' (dua bahasa), 'id', 'en'
    const L = mode === 'bi' ? ['id', 'en'] : [mode];
    let h = `<article class="report"><header style="border-bottom:4px solid #f7941d;padding-bottom:10px;margin-bottom:14px"><div class="muted small">QSHE DEPARTMENT · PT HEXINDO ADIPERKASA TBK</div>
      <h2 style="font:700 24px var(--disp);margin:6px 0">${L.map(l => e(cellTxt(R.title, l))).join('<br><span style="font-weight:500;font-size:18px;color:#6b6b6b">') + (L.length > 1 ? '</span>' : '')}</h2>
      <div>${L.map(l => e(cellTxt(R.sub, l))).join(' · ')}</div></header>`;
    R.sec.forEach((s, i) => {
      h += `<h3 class="sec">${i + 1}. ${L.map(l => e(cellTxt(s.h, l))).join(' / ')}</h3>`;
      (s.p || []).forEach(p => { L.forEach((l, j) => { h += `<p style="${j ? 'color:#555;font-style:italic' : ''};line-height:1.6;margin:4px 0 8px">${e(cellTxt(p, l))}</p>`; }); });
      if (s.list) h += '<ul>' + s.list.map(x => L.map((l, j) => `<li style="${j ? 'color:#555;font-style:italic;list-style:circle' : ''}">${e(cellTxt(x, l))}</li>`).join('')).join('') + '</ul>';
      if (s.table) h += `<div class="tbox free"><table class="t"><thead><tr>${s.table.head.map(c => `<th>${L.map(l => e(cellTxt(c, l))).join('<br>')}</th>`).join('')}</tr></thead><tbody>${s.table.rows.map(r => `<tr>${r.map(c => `<td>${L.map(l => e(cellTxt(c, l))).filter((v, j, a) => a.indexOf(v) === j).join('<br><i class="muted">')}</td>`).join('')}</tr>`).join('')}</tbody></table></div>`;
    });
    h += `<p class="note" style="margin-top:20px">Disusun otomatis oleh aplikasi Vehicle Telematic Score and Report v${VT.VERSION} pada ${new Date().toLocaleString('id-ID')}. / Generated automatically by Vehicle Telematic Score and Report v${VT.VERSION}.</p></article>`;
    return h;
  }
  async function reportDocx(R, mode) {
    const D = window.docx; if (!D) throw new Error('Pustaka docx belum termuat');
    const L = mode === 'bi' ? ['id', 'en'] : [mode];
    const OR = 'F7941D';
    const run = (t, o) => new D.TextRun(Object.assign({ text: String(t == null ? '' : t), font: 'Segoe UI', size: 20 }, o || {}));
    const para = (t, o, po) => new D.Paragraph(Object.assign({ children: [run(t, o)], spacing: { after: 120, line: 300 } }, po || {}));
    const kids = [];
    kids.push(para('QSHE DEPARTMENT · PT HEXINDO ADIPERKASA TBK', { color: '6B6B6B', size: 18, bold: true }));
    L.forEach((l, j) => kids.push(para(cellTxt(R.title, l), { bold: true, size: j ? 30 : 40, color: j ? '6B6B6B' : '111111', font: 'Poppins' }, { spacing: { after: 60 } })));
    kids.push(para(L.map(l => cellTxt(R.sub, l)).join('  ·  '), { size: 20 }, { border: { bottom: { color: OR, size: 18, style: D.BorderStyle.SINGLE, space: 6 } }, spacing: { after: 240 } }));
    R.sec.forEach((s, i) => {
      kids.push(new D.Paragraph({ heading: D.HeadingLevel.HEADING_2, children: [run((i + 1) + '. ' + L.map(l => cellTxt(s.h, l)).join(' / '), { bold: true, size: 26, color: '111111' })], spacing: { before: 280, after: 120 }, border: { left: { color: OR, size: 24, style: D.BorderStyle.SINGLE, space: 8 } } }));
      (s.p || []).forEach(p => L.forEach((l, j) => kids.push(para(cellTxt(p, l), j ? { italics: true, color: '555555' } : {}))));
      (s.list || []).forEach(x => L.forEach((l, j) => kids.push(new D.Paragraph({ bullet: { level: j }, children: [run(cellTxt(x, l), j ? { italics: true, color: '555555' } : {})], spacing: { after: 80 } }))));
      if (s.table && s.table.rows.length) {
        const ncol = s.table.head.length;
        const mk = (c, hdr) => new D.TableCell({ shading: hdr ? { type: D.ShadingType.CLEAR, color: 'auto', fill: '111111' } : undefined, margins: { top: 60, bottom: 60, left: 90, right: 90 },
          children: L.map(l => cellTxt(c, l)).filter((v, j, a) => a.indexOf(v) === j).map((v, j) => new D.Paragraph({ children: [run(v, { size: 17, bold: hdr && j === 0, color: hdr ? 'FFFFFF' : j ? '555555' : '111111', italics: j > 0 })] })) });
        kids.push(new D.Table({ width: { size: 100, type: D.WidthType.PERCENTAGE }, rows: [new D.TableRow({ tableHeader: true, children: s.table.head.map(c => mk(c, true)) }), ...s.table.rows.map(r => new D.TableRow({ children: r.slice(0, ncol).map(c => mk(c, false)) }))] }));
        kids.push(para(''));
      }
    });
    kids.push(para('Disusun otomatis oleh Vehicle Telematic Score and Report v' + VT.VERSION + ' · Generated automatically by Vehicle Telematic Score and Report.', { size: 16, color: '9E9E9E', italics: true }, { spacing: { before: 300 } }));
    const doc = new D.Document({
      creator: 'QSHE Department – PT Hexindo Adiperkasa Tbk', title: cellTxt(R.title, 'id'), styles: { default: { document: { run: { font: 'Segoe UI', size: 20 } } } },
      sections: [{ properties: { page: { margin: { top: 1000, bottom: 1000, left: 1100, right: 1100 } } },
        headers: { default: new D.Header({ children: [para('Vehicle Telematic Score and Report  ·  ' + R.p.lbl, { size: 16, color: '9E9E9E' })] }) },
        footers: { default: new D.Footer({ children: [new D.Paragraph({ alignment: D.AlignmentType.RIGHT, children: [run('QSHE Department · Halaman / Page ', { size: 16, color: '9E9E9E' }), new D.TextRun({ children: [D.PageNumber.CURRENT], size: 16, color: '9E9E9E', font: 'Segoe UI' })] })] }) },
        children: kids }]
    });
    return D.Packer.toBlob(doc);
  }
  return { exec, execHtml, execText, buildReport, reportHtml, reportDocx };
})();
