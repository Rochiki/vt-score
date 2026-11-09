/* =====================================================================
   app.js — kerangka aplikasi: state, penyimpanan, router, komponen UI,
   dan halaman DASHBOARD (format sama dengan sheet 00_DASHBOARD Excel).
   ===================================================================== */
'use strict';
const $ = (s, el) => (el || document).querySelector(s);
const $$ = (s, el) => [...(el || document).querySelectorAll(s)];
const esc = s => String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const fx = VT.fx;
const f0 = v => fx(v, 0), f1 = v => fx(v, 1), f2 = v => fx(v, 2);
const pc = (v, dp) => (v == null || !isFinite(v) ? '–' : fx(v * 100, dp || 0) + '%');
const sg = (v, dp) => (v == null || !isFinite(v) ? '–' : (v > 0 ? '+' : '') + fx(v, dp == null ? 1 : dp));
const uid = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 7);
const today = () => new Date().toISOString().slice(0, 10);

/* ------------------------------------------------------------ state */
const S = {
  trips: [], pstats: {}, res: null,
  inputs: { params: {}, corrections: [], factors: {}, master: [], lastPid: null, scoreFromPid: null, imports: [] },
  ga: { rows: [], months: [], info: {}, log: [], concerns: {} },
  settings: { ai: { provider: 'claude', mode: 'web', apiKey: '', model: 'claude-sonnet-5-5', anonymize: true }, lang: 'id' },
  ui: { pid: null, cab: 'Semua cabang' },
  charts: {}, dirty: false
};
VT.FACTOR_TYPES.forEach(f => { S.inputs.factors[f.id] = []; });

/* ------------------------------------------------------------ komponen kecil */
function toast(msg, ms) {
  const t = document.createElement('div'); t.className = 'toast'; t.setAttribute('role', 'status'); t.textContent = msg;
  document.body.appendChild(t); setTimeout(() => t.remove(), ms || 3200);
}
function busy(msg) {
  let b = $('#busy');
  if (msg === false) { if (b) b.remove(); return; }
  if (!b) { b = document.createElement('div'); b.id = 'busy'; b.className = 'busy'; b.innerHTML = '<div></div>'; document.body.appendChild(b); }
  b.firstChild.textContent = msg || 'Memproses…';
}
function modal(title, html, opts) {
  opts = opts || {};
  const m = document.createElement('div'); m.className = 'modal'; m.setAttribute('role', 'dialog'); m.setAttribute('aria-modal', 'true');
  m.innerHTML = `<div class="box" style="${opts.wide ? 'width:min(1100px,100%)' : ''}"><header><b>${esc(title)}</b><button type="button" aria-label="Tutup">×</button></header><div class="bd"></div></div>`;
  const bd = $('.bd', m); if (typeof html === 'string') bd.innerHTML = html; else bd.appendChild(html);
  const close = () => { m.remove(); if (opts.onClose) opts.onClose(); document.removeEventListener('keydown', key); };
  const key = e => { if (e.key === 'Escape') close(); };
  $('header button', m).onclick = close; m.addEventListener('click', e => { if (e.target === m) close(); });
  document.addEventListener('keydown', key);
  document.body.appendChild(m); return { el: m, body: bd, close };
}
function confirmBox(msg, okLabel) {
  return new Promise(res => {
    const md = modal('Konfirmasi', `<p>${esc(msg)}</p><div class="tools"><span class="sp"></span><button class="btn" data-x="0">Batal</button><button class="btn pri" data-x="1">${esc(okLabel || 'Ya, lanjutkan')}</button></div>`, { onClose: () => res(false) });
    $$('[data-x]', md.body).forEach(b => b.onclick = () => { const v = b.dataset.x === '1'; md.el.remove(); res(v); });
  });
}
function download(blob, name) {
  const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = name; document.body.appendChild(a); a.click();
  setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1500);
}
async function copyText(t) {
  try { await navigator.clipboard.writeText(t); toast('Tersalin ke clipboard.'); return true; }
  catch (e) { const ta = document.createElement('textarea'); ta.value = t; document.body.appendChild(ta); ta.select(); let ok = false; try { ok = document.execCommand('copy'); } catch (x) { /* */ } ta.remove(); toast(ok ? 'Tersalin ke clipboard.' : 'Tidak bisa menyalin otomatis — pilih teks lalu salin manual.'); return ok; }
}
const loaded = {};
function loadLib(src) {
  if (loaded[src]) return loaded[src];
  loaded[src] = new Promise((res, rej) => { const s = document.createElement('script'); s.src = src; s.onload = res; s.onerror = () => rej(new Error('Gagal memuat ' + src)); document.head.appendChild(s); });
  return loaded[src];
}
function gradeChip(g) { return g && 'ABCDE'.includes(g) ? `<span class="g g${g}">${g}</span>` : esc(g || '–'); }
function ossChip(s) { if (!s) return '–'; const k = s.split(' ')[0]; return `<span class="st ${k}">${esc(s)}</span>`; }

/* ------------------------------------------------------------ tabel generik */
/** cols: [{k, l, f?(row)->html, v?(row)->sort value, cls?, x?(row)->export value}] */
function table(host, cols, rows, opts) {
  opts = opts || {};
  const st = { q: '', sort: opts.sort || null, dir: opts.dir || 1, limit: opts.limit || 500 };
  host.innerHTML = `<div class="tools" ${opts.search === false && !opts.tools && !opts.exportName ? 'hidden' : ''}>${opts.search === false ? '' : '<input type="search" placeholder="Cari…" aria-label="Cari">'}${opts.tools || ''}<span class="sp"></span><span class="muted small cnt"></span>${opts.exportName ? '<button class="btn sm" data-exp>Ekspor Excel</button>' : ''}</div><div class="tbox ${opts.free ? 'free' : ''}"><table class="t"><thead><tr>${cols.map((c, i) => `<th class="${c.cls || ''} ${c.nosort ? '' : 'sort'}" data-i="${i}" scope="col">${esc(c.l)}</th>`).join('')}</tr></thead><tbody></tbody></table></div><div class="more"></div>`;
  const tb = $('tbody', host), cnt = $('.cnt', host), more = $('.more', host);
  const val = (c, r) => (c.v ? c.v(r) : r[c.k]);
  function draw() {
    let a = rows;
    if (st.q) { const q = st.q.toLowerCase(); a = a.filter(r => cols.some(c => String(c.x ? c.x(r) : val(c, r) == null ? '' : val(c, r)).toLowerCase().indexOf(q) >= 0)); }
    if (st.sort != null) { const c = cols[st.sort]; a = a.slice().sort((x, y) => { const p = val(c, x), q2 = val(c, y); if (p == null && q2 == null) return 0; if (p == null) return 1; if (q2 == null) return -1; return (typeof p === 'number' && typeof q2 === 'number' ? p - q2 : String(p).localeCompare(String(q2), 'id')) * st.dir; }); }
    cnt.textContent = a.length + ' baris';
    const show = a.slice(0, st.limit);
    tb.innerHTML = show.map((r, i) => `<tr class="${opts.onRow ? 'click' : ''}" data-i="${rows.indexOf(r)}">${cols.map(c => `<td class="${c.cls || ''}">${c.f ? c.f(r, i) : esc(val(c, r) == null ? '' : val(c, r))}</td>`).join('')}</tr>`).join('') || `<tr><td colspan="${cols.length}" class="muted">Tidak ada data.</td></tr>`;
    more.innerHTML = a.length > st.limit ? `<div class="tools"><button class="btn sm">Tampilkan ${Math.min(500, a.length - st.limit)} baris lagi (dari ${a.length})</button></div>` : '';
    if (a.length > st.limit) $('button', more).onclick = () => { st.limit += 500; draw(); };
    $$('th', host).forEach(th => { th.removeAttribute('aria-sort'); if (+th.dataset.i === st.sort) th.setAttribute('aria-sort', st.dir > 0 ? 'ascending' : 'descending'); });
    host._filtered = a;
  }
  const qi = $('input[type=search]', host); if (qi) qi.oninput = () => { st.q = qi.value.trim(); st.limit = opts.limit || 500; draw(); };
  $$('th.sort', host).forEach(th => th.onclick = () => { const i = +th.dataset.i; if (st.sort === i) st.dir = -st.dir; else { st.sort = i; st.dir = 1; } draw(); });
  if (opts.onRow) tb.addEventListener('click', e => { const tr = e.target.closest('tr[data-i]'); if (tr && !e.target.closest('button,input,select,a')) opts.onRow(rows[+tr.dataset.i]); });
  const ex = $('[data-exp]', host); if (ex) ex.onclick = () => exportXlsx(opts.exportName, [{ name: opts.sheetName || 'Data', cols, rows: host._filtered || rows }]);
  draw();
  return { redraw: draw, host };
}
async function exportXlsx(name, sheets) {
  busy('Menyiapkan file Excel…');
  try {
    await loadLib('lib/exceljs.min.js');
    const wb = new ExcelJS.Workbook(); wb.creator = 'Vehicle Telematic Score and Report';
    for (const sh of sheets) {
      const ws = wb.addWorksheet(sh.name.slice(0, 31));
      ws.addRow(sh.cols.map(c => c.l));
      const hr = ws.getRow(1); hr.font = { bold: true, color: { argb: 'FFFFFFFF' } }; hr.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF111111' } }; hr.alignment = { vertical: 'middle', wrapText: true }; hr.height = 30;
      for (const r of sh.rows) ws.addRow(sh.cols.map(c => { let v = c.x ? c.x(r) : (c.v ? c.v(r) : r[c.k]); if (typeof v === 'number' && !isFinite(v)) v = null; return v == null ? '' : v; }));
      sh.cols.forEach((c, i) => { ws.getColumn(i + 1).width = c.w || Math.min(48, Math.max(10, String(c.l).length + 4)); if (c.nf) ws.getColumn(i + 1).numFmt = c.nf; });
      ws.views = [{ state: 'frozen', ySplit: 1 }];
      ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: sh.cols.length } };
    }
    const buf = await wb.xlsx.writeBuffer();
    download(new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), name.endsWith('.xlsx') ? name : name + '.xlsx');
  } catch (e) { toast('Ekspor gagal: ' + e.message, 6000); } finally { busy(false); }
}

/* ------------------------------------------------------------ chart helper */
function chart(id, cfg) {
  const cv = document.getElementById(id); if (!cv) return null;
  const box = cv.parentElement;
  if (S.charts[id]) { try { S.charts[id].destroy(); } catch (e) { /* */ } delete S.charts[id]; }
  const fail = msg => { const p = document.createElement('p'); p.className = 'note chart-err'; p.textContent = msg; box.querySelectorAll('.chart-err').forEach(x => x.remove()); box.appendChild(p); };
  if (typeof Chart === 'undefined') { fail('Grafik tidak bisa dimuat (pustaka grafik gagal diunduh). Muat ulang halaman dengan Ctrl+F5.'); return null; }
  try {
    cfg.options = cfg.options || {}; cfg.options.animation = false; cfg.options.responsiveAnimationDuration = 0;
    const c = new Chart(cv, cfg); S.charts[id] = c;
    c.draw(); // gambar langsung, tidak menunggu frame animasi browser
    return c;
  } catch (e) { console.error('[VT] grafik ' + id, e); fail('Grafik gagal digambar: ' + e.message); return null; }
}
// gambar ulang grafik bila tab/jendela baru terlihat atau ukurannya berubah
document.addEventListener('visibilitychange', () => { if (!document.hidden) Object.values(S.charts).forEach(c => { try { c.resize(); c.update('none'); } catch (e) { /* */ } }); });
function chartsInit() {
  if (typeof Chart === 'undefined') return;
  Chart.defaults.font.family = "'Segoe UI', 'Barlow', system-ui, sans-serif";
  Chart.defaults.font.size = 12; Chart.defaults.color = '#3a3a3a';
  Chart.defaults.plugins.legend.labels.boxWidth = 14;
  Chart.defaults.maintainAspectRatio = false;
  Chart.defaults.animation = false;
  Chart.defaults.plugins.tooltip.callbacks.label = ctx => (ctx.dataset.label ? ctx.dataset.label + ': ' : '') + (ctx.parsed.y == null ? '–' : fx(ctx.parsed.y, Math.abs(ctx.parsed.y) < 10 ? 2 : 1));
}
const gridY = { grid: { color: '#e5e5e5' }, border: { display: false } };

/* ------------------------------------------------------------ penyimpanan */
let saveT = null;
function saveInputs(recompute) {
  clearTimeout(saveT);
  saveT = setTimeout(async () => { try { await DB.set('inputs', S.inputs); } catch (e) { toast('Gagal menyimpan: ' + e.message, 6000); } }, 250);
  if (recompute) { S.dirty = true; }
}
const saveGA = () => DB.set('ga', S.ga).catch(e => toast('Gagal menyimpan data GA: ' + e.message));
const saveSettings = () => DB.set('settings', S.settings).catch(() => {});

async function loadAll() {
  const [inp, ga, st, per] = await Promise.all([DB.get('inputs'), DB.get('ga'), DB.get('settings'), DB.periods()]);
  if (inp) { Object.assign(S.inputs, inp); VT.FACTOR_TYPES.forEach(f => { S.inputs.factors[f.id] = S.inputs.factors[f.id] || []; }); }
  if (ga) Object.assign(S.ga, ga);
  if (st) { S.settings = Object.assign(S.settings, st); S.settings.ai = Object.assign({ provider: 'claude', mode: 'web', apiKey: '', model: 'claude-sonnet-5-5', anonymize: true }, st.ai || {}); }
  setTrips(per);
}
function setTrips(per) {
  S.trips = []; S.pstats = {};
  for (const p of per) { S.pstats[p.lbl] = p.stats; for (const t of p.trips) S.trips.push(t); }
}
function compute() {
  if (!S.trips.length) { S.res = null; return; }
  const t = performance.now();
  S.res = VT.compute({ trips: S.trips, pstats: S.pstats, inputs: S.inputs });
  S.dirty = false;
  if (!S.ui.pid || !S.res.periods.find(p => p.pid === S.ui.pid)) S.ui.pid = S.res.lastPid;
  console.info('[VT] hitung ulang', Math.round(performance.now() - t), 'ms', S.res.rows.length, 'baris');
}
async function recompute(msg) {
  busy(msg || 'Menghitung ulang skor…');
  await new Promise(r => setTimeout(r, 30));
  try { compute(); } catch (e) { console.error(e); toast('Perhitungan gagal: ' + e.message, 8000); } finally { busy(false); }
}
const P = () => (S.res ? S.res.P : VT.params(S.inputs.params));
const per = pid => S.res && S.res.periods.find(p => p.pid === pid);
const perLabel = pid => { const p = per(pid); return p ? p.lbl : '–'; };
const perRange = pid => { const p = per(pid); return p ? VT.dtxt(p.start) + ' – ' + VT.dtxt(p.end) : ''; };

/* ------------------------------------------------------------ worker impor */
let worker = null, wid = 0; const wcb = {};
function getWorker() {
  if (worker) return worker;
  worker = new Worker('js/worker.js');
  worker.onmessage = ev => { const m = ev.data; const cb = wcb[m.id]; if (cb) cb(m); };
  worker.onerror = e => { toast('Worker error: ' + (e.message || e), 8000); };
  return worker;
}
function importFile(file, onProgress) {
  return new Promise((res, rej) => {
    const id = ++wid;
    wcb[id] = m => { if (m.type === 'progress') onProgress && onProgress(m.pct); else { delete wcb[id]; m.type === 'done' ? res(m) : rej(new Error(m.message)); } };
    getWorker().postMessage({ id, cmd: 'import', file });
  });
}

/* ------------------------------------------------------------ router & layout */
const PAGES = {};
const NAV = [
  ['MULAI DI SINI', [['panduan', 'Panduan'], ['unggah', 'Unggah data']]],
  ['LAPORAN', [['dashboard', 'Dashboard & executive report'], ['notifikasi', 'Notifikasi email atasan'], ['rekap-pengemudi', 'Rekap pengemudi'], ['rekap-cabang', 'Rekap cabang'], ['rekap-pemimpin', 'Rekap pemimpin'], ['laporan', 'Laporan bilingual (ID/EN)']]],
  ['BULANAN (DATA GA)', [['ga', 'Laporan bulanan & concern'], ['ga-log', 'Log tindak lanjut']]],
  ['INPUT DATA', [['koreksi', 'Faktor koreksi'], ['faktor', 'Faktor lain (insiden, promosi…)'], ['master', 'Master karyawan & email']]],
  ['ANALISA', [['banding', 'Perbandingan periode / bulan'], ['keadilan', 'Uji keadilan travel load'], ['tren', 'Analisa tren'], ['pengemudi', 'Detail pengemudi']]],
  ['REFERENSI', [['parameter', 'Parameter & standar'], ['matriks', 'Matriks tindak lanjut'], ['periode', 'Periode 2-mingguan']]],
  ['PENGATURAN', [['ai', 'AI (Claude)'], ['instal', 'Instal aplikasi'], ['data', 'Data & cadangan']]]
];
function renderSide() {
  const cur = (location.hash.replace(/^#\/?/, '').split('?')[0]) || 'dashboard';
  const nNotif = S.res ? S.res.rows.filter(r => r.PID === S.res.lastPid && r.NOTIF > 0 && !r.KELUAR).length : 0;
  const pend = (S.inputs.corrections || []).filter(c => c.status === 'Diajukan').length;
  let h = `<div class="brand"><b>VEHICLE TELEMATIC</b><span>Score &amp; Report</span><small>QSHE Department · PT Hexindo Adiperkasa Tbk</small></div>`;
  for (const [g, items] of NAV) {
    h += `<div class="grp">${g}</div>`;
    for (const [id, lbl] of items) {
      const badge = id === 'notifikasi' && nNotif ? `<span class="badge">${nNotif}</span>` : id === 'koreksi' && pend ? `<span class="badge">${pend}</span>` : '';
      h += `<a class="nav ${cur === id ? 'on' : ''}" href="#/${id}"><span class="dot"></span>${esc(lbl)}${badge}</a>`;
    }
  }
  if (S.res) {
    h += `<div class="meta"><div class="k">Periode ditampilkan</div><div class="v">${esc(perLabel(S.ui.pid))}</div><div class="s">${esc(S.ui.cab)}</div>
      <div class="k">Periode laporan terakhir</div><div class="v">${esc(perLabel(S.res.lastPid))}</div>
      <div class="k">Data tersimpan</div><div class="s">${f0(S.trips.length)} trip Driving · ${S.res.periods.length} periode</div>
      <div class="s" style="margin-top:8px">Data diolah di perangkat ini; tidak dikirim ke server.</div></div>`;
  } else h += `<div class="meta"><div class="s">Belum ada data. Mulai dari menu <b style="color:#fff">Unggah data</b>.</div></div>`;
  $('#side').innerHTML = h;
}
async function route() {
  const [id, qs] = (location.hash.replace(/^#\/?/, '') || 'dashboard').split('?');
  const pg = PAGES[id] || PAGES.dashboard;
  if (S.dirty) await recompute();
  $('#side').classList.remove('open');
  renderSide();
  Object.keys(S.charts).forEach(k => { S.charts[k].destroy(); delete S.charts[k]; });
  const v = $('#view'); v.innerHTML = '';
  const params = new URLSearchParams(qs || '');
  try { await pg.render(v, params); } catch (e) { console.error(e); v.innerHTML = `<div class="pad"><div class="empty"><b>Halaman gagal ditampilkan</b>${esc(e.message)}</div></div>`; }
  v.focus({ preventScroll: true }); window.scrollTo(0, 0);
  document.title = (pg.title ? pg.title + ' · ' : '') + 'Vehicle Telematic Score and Report';
}
function head(title, sub) { return `<div class="pagehead"><h2>${esc(title)}</h2>${sub ? `<p>${sub}</p>` : ''}</div>`; }
function needData(v) {
  if (S.res) return false;
  v.innerHTML = head('Belum ada data', '') + `<div class="pad"><div class="empty"><b>Unggah data mentah Cartrack dulu</b>Menu <a href="#/unggah">Unggah data</a> → pilih file "Data Score Behavior" (.xlsx/.csv) dari Tim VT. Semua diolah di perangkat ini.</div></div>`;
  return true;
}
function periodOptions(sel, withLast) {
  const ps = S.res.periods.slice().reverse();
  return (withLast ? `<option value="last">Periode terakhir</option>` : '') + ps.map(p => `<option value="${p.pid}" ${p.pid === sel ? 'selected' : ''}>${esc(p.lbl)}${p.days < 7 ? ' (data ' + p.days + ' hari)' : ''}</option>`).join('');
}
function cabList() { return [...new Set(S.res.rows.map(r => r.cab))].filter(Boolean).sort(); }

/* ------------------------------------------------------------ DASHBOARD */
PAGES.dashboard = {
  title: 'Dashboard', render(v) {
    const p0 = S.res ? per(S.res.lastPid) : null;
    const first = S.res ? S.res.periods[0] : null;
    const fyTxt = p0 ? 'FY' + p0.fy : '';
    const ml = d => { const t = VT.dayToDate(d); return VT.MON[t.getUTCMonth()]; };
    const foot = S.res ? `Data Vehicle Telematics ${ml(first.start)} ${VT.dayToDate(first.start).getUTCFullYear()}–${ml(p0.end)} ${VT.dayToDate(p0.end).getUTCFullYear()} · skoring berimbang travel load` : 'Unggah data mentah Cartrack untuk mulai';
    const nightTxt = String(P().NSTART).padStart(2, '0') + ':00 – ' + String((P().NEND + 23) % 24).padStart(2, '0') + ':59';
    let h = `<section class="hero" aria-label="Vehicle Telematic Score and Report"><div class="in">
      <span class="pill">QSHE DEPARTMENT · PT HEXINDO ADIPERKASA TBK${fyTxt ? ' · ' + fyTxt : ''}</span>
      <h1>Vehicle Telematic<br>Score and Report<i>.</i></h1>
      <p class="tag">Adil untuk travel load tinggi. Tegas untuk pelanggaran berulang.</p>
      <div class="chips"><span>Travel load 5 band</span><span>Berulang 2 minggu</span><span>Insiden &amp; pemimpin</span><span>Notifikasi atasan</span></div>
      <div class="foot">${esc(foot)}</div></div>
      <div class="callouts" aria-hidden="true">
        <div class="co" style="left:48.4%;top:10%"><b>TRAVEL LOAD</b>dibanding band km sejenis</div>
        <div class="co" style="left:69.7%;top:7%"><b>NIGHT DRIVING</b>mulai ${esc(nightTxt)}</div>
        <div class="co" style="right:1.2%;top:21%"><b>OVERSPEED</b>toleransi ≤${P().OSTOL} count/trip</div>
        <div class="co" style="right:1.2%;top:82%"><b>HARSH BRAKING</b>laju per 1.000 km</div></div></section>`;
    if (!S.res) { v.innerHTML = h + `<div class="pad"><div class="empty"><b>Belum ada data</b>Buka <a href="#/unggah">Unggah data</a> lalu pilih file mentah Cartrack ("Data Score Behavior"). Aplikasi menghitung skor, notifikasi atasan, dan analisa secara otomatis.</div></div>`; return; }
    const cabs = cabList();
    h += `<div class="filters noprint"><label for="fPer">PERIODE</label><select id="fPer">${periodOptions(S.ui.pid === S.res.lastPid ? null : S.ui.pid, true)}</select>
      <label for="fCab">CABANG</label><select id="fCab"><option>Semua cabang</option>${cabs.map(c => `<option ${c === S.ui.cab ? 'selected' : ''}>${esc(c)}</option>`).join('')}</select>
      <span class="hint" id="fHint"></span><span class="sp" style="flex:1"></span><button class="btn pri" id="dlAll" type="button" title="Satu file Excel: dashboard di halaman depan + semua menu">⬇ Unduh semua (Excel)</button></div><div class="pad" id="dash"></div>
      <footer class="f">Vehicle Telematic Score and Report · QSHE Department PT Hexindo Adiperkasa Tbk · model v2.0 (sama dengan workbook Excel) · ilustrasi kendaraan generik · aplikasi v${VT.VERSION}</footer>`;
    v.innerHTML = h;
    $('#fPer').value = S.ui.pid === S.res.lastPid ? 'last' : String(S.ui.pid);
    $('#fPer').onchange = e => { S.ui.pid = e.target.value === 'last' ? S.res.lastPid : +e.target.value; renderSide(); drawDash(); };
    $('#fCab').onchange = e => { S.ui.cab = e.target.value; renderSide(); drawDash(); };
    $('#dlAll').onclick = () => exportAll(S.ui.pid);
    drawDash();
  }
};
function drawDash() {
  try { drawDash0(); } catch (e) { console.error(e); toast('Sebagian dashboard gagal ditampilkan: ' + e.message, 8000); }
}
function drawDash0() {
  const pid = S.ui.pid, cab = S.ui.cab, p = per(pid), Pp = P();
  const k = VT.kpis(S.res, pid, cab);
  $('#fHint').textContent = '▾ pilih di kotak oranye · ' + f0(k.n) + ' pengemudi aktif';
  const delta = k.finPrev != null && k.fin != null ? k.fin - k.finPrev : null;
  const nightTxt = String(Pp.NSTART).padStart(2, '0') + ':00–' + String((Pp.NEND + 23) % 24).padStart(2, '0') + ':59';
  const d = $('#dash');
  d.innerHTML = `<div class="h-sum">RINGKASAN · ${esc(p.lbl.toUpperCase())} (${esc(perRange(pid))}) · ${cab === 'Semua cabang' ? 'NASIONAL' : esc(cab.toUpperCase())}</div>
  <div class="kpis">
    <div class="kpi acc"><span class="l">SKOR RATA-RATA (METODE BARU)</span><span class="v">${f1(k.fin)}</span><span class="s">metode lama ${f1(k.old)}${delta != null ? ' · ' + (delta >= 0 ? '▲ ' : '▼ ') + f1(Math.abs(delta)) + ' vs periode lalu' : ''}</span></div>
    <div class="kpi"><span class="l">PENGEMUDI AKTIF</span><span class="v">${f0(k.n)}</span><span class="s">${f0(k.km)} km · rata-rata ${f0(k.kmPer)} km/orang</span></div>
    <div class="kpi"><span class="l">GRADE A–B</span><span class="v">${pc(k.ab)}</span><span class="s">${k.de} orang grade D/E perlu pembinaan</span></div>
    <a class="kpi" href="#/notifikasi?pid=${pid}"><span class="l">NOTIFIKASI KE ATASAN</span><span class="v">${f0(k.notif)}</span><span class="s">${k.esk} eskalasi · buka menu Notifikasi email</span></a>
    <div class="kpi"><span class="l">PELANGGARAN BERULANG (2 MINGGU)</span><span class="v">${f0(k.rep)}</span><span class="s">${k.n ? fx(k.rep / k.n * 100, 0) : 0}% dari pengemudi aktif</span></div>
    <div class="kpi"><span class="l">OVERSPEED STATUS MERAH</span><span class="v">${f0(k.merah)}</span><span class="s">KUNING ${k.kuning} · count dihitung ${f0(k.osv)}</span></div>
    <div class="kpi"><span class="l">NIGHT DRIVING (JAM)</span><span class="v">${f0(k.nh)}</span><span class="s">${f1(k.nhPct)}% jam mengemudi · mulai ${nightTxt}</span></div>
    <div class="kpi"><span class="l">HARSH BRAKING / 1.000 KM</span><span class="v">${f2(k.brk)}</span><span class="s">periode lalu ${k.brkPrev == null ? '–' : f2(k.brkPrev)}</span></div>
  </div>
  <h3 class="sec">EXECUTIVE REPORT</h3><div id="execBox"></div>
  <div class="grid2" style="margin-top:14px">
    <div class="panel"><h4>Tren skor rata-rata per periode</h4><div class="sub">Garis oranye = metode baru · garis abu putus-putus = metode lama (11 periode terakhir)</div><div class="chartbox"><canvas id="cTren" role="img" aria-label="Tren skor rata-rata"></canvas></div></div>
    <div class="panel"><h4>Sebaran grade · periode terpilih</h4><div class="sub">Jumlah pengemudi per grade (A ≥${Pp.GRADE[0]} · B ≥${Pp.GRADE[1]} · C ≥${Pp.GRADE[2]} · D ≥${Pp.GRADE[3]} · E &lt;${Pp.GRADE[3]})</div><div class="chartbox"><canvas id="cGrade" role="img" aria-label="Sebaran grade"></canvas></div></div>
    <div class="panel"><h4>Laju pelanggaran per 1.000 km</h4><div class="sub">Overspeed dihitung (setelah toleransi), harsh braking, harsh acceleration</div><div class="chartbox"><canvas id="cLaju" role="img" aria-label="Laju pelanggaran"></canvas></div></div>
    <div class="panel"><h4>Keadilan travel load · skor per band</h4><div class="sub">Metode lama turun makin jauh jalan; metode baru rata antar band</div><div class="chartbox"><canvas id="cBand" role="img" aria-label="Skor per band"></canvas></div></div>
  </div>
  <h3 class="sec">10 PENGEMUDI PRIORITAS COACHING · skor terendah di periode &amp; cabang terpilih</h3><div id="top10"></div>
  <div class="grid2" style="margin-top:6px"><div><h3 class="sec">CABANG TERBAIK · periode terpilih</h3><div id="best"></div></div><div><h3 class="sec">CABANG PERLU PERHATIAN · periode terpilih</h3><div id="worst"></div></div></div>
  <h3 class="sec">STATUS DATA &amp; TINDAK LANJUT</h3><ul class="status" id="stat"></ul>`;
  // executive report
  renderExec($('#execBox'), pid, cab);
  // charts
  const tr = VT.trend(S.res, S.res.lastPid, cab, 11);
  chart('cTren', { type: 'line', data: { labels: tr.map(x => x.lbl), datasets: [
    { label: 'Skor metode lama', data: tr.map(x => x.old), borderColor: '#9e9e9e', borderDash: [6, 4], borderWidth: 1.75, pointRadius: 2, tension: 0 },
    { label: 'Skor baru', data: tr.map(x => x.fin), borderColor: '#f7941d', backgroundColor: '#f7941d', borderWidth: 3, pointRadius: 3, tension: 0 }] },
    options: { scales: { y: Object.assign({ min: 60, max: 100 }, gridY), x: { grid: { display: false } } }, plugins: { legend: { position: 'bottom' } } } });
  chart('cGrade', { type: 'bar', data: { labels: ['A', 'B', 'C', 'D', 'E'], datasets: [{ label: 'Pengemudi', data: k.grades, backgroundColor: ['#111111', '#6b6b6b', '#f9c38a', '#f7941d', '#c2410c'], borderRadius: 3, maxBarThickness: 64 }] },
    options: { plugins: { legend: { display: false } }, scales: { y: Object.assign({ beginAtZero: true, ticks: { precision: 0 } }, gridY), x: { grid: { display: false } } } } });
  chart('cLaju', { type: 'line', data: { labels: tr.map(x => x.lbl), datasets: [
    { label: 'Overspeed dihitung /1.000 km', data: tr.map(x => x.osv), borderColor: '#f7941d', borderWidth: 3, pointRadius: 2, tension: 0 },
    { label: 'Harsh braking /1.000 km', data: tr.map(x => x.brk), borderColor: '#111111', borderWidth: 2.25, pointRadius: 2, tension: 0 },
    { label: 'Harsh accel /1.000 km', data: tr.map(x => x.acc), borderColor: '#9e9e9e', borderWidth: 2, pointRadius: 2, tension: 0 }] },
    options: { scales: { y: Object.assign({ beginAtZero: true }, gridY), x: { grid: { display: false } } }, plugins: { legend: { position: 'bottom' } } } });
  chart('cBand', { type: 'bar', data: { labels: k.bands.map(b => b.band), datasets: [
    { label: 'Skor metode lama', data: k.bands.map(b => b.old || null), backgroundColor: '#bdbdbd', borderRadius: 3, maxBarThickness: 40 },
    { label: 'Skor baru', data: k.bands.map(b => b.fin || null), backgroundColor: '#f7941d', borderRadius: 3, maxBarThickness: 40 }] },
    options: { scales: { y: Object.assign({ min: 60, max: 100 }, gridY), x: { grid: { display: false } } }, plugins: { legend: { position: 'bottom' } } } });
  // top 10
  const s = VT.sel(S.res, pid, cab).slice().sort((a, b) => a.FIN - b.FIN || (a.drv < b.drv ? -1 : 1)).slice(0, 10);
  table($('#top10'), [
    { k: 'i', l: '#', f: (r, i) => i + 1, nosort: true, cls: 'c' }, { k: 'drv', l: 'Nama pengemudi', f: r => `<b>${esc(r.drv)}</b>` }, { k: 'cab', l: 'Cabang' }, { k: 'band', l: 'Band', cls: 'c' },
    { k: 'km', l: 'KM', f: r => f0(r.km), cls: 'r' }, { k: 'FIN', l: 'Skor', f: r => `<span class="bar"><i style="width:${Math.max(0, r.FIN)}%"></i><span><b>${f1(r.FIN)}</b></span></span>` },
    { k: 'GR', l: 'Grade', f: r => gradeChip(r.GR), cls: 'c' }, { k: 'OSS', l: 'Status overspeed', f: r => ossChip(r.OSS) },
    { k: 'NREP', l: 'Berulang', f: r => (r.NREP > 0 ? r.NREP + ' komponen' : '–') }, { k: 'WK', l: 'Fokus / tindak lanjut', f: r => (r.NOTIF > 0 ? '<b>NOTIF ATASAN</b> · ' : '') + esc(r.WK) }
  ], s, { search: false, free: true, onRow: r => { location.hash = '#/pengemudi?d=' + encodeURIComponent(r.drv) + '&pid=' + r.PID; } });
  // cabang
  const br = VT.branches(S.res, pid, Pp).filter(b => b.n >= Pp.MIN_BRANCH);
  const bt = (arr) => [{ k: 'i', l: '#', f: (r, i) => i + 1, nosort: true, cls: 'c' }, { k: 'cab', l: 'Cabang', f: r => `<b>${esc(r.cab)}</b>` }, { k: 'n', l: 'Pengemudi', cls: 'c' }, { k: 'fin', l: 'Skor', f: r => f1(r.fin), cls: 'r' }, { k: 'de', l: '% D/E', f: r => pc(r.de), cls: 'r' }];
  table($('#best'), bt(), br.slice().sort((a, b) => b.fin - a.fin).slice(0, 5), { search: false, free: true });
  table($('#worst'), bt(), br.slice().sort((a, b) => a.fin - b.fin).slice(0, 5), { search: false, free: true });
  // status
  const nt = VT.buildNotifications(S.res, S.res.lastPid, Pp, S.res.master, S.res.tripsByKey);
  const ms = S.inputs.master || [];
  const ok = ms.filter(m => /@/.test(m.atasanEmail || '')).length;
  const corr = S.inputs.corrections || [];
  const F = S.inputs.factors;
  const leaders = VT.recapLeaders(S.res, S.res.lastPid, Pp).length;
  $('#stat').innerHTML = [
    [`Notifikasi periode laporan terakhir: ${nt.filter(x => x.kirim === 'YA').length} siap kirim · ${nt.filter(x => x.kirim !== 'YA').length} tertahan karena email atasan belum ada`, 'notifikasi'],
    [`Master karyawan: ${ok} lengkap · ${ms.length - ok} belum ada email atasan langsung · ${new Set(S.res.rows.filter(r => r.PID === S.res.lastPid).map(r => r.drv)).size - S.res.rows.filter(r => r.PID === S.res.lastPid && S.res.master.has(r.drv)).length} pengemudi aktif belum ada di master`, 'master'],
    [`Faktor koreksi: ${corr.filter(c => c.status === 'Disetujui').length} disetujui · ${corr.filter(c => c.status === 'Diajukan').length} menunggu persetujuan QSHE`, 'koreksi'],
    [`Faktor lain: ${(F.insiden || []).length} insiden · ${(F.promosi || []).length} promosi · ${(F.mutasi || []).length} mutasi · ${(F.keluar || []).length} keluar · ${(F.pelatihan || []).length} pelatihan`, 'faktor'],
    [`Pemimpin terdeteksi (FR/CSS/Unit/Section Head): ${leaders} orang`, 'rekap-pemimpin'],
    [`Kualitas data: ${f0(S.res.dq.generic)} trip ID generik (DRIVER…/MASTER…, ${S.res.dq.genericIds} ID) · ${f0(S.res.dq.v120)} trip kecepatan rata-rata >120 km/jam`, 'data'],
    [`Laporan bulanan GA: ${S.ga.months.length ? S.ga.months.map(m => VT.monthLabel(m)).join(', ') : 'belum diunggah'}`, 'ga']
  ].map(([t, l]) => `<li><span>${esc(t)}</span><a href="#/${l}">Buka ›</a></li>`).join('');
}
function renderExec(box, pid, cab) {
  const lang = S.settings.lang || 'id';
  const ex = VTR.exec(S.res, pid, cab, lang, S);
  const key = 'exec|' + pid + '|' + cab + '|' + lang;
  const ai = (S.settings.aiCache || {})[key];
  box.innerHTML = `<div class="exec"><div class="top"><h4>${lang === 'en' ? 'Executive report' : 'Executive report'} · ${esc(perLabel(pid))}${cab !== 'Semua cabang' ? ' · ' + esc(cab) : ''}</h4>
    <div class="tools noprint" style="margin:0"><button class="btn sm ${lang === 'id' ? 'dark' : ''}" data-l="id">ID</button><button class="btn sm ${lang === 'en' ? 'dark' : ''}" data-l="en">EN</button>
    <button class="btn sm" data-copy>Salin</button><button class="btn sm pri" data-ai>Rapikan dengan AI</button>${ai ? '<button class="btn sm" data-orig>Versi asli</button>' : ''}</div></div>
    <div id="execBody">${ai && !box._orig ? `<div class="aitxt">${esc(ai)}</div><p class="note">Dirapikan dengan AI (${esc(S.settings.ai.provider)}). Angka tetap dari perhitungan aplikasi.</p>` : VTR.execHtml(ex)}</div></div>`;
  $$('[data-l]', box).forEach(b => b.onclick = () => { S.settings.lang = b.dataset.l; saveSettings(); renderExec(box, pid, cab); });
  $('[data-copy]', box).onclick = () => copyText(ai && !box._orig ? ai : VTR.execText(ex));
  const o = $('[data-orig]', box); if (o) o.onclick = () => { box._orig = !box._orig; renderExec(box, pid, cab); };
  $('[data-ai]', box).onclick = async () => {
    const txt = await AI.tidy(VTR.execText(ex), lang === 'en' ? 'exec-en' : 'exec-id');
    if (txt) { S.settings.aiCache = S.settings.aiCache || {}; S.settings.aiCache[key] = txt; saveSettings(); box._orig = false; renderExec(box, pid, cab); }
  };
}

/* ------------------------------------------------------------ boot */
async function boot() {
  chartsInit();
  $('#menuBtn').onclick = () => $('#side').classList.toggle('open');
  window.addEventListener('hashchange', route);
  const upd = () => { $('#offline').hidden = navigator.onLine; }; window.addEventListener('online', upd); window.addEventListener('offline', upd); upd();
  busy('Memuat data tersimpan…');
  try { await loadAll(); compute(); }
  catch (e) { console.error(e); toast('Gagal memuat data lokal: ' + e.message, 8000); }
  finally { busy(false); }
  if ('serviceWorker' in navigator && location.protocol !== 'file:') {
    navigator.serviceWorker.register('sw.js').then(reg => {
      reg.addEventListener('updatefound', () => { const w = reg.installing; if (w) w.addEventListener('statechange', () => { if (w.state === 'installed' && navigator.serviceWorker.controller) toast('Versi baru aplikasi siap. Muat ulang halaman untuk memakainya.', 8000); }); });
    }).catch(() => {});
  }
  route();
}
document.addEventListener('DOMContentLoaded', boot);
