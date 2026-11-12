/* =====================================================================
   Service worker — aplikasi tetap jalan tanpa internet.
   SETIAP KALI ADA PERUBAHAN FILE: naikkan VERSI (v1 -> v2 -> ...)
   supaya perangkat pengguna mengambil versi terbaru.
   ===================================================================== */
const VERSI = 'v6';
const CACHE = 'vt-score-' + VERSI;
const FILES = [
  './', './index.html', './manifest.webmanifest', './icon-192.png', './icon-512.png',
  './css/style.css', './assets/hero.jpg',
  './js/xlsx-stream.js', './js/engine.js', './js/db.js', './js/worker.js', './js/report.js', './js/app.js', './js/ai.js', './js/pages1.js', './js/pages2.js', './js/export-all.js',
  './lib/chart.umd.js', './lib/fflate.min.js', './lib/exceljs.min.js', './lib/docx.iife.js',
  './fonts/poppins-Bold.woff2', './fonts/poppins-Medium.woff2', './fonts/poppins-Regular.woff2',
  './fonts/barlow-latin-400-normal.woff2', './fonts/barlow-latin-600-normal.woff2', './fonts/barlow-latin-700-normal.woff2'
];
self.addEventListener('install', e => { e.waitUntil(caches.open(CACHE).then(c => c.addAll(FILES)).then(() => self.skipWaiting())); });
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k.startsWith('vt-score-') && k !== CACHE).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});
// Jaringan dulu (selalu versi terbaru saat online), cadangan dari cache saat offline.
self.addEventListener('fetch', e => {
  if (e.request.method !== 'GET' || new URL(e.request.url).origin !== location.origin) return;
  e.respondWith((async () => {
    const c = await caches.open(CACHE);
    try {
      const ctl = new AbortController(); const t = setTimeout(() => ctl.abort(), 6000);
      const r = await fetch(e.request, { signal: ctl.signal, cache: 'no-cache' }); clearTimeout(t);
      if (r.ok) c.put(e.request, r.clone());
      return r;
    } catch (err) {
      return (await c.match(e.request, { ignoreSearch: true })) || (await c.match('./index.html')) || Response.error();
    }
  })());
});
