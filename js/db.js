/* db.js — penyimpanan lokal (IndexedDB). Data tidak pernah dikirim ke server. */
(function (root) {
  'use strict';
  const NAME = 'vt-score', VER = 1;
  let dbp = null;
  function open() {
    if (dbp) return dbp;
    dbp = new Promise((res, rej) => {
      const r = indexedDB.open(NAME, VER);
      r.onupgradeneeded = () => {
        const d = r.result;
        if (!d.objectStoreNames.contains('periods')) d.createObjectStore('periods', { keyPath: 'lbl' });
        if (!d.objectStoreNames.contains('kv')) d.createObjectStore('kv');
      };
      r.onsuccess = () => res(r.result);
      r.onerror = () => rej(r.error || new Error('IndexedDB tidak tersedia (mode privat?)'));
      r.onblocked = () => rej(new Error('Tutup tab lain aplikasi ini lalu muat ulang.'));
    });
    return dbp;
  }
  const tx = async (store, mode, fn) => {
    const d = await open();
    return new Promise((res, rej) => {
      const t = d.transaction(store, mode); const s = t.objectStore(store); let out;
      Promise.resolve(fn(s)).then(v => { out = v; });
      t.oncomplete = () => res(out); t.onerror = () => rej(t.error); t.onabort = () => rej(t.error || new Error('Transaksi dibatalkan (ruang penyimpanan penuh?)'));
    });
  };
  const req = r => new Promise((res, rej) => { r.onsuccess = () => res(r.result); r.onerror = () => rej(r.error); });
  const DB = {
    open,
    get: (k) => tx('kv', 'readonly', s => req(s.get(k))),
    set: (k, v) => tx('kv', 'readwrite', s => { s.put(v, k); }),
    del: (k) => tx('kv', 'readwrite', s => { s.delete(k); }),
    keys: () => tx('kv', 'readonly', s => req(s.getAllKeys())),
    periods: () => tx('periods', 'readonly', s => req(s.getAll())),
    periodMeta: async () => { const all = await DB.periods(); return all.map(p => ({ lbl: p.lbl, stats: p.stats, n: p.trips.length })); },
    putPeriods: (arr) => tx('periods', 'readwrite', s => { arr.forEach(p => s.put(p)); }),
    delPeriod: (lbl) => tx('periods', 'readwrite', s => { s.delete(lbl); }),
    getPeriod: (lbl) => tx('periods', 'readonly', s => req(s.get(lbl))),
    clearAll: async () => { await tx('periods', 'readwrite', s => { s.clear(); }); await tx('kv', 'readwrite', s => { s.clear(); }); }
  };
  root.DB = DB;
})(typeof self !== 'undefined' ? self : globalThis);
