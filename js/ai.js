/* =====================================================================
   ai.js — opsi "Rapikan dengan AI" (opsional). Default: Claude.
   Mode "web": prompt disalin lalu Claude.ai dibuka — pengguna login
   dengan akunnya sendiri, tempel, lalu tempel kembali hasilnya.
   Mode "api": memakai API key Anthropic milik pengguna (disimpan hanya
   di perangkat ini) untuk merapikan langsung di aplikasi.
   ===================================================================== */
'use strict';
const AI = (function () {
  const PROVIDERS = {
    claude: { name: 'Claude (Anthropic)', web: 'https://claude.ai/new', login: 'https://claude.ai/login' },
    chatgpt: { name: 'ChatGPT', web: 'https://chatgpt.com/', login: 'https://chatgpt.com/auth/login' },
    gemini: { name: 'Gemini', web: 'https://gemini.google.com/app', login: 'https://gemini.google.com/app' }
  };
  const PURPOSE = {
    'exec-id': 'Rapikan executive report berikut menjadi bahasa Indonesia korporat yang ringkas, strategis, dan mudah dipahami direksi. Pertahankan SEMUA angka, nama cabang, dan struktur (judul, kondisi, risiko strategis, keputusan & aksi, target). Jangan menambah fakta atau angka baru. Keluarkan teks polos tanpa markdown.',
    'exec-en': 'Polish the following executive report into concise, strategic corporate English suitable for the board. Keep ALL numbers, branch names and the structure (title, situation, strategic risks, decisions & actions, targets). Do not add new facts or numbers. Output plain text without markdown.',
    'email': 'Rapikan email notifikasi keselamatan berkendara berikut agar tetap sopan, komunikatif, dan jelas untuk atasan di lapangan. Pertahankan semua angka, tanggal, jam, plat nomor, dan daftar trip apa adanya. Jangan menambah fakta. Keluarkan teks email polos tanpa markdown.',
    'report': 'Rapikan narasi laporan berikut (bahasa Indonesia lalu English) menjadi gaya laporan profesional. Pertahankan semua angka dan istilah. Jangan menambah fakta. Teks polos.'
  };
  const cfg = () => S.settings.ai;
  /** Samarkan nama pengemudi sebelum dikirim ke AI */
  function anonymize(text) {
    if (!cfg().anonymize || !S.res) return { text, map: null };
    const names = [...new Set(S.res.rows.map(r => r.drv))].sort((a, b) => b.length - a.length);
    const map = {}; let i = 0; let out = text;
    for (const n of names) {
      if (n.length < 4 || out.indexOf(n) < 0) continue;
      const tag = '[PENGEMUDI-' + (++i) + ']'; map[tag] = n; out = out.split(n).join(tag);
    }
    (S.inputs.master || []).forEach(m => { [m.nama, m.atasan, m.atasan2].forEach(x => { if (x && x.length > 3 && out.indexOf(x) >= 0) { const tag = '[ORANG-' + (++i) + ']'; map[tag] = x; out = out.split(x).join(tag); } }); });
    return { text: out, map };
  }
  const restore = (t, map) => { if (!map) return t; for (const k in map) t = t.split(k).join(map[k]); return t; };

  async function callClaude(prompt) {
    const c = cfg();
    const r = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: { 'content-type': 'application/json', 'x-api-key': c.apiKey, 'anthropic-version': '2023-06-01', 'anthropic-dangerous-direct-browser-access': 'true' },
      body: JSON.stringify({ model: c.model || 'claude-sonnet-5-5', max_tokens: 4000, messages: [{ role: 'user', content: prompt }] })
    });
    const j = await r.json().catch(() => ({}));
    if (!r.ok) throw new Error((j.error && j.error.message) || ('HTTP ' + r.status));
    return (j.content || []).filter(x => x.type === 'text').map(x => x.text).join('\n').trim();
  }

  /** Kembalikan teks hasil (atau null bila dibatalkan) */
  async function tidy(text, purpose) {
    const c = cfg(); const pv = PROVIDERS[c.provider] || PROVIDERS.claude;
    const an = anonymize(text);
    const prompt = PURPOSE[purpose] + '\n\n---\n' + an.text;
    if (c.provider === 'claude' && c.mode === 'api' && c.apiKey) {
      busy('Merapikan dengan Claude…');
      try { const out = await callClaude(prompt); toast('Selesai dirapikan oleh Claude.'); return restore(out, an.map); }
      catch (e) { toast('Claude API gagal: ' + e.message + '. Coba mode "Buka di Claude.ai".', 8000); return null; }
      finally { busy(false); }
    }
    // mode web: salin prompt, buka situs AI, minta tempel hasil
    return new Promise(res => {
      const md = modal('Rapikan dengan ' + pv.name, `
        <ol class="steps" style="margin-bottom:12px">
          <li>Klik <b>Salin prompt &amp; buka ${esc(pv.name)}</b>. Login dengan akun Anda bila diminta.</li>
          <li>Tempel (Ctrl+V) di kotak chat ${esc(pv.name)} lalu kirim.</li>
          <li>Salin jawabannya, tempel di kotak di bawah, lalu klik <b>Pakai hasil</b>.</li></ol>
        ${an.map ? '<p class="note">Nama orang sudah disamarkan (mis. [PENGEMUDI-1]) dan otomatis dikembalikan saat hasil dipakai.</p>' : '<p class="note">Penyamaran nama nonaktif (menu AI). Data nama ikut terkirim ke layanan AI.</p>'}
        <div class="tools"><button class="btn pri" data-go>Salin prompt &amp; buka ${esc(pv.name)}</button><button class="btn" data-cp>Salin prompt saja</button></div>
        <div class="fld"><label for="aiOut">Hasil dari ${esc(pv.name)}</label><textarea id="aiOut" rows="12" placeholder="Tempel jawaban AI di sini…"></textarea></div>
        <div class="tools"><span class="sp"></span><button class="btn" data-x>Batal</button><button class="btn dark" data-ok>Pakai hasil</button></div>`, { wide: true, onClose: () => res(null) });
      $('[data-go]', md.body).onclick = async () => { await copyText(prompt); window.open(pv.web, '_blank', 'noopener'); };
      $('[data-cp]', md.body).onclick = () => copyText(prompt);
      $('[data-x]', md.body).onclick = () => { md.el.remove(); res(null); };
      $('[data-ok]', md.body).onclick = () => { const t = $('#aiOut', md.body).value.trim(); if (!t) { toast('Tempel hasil dari AI dulu.'); return; } md.el.remove(); res(restore(t, an.map)); };
    });
  }
  async function test() {
    if (cfg().mode !== 'api' || !cfg().apiKey) { window.open((PROVIDERS[cfg().provider] || PROVIDERS.claude).login, '_blank', 'noopener'); return 'Halaman login dibuka di tab baru.'; }
    const out = await callClaude('Jawab satu kata: OK');
    return 'Terhubung ke Claude API. Jawaban: ' + out.slice(0, 40);
  }
  return { tidy, test, PROVIDERS };
})();
