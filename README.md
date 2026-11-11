# Vehicle Telematic Score and Report

**QSHE Department – PT Hexindo Adiperkasa Tbk**

Aplikasi web (bisa di-install seperti aplikasi) yang mengolah data mentah Vehicle Telematics **Cartrack** ("Data Score Behavior") menjadi skor berkendara yang adil terhadap travel load, notifikasi email ke atasan, laporan bulanan GA, dan laporan profesional bilingual. Rumusnya sama 1:1 dengan workbook Excel *Vehicle_Telematic_Score_and_Report_Hexindo.xlsx*.

*A web app (installable like a native app) that turns raw Cartrack Vehicle Telematics data ("Data Score Behavior") into travel-load-fair driving scores, supervisor email notifications, GA monthly reports, and a professional bilingual report. The formulas match the Excel workbook* Vehicle_Telematic_Score_and_Report_Hexindo.xlsx *one-to-one.*

---

## Bahasa Indonesia

### Fitur utama
- **Unggah data mentah**: file Cartrack .xlsx (bisa >100 MB, dibaca streaming) atau .csv; lampiran GA "Data Case" bulanan; master karyawan (sheet 10 workbook Excel atau template CSV); file cadangan .json. Trip ganda otomatis dilewati.
- **Dashboard** dengan format sama seperti sheet 00_DASHBOARD: hero, filter periode & cabang, 8 kartu KPI, 4 grafik, 10 prioritas coaching, cabang terbaik/perlu perhatian, status data.
- **Executive report strategis** (ID/EN) langsung di dashboard: kondisi, risiko strategis, keputusan & aksi (PIC, tenggat), target periode berikut.
- **Notifikasi email atasan**: isi email sama dengan workbook (kapan, di mana, apa pelanggarannya), eskalasi ke atasan tidak langsung, salin / buka di aplikasi email / ekspor CSV & Excel untuk Power Automate.
- **Faktor koreksi**: DID dipinjam (pindahkan ke pemakai sebenarnya), DID tertukar, error GPS/sensor, lonjakan GPS >120 km/jam, darurat medis, pengawalan resmi, perjalanan malam ber-JRM, test drive mekanik, unit pihak ketiga, data ganda, alasan lain. Wajib bukti; hanya status *Disetujui* QSHE yang mengubah skor. Ada saran koreksi otomatis dan jejak audit.
- **Faktor lain**: insiden/kecelakaan (pengurang pengemudi & pemimpin), promosi/perubahan jabatan, mutasi cabang, cuti, resign, pelatihan (dengan analisa efektivitas), ID bersama.
- **Analisa**: perbandingan dua periode 2-mingguan atau dua bulan, uji keadilan travel load, tren bulanan, detail perhitungan per pengemudi, rekap pengemudi/cabang/pemimpin.
- **Laporan bulanan GA**: sebelum vs sesudah, prioritas P1/P2, per unit, concern & kesepakatan, log tindak lanjut (close-loop), prompt flyer.
- **Laporan bilingual** Indonesia + English: unduh Word (.docx) atau cetak/PDF.
- **AI opsional (Claude)**: rapikan bahasa executive report, email, dan laporan. Mode login ke Claude.ai dengan akun sendiri, atau API key Anthropic. Nama orang disamarkan sebelum dikirim.
- **Install & offline**: PWA (Android, iPhone/iPad, Windows/Mac), tetap jalan tanpa internet.

### Privasi
Semua data diolah dan disimpan **di browser perangkat pengguna** (IndexedDB). Tidak ada server, tidak ada data yang diunggah ke GitHub. Repo ini hanya berisi kode. Cadangkan data lewat menu *Data & cadangan*.

### Cara memakai
1. Buka aplikasi (GitHub Pages), menu **Unggah data**, pilih file mentah Cartrack.
2. Lengkapi **Master karyawan** (email atasan) dan catat **Faktor koreksi / Faktor lain** bila ada.
3. Baca **Dashboard & executive report**, kirim **Notifikasi email atasan**, unduh **Laporan bilingual**.
4. Panduan lengkap ada di menu **Panduan** di dalam aplikasi.

### Publikasi (GitHub Pages)
Settings → Pages → *Build and deployment*: Source **Deploy from a branch**, Branch **main** / **(root)** → Save. Alamat: `https://rochiki.github.io/vt-score/`.

### Memperbarui aplikasi
Setiap mengubah file, naikkan `VERSI` di `sw.js` (v1 → v2 …) agar perangkat pengguna mengambil versi baru.

---

## English

### Key features
- **Raw data upload**: Cartrack .xlsx (100 MB+, streamed) or .csv; GA monthly "Data Case" attachment; employee master (sheet 10 of the Excel workbook or CSV template); .json backups. Duplicate trips are skipped automatically.
- **Dashboard** in the same layout as sheet 00_DASHBOARD: hero, period & branch filters, 8 KPI cards, 4 charts, top-10 coaching list, best / attention branches, data status.
- **Strategic executive report** (ID/EN) on the dashboard: situation, strategic risks, decisions & actions (owner, timing), next-period targets.
- **Supervisor email notifications** identical to the workbook (when, where, what), escalation to the indirect supervisor, copy / open in mail app / export CSV & Excel for Power Automate.
- **Corrections**: lent DID tag (reassign to the actual driver), swapped tag, verified GPS/sensor error, GPS spike >120 km/h, medical emergency, official escort, JRM-approved night trip, mechanic road test, third-party operation, duplicates, other. Evidence required; only QSHE-*approved* corrections change scores. Includes automatic suggestions and an audit trail.
- **Other factors**: incidents/accidents (driver & leader deductions), promotions/role changes, branch transfers, leave, resignations, training (with effectiveness analysis), shared IDs.
- **Analysis**: compare two bi-weekly periods or two months, travel-load fairness test, monthly trend, per-driver calculation breakdown, driver/branch/leader recaps.
- **GA monthly report**: before vs after, P1/P2 priorities, per unit, concerns, follow-up log (closed loop), flyer prompt.
- **Bilingual report** Indonesian + English: download Word (.docx) or print/PDF.
- **Optional AI (Claude)** to polish the executive report, emails, and report wording — sign in to Claude.ai with your own account, or use an Anthropic API key. Names are masked before sending.
- **Installable & offline** PWA (Android, iPhone/iPad, Windows/Mac).

### Privacy
All data is processed and stored **in the user's browser** (IndexedDB). There is no server and no data is uploaded to GitHub; this repository contains code only.

### Publishing
Settings → Pages → Source **Deploy from a branch**, Branch **main** / **(root)**.

---

### Struktur / Structure
| Path | Isi / Contents |
|---|---|
| `index.html`, `css/style.css` | Tampilan / UI shell |
| `js/xlsx-stream.js` | Pembaca .xlsx streaming / streaming .xlsx reader |
| `js/engine.js` | Mesin skor (port rumus Excel) / scoring engine |
| `js/worker.js`, `js/db.js` | Impor latar belakang & penyimpanan lokal / background import & local storage |
| `js/app.js`, `js/pages1.js`, `js/pages2.js` | Halaman / pages |
| `js/report.js` | Executive report & laporan bilingual / reports |
| `js/ai.js` | Opsi AI / AI option |
| `sw.js`, `manifest.webmanifest` | Offline & instal / offline & install |
| `lib/` | Chart.js, ExcelJS, docx (lisensi di `licenses/`) |

Pustaka pihak ketiga: Chart.js (MIT), ExcelJS (MIT), docx (MIT), font Poppins & Barlow (SIL OFL 1.1).

## v1.2 — Unduh semua (Excel) / Download everything (Excel)
**ID:** Tombol "⬇ Unduh semua (Excel)" di Dashboard menghasilkan satu file Excel: sheet pertama adalah Dashboard (KPI, executive report ID/EN, 4 grafik, 10 prioritas coaching, cabang), lalu "Daftar isi" dan satu sheet untuk setiap menu sesuai urutan sidebar.
**EN:** The "⬇ Unduh semua (Excel)" button on the Dashboard produces a single Excel file: the first sheet is the Dashboard (KPIs, ID/EN executive report, 4 charts, top-10 coaching list, branches), followed by a table of contents and one sheet per sidebar menu, in sidebar order.

## v1.3 — Overspeed singkat tidak dihitung + perbaikan kestabilan / Short overspeed exclusion + stability fixes
**ID:**
- **Overspeed singkat diabaikan.** Kejadian overspeed berdurasi ≤ 10 detik (bisa diubah di Parameter → "Durasi minimum overspeed dihitung") tidak dihitung sebagai kasus. Caranya: unggah laporan **detail kejadian overspeed** dari Cartrack (laporan Speeding/Violations, .xlsx/.csv — kolom minimal: waktu kejadian, durasi atau waktu selesai, nama pengemudi atau plat) di menu **Unggah data**. Aplikasi mencocokkan tiap kejadian ke trip (plat/nama + waktu) dan menghitung ulang count overspeed trip tersebut; trip tanpa kejadian yang cocok tetap memakai count asli. Statistiknya tampil di Unggah data, Dashboard (Status data), dan laporan bilingual. Template CSV tersedia di Unggah data → Template. Tanpa file detail, perilaku lama berlaku (data mentah tidak memuat durasi).
- **Lebih tahan error:** impor otomatis pindah ke mode langsung bila Web Worker gagal (mis. dibuka dari file://); bila penyimpanan browser gagal (mode privat / penuh), data tetap dipakai untuk sesi itu dengan peringatan; error tak terduga kini muncul sebagai pesan (bukan layar diam); memilih file yang sama dua kali kini tetap terbaca; simpanan yang tertunda di-flush saat halaman ditutup.
**EN:**
- **Short overspeed ignored.** Overspeed events lasting ≤ 10 seconds (configurable in Parameter) are not counted as cases. Upload the Cartrack **speeding/violations detail report** (.xlsx/.csv with event time, duration or end time, and driver name or plate) on the **Unggah data** page; the app matches each event to its trip and recounts that trip's overspeed count. Trips without matching events keep their raw counts. Without the detail file, previous behaviour applies (raw data has no durations).
- **More robust:** imports fall back to the main thread when the Web Worker fails; browser-storage failures degrade to in-memory for the session with a clear warning; unexpected errors surface as messages; re-selecting the same file works; pending saves are flushed on page close.

### Mengaktifkan GitHub Pages / Enabling GitHub Pages
Repo → **Settings → Pages → Build and deployment → Source: Deploy from a branch → Branch: main, folder: / (root) → Save.** Tunggu ±1 menit, lalu buka https://rochiki.github.io/vt-score/ . Setiap push ke main otomatis memperbarui aplikasi; pengguna cukup muat ulang (aplikasi memakai network-first, jadi versi terbaru terambil saat online).
