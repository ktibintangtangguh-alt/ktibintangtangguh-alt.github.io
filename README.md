# Finansialku

Aplikasi perencana keuangan pribadi (PWA): pemasukan, pengeluaran, anggaran bulanan,
investasi, target tabungan, dan kalkulator investasi.

- Tampilan aplikasi di-hosting di GitHub Pages dan bisa diinstal di Android (Chrome > Instal aplikasi).
- Data disimpan di Google Sheet "Finansialku — Data" di Google Drive pemilik, lewat Google Apps Script
  (`apps-script/Code.gs`). Setiap permintaan dilindungi kunci rahasia.
- Bisa dipakai tanpa internet: perubahan disimpan di HP lalu dikirim otomatis saat online.
- Banyak pengguna: admin membuat kode undangan; setiap pengguna mendapat Google Sheet sendiri di Drive admin.

## Memasang backend
1. Ganti isi `Kode.gs` di project Apps Script dengan `apps-script/Code.gs` (file Index boleh dihapus).
2. Terapkan > Kelola deployment > Edit > Versi: Versi baru, Jalankan sebagai: Saya, Akses: Siapa saja > Terapkan.
3. Buka aplikasi, tempel URL `/exec`, ketuk Hubungkan.

Lupa kunci? Jalankan fungsi `resetKunci` di editor Apps Script, lalu hubungkan ulang.
