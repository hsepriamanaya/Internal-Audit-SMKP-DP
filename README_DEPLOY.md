# Deploy Web Audit SMKP

Arsitektur: Google Sheets (DB) + Google Apps Script (Code.gs) + file statis di GitHub Pages. Tanpa server sendiri.

## 1. Backend (Google Sheets + Apps Script)
1. Buat Google Sheet baru, mis. "DB Audit SMKP". Sheet "Data" dibuat otomatis (key, value, updated).
2. Extensions > Apps Script, tempel isi `Code.gs`.
3. Project Settings > Script Properties, tambahkan tiga kunci (acak, panjang):
   - `ADMIN_KEY`   : tim auditor (baca/tulis semua)
   - `READER_KEY` : pembaca (baca saja)
   - `AUDITEE_KEY`: auditi (hanya menulis tindak lanjut: key berakhiran `/tl_auditi`)
4. Deploy > New deployment > Web app. Execute as: Me. Who has access: Anyone. Salin URL `/exec`.
5. Setiap Code.gs diubah: Deploy > Manage deployments > Edit > Version: New version.

## 2. Frontend (GitHub Pages)
Unggah seluruh isi folder `web/` ke repo (branch main, root atau /docs) lalu aktifkan Pages.
File: app.html, core.js, docgen.js, master_web.js, vendor/docx.iife.js, vendor/jszip.min.js (tanpa CDN, bisa offline).
Buka `app.html` > tombol "Koneksi" > isi URL Web App; masukkan kunci akses sesuai peran. Kosongkan URL = mode lokal (data di browser).

## 3. Catatan operasional
- Sesi otomatis keluar setelah 30 menit tidak aktif; kunci akses hanya disimpan di sessionStorage.
- Batas satu sel Sheets ~50.000 karakter: data dipecah per key (ba, plan, nilai, lap, tl_auditi, tl_verif, snap/<id>, index). Gambar besar di tab Lampiran dikompres otomatis; hindari >10 gambar.
- Arsip periode: tab "6. Arsip & Perbandingan" > "Arsipkan snapshot", lalu "+ Audit baru" untuk periode berikutnya.
- Keamanan: kunci dicek di server (LockService untuk tulis). Auditi tidak bisa menulis key selain tindak lanjutnya; pembaca ditolak server saat menulis.
- Cadangan: File > Make a copy pada Sheet secara berkala; atau salin kolom value.
