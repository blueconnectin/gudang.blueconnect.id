# Gudang Scan

Aplikasi web responsif untuk mencatat inventaris dengan pemindai barcode kamera ponsel, Google Apps Script sebagai backend, Google Sheets sebagai database, serta ekspor Excel `.xlsx`.

## Jalankan antarmuka

Buka `index.html` melalui HTTPS (misalnya GitHub Pages) atau localhost. Izin kamera hanya tersedia dalam secure context. Pemindai memakai `html5-qrcode`, ekspor memakai SheetJS; keduanya dimuat dari CDN. Jika library ekspor gagal dimuat, aplikasi menyediakan unduhan CSV sebagai fallback.

## Hubungkan Google Sheets

1. Buat Google Spreadsheet untuk database.
2. Buka **Extensions → Apps Script**, tempel isi `Code.gs` ke editor, dan simpan.
3. Di Apps Script, buka **Project Settings → Script Properties** dan tambahkan `SPREADSHEET_ID` dengan ID spreadsheet (bagian antara `/d/` dan `/edit` pada URL).
4. Pilih **Deploy → New deployment → Web app**. Jalankan sebagai akun Anda; atur akses sesuai kebijakan organisasi. Salin URL deployment berakhiran `/exec`.
5. Isi nilai `API_URL` di bagian awal skrip pada `index.html` dengan URL tersebut. Publikasikan ulang frontend.
6. Buka frontend. Status **Sheets tersambung** menandakan data telah berhasil dimuat.

Web app Apps Script yang menerima akses publik dapat menulis ke spreadsheet milik akun deployer. Bagikan URL frontend hanya kepada pengguna yang diizinkan; batasi akses deployment bila sesuai kebutuhan organisasi.

## Penggunaan

Pindai barcode menggunakan kamera belakang, lengkapi nama barang, jumlah, dan catatan, lalu simpan. Barcode juga dapat dimasukkan manual. Daftar dan pencarian ditampilkan pada halaman; tombol **Ekspor Excel** menghasilkan berkas `.xlsx`. Tanpa konfigurasi API, pencatatan berjalan lokal pada browser yang sama (`localStorage`) dan tidak tersinkron antarperangkat.

## Deploy otomatis dengan clasp

Source Apps Script siap di folder `apps-script/`. Pasang CLI `@google/clasp`, login ke akun Google yang memiliki akses, lalu dari folder tersebut jalankan `clasp create --type webapp --title "Gudang Scan"` (atau `clasp clone <SCRIPT_ID>` untuk project yang sudah ada). Atur Script Property `SPREADSHEET_ID` di Apps Script. Dari root repository jalankan `./deploy-appscript.sh`; perintah ini mengunggah source lalu membuat deployment web app. Untuk memperbarui deployment web app yang sama, set `CLASP_DEPLOYMENT_ID` ke deployment ID sebelum menjalankannya. Deployment pertama mengeluarkan URL web app; masukkan URL `/exec` tersebut ke `API_URL` di `index.html`. Otorisasi dan pengaturan siapa yang dapat mengakses web app tetap dikendalikan oleh akun Google di Apps Script.

## Deploy otomatis dari GitHub ke Apps Script

Workflow `.github/workflows/deploy-appscript.yml` mengunggah folder `apps-script/` dan menerbitkan versi web app setiap ada perubahan pada folder itu di branch `main` (atau saat dijalankan manual di tab **Actions**). Workflow memakai GitHub Actions Secrets berikut:

- `SCRIPT_ID`: Script ID Apps Script tujuan.
- `CLASP_CREDENTIALS`: isi file autentikasi `.clasprc.json` dari akun Google yang berwenang pada project.
- `DEPLOYMENT_ID` (opsional, disarankan setelah deployment pertama): ID deployment web app yang akan diperbarui supaya URL tetap sama.

Untuk membuat `CLASP_CREDENTIALS`, di komputer yang sudah berhasil login dengan clasp, tampilkan isi file `.clasprc.json` (Windows: `%USERPROFILE%\\.clasprc.json`; macOS/Linux: `~/.clasprc.json`) dan salin seluruh JSON ke **GitHub → repository → Settings → Secrets and variables → Actions → New repository secret** bernama `CLASP_CREDENTIALS`. Jangan commit file kredensial itu atau kirimkan isinya melalui chat. Tambahkan dua secret lainnya di tempat yang sama.

Workflow mendeploy web app, tetapi tidak membuat/mengatur spreadsheet atau Script Property `SPREADSHEET_ID`. Pastikan property tersebut sudah diatur di Apps Script. Jalankan workflow sekali dari tab **Actions**. Jika deployment pertama membuat deployment baru, salin ID deployment dari output Actions, simpan sebagai secret `DEPLOYMENT_ID`, lalu jalankan ulang workflow untuk deployment yang konsisten. Pastikan setting **Execute as** dan **Who has access** pada deployment Apps Script sesuai kebutuhan. URL `/exec` dapat dipakai sebagai `API_URL` pada frontend.

### Penyimpanan cepat dengan antrean

Tombol Simpan memasukkan catatan ke antrean browser dan langsung menyiapkan formulir berikutnya. Kolom Status membedakan Menunggu, Mengirim, dan Tersimpan (setelah dikonfirmasi backend). Antrean yang gagal dikirim tetap disimpan pada browser yang sama, dicoba ulang otomatis, dan dapat dikirim ulang melalui tombol Kirim ulang. Membuka kembali halaman pada browser tersebut melanjutkan antrean; menghapus data situs/browser juga menghapus antrean lokal. Ekspor hanya memuat data yang sudah dikonfirmasi. Backend menulis hingga 20 catatan per permintaan dan mencegah duplikasi pengiriman ulang berdasarkan ID catatan. Pemindaian barcode yang sama sebagai catatan baru tetap memiliki ID berbeda.

### Kategori ONT

Catatan baru memilih kategori **ONT Rusak**, **ONT Dismantle**, atau **ONT Baru**, terpisah dari jenis Huawei/Fiberhome. Kategori tersedia pada log, pencarian, antrean penyimpanan dan ekspor Excel. Backend menambahkan kolom Kategori ONT jika belum ada. Catatan dan antrean lama yang tidak memiliki kategori tetap kosong tanpa penetapan otomatis.

### Edit dan hapus catatan

Kolom Aksi menyediakan Edit dan Hapus untuk catatan yang sudah tersimpan. Edit memperbarui barcode, jenis, kategori, jumlah dan catatan sambil mempertahankan ID serta waktu pencatatan asli. Hapus meminta konfirmasi sebelum menghapus baris pada Google Sheets. Catatan yang masih di antrean belum dapat diedit/dihapus sampai sinkronisasi dikonfirmasi. Perubahan pada layar berlaku setelah respons backend berhasil; kegagalan jaringan ditampilkan tanpa menghapus data di layar. Jika konfirmasi terlalu lama, refresh log untuk memeriksa hasil operasi.

### Keterangan Close WO

Kolom Catatan pada aplikasi diganti dengan dropdown Keterangan Close WO: Sudah atau Belum. Status wajib dipilih untuk input baru dan saat mengedit; muncul pada log, pencarian, serta ekspor Excel. Backend menambahkan kolom Keterangan Close WO pada Sheets tanpa menghapus catatan lama. Data/antrean lama yang belum memiliki status tidak diberi nilai otomatis dan tetap dapat disinkronkan.

### Rekap QTY tersimpan

Rekap menampilkan jumlah QTY untuk ONT Rusak, ONT Dismantle, dan ONT Baru dari seluruh catatan tersimpan. Antrean yang belum dikonfirmasi tidak dihitung; filter pencarian tidak mengubah rekap. Rekap diperbarui setelah sinkronisasi, edit, hapus, dan refresh berhasil. Saat offline, rekap memakai cache data tersimpan terakhir. Catatan lama tanpa kategori tidak dimasukkan ke kategori mana pun.
