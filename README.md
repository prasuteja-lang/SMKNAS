# Aplikasi Absensi Guru (Google Apps Script)

Aplikasi absensi guru berbasis Google Sheets + Apps Script Web App
(login, absen dengan radius lokasi, rekap bulanan/tahunan, hari libur, koreksi admin, bukti izin/sakit).

## Struktur

```
src/
├── appsscript.json      # manifest Apps Script
├── Index.html           # kerangka halaman (markup) + include file lain
├── Styles.html          # seluruh CSS
├── JS_Core.html         # utilitas, tema, jam, navigasi
├── JS_Home.html         # home, lokasi, absen guru, bukti izin/sakit
├── JS_Auth.html         # login, sesi, password, kelola akun
├── JS_Rekap.html        # rekap bulanan/tahunan, cetak, koreksi
├── JS_Admin.html        # pengaturan, hari libur, rapikan data, start
├── Config.gs            # konstanta, ID Drive, nama sheet, DEFAULTS
├── HolidayData.gs       # daftar libur nasional / cuti bersama
├── Main.gs              # menu spreadsheet, doGet, include(), setup
├── Helpers.gs           # fungsi bantu dasar
├── Settings.gs          # pengaturan & log absen
├── Auth.gs              # login, sesi, kelola akun
├── MonthSheets.gs       # sheet bulanan
├── Attendance.gs        # penulisan absen, API guru, cek lokasi
├── Admin.gs             # API admin (rekap, koreksi, pengaturan)
├── Holidays.gs          # hari libur + API libur
├── Maintenance.gs       # backup & rapikan data
├── MasterSheet.gs       # sheet "Nama Guru" & setup semua bulan
├── Proof.gs             # upload bukti izin/sakit ke Drive
├── WorkRequests.gs      # izin buka absen di hari libur
└── Logo.gs              # logo sekolah
```

Semua file `.gs` berbagi satu scope global, jadi urutan file tidak berpengaruh.

## Cara pakai

**Manual:** buka Extensions → Apps Script di spreadsheet, lalu buat file dengan nama yang sama
(`.gs` untuk script, `.html` untuk HTML — tanpa ekstensi saat memberi nama) dan salin isinya.

**Dengan clasp:**
```bash
npm i -g @google/clasp
clasp login
clasp clone <SCRIPT_ID> --rootDir src   # atau buat .clasp.json manual
clasp push
```

Setelah itu jalankan menu **📋 Absensi Guru → Siapkan aplikasi** di spreadsheet, lalu *Deploy → Web app*.

## Catatan
- Ubah ID file/folder Google Drive (kop, logo, folder bukti) di `src/Config.gs`.
- Jangan commit `.clasp.json` (sudah ada di `.gitignore`).
