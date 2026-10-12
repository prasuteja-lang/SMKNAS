/* ------------------------------------------------------------ */
/*  Menu & web app                                               */
/* ------------------------------------------------------------ */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('📋 Absensi Guru')
    .addItem('1. Siapkan aplikasi (jalankan sekali)', 'setupApp')
    .addItem('Pasang warna otomatis di semua sheet bulan', 'formatAllMonths')
    .addItem('Buat sheet bulan ini', 'createCurrentMonth')
    .addSeparator()
    .addItem('Buat & format semua sheet bulan (Agu–Jul)', 'setupMonthsMenu')
    .addItem('Rapikan data semua sheet bulan', 'cleanDataMenu')
    .addItem('Sinkronkan libur nasional', 'syncNationalMenu')
    .addItem('Aktifkan penandaan BA otomatis (harian)', 'installDailyTrigger')
    .addSeparator()
    .addItem('Buat akun login untuk semua guru', 'syncAccountsMenu')
    .addItem('Tampilkan alamat aplikasi', 'showUrl')
    .addSeparator()
    .addItem('API: buat / ganti API Key (baca database)', 'createApiKeyMenu')
    .addItem('API: hapus API Key', 'revokeApiKeyMenu')
    .addSeparator()
    .addItem('Izinkan akses Drive (folder bukti)', 'authorizeDrive')
    .addToUi();
}

function doGet(e) {
  // ?action=...  -> JSON API (lihat Api.gs); selain itu -> halaman web
  if (e && e.parameter && e.parameter.action) return apiGateway_(e.parameter, 'GET');
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle('Absensi Guru')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** JSON API: POST dengan body JSON {action, token, args:[...]} (lihat Api.gs / API.md). */
function doPost(e) { return apiPost_(e); }

/** Menyisipkan file HTML lain (CSS / JS) ke dalam Index.html:  <?!= include('Styles'); ?> */
function include(name) {
  return HtmlService.createHtmlOutputFromFile(name).getContent();
}

function setupApp() {
  const set = getSettings_();
  ensureLog_();
  getAccSheet_();
  syncAccounts_();   // buat/perbarui akun dari sheet Nama Guru
  let kop = 'Kop surat: OK (terbaca).';
  try { DriveApp.getFileById(KOP_FILE_ID).getName(); }
  catch (e) { kop = 'Kop surat: BELUM bisa dibaca → pastikan akun Anda punya akses ke file kop di Drive. (' + e.message + ')'; }
  SpreadsheetApp.getUi().alert('Siap!',
    'Sheet "Pengaturan", "LogAbsen" dan "Akun" sudah tersedia.\n' + kop +
    '\n\nAkun admin awal: ID "admin" (password = PIN lama / Password Awal). Akun guru/admin lain: sesuai kolom id_login di sheet Nama Guru (password = Password Awal). Semua wajib ganti password saat login pertama.' +
    '\n\nLangkah berikutnya: isi sheet Pengaturan (nama sekolah, tahun ajaran, lokasi sekolah), lalu Deploy › New deployment › Web app.',
    SpreadsheetApp.getUi().ButtonSet.OK);
}

function showUrl() {
  const url = ScriptApp.getService().getUrl();
  SpreadsheetApp.getUi().alert(url ? 'Alamat aplikasi:\n\n' + url : 'Aplikasi belum di-deploy. Buka Deploy › New deployment › Web app.');
}

function createCurrentMonth() {
  const set = getSettings_(), n = nowInfo_();
  if (monthYear_(set, n.m) !== n.y) {
    SpreadsheetApp.getUi().alert('Tahun ajaran di sheet Pengaturan belum sesuai dengan tanggal hari ini.');
    return;
  }
  ensureMonthSheet_(n.m, set);
  SpreadsheetApp.getUi().alert('Sheet bulan ' + MONTHS[n.m] + ' siap.');
}

function formatAllMonths() {
  const set = getSettings_();
  existingMonths_(set).forEach(function (mo) {
    const sh = getMonthSheet_(mo.idx);
    applyFormat_(sh);
    refreshMonthMarks_(sh, mo.idx, set, true);
  });
  SpreadsheetApp.getUi().alert('Warna otomatis terpasang: H hijau, I biru, S oranye, T merah, TJ abu-abu, A merah tua (alpa), BA ungu (belum absen), L merah muda (libur). Kode di luar daftar ditandai kuning.');
}

