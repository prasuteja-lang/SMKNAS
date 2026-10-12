/* ------------------------------------------------------------ */
/*  Pengaturan & log                                             */
/* ------------------------------------------------------------ */
function getSettings_() {
  const ss = ss_();
  let sh = ss.getSheetByName(SHEET_SET);
  if (!sh) {
    sh = ss.insertSheet(SHEET_SET);
    sh.getRange(1, 1, 1, 3).setValues([['Pengaturan', 'Nilai', 'Keterangan']]).setFontWeight('bold').setBackground('#e6f4f1');
    sh.getRange(2, 2, 60, 1).setNumberFormat('@');
    sh.getRange(2, 1, DEFAULTS.length, 3).setValues(DEFAULTS.map(function (d) { return [d[0], String(d[1]), d[2]]; }));
    sh.setColumnWidth(1, 190); sh.setColumnWidth(2, 280); sh.setColumnWidth(3, 520);
    sh.setFrozenRows(1);
  }
  const last = sh.getLastRow();
  const vals = last >= 2 ? sh.getRange(2, 1, last - 1, 2).getDisplayValues() : [];
  const map = {};
  vals.forEach(function (r) { if (r[0]) map[String(r[0]).trim()] = String(r[1]).trim(); });
  DEFAULTS.forEach(function (d) {
    if (map[d[0]] === undefined) {
      const r = sh.getLastRow() + 1;
      sh.getRange(r, 2).setNumberFormat('@');
      sh.getRange(r, 1, 1, 3).setValues([[d[0], String(d[1]), d[2]]]);
      map[d[0]] = String(d[1]);
    }
  });
  // migrasi: ganti nama placeholder lama dengan nama sekolah yang benar
  if (map.NamaSekolah === 'NAMA SEKOLAH ANDA') {
    for (let i = 0; i < vals.length; i++) {
      if (String(vals[i][0]).trim() === 'NamaSekolah') { sh.getRange(2 + i, 2).setValue('SMK NASIONAL BANDUNG'); break; }
    }
    map.NamaSekolah = 'SMK NASIONAL BANDUNG';
  }
  const m = /(\d{1,2})[:.](\d{2})/.exec(map.JamMasuk || '');
  map.JamMasuk = m ?('0' + m[1]).slice(-2) + ':' + m[2] : '07:00';
  map.HariKerja = String(map.HariKerja) === '6' ? '6' : '5';
  map.CutiBersama = /masuk/i.test(map.CutiBersama || '') ? 'Masuk' : 'Libur';
  map.TahunAjaranMulai = parseInt(map.TahunAjaranMulai, 10) || new Date().getFullYear();
  map.LokasiWajib = /^(tidak|no|off|0|false)/i.test(map.LokasiWajib || '') ? 'Tidak' : 'Ya';
  map.RadiusMeter = Math.max(10, parseInt(map.RadiusMeter, 10) || 100);
  map.MaxAkurasi = Math.max(10, parseInt(map.MaxAkurasi, 10) || 100);
  map.PasswordAwal = String(map.PasswordAwal || '123456');
  return map;
}

function publicSettings_(set) {
  const o = {};
  Object.keys(set).forEach(function (k) { if (k !== 'PinAdmin') o[k] = set[k]; });
  return o;
}

const LOG_HEAD = ['Waktu Catat', 'Tanggal', 'Jam', 'No', 'Nama', 'Status', 'Catatan', 'Menit Telat', 'Sumber', 'ID Login', 'Lat', 'Lng', 'Jarak (m)', 'Akurasi (m)', 'Bukti'];
function ensureLog_() {
  const ss = ss_();
  let sh = ss.getSheetByName(SHEET_LOG);
  if (!sh) {
    sh = ss.insertSheet(SHEET_LOG);
    sh.getRange(1, 1, 1, 15).setValues([LOG_HEAD]).setFontWeight('bold').setBackground('#e6f4f1');
    sh.setFrozenRows(1);
    sh.setColumnWidths(1, 15, 110); sh.setColumnWidth(5, 240); sh.setColumnWidth(7, 380); sh.setColumnWidth(15, 300);
  } else {
    if (sh.getMaxColumns() < 15) sh.insertColumnsAfter(sh.getMaxColumns(), 15 - sh.getMaxColumns());
    if (!sh.getRange(1, 10).getDisplayValue()) sh.getRange(1, 10, 1, 5).setValues([LOG_HEAD.slice(9, 14)]).setFontWeight('bold').setBackground('#e6f4f1');
    if (!sh.getRange(1, 15).getDisplayValue()) sh.getRange(1, 15).setValue('Bukti').setFontWeight('bold').setBackground('#e6f4f1');
  }
  return sh;
}

function appendLog_(no, name, dateStr, timeStr, st, note, minLate, source, extra) {
  const sh = ensureLog_(), x = extra || {};
  const r = sh.getLastRow() + 1;
  const rng = sh.getRange(r, 1, 1, 15);
  rng.setNumberFormats([['yyyy-mm-dd hh:mm:ss', '@', '@', '0', '@', '@', '@', '0', '@', '@', '0.000000', '0.000000', '0', '0', '@']]);
  rng.setValues([[new Date(), dateStr, timeStr, no || '', name, st || 'KOSONG', note || '', minLate || 0, source,
    x.user || '', x.lat == null ? '' : x.lat, x.lng == null ? '' : x.lng, x.dist == null ? '' : x.dist, x.acc == null ? '' : x.acc, x.proof || '']]);
}

