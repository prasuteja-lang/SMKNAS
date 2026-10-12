/* ------------------------------------------------------------ */
/*  Sheet "Nama Guru" (induk) & pembuatan sheet bulan            */
/* ------------------------------------------------------------ */
function getGuruSheet_() {
  const ss = ss_();
  let sh = ss.getSheetByName(SHEET_GURU);
  if (!sh) {
    sh = ss.insertSheet(SHEET_GURU);
    sh.getRange(1, 1, 1, 5).setValues([['NO', 'NAMA', 'Status', 'Job', 'id_login']]);
    const src = AY_ORDER.map(getMonthSheet_).filter(hasNames_)[0];
    if (src) {
      const names = src.getRange(FIRST_ROW, 2, src.getLastRow() - FIRST_ROW + 1, 1).getDisplayValues()
        .map(function (r) { return String(r[0]).replace(/\s+/g, ' ').trim(); }).filter(String);
      if (names.length) sh.getRange(2, 1, names.length, 5).setValues(names.map(function (n, i) { return [i + 1, n, 'Aktif', 'Guru', '']; }));
    }
  }
  return sh;
}

/** Daftar induk guru: [{ row, name, status, job, loginId }]. Tidak membuat sheet jika belum ada. */
function readMaster_() {
  const sh = ss_().getSheetByName(SHEET_GURU);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues().map(function (r, i) {
    return { row: 2 + i, name: String(r[1] == null ? '' : r[1]).replace(/\s+/g, ' ').trim(), status: String(r[2] == null ? '' : r[2]).trim(),
      job: normJob_(r[3]), loginId: String(r[4] == null ? '' : r[4]).trim() };
  }).filter(function (t) { return t.name; });
}

/** Nama guru aktif (tanpa duplikat: orang yang punya 2 akun, mis. Admin & Guru, hanya satu baris di sheet bulan). */
function activeTeachers_() {
  const seen = {};
  return readMaster_().filter(function (t) { return !/^(non|tidak|keluar|pindah|resign|off)/i.test(t.status); })
    .map(function (t) { return t.name; })
    .filter(function (n) { const k = normName_(n); if (seen[k]) return false; seen[k] = true; return true; });
}

function tidyMaster_() {
  const sh = getGuruSheet_(), last = sh.getLastRow();
  if (last >= 2) {
    const rng = sh.getRange(2, 1, last - 1, 5), v = rng.getValues();
    let no = 0;
    v.forEach(function (r) {
      const nm = String(r[1] == null ? '' : r[1]).replace(/\s+/g, ' ').trim();
      r[1] = nm; r[0] = nm ? ++no : ''; r[2] = nm ? (String(r[2] || '').trim() || 'Aktif') : '';
      r[3] = nm ? normJob_(r[3]) : ''; r[4] = String(r[4] == null ? '' : r[4]).trim();
    });
    rng.setValues(v);
  }
  const n = Math.max(sh.getLastRow(), 2);
  sh.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#e6f4f1').setHorizontalAlignment('center');
  sh.getRange(1, 1, n, 5).setBorder(true, true, true, true, true, true, '#cfd8dc', SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange(2, 1, n - 1, 1).setHorizontalAlignment('center');
  sh.getRange(2, 3, Math.max(sh.getMaxRows() - 1, 1), 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['Aktif', 'Nonaktif'], true).setAllowInvalid(false).build());
  sh.getRange(2, 4, Math.max(sh.getMaxRows() - 1, 1), 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['Guru', 'Admin'], true).setAllowInvalid(false).build());
  sh.setColumnWidth(1, 45); sh.setColumnWidth(2, 300); sh.setColumnWidth(3, 90); sh.setColumnWidth(4, 80); sh.setColumnWidth(5, 130);
  sh.setFrozenRows(1);
  return sh;
}

/** Isi nama pada sheet bulan yang masih kosong: dari sheet Nama Guru, atau dari bulan terdekat. */
function seedNames_(sh, mIdx) {
  const act = activeTeachers_();
  if (!act.length) return copyNamesFrom_(sh, findSourceSheet_(mIdx));
  sh.getRange(FIRST_ROW, 1, act.length, 2).setValues(act.map(function (n, i) { return [i + 1, n]; }));
  return act.length;
}

/** Tambahkan guru aktif yang belum ada di sheet bulan (data yang sudah ada tidak diubah). Mengembalikan jumlah baris baru. */
function syncMonthNames_(sh) {
  const act = activeTeachers_();
  if (!act.length || !sh) return 0;
  const last = sh.getLastRow();
  const rows = last >= FIRST_ROW ? sh.getRange(FIRST_ROW, 1, last - FIRST_ROW + 1, 2).getValues() : [];
  const have = {};
  let lastName = FIRST_ROW - 1, no = 0;
  rows.forEach(function (r, i) {
    const k = normName_(r[1]);
    if (k) { have[k] = true; lastName = FIRST_ROW + i; }
    no = Math.max(no, Number(r[0]) || 0);
  });
  const vals = act.filter(function (n) { return !have[normName_(n)]; }).map(function (n) { return [++no, n]; });
  if (!vals.length) return 0;
  const start = lastName + 1, need = start + vals.length - 1;
  if (need > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), need - sh.getMaxRows() + 5);
  sh.getRange(start, 1, vals.length, 2).setValues(vals);
  if (lastName >= FIRST_ROW) sh.getRange(lastName, 1, 1, REKAP_COL).copyTo(sh.getRange(start, 1, vals.length, REKAP_COL), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  return vals.length;
}

/** Sheet dianggap sudah berisi absensi jika ada kode H/I/S/T/TJ pada hari kerja biasa (kolom libur diabaikan). */
function sheetHasData_(sh, mIdx, set) {
  const last = sh.getLastRow();
  if (last < FIRST_ROW) return false;
  const year = monthYear_(set, mIdx), days = daysIn_(year, mIdx), hol = readHolidays_();
  const vals = sh.getRange(FIRST_ROW, DAY_COL, last - FIRST_ROW + 1, 31).getValues();
  const work = [];
  for (let d = 1; d <= days; d++) work.push(!dayInfo_(set, hol, year, mIdx, d).off);
  return vals.some(function (row) {
    return row.some(function (c, i) { return work[i] && VALID.indexOf(normalizeCode_(c)) >= 0; });
  });
}

function rebuildNames_(sh) {
  const act = activeTeachers_();
  if (!act.length) return 0;
  const last = sh.getLastRow();
  if (last >= FIRST_ROW) sh.getRange(FIRST_ROW, 1, last - FIRST_ROW + 1, REKAP_COL).clearContent();
  if (FIRST_ROW + act.length > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), FIRST_ROW + act.length - sh.getMaxRows());
  sh.getRange(FIRST_ROW, 1, act.length, 2).setValues(act.map(function (n, i) { return [i + 1, n]; }));
  return act.length;
}

/** Buat (jika belum ada) dan seragamkan format 12 sheet bulan, Agustus s/d Juli. */
function setupAllMonths_() {
  const set = getSettings_();
  HOL_CACHE_ = null;
  readHolidays_();
  tidyMaster_();
  const out = [];
  AY_ORDER.forEach(function (i) {
    const existed = !!getMonthSheet_(i);
    const sh = ensureMonthSheet_(i, set), year = monthYear_(set, i);
    let mode;
    if (!sheetHasData_(sh, i, set)) { rebuildNames_(sh); mode = existed ? 'belum ada absensi → daftar guru diisi ulang dari Nama Guru' : 'dibuat baru'; }
    else { const a = syncMonthNames_(sh); mode = 'data absensi dipertahankan' + (a ? ', +' + a + ' guru baru' : ''); }
    sh.getRange(1, 1).setValue(year);
    sh.getRange(1, 2).setValue('REKAP ABSENSI GURU — ' + MONTHS[i].toUpperCase() + ' ' + year).setFontWeight('bold').setFontSize(14);
    sh.getRange(2, REKAP_COL).setValue('REKAP');
    applyFormat_(sh);
    beautifySheet_(sh);
    refreshMonthMarks_(sh, i, set, true);
    out.push(MONTHS[i] + ' ' + year + ': ' + mode + ' (' + Math.max(sh.getLastRow() - FIRST_ROW + 1, 0) + ' guru)');
  });
  return out;
}

function setupMonthsMenu() {
  const ui = SpreadsheetApp.getUi();
  if (ui.alert('Buat & format sheet bulan',
    'Semua sheet bulan (Agustus–Juli) akan dibuat bila belum ada dan diseragamkan formatnya. Daftar guru diambil dari sheet "Nama Guru" (yang berstatus Aktif). Sheet yang sudah berisi absensi TIDAK dihapus datanya.\n\nLanjutkan?', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  ui.alert('Selesai', setupAllMonths_().join('\n'), ui.ButtonSet.OK);
}

function apiSetupMonths(token) {
  requireAdmin_(token);
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try { return { ok: true, summary: setupAllMonths_().join('\n') }; }
  finally { lock.releaseLock(); }
}


