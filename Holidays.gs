/* ------------------------------------------------------------ */
/*  Hari libur (nasional + libur khusus sekolah)                 */
/* ------------------------------------------------------------ */
let HOL_CACHE_ = null;

function pad2_(n) { return ('0' + n).slice(-2); }
function iso_(y, m, d) { return y + '-' + pad2_(m + 1) + '-' + pad2_(d); }
function dowOf_(y, m, d) { return new Date(y, m, d).getDay(); }

function parseIso_(s) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(s || '').trim());
  if (!m) return null;
  const y = +m[1], mo = +m[2] - 1, d = +m[3], dt = new Date(y, mo, d);
  return (dt.getFullYear() === y && dt.getMonth() === mo && dt.getDate() === d) ? { y: y, m: mo, d: d } : null;
}

function holType_(s) {
  s = String(s || '');
  if (/cuti/i.test(s)) return 'Cuti Bersama';
  if (/khusus|sekolah/i.test(s)) return 'Khusus';
  return 'Nasional';
}

function cellIso_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  const s = String(v == null ? '' : v).trim();
  if (parseIso_(s)) return s;
  const m = /^(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{4})$/.exec(s);   // dd/mm/yyyy yang diketik manual
  if (m) { const r = m[3] + '-' + pad2_(+m[2]) + '-' + pad2_(+m[1]); return parseIso_(r) ? r : ''; }
  return '';
}

function seedNational_(sh) {
  const have = {}, last = sh.getLastRow();
  if (last >= 2) sh.getRange(2, 1, last - 1, 1).getValues().forEach(function (r) { const k = cellIso_(r[0]); if (k) have[k] = true; });
  const add = NATIONAL.filter(function (h) { return !have[h[0]]; })
    .map(function (h) { return [h[0], h[1], h[2] === 'C' ? 'Cuti Bersama' : 'Nasional', 'Ya']; });
  if (!add.length) return 0;
  const start = Math.max(sh.getLastRow(), 1) + 1, need = start + add.length - 1;
  if (need > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), need - sh.getMaxRows() + 20);
  sh.getRange(start, 1, add.length, 1).setNumberFormat('@');
  sh.getRange(start, 1, add.length, 4).setValues(add);
  sh.getRange(2, 1, sh.getLastRow() - 1, 4).sort(1);
  return add.length;
}

function getHolSheet_() {
  const ss = ss_();
  let sh = ss.getSheetByName(SHEET_HOL);
  if (!sh) {
    sh = ss.insertSheet(SHEET_HOL);
    sh.getRange(1, 1, 1, 4).setValues([['Tanggal', 'Keterangan', 'Jenis', 'Aktif']]).setFontWeight('bold').setBackground('#e6f4f1');
    sh.getRange(2, 1, sh.getMaxRows() - 1, 1).setNumberFormat('@');
    sh.setColumnWidth(1, 110); sh.setColumnWidth(2, 380); sh.setColumnWidth(3, 130); sh.setColumnWidth(4, 70);
    sh.setFrozenRows(1);
    seedNational_(sh);
  }
  return sh;
}

/** Peta tanggal (yyyy-MM-dd) -> { row, name, type, active }. */
function readHolidays_() {
  if (HOL_CACHE_) return HOL_CACHE_;
  const sh = getHolSheet_(), map = {}, last = sh.getLastRow();
  if (last >= 2) {
    sh.getRange(2, 1, last - 1, 4).getValues().forEach(function (r, i) {
      const k = cellIso_(r[0]);
      if (!k) return;
      map[k] = {
        row: 2 + i,
        name: String(r[1] || '').trim() || 'Libur',
        type: holType_(r[2]),
        active: !/^(tidak|no|0|false|n)\b/i.test(String(r[3] == null ? '' : r[3]).trim())
      };
    });
  }
  HOL_CACHE_ = map;
  return map;
}

function syncNationalHolidays_() {
  const n = seedNational_(getHolSheet_());
  HOL_CACHE_ = null;
  return n;
}

/** Info satu tanggal: akhir pekan, libur (nasional / cuti bersama / khusus), dan apakah hari itu libur. */
function dayInfo_(set, hol, y, m, d) {
  const iso = iso_(y, m, d), h = hol[iso], dw = dowOf_(y, m, d);
  const weekend = dw === 0 || (dw === 6 && String(set.HariKerja) !== '6');
  let holiday = null;
  if (h && h.active && !(h.type === 'Cuti Bersama' && set.CutiBersama === 'Masuk')) holiday = { name: h.name, type: h.type };
  return {
    iso: iso, dow: dw, weekend: weekend, holiday: holiday, off: weekend || !!holiday,
    name: holiday ? holiday.name : (weekend ? (dw === 0 ? 'Hari Minggu' : 'Hari Sabtu') : '')
  };
}

/** Warnai kolom tanggal (akhir pekan / libur) dan isi "L" untuk hari libur pada sheet bulan. */
function refreshMonthMarks_(sh, mIdx, set, force) {
  const year = monthYear_(set, mIdx), days = daysIn_(year, mIdx), hol = readHolidays_();
  const last = sh.getLastRow(), info = [];
  let sig = set.HariKerja + '|' + set.CutiBersama + '|' + year + '|' + last + '|';
  for (let d = 1; d <= 31; d++) {
    const inf = d <= days ? dayInfo_(set, hol, year, mIdx, d) : null;
    info.push(inf);
    sig += !inf ? 'x' : inf.holiday ? 'H' + inf.holiday.name + ';' : inf.weekend ? 'W' : '-';
  }
  const props = PropertiesService.getDocumentProperties(), key = 'hl_' + mIdx;
  if (!force && props.getProperty(key) === sig) return;

  const hb = [[]], nt = [[]];
  info.forEach(function (inf) {
    hb[0].push(!inf ? '#bdbdbd' : inf.holiday ? '#f4b6b0' : inf.weekend ? '#fbd5d1' : '#e6f4f1');
    nt[0].push(inf && inf.holiday ? inf.holiday.name + ' (' + inf.holiday.type + ')' : '');
  });
  const hr = sh.getRange(2, DAY_COL, 1, 31);
  hr.setBackgrounds(hb); hr.setNotes(nt);

  const bodyRows = Math.max(sh.getMaxRows() - 2, 1);
  const colBg = info.map(function (inf) { return !inf ? '#d9d9d9' : inf.holiday ? '#fde2e0' : inf.weekend ? '#f1f3f4' : '#ffffff'; });
  const bgs = [];
  for (let r = 0; r < bodyRows; r++) bgs.push(colBg.slice());
  sh.getRange(3, DAY_COL, bodyRows, 31).setBackgrounds(bgs);

  if (last >= FIRST_ROW) {
    const rows = last - FIRST_ROW + 1;
    const names = sh.getRange(FIRST_ROW, 2, rows, 1).getDisplayValues();
    const rng = sh.getRange(FIRST_ROW, DAY_COL, rows, 31), vals = rng.getValues();
    let changed = false;
    for (let i = 0; i < rows; i++) {
      const has = String(names[i][0]).trim() !== '';
      for (let d = 1; d <= 31; d++) {
        const inf = info[d - 1], c = vals[i][d - 1], raw = String(c == null ? '' : c).trim().toUpperCase();
        const wantL = has && inf && inf.holiday && !inf.weekend;
        if (wantL && (raw === '' || raw === 'BA')) { vals[i][d - 1] = 'L'; changed = true; }
        else if (!wantL && raw === 'L') { vals[i][d - 1] = ''; changed = true; }
      }
    }
    if (changed) rng.setValues(vals);
  }
  props.setProperty(key, sig);
}

/**
 * Isi "BA" (Belum Absen) pada sel kosong di hari kerja yang SUDAH LEWAT (hari ini tidak ikut: guru masih bisa absen).
 * Hari hanya dihitung bila minimal satu guru tercatat pada hari itu — jadi hari yang belum dipakai / libur
 * yang belum didaftarkan tidak membanjiri sheet dengan BA. Kolom hari ini & seterusnya tidak pernah disentuh.
 */
function markBelumAbsen_(sh, mIdx, set) {
  if (!sh) return;
  const n = nowInfo_(), year = monthYear_(set, mIdx);
  const cur = year * 12 + mIdx, now = n.y * 12 + n.m;
  if (cur > now) return;
  const last = sh.getLastRow();
  if (last < FIRST_ROW) return;
  const until = cur < now ? daysIn_(year, mIdx) : n.day - 1;
  if (until < 1) return;
  const hol = readHolidays_(), rows = last - FIRST_ROW + 1;
  const names = sh.getRange(FIRST_ROW, 2, rows, 1).getDisplayValues();
  const rng = sh.getRange(FIRST_ROW, DAY_COL, rows, until), vals = rng.getValues();
  let changed = false;
  for (let d = 1; d <= until; d++) {
    if (dayInfo_(set, hol, year, mIdx, d).off) continue;
    let used = false;
    for (let i = 0; i < rows; i++) {
      const code = normalizeCode_(vals[i][d - 1]);
      if (code && code !== 'L' && code !== 'BA') { used = true; break; }
    }
    if (!used) continue;
    for (let i = 0; i < rows; i++) {
      if (String(names[i][0]).trim() && String(vals[i][d - 1] == null ? '' : vals[i][d - 1]).trim() === '') { vals[i][d - 1] = 'BA'; changed = true; }
    }
  }
  if (changed) rng.setValues(vals);
}

/** Dijalankan pemicu harian: perbarui BA di semua sheet bulan supaya sheet langsung benar walau aplikasi belum dibuka. */
function dailyMarkBA() {
  const set = getSettings_();
  HOL_CACHE_ = null;
  existingMonths_(set).forEach(function (mo) { markBelumAbsen_(getMonthSheet_(mo.idx), mo.idx, set); });
}

function installDailyTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'dailyMarkBA') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('dailyMarkBA').timeBased().everyDays(1).atHour(1).inTimezone(TZ).create();
  dailyMarkBA();
  SpreadsheetApp.getUi().alert('Aktif', 'Setiap hari sekitar pukul 01.00 WIB sistem menandai BA pada sel hari kerja yang kosong (guru lupa absen). Penandaan BA juga berjalan setiap kali aplikasi dibuka.', SpreadsheetApp.getUi().ButtonSet.OK);
}

function refreshMonthByIso_(set, iso) {
  const p = parseIso_(iso);
  if (!p || monthYear_(set, p.m) !== p.y) return;
  const sh = getMonthSheet_(p.m);
  if (sh) refreshMonthMarks_(sh, p.m, set, true);
}


/* ------------------------------------------------------------ */
/*  API admin: hari libur & rapikan data                         */
/* ------------------------------------------------------------ */
function apiHolidays(token) {
  requireAdmin_(token);
  const set = getSettings_(), hol = readHolidays_();
  const a = set.TahunAjaranMulai, from = a + '-08-01', to = (a + 1) + '-07-31';
  const items = Object.keys(hol).filter(function (k) { return k >= from && k <= to; }).sort().map(function (k) {
    const h = hol[k], effective = h.active && !(h.type === 'Cuti Bersama' && set.CutiBersama === 'Masuk');
    return { date: k, name: h.name, type: h.type, active: h.active, effective: effective };
  });
  return { ay: a + '/' + (a + 1), items: items, cutiBersama: set.CutiBersama, hariKerja: set.HariKerja };
}

function apiAddHoliday(token, p) {
  requireAdmin_(token);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const set = getSettings_();
    const s = parseIso_(p && p.start), e = (p && p.end) ? parseIso_(p.end) : s;
    if (!s || !e) throw new Error('Tanggal tidak valid.');
    const name = cleanNote_(p.name);
    if (name.length < 3) throw new Error('Keterangan libur wajib diisi (minimal 3 huruf).');
    const n = Math.round((Date.UTC(e.y, e.m, e.d) - Date.UTC(s.y, s.m, s.d)) / 86400000) + 1;
    if (n < 1) throw new Error('Tanggal selesai tidak boleh sebelum tanggal mulai.');
    if (n > 120) throw new Error('Rentang libur maksimal 120 hari.');

    HOL_CACHE_ = null;
    const hol = readHolidays_(), sh = getHolSheet_();
    const adds = [], touched = {};
    let updated = 0, skipped = 0, weekend = 0;
    for (let i = 0; i < n; i++) {
      const dt = new Date(s.y, s.m, s.d + i), iso = iso_(dt.getFullYear(), dt.getMonth(), dt.getDate());
      const dw = dt.getDay();
      if (dw === 0 || (dw === 6 && String(set.HariKerja) !== '6')) { weekend++; continue; }
      const h = hol[iso];
      if (h) {
        const eff = h.active && !(h.type === 'Cuti Bersama' && set.CutiBersama === 'Masuk');
        if (eff && h.type !== 'Khusus') { skipped++; continue; }          // sudah libur nasional
        sh.getRange(h.row, 2, 1, 3).setValues([[name, 'Khusus', 'Ya']]);
        updated++;
      } else {
        adds.push([iso, name, 'Khusus', 'Ya']);
      }
      touched[iso] = true;
    }
    if (adds.length) {
      const start = Math.max(sh.getLastRow(), 1) + 1, need = start + adds.length - 1;
      if (need > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), need - sh.getMaxRows() + 20);
      sh.getRange(start, 1, adds.length, 1).setNumberFormat('@');
      sh.getRange(start, 1, adds.length, 4).setValues(adds);
    }
    if (adds.length) sh.getRange(2, 1, sh.getLastRow() - 1, 4).sort(1);
    HOL_CACHE_ = null;
    const months = {};
    Object.keys(touched).forEach(function (iso) { months[iso.slice(0, 7)] = iso; });
    Object.keys(months).forEach(function (k) { refreshMonthByIso_(set, months[k]); });
    return { ok: true, added: adds.length, updated: updated, skippedNational: skipped, skippedWeekend: weekend };
  } finally {
    lock.releaseLock();
  }
}

function apiToggleHoliday(token, date, active) {
  requireAdmin_(token);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const set = getSettings_(), h = readHolidays_()[String(date || '')];
    if (!h) throw new Error('Tanggal libur tidak ditemukan.');
    if (h.type === 'Khusus') throw new Error('Libur khusus dihapus, bukan dinonaktifkan.');
    getHolSheet_().getRange(h.row, 4).setValue(active ? 'Ya' : 'Tidak');
    HOL_CACHE_ = null;
    refreshMonthByIso_(set, date);
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

function apiDeleteHoliday(token, date) {
  requireAdmin_(token);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const set = getSettings_(), h = readHolidays_()[String(date || '')];
    if (!h) throw new Error('Tanggal libur tidak ditemukan.');
    if (h.type !== 'Khusus') throw new Error('Libur nasional tidak dihapus. Gunakan tombol "Tetap masuk" jika sekolah masuk pada tanggal itu.');
    getHolSheet_().deleteRow(h.row);
    HOL_CACHE_ = null;
    refreshMonthByIso_(set, date);
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

function apiSyncNational(token) {
  requireAdmin_(token);
  const set = getSettings_(), n = syncNationalHolidays_();
  existingMonths_(set).forEach(function (mo) { refreshMonthMarks_(getMonthSheet_(mo.idx), mo.idx, set, true); });
  return { ok: true, added: n };
}

function apiCleanData(token) {
  requireAdmin_(token);
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const res = cleanAllData_();
    return { ok: true, backup: res.backup, summary: cleanSummary_(res) };
  } finally {
    lock.releaseLock();
  }
}


