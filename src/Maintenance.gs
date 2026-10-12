/* ------------------------------------------------------------ */
/*  Rapikan data                                                 */
/* ------------------------------------------------------------ */
function backupSpreadsheet_() {
  const name = 'Backup ' + ss_().getName() + ' ' + Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH.mm');
  return DriveApp.getFileById(ss_().getId()).makeCopy(name).getUrl();
}

function logKeys_() {
  const sh = ss_().getSheetByName(SHEET_LOG), keys = {};
  if (!sh || sh.getLastRow() < 2) return keys;
  sh.getRange(2, 1, sh.getLastRow() - 1, 9).getDisplayValues().forEach(function (r) {
    if (VALID.indexOf(String(r[5]).trim().toUpperCase()) >= 0) keys[String(r[1]).trim() + '|' + normName_(r[4])] = true;
  });
  return keys;
}

function beautifySheet_(sh) {
  const last = Math.max(sh.getLastRow(), FIRST_ROW), n = last - FIRST_ROW + 1;
  sh.getRange(2, 1, last - 1, REKAP_COL).setFontSize(10).setVerticalAlignment('middle')
    .setBorder(true, true, true, true, true, true, '#cfd8dc', SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange(2, 1, 1, REKAP_COL).setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange(FIRST_ROW, 1, n, 1).setHorizontalAlignment('center');
  sh.getRange(FIRST_ROW, 2, n, 1).setHorizontalAlignment('left');
  sh.getRange(FIRST_ROW, DAY_COL, n, 31).setHorizontalAlignment('center').setFontWeight('bold');
  sh.getRange(FIRST_ROW, REKAP_COL, n, 1).setHorizontalAlignment('left').setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  sh.setColumnWidth(1, 40); sh.setColumnWidth(2, 260);
  sh.setColumnWidths(DAY_COL, 31, 34); sh.setColumnWidth(REKAP_COL, 520);
  sh.setFrozenRows(2); sh.setFrozenColumns(2);
}

/**
 * Merapikan satu sheet bulan:
 *  - nama: spasi ganda dirapikan, NO diurutkan ulang
 *  - kode: huruf kecil / "O" -> H
 *  - kolom hari libur: sisa tulisan manual (mis. "LIBUR" ditulis vertikal) diganti "L",
 *    kecuali ada catatan absen asli di LogAbsen
 *  - kolom akhir pekan: sisa tulisan dibersihkan, absen valid dipertahankan
 *  - kode lain yang tidak dikenal dibiarkan (ditandai kuning oleh format otomatis) dan dilaporkan
 */
function cleanMonthSheet_(sh, mIdx, set, logKeys, hol) {
  const year = monthYear_(set, mIdx), days = daysIn_(year, mIdx);
  const st = { sheet: MONTHS[mIdx], namesFixed: 0, converted: 0, holidayFilled: 0, junkCleared: 0, unknown: [] };
  const last = sh.getLastRow();
  if (last < FIRST_ROW) return st;
  const rows = last - FIRST_ROW + 1;

  const nameRng = sh.getRange(FIRST_ROW, 1, rows, 2), nv = nameRng.getValues();
  let no = 0;
  nv.forEach(function (r) {
    const raw = String(r[1] == null ? '' : r[1]), nm = raw.replace(/\s+/g, ' ').trim();
    if (nm !== raw && raw !== '') st.namesFixed++;
    r[1] = nm; r[0] = nm ? ++no : '';
  });
  nameRng.setValues(nv);

  const rng = sh.getRange(FIRST_ROW, DAY_COL, rows, 31), vals = rng.getValues();
  const infos = [];
  for (let d = 1; d <= 31; d++) infos.push(d <= days ? dayInfo_(set, hol, year, mIdx, d) : null);

  for (let i = 0; i < rows; i++) {
    const name = nv[i][1];
    if (!name) continue;
    for (let d = 1; d <= 31; d++) {
      const c = vals[i][d - 1], raw = String(c == null ? '' : c).trim().toUpperCase(), inf = infos[d - 1], code = normalizeCode_(raw);
      if (!inf) { if (raw !== '') { vals[i][d - 1] = ''; st.junkCleared++; } continue; }
      if (inf.holiday && !inf.weekend) {
        const real = VALID.indexOf(code) >= 0 && logKeys[inf.iso + '|' + normName_(name)];
        if (real) { vals[i][d - 1] = code; }
        else if (raw !== 'L') { if (raw === '') st.holidayFilled++; else st.junkCleared++; vals[i][d - 1] = 'L'; }
      } else if (inf.off) {
        if (VALID.indexOf(code) >= 0) vals[i][d - 1] = code;
        else if (raw !== '') { vals[i][d - 1] = ''; st.junkCleared++; }
      } else if (code === 'L') {
        vals[i][d - 1] = ''; st.junkCleared++;
      } else if (code) {
        if (String(c) !== code) st.converted++;
        vals[i][d - 1] = code;
      } else if (raw !== '') {
        st.unknown.push(sh.getRange(FIRST_ROW + i, DAY_COL + d - 1).getA1Notation() + '="' + raw + '"');
      }
    }
  }
  rng.setValues(vals);
  return st;
}

function cleanAllData_() {
  const set = getSettings_();
  HOL_CACHE_ = null;
  const hol = readHolidays_(), keys = logKeys_(), out = { backup: '', sheets: [] };
  out.backup = backupSpreadsheet_();          // gagal -> berhenti, data tidak disentuh
  existingMonths_(set).forEach(function (mo) {
    const sh = getMonthSheet_(mo.idx);
    if (!sh.getRange(1, 1).getDisplayValue()) sh.getRange(1, 1).setValue(mo.year);
    if (!hasNames_(sh)) seedNames_(sh, mo.idx);
    syncMonthNames_(sh);
    const s = cleanMonthSheet_(sh, mo.idx, set, keys, hol);
    applyFormat_(sh);
    beautifySheet_(sh);
    refreshMonthMarks_(sh, mo.idx, set, true);
    out.sheets.push(s);
  });
  return out;
}

function cleanSummary_(res) {
  const lines = ['Cadangan (backup): ' + res.backup, ''];
  res.sheets.forEach(function (s) {
    let t = s.sheet + ': ' + s.holidayFilled + ' sel libur diisi L, ' + s.junkCleared + ' sisa tulisan dibersihkan, ' +
      s.converted + ' kode dirapikan, ' + s.namesFixed + ' nama dirapikan';
    if (s.unknown.length) t += '\n   ⚠ ' + s.unknown.length + ' sel berkode tak dikenal (ditandai kuning): ' + s.unknown.slice(0, 8).join(', ') + (s.unknown.length > 8 ? ', …' : '');
    lines.push(t);
  });
  return lines.join('\n');
}

function cleanDataMenu() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.alert('Rapikan data',
    'Sistem akan MEMBUAT SALINAN CADANGAN di Google Drive terlebih dahulu, lalu merapikan semua sheet bulan:\n' +
    '• tanggal libur nasional/khusus diisi "L" (tulisan LIBUR manual dibersihkan)\n' +
    '• kode O → H, nama & nomor dirapikan, tampilan tabel diseragamkan\n\nLanjutkan?', ui.ButtonSet.YES_NO);
  if (r !== ui.Button.YES) return;
  ui.alert('Selesai', cleanSummary_(cleanAllData_()), ui.ButtonSet.OK);
}

function syncNationalMenu() {
  const n = syncNationalHolidays_();
  existingMonths_(getSettings_()).forEach(function (mo) { refreshMonthMarks_(getMonthSheet_(mo.idx), mo.idx, getSettings_(), true); });
  SpreadsheetApp.getUi().alert(n ? n + ' tanggal libur nasional/cuti bersama ditambahkan ke sheet HariLibur.' : 'Sheet HariLibur sudah lengkap.');
}

