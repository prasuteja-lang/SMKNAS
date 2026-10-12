/* ------------------------------------------------------------ */
/*  Sheet bulanan                                                */
/* ------------------------------------------------------------ */
function getMonthSheet_(mIdx) { return ss_().getSheetByName(MONTHS[mIdx]); }

function existingMonths_(set) {
  return AY_ORDER.filter(function (i) { return !!getMonthSheet_(i); })
    .map(function (i) { return { idx: i, name: MONTHS[i], year: monthYear_(set, i) }; });
}

function hasNames_(sh) {
  if (!sh || sh.getLastRow() < FIRST_ROW) return false;
  return sh.getRange(FIRST_ROW, 2, sh.getLastRow() - FIRST_ROW + 1, 1).getDisplayValues()
    .some(function (r) { return String(r[0]).trim(); });
}

function findSourceSheet_(mIdx) {
  const pos = AY_ORDER.indexOf(mIdx);
  for (let p = pos - 1; p >= 0; p--) { const s = getMonthSheet_(AY_ORDER[p]); if (hasNames_(s)) return s; }
  for (let p = pos + 1; p < AY_ORDER.length; p++) { const s = getMonthSheet_(AY_ORDER[p]); if (hasNames_(s)) return s; }
  return null;
}

function copyNamesFrom_(sh, src) {
  if (!src || src.getLastRow() < FIRST_ROW) return 0;
  const names = src.getRange(FIRST_ROW, 1, src.getLastRow() - FIRST_ROW + 1, 2).getDisplayValues()
    .filter(function (r) { return String(r[1]).trim(); });
  if (names.length) sh.getRange(FIRST_ROW, 1, names.length, 2).setValues(names);
  return names.length;
}

function ensureMonthSheet_(mIdx, set) {
  let sh = getMonthSheet_(mIdx);
  const year = monthYear_(set, mIdx);
  if (sh) {
    if (!sh.getRange(1, 1).getDisplayValue()) sh.getRange(1, 1).setValue(year);
    if (!sh.getRange(2, REKAP_COL).getDisplayValue()) sh.getRange(2, REKAP_COL).setValue('REKAP');
    // sheet bulan sudah ada tetapi daftar guru kosong -> salin dari bulan terdekat yang berisi
    if (!hasNames_(sh)) seedNames_(sh, mIdx);
    refreshMonthMarks_(sh, mIdx, set, false);
    return sh;
  }
  sh = ss_().insertSheet(MONTHS[mIdx]);
  const head = ['NO', 'NAMA'];
  for (let d = 1; d <= 31; d++) head.push(d);
  head.push('REKAP');
  sh.getRange(1, 1).setValue(year).setFontSize(20);
  sh.getRange(2, 1, 1, head.length).setValues([head]).setFontWeight('bold').setHorizontalAlignment('center').setBackground('#e6f4f1');
  seedNames_(sh, mIdx);
  sh.setColumnWidth(1, 40); sh.setColumnWidth(2, 260);
  sh.setColumnWidths(DAY_COL, 31, 34); sh.setColumnWidth(REKAP_COL, 520);
  sh.getRange(FIRST_ROW, DAY_COL, Math.max(sh.getMaxRows() - 2, 1), 31).setHorizontalAlignment('center');
  applyFormat_(sh);
  refreshMonthMarks_(sh, mIdx, set, true);
  return sh;
}

function applyFormat_(sh) {
  sh.setFrozenRows(2); sh.setFrozenColumns(2);
  const rng = sh.getRange(3, DAY_COL, Math.max(sh.getMaxRows() - 2, 1), 31);
  const mk = function (txts, bg, fc) {
    return txts.map(function (t) {
      return SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(t).setBackground(bg).setFontColor(fc).setRanges([rng]).build();
    });
  };
  // isi sel di luar kode yang dikenal (mis. sisa tulisan manual) ditandai kuning supaya mudah dicek
  const cond = ['H', 'O', 'I', 'S', 'T', 'TJ', 'A', 'BA', 'L'].map(function (k) { return 'UPPER(TRIM(C3))<>"' + k + '"'; }).join(',');
  const unknown = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND(LEN(TRIM(C3))>0,' + cond + ')')
    .setBackground('#fff2cc').setFontColor('#7f6000').setRanges([rng]).build();
  const rules = [].concat(
    mk(['H', 'O'], '#d9f2e1', '#137333'), mk(['I'], '#dbe8fd', '#1a56b8'),
    mk(['S'], '#ffe9c7', '#a15c00'), mk(['T'], '#fcd9d6', '#b3261e'), mk(['TJ'], '#e8eaed', '#5f6368'),
    mk(['A'], '#c0392b', '#ffffff'), mk(['BA'], '#ead7fb', '#6b21a8'),
    mk(['L'], '#fde2e0', '#b3261e'), [unknown]);
  sh.setConditionalFormatRules(rules);
}

function findTeacherRow_(sh, name) {
  const last = sh.getLastRow();
  if (last < FIRST_ROW) return 0;
  const vals = sh.getRange(FIRST_ROW, 2, last - FIRST_ROW + 1, 1).getDisplayValues();
  const key = normName_(name);
  for (let i = 0; i < vals.length; i++) if (normName_(vals[i][0]) === key) return FIRST_ROW + i;
  return 0;
}

function teacherNames_(set) {
  const act = activeTeachers_();
  if (act.length) {
    const n0 = nowInfo_();
    if (monthYear_(set, n0.m) === n0.y) syncMonthNames_(ensureMonthSheet_(n0.m, set));
    return act;
  }
  const n = nowInfo_();
  let sh = null;
  if (monthYear_(set, n.m) === n.y) sh = ensureMonthSheet_(n.m, set);
  else { const ex = existingMonths_(set); if (ex.length) sh = getMonthSheet_(ex[0].idx); }
  if (!sh || sh.getLastRow() < FIRST_ROW) return [];
  return sh.getRange(FIRST_ROW, 2, sh.getLastRow() - FIRST_ROW + 1, 1).getDisplayValues()
    .map(function (r) { return String(r[0]).trim(); }).filter(String);
}

