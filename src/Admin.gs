/* ------------------------------------------------------------ */
/*  API admin (butuh PIN)                                        */
/* ------------------------------------------------------------ */
function apiAdminInit(token) {
  requireAdmin_(token);
  const set = getSettings_(), n = nowInfo_();
  return {
    settings: publicSettings_(set),
    months: existingMonths_(set),
    teachers: teacherNames_(set),
    now: n,
    validYear: monthYear_(set, n.m) === n.y
  };
}

function readMonth_(sh, mIdx, set) {
  const year = monthYear_(set, mIdx), days = daysIn_(year, mIdx);
  const res = { idx: mIdx, name: MONTHS[mIdx], year: year, days: days, teachers: [], totals: zero_(), avgPct: null, dayType: [], holidays: [] };
  const hol = readHolidays_();
  for (let d = 1; d <= days; d++) {
    const inf = dayInfo_(set, hol, year, mIdx, d);
    res.dayType.push(inf.holiday ? 'H' : inf.weekend ? 'W' : '');
    if (inf.holiday) res.holidays.push({ day: d, name: inf.holiday.name, type: inf.holiday.type });
  }
  markBelumAbsen_(sh, mIdx, set);                 // guru yang lupa absen pada hari yang sudah lewat -> BA
  const last = sh.getLastRow();
  if (last < FIRST_ROW) return res;
  const vals = sh.getRange(FIRST_ROW, 1, last - FIRST_ROW + 1, REKAP_COL).getDisplayValues();
  const pcts = [];
  vals.forEach(function (r, i) {
    const name = String(r[1]).trim();
    if (!name) return;
    const c = zero_(), cells = [];
    for (let d = 1; d <= days; d++) {
      const code = normalizeCode_(r[DAY_COL - 1 + d - 1]);
      cells.push(code);
      if (code && code !== 'L') c[code]++;
    }
    const base = c.H + c.T + c.I + c.S + c.A;      // BA belum dihitung sampai admin mengoreksinya
    const pct = base ? Math.round((c.H + c.T) / base * 1000) / 10 : null;
    if (pct !== null) pcts.push(pct);
    CNT.forEach(function (k) { res.totals[k] += c[k]; });
    res.teachers.push({ row: FIRST_ROW + i, no: r[0], name: name, cells: cells, H: c.H, I: c.I, S: c.S, T: c.T, TJ: c.TJ, A: c.A, BA: c.BA, pct: pct, note: String(r[REKAP_COL - 1] || '').trim() });
  });
  if (pcts.length) res.avgPct = Math.round(pcts.reduce(function (a, b) { return a + b; }, 0) / pcts.length * 10) / 10;
  return res;
}

function todayLogs_(dateStr) {
  const sh = ss_().getSheetByName(SHEET_LOG);
  const out = {};
  if (!sh || sh.getLastRow() < 2) return out;
  const start = Math.max(2, sh.getLastRow() - 800);
  const cols = Math.min(15, sh.getMaxColumns());
  sh.getRange(start, 1, sh.getLastRow() - start + 1, cols).getDisplayValues().forEach(function (r) {
    if (r[1] === dateStr && String(r[8]).indexOf('Ditolak') !== 0) out[normName_(r[4])] = { time: String(r[2]).slice(0, 5), note: r[6], minLate: r[7], source: r[8], proof: String(r[14] || '') };
  });
  return out;
}

function apiToday(token) {
  requireAdmin_(token);
  const set = getSettings_(), n = nowInfo_();
  if (monthYear_(set, n.m) !== n.y) throw new Error('Tahun ajaran di Pengaturan belum sesuai dengan tanggal hari ini.');
  const sh = ensureMonthSheet_(n.m, set);
  const md = readMonth_(sh, n.m, set);
  const logs = todayLogs_(n.date);
  const rows = md.teachers.map(function (t) {
    const lg = logs[normName_(t.name)] || {};
    return { name: t.name, status: t.cells[n.day - 1] || '', time: lg.time || '', note: lg.note || '', minLate: lg.minLate || 0, proof: lg.proof || '' };
  });
  const di = dayInfo_(set, readHolidays_(), n.y, n.m, n.day);
  const cnt = { H: 0, I: 0, S: 0, T: 0, TJ: 0, A: 0, BA: 0, belum: 0 };
  rows.forEach(function (r) {
    if (r.status === 'L') return;
    if (r.status) cnt[r.status]++; else if (!di.off) cnt.belum++;
  });
  return { date: n.date, day: n.day, rows: rows, counts: cnt, jamMasuk: set.JamMasuk, libur: di.off ? (di.name || 'Libur') : '', pendingReq: countPendingReq_(n.date) };
}

function apiMonth(token, mIdx) {
  requireAdmin_(token);
  const set = getSettings_();
  mIdx = parseInt(mIdx, 10);
  const sh = getMonthSheet_(mIdx);
  if (!sh) throw new Error('Sheet bulan ' + MONTHS[mIdx] + ' belum ada.');
  const md = readMonth_(sh, mIdx, set), px = monthProofs_(md.year, mIdx);
  md.teachers.forEach(function (t) { t.proofs = px[normName_(t.name)] || []; });
  return md;
}

/** Tautan bukti Izin / Sakit pada satu bulan: { namaTernormalisasi: [{ day, st, url }] }. */
function monthProofs_(year, mIdx) {
  const out = {}, sh = ss_().getSheetByName(SHEET_LOG);
  if (!sh || sh.getLastRow() < 2 || sh.getMaxColumns() < 15) return out;
  const pre = year + '-' + pad2_(mIdx + 1) + '-';
  sh.getRange(2, 1, sh.getLastRow() - 1, 15).getDisplayValues().forEach(function (r) {
    const d = String(r[1]).trim(), url = String(r[14] || '').trim();
    if (!url || d.indexOf(pre) !== 0) return;
    const k = normName_(r[4]);
    (out[k] = out[k] || []).push({ day: +d.slice(8, 10), st: String(r[5]).trim(), url: url });
  });
  return out;
}

function apiYear(token) {
  requireAdmin_(token);
  const set = getSettings_();
  const months = [], datas = [];
  AY_ORDER.forEach(function (i) {
    const sh = getMonthSheet_(i);
    if (!sh) return;
    const md = readMonth_(sh, i, set);
    const has = md.teachers.some(function (t) { return t.H + t.I + t.S + t.T + t.TJ + t.A > 0; });
    if (has) { months.push({ idx: i, name: MONTHS[i], year: md.year }); datas.push(md); }
  });
  const order = [], map = {};
  datas.forEach(function (md, mi) {
    md.teachers.forEach(function (t) {
      const k = normName_(t.name);
      if (!map[k]) { const z = zero_(); z.name = t.name; z.perMonth = months.map(function () { return null; }); map[k] = z; order.push(k); }
      const o = map[k];
      o.perMonth[mi] = { H: t.H, I: t.I, S: t.S, T: t.T, TJ: t.TJ, A: t.A, BA: t.BA, pct: t.pct, note: t.note };
      CNT.forEach(function (c) { o[c] += t[c]; });
    });
  });
  const teachers = order.map(function (k) {
    const o = map[k], base = o.H + o.T + o.I + o.S + o.A;
    o.pct = base ? Math.round((o.H + o.T) / base * 1000) / 10 : null;
    return o;
  });
  const tot = zero_();
  teachers.forEach(function (t) { CNT.forEach(function (c) { tot[c] += t[c]; }); });
  const ps = teachers.filter(function (t) { return t.pct !== null; }).map(function (t) { return t.pct; });
  return {
    ay: set.TahunAjaranMulai + '/' + (set.TahunAjaranMulai + 1), months: months, teachers: teachers, totals: tot,
    avgPct: ps.length ? Math.round(ps.reduce(function (a, b) { return a + b; }, 0) / ps.length * 10) / 10 : null
  };
}

function apiAdminSet(token, p) {
  const adm = requireAdmin_(token);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const set = getSettings_();
    const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(p.date || ''));
    if (!m) throw new Error('Tanggal tidak valid.');
    const y = +m[1], mIdx = +m[2] - 1, day = +m[3];
    if (monthYear_(set, mIdx) !== y) throw new Error('Tanggal di luar tahun ajaran ' + set.TahunAjaranMulai + '/' + (set.TahunAjaranMulai + 1) + '.');
    if (day < 1 || day > daysIn_(y, mIdx)) throw new Error('Tanggal tidak valid.');
    let st = String(p.status || '').toUpperCase();
    if (['', 'H', 'I', 'S', 'T', 'TJ', 'A'].indexOf(st) < 0) throw new Error('Keterangan tidak valid.');
    let note = cleanNote_(p.note), minLate = 0, timeStr = nowInfo_().time;
    const tm = String(p.time || '').trim();
    if (tm) {
      // Admin mengisi jam masuk -> sistem menentukan Hadir atau Telat dari batas jam masuk
      const m2 = /^(\d{1,2})[:.](\d{2})$/.exec(tm);
      if (!m2 || +m2[1] > 23 || +m2[2] > 59) throw new Error('Format jam masuk harus HH:mm, contoh 07:25.');
      if (st !== 'H' && st !== 'T') throw new Error('Jam masuk hanya dipakai untuk status Hadir atau Telat.');
      const sec = (+m2[1]) * 3600 + (+m2[2]) * 60, lim = timeToSec_(set.JamMasuk), hhmm = pad2_(+m2[1]) + ':' + m2[2];
      timeStr = hhmm + ':00';
      if (sec > lim) {
        st = 'T'; minLate = Math.ceil((sec - lim) / 60);
        if (note.length < 3) note = 'Jam masuk ' + hhmm + ' (dikoreksi admin)';
      } else st = 'H';
    } else if (st === 'T') {
      minLate = Math.max(0, parseInt(p.minLate, 10) || 0);
    }
    if ((st === 'I' || st === 'S' || st === 'T') && note.length < 3) throw new Error('Catatan alasan wajib diisi untuk ' + STATUS[st] + '.');
    const sh = ensureMonthSheet_(mIdx, set);
    const row = findTeacherRow_(sh, p.name);
    if (!row) throw new Error('Nama guru tidak ditemukan di sheet ' + MONTHS[mIdx] + '.');
    writeEntry_(sh, row, sh.getRange(row, 1).getDisplayValue(), String(p.name).trim(), day, p.date, timeStr, st, note, minLate, 'Admin', { user: adm.id });
    return { ok: true, status: st, minLate: minLate, time: timeStr.slice(0, 5) };
  } finally {
    lock.releaseLock();
  }
}

function apiAddTeacher(token, name) {
  requireAdmin_(token);
  name = String(name || '').replace(/\s+/g, ' ').trim();
  if (name.length < 3) throw new Error('Nama guru terlalu pendek.');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const set = getSettings_(), sh = getGuruSheet_(), key = normName_(name);
    const found = readMaster_().filter(function (t) { return normName_(t.name) === key; })[0];
    if (found) sh.getRange(found.row, 3).setValue('Aktif');
    else {
      const last = Math.max(sh.getLastRow(), 1), nums = last >= 2 ? sh.getRange(2, 1, last - 1, 1).getValues() : [];
      const next = Math.max.apply(null, nums.map(function (r) { return Number(r[0]) || 0; }).concat([0])) + 1;
      sh.getRange(last + 1, 1, 1, 5).setValues([[next, name, 'Aktif', 'Guru', '']]);   // id_login diisi otomatis oleh syncAccounts_
    }
    let added = 0;
    syncAccounts_();
    existingMonths_(set).forEach(function (mo) {
      const msh = getMonthSheet_(mo.idx);
      added += syncMonthNames_(msh) ? 1 : 0;
      refreshMonthMarks_(msh, mo.idx, set, false);
    });
    return { ok: true, added: added };
  } finally {
    lock.releaseLock();
  }
}

function apiSaveSettings(token, o) {
  requireAdmin_(token);
  const sh = ss_().getSheetByName(SHEET_SET);
  const before = getSettings_();
  const allowed = DEFAULTS.map(function (d) { return d[0]; });
  const vals = sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 2).getDisplayValues();
  Object.keys(o || {}).forEach(function (k) {
    if (allowed.indexOf(k) < 0) return;
    let v = String(o[k] == null ? '' : o[k]).trim();
    if (k === 'PasswordAwal' && v.length < 6) throw new Error('Password awal minimal 6 karakter.');
    if (k === 'LokasiWajib' && v !== 'Ya' && v !== 'Tidak') throw new Error('Lokasi wajib harus Ya atau Tidak.');
    if (k === 'SekolahLat' && !(isFinite(parseFloat(v)) && Math.abs(parseFloat(v)) <= 90)) throw new Error('Latitude sekolah tidak valid (contoh -6.890437).');
    if (k === 'SekolahLng' && !(isFinite(parseFloat(v)) && Math.abs(parseFloat(v)) <= 180)) throw new Error('Longitude sekolah tidak valid (contoh 107.626437).');
    if (k === 'RadiusMeter' && !(/^\d+$/.test(v) && +v >= 10 && +v <= 5000)) throw new Error('Radius harus 10 - 5000 meter.');
    if (k === 'MaxAkurasi' && !(/^\d+$/.test(v) && +v >= 10 && +v <= 1000)) throw new Error('Batas akurasi harus 10 - 1000 meter.');
    if (k === 'JamMasuk' && !/^\d{1,2}[:.]\d{2}$/.test(v)) throw new Error('Format jam masuk harus HH:mm, contoh 07:00.');
    if (k === 'TahunAjaranMulai' && !/^\d{4}$/.test(v)) throw new Error('Tahun ajaran harus 4 digit, contoh 2026.');
    if (k === 'HariKerja' && v !== '5' && v !== '6') throw new Error('Hari kerja harus 5 atau 6.');
    if (k === 'CutiBersama' && v !== 'Libur' && v !== 'Masuk') throw new Error('Cuti bersama harus Libur atau Masuk.');
    for (let i = 0; i < vals.length; i++) {
      if (String(vals[i][0]).trim() === k) { const c = sh.getRange(2 + i, 2); c.setNumberFormat('@'); c.setValue(v); break; }
    }
  });
  const after = getSettings_();
  if (['HariKerja', 'CutiBersama', 'TahunAjaranMulai'].some(function (k) { return String(before[k]) !== String(after[k]); })) {
    HOL_CACHE_ = null;
    existingMonths_(after).forEach(function (mo) { refreshMonthMarks_(getMonthSheet_(mo.idx), mo.idx, after, true); });
  }
  return { ok: true, settings: publicSettings_(after) };
}

