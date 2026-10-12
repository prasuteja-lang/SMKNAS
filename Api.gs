/* ------------------------------------------------------------ */
/*  REST API (JSON)  —  akses database dari aplikasi luar         */
/*                                                                */
/*  Web app yang sama melayani 2 hal:                             */
/*    • GET  tanpa ?action  -> halaman web (Index.html)           */
/*    • POST {action,...}   -> JSON API (file ini)                */
/*  Lihat API.md untuk contoh pemakaian.                          */
/* ------------------------------------------------------------ */

const API_VERSION = '1.0.0';
const DB_MAX_ROWS = 5000;

/** Daftar aksi API. Nama aksi -> fungsi yang sudah ada (login/akses dicek di dalam fungsinya). */
function apiActions_() {
  // [fungsi, apakahMemakaiToken]
  return {
    // ---- publik (tanpa token) ----
    'login':           [apiLogin, false],      // args: [id, password]
    'publicInit':      [apiPublicInit, false],
    'logo':            [apiLogo, false],
    // ---- sesi ----
    'resume':          [apiResume, true],
    'logout':          [apiLogout, true],
    'changePassword':  [apiChangePassword, true],   // args: [passwordLama, passwordBaru]
    // ---- guru ----
    'homeSummary':     [apiHomeSummary, true],
    'submit':          [apiSubmit, true],           // args: [{status, note, lat, lng, acc, proof...}]
    'dayStatus':       [apiDayStatus, true],
    'requestWork':     [apiRequestWork, true],
    // ---- admin ----
    'adminInit':       [apiAdminInit, true],
    'today':           [apiToday, true],
    'month':           [apiMonth, true],            // args: [indeksBulan 0-11]
    'year':            [apiYear, true],
    'adminSet':        [apiAdminSet, true],
    'addTeacher':      [apiAddTeacher, true],
    'saveSettings':    [apiSaveSettings, true],
    'accounts':        [apiAccounts, true],
    'syncAccounts':    [apiSyncAccounts, true],
    'saveAccount':     [apiSaveAccount, true],
    'resetPassword':   [apiResetPassword, true],
    'holidays':        [apiHolidays, true],
    'addHoliday':      [apiAddHoliday, true],
    'toggleHoliday':   [apiToggleHoliday, true],
    'deleteHoliday':   [apiDeleteHoliday, true],
    'syncNational':    [apiSyncNational, true],
    'cleanData':       [apiCleanData, true],
    'setupMonths':     [apiSetupMonths, true],
    'workRequests':    [apiWorkRequests, true],
    'decideWork':      [apiDecideWork, true],
    'kop':             [apiKop, true],
    // ---- database (lihat bagian bawah) ----
    'my.month':        [apiMyMonth, true],          // guru: absensi diri sendiri; args: [{month}]
    'db.sheets':       [apiDbSheets, true],         // admin / apiKey
    'db.read':         [apiDbRead, true]            // admin / apiKey; args: [{sheet, limit, offset, where, q}]
  };
}

/** Aksi yang boleh lewat GET (tanpa data sensitif). */
const API_GET_OK = { ping: 1, publicInit: 1, logo: 1 };

/* ---------------- pintu masuk ---------------- */
function apiGateway_(req, method) {
  let out;
  try {
    const action = String(req.action || '').trim();
    if (!action) throw new Error('Parameter "action" wajib diisi.');
    if (action === 'ping') {
      out = { ok: true, data: { app: 'Absensi Guru', version: API_VERSION, serverTime: Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd'T'HH:mm:ssXXX") } };
    } else {
      if (method === 'GET' && !API_GET_OK[action]) throw new Error('Aksi "' + action + '" hanya boleh lewat POST.');
      const def = apiActions_()[action];
      if (!def) throw new Error('Aksi tidak dikenal: ' + action);
      let args = Array.isArray(req.args) ? req.args.slice() : [];
      if (def[1]) {
        // apiKey (server ke server) hanya berlaku untuk db.*; selain itu wajib token sesi.
        const key = /^db\./.test(action) && apiKeyOk_(req.apiKey) ? { viaApiKey: true } : String(req.token || '');
        args = [key].concat(args);
      }
      out = { ok: true, data: def[0].apply(null, args) };
    }
  } catch (err) {
    out = { ok: false, error: String((err && err.message) || err).replace(/^(Error|Exception):\s*/, '') };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function apiPost_(e) {
  let req = {};
  try {
    const raw = e && e.postData && e.postData.contents;
    req = raw ? JSON.parse(raw) : (e && e.parameter) || {};
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: 'Body harus berupa JSON yang valid.' }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  return apiGateway_(req, 'POST');
}

/* ---------------- API key (opsional, untuk akses baca database) ---------------- */
function apiKeyOk_(key) {
  const saved = PropertiesService.getScriptProperties().getProperty('API_KEY');
  return !!(saved && key && String(key) === saved);
}

/** Menu: buat / ganti API key. Key disimpan di Script Properties (tidak ada di GitHub). */
function createApiKeyMenu() {
  const ui = SpreadsheetApp.getUi();
  const key = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '').slice(0, 16);
  PropertiesService.getScriptProperties().setProperty('API_KEY', key);
  ui.alert('API Key baru dibuat (hanya untuk BACA database).\n\n' + key + '\n\nSimpan sekarang; key lama otomatis tidak berlaku. Jangan dibagikan atau di-commit ke GitHub.');
}
function revokeApiKeyMenu() {
  PropertiesService.getScriptProperties().deleteProperty('API_KEY');
  SpreadsheetApp.getUi().alert('API Key dihapus. Akses database hanya bisa lewat login Admin.');
}

/** Admin (token sesi) atau API key. */
function requireDbAccess_(tokenOrKey) {
  // objek hanya dibuat oleh apiGateway_ setelah API key terverifikasi (string dari klien tidak bisa menyamar)
  if (tokenOrKey && typeof tokenOrKey === 'object' && tokenOrKey.viaApiKey === true) return { id: 'apikey', job: 'ApiKey' };
  return requireAdmin_(tokenOrKey);
}

/* ---------------- database: baca sheet ---------------- */
/** Sheet yang boleh dibaca + baris header-nya. Sheet Akun: kolom Password disembunyikan. */
function dbSheetDef_(name) {
  if (MONTHS.indexOf(name) >= 0) return { headerRow: 2, kind: 'month' };
  const map = {};
  map[SHEET_GURU] = { headerRow: 1 };
  map[SHEET_LOG]  = { headerRow: 1 };
  map[SHEET_HOL]  = { headerRow: 1 };
  map[SHEET_REQ]  = { headerRow: 1 };
  map[SHEET_SET]  = { headerRow: 1, kind: 'settings' };
  map[SHEET_ACC]  = { headerRow: 1, kind: 'accounts' };
  return map[name] || null;
}

function apiDbSheets(tokenOrKey) {
  requireDbAccess_(tokenOrKey);
  const names = MONTHS.concat([SHEET_GURU, SHEET_LOG, SHEET_HOL, SHEET_REQ, SHEET_SET, SHEET_ACC]);
  const out = [];
  names.forEach(function (n) {
    const sh = ss_().getSheetByName(n);
    if (sh) out.push({ sheet: n, rows: sh.getLastRow(), cols: sh.getLastColumn(), headerRow: dbSheetDef_(n).headerRow });
  });
  return out;
}

/**
 * p = { sheet, limit (default 500, maks 5000), offset, where: {Kolom: nilai}, q: 'teks cari' }
 * Hasil: { sheet, header, total, offset, rows: [ {Kolom: nilai, ...} ] }  (nilai = teks seperti tampil di sheet)
 */
function apiDbRead(tokenOrKey, p) {
  requireDbAccess_(tokenOrKey);
  p = p || {};
  const name = String(p.sheet || '');
  const def = dbSheetDef_(name);
  if (!def) throw new Error('Sheet tidak diizinkan / tidak dikenal: "' + name + '".');
  const sh = ss_().getSheetByName(name);
  if (!sh) throw new Error('Sheet "' + name + '" belum ada.');
  const hr = def.headerRow, last = sh.getLastRow(), lc = sh.getLastColumn();
  if (last <= hr || lc < 1) return { sheet: name, header: [], total: 0, offset: 0, rows: [] };

  const vals = sh.getRange(hr, 1, last - hr + 1, lc).getDisplayValues();
  const header = vals[0].map(function (h, i) { return String(h).trim() || ('kolom' + (i + 1)); });
  let body = vals.slice(1).filter(function (r) { return r.some(function (c) { return String(c).trim() !== ''; }); });

  // sembunyikan data rahasia
  let hide = -1;
  if (def.kind === 'accounts') hide = header.indexOf('Password');
  if (def.kind === 'settings') body = body.map(function (r) {
    return /pin|password|kunci|key/i.test(String(r[0])) ? [r[0], '***'].concat(r.slice(2)) : r;
  });

  // filter
  const where = p.where && typeof p.where === 'object' ? p.where : null;
  if (where) Object.keys(where).forEach(function (k) {
    const ci = header.indexOf(k);
    if (ci < 0) throw new Error('Kolom tidak ada: "' + k + '".');
    body = body.filter(function (r) { return String(r[ci]).trim() === String(where[k]).trim(); });
  });
  const q = String(p.q || '').toLowerCase().trim();
  if (q) body = body.filter(function (r) { return r.join(' ').toLowerCase().indexOf(q) >= 0; });

  const total = body.length;
  const off = Math.max(0, parseInt(p.offset, 10) || 0);
  const lim = Math.min(DB_MAX_ROWS, Math.max(1, parseInt(p.limit, 10) || 500));
  const rows = body.slice(off, off + lim).map(function (r) {
    const o = {};
    header.forEach(function (h, i) { if (i !== hide) o[h] = r[i]; });
    return o;
  });
  return { sheet: name, header: header.filter(function (h, i) { return i !== hide; }), total: total, offset: off, rows: rows };
}

/* ---------------- guru: absensi diri sendiri ---------------- */
/** p = { month: 0-11 }  -> { name, month, year, days: {"1":"H",...}, rekap } */
function apiMyMonth(token, p) {
  const s = requireSession_(token);
  p = p || {};
  const set = getSettings_();
  const m = Math.min(11, Math.max(0, parseInt(p.month, 10)));
  if (isNaN(m)) throw new Error('Parameter "month" (0-11) wajib diisi.');
  if (!s.name) throw new Error('Akun ini tidak terhubung ke data guru.');
  const sh = getMonthSheet_(m);
  if (!sh) throw new Error('Sheet bulan ' + MONTHS[m] + ' belum dibuat.');
  const row = findTeacherRow_(sh, s.name);
  if (!row) throw new Error('Nama Anda tidak ditemukan di sheet ' + MONTHS[m] + '.');
  const vals = sh.getRange(row, 1, 1, REKAP_COL).getDisplayValues()[0];
  const nd = daysIn_(monthYear_(set, m), m), days = {};
  for (let d = 1; d <= nd; d++) days[d] = vals[DAY_COL - 1 + d - 1];
  return { name: s.name, month: MONTHS[m], year: monthYear_(set, m), days: days, rekap: vals[REKAP_COL - 1] };
}
