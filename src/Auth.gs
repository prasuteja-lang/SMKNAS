/* ------------------------------------------------------------ */
/*  Login, sesi & akun                                           */
/* ------------------------------------------------------------ */
const ACC_HEAD = ['ID Login', 'Password', 'Nama Guru', 'Job', 'Status', 'Wajib Ganti Password', 'Login Terakhir'];
const SESSION_TTL = 21600;   // detik (6 jam, diperpanjang tiap aktivitas)

function normJob_(v) { return /^admin/i.test(String(v || '').trim()) ? 'Admin' : 'Guru'; }
function accKey_(id) { return String(id || '').trim().toLowerCase(); }

function hashPw_(pw, salt) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + ':' + pw, Utilities.Charset.UTF_8);
  return bytes.map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');
}
function makePw_(pw) {
  const salt = Utilities.getUuid().replace(/-/g, '').slice(0, 16);
  return 'h1$' + salt + '$' + hashPw_(String(pw), salt);
}
/** Password di sheet berformat "h1$salt$hash". Teks biasa juga diterima (admin mengetik password baru langsung di sheet) lalu otomatis dienkripsi saat login. */
function checkPw_(stored, pw) {
  stored = String(stored == null ? '' : stored);
  if (stored.indexOf('h1$') === 0) { const p = stored.split('$'); return p.length === 3 && p[2] === hashPw_(pw, p[1]); }
  return stored !== '' && stored === pw;
}

function getAccSheet_() {
  const ss = ss_();
  let sh = ss.getSheetByName(SHEET_ACC);
  if (!sh) {
    sh = ss.insertSheet(SHEET_ACC);
    sh.getRange(1, 1, 1, ACC_HEAD.length).setValues([ACC_HEAD]).setFontWeight('bold').setBackground('#e6f4f1');
    const rows = Math.max(sh.getMaxRows() - 1, 1);
    sh.getRange(2, 1, rows, 3).setNumberFormat('@');
    sh.getRange(2, 4, rows, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Guru', 'Admin'], true).setAllowInvalid(false).build());
    sh.getRange(2, 5, rows, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Aktif', 'Nonaktif'], true).setAllowInvalid(false).build());
    sh.getRange(2, 6, rows, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Ya', 'Tidak'], true).setAllowInvalid(false).build());
    sh.setColumnWidth(1, 130); sh.setColumnWidth(2, 300); sh.setColumnWidth(3, 280); sh.setColumnWidth(4, 80); sh.setColumnWidth(5, 90); sh.setColumnWidth(6, 150); sh.setColumnWidth(7, 150);
    sh.setFrozenRows(1);
    const set = getSettings_();
    // Akun 'admin' cadangan hanya dibuat bila sheet Nama Guru belum punya Admin ber-id_login
    const hasAdmin = readMaster_().some(function (t) { return t.job === 'Admin' && t.loginId; });
    if (!hasAdmin) sh.getRange(2, 1, 1, 7).setValues([['admin', makePw_(set.PinAdmin || set.PasswordAwal), '', 'Admin', 'Aktif', 'Ya', '']]);
  }
  return sh;
}

function readAccounts_() {
  const sh = getAccSheet_(), last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, 7).getValues().map(function (r, i) {
    return {
      row: 2 + i, id: String(r[0] == null ? '' : r[0]).trim(), pw: r[1], name: String(r[2] == null ? '' : r[2]).replace(/\s+/g, ' ').trim(),
      job: normJob_(r[3]), active: !/^(non|tidak|off|no|0)/i.test(String(r[4] == null ? '' : r[4]).trim()),
      mustChange: /^(ya|y|1|true)/i.test(String(r[5] == null ? '' : r[5]).trim()),
      last: r[6] instanceof Date ? Utilities.formatDate(r[6], TZ, 'yyyy-MM-dd HH:mm') : String(r[6] || '')
    };
  }).filter(function (a) { return a.id; });
}

/** ID baru berformat SMKNAS001, SMKNAS002, ... yang belum dipakai. */
function nextLoginId_(taken) {
  let i = 1;
  while (taken[accKey_('SMKNAS' + ('00' + i).slice(-3))]) i++;
  return 'SMKNAS' + ('00' + i).slice(-3);
}

/**
 * Sinkronkan sheet "Akun" dengan sheet "Nama Guru" (kolom Status, Job, id_login). Sheet Nama Guru adalah sumber utama:
 *  - id_login di Nama Guru dipakai sebagai ID Login; akun baru dibuat dengan Password Awal
 *  - nama / Job / status akun mengikuti Nama Guru; password akun yang sudah ada TIDAK diubah
 *  - baris tanpa id_login diberi ID otomatis (SMKNAS###) dan ditulis balik ke Nama Guru
 * Mengembalikan { added, updated }.
 */
function syncAccounts_() {
  const set = getSettings_(), sh = getAccSheet_(), accs = readAccounts_(), byId = {}, taken = {};
  accs.forEach(function (a) { byId[accKey_(a.id)] = a; taken[accKey_(a.id)] = true; });
  let master = readMaster_();
  const mSh = ss_().getSheetByName(SHEET_GURU);
  if (!master.length) master = teacherNames_(set).map(function (n) { return { row: 0, name: n, status: 'Aktif', job: 'Guru', loginId: '' }; });
  master.forEach(function (t) { if (t.loginId) taken[accKey_(t.loginId)] = true; });
  const pw = makePw_(set.PasswordAwal), add = [];
  let updated = 0;
  master.forEach(function (t) {
    if (!t.loginId) {
      t.loginId = nextLoginId_(taken); taken[accKey_(t.loginId)] = true;
      if (t.row && mSh) { mSh.getRange(t.row, 5).setNumberFormat('@').setValue(t.loginId); if (!mSh.getRange(t.row, 4).getDisplayValue()) mSh.getRange(t.row, 4).setValue(t.job); }
    }
    const active = !/^(non|tidak|keluar|pindah|resign|off)/i.test(t.status), a = byId[accKey_(t.loginId)];
    if (a && a.row) {
      if (a.name !== t.name || a.job !== t.job || a.active !== active) {
        sh.getRange(a.row, 3, 1, 3).setValues([[t.name, t.job, active ? 'Aktif' : 'Nonaktif']]);
        revoke_(a.id); updated++;
      }
    } else if (!a) {
      add.push([t.loginId, pw, t.name, t.job, active ? 'Aktif' : 'Nonaktif', 'Ya', '']);
      byId[accKey_(t.loginId)] = { id: t.loginId };
    }
  });
  if (add.length) {
    const start = Math.max(sh.getLastRow(), 1) + 1, need = start + add.length - 1;
    if (need > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), need - sh.getMaxRows() + 10);
    sh.getRange(start, 1, add.length, 3).setNumberFormat('@');
    sh.getRange(start, 1, add.length, 7).setValues(add);
  }
  return { added: add.length, updated: updated };
}

function syncAccountsMenu() {
  const r = syncAccounts_();
  SpreadsheetApp.getUi().alert('Sinkron akun dari sheet Nama Guru selesai.\n' + r.added + ' akun baru dibuat (password awal = PasswordAwal di sheet Pengaturan), ' + r.updated + ' akun diperbarui (nama / Job / status).');
}

/* ---- sesi (CacheService) ---- */
function revoke_(id) { CacheService.getScriptCache().put('rv_' + accKey_(id), String(Date.now()), SESSION_TTL); }

function newSession_(acc) {
  const token = (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '');
  const s = { id: acc.id, name: acc.name, job: acc.job, mc: !!acc.mustChange, t: Date.now() };
  CacheService.getScriptCache().put('s_' + token, JSON.stringify(s), SESSION_TTL);
  return token;
}

/** Validasi token. Selama wajib-ganti-password, hanya apiChangePassword/apiResume yang boleh (allowMC). */
function requireSession_(token, allowMC) {
  const cache = CacheService.getScriptCache();
  const raw = token ? cache.get('s_' + String(token)) : null;
  if (!raw) throw new Error('Sesi berakhir, silakan login kembali.');
  const s = JSON.parse(raw), rv = cache.get('rv_' + accKey_(s.id));
  if (rv && +rv > s.t) { cache.remove('s_' + String(token)); throw new Error('Sesi berakhir, silakan login kembali.'); }
  if (s.mc && !allowMC) throw new Error('Anda wajib mengganti password terlebih dahulu.');
  cache.put('s_' + String(token), raw, SESSION_TTL);
  return s;
}
function requireAdmin_(token) {
  const s = requireSession_(token);
  if (s.job !== 'Admin') throw new Error('Akses ditolak: menu ini khusus Admin.');
  return s;
}

function cfgOf_(set) {
  return { school: set.NamaSekolah, jamMasuk: set.JamMasuk, lokasiWajib: set.LokasiWajib,
    lat: parseFloat(set.SekolahLat), lng: parseFloat(set.SekolahLng), radius: set.RadiusMeter, maxAcc: set.MaxAkurasi };
}
function userOf_(s) { return { id: s.id, name: s.name, job: s.job, disp: s.name || s.id, mustChange: !!s.mc }; }

function apiLogin(id, pw) {
  id = String(id || '').trim().slice(0, 40); pw = String(pw || '');
  if (!id || !pw) throw new Error('ID Login dan password wajib diisi.');
  const cache = CacheService.getScriptCache(), fk = 'fail_' + accKey_(id), fails = +(cache.get(fk) || 0);
  if (fails >= 5) throw new Error('Terlalu banyak percobaan gagal. Coba lagi dalam 10 menit.');
  const acc = readAccounts_().filter(function (a) { return accKey_(a.id) === accKey_(id); })[0];
  if (!acc || !acc.active || !checkPw_(acc.pw, pw)) {
    cache.put(fk, String(fails + 1), 600);
    Utilities.sleep(700);
    throw new Error('ID Login atau password salah, atau akun nonaktif.');
  }
  cache.remove(fk);
  const sh = getAccSheet_();
  if (String(acc.pw).indexOf('h1$') !== 0) sh.getRange(acc.row, 2).setValue(makePw_(pw));   // enkripsi password yang diketik manual
  sh.getRange(acc.row, 7).setValue(Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm'));
  const set = getSettings_(), token = newSession_(acc);
  return { ok: true, token: token, user: userOf_({ id: acc.id, name: acc.name, job: acc.job, mc: acc.mustChange }), cfg: cfgOf_(set), serverMs: Date.now() };
}

/** Pulihkan sesi setelah halaman di-refresh. */
function apiResume(token) {
  const s = requireSession_(token, true), set = getSettings_();
  return { ok: true, token: token, user: userOf_(s), cfg: cfgOf_(set), serverMs: Date.now() };
}

function apiLogout(token) {
  if (token) CacheService.getScriptCache().remove('s_' + String(token));
  return { ok: true };
}

function apiChangePassword(token, oldPw, newPw) {
  const s = requireSession_(token, true);
  oldPw = String(oldPw || ''); newPw = String(newPw || '');
  if (newPw.length < 6) throw new Error('Password baru minimal 6 karakter.');
  if (newPw === oldPw) throw new Error('Password baru harus berbeda dari password lama.');
  if (accKey_(newPw) === accKey_(s.id)) throw new Error('Password tidak boleh sama dengan ID Login.');
  if (newPw === getSettings_().PasswordAwal) throw new Error('Password tidak boleh sama dengan password awal.');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const acc = readAccounts_().filter(function (a) { return accKey_(a.id) === accKey_(s.id); })[0];
    if (!acc || !checkPw_(acc.pw, oldPw)) throw new Error('Password lama salah.');
    const sh = getAccSheet_();
    sh.getRange(acc.row, 2).setValue(makePw_(newPw));
    sh.getRange(acc.row, 6).setValue('Tidak');
    const cache = CacheService.getScriptCache(), raw = cache.get('s_' + String(token));
    if (raw) { const o = JSON.parse(raw); o.mc = false; cache.put('s_' + String(token), JSON.stringify(o), SESSION_TTL); }
    return { ok: true };
  } finally { lock.releaseLock(); }
}

/* ---- kelola akun (Admin) ---- */
function apiAccounts(token) {
  requireAdmin_(token);
  const accs = readAccounts_(), set = getSettings_();
  return {
    items: accs.map(function (a) { return { id: a.id, name: a.name, job: a.job, active: a.active, mustChange: a.mustChange, last: a.last }; }),
    teachers: teacherNames_(set),
    defaultPw: set.PasswordAwal
  };
}

function apiSyncAccounts(token) {
  requireAdmin_(token);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { const r = syncAccounts_(); return { ok: true, added: r.added, updated: r.updated }; } finally { lock.releaseLock(); }
}

function apiSaveAccount(token, p) {
  const me = requireAdmin_(token);
  p = p || {};
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const id = String(p.id || '').trim(), orig = String(p.origId || '').trim(), name = String(p.name || '').replace(/\s+/g, ' ').trim();
    const job = normJob_(p.job), active = !!p.active;
    if (!/^[A-Za-z0-9._-]{3,30}$/.test(id)) throw new Error('ID Login 3-30 karakter: huruf, angka, titik, strip, atau garis bawah (tanpa spasi).');
    if (job === 'Guru' && !name) throw new Error('Akun berjob Guru harus dihubungkan ke nama guru.');
    const accs = readAccounts_(), cur = orig ? accs.filter(function (a) { return accKey_(a.id) === accKey_(orig); })[0] : null;
    if (orig && !cur) throw new Error('Akun tidak ditemukan.');
    if (accs.some(function (a) { return accKey_(a.id) === accKey_(id) && (!cur || a.row !== cur.row); })) throw new Error('ID Login "' + id + '" sudah dipakai.');
    if (cur && accKey_(cur.id) === accKey_(me.id) && (job !== 'Admin' || !active)) throw new Error('Anda tidak dapat menurunkan / menonaktifkan akun Anda sendiri.');
    const sh = getAccSheet_();
    if (cur) {
      sh.getRange(cur.row, 1).setValue(id);
      sh.getRange(cur.row, 3, 1, 3).setValues([[name, job, active ? 'Aktif' : 'Nonaktif']]);
      revoke_(cur.id);
      updateMasterFor_(cur.id, id, job, active);
    } else {
      const r = Math.max(sh.getLastRow(), 1) + 1;
      if (r > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), 10);
      sh.getRange(r, 1, 1, 3).setNumberFormat('@');
      sh.getRange(r, 1, 1, 7).setValues([[id, makePw_(getSettings_().PasswordAwal), name, job, active ? 'Aktif' : 'Nonaktif', 'Ya', '']]);
    }
    return { ok: true, created: !cur, defaultPw: getSettings_().PasswordAwal };
  } finally { lock.releaseLock(); }
}

/** Samakan perubahan akun (ID / Job / status) ke sheet Nama Guru supaya sinkron berikutnya tidak menimpanya. */
function updateMasterFor_(origId, id, job, active) {
  const mSh = ss_().getSheetByName(SHEET_GURU);
  if (!mSh) return;
  const m = readMaster_().filter(function (t) { return accKey_(t.loginId) === accKey_(origId); })[0];
  if (!m) return;
  mSh.getRange(m.row, 4).setValue(job);
  mSh.getRange(m.row, 5).setNumberFormat('@').setValue(id);
  mSh.getRange(m.row, 3).setValue(active ? 'Aktif' : 'Nonaktif');
}

function apiResetPassword(token, id) {
  requireAdmin_(token);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const acc = readAccounts_().filter(function (a) { return accKey_(a.id) === accKey_(id); })[0];
    if (!acc) throw new Error('Akun tidak ditemukan.');
    const set = getSettings_(), sh = getAccSheet_();
    sh.getRange(acc.row, 2).setValue(makePw_(set.PasswordAwal));
    sh.getRange(acc.row, 6).setValue('Ya');
    revoke_(acc.id);
    CacheService.getScriptCache().remove('fail_' + accKey_(acc.id));
    return { ok: true, password: set.PasswordAwal };
  } finally { lock.releaseLock(); }
}

