/* ------------------------------------------------------------ */
/*  Buka Absen di hari libur / Sabtu / Minggu                    */
/*  Guru mengajukan  ->  Admin menyetujui  ->  guru bisa absen    */
/* ------------------------------------------------------------ */
const REQ_HEAD = ['Tanggal', 'ID Login', 'Nama', 'Alasan', 'Status', 'Diajukan', 'Diputuskan Oleh', 'Diputuskan Pada'];

function getReqSheet_() {
  const ss = ss_();
  let sh = ss.getSheetByName(SHEET_REQ);
  if (!sh) {
    sh = ss.insertSheet(SHEET_REQ);
    sh.getRange(1, 1, 1, REQ_HEAD.length).setValues([REQ_HEAD]).setFontWeight('bold').setBackground('#e6f4f1');
    sh.getRange(2, 1, Math.max(sh.getMaxRows() - 1, 1), REQ_HEAD.length).setNumberFormat('@');
    sh.setColumnWidth(1, 100); sh.setColumnWidth(2, 120); sh.setColumnWidth(3, 260); sh.setColumnWidth(4, 360); sh.setColumnWidth(5, 100);
    sh.setColumnWidths(6, 3, 140);
    sh.setFrozenRows(1);
  }
  return sh;
}

function readRequests_() {
  const sh = getReqSheet_(), last = sh.getLastRow();
  if (last < 2) return [];
  const fmt = function (v) { return v instanceof Date ? Utilities.formatDate(v, TZ, 'yyyy-MM-dd HH:mm') : String(v == null ? '' : v); };
  return sh.getRange(2, 1, last - 1, 8).getValues().map(function (r, i) {
    const stt = String(r[4] || '').trim();
    return {
      row: 2 + i, date: cellIso_(r[0]), id: String(r[1] == null ? '' : r[1]).trim(), name: String(r[2] || '').trim(), note: String(r[3] || ''),
      status: /setuju|^acc|^ya$/i.test(stt) ? 'Disetujui' : /tolak/i.test(stt) ? 'Ditolak' : 'Menunggu',
      at: fmt(r[5]), by: String(r[6] || ''), decAt: fmt(r[7])
    };
  }).filter(function (q) { return q.date && q.id; });
}

function findReq_(list, iso, id) {
  for (let i = 0; i < list.length; i++) if (list[i].date === iso && accKey_(list[i].id) === accKey_(id)) return list[i];
  return null;
}

function countPendingReq_(iso) {
  return readRequests_().filter(function (q) { return q.date === iso && q.status === 'Menunggu'; }).length;
}

/** Status hari ini untuk halaman Absensi guru: libur atau tidak, dan status permintaan buka absen miliknya. */
function apiDayStatus(token) {
  const s = requireSession_(token), set = getSettings_(), n = nowInfo_();
  const di = dayInfo_(set, readHolidays_(), n.y, n.m, n.day);
  const out = { off: !!di.off && s.job !== 'Admin', adminOpen: !!di.off && s.job === 'Admin', name: di.name || '', date: n.date, req: null };
  if (out.off) {
    const q = findReq_(readRequests_(), n.date, s.id);
    if (q) out.req = { status: q.status, note: q.note };
  }
  return out;
}

/** Guru meminta admin membuka absen untuk hari ini (hanya bila hari ini libur / akhir pekan). */
function apiRequestWork(token, p) {
  const sess = requireSession_(token);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const set = getSettings_(), n = nowInfo_();
    if (!String(sess.name || '').trim()) return { ok: false, message: 'Akun Anda belum dihubungkan dengan data guru. Hubungi admin.' };
    const di = dayInfo_(set, readHolidays_(), n.y, n.m, n.day);
    if (!di.off) return { ok: false, message: 'Hari ini hari kerja biasa, Anda bisa langsung absen.' };
    if (sess.job === 'Admin') return { ok: true, status: 'Disetujui' };
    const note = cleanNote_(p && p.note);
    if (note.length < 3) return { ok: false, needNote: true, message: 'Isi alasan masuk di hari libur (minimal 3 huruf).' };
    const sh = getReqSheet_(), cur = findReq_(readRequests_(), n.date, sess.id), stamp = n.date + ' ' + n.time.slice(0, 5);
    if (cur && cur.status === 'Disetujui') return { ok: true, status: 'Disetujui' };
    if (cur) {
      sh.getRange(cur.row, 4, 1, 5).setValues([[note, 'Menunggu', stamp, '', '']]);     // ajukan ulang setelah ditolak / ubah alasan
    } else {
      const r = Math.max(sh.getLastRow(), 1) + 1;
      if (r > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), 50);
      sh.getRange(r, 1, 1, 8).setNumberFormat('@');
      sh.getRange(r, 1, 1, 8).setValues([[n.date, sess.id, String(sess.name).trim(), note, 'Menunggu', stamp, '', '']]);
    }
    return { ok: true, status: 'Menunggu' };
  } finally { lock.releaseLock(); }
}

/** Admin: daftar permintaan buka absen (yang menunggu hari ini di atas, lalu terbaru). */
function apiWorkRequests(token) {
  requireAdmin_(token);
  const n = nowInfo_();
  const rank = function (q) { return q.date === n.date && q.status === 'Menunggu' ? 0 : 1; };
  const items = readRequests_().sort(function (a, b) {
    return rank(a) - rank(b) || (a.date < b.date ? 1 : a.date > b.date ? -1 : 0) || (b.row - a.row);
  });
  return { today: n.date, items: items.slice(0, 60) };
}

/** Admin: setujui ('Disetujui') atau tolak / cabut ('Ditolak') permintaan buka absen. Hanya berlaku untuk hari ini. */
function apiDecideWork(token, date, id, decision) {
  const adm = requireAdmin_(token);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const n = nowInfo_();
    const q = findReq_(readRequests_(), String(date || ''), id);
    if (!q) throw new Error('Permintaan tidak ditemukan.');
    if (q.date < n.date) throw new Error('Permintaan ini sudah lewat tanggalnya.');
    const st = decision === 'Disetujui' ? 'Disetujui' : 'Ditolak', sh = getReqSheet_();
    sh.getRange(q.row, 5).setValue(st);
    sh.getRange(q.row, 7, 1, 2).setValues([[adm.name || adm.id, n.date + ' ' + n.time.slice(0, 5)]]);
    return { ok: true, status: st };
  } finally { lock.releaseLock(); }
}

/** Jalankan SEKALI (menu "Izinkan akses Drive" atau Run di editor) agar pemilik script menyetujui izin menulis ke Drive. */
function authorizeDrive() {
  const f = DriveApp.getFolderById(PROOF_FOLDER_ID);
  const t = f.createFile('tes-izin.txt', 'ok');
  t.setTrashed(true);
  const msg = 'Izin Drive aktif. Folder bukti: ' + f.getName();
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) {}
}

