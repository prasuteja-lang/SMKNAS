/* ------------------------------------------------------------ */
/*  Penulisan absensi                                            */
/* ------------------------------------------------------------ */
function segment_(day, st, note, minLate) {
  const n = note || '';
  switch (st) {
    case 'I': return 'tgl ' + day + ' izin: ' + n;
    case 'S': return 'tgl ' + day + ' sakit: ' + n;
    case 'T': return 'tgl ' + day + ' telat ' + minLate + ' mnt' + (n ? ' (' + n + ')' : '');
    case 'TJ': return 'tgl ' + day + ' tidak ada jadwal' + (n ? ' (' + n + ')' : '');
    case 'A': return 'tgl ' + day + ' alpa' + (n ? ' (' + n + ')' : '');
    case 'H': return n ? 'tgl ' + day + ' hadir (' + n + ')' : '';
  }
  return '';
}

function mergeRekap_(old, day, seg) {
  const re = new RegExp('^tgl\\s*' + day + '(?!\\d)', 'i');
  // hanya hapus segmen buatan aplikasi untuk tanggal yg sama (catatan manual yg memuat banyak "tgl" dibiarkan)
  const parts = String(old || '').split(';').map(function (s) { return s.trim(); }).filter(Boolean)
    .filter(function (s) { return !(re.test(s) && s.toLowerCase().indexOf('tgl', 3) < 0); });
  if (seg) parts.push(seg);
  const items = parts.map(function (s, i) {
    const m = /^tgl\s*(\d+)/i.exec(s);
    return { s: s, d: m ? +m[1] : 999, i: i };
  });
  items.sort(function (a, b) { return a.d - b.d || a.i - b.i; });
  return items.map(function (o) { return o.s; }).join('; ');
}

function writeEntry_(sh, row, no, name, day, dateStr, timeStr, st, note, minLate, source, extra) {
  const cell = sh.getRange(row, DAY_COL + day - 1);
  if (st) cell.setValue(st); else cell.clearContent();
  const rc = sh.getRange(row, REKAP_COL);
  rc.setValue(mergeRekap_(String(rc.getValue() || ''), day, segment_(day, st, note, minLate)));
  appendLog_(no, name, dateStr, timeStr, st, note, minLate, source, extra);
}

/* ------------------------------------------------------------ */
/*  API publik (halaman absen guru)                              */
/* ------------------------------------------------------------ */
/** Data minimal untuk halaman login (tanpa data guru). */
function apiPublicInit() {
  const set = getSettings_();
  return { school: set.NamaSekolah, jamMasuk: set.JamMasuk, serverMs: Date.now() };
}

/** Ringkasan jumlah absensi hari ini untuk halaman Home (tanpa nama/catatan, aman untuk publik). */
function apiHomeSummary(token) {
  requireSession_(token);
  const set = getSettings_(), n = nowInfo_();
  const out = { school: set.NamaSekolah, jamMasuk: set.JamMasuk, counts: { H: 0, I: 0, S: 0, T: 0, TJ: 0, A: 0, BA: 0, belum: 0, total: 0 }, error: '' };
  if (monthYear_(set, n.m) !== n.y) {
    out.error = 'Tahun ajaran pada sheet Pengaturan belum sesuai dengan tanggal hari ini. Hubungi admin.';
    return out;
  }
  const di = dayInfo_(set, readHolidays_(), n.y, n.m, n.day);
  if (di.off) { out.libur = di.name || 'Libur'; return out; }
  const md = readMonth_(ensureMonthSheet_(n.m, set), n.m, set);
  md.teachers.forEach(function (t) {
    const c = t.cells[n.day - 1];
    out.counts.total++;
    if (c && c !== 'L') out.counts[c]++; else out.counts.belum++;
  });
  return out;
}

function apiSubmit(token, p) {
  const sess = requireSession_(token);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const set = getSettings_(), n = nowInfo_();
    if (monthYear_(set, n.m) !== n.y) return { ok: false, message: 'Tahun ajaran belum diatur dengan benar. Hubungi admin.' };
    // Nama diambil dari AKUN yang login — bukan dari kiriman browser — sehingga guru tidak bisa mengabsenkan guru lain.
    const name = String(sess.name || '').trim();
    const st = String((p && p.status) || '').toUpperCase();
    const note = cleanNote_(p && p.note);
    if (!name) return { ok: false, message: 'Akun Anda belum dihubungkan dengan data guru. Hubungi admin.' };
    if (['H', 'I', 'S', 'TJ'].indexOf(st) < 0) return { ok: false, message: 'Keterangan tidak valid.' };

    const di = dayInfo_(set, readHolidays_(), n.y, n.m, n.day);
    if (di.off) {
      // hari libur / Sabtu / Minggu: guru hanya boleh absen bila admin sudah menyetujui permintaan "Buka Absen"
      // akun Admin yang terhubung ke nama guru tidak perlu meminta: akses hari libur langsung terbuka
      const rq = findReq_(readRequests_(), n.date, sess.id);
      if (sess.job !== 'Admin' && (!rq || rq.status !== 'Disetujui')) {
        return { ok: false, off: true, message: 'Hari ini libur' + (di.name ? ': ' + di.name : '') + '. Jika Anda tetap masuk, ajukan "Minta Buka Absen" ke admin dari menu Absensi.' };
      }
    }

    const sh = ensureMonthSheet_(n.m, set);
    let row = findTeacherRow_(sh, name);
    if (!row && syncMonthNames_(sh)) row = findTeacherRow_(sh, name);
    if (!row) return { ok: false, message: 'Nama "' + name + '" tidak ditemukan di sheet ' + MONTHS[n.m] + '. Hubungi admin.' };

    const cell = sh.getRange(row, DAY_COL + n.day - 1);
    const raw = cell.getDisplayValue().trim();
    const exist = normalizeCode_(raw);
    if (exist && exist !== 'L' && exist !== 'BA') return { ok: false, duplicate: true, message: 'Anda sudah tercatat "' + STATUS[exist] + '" hari ini. Jika salah, hubungi admin.' };
    if (raw && exist !== 'L' && exist !== 'BA') return { ok: false, message: 'Kolom tanggal hari ini sudah berisi tanda "' + raw + '" (mis. libur). Hubungi admin.' };

    // Pengecekan lokasi: wajib untuk "Hadir". Izin / Sakit / Tidak ada jadwal tidak harus di sekolah.
    const loc = checkLocation_(set, p);
    const no = sh.getRange(row, 1).getDisplayValue();
    const extra = { user: sess.id, lat: loc.lat, lng: loc.lng, dist: loc.dist, acc: loc.acc };
    if (st === 'H' && set.LokasiWajib === 'Ya' && !loc.ok) {
      appendLog_(no, name, n.date, n.time, 'DITOLAK', 'Lokasi: ' + loc.message, 0, 'Ditolak-Lokasi', extra);
      return { ok: false, loc: true, message: loc.message, dist: loc.dist, radius: set.RadiusMeter };
    }

    let final = st, minLate = 0;
    if (st === 'H') {
      const lim = timeToSec_(set.JamMasuk);
      if (n.sec > lim) { final = 'T'; minLate = Math.ceil((n.sec - lim) / 60); }
    }
    if ((final === 'I' || final === 'S' || final === 'T') && note.length < 3) {
      return {
        ok: false, needNote: true, status: final, minLate: minLate,
        message: final === 'T'
          ? 'Anda terlambat ' + minLate + ' menit (lewat ' + set.JamMasuk + '). Isi alasan keterlambatan, lalu kirim lagi.'
          : 'Catatan alasan wajib diisi untuk ' + STATUS[final] + '.'
      };
    }
    // bukti (foto / PDF) hanya untuk Izin & Sakit; diunggah ke folder Drive sekolah sebelum absen dicatat
    let proofUrl = '';
    if ((final === 'I' || final === 'S') && p && p.file && p.file.data) {
      try { proofUrl = saveProof_(p.file, name, n.date, final); }
      catch (e) { return { ok: false, message: 'Bukti gagal diunggah: ' + e.message + ' Coba lagi, atau kirim tanpa bukti.' }; }
      extra.proof = proofUrl;
    }
    writeEntry_(sh, row, no, name, n.day, n.date, n.time, final, note, minLate, 'Guru', extra);
    return { ok: true, status: final, minLate: minLate, time: n.time.slice(0, 5), name: name, jamMasuk: set.JamMasuk, dist: loc.dist, proof: !!proofUrl };
  } finally {
    lock.releaseLock();
  }
}

/* ------------------------------------------------------------ */
/*  Lokasi (radius sekolah)                                      */
/* ------------------------------------------------------------ */
function distM_(lat1, lng1, lat2, lng2) {
  const R = 6371000, rad = Math.PI / 180;
  const dLat = (lat2 - lat1) * rad, dLng = (lng2 - lng1) * rad;
  const a = Math.sin(dLat / 2) * Math.sin(dLat / 2) + Math.cos(lat1 * rad) * Math.cos(lat2 * rad) * Math.sin(dLng / 2) * Math.sin(dLng / 2);
  return Math.round(R * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a)));
}

/** Hasil: { ok, message, lat, lng, acc, dist }. ok=false bila lokasi tidak ada / terlalu jauh / terlalu tidak akurat. */
function checkLocation_(set, p) {
  const lat = parseFloat(p && p.lat), lng = parseFloat(p && p.lng), acc = Math.round(parseFloat(p && p.acc));
  const out = { ok: false, message: '', lat: null, lng: null, acc: isFinite(acc) ? acc : null, dist: null };
  if (!isFinite(lat) || !isFinite(lng) || Math.abs(lat) > 90 || Math.abs(lng) > 180) {
    out.message = 'Lokasi Anda tidak terbaca. Aktifkan GPS dan izinkan akses lokasi di browser, lalu coba lagi.';
    return out;
  }
  out.lat = Math.round(lat * 1e6) / 1e6; out.lng = Math.round(lng * 1e6) / 1e6;
  const sLat = parseFloat(set.SekolahLat), sLng = parseFloat(set.SekolahLng);
  if (!isFinite(sLat) || !isFinite(sLng)) { out.message = 'Titik lokasi sekolah belum diatur oleh admin (menu Pengaturan).'; return out; }
  out.dist = distM_(lat, lng, sLat, sLng);
  if (out.acc != null && out.acc > set.MaxAkurasi) {
    out.message = 'Akurasi GPS terlalu rendah (±' + out.acc + ' m, maksimal ' + set.MaxAkurasi + ' m). Pindah ke tempat terbuka / aktifkan GPS "akurasi tinggi", lalu perbarui lokasi.';
    return out;
  }
  if (out.dist > set.RadiusMeter) {
    out.message = 'Anda berada ±' + out.dist + ' m dari sekolah. Absen Hadir hanya bisa dilakukan dalam radius ' + set.RadiusMeter + ' m dari lingkungan sekolah.';
    return out;
  }
  out.ok = true;
  return out;
}

function apiKop(token) {
  requireAdmin_(token);
  try {
    const f = DriveApp.getFileById(KOP_FILE_ID);
    const b = f.getBlob();
    const mt = b.getContentType();
    if (/^image\//.test(mt)) return { ok: true, src: 'data:' + mt + ';base64,' + Utilities.base64Encode(b.getBytes()) };
    const th = f.getThumbnail();
    if (th) return { ok: true, lowRes: true, src: 'data:' + th.getContentType() + ';base64,' + Utilities.base64Encode(th.getBytes()) };
    return { ok: false, message: 'Format file kop tidak didukung. Gunakan PNG/JPG.' };
  } catch (e) {
    return { ok: false, message: 'Kop surat tidak dapat dibuka: ' + e.message };
  }
}

