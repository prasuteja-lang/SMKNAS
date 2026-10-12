/* ------------------------------------------------------------ */
/*  Logo sekolah (publik, tampil di halaman login & sidebar)      */
/* ------------------------------------------------------------ */
function apiLogo() {
  const cache = CacheService.getScriptCache();
  const n = +(cache.get('logo_n') || 0);
  if (n > 0) {
    const keys = [];
    for (let i = 0; i < n; i++) keys.push('logo_' + i);
    const got = cache.getAll(keys);
    let src = '', ok = true;
    for (let i = 0; i < n; i++) { if (!got['logo_' + i]) { ok = false; break; } src += got['logo_' + i]; }
    if (ok) return { ok: true, src: src };
  }
  try {
    const f = DriveApp.getFileById(LOGO_FILE_ID);
    let b = f.getBlob();
    if (!/^image\//.test(b.getContentType()) || b.getBytes().length > 300 * 1024) b = f.getThumbnail() || b;   // file besar / bukan gambar -> pakai thumbnail
    if (!/^image\//.test(b.getContentType())) return { ok: false, message: 'Logo bukan gambar.' };
    const src = 'data:' + b.getContentType() + ';base64,' + Utilities.base64Encode(b.getBytes());
    const parts = {};
    let cnt = 0;
    for (let i = 0; i < src.length; i += 90000) parts['logo_' + (cnt++)] = src.slice(i, i + 90000);
    parts.logo_n = String(cnt);
    try { cache.putAll(parts, 21600); } catch (e) {}
    return { ok: true, src: src };
  } catch (e) {
    return { ok: false, message: e.message };
  }
