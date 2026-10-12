/* ------------------------------------------------------------ */
/*  Bukti Izin / Sakit  ->  folder Google Drive                  */
/* ------------------------------------------------------------ */
function saveProof_(f, name, dateStr, st) {
  const mime = String(f.mime || '').toLowerCase();
  const okMime = ['image/jpeg', 'image/png', 'image/webp', 'image/heic', 'image/heif', 'application/pdf'];
  if (okMime.indexOf(mime) < 0) throw new Error('Format file harus foto (JPG/PNG) atau PDF.');
  const bytes = Utilities.base64Decode(String(f.data));
  if (!bytes.length) throw new Error('File kosong.');
  if (bytes.length > PROOF_MAX_BYTES) throw new Error('Ukuran file maksimal ' + Math.round(PROOF_MAX_BYTES / 1048576) + ' MB.');
  const ext = mime === 'application/pdf' ? 'pdf' : mime === 'image/png' ? 'png' : mime === 'image/webp' ? 'webp' : /hei[cf]/.test(mime) ? 'heic' : 'jpg';
  const safe = String(name || '').replace(/[^A-Za-z0-9 ]+/g, '').trim().replace(/\s+/g, '_').slice(0, 40) || 'guru';
  const fname = dateStr + '_' + STATUS[st] + '_' + safe + '_' + Utilities.formatDate(new Date(), TZ, 'HHmmss') + '.' + ext;
  const file = DriveApp.getFolderById(PROOF_FOLDER_ID).createFile(Utilities.newBlob(bytes, mime, fname));
  file.setDescription('Bukti ' + STATUS[st] + ' - ' + name + ' - ' + dateStr);
  return file.getUrl();
}

