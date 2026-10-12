/**
 * ============================================================
 *  APLIKASI ABSENSI GURU  —  Google Apps Script (container-bound)
 * ============================================================
 *  Struktur sheet bulanan (sama dengan file "Data Absen" Anda):
 *    A1  = tahun            Baris 2 = NO | NAMA | 1..31 | REKAP
 *    Baris 3 dst = data guru. Tanggal 1 = kolom C, tanggal 31 = kolom AG, REKAP = kolom AH
 *
 *  Keterangan : H = Hadir (kode lama "O" dianggap Hadir), I = Izin, S = Sakit,
 *               T = Telat, TJ = Tidak ada jadwal mengajar,
 *               A = Alpa (tidak masuk tanpa keterangan, diinput admin),
 *               BA = Belum Absen (diisi OTOMATIS bila guru lupa absen pada hari kerja yang sudah lewat),
 *               L = Libur (diisi otomatis dari sheet "HariLibur": libur nasional + libur khusus)
 *  LOGIN      : sheet "Akun" (ID Login, Password terenkripsi, Nama Guru, Job Admin/Guru, Status).
 *               Job Guru  = hanya bisa absen untuk dirinya sendiri.  Job Admin = semua menu terbuka.
 *  LOKASI     : absen "Hadir" hanya diterima bila HP berada dalam radius sekolah (sheet Pengaturan).
 *  Aturan     : Absen "Hadir" setelah jam masuk (default 07:00) otomatis menjadi T (Telat).
 *               Sel BA dikoreksi admin (menu Koreksi Absensi) dengan mengisi jam masuk -> sistem menentukan H atau T.
 */

const TZ = 'Asia/Jakarta';
const MONTHS = ['Januari','Februari','Maret','April','Mei','Juni','Juli','Agustus','September','Oktober','November','Desember'];
const AY_ORDER = [7,8,9,10,11,0,1,2,3,4,5,6];          // urutan tahun ajaran: Agustus ... Juli
const FIRST_ROW = 3, DAY_COL = 3, REKAP_COL = 34;
const KOP_FILE_ID = '1OWMGOVpoGccenAmsM8cR7yvX0WQg3O8a'; // file kop surat di Google Drive
const SHEET_ACC = 'Akun';
const SHEET_SET = 'Pengaturan', SHEET_LOG = 'LogAbsen', SHEET_HOL = 'HariLibur', SHEET_GURU = 'Nama Guru';
const SHEET_REQ = 'IzinMasuk';                                    // permintaan guru untuk absen di hari libur / Sabtu / Minggu
const PROOF_FOLDER_ID = '1ephYczAoml0nh4TNM3E1JwtSLO6Hah3A';     // folder Google Drive bukti Izin / Sakit
const PROOF_MAX_BYTES = 6 * 1024 * 1024;
const LOGO_FILE_ID = '1lZenwkPyvb2GwpDsGoea9R1VnSac8ZsT';        // logo sekolah di Google Drive (ganti ID bila logo berubah)                          // batas ukuran file bukti (6 MB)
const STATUS = { H: 'Hadir', I: 'Izin', S: 'Sakit', T: 'Telat', TJ: 'Tidak Ada Jadwal', A: 'Alpa', BA: 'Belum Absen' };
const VALID = ['H', 'I', 'S', 'T', 'TJ', 'A'];            // kode yang benar-benar diinput (BA diisi otomatis oleh sistem)
const CNT = ['H', 'I', 'S', 'T', 'TJ', 'A', 'BA'];
function zero_() { return { H: 0, I: 0, S: 0, T: 0, TJ: 0, A: 0, BA: 0 }; }

const DEFAULTS = [
  ['NamaSekolah', 'SMK NASIONAL BANDUNG', 'Nama sekolah (tampil di judul aplikasi)'],
  ['JamMasuk', '07:00', 'Batas jam masuk (HH:mm). Absen Hadir setelah jam ini otomatis Telat'],
  ['TahunAjaranMulai', 2026, 'Tahun mulai tahun ajaran (Agustus s/d Juli). Contoh 2026 = TA 2026/2027'],
  ['HariKerja', '5', 'Hari kerja per minggu: 5 (Senin-Jumat) atau 6 (Senin-Sabtu). Sabtu/Minggu di luar hari kerja otomatis libur'],
  ['CutiBersama', 'Libur', 'Cuti bersama dianggap: Libur atau Masuk (sekolah swasta boleh memilih Masuk)'],
  ['Kota', 'Bandung', 'Kota untuk tanda tangan pada cetakan'],
  ['Jabatan', 'Kepala Sekolah', 'Jabatan penandatangan'],
  ['NamaPenandatangan', 'Nama Kepala Sekolah, M.Pd.', 'Nama penandatangan'],
  ['NIP', '', 'NIP penandatangan (kosongkan jika tidak ada)'],
  ['PasswordAwal', '123456', 'Password awal akun baru / hasil reset. Wajib diganti pengguna saat login pertama'],
  ['LokasiWajib', 'Ya', 'Ya = absen Hadir hanya bisa dari dalam radius sekolah. Tidak = tanpa pengecekan lokasi'],
  ['SekolahLat', '-6.890437', 'Latitude titik pusat sekolah (klik kanan titik sekolah di Google Maps untuk menyalin)'],
  ['SekolahLng', '107.626437', 'Longitude titik pusat sekolah'],
  ['RadiusMeter', '100', 'Radius area absen dari titik pusat sekolah (meter)'],
  ['MaxAkurasi', '100', 'Batas akurasi GPS (meter). Sinyal lokasi yang lebih buruk dari ini ditolak']
];

