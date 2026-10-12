/* ------------------------------------------------------------ */
/*  Helper dasar                                                 */
/* ------------------------------------------------------------ */
function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }

function nowInfo_() {
  const d = new Date();
  const date = Utilities.formatDate(d, TZ, 'yyyy-MM-dd');
  const time = Utilities.formatDate(d, TZ, 'HH:mm:ss');
  const p = date.split('-'), t = time.split(':');
  return { date: date, time: time, y: +p[0], m: +p[1] - 1, day: +p[2], sec: (+t[0]) * 3600 + (+t[1]) * 60 + (+t[2]) };
}

function timeToSec_(hhmm) {
  const m = /(\d{1,2}):(\d{2})/.exec(hhmm || '07:00');
  return m ? (+m[1]) * 3600 + (+m[2]) * 60 : 7 * 3600;
}

function monthYear_(set, mIdx) { return mIdx >= 7 ? set.TahunAjaranMulai : set.TahunAjaranMulai + 1; }
function daysIn_(year, mIdx) { return new Date(year, mIdx + 1, 0).getDate(); }
function normName_(s) { return String(s || '').toLowerCase().replace(/\s+/g, ' ').trim(); }
function cleanNote_(s) { return String(s || '').replace(/[;\r\n]+/g, ', ').replace(/\s+/g, ' ').trim().slice(0, 300); }

function normalizeCode_(v) {
  const s = String(v == null ? '' : v).trim().toUpperCase();
  if (s === 'O') return 'H';
  if (s === 'L') return 'L';
  return STATUS[s] ? s : '';
}

