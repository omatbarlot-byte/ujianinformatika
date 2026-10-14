/**
 * =====================================================================
 *  CBT INFORMATIKA SMA NUGRAHA — BACKEND (Google Apps Script Web App)
 * =====================================================================
 *  Deploy: Extensions > Apps Script > tempel file ini > Deploy > New deployment
 *          Type: Web app | Execute as: Me | Who has access: Anyone
 *  Lalu salin URL "/exec" ke CONFIG.API_URL pada index.html.
 *
 *  Prinsip keamanan:
 *   - Password TIDAK PERNAH dikirim ke browser. Validasi dilakukan di server.
 *   - Link Google Form hanya dikirim saat siswa menekan "Mulai Ujian".
 *   - Hitungan pelanggaran, kunci, dan waktu berakhir dijaga SERVER (bukan browser),
 *     sehingga refresh / buka tab baru tidak mereset pelanggaran.
 * =====================================================================
 */

const SS_ID = '1Xh-Bqr-lOCrUCopS-VB0e5xCTvWIF-17lXDd-CFsU6c';

const SHEET_NAMES = {
  SISWA: ['Siswa'],
  GURU: ['Guru / Admin', 'Guru/Admin', 'Guru', 'Admin'],
  EXAM: ['Pengaturan Ujian'],
  SESI: ['Sesi Ujian'],
  LOG: ['Log Pelanggaran']
};
const EXAM_HEAD = ['ID Ujian', 'Nama Sekolah', 'Nama Ujian', 'Link Google Form', 'Batas Pelanggaran', 'Durasi (Menit)', 'Status (Aktif/Nonaktif)'];
const SESI_HEAD = ['Sesi ID', 'Nama Sekolah', 'Nomor Peserta', 'Nama Siswa', 'Kelas', 'ID Ujian', 'Nama Ujian', 'Mulai (ms)', 'Berakhir (ms)', 'Pelanggaran', 'Status', 'Terakhir Aktif (ms)', 'Waktu Selesai', 'Catatan'];
const LOG_HEAD = ['Waktu', 'Sesi ID', 'Nama Sekolah', 'Nomor Peserta', 'Nama Siswa', 'ID Ujian', 'Jenis', 'Detail', 'Pelanggaran Ke'];

const ST = { RUN: 'Sedang Mengerjakan', LOCK: 'Terkunci', DONE: 'Selesai' };
const TOKEN_TTL = 6 * 60 * 60;      // detik
const ONLINE_WINDOW_MS = 60 * 1000; // dianggap online jika heartbeat < 60 detik
const MAX_LOGIN_FAIL = 5;

/* ----------------------------- Router ------------------------------ */

const ACTIONS = {
  getSchools: getSchools_,
  loginSiswa: loginSiswa_,
  loginAdmin: loginAdmin_,
  getExams: getExams_,
  startExam: startExam_,
  heartbeat: heartbeat_,
  logViolation: logViolation_,
  finishExam: finishExam_,
  saveExam: saveExam_,
  deleteExam: deleteExam_,
  getMonitor: getMonitor_,
  resetSession: resetSession_,
  getLogs: getLogs_
};
const WRITE_ACTIONS = ['startExam', 'heartbeat', 'logViolation', 'finishExam', 'saveExam', 'deleteExam', 'resetSession'];

function doGet() {
  return json_({ ok: true, service: 'CBT Informatika SMA Nugraha API', time: Date.now() });
}

function doPost(e) {
  let lock = null;
  try {
    const req = JSON.parse(e.postData.contents);
    const fn = ACTIONS[req.action];
    if (!fn) return json_({ ok: false, error: 'Aksi tidak dikenal.' });
    if (WRITE_ACTIONS.indexOf(req.action) > -1) {
      lock = LockService.getScriptLock();
      lock.waitLock(15000);
    }
    return json_(Object.assign({ ok: true }, fn(req)));
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message ? err.message : err) });
  } finally {
    if (lock) { try { lock.releaseLock(); } catch (x) {} }
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* ----------------------------- Helpers ----------------------------- */

function sheet_(names, create, head) {
  const ss = SpreadsheetApp.openById(SS_ID);
  const list = [].concat(names);
  for (let i = 0; i < list.length; i++) {
    const s = ss.getSheetByName(list[i]);
    if (s) return s;
  }
  const all = ss.getSheets();
  for (let i = 0; i < list.length; i++) {
    const hit = all.filter(function (x) { return x.getName().trim().toLowerCase() === list[i].toLowerCase(); })[0];
    if (hit) return hit;
  }
  if (create) {
    const s = ss.insertSheet(list[0]);
    s.appendRow(head);
    s.setFrozenRows(1);
    s.getRange(1, 1, 1, head.length).setFontWeight('bold').setBackground('#4F46E5').setFontColor('#FFFFFF');
    return s;
  }
  throw new Error('Sheet "' + list[0] + '" tidak ditemukan di spreadsheet.');
}

function norm_(s) { return String(s).trim().toLowerCase().replace(/\s+/g, ' '); }
function s_(v) { return String(v == null ? '' : v).trim(); }

/** Baca sheet menjadi array objek; kunci = header (lowercase). */
function table_(sh) {
  const v = sh.getDataRange().getValues();
  if (v.length < 2) return [];
  const h = v[0].map(norm_);
  const out = [];
  for (let i = 1; i < v.length; i++) {
    const o = { _row: i + 1 };
    let filled = false;
    for (let j = 0; j < h.length; j++) { o[h[j]] = v[i][j]; if (s_(v[i][j]) !== '') filled = true; }
    if (filled) out.push(o);
  }
  return out;
}

function auth_(req, role) {
  if (!req.token) throw new Error('Token tidak ada. Silakan login.');
  const raw = CacheService.getScriptCache().get('tok_' + req.token);
  if (!raw) throw new Error('Sesi login berakhir. Silakan login ulang.');
  const u = JSON.parse(raw);
  if (role && u.role !== role) throw new Error('Akses ditolak.');
  return u;
}

function issueToken_(user) {
  const token = Utilities.getUuid() + Utilities.getUuid().replace(/-/g, '');
  CacheService.getScriptCache().put('tok_' + token, JSON.stringify(user), TOKEN_TTL);
  return token;
}

function isWildcard_(school) {
  const a = s_(school).toLowerCase();
  return a === 'semua' || a === '*' || a === 'all';
}
function canSchool_(admin, school) {
  return isWildcard_(admin.school) || s_(admin.school).toLowerCase() === s_(school).toLowerCase();
}

function throttle_(key) {
  const c = CacheService.getScriptCache();
  const n = Number(c.get(key) || 0);
  if (n >= MAX_LOGIN_FAIL) throw new Error('Terlalu banyak percobaan gagal. Coba lagi dalam 5 menit.');
  return {
    fail: function () { c.put(key, String(n + 1), 300); },
    ok: function () { c.remove(key); }
  };
}

/** Ambil src jika admin menempelkan kode <iframe>, pastikan ?embedded=true. */
function toEmbed_(raw) {
  let url = s_(raw);
  const m = url.match(/src=["']([^"']+)["']/i);
  if (m) url = m[1].replace(/&amp;/g, '&');
  if (!/^https:\/\/docs\.google\.com\/forms\//i.test(url)) {
    throw new Error('Link harus berupa Google Form (https://docs.google.com/forms/...). Link pendek forms.gle tidak bisa di-embed.');
  }
  if (url.indexOf('embedded=true') === -1) url += (url.indexOf('?') > -1 ? '&' : '?') + 'embedded=true';
  return url;
}

function examRows_() { return table_(sheet_(SHEET_NAMES.EXAM)); }

function examOut_(r, withLink) {
  const o = {
    id: s_(r['id ujian']),
    school: s_(r['nama sekolah']),
    name: s_(r['nama ujian']),
    max: Number(r['batas pelanggaran']) || 3,
    duration: Number(r['durasi (menit)']) || 60,
    status: /^aktif$/i.test(s_(r['status (aktif/nonaktif)'])) ? 'Aktif' : 'Nonaktif'
  };
  if (withLink) o.link = s_(r['link google form']);
  return o;
}

function findExam_(id) {
  const r = examRows_().filter(function (x) { return s_(x['id ujian']) === s_(id); })[0];
  if (!r) throw new Error('Ujian tidak ditemukan.');
  return r;
}

/* ---- Sesi ---- */
function sesiSheet_() { return sheet_(SHEET_NAMES.SESI, true, SESI_HEAD); }
function logSheet_() { return sheet_(SHEET_NAMES.LOG, true, LOG_HEAD); }

function rowToSesi_(row, v) {
  return {
    _row: row, id: s_(v[0]), school: s_(v[1]), nomor: s_(v[2]), nama: s_(v[3]), kelas: s_(v[4]),
    examId: s_(v[5]), examName: s_(v[6]), start: Number(v[7]), end: Number(v[8]),
    violations: Number(v[9]) || 0, status: s_(v[10]), last: Number(v[11]) || 0, finishedAt: s_(v[12]), note: s_(v[13])
  };
}
function findSesi_(sh, id) {
  if (!id) throw new Error('Sesi tidak valid.');
  const f = sh.getRange('A:A').createTextFinder(String(id)).matchEntireCell(true).findNext();
  if (!f) throw new Error('Sesi tidak ditemukan.');
  const row = f.getRow();
  return rowToSesi_(row, sh.getRange(row, 1, 1, SESI_HEAD.length).getValues()[0]);
}
function allSesi_(sh) {
  const last = sh.getLastRow();
  if (last < 2) return [];
  const vals = sh.getRange(2, 1, last - 1, SESI_HEAD.length).getValues();
  return vals.map(function (v, i) { return rowToSesi_(i + 2, v); }).filter(function (x) { return x.id; });
}
function setSesi_(sh, s, patch) {
  Object.keys(patch).forEach(function (k) {
    const col = { violations: 10, status: 11, last: 12, finishedAt: 13, note: 14, end: 9 }[k];
    if (col) { sh.getRange(s._row, col).setValue(patch[k]); s[k] = patch[k]; }
  });
}
function stateOf_(s, max) {
  return { id: s.id, status: s.status, violations: s.violations, max: max, endsAt: s.end, serverNow: Date.now(), note: s.note };
}
/** Tutup sesi otomatis jika waktu habis. */
function expire_(sh, s) {
  if (s.status === ST.RUN && Date.now() > s.end + 3000) {
    setSesi_(sh, s, { status: ST.DONE, finishedAt: new Date(), note: 'Waktu habis' });
  }
  return s;
}
function ownSesi_(sh, req, user) {
  const s = findSesi_(sh, req.sessionId);
  if (s.nomor !== user.nomor || s.school !== user.school) throw new Error('Akses ditolak.');
  return s;
}

/* ----------------------------- Actions ----------------------------- */

function getSchools_() {
  const set = {};
  table_(sheet_(SHEET_NAMES.SISWA)).forEach(function (r) { const n = s_(r['nama sekolah']); if (n) set[n] = 1; });
  return { schools: Object.keys(set).sort() };
}

function loginSiswa_(req) {
  const school = s_(req.school), nomor = s_(req.nomor), pass = s_(req.password);
  if (!school || !nomor || !pass) throw new Error('Lengkapi sekolah, nomor peserta, dan password.');
  const th = throttle_('fail_s_' + school + '_' + nomor);
  const r = table_(sheet_(SHEET_NAMES.SISWA)).filter(function (x) {
    return s_(x['nama sekolah']).toLowerCase() === school.toLowerCase() && s_(x['nomor peserta']) === nomor;
  })[0];
  if (!r || s_(r['password']) !== pass) { th.fail(); throw new Error('Nomor peserta atau password salah.'); }
  th.ok();
  const user = { role: 'siswa', school: s_(r['nama sekolah']), nomor: nomor, nama: s_(r['nama siswa']), kelas: s_(r['kelas']) };
  return { token: issueToken_(user), user: user, exams: listExamsForStudent_(user) };
}

function loginAdmin_(req) {
  const username = s_(req.username), pass = s_(req.password);
  if (!username || !pass) throw new Error('Lengkapi username dan password.');
  const th = throttle_('fail_a_' + username);
  const r = table_(sheet_(SHEET_NAMES.GURU)).filter(function (x) { return s_(x['username / nip']) === username; })[0];
  if (!r || s_(r['password']) !== pass) { th.fail(); throw new Error('Username/NIP atau password salah.'); }
  th.ok();
  const user = { role: 'admin', school: s_(r['nama sekolah']), nama: s_(r['nama guru']), username: username };
  user.wildcard = isWildcard_(user.school);
  return { token: issueToken_(user), user: user, exams: listExamsForAdmin_(user), schools: getSchools_().schools };
}

function listExamsForStudent_(user) {
  const sh = sesiSheet_();
  const mine = allSesi_(sh).filter(function (s) { return s.nomor === user.nomor && s.school === user.school; });
  return examRows_()
    .filter(function (r) { return s_(r['nama sekolah']).toLowerCase() === user.school.toLowerCase() && /^aktif$/i.test(s_(r['status (aktif/nonaktif)'])); })
    .map(function (r) {
      const e = examOut_(r, false);
      const s = mine.filter(function (x) { return x.examId === e.id; })[0];
      e.myStatus = s ? s.status : null;
      return e;
    });
}
function listExamsForAdmin_(user) {
  return examRows_()
    .filter(function (r) { return canSchool_(user, r['nama sekolah']); })
    .map(function (r) { return examOut_(r, true); });
}

function getExams_(req) {
  const u = auth_(req);
  return { exams: u.role === 'admin' ? listExamsForAdmin_(u) : listExamsForStudent_(u) };
}

function startExam_(req) {
  const u = auth_(req, 'siswa');
  const r = findExam_(req.examId);
  const ex = examOut_(r, true);
  if (ex.school.toLowerCase() !== u.school.toLowerCase()) throw new Error('Ujian ini bukan untuk sekolah Anda.');
  if (ex.status !== 'Aktif') throw new Error('Ujian sedang tidak aktif.');

  const sh = sesiSheet_();
  let s = allSesi_(sh).filter(function (x) { return x.nomor === u.nomor && x.school === u.school && x.examId === ex.id; })[0];
  const now = Date.now();

  if (!s) {
    const id = Utilities.getUuid().substring(0, 8);
    const end = now + ex.duration * 60000;
    sh.getRange(sh.getLastRow() + 1, 1, 1, SESI_HEAD.length).setValues([[id, u.school, "'" + u.nomor, u.nama, u.kelas, ex.id, ex.name, now, end, 0, ST.RUN, now, '', 'Mulai']]);
    s = findSesi_(sh, id);
  } else {
    expire_(sh, s);
  }
  if (s.status === ST.DONE) throw new Error('Anda sudah menyelesaikan ujian ini.');
  setSesi_(sh, s, { last: now });

  const out = { session: stateOf_(s, ex.max), exam: { id: ex.id, name: ex.name, school: ex.school, duration: ex.duration, max: ex.max } };
  if (s.status === ST.RUN) out.exam.link = toEmbed_(ex.link); // link hanya diberikan jika tidak terkunci
  return out;
}

function heartbeat_(req) {
  const u = auth_(req, 'siswa');
  const sh = sesiSheet_();
  const s = ownSesi_(sh, req, u);
  expire_(sh, s);
  setSesi_(sh, s, { last: Date.now() });
  const max = examOut_(findExam_(s.examId), false).max;
  const out = { session: stateOf_(s, max) };
  if (s.status === ST.RUN && req.wantLink) out.link = toEmbed_(examOut_(findExam_(s.examId), true).link);
  return out;
}

function logViolation_(req) {
  const u = auth_(req, 'siswa');
  const sh = sesiSheet_();
  const s = ownSesi_(sh, req, u);
  expire_(sh, s);
  const ex = examOut_(findExam_(s.examId), false);
  if (s.status === ST.RUN) {
    const n = s.violations + 1;
    const patch = { violations: n, last: Date.now() };
    if (n >= ex.max) { patch.status = ST.LOCK; patch.note = 'Terkunci otomatis: melebihi batas pelanggaran'; }
    setSesi_(sh, s, patch);
    const lg = logSheet_();
    lg.appendRow([new Date(), s.id, s.school, "'" + s.nomor, s.nama, s.examId, s_(req.type), s_(req.detail).substring(0, 200), n]);
  }
  return { session: stateOf_(s, ex.max) };
}

function finishExam_(req) {
  const u = auth_(req, 'siswa');
  const sh = sesiSheet_();
  const s = ownSesi_(sh, req, u);
  if (s.status !== ST.DONE) {
    setSesi_(sh, s, { status: ST.DONE, finishedAt: new Date(), last: Date.now(), note: s_(req.reason) === 'timeout' ? 'Waktu habis' : 'Selesai oleh siswa' });
  }
  return { session: stateOf_(s, examOut_(findExam_(s.examId), false).max) };
}

function saveExam_(req) {
  const u = auth_(req, 'admin');
  const e = req.exam || {};
  const school = s_(e.school), name = s_(e.name);
  const max = parseInt(e.max, 10), dur = parseInt(e.duration, 10);
  const status = /^aktif$/i.test(s_(e.status)) ? 'Aktif' : 'Nonaktif';
  if (!school || !name) throw new Error('Nama sekolah dan nama ujian wajib diisi.');
  if (!canSchool_(u, school)) throw new Error('Anda tidak berhak mengelola sekolah "' + school + '".');
  if (!(max >= 1 && max <= 50)) throw new Error('Batas pelanggaran harus 1–50.');
  if (!(dur >= 1 && dur <= 600)) throw new Error('Durasi harus 1–600 menit.');
  const link = toEmbed_(e.link);

  const sh = sheet_(SHEET_NAMES.EXAM);
  let id = s_(e.id);
  if (id) {
    const f = sh.getRange('A:A').createTextFinder(id).matchEntireCell(true).findNext();
    if (!f) throw new Error('Ujian tidak ditemukan.');
    const old = findExam_(id);
    if (!canSchool_(u, old['nama sekolah'])) throw new Error('Akses ditolak.');
    sh.getRange(f.getRow(), 1, 1, 7).setValues([[id, school, name, link, max, dur, status]]);
  } else {
    id = 'UJ-' + Date.now().toString(36).toUpperCase();
    sh.appendRow([id, school, name, link, max, dur, status]);
  }
  return { id: id, exams: listExamsForAdmin_(u) };
}

function deleteExam_(req) {
  const u = auth_(req, 'admin');
  const r = findExam_(req.id);
  if (!canSchool_(u, r['nama sekolah'])) throw new Error('Akses ditolak.');
  sheet_(SHEET_NAMES.EXAM).deleteRow(r._row);
  return { exams: listExamsForAdmin_(u) };
}

function getMonitor_(req) {
  const u = auth_(req, 'admin');
  const now = Date.now();
  const sh = sesiSheet_();
  let list = allSesi_(sh).filter(function (s) { return canSchool_(u, s.school); });
  if (req.examId) list = list.filter(function (s) { return s.examId === req.examId; });
  list.forEach(function (s) { if (s.status === ST.RUN && now > s.end + 3000) expire_(sh, s); });
  list.sort(function (a, b) { return b.last - a.last; });
  return {
    serverNow: now,
    sessions: list.slice(0, 400).map(function (s) {
      return {
        id: s.id, school: s.school, nomor: s.nomor, nama: s.nama, kelas: s.kelas, examId: s.examId, examName: s.examName,
        violations: s.violations, status: s.status, last: s.last, end: s.end, note: s.note,
        online: s.status === ST.RUN && now - s.last < ONLINE_WINDOW_MS
      };
    })
  };
}

function resetSession_(req) {
  const u = auth_(req, 'admin');
  const sh = sesiSheet_();
  const s = findSesi_(sh, req.sessionId);
  if (!canSchool_(u, s.school)) throw new Error('Akses ditolak.');
  const who = u.nama + ' (' + u.username + ')';
  if (req.op === 'unlock') {
    const extra = Math.max(0, parseInt(req.extraMinutes, 10) || 0);
    setSesi_(sh, s, { status: ST.RUN, violations: 0, last: Date.now(), end: Math.max(s.end, Date.now()) + extra * 60000, note: 'Dibuka proktor: ' + who });
  } else if (req.op === 'finish') {
    setSesi_(sh, s, { status: ST.DONE, finishedAt: new Date(), note: 'Diselesaikan proktor: ' + who });
  } else if (req.op === 'lock') {
    setSesi_(sh, s, { status: ST.LOCK, note: 'Dikunci proktor: ' + who });
  } else {
    throw new Error('Operasi tidak dikenal.');
  }
  return { ok: true };
}

function getLogs_(req) {
  const u = auth_(req, 'admin');
  const lg = logSheet_();
  const last = lg.getLastRow();
  if (last < 2) return { logs: [] };
  const start = Math.max(2, last - 299);
  const vals = lg.getRange(start, 1, last - start + 1, LOG_HEAD.length).getValues();
  const logs = vals.filter(function (v) { return canSchool_(u, v[2]) && (!req.sessionId || s_(v[1]) === s_(req.sessionId)); })
    .map(function (v) {
      return { time: v[0] instanceof Date ? v[0].getTime() : 0, sessionId: s_(v[1]), school: s_(v[2]), nomor: s_(v[3]), nama: s_(v[4]), examId: s_(v[5]), type: s_(v[6]), detail: s_(v[7]), n: v[8] };
    }).reverse();
  return { logs: logs };
}

/* ------------------- Jalankan SEKALI untuk inisialisasi ------------------- */
function setup() {
  sesiSheet_();
  logSheet_();
  try { sheet_(SHEET_NAMES.EXAM); } catch (e) { sheet_(SHEET_NAMES.EXAM, true, EXAM_HEAD); }
  Logger.log('Setup selesai. Pastikan sheet Siswa & Guru / Admin sudah terisi.');
}