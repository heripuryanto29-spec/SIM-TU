/**
 * SIM-TU v2 - Backend API (Google Apps Script)
 * Frontend dihosting di Vercel (index.html), backend ini hanya melayani JSON.
 *
 * Cara pasang:
 * 1. Buat Google Spreadsheet baru > Extensions > Apps Script
 * 2. Tempel file ini sebagai Code.gs (tidak perlu file HTML di Apps Script)
 * 3. Jalankan fungsi setup() sekali (beri izin akses)
 * 4. Deploy > New deployment > Web app
 *    Execute as: Me | Who has access: Anyone
 * 5. Salin URL Web app (berakhiran /exec) ke API_URL di index.html
 * Login awal: admin / admin123 (segera ganti di sheet Users)
 *
 * Setiap kali kode ini diubah: Deploy > Manage deployments > Edit > New version.
 */

const SHEETS = {
  Users: ['ID', 'Username', 'Password', 'Nama', 'Role'],
  Siswa: ['ID', 'NIS', 'NISN', 'Nama', 'JK', 'Kelas', 'TempatLahir', 'TglLahir', 'Alamat', 'NamaOrtu'],
  Guru: ['ID', 'NIP', 'Nama', 'JK', 'Mapel', 'Jabatan', 'NoHP', 'Status'],
  TahunPelajaran: ['ID', 'Tahun', 'Semester', 'Aktif'],
  Alumni: ['ID', 'NIS', 'NISN', 'Nama', 'JK', 'Kelas', 'TempatLahir', 'TglLahir', 'Alamat', 'NamaOrtu', 'TahunLulus', 'TglLulus']
};
const DATA_SHEETS = ['Siswa', 'Guru', 'TahunPelajaran', 'Alumni'];
const UNIQUE_KEY = { Siswa: 'NIS', Guru: 'NIP' };
const SESSION_SECONDS = 6 * 60 * 60;

/* ---------- Router API ---------- */
// Hanya fungsi di daftar ini yang bisa dipanggil dari luar.
const API = {
  login: function (a) { return login(a[0], a[1]); },
  logout: function (a) { return logout(a[0]); },
  bootstrap: function (a) { return bootstrap(a[0]); },
  saveData: function (a) { return saveData(a[0], a[1], a[2]); },
  deleteData: function (a) { return deleteData(a[0], a[1], a[2]); },
  importData: function (a) { return importData(a[0], a[1], a[2]); },
  graduate: function (a) { return graduate(a[0], a[1]); },
  publicInfo: function () { return publicInfo(); },
  saveSettings: function (a) { return saveSettings(a[0], a[1]); }
};

function doPost(e) {
  let out;
  try {
    const req = JSON.parse(e.postData.contents);
    const fn = API[req.fn];
    if (!fn) throw new Error('Perintah tidak dikenal.');
    out = { ok: true, data: fn(req.args || []) };
  } catch (err) {
    out = { ok: false, error: String(err && err.message ? err.message : err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  return ContentService.createTextOutput(JSON.stringify({ ok: true, app: 'SIM-TU API' })).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- Setup awal ---------- */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(SHEETS).forEach(function (name) {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    const headers = SHEETS[name];
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#12383A').setFontColor('#ffffff');
    sh.getRange(2, 1, sh.getMaxRows() - 1, headers.length).setNumberFormat('@');
    sh.setFrozenRows(1);
  });
  const users = ss.getSheetByName('Users');
  if (users.getLastRow() < 2) users.appendRow([uid_(), 'admin', hash_('admin123'), 'Administrator', 'admin']);
  const ta = ss.getSheetByName('TahunPelajaran');
  if (ta.getLastRow() < 2) {
    const y = new Date().getFullYear();
    ta.appendRow([uid_(), y + '/' + (y + 1), 'Ganjil', 'Aktif']);
  }
}

/* ---------- Helper ---------- */
function uid_() { return Utilities.getUuid().slice(0, 8); }

function hash_(text) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, text, Utilities.Charset.UTF_8)
    .map(function (b) { return ('0' + (b < 0 ? b + 256 : b).toString(16)).slice(-2); }).join('');
}

function sheet_(name) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sh) throw new Error('Sheet "' + name + '" belum ada. Jalankan setup() dulu.');
  return sh;
}

function cell_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return v === null || v === undefined ? '' : String(v);
}

function readAll_(name) {
  const sh = sheet_(name), headers = SHEETS[name], last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, headers.length).getValues().map(function (r) {
    const o = {};
    headers.forEach(function (h, i) { o[h] = cell_(r[i]); });
    return o;
  }).filter(function (o) { return o.ID; });
}

function allData_() {
  const d = {};
  DATA_SHEETS.forEach(function (n) { d[n] = readAll_(n); });
  return d;
}

function rowOf_(name, obj) { return SHEETS[name].map(function (h) { return obj[h] === undefined ? '' : String(obj[h]); }); }

function auth_(token) {
  const raw = token ? CacheService.getScriptCache().get('sess_' + token) : null;
  if (!raw) throw new Error('SESI_HABIS');
  return JSON.parse(raw);
}

function guard_(name) {
  if (DATA_SHEETS.indexOf(name) < 0) throw new Error('Data tidak valid.');
}

function findRow_(sh, id) {
  const last = sh.getLastRow();
  if (last < 2) return -1;
  const ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  return -1;
}

/** Hapus banyak baris sekaligus (jauh lebih cepat daripada deleteRow satu-satu). */
function removeIds_(name, idList) {
  const sh = sheet_(name), n = SHEETS[name].length, last = sh.getLastRow();
  if (last < 2) return 0;
  const set = {}; idList.forEach(function (i) { set[String(i)] = true; });
  const rng = sh.getRange(2, 1, last - 1, n);
  const vals = rng.getValues();
  const keep = vals.filter(function (r) { return !set[String(r[0])]; });
  const removed = vals.length - keep.length;
  if (removed) {
    rng.clearContent();
    if (keep.length) {
      sh.getRange(2, 1, keep.length, n).setNumberFormat('@')
        .setValues(keep.map(function (r) { return r.map(cell_); }));
    }
  }
  return removed;
}

/* ---------- Auth ---------- */
function login(username, password) {
  const u = readAll_('Users').filter(function (x) {
    return x.Username.toLowerCase() === String(username).trim().toLowerCase() && x.Password === hash_(String(password));
  })[0];
  if (!u) return { ok: false, message: 'Username atau password salah.' };
  const token = Utilities.getUuid();
  const user = { nama: u.Nama, role: u.Role, username: u.Username };
  CacheService.getScriptCache().put('sess_' + token, JSON.stringify(user), SESSION_SECONDS);
  // data langsung dikirim bersama login supaya tidak perlu request tambahan
  return { ok: true, token: token, user: user, data: allData_(), settings: getSettings_() };
}

function logout(token) {
  CacheService.getScriptCache().remove('sess_' + token);
  return { ok: true };
}

/** Satu panggilan untuk semua data: dipakai saat buka aplikasi dan sinkron ulang. */
function bootstrap(token) {
  const user = auth_(token);
  return { user: user, data: allData_(), settings: getSettings_() };
}

/* ---------- Data ---------- */
function saveData(token, name, obj) {
  auth_(token); guard_(name);
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const sh = sheet_(name), key = UNIQUE_KEY[name];
    if (key) {
      const v = String(obj[key] || '').trim();
      if (!v) throw new Error(key + ' wajib diisi.');
      obj[key] = v;
      const dup = readAll_(name).filter(function (x) { return x[key] === v && x.ID !== obj.ID; })[0];
      if (dup) throw new Error(key + ' ' + v + ' sudah terdaftar.');
    }
    if (name === 'TahunPelajaran' && obj.Aktif === 'Aktif') {
      readAll_('TahunPelajaran').forEach(function (r) {
        if (r.ID !== obj.ID && r.Aktif === 'Aktif') sh.getRange(findRow_(sh, r.ID), 4).setValue('Tidak');
      });
    }
    const n = SHEETS[name].length;
    if (obj.ID) {
      const idx = findRow_(sh, obj.ID);
      if (idx < 0) throw new Error('Data tidak ditemukan.');
      sh.getRange(idx, 1, 1, n).setNumberFormat('@').setValues([rowOf_(name, obj)]);
    } else {
      obj.ID = uid_();
      sh.getRange(sh.getLastRow() + 1, 1, 1, n).setNumberFormat('@').setValues([rowOf_(name, obj)]);
    }
    return { ok: true, id: obj.ID };
  } finally { lock.releaseLock(); }
}

function deleteData(token, name, ids) {
  auth_(token); guard_(name);
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try { return { ok: true, removed: removeIds_(name, ids) }; }
  finally { lock.releaseLock(); }
}

/** Import massal dari Excel. Mengembalikan data terbaru agar klien tidak perlu request lagi. */
function importData(token, name, rows) {
  auth_(token);
  if (!UNIQUE_KEY[name]) throw new Error('Import tidak tersedia.');
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    const key = UNIQUE_KEY[name], exist = {};
    readAll_(name).forEach(function (x) { exist[x[key]] = true; });
    const out = []; let skip = 0;
    rows.forEach(function (r) {
      const k = String(r[key] || '').trim();
      if (!k || exist[k]) { skip++; return; }
      exist[k] = true; r.ID = uid_(); r[key] = k;
      out.push(rowOf_(name, r));
    });
    if (out.length) {
      const sh = sheet_(name);
      sh.getRange(sh.getLastRow() + 1, 1, out.length, SHEETS[name].length).setNumberFormat('@').setValues(out);
    }
    return { ok: true, added: out.length, skipped: skip, data: allData_() };
  } finally { lock.releaseLock(); }
}

/** Luluskan siswa (ids = array ID atau 'ALL'). Siswa dipindah ke sheet Alumni. */
function graduate(token, ids) {
  auth_(token);
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    const siswa = readAll_('Siswa');
    const pick = ids === 'ALL' ? siswa : siswa.filter(function (s) { return ids.indexOf(s.ID) > -1; });
    if (!pick.length) return { ok: true, count: 0, data: allData_() };
    const ta = readAll_('TahunPelajaran').filter(function (t) { return t.Aktif === 'Aktif'; })[0];
    const tahun = ta ? ta.Tahun : String(new Date().getFullYear());
    const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const rows = pick.map(function (s) {
      return rowOf_('Alumni', Object.assign({}, s, { ID: uid_(), TahunLulus: tahun, TglLulus: today }));
    });
    const al = sheet_('Alumni');
    al.getRange(al.getLastRow() + 1, 1, rows.length, SHEETS.Alumni.length).setNumberFormat('@').setValues(rows);
    removeIds_('Siswa', pick.map(function (s) { return s.ID; }));
    return { ok: true, count: pick.length, tahun: tahun, data: allData_() };
  } finally { lock.releaseLock(); }
}

/* ---------- Pengaturan sekolah (nama & logo) ---------- */
const LOGO_MAX = 45000; // batas karakter per sel Google Sheets = 50.000

function settingsSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('Settings');
  if (!sh) {
    sh = ss.insertSheet('Settings');
    sh.getRange(1, 1, 1, 2).setValues([['Key', 'Value']]).setFontWeight('bold');
    sh.getRange(2, 1, sh.getMaxRows() - 1, 2).setNumberFormat('@');
  }
  return sh;
}

function getSettings_() {
  const sh = settingsSheet_(), last = sh.getLastRow();
  const s = { NamaSekolah: '', Logo: '' };
  if (last > 1) {
    sh.getRange(2, 1, last - 1, 2).getValues().forEach(function (r) {
      if (r[0] === 'NamaSekolah' || r[0] === 'Logo') s[r[0]] = String(r[1]);
    });
  }
  return s;
}

function setSetting_(key, val) {
  const sh = settingsSheet_(), last = sh.getLastRow();
  const keys = last > 1 ? sh.getRange(2, 1, last - 1, 1).getValues() : [];
  for (let i = 0; i < keys.length; i++) {
    if (keys[i][0] === key) { sh.getRange(i + 2, 2).setNumberFormat('@').setValue(val); return; }
  }
  sh.getRange(last + 1, 1, 1, 2).setNumberFormat('@').setValues([[key, val]]);
}

/** Tanpa login: dipakai halaman login untuk menampilkan logo & nama sekolah. */
function publicInfo() {
  return getSettings_();
}

function saveSettings(token, obj) {
  const user = auth_(token);
  if (user.role !== 'admin') throw new Error('Hanya administrator yang dapat mengubah pengaturan.');
  const nama = String(obj.NamaSekolah || '').trim().slice(0, 100);
  const logo = String(obj.Logo || '');
  if (logo && !/^data:image\/(png|jpeg|webp);base64,[A-Za-z0-9+\/=]+$/.test(logo)) throw new Error('Format logo tidak valid.');
  if (logo.length > LOGO_MAX) throw new Error('Ukuran logo terlalu besar.');
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    setSetting_('NamaSekolah', nama);
    setSetting_('Logo', logo);
    return { ok: true, settings: getSettings_() };
  } finally { lock.releaseLock(); }
}
