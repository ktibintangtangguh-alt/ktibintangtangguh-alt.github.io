/**
 * Finansialku — perencana keuangan pribadi
 * Backend Google Apps Script (versi 4: banyak pengguna, daftar dengan username & password).
 *
 * Semua data disimpan di Google Sheet "Finansialku — Data" di Google Drive kamu.
 * Sheet dibuat otomatis saat aplikasi pertama kali dibuka.
 * Jika skrip ini dibuat dari dalam sebuah Google Sheet (Ekstensi > Apps Script),
 * sheet itu yang dipakai sebagai tempat data.
 */

const APP_NAME = 'Finansialku';

// Kolom setiap tabel (urutan = urutan kolom di Google Sheet)
const TABLES = {
  Transaksi: ['id', 'tanggal', 'jenis', 'kategori', 'jumlah', 'akun', 'catatan', 'dibuat'],
  Anggaran: ['kategori', 'jumlah'],
  Investasi: ['id', 'nama', 'jenis', 'modal', 'nilai', 'diperbarui', 'catatan'],
  Target: ['id', 'nama', 'target', 'terkumpul', 'tenggat', 'catatan'],
  Pengaturan: ['kunci', 'isi']
};
const NUMERIC = { jumlah: true, modal: true, nilai: true, target: true, terkumpul: true };

const DEFAULT_SETTINGS = {
  kategoriPengeluaran: [
    'Makan & Minum', 'Belanja Harian', 'Transportasi', 'Tagihan & Utilitas', 'Tempat Tinggal',
    'Cicilan & Utang', 'Kesehatan', 'Pendidikan', 'Keluarga', 'Hiburan', 'Belanja Pribadi',
    'Sedekah & Donasi', 'Tabungan & Investasi', 'Lainnya'
  ],
  kategoriPemasukan: ['Gaji', 'Bonus', 'Usaha / Sampingan', 'Hasil Investasi', 'Hadiah', 'Saldo Awal', 'Lainnya'],
  akun: ['Tunai', 'Rekening Bank', 'E-Wallet'],
  jenisInvestasi: ['Reksa Dana', 'Saham', 'Emas', 'Deposito', 'Obligasi / SBN', 'Kripto', 'Properti', 'Lainnya'],
  targetPemasukan: 0
};

/* ---------- API untuk aplikasi Finansialku (PWA) ---------- */
// Deploy sebagai Aplikasi web: Jalankan sebagai "Saya", akses "Siapa saja".
// Versi 3: banyak pengguna. Kunci pertama yang terhubung menjadi kunci ADMIN.
// Admin membuat kode undangan untuk pengguna lain; setiap pengguna mendapat
// Google Sheet sendiri di Drive admin. Daftar pengguna ada di tab "Pengguna"
// pada sheet milik admin. Lupa kunci admin? Jalankan resetKunci().

let CTX = null; // { role: 'admin' | 'user', name, kode, ss }

function doGet() {
  return json_({ ok: true, app: APP_NAME, info: 'API Finansialku aktif. Buka aplikasi Finansialku untuk memakainya.' });
}

function doPost(e) {
  try {
    const req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const publicFns = { register: register, login: login };
    if (publicFns[req.fn]) return json_({ ok: true, data: publicFns[req.fn].apply(null, req.args || []) });
    const key = String(req.key || '');
    if (key.length < 12) return json_({ ok: false, auth: true, error: 'Kode akses tidak valid.' });
    const ctx = auth_(key);
    if (!ctx) return json_({ ok: false, auth: true, error: 'Kode akses salah atau sudah dinonaktifkan. Minta kode baru ke admin.' });
    CTX = ctx;
    const userFns = { getAllData: getAllData, saveRecord: saveRecord, saveRecords: saveRecords, deleteRecord: deleteRecord, saveBudgets: saveBudgets, saveSettings: saveSettings, changePassword: changePassword };
    const adminFns = { listUsers: listUsers, addUser: addUser, setUserActive: setUserActive, resetUserCode: resetUserCode, renameUser: renameUser, userSummary: userSummary, approveUser: approveUser, rejectUser: rejectUser, resetPassword: resetPassword };
    const fn = userFns[req.fn] || (ctx.role === 'admin' ? adminFns[req.fn] : null);
    if (!fn) return json_({ ok: false, error: 'Fungsi tidak dikenal atau khusus admin: ' + req.fn });
    return json_({ ok: true, data: fn.apply(null, req.args || []) });
  } catch (err) {
    return json_({ ok: false, error: String((err && err.message) || err) });
  }
}

function auth_(key) {
  const props = PropertiesService.getScriptProperties();
  let admin = props.getProperty('API_KEY');
  if (!admin) {
    withLock_(function () {
      admin = props.getProperty('API_KEY');
      if (!admin) { props.setProperty('API_KEY', key); admin = key; }
    });
  }
  if (key === admin) return { role: 'admin', name: 'Admin', kode: '', ss: adminDb_() };
  const u = findUser_(key);
  if (!u || !u.aktif || !u.sheetId) return null;
  let ss;
  try { ss = SpreadsheetApp.openById(u.sheetId); } catch (e) { throw new Error('Sheet milik ' + u.nama + ' tidak ditemukan. Minta admin memeriksanya.'); }
  return { role: 'user', name: u.nama, kode: u.kode, ss: ss };
}

/** Jalankan dari editor bila lupa kunci admin; perangkat admin berikutnya yang terhubung menjadi admin. */
function resetKunci() {
  PropertiesService.getScriptProperties().deleteProperty('API_KEY');
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- Pengguna (khusus admin) ---------- */

const USER_COLS = ['kode', 'nama', 'sheetId', 'aktif', 'dibuat', 'username', 'hash', 'salt', 'gagal', 'kunciSampai'];

function usersSheet_() {
  const ss = adminDb_();
  let sh = ss.getSheetByName('Pengguna');
  if (sh) {
    if (sh.getLastColumn() < USER_COLS.length) { // tingkatkan tabel lama (versi 3)
      sh.getRange(1, 1, 1, USER_COLS.length).setValues([USER_COLS]).setFontWeight('bold').setBackground('#0e6b57').setFontColor('#ffffff');
      sh.getRange(2, 1, sh.getMaxRows() - 1, USER_COLS.length).setNumberFormat('@');
    }
    return sh;
  }
  sh = ss.insertSheet('Pengguna');
  sh.getRange(1, 1, 1, USER_COLS.length).setValues([USER_COLS])
    .setFontWeight('bold').setBackground('#0e6b57').setFontColor('#ffffff');
  sh.setFrozenRows(1);
  sh.getRange(2, 1, sh.getMaxRows() - 1, USER_COLS.length).setNumberFormat('@');
  return sh;
}

function readUsers_() {
  const sh = usersSheet_();
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, USER_COLS.length).getValues()
    .map(function (r, i) {
      const st = String(r[3]) === 'tidak' ? 'nonaktif' : String(r[3]) === 'menunggu' ? 'menunggu' : 'aktif';
      return { row: i + 2, kode: String(r[0]), nama: String(r[1]), sheetId: String(r[2]), status: st, aktif: st === 'aktif', dibuat: String(r[4]),
        username: String(r[5] || ''), hash: String(r[6] || ''), salt: String(r[7] || ''), gagal: Number(r[8]) || 0, kunciSampai: Number(r[9]) || 0 };
    })
    .filter(function (u) { return u.kode; });
}

function findUser_(kode) {
  return readUsers_().filter(function (u) { return u.kode === kode; })[0] || null;
}

function publicUser_(u) {
  return { kode: u.kode, nama: u.nama, aktif: u.aktif, status: u.status, username: u.username, dibuat: u.dibuat, sheetUrl: u.sheetId ? 'https://docs.google.com/spreadsheets/d/' + u.sheetId + '/edit' : '' };
}

function genCode_() {
  return (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '').slice(0, 16);
}

function listUsers() {
  return readUsers_().map(publicUser_);
}

function addUser(nama) {
  nama = String(nama || '').trim().slice(0, 40);
  if (!nama) throw new Error('Isi nama pengguna.');
  return withLock_(function () {
    if (readUsers_().some(function (u) { return u.nama.toLowerCase() === nama.toLowerCase(); })) throw new Error('Nama "' + nama + '" sudah dipakai.');
    const ss = SpreadsheetApp.create(APP_NAME + ' — ' + nama);
    initSpreadsheet_(ss);
    const u = { kode: genCode_(), nama: nama, sheetId: ss.getId(), aktif: true, dibuat: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd') };
    usersSheet_().appendRow([u.kode, u.nama, u.sheetId, 'ya', u.dibuat, '', '', '', 0, 0]);
    return publicUser_(u);
  });
}

function setUserActive(kode, aktif) {
  return withLock_(function () {
    const u = findUser_(kode); if (!u) throw new Error('Pengguna tidak ditemukan.');
    if (!u.sheetId) throw new Error('Pendaftaran ini belum disetujui.');
    usersSheet_().getRange(u.row, 4).setValue(aktif ? 'ya' : 'tidak');
    u.status = aktif ? 'aktif' : 'nonaktif';
    u.aktif = !!aktif; return publicUser_(u);
  });
}

function resetUserCode(kode) {
  return withLock_(function () {
    const u = findUser_(kode); if (!u) throw new Error('Pengguna tidak ditemukan.');
    u.kode = genCode_();
    usersSheet_().getRange(u.row, 1).setValue(u.kode);
    return publicUser_(u);
  });
}

function renameUser(kode, nama) {
  nama = String(nama || '').trim().slice(0, 40);
  if (!nama) throw new Error('Isi nama pengguna.');
  return withLock_(function () {
    const u = findUser_(kode); if (!u) throw new Error('Pengguna tidak ditemukan.');
    usersSheet_().getRange(u.row, 2).setValue(nama);
    try { SpreadsheetApp.openById(u.sheetId).rename(APP_NAME + ' — ' + nama); } catch (e) { /* sheet tetap */ }
    u.nama = nama; return publicUser_(u);
  });
}

/** Ringkasan satu bulan (yyyy-MM) untuk satu pengguna. */
function userSummary(kode, month) {
  const u = findUser_(kode); if (!u) throw new Error('Pengguna tidak ditemukan.');
  if (!u.sheetId) return { kode: kode, inc: 0, exp: 0, n: 0, last: '' };
  const tx = readTable_(SpreadsheetApp.openById(u.sheetId), 'Transaksi');
  let inc = 0, exp = 0, n = 0, last = '';
  tx.forEach(function (t) {
    if (t.tanggal > last) last = t.tanggal;
    if (String(t.tanggal).slice(0, 7) !== month) return;
    n++;
    if (t.jenis === 'Pemasukan') { if (t.kategori !== 'Saldo Awal') inc += t.jumlah; } else exp += t.jumlah;
  });
  return { kode: kode, inc: inc, exp: exp, n: n, last: last };
}

/* ---------- Fungsi yang dipanggil aplikasi ---------- */

function getAllData() {
  const ss = getDb_();
  const settings = Object.assign({}, DEFAULT_SETTINGS);
  readTable_(ss, 'Pengaturan').forEach(function (r) {
    try { settings[r.kunci] = JSON.parse(r.isi); } catch (e) { /* abaikan isi rusak */ }
  });
  return {
    transaksi: readTable_(ss, 'Transaksi'),
    anggaran: readTable_(ss, 'Anggaran'),
    investasi: readTable_(ss, 'Investasi'),
    target: readTable_(ss, 'Target'),
    settings: settings,
    sheetUrl: ss.getUrl(),
    role: CTX ? CTX.role : 'admin',
    userName: CTX ? CTX.name : 'Admin'
  };
}

/** Tambah atau ubah satu baris (berdasarkan id). */
function saveRecord(table, rec) {
  if (['Transaksi', 'Investasi', 'Target'].indexOf(table) < 0) throw new Error('Tabel tidak dikenal: ' + table);
  if (!rec || typeof rec !== 'object') throw new Error('Data kosong');
  return withLock_(function () {
    const ss = getDb_();
    const sh = getSheet_(ss, table);
    if (!rec.id) rec.id = Utilities.getUuid().slice(0, 12);
    const row = toRow_(table, rec);
    const r = findRow_(sh, rec.id);
    if (r) sh.getRange(r, 1, 1, row.length).setValues([row]);
    else sh.appendRow(row);
    return rec;
  });
}

/** Simpan beberapa baris sekaligus (dipakai saat setoran investasi/target juga dicatat sebagai transaksi). */
function saveRecords(items) {
  (items || []).forEach(function (it) { saveRecord(it.table, it.rec); });
  return true;
}

function deleteRecord(table, id) {
  if (['Transaksi', 'Investasi', 'Target'].indexOf(table) < 0) throw new Error('Tabel tidak dikenal: ' + table);
  return withLock_(function () {
    const sh = getSheet_(getDb_(), table);
    const r = findRow_(sh, id);
    if (r) sh.deleteRow(r);
    return !!r;
  });
}

/** Ganti seluruh anggaran bulanan: [{kategori, jumlah}] */
function saveBudgets(list) {
  return withLock_(function () {
    const sh = getSheet_(getDb_(), 'Anggaran');
    const last = sh.getLastRow();
    if (last > 1) sh.getRange(2, 1, last - 1, 2).clearContent();
    const rows = (list || [])
      .filter(function (b) { return b && b.kategori && Number(b.jumlah) > 0; })
      .map(function (b) { return toRow_('Anggaran', b); });
    if (rows.length) sh.getRange(2, 1, rows.length, 2).setValues(rows);
    return true;
  });
}

/** Simpan pengaturan (kategori, akun, target pemasukan, dll). */
function saveSettings(obj) {
  return withLock_(function () {
    const sh = getSheet_(getDb_(), 'Pengaturan');
    const keys = Object.keys(obj || {});
    keys.forEach(function (k) {
      const row = [k, JSON.stringify(obj[k])];
      const r = findRow_(sh, k);
      if (r) sh.getRange(r, 1, 1, 2).setValues([row]);
      else sh.appendRow(row);
    });
    return true;
  });
}

/* ---------- Penyimpanan ---------- */

function getDb_() {
  return CTX ? CTX.ss : adminDb_();
}

/** Sheet milik admin (sheet data pertama). */
function adminDb_() {
  let ss = null;
  try { ss = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) { ss = null; }
  if (ss) return ss;

  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('SHEET_ID');
  if (id) {
    try { ss = SpreadsheetApp.openById(id); } catch (e) { ss = null; }
  }
  if (!ss) {
    ss = SpreadsheetApp.create(APP_NAME + ' — Data');
    props.setProperty('SHEET_ID', ss.getId());
    initSpreadsheet_(ss);
  }
  return ss;
}

function initSpreadsheet_(ss) {
  Object.keys(TABLES).forEach(function (name) { getSheet_(ss, name); });
  // hapus sheet bawaan yang kosong ("Sheet1" / "Lembar1")
  ss.getSheets().forEach(function (sh) {
    if (!TABLES[sh.getName()] && sh.getName() !== 'Pengguna' && sh.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(sh);
  });
}

function getSheet_(ss, name) {
  let sh = ss.getSheetByName(name);
  if (sh) return sh;
  const cols = TABLES[name];
  sh = ss.insertSheet(name);
  sh.getRange(1, 1, 1, cols.length).setValues([cols])
    .setFontWeight('bold').setBackground('#0e6b57').setFontColor('#ffffff');
  sh.setFrozenRows(1);
  cols.forEach(function (c, i) {
    const col = sh.getRange(2, i + 1, sh.getMaxRows() - 1, 1);
    col.setNumberFormat(NUMERIC[c] ? '#,##0' : '@');
  });
  sh.autoResizeColumns(1, cols.length);
  return sh;
}

function readTable_(ss, name) {
  const sh = getSheet_(ss, name);
  const last = sh.getLastRow();
  if (last < 2) return [];
  const cols = TABLES[name];
  const tz = ss.getSpreadsheetTimeZone() || Session.getScriptTimeZone();
  return sh.getRange(2, 1, last - 1, cols.length).getValues()
    .filter(function (r) { return r[0] !== '' && r[0] !== null; })
    .map(function (r) {
      const o = {};
      cols.forEach(function (c, i) {
        let v = r[i];
        if (v instanceof Date) v = Utilities.formatDate(v, tz, 'yyyy-MM-dd');
        if (NUMERIC[c]) { o[c] = Number(v) || 0; return; }
        v = v === null || v === undefined ? '' : String(v);
        if (/^'[=+\-@]/.test(v)) v = v.slice(1);
        o[c] = v;
      });
      return o;
    });
}

function toRow_(table, rec) {
  return TABLES[table].map(function (c) {
    if (NUMERIC[c]) return Number(rec[c]) || 0;
    let v = rec[c] === null || rec[c] === undefined ? '' : String(rec[c]);
    if (/^[=+\-@]/.test(v)) v = "'" + v; // jangan dianggap rumus
    return v;
  });
}

function findRow_(sh, id) {
  const last = sh.getLastRow();
  if (last < 2 || id === undefined || id === null || id === '') return 0;
  const cell = sh.getRange(2, 1, last - 1, 1)
    .createTextFinder(String(id)).matchEntireCell(true).matchCase(true).findNext();
  return cell ? cell.getRow() : 0;
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

/* ---------- Akun dengan username & password ---------- */
// Siapa pun yang punya link pendaftaran bisa mendaftar; akun baru berstatus
// "menunggu" sampai admin menyetujuinya. Password disimpan sebagai hash bergaram.

const MAX_PENDING = 30;

function hash_(password, salt) {
  let h = salt;
  for (let i = 0; i < 300; i++) {
    h = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, h + '|' + password, Utilities.Charset.UTF_8));
  }
  return h;
}

function normUser_(x) { return String(x || '').trim().toLowerCase(); }

function checkPassword_(p) {
  p = String(p || '');
  if (p.length < 8) throw new Error('Password minimal 8 karakter.');
  if (p.length > 100) throw new Error('Password terlalu panjang.');
  return p;
}

function register(nama, username, password) {
  nama = String(nama || '').trim().slice(0, 40);
  username = normUser_(username);
  if (!nama) throw new Error('Isi nama kamu.');
  if (!/^[a-z0-9._]{3,20}$/.test(username)) throw new Error('Username 3–20 karakter: huruf kecil, angka, titik, atau garis bawah.');
  password = checkPassword_(password);
  return withLock_(function () {
    const users = readUsers_();
    if (users.some(function (u) { return u.username === username; })) throw new Error('Username "' + username + '" sudah dipakai. Pilih yang lain.');
    if (users.filter(function (u) { return u.status === 'menunggu'; }).length >= MAX_PENDING) throw new Error('Pendaftaran sedang penuh. Hubungi admin.');
    const salt = genCode_();
    usersSheet_().appendRow([genCode_(), nama, '', 'menunggu', Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'), username, hash_(password, salt), salt, 0, 0]);
    return { status: 'menunggu' };
  });
}

function login(username, password) {
  username = normUser_(username); password = String(password || '');
  const generic = 'Username atau password salah.';
  return withLock_(function () {
    const u = readUsers_().filter(function (x) { return x.username && x.username === username; })[0];
    if (!u) { Utilities.sleep(400); throw new Error(generic); }
    const now = Date.now(), sh = usersSheet_();
    if (u.kunciSampai > now) throw new Error('Terlalu banyak percobaan salah. Coba lagi dalam ' + Math.ceil((u.kunciSampai - now) / 60000) + ' menit.');
    if (hash_(password, u.salt) !== u.hash) {
      const g = u.gagal + 1;
      if (g >= 5) sh.getRange(u.row, 9, 1, 2).setValues([[0, now + 15 * 60000]]);
      else sh.getRange(u.row, 9).setValue(g);
      throw new Error(g >= 5 ? 'Terlalu banyak percobaan salah. Akun dikunci 15 menit.' : generic);
    }
    if (u.gagal) sh.getRange(u.row, 9, 1, 2).setValues([[0, 0]]);
    if (u.status === 'menunggu') throw new Error('Akunmu masih menunggu persetujuan admin. Coba masuk lagi nanti.');
    if (u.status === 'nonaktif') throw new Error('Akunmu dinonaktifkan admin.');
    return { kode: u.kode, nama: u.nama, username: u.username };
  });
}

function changePassword(oldPw, newPw) {
  if (!CTX || CTX.role !== 'user') throw new Error('Hanya untuk akun pengguna.');
  newPw = checkPassword_(newPw);
  return withLock_(function () {
    const u = findUser_(CTX.kode);
    if (!u || !u.username) throw new Error('Akun ini masuk lewat link undangan dan tidak memakai password.');
    if (hash_(String(oldPw || ''), u.salt) !== u.hash) throw new Error('Password lama salah.');
    const salt = genCode_();
    usersSheet_().getRange(u.row, 7, 1, 2).setValues([[hash_(newPw, salt), salt]]);
    return true;
  });
}

function approveUser(kode) {
  return withLock_(function () {
    const u = findUser_(kode); if (!u) throw new Error('Pengguna tidak ditemukan.');
    const sh = usersSheet_();
    if (!u.sheetId) {
      const ss = SpreadsheetApp.create(APP_NAME + ' — ' + u.nama);
      initSpreadsheet_(ss);
      u.sheetId = ss.getId();
      sh.getRange(u.row, 3).setValue(u.sheetId);
    }
    sh.getRange(u.row, 4).setValue('ya');
    u.status = 'aktif'; u.aktif = true;
    return publicUser_(u);
  });
}

function rejectUser(kode) {
  return withLock_(function () {
    const u = findUser_(kode); if (!u) throw new Error('Pengguna tidak ditemukan.');
    if (u.status !== 'menunggu') throw new Error('Hanya pendaftaran yang menunggu yang bisa ditolak. Gunakan Nonaktifkan.');
    usersSheet_().deleteRow(u.row);
    return true;
  });
}

/** Admin mengatur password baru; pengguna keluar dari semua HP dan harus masuk lagi. */
function resetPassword(kode, newPw) {
  newPw = checkPassword_(newPw);
  return withLock_(function () {
    const u = findUser_(kode); if (!u) throw new Error('Pengguna tidak ditemukan.');
    if (!u.username) throw new Error('Pengguna ini masuk lewat link undangan. Gunakan Ganti kode akses.');
    const salt = genCode_(); u.kode = genCode_();
    usersSheet_().getRange(u.row, 1).setValue(u.kode);
    usersSheet_().getRange(u.row, 7, 1, 4).setValues([[hash_(newPw, salt), salt, 0, 0]]);
    return publicUser_(u);
  });
}

