/**
 * Finansialku — perencana keuangan pribadi
 * Backend Google Apps Script (versi 2: API untuk aplikasi mandiri).
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
// Semua permintaan wajib membawa kunci rahasia. Kunci pertama yang dipakai
// aplikasi akan dikunci di sini. Lupa kunci? Jalankan fungsi resetKunci().

function doGet() {
  return json_({ ok: true, app: APP_NAME, info: 'API Finansialku aktif. Buka aplikasi Finansialku untuk memakainya.' });
}

function doPost(e) {
  try {
    const req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const key = String(req.key || '');
    if (key.length < 12) return json_({ ok: false, auth: true, error: 'Kunci rahasia tidak valid.' });
    const props = PropertiesService.getScriptProperties();
    let saved = props.getProperty('API_KEY');
    if (!saved) {
      withLock_(function () {
        saved = props.getProperty('API_KEY');
        if (!saved) { props.setProperty('API_KEY', key); saved = key; }
      });
    }
    if (key !== saved) return json_({ ok: false, auth: true, error: 'Kunci rahasia salah. Salin kunci dari Pengaturan di HP pertama, atau jalankan fungsi resetKunci di Apps Script.' });
    const fns = { getAllData: getAllData, saveRecord: saveRecord, saveRecords: saveRecords, deleteRecord: deleteRecord, saveBudgets: saveBudgets, saveSettings: saveSettings };
    const fn = fns[req.fn];
    if (!fn) return json_({ ok: false, error: 'Fungsi tidak dikenal: ' + req.fn });
    return json_({ ok: true, data: fn.apply(null, req.args || []) });
  } catch (err) {
    return json_({ ok: false, error: String((err && err.message) || err) });
  }
}

/** Jalankan dari editor bila lupa kunci; kunci baru akan dikunci saat aplikasi terhubung lagi. */
function resetKunci() {
  PropertiesService.getScriptProperties().deleteProperty('API_KEY');
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
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
    sheetUrl: ss.getUrl()
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
    Object.keys(TABLES).forEach(function (name) { getSheet_(ss, name); });
    // hapus sheet bawaan yang kosong ("Sheet1" / "Lembar1")
    ss.getSheets().forEach(function (sh) {
      if (!TABLES[sh.getName()] && sh.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(sh);
    });
  }
  return ss;
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
