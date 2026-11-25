// ====== GANTI PIN INI dengan PIN pilihan Anda (angka/huruf) ======
const PIN = '1234';
// =================================================================

const KATEGORI = ['BPJS', 'Asuransi', 'Umum', 'Baksos'];
const HEAD_REKAP = ['id', 'waktu_ms', 'BPJS', 'Asuransi', 'Umum', 'Baksos'];

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Hitung Otomatis Pasien Bokingan')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function sheet_(name, head) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(name);
  if (!s) { s = ss.insertSheet(name); s.appendRow(head); }
  return s;
}

function cek_(pin) {
  if (String(pin) !== String(PIN)) throw new Error('PIN salah');
}

function state_() {
  const rows = sheet_('Rekap', HEAD_REKAP).getDataRange().getValues().slice(1);
  const batches = rows.filter(r => r[0] !== '').map(r => ({
    id: String(r[0]), t: Number(r[1]),
    c: { BPJS: +r[2] || 0, Asuransi: +r[3] || 0, Umum: +r[4] || 0, Baksos: +r[5] || 0 }
  }));
  let resetT = 0;
  sheet_('Meta', ['key', 'value']).getDataRange().getValues().slice(1).forEach(r => {
    if (r[0] === 'resetT') resetT = Number(r[1]) || 0;
  });
  return { batches: batches, resetT: resetT };
}

function getState(pin) { cek_(pin); return state_(); }

function addBatch(pin, c) {
  cek_(pin);
  const row = KATEGORI.map(k => {
    const n = Math.floor(Number(c && c[k]) || 0);
    if (n < 0 || n > 5000) throw new Error('Jumlah tidak valid');
    return n;
  });
  if (row.every(n => n === 0)) throw new Error('Tidak ada data');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    sheet_('Rekap', HEAD_REKAP).appendRow([Utilities.getUuid(), Date.now()].concat(row));
    SpreadsheetApp.flush();
  } finally { lock.releaseLock(); }
  return state_();
}

function resetCounts(pin) {
  cek_(pin);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const m = sheet_('Meta', ['key', 'value']);
    const v = m.getDataRange().getValues();
    let done = false;
    for (let i = 1; i < v.length; i++) {
      if (v[i][0] === 'resetT') { m.getRange(i + 1, 2).setValue(Date.now()); done = true; break; }
    }
    if (!done) m.appendRow(['resetT', Date.now()]);
    SpreadsheetApp.flush();
  } finally { lock.releaseLock(); }
  return state_();
}

function clearAll(pin) {
  cek_(pin);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const s = sheet_('Rekap', HEAD_REKAP);
    if (s.getLastRow() > 1) s.deleteRows(2, s.getLastRow() - 1);
    SpreadsheetApp.flush();
  } finally { lock.releaseLock(); }
  return state_();
}
