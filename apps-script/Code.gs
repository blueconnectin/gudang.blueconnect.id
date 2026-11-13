/** Google Apps Script API untuk aplikasi Gudang Scan. */
const SHEET_NAME = 'Pemindaian';
const HEADERS = ['ID', 'Waktu', 'Barcode', 'Nama Barang', 'Jumlah', 'Catatan'];
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents || '{}');
    const sheet = getSheet_();
    let result;
    if (body.action === 'add') result = addEntry_(sheet, body.entry);
    else if (body.action === 'list') result = { ok: true, entries: listEntries_(sheet) };
    else throw new Error('Aksi tidak dikenal.');
    return json_(result);
  } catch (err) { return json_({ ok: false, error: err.message || String(err) }); }
}
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Gudang Scan');
}
function getSheet_() {
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!id) throw new Error('Script Property SPREADSHEET_ID belum diatur.');
  const ss = SpreadsheetApp.openById(id);
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) sheet.appendRow(HEADERS);
  return sheet;
}
function addEntry_(sheet, entry) {
  if (!entry || !String(entry.barcode || '').trim()) throw new Error('Barcode wajib diisi.');
  const qty = Number(entry.qty);
  if (!Number.isInteger(qty) || qty < 1) throw new Error('Jumlah harus bilangan bulat minimal 1.');
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    sheet.appendRow([String(entry.id || Utilities.getUuid()), new Date(entry.timestamp || Date.now()), String(entry.barcode).trim(), String(entry.name || '').trim(), qty, String(entry.note || '').trim()]);
  } finally { lock.releaseLock(); }
  return { ok: true };
}
function listEntries_(sheet) {
  const last = sheet.getLastRow(); if (last < 2) return [];
  return sheet.getRange(2, 1, last - 1, HEADERS.length).getValues().reverse().map(r => ({ id: String(r[0]), timestamp: r[1] instanceof Date ? r[1].toISOString() : new Date(r[1]).toISOString(), barcode: String(r[2]), name: String(r[3]), qty: Number(r[4]), note: String(r[5]) }));
}
function json_(value) { return ContentService.createTextOutput(JSON.stringify(value)).setMimeType(ContentService.MimeType.JSON); }
