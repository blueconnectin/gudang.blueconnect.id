/** Google Apps Script API untuk aplikasi Gudang Scan. */
const SHEET_NAME = 'Pemindaian';
const HEADERS = ['ID', 'Waktu', 'Barcode', 'Nama Barang', 'Jenis ONT', 'Jumlah', 'Catatan'];
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
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.openById(id);
    let sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
    if (sheet.getLastRow() === 0) sheet.appendRow(HEADERS);
    else {
      const headers = sheetHeaders_(sheet);
      const missing = HEADERS.filter(header => !headers.includes(header));
      // Add missing columns at the end; existing data stays in its original columns.
      if (missing.length) sheet.getRange(1, headers.length + 1, 1, missing.length).setValues([missing]);
    }
    return sheet;
  } finally { lock.releaseLock(); }
}
function sheetHeaders_(sheet) {
  return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
}
function addEntry_(sheet, entry) {
  if (!entry || !String(entry.barcode || '').trim()) throw new Error('Barcode wajib diisi.');
  const type = String(entry.type || '').trim();
  if (!['Huawei', 'Fiberhome'].includes(type)) throw new Error('Pilih jenis ONT terlebih dahulu.');
  const qty = Number(entry.qty);
  if (!Number.isInteger(qty) || qty < 1) throw new Error('Jumlah harus bilangan bulat minimal 1.');
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const values = {
      'ID': String(entry.id || Utilities.getUuid()),
      'Waktu': new Date(entry.timestamp || Date.now()),
      'Barcode': String(entry.barcode).trim(), 'Nama Barang': 'ONT',
      'Jenis ONT': type, 'Jumlah': qty, 'Catatan': String(entry.note || '').trim()
    };
    sheet.appendRow(sheetHeaders_(sheet).map(header => values[header] ?? ''));
  } finally { lock.releaseLock(); }
  return { ok: true };
}
function listEntries_(sheet) {
  const last = sheet.getLastRow(); if (last < 2) return [];
  const headers = sheetHeaders_(sheet);
  return sheet.getRange(2, 1, last - 1, headers.length).getValues().reverse().map(row => {
    const value = header => row[headers.indexOf(header)];
    const time = value('Waktu');
    return {
      id: String(value('ID') || ''),
      timestamp: time instanceof Date ? time.toISOString() : new Date(time).toISOString(),
      barcode: String(value('Barcode') || ''), name: String(value('Nama Barang') || ''),
      type: String(value('Jenis ONT') || ''), qty: Number(value('Jumlah')),
      note: String(value('Catatan') || '')
    };
  });
}
function json_(value) { return ContentService.createTextOutput(JSON.stringify(value)).setMimeType(ContentService.MimeType.JSON); }
