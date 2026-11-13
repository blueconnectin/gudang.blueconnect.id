/** Google Apps Script API untuk aplikasi Gudang Scan. */
const SHEET_NAME = 'Pemindaian';
const HEADERS = ['ID', 'Waktu', 'Barcode', 'Nama Barang', 'Jenis ONT', 'Jumlah', 'Catatan'];
function doPost(e) {
  try { return json_(handleRequest(JSON.parse(e.postData.contents || '{}'))); }
  catch (error) { return json_({ ok: false, error: error.message || String(error) }); }
}
// Shared by HTTP clients and the Apps Script page (google.script.run).
function handleRequest(body) {
  try {
    if (!body || !['add', 'list'].includes(body.action)) throw new Error('Aksi tidak dikenal.');
    const context = getSheetContext_();
    if (body.action === 'add') return addEntry_(context.sheet, body.entry, context.headers);
    return { ok: true, entries: listEntries_(context.sheet, context.headers) };
  } catch (error) { return { ok: false, error: error.message || String(error) }; }
}
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Gudang Scan');
}
function getSheet_() { return getSheetContext_().sheet; }
function getSheetContext_() {
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!id) throw new Error('Script Property SPREADSHEET_ID belum diatur.');
  const ss = SpreadsheetApp.openById(id);
  let sheet = ss.getSheetByName(SHEET_NAME);
  let headers = sheet && sheet.getLastRow() > 0 ? sheetHeaders_(sheet) : [];
  // Normal reads do not block a save. Lock only when creating/upgrading the sheet.
  if (sheet && HEADERS.every(header => headers.includes(header))) return { sheet, headers };
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
    if (sheet.getLastRow() === 0) { sheet.appendRow(HEADERS); headers = HEADERS.slice(); }
    else {
      headers = sheetHeaders_(sheet);
      const missing = HEADERS.filter(header => !headers.includes(header));
      if (missing.length) {
        sheet.getRange(1, headers.length + 1, 1, missing.length).setValues([missing]);
        headers = headers.concat(missing);
      }
    }
    return { sheet, headers };
  } finally { lock.releaseLock(); }
}
function sheetHeaders_(sheet) {
  return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
}
function addEntry_(sheet, entry, headers) {
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
    sheet.appendRow((headers || sheetHeaders_(sheet)).map(header => values[header] ?? ''));
  } finally { lock.releaseLock(); }
  return { ok: true };
}
function listEntries_(sheet, suppliedHeaders) {
  const last = sheet.getLastRow(); if (last < 2) return [];
  const headers = suppliedHeaders || sheetHeaders_(sheet);
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
