/** Google Apps Script API untuk aplikasi Gudang Scan. */
const SHEET_NAME = 'Pemindaian';
const HEADERS = ['ID', 'Waktu', 'Barcode', 'Nama Barang', 'Jenis ONT', 'Kategori ONT', 'Jumlah', 'Catatan', 'Keterangan Close WO'];
function doPost(e) {
  try { return json_(handleRequest(JSON.parse(e.postData.contents || '{}'))); }
  catch (error) { return json_({ ok: false, error: error.message || String(error) }); }
}
// Shared by HTTP clients and the Apps Script page (google.script.run).
function handleRequest(body) {
  try {
    if (!body || !['add', 'addMany', 'list', 'update', 'delete'].includes(body.action)) throw new Error('Aksi tidak dikenal.');
    const context = getSheetContext_();
    if (body.action === 'add') return addEntry_(context.sheet, body.entry, context.headers);
    if (body.action === 'addMany') return addManyEntries_(context.sheet, body.entries, context.headers);
    if (body.action === 'update') return updateEntry_(context.sheet, body.entry, context.headers);
    if (body.action === 'delete') return deleteEntry_(context.sheet, body.id, context.headers);
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
  const category = categoryValue_(entry);
  const closeWo = closeWoValue_(entry);
  const qty = Number(entry.qty);
  if (!Number.isInteger(qty) || qty < 1) throw new Error('Jumlah harus bilangan bulat minimal 1.');
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const values = {
      'ID': String(entry.id || Utilities.getUuid()),
      'Waktu': new Date(entry.timestamp || Date.now()),
      'Barcode': String(entry.barcode).trim(), 'Nama Barang': 'ONT',
      'Jenis ONT': type, 'Kategori ONT': category, 'Jumlah': qty, 'Catatan': String(entry.note || '').trim(), 'Keterangan Close WO': closeWo
    };
    sheet.appendRow((headers || sheetHeaders_(sheet)).map(header => values[header] ?? ''));
  } finally { lock.releaseLock(); }
  return { ok: true };
}
// Retry-safe batch endpoint: repeated IDs are acknowledged without adding another row.
function addManyEntries_(sheet, entries, headers) {
  if (!Array.isArray(entries) || !entries.length || entries.length > 20) throw new Error('Antrean tidak valid.');
  const values = entries.map(entry => {
    if (!entry || typeof entry.id !== 'string' || !entry.id || entry.id.length > 200) throw new Error('ID catatan tidak valid.');
    const barcode = String(entry.barcode || '').trim();
    if (!barcode) throw new Error('Barcode wajib diisi.');
    if (!['Huawei', 'Fiberhome'].includes(entry.type)) throw new Error('Pilih jenis ONT terlebih dahulu.');
    const category = categoryValue_(entry);
    const closeWo = closeWoValue_(entry);
    const qty = Number(entry.qty);
    if (!Number.isInteger(qty) || qty < 1) throw new Error('Jumlah harus bilangan bulat minimal 1.');
    const time = new Date(entry.timestamp);
    if (isNaN(time.getTime())) throw new Error('Waktu catatan tidak valid.');
    return {'ID': entry.id, 'Waktu': time, 'Barcode': barcode, 'Nama Barang': 'ONT',
      'Jenis ONT': entry.type, 'Kategori ONT': category, 'Jumlah': qty, 'Catatan': String(entry.note || '').trim(), 'Keterangan Close WO': closeWo};
  });
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const last = sheet.getLastRow();
    const idColumn = headers.indexOf('ID') + 1;
    const existing = new Set(last > 1 ? sheet.getRange(2, idColumn, last - 1, 1).getValues().map(row => String(row[0])) : []);
    const fresh = values.filter(value => {
      if (existing.has(value.ID)) return false;
      existing.add(value.ID);return true;
    });
    if (fresh.length) {
      const needed = last + fresh.length - sheet.getMaxRows();
      if (needed > 0) sheet.insertRowsAfter(sheet.getMaxRows(), needed);
      sheet.getRange(last + 1, 1, fresh.length, headers.length)
        .setValues(fresh.map(value => headers.map(header => value[header] ?? '')));
    }
    return {ok: true, ids: values.map(value => value.ID)};
  } finally { lock.releaseLock(); }
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
      category: String(value('Kategori ONT') || ''),
      type: String(value('Jenis ONT') || ''), qty: Number(value('Jumlah')),
      note: String(value('Catatan') || ''), closeWo: String(value('Keterangan Close WO') || '')
    };
  });
}
// Older queued records can have no category. Never guess a category for them.
function categoryValue_(entry) {
  const category = String(entry.category || '').trim();
  if (category && !['ONT Rusak', 'ONT Dismantle', 'ONT Baru'].includes(category)) throw new Error('Kategori ONT tidak valid.');
  return category;
}
function findRecordRow_(sheet, headers, id) {
  if (typeof id !== 'string' || !id || id.length > 200) throw new Error('ID catatan tidak valid.');
  const last = sheet.getLastRow();
  if (last < 2) return 0;
  const matches = sheet.getRange(2, headers.indexOf('ID') + 1, last - 1, 1).getValues()
    .map((row, index) => String(row[0]) === id ? index + 2 : 0).filter(Boolean);
  if (matches.length > 1) throw new Error('ID catatan ganda ditemukan. Periksa Google Sheets terlebih dahulu.');
  return matches[0] || 0;
}
function updateEntry_(sheet, entry, headers) {
  if (!entry || !String(entry.barcode || '').trim()) throw new Error('Barcode wajib diisi.');
  if (!['Huawei', 'Fiberhome'].includes(entry.type)) throw new Error('Pilih jenis ONT terlebih dahulu.');
  const category = categoryValue_(entry);
  const closeWo = closeWoValue_(entry);
  const qty = Number(entry.qty);
  if (!Number.isInteger(qty) || qty < 1) throw new Error('Jumlah harus bilangan bulat minimal 1.');
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const rowNumber = findRecordRow_(sheet, headers, entry.id);
    if (!rowNumber) throw new Error('Catatan tidak ditemukan. Muat ulang log.');
    const range = sheet.getRange(rowNumber, 1, 1, headers.length);
    const row = range.getValues()[0];
    const timestamp = row[headers.indexOf('Waktu')];
    const record = {id: entry.id, timestamp: new Date(timestamp).toISOString(),
      barcode: String(entry.barcode).trim(), name: 'ONT', type: entry.type,
      category, qty, closeWo, note: entry.note === undefined ? String(row[headers.indexOf('Catatan')] || '') : String(entry.note || '').trim()};
    const changes = {'Barcode': record.barcode, 'Nama Barang': record.name, 'Jenis ONT': record.type,
      'Kategori ONT': category, 'Jumlah': qty, 'Keterangan Close WO': closeWo};
    if (entry.note !== undefined) changes['Catatan'] = record.note;
    range.setValues([headers.map((header, index) => Object.prototype.hasOwnProperty.call(changes, header) ? changes[header] : row[index])]);
    return {ok: true, entry: record};
  } finally { lock.releaseLock(); }
}
function deleteEntry_(sheet, id, headers) {
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const rowNumber = findRecordRow_(sheet, headers, id);
    if (rowNumber) sheet.deleteRow(rowNumber);
    return {ok: true, id};
  } finally { lock.releaseLock(); }
}

// Legacy records/queued entries have no WO status; do not infer one from notes.
function closeWoValue_(entry) {
  const status = String(entry.closeWo || '').trim();
  if (status && !['Sudah', 'Belum'].includes(status)) throw new Error('Status Close WO tidak valid.');
  return status;
}
function json_(value) { return ContentService.createTextOutput(JSON.stringify(value)).setMimeType(ContentService.MimeType.JSON); }
