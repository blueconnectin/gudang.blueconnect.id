const SHEET_NAME = 'Pemindaian';
const HEADERS = ['ID', 'Waktu', 'Barcode', 'Nama Barang', 'Jumlah', 'Catatan'];

function doGet() {
  return json_({ ok: true, service: 'Gudang Scan API' });
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents || '{}');
    const sheet = getSheet_();

    if (body.action === 'add') {
      return json_(addEntry_(sheet, body.entry));
    }

    if (body.action === 'list') {
      return json_({ ok: true, entries: listEntries_(sheet) });
    }

    throw new Error('Aksi tidak dikenal.');
  } catch (err) {
    return json_({ ok: false, error: err.message || String(err) });
  }
}

function getSheet_() {
  const id = PropertiesService.getScriptProperties()
    .getProperty('SPREADSHEET_ID');

  if (!id) {
    throw new Error('Script Property SPREADSHEET_ID belum diatur.');
  }

  const spreadsheet = SpreadsheetApp.openById(id);
  let sheet = spreadsheet.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(SHEET_NAME);
  }

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
  }

  return sheet;
}

function addEntry_(sheet, entry) {
  if (!entry || !String(entry.barcode || '').trim()) {
    throw new Error('Barcode wajib diisi.');
  }

  const qty = Number(entry.qty);

  if (!Number.isInteger(qty) || qty < 1) {
    throw new Error('Jumlah harus bilangan bulat minimal 1.');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    sheet.appendRow([
      String(entry.id || Utilities.getUuid()),
      new Date(entry.timestamp || Date.now()),
      String(entry.barcode).trim(),
      String(entry.name || '').trim(),
      qty,
      String(entry.note || '').trim()
    ]);
  } finally {
    lock.releaseLock();
  }

  return { ok: true };
}

function listEntries_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  return sheet
    .getRange(2, 1, lastRow - 1, HEADERS.length)
    .getValues()
    .reverse()
    .map(row => ({
      id: String(row[0]),
      timestamp: row[1] instanceof Date
        ? row[1].toISOString()
        : new Date(row[1]).toISOString(),
      barcode: String(row[2]),
      name: String(row[3]),
      qty: Number(row[4]),
      note: String(row[5])
    }));
}

function json_(value) {
  return ContentService
    .createTextOutput(JSON.stringify(value))
    .setMimeType(ContentService.MimeType.JSON);
}
