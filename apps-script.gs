const SHEET_NAME = 'Registros';
const HEADERS = ['date', 'team', 'owner', 'safety', 'attendance', 'base', 'performance', 'quality', 'cost', 'motivation', 'message'];

function getSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = spreadsheet.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) sheet.appendRow(HEADERS);
  return sheet;
}

function doGet(event) {
  if (event.parameter.action !== 'list') return json({ ok: true, message: 'Tablero de comunicaciones activo' });
  const sheet = getSheet();
  const rows = sheet.getDataRange().getValues();
  const records = rows.slice(1).filter((row) => row[0]).map((row) => Object.fromEntries(HEADERS.map((header, index) => [header, String(row[index] ?? '')])));
  return json(records);
}

function doPost(event) {
  const record = JSON.parse(event.postData.contents);
  getSheet().appendRow(HEADERS.map((header) => record[header] || ''));
  return json({ ok: true });
}

function json(value) {
  return ContentService.createTextOutput(JSON.stringify(value)).setMimeType(ContentService.MimeType.JSON);
}
