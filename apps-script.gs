const SHEET_NAME = 'Registros';
const HEADERS = ['date', 'team', 'owner', 'safety', 'attendance', 'base', 'performance', 'quality', 'cost', 'motivation', 'message'];
const SESSION_TTL_SECONDS = 21600;

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
  const records = rows.slice(1).map((row, rowIndex) => ({ id: String(rowIndex + 2), ...Object.fromEntries(HEADERS.map((header, index) => [header, valueFor(header, row[index])])) })).filter((record) => record.date);
  return json(records);
}

function valueFor(header, value) {
  if (header === 'date' && value) return Utilities.formatDate(new Date(value), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return String(value ?? '');
}

function doPost(event) {
  const payload = JSON.parse(event.postData.contents);
  if (payload.action === 'login') return login(payload);
  if (!isValidSession(payload.token)) return json({ ok: false, message: 'Sesión no autorizada.' });
  const sheet = getSheet();
  if (payload.action === 'save') {
    sheet.appendRow(HEADERS.map((header) => payload.record[header] || ''));
    return json({ ok: true, id: String(sheet.getLastRow()) });
  }
  else if (payload.action === 'update') updateRecord(sheet, payload.id, payload.record);
  else if (payload.action === 'delete') deleteRecord(sheet, payload.id);
  else return json({ ok: false, message: 'Acción no válida.' });
  return json({ ok: true });
}

function rowNumber(id, sheet) {
  const row = Number(id);
  if (!Number.isInteger(row) || row < 2 || row > sheet.getLastRow()) throw new Error('Registro no válido.');
  return row;
}

function updateRecord(sheet, id, record) {
  sheet.getRange(rowNumber(id, sheet), 1, 1, HEADERS.length).setValues([HEADERS.map((header) => record[header] || '')]);
}

function deleteRecord(sheet, id) {
  sheet.deleteRow(rowNumber(id, sheet));
}

function login(payload) {
  const properties = PropertiesService.getScriptProperties();
  const username = properties.getProperty('ADMIN_USERNAME');
  const password = properties.getProperty('ADMIN_PASSWORD');
  if (!username || !password || payload.username !== username || payload.password !== password) return json({ ok: false, message: 'Usuario o contraseña incorrectos.' });
  const token = Utilities.getUuid();
  CacheService.getScriptCache().put(`session_${token}`, username, SESSION_TTL_SECONDS);
  return json({ ok: true, token });
}

function isValidSession(token) {
  return Boolean(token && CacheService.getScriptCache().get(`session_${token}`));
}

function json(value) {
  return ContentService.createTextOutput(JSON.stringify(value)).setMimeType(ContentService.MimeType.JSON);
}
