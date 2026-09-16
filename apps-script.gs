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
  const records = rows.slice(1).filter((row) => row[0]).map((row) => Object.fromEntries(HEADERS.map((header, index) => [header, String(row[index] ?? '')])));
  return json(records);
}

function doPost(event) {
  const payload = JSON.parse(event.postData.contents);
  if (payload.action === 'login') return login(payload);
  if (payload.action !== 'save' || !isValidSession(payload.token)) return json({ ok: false, message: 'Sesión no autorizada.' });
  const record = payload.record;
  getSheet().appendRow(HEADERS.map((header) => record[header] || ''));
  return json({ ok: true });
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
