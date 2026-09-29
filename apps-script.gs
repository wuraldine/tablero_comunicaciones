const SHEET_NAME = 'Registros';
const USERS_SHEET_NAME = 'Usuarios';
const HEADERS = ['date', 'team', 'owner', 'safety', 'attendance', 'base', 'performance', 'quality', 'cost', 'motivation', 'message', 'cause', 'planAccion'];
const USER_HEADERS = ['username', 'name', 'passwordHash', 'initialPassword', 'active'];
const SESSION_TTL_SECONDS = 21600;
const USER_NAMES = [
  'Erika Arango', 'Valentina Arango', 'Dayana Argumedo', 'Anyelis Avila', 'Daniela Blanco', 'Ivan Bravo',
  'Ronald Cardoza', 'Angela Chacin', 'Yeinys Contreras', 'Adriana De Avila', 'Paola Ferreira', 'Fernanda Gallego',
  'Jaiber Galvan', 'Claudia Garcia', 'Carolina Gonzalez', 'Libardo Grisales', 'Mario Hernandez', 'Geraldine Lopez',
  'Yiseth Medina', 'Danilson Mendoza', 'Alba Muñoz', 'Fanelis Narvaez', 'Sophia Pacheco', 'Margara Piedrahita',
  'Carolina Sanchez', 'Yurianys Serna', 'Rosa Serna', 'Adelmo Urariyu', 'Yeisly Valladares', 'Natalia Amariles',
  'Vergara Solanyi'
];

function getSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = spreadsheet.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) sheet.appendRow(HEADERS);
  else sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  return sheet;
}

function getUsersSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(USERS_SHEET_NAME);
  if (!sheet) sheet = spreadsheet.insertSheet(USERS_SHEET_NAME);
  if (sheet.getLastRow() === 0) sheet.appendRow(USER_HEADERS);
  else sheet.getRange(1, 1, 1, USER_HEADERS.length).setValues([USER_HEADERS]);
  return sheet;
}

function setupUsers() {
  const sheet = getUsersSheet();
  const existing = sheet.getLastRow() > 1 ? sheet.getRange(2, 1, sheet.getLastRow() - 1, USER_HEADERS.length).getValues() : [];
  const existingNames = new Set(existing.map((row) => String(row[1] || '')));
  const existingUsernames = new Set(existing.map((row) => String(row[0] || '')));
  const rows = USER_NAMES.filter((name) => !existingNames.has(name)).map((name) => {
    const username = uniqueUsername(name, existingUsernames);
    const password = numericPassword();
    existingUsernames.add(username);
    return [username, name, hashPassword(password), password, true];
  });
  if (rows.length) sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, USER_HEADERS.length).setValues(rows);
  return `Usuarios creados: ${rows.length}. Revisá la hoja ${USERS_SHEET_NAME} para entregar las credenciales iniciales.`;
}

function resetUserPasswords() {
  const sheet = getUsersSheet();
  if (sheet.getLastRow() < 2) return 'No hay usuarios para actualizar.';
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, USER_HEADERS.length).getValues();
  const used = new Set();
  const updated = rows.map((row) => {
    let password = numericPassword();
    while (used.has(password)) password = numericPassword();
    used.add(password);
    row[2] = hashPassword(password);
    row[3] = password;
    return row;
  });
  sheet.getRange(2, 1, updated.length, USER_HEADERS.length).setValues(updated);
  return `Contraseñas actualizadas: ${updated.length}.`;
}

function numericPassword() {
  return String(Math.floor(1000 + Math.random() * 9000));
}

function uniqueUsername(name, used) {
  const base = normalize(name).replace(/[^a-z0-9]+/g, '.').replace(/^\.|\.$/g, '');
  let username = base;
  let suffix = 2;
  while (used.has(username)) username = `${base}${suffix++}`;
  return username;
}

function normalize(value) {
  return String(value || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().trim();
}

function hashPassword(password) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(password), Utilities.Charset.UTF_8);
  return bytes.map((byte) => (byte < 0 ? byte + 256 : byte).toString(16).padStart(2, '0')).join('');
}

function doGet(event) {
  if (event.parameter.action !== 'list') return json({ ok: true, message: 'Tablero de comunicaciones activo' });
  const session = getSession(event.parameter.token);
  if (!session) return json({ ok: false, message: 'Sesión no autorizada.' });
  const sheet = getSheet();
  const rows = sheet.getDataRange().getValues();
  const records = rows.slice(1).map((row, rowIndex) => ({ id: String(rowIndex + 2), ...Object.fromEntries(HEADERS.map((header, index) => [header, valueFor(header, row[index])])) })).filter((record) => record.date && (session.role === 'admin' || normalize(record.owner) === normalize(session.name)));
  return json(records);
}

function valueFor(header, value) {
  if (header === 'date' && value) return Utilities.formatDate(new Date(value), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return String(value ?? '');
}

function doPost(event) {
  const payload = JSON.parse(event.postData.contents);
  if (payload.action === 'login') return login(payload);
  const session = getSession(payload.token);
  if (!session || session.role !== 'admin') return json({ ok: false, message: 'Sesión no autorizada.' });
  const sheet = getSheet();
  if (payload.action === 'save') {
    sheet.appendRow(HEADERS.map((header) => payload.record[header] || ''));
    return json({ ok: true, id: String(sheet.getLastRow()) });
  }
  else if (payload.action === 'bulkSave') {
    const records = Array.isArray(payload.records) ? payload.records : [];
    if (!records.length) return json({ ok: false, message: 'No se recibieron filas para cargar.' });
    const rows = records.map((record) => HEADERS.map((header) => record[header] || ''));
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, HEADERS.length).setValues(rows);
    return json({ ok: true, count: rows.length });
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
  let session;
  if (username && password && payload.username === username && payload.password === password) session = { role: 'admin', username, name: 'Administrador' };
  else session = findUser(payload.username, payload.password);
  if (!session) return json({ ok: false, message: 'Usuario o contraseña incorrectos.' });
  const token = Utilities.getUuid();
  CacheService.getScriptCache().put(`session_${token}`, JSON.stringify(session), SESSION_TTL_SECONDS);
  return json({ ok: true, token, role: session.role, name: session.name });
}

function findUser(username, password) {
  if (!username || !password) return null;
  const sheet = getUsersSheet();
  if (sheet.getLastRow() < 2) return null;
  const normalizedUsername = String(username).trim().toLowerCase();
  const row = sheet.getRange(2, 1, sheet.getLastRow() - 1, USER_HEADERS.length).getValues().find((item) => String(item[0]).trim().toLowerCase() === normalizedUsername && item[4] !== false && item[4] !== 'false');
  return row && hashPassword(password) === String(row[2]) ? { role: 'user', username: row[0], name: row[1] } : null;
}

function getSession(token) {
  if (!token) return null;
  const value = CacheService.getScriptCache().get(`session_${token}`);
  if (!value) return null;
  try { return JSON.parse(value); } catch (error) { return null; }
}

function json(value) {
  return ContentService.createTextOutput(JSON.stringify(value)).setMimeType(ContentService.MimeType.JSON);
}
