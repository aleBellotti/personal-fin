/**
 * Respaldo de Finanzas en Google Sheets.
 * Pegar este código en Extensiones > Apps Script de la planilla destino.
 * Cambiar TOKEN por una clave larga y propia (la misma se carga en la app).
 */
const TOKEN = 'CAMBIA-ESTA-CLAVE';
const CHUNK = 40000; // una celda admite 50000 caracteres

function doGet() {
  return out({ ok: true, msg: 'Respaldo activo' });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const req = JSON.parse(e.postData.contents);
    if (req.token !== TOKEN) return out({ ok: false, error: 'Clave incorrecta' });
    if (req.action === 'ping') return out({ ok: true });
    if (req.action === 'backup') return out(backup(req.data));
    if (req.action === 'restore') return out(restore());
    return out({ ok: false, error: 'Acción desconocida' });
  } catch (err) {
    return out({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function out(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function sheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function backup(data) {
  if (!data || !Array.isArray(data.expenses) || !Array.isArray(data.movements)) {
    return { ok: false, error: 'Datos inválidos' };
  }
  const ahora = new Date();

  // 1) copia exacta (para restaurar) + una copia anterior de seguridad
  const js = sheet('_json');
  const prev = sheet('_json_prev');
  const anterior = js.getLastRow() > 0 ? js.getRange(1, 1, js.getLastRow(), 1).getValues() : [];
  if (anterior.length) {
    prev.clear();
    prev.getRange(1, 1, anterior.length, 1).setValues(anterior);
  }
  const txt = JSON.stringify(data);
  const filas = [[ahora.toISOString()]];
  for (let i = 0; i < txt.length; i += CHUNK) filas.push([txt.substring(i, i + CHUNK)]);
  js.clear();
  js.getRange(1, 1, filas.length, 1).setValues(filas);
  js.hideSheet();
  prev.hideSheet();

  // 2) pestañas legibles
  writeTable('Gastos', data.expenses);
  writeTable('Movimientos', data.movements);
  writeTable('Extras', data.trabajos || []);
  writeTable('USD', (data.usd && data.usd.movimientos) || []);

  const saldos = [['Dato', 'Valor']];
  const s = data.settings || {};
  Object.keys(s).forEach(function (k) {
    if (typeof s[k] !== 'object') saldos.push([k, s[k]]);
  });
  saldos.push(['usd.objetivo', (data.usd && data.usd.objetivo) || 0]);
  saldos.push(['Último respaldo', ahora]);
  const sh = sheet('Saldos');
  sh.clear();
  sh.getRange(1, 1, saldos.length, 2).setValues(saldos);

  return { ok: true, registros: data.expenses.length + data.movements.length };
}

function writeTable(name, rows) {
  const sh = sheet(name);
  sh.clear();
  if (!rows.length) return;
  const cols = [];
  rows.forEach(function (r) {
    Object.keys(r).forEach(function (k) { if (cols.indexOf(k) < 0) cols.push(k); });
  });
  const values = [cols].concat(rows.map(function (r) {
    return cols.map(function (k) {
      const v = r[k];
      if (v === undefined || v === null) return '';
      return typeof v === 'object' ? JSON.stringify(v) : v;
    });
  }));
  sh.getRange(1, 1, values.length, cols.length).setValues(values);
  sh.setFrozenRows(1);
}

function restore() {
  const js = sheet('_json');
  const n = js.getLastRow();
  if (n < 2) return { ok: false, error: 'Todavía no hay respaldo en la planilla' };
  const filas = js.getRange(1, 1, n, 1).getValues();
  const txt = filas.slice(1).map(function (f) { return f[0]; }).join('');
  return { ok: true, fecha: String(filas[0][0]), data: JSON.parse(txt) };
}
