// Mis gastos — conexión con Google Sheets
// Pegá este código en la planilla: Extensiones > Apps Script.

const HOJA_GASTOS = "Gastos";
const HOJA_AJUSTES = "Ajustes";

function hoja_(nombre, encabezados) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(nombre);
  if (!sh) {
    sh = ss.insertSheet(nombre);
    sh.appendRow(encabezados);
    sh.setFrozenRows(1);
  }
  return sh;
}
const gastos_ = () => hoja_(HOJA_GASTOS, ["id", "fecha", "monto", "descripcion", "categoria"]);
const ajustes_ = () => hoja_(HOJA_AJUSTES, ["clave", "valor"]);

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function fecha_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), "yyyy-MM-dd");
  return String(v).slice(0, 10);
}

function leerPresupuesto_() {
  const filas = ajustes_().getDataRange().getValues();
  for (let i = 1; i < filas.length; i++) if (filas[i][0] === "presupuesto") return Number(filas[i][1]) || 0;
  return 0;
}

function doGet() {
  const filas = gastos_().getDataRange().getValues().slice(1).filter(f => f[0] !== "");
  const expenses = filas.map(f => ({ id: String(f[0]), date: fecha_(f[1]), amount: Number(f[2]) || 0, desc: String(f[3]), cat: String(f[4]) }));
  return json_({ ok: true, expenses, budget: leerPresupuesto_() });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const b = JSON.parse(e.postData.contents);
    const sh = gastos_();
    if (b.op === "add") {
      const ids = sh.getRange(1, 1, sh.getLastRow(), 1).getValues().flat().map(String);
      if (!ids.includes(String(b.e.id))) {
        sh.appendRow([String(b.e.id), b.e.date, Number(b.e.amount), b.e.desc || "", b.e.cat || "otros"]);
        sh.getRange(sh.getLastRow(), 2).setNumberFormat("yyyy-mm-dd");
      }
    } else if (b.op === "delete") {
      const ids = sh.getRange(1, 1, sh.getLastRow(), 1).getValues().flat().map(String);
      const i = ids.indexOf(String(b.id));
      if (i > 0) sh.deleteRow(i + 1);
    } else if (b.op === "budget") {
      const aj = ajustes_();
      const filas = aj.getDataRange().getValues();
      let fila = filas.findIndex(f => f[0] === "presupuesto");
      if (fila < 0) aj.appendRow(["presupuesto", Number(b.budget) || 0]);
      else aj.getRange(fila + 1, 2).setValue(Number(b.budget) || 0);
    } else {
      return json_({ ok: false, error: "operación desconocida" });
    }
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}
