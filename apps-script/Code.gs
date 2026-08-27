/**
 * Registro de No Conforme -> Google Sheets
 *
 * Cómo publicarlo (una sola vez):
 * 1. Abre https://script.google.com
 * 2. Nuevo proyecto, pega este archivo y guarda.
 * 3. Implementar > Nueva implementación > Tipo: Aplicación web
 *    - Descripción: Registro NC
 *    - Ejecutar como: Yo
 *    - Quién tiene acceso: Cualquiera
 * 4. Copiar la URL y pegarla en index.html (SHEETS_WEBAPP_URL).
 *
 * La primera vez que llegue un registro se crea sola la hoja
 * "Registro de No Conforme" en tu Drive.
 */

const HOJA = "Registros";
const HEADERS = [
  "ID",
  "Fecha captura",
  "Fecha",
  "Mes",
  "Orden",
  "Cliente",
  "Código",
  "Producto",
  "Cantidad KG",
  "UM",
  "Cantidad Pzas/ML",
  "Unidad",
  "Proceso que detecta",
  "Máquina que detecta",
  "Operador que detecta",
  "Proceso que origina",
  "Máquina que origina",
  "Operador que origina",
  "Defecto",
  "Reporta",
  "Se autoriza",
  "Se sanea",
  "Material recuperado",
  "Material rechazado"
];

function getSpreadsheet() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty("SHEET_ID");
  if (id) {
    try {
      return SpreadsheetApp.openById(id);
    } catch (err) {
      // se recrea abajo
    }
  }
  const ss = SpreadsheetApp.create("Registro de No Conforme");
  props.setProperty("SHEET_ID", ss.getId());
  return ss;
}

function ensureSheet(ss) {
  let sh = ss.getSheetByName(HOJA);
  if (!sh) sh = ss.insertSheet(HOJA);
  const actuales = sh.getRange(1, 1, 1, HEADERS.length).getValues()[0];
  if (String(actuales[0] || "").trim() !== HEADERS[0]) {
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight("bold");
    sh.setFrozenRows(1);
    sh.setColumnWidth(7, 180);
    sh.setColumnWidth(8, 320);
  }
  const extra = ss.getSheetByName("Hoja 1") || ss.getSheetByName("Sheet1");
  if (extra && ss.getSheets().length > 1) {
    ss.deleteSheet(extra);
  }
  return sh;
}

function siguienteId(sh) {
  const last = sh.getLastRow();
  if (last < 2) return 1;
  const vals = sh.getRange(2, 1, last - 1, 1).getValues();
  let max = 0;
  vals.forEach(function (row) {
    const n = Number(row[0]);
    if (!isNaN(n) && n > max) max = n;
  });
  return max + 1;
}

function parseBody(e) {
  if (!e) return {};
  if (e.postData && e.postData.contents) {
    try {
      return JSON.parse(e.postData.contents);
    } catch (err) {
      return {};
    }
  }
  if (e.parameter && e.parameter.payload) {
    try {
      return JSON.parse(e.parameter.payload);
    } catch (err) {
      return {};
    }
  }
  return e.parameter || {};
}

function guardarRegistro(data) {
  const ss = getSpreadsheet();
  const sh = ensureSheet(ss);
  const id = siguienteId(sh);
  const captura = Utilities.formatDate(new Date(), "America/Mexico_City", "yyyy-MM-dd HH:mm:ss");
  sh.appendRow([
    id,
    captura,
    data.fecha || "",
    data.mes || "",
    data.orden || "",
    data.cliente || "",
    data.codigo || "",
    data.producto || "",
    data.cantidad || "",
    data.um || "KG",
    data.cantidad_pzas || "",
    data.um_pzas || "",
    data.proceso_detecta || "",
    data.maquina_detecta || "",
    data.operador_detecta || "",
    data.proceso_origina || "",
    data.maquina || "",
    data.operador || "",
    data.defecto || "",
    data.reporta || "",
    data.se_autoriza || "",
    data.se_sanea || "",
    data.material_recuperado || "",
    data.material_rechazado || ""
  ]);
  return { ok: true, id: id, sheetUrl: ss.getUrl() };
}

function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    return jsonOut(guardarRegistro(parseBody(e)));
  } catch (err) {
    return jsonOut({ ok: false, error: String(err) });
  }
}

function doGet(e) {
  try {
    if (e && e.parameter && e.parameter.payload) {
      return jsonOut(guardarRegistro(parseBody(e)));
    }
    const ss = getSpreadsheet();
    ensureSheet(ss);
    return jsonOut({ ok: true, sheetUrl: ss.getUrl() });
  } catch (err) {
    return jsonOut({ ok: false, error: String(err) });
  }
}
