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
 * Si ya estaba publicado y cambias este archivo:
 * Implementar > Administrar implementaciones > lápiz > Nueva versión > Implementar.
 * GET sin parámetros: siguiente ID.
 * GET ?historial=1: últimos 50 registros para la pestaña Historial.
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
  "SKU",
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
  "Supervisor que origina",
  "Defecto",
  "Reporta",
  "Autorizo",
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

function leerEncabezados(sh) {
  const lastCol = Math.max(sh.getLastColumn(), 1);
  return sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function (v) {
    return String(v || "").trim();
  });
}

function columnaSinDatos(sh, col) {
  const lastRow = Math.max(sh.getLastRow(), 1);
  const vals = sh.getRange(1, col, lastRow, 1).getValues();
  for (let i = 0; i < vals.length; i++) {
    if (String(vals[i][0] || "").trim() !== "") return false;
  }
  return true;
}

function quitarColumnasSinTitulo(sh) {
  for (let c = sh.getLastColumn(); c >= 1; c--) {
    const titulo = String(sh.getRange(1, c).getValue() || "").trim();
    if (titulo === "" && columnaSinDatos(sh, c)) {
      sh.deleteColumn(c);
    }
  }
}

function ensureSheet(ss) {
  let sh = ss.getSheetByName(HOJA);
  if (!sh) sh = ss.insertSheet(HOJA);

  quitarColumnasSinTitulo(sh);

  let actuales = leerEncabezados(sh);
  const codigoIdx = actuales.indexOf("Código");
  if (codigoIdx !== -1) {
    sh.getRange(1, codigoIdx + 1).setValue("SKU").setFontWeight("bold");
  }

  HEADERS.forEach(function (header, i) {
    actuales = leerEncabezados(sh);
    const pos = actuales.indexOf(header);
    const expected = i + 1;
    if (pos === -1) {
      if (expected <= Math.max(sh.getLastColumn(), 1)) {
        sh.insertColumnBefore(expected);
      }
      sh.getRange(1, expected).setValue(header).setFontWeight("bold");
    }
  });

  sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  sh.getRange(1, 1, 1, HEADERS.length).setFontWeight("bold");
  sh.setFrozenRows(1);
  sh.setColumnWidth(7, 180);
  sh.setColumnWidth(8, 320);

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

function celdaTexto(v) {
  if (Object.prototype.toString.call(v) === "[object Date]" && !isNaN(v.getTime())) {
    return Utilities.formatDate(v, "America/Mexico_City", "yyyy-MM-dd HH:mm:ss");
  }
  return v == null ? "" : String(v);
}

function leerUltimos(sh, limite) {
  limite = limite || 50;
  const last = sh.getLastRow();
  if (last < 2) return [];
  const num = Math.min(limite, last - 1);
  const start = last - num + 1;
  const valores = sh.getRange(start, 1, num, HEADERS.length).getValues();
  const registros = [];
  for (let i = valores.length - 1; i >= 0; i--) {
    const fila = {};
    HEADERS.forEach(function (h, c) {
      fila[h] = celdaTexto(valores[i][c]);
    });
    registros.push(fila);
  }
  return registros;
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
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
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
      data.supervisor || "",
      data.defecto || "",
      data.reporta || "",
      data.autorizo || "",
      data.se_autoriza || "",
      data.se_sanea || "",
      data.material_recuperado || "",
      data.material_rechazado || ""
    ]);
    return { ok: true, id: id, siguienteId: id + 1, sheetUrl: ss.getUrl() };
  } finally {
    lock.releaseLock();
  }
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
    const p = (e && e.parameter) || {};
    if (p.payload) {
      return jsonOut(guardarRegistro(parseBody(e)));
    }
    const ss = getSpreadsheet();
    const sh = ensureSheet(ss);
    const out = { ok: true, sheetUrl: ss.getUrl(), siguienteId: siguienteId(sh) };
    if (p.historial === "1" || p.accion === "historial") {
      out.registros = leerUltimos(sh, 50);
      out.totalHoja = Math.max(0, sh.getLastRow() - 1);
    }
    return jsonOut(out);
  } catch (err) {
    return jsonOut({ ok: false, error: String(err) });
  }
}
