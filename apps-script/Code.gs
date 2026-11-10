// Caja Bufet - Apps Script (ventas + catálogo compartidos entre dispositivos)
//
// Hoja de ventas: la misma de siempre (hoja activa), columnas A-E:
//   Fecha | Hora | Producto | Monto | Medio de pago   (fila 1 = encabezados)
// Catálogo: se guarda como JSON en la celda A1 de una hoja llamada "Catalogo"
//   (se crea sola la primera vez).

var CATALOG_SHEET = 'Catalogo';

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function salesSheet_() {
  return SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
}

function catalogSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(CATALOG_SHEET) || ss.insertSheet(CATALOG_SHEET);
}

// Fecha y hora se guardan como TEXTO para que Sheets no las transforme.
function writeSalesRows_(sheet, startRow, rows) {
  if (!rows.length) return;
  sheet.getRange(startRow, 1, rows.length, 2).setNumberFormat('@');
  sheet.getRange(startRow, 1, rows.length, 5).setValues(rows);
}

// LECTURA: devuelve ventas + catálogo
function doGet(e) {
  try {
    var sheet = salesSheet_();
    var sales = [];
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      var range = sheet.getRange(2, 1, lastRow - 1, 5);
      var values = range.getValues();
      var shown = range.getDisplayValues();
      for (var i = 0; i < values.length; i++) {
        if (shown[i][2] === '' && shown[i][3] === '') continue; // fila vacía
        sales.push({
          date: shown[i][0],
          time: shown[i][1],
          product: shown[i][2],
          amount: Number(values[i][3]) || 0,
          type: shown[i][4]
        });
      }
    }

    var catalog = null;
    var raw = catalogSheet_().getRange(1, 1).getValue();
    if (raw) {
      try { catalog = JSON.parse(raw); } catch (err) { catalog = null; }
    }

    return json_({ status: 'success', sales: sales, catalog: catalog });
  } catch (err) {
    return json_({ status: 'error', message: err.toString() });
  }
}

// ESCRITURA: venta nueva, reemplazo total de ventas o guardado del catálogo
function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var data = JSON.parse(e.postData.contents);

    if (data.action === 'saveCatalog') {
      catalogSheet_().getRange(1, 1).setValue(JSON.stringify(data.catalog));
    } else if (data.action === 'syncAll') {
      var sheet = salesSheet_();
      var lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        sheet.getRange(2, 1, lastRow - 1, 5).clearContent();
      }
      var rows = (data.sales || []).map(function (s) {
        return [s.date, s.time, s.product, s.amount, s.type];
      });
      writeSalesRows_(sheet, 2, rows);
    } else {
      var sh = salesSheet_();
      writeSalesRows_(sh, sh.getLastRow() + 1,
        [[data.date, data.time, data.product, data.amount, data.type]]);
    }

    return json_({ status: 'success' });
  } catch (err) {
    return json_({ status: 'error', message: err.toString() });
  } finally {
    try { lock.releaseLock(); } catch (e2) { }
  }
}
