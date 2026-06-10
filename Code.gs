// ── CONTROL FINANCIERO — Apps Script Backend ──────────────────────────────
// Pegá este código en Extensiones → Apps Script de tu Google Sheet
// Después publicalo como Web App (ver instrucciones abajo)

const SHEET_NAME = 'datos';

function getOrCreateSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.getRange('A1:B1').setValues([['clave', 'valor']]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// ── Lee todos los datos ────────────────────────────────────────────────────
function doGet(e) {
  try {
    const sheet = getOrCreateSheet();
    const data  = sheet.getDataRange().getValues();
    const result = {};
    for (let i = 1; i < data.length; i++) {
      const clave = data[i][0];
      const valor = data[i][1];
      if (clave) {
        try { result[clave] = JSON.parse(valor); }
        catch(err) { result[clave] = valor; }
      }
    }
    return ContentService
      .createTextOutput(JSON.stringify({ ok: true, data: result }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── Escribe/actualiza datos ────────────────────────────────────────────────
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const sheet   = getOrCreateSheet();
    const data    = sheet.getDataRange().getValues();

    // Construir índice de filas existentes
    const rowIndex = {};
    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) rowIndex[data[i][0]] = i + 1; // +1 porque sheets es 1-indexed
    }

    // Procesar cada clave del payload
    const updates = payload.updates || {};
    Object.keys(updates).forEach(clave => {
      const valor = JSON.stringify(updates[clave]);
      if (rowIndex[clave]) {
        sheet.getRange(rowIndex[clave], 2).setValue(valor);
      } else {
        sheet.appendRow([clave, valor]);
      }
    });

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
