// Google Apps Script — Web App para conteo de leads TikTok
// leads-tiktok-equilibrio (Contact Center Equilibrio Total)

const SHEET_NAME = 'Leads';
const ZONA_HORARIA = 'America/Mexico_City';

function doGet(e) {
  const accion = e.parameter.action;
  if (accion === 'resumen') {
    return respuestaJSON(getResumen());
  }
  return respuestaJSON({ ok: true, mensaje: 'API de Leads TikTok Equilibrio Total activa' });
}

function doPost(e) {
  try {
    const p = e.parameter;
    const fecha = p.fecha;
    const leadsDia = parseInt(p.leadsDia, 10) || 0;
    const leadsAcumulados = parseInt(p.leadsAcumulados, 10) || 0;

    if (!fecha) {
      return respuestaJSON({ ok: false, error: 'Falta la fecha' });
    }

    guardarRegistro(fecha, leadsDia, leadsAcumulados);

    return respuestaJSON({ ok: true });
  } catch (err) {
    return respuestaJSON({ ok: false, error: err.toString() });
  }
}

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  const headers = ['Fecha', 'Leads del día', 'Leads acumulados', 'Total', 'Actualizado'];
  const primeraFila = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  const yaTieneHeaders = headers.every(function (h, i) { return primeraFila[i] === h; });
  if (!yaTieneHeaders) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    // Columna A como texto plano para que "2026-09-04" nunca se reinterprete como Date.
    sheet.getRange(1, 1, sheet.getMaxRows(), 1).setNumberFormat('@');
  }
  return sheet;
}

function normalizarFecha(valor) {
  if (valor instanceof Date) {
    return Utilities.formatDate(valor, ZONA_HORARIA, 'yyyy-MM-dd');
  }
  return valor;
}

function guardarRegistro(fechaStr, leadsDia, leadsAcumulados) {
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  const total = leadsDia + leadsAcumulados;
  const ahora = new Date();

  for (let i = 1; i < data.length; i++) {
    const filaFechaStr = normalizarFecha(data[i][0]);
    if (filaFechaStr === fechaStr) {
      sheet.getRange(i + 1, 2, 1, 3).setValues([[leadsDia, leadsAcumulados, total]]);
      sheet.getRange(i + 1, 5).setValue(ahora);
      return;
    }
  }

  // El apóstrofo fuerza a Sheets a guardar como texto literal y evita
  // que reinterprete "2026-09-04" como un objeto Date al escribir vía API
  // (setNumberFormat('@') por sí solo no es suficiente para writes por API).
  sheet.appendRow(["'" + fechaStr, leadsDia, leadsAcumulados, total, ahora]);
}

function getResumen() {
  const sheet = getSheet();
  const values = sheet.getDataRange().getValues();
  values.shift();

  const ahora = new Date();
  const hoyStr = Utilities.formatDate(ahora, ZONA_HORARIA, 'yyyy-MM-dd');

  const hace7Date = new Date(ahora);
  hace7Date.setDate(hace7Date.getDate() - 6);
  const hace7Str = Utilities.formatDate(hace7Date, ZONA_HORARIA, 'yyyy-MM-dd');

  let hoyTotal = 0;
  let semanaTotal = 0;

  values.forEach(function (row) {
    const fechaStr = normalizarFecha(row[0]);
    if (!fechaStr) return;
    const total = Number(row[3]) || 0;

    if (fechaStr === hoyStr) {
      hoyTotal += total;
    }
    if (fechaStr >= hace7Str && fechaStr <= hoyStr) {
      semanaTotal += total;
    }
  });

  return {
    ok: true,
    hoy: hoyTotal,
    semana: semanaTotal
  };
}

function respuestaJSON(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
