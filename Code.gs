// Google Apps Script — Web App para conteo de leads TikTok
// leads-tiktok-equilibrio (Contact Center Equilibrio Total)

const SHEET_NAME = 'Leads';
const ZONA_HORARIA = 'America/Mexico_City';
const SEDES_VALIDAS = ['Balbuena', 'Coacalco', 'Chalco', 'Neza'];

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
    const sede = p.sede;
    const leadsDia = parseInt(p.leadsDia, 10) || 0;
    const leadsAcumulados = parseInt(p.leadsAcumulados, 10) || 0;

    if (!fecha || !sede || SEDES_VALIDAS.indexOf(sede) === -1) {
      return respuestaJSON({ ok: false, error: 'Faltan campos obligatorios o sede inválida' });
    }

    guardarRegistro(fecha, sede, leadsDia, leadsAcumulados);

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
  const headers = ['Fecha', 'Sede', 'Leads del día', 'Leads acumulados', 'Total', 'Actualizado'];
  const primeraFila = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  const yaTieneHeaders = headers.every(function (h, i) { return primeraFila[i] === h; });
  if (!yaTieneHeaders) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function guardarRegistro(fechaStr, sede, leadsDia, leadsAcumulados) {
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  const total = leadsDia + leadsAcumulados;
  const ahora = new Date();

  for (let i = 1; i < data.length; i++) {
    const filaFecha = data[i][0];
    const filaSede = data[i][1];
    if (!(filaFecha instanceof Date)) continue;
    const filaFechaStr = Utilities.formatDate(filaFecha, ZONA_HORARIA, 'yyyy-MM-dd');
    if (filaFechaStr === fechaStr && filaSede === sede) {
      sheet.getRange(i + 1, 3, 1, 4).setValues([[leadsDia, leadsAcumulados, total, ahora]]);
      return;
    }
  }

  const fechaDate = new Date(fechaStr + 'T00:00:00');
  sheet.appendRow([fechaDate, sede, leadsDia, leadsAcumulados, total, ahora]);
}

function getResumen() {
  const sheet = getSheet();
  const values = sheet.getDataRange().getValues();
  values.shift(); // quitar encabezado

  const ahora = new Date();
  const hoyStr = Utilities.formatDate(ahora, ZONA_HORARIA, 'yyyy-MM-dd');

  const hace7 = new Date(ahora);
  hace7.setDate(hace7.getDate() - 6);
  hace7.setHours(0, 0, 0, 0);

  let hoyTotal = 0;
  let semanaTotal = 0;
  const porSedeHoy = {};
  SEDES_VALIDAS.forEach(function (s) { porSedeHoy[s] = 0; });

  values.forEach(function (row) {
    const fecha = row[0];
    if (!(fecha instanceof Date)) return;
    const total = Number(row[4]) || 0;
    const fechaStr = Utilities.formatDate(fecha, ZONA_HORARIA, 'yyyy-MM-dd');

    if (fechaStr === hoyStr) {
      hoyTotal += total;
      if (porSedeHoy.hasOwnProperty(row[1])) {
        porSedeHoy[row[1]] += total;
      }
    }
    if (fecha >= hace7) {
      semanaTotal += total;
    }
  });

  return {
    ok: true,
    hoy: hoyTotal,
    semana: semanaTotal,
    porSedeHoy: porSedeHoy
  };
}

function respuestaJSON(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
