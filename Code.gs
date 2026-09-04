// Google Apps Script — Web App para recibir leads del formulario
// leads-tiktok-equilibrio (Contact Center Equilibrio Total)
//
// Este archivo es SOLO REFERENCIA/BACKUP. El código real vive en el
// proyecto de Google Apps Script ligado al Google Sheet (script.google.com).

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
    const sheet = getSheet();
    const p = e.parameter;

    if (!p.sede || !p.nombre || !p.telefono || !p.campana || !p.estatus) {
      return respuestaJSON({ ok: false, error: 'Faltan campos obligatorios' });
    }

    sheet.appendRow([
      new Date(),
      p.sede,
      p.nombre,
      p.telefono,
      p.campana,
      p.estatus,
      p.notas || ''
    ]);

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
    sheet.appendRow(['Fecha y hora', 'Sede', 'Nombre', 'Teléfono', 'Campaña de origen', 'Estatus', 'Notas']);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getResumen() {
  const sheet = getSheet();
  const values = sheet.getDataRange().getValues();
  values.shift(); // quitar encabezado

  const ahora = new Date();
  const hoyStr = Utilities.formatDate(ahora, ZONA_HORARIA, 'yyyy-MM-dd');

  const inicioSemana = new Date(ahora);
  const diaSemana = inicioSemana.getDay(); // 0 = domingo
  const offset = diaSemana === 0 ? 6 : diaSemana - 1; // semana inicia lunes
  inicioSemana.setDate(inicioSemana.getDate() - offset);
  inicioSemana.setHours(0, 0, 0, 0);

  let hoy = 0, semana = 0, agendaronSemana = 0;

  values.forEach(function (row) {
    const fecha = row[0];
    if (!(fecha instanceof Date)) return;

    const fechaStr = Utilities.formatDate(fecha, ZONA_HORARIA, 'yyyy-MM-dd');
    if (fechaStr === hoyStr) hoy++;

    if (fecha >= inicioSemana) {
      semana++;
      if (row[5] === 'Agendó cita') agendaronSemana++;
    }
  });

  const pctAgendo = semana > 0 ? Math.round((agendaronSemana / semana) * 100) : 0;

  return { ok: true, hoy: hoy, semana: semana, pctAgendo: pctAgendo, total: values.length };
}

function respuestaJSON(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
