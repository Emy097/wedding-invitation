var SHEET_NAME   = 'Confirmaciones';
var VITACORA     = 'Bitacora';
var INVITADOS    = 'Invitados';   /* lista oficial — columna A: nombre */

/**
 * La página permite ABRIR la invitación con cualquier nombre cuando la URL
 * lleva ?extra. Por eso la lista de invitados se valida aquí, del lado del
 * servidor: sin esta validación cualquiera podría insertar filas basura en
 * Confirmaciones. ?extra deja ver la invitación; sólo la hoja "Invitados"
 * decide quién puede escribir.
 */

function doGet(e) {
  return e.parameter.action === 'status' ? getStatus(e) : submitRsvp(e);
}

/**
 * Normaliza un nombre para comparar: unifica la forma Unicode de los acentos,
 * recorta espacios y no distingue mayúsculas.
 *
 * Lo de Unicode importa: "José" tecleado en Sheets y "José" enviado por el
 * navegador pueden ser secuencias de bytes distintas que fallan con ===.
 */
function normalizar_(s) {
  return String(s == null ? '' : s)
    .normalize('NFC')
    .trim()
    .replace(/\s+/g, ' ')
    .toLowerCase();
}

/**
 * Busca el nombre en la hoja "Invitados".
 * @return {string} 'invitado' | 'no-invitado' | 'sin-lista'
 *
 * 'sin-lista' (hoja ausente o vacía) NO bloquea: si la lista todavía no está
 * cargada, es preferible aceptar respuestas a dejar fuera a todos los
 * invitados reales. La validación empieza a aplicar en cuanto haya nombres.
 */
function estadoInvitado_(nombre) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(INVITADOS);
  if (!sheet) return 'sin-lista';

  var ultima = sheet.getLastRow();
  if (ultima < 2) return 'sin-lista';

  var objetivo = normalizar_(nombre);
  if (!objetivo) return 'no-invitado';

  var nombres = sheet.getRange(2, 1, ultima - 1, 1).getValues();
  for (var i = 0; i < nombres.length; i++) {
    if (normalizar_(nombres[i][0]) === objetivo) return 'invitado';
  }
  return 'no-invitado';
}

function registrarBitacora_(nombre, asistencia, accion) {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var log = ss.getSheetByName(VITACORA);
  if (!log) {
    log = ss.insertSheet(VITACORA);
    log.appendRow(['Fecha', 'Nombre', 'Asistencia', 'Acción']);
    log.getRange('1:1').setFontWeight('bold');
  }
  log.appendRow([new Date(), nombre, asistencia, accion]);
}

function getStatus(e) {
  try {
    var nombre = (e.parameter.nombre || '').trim();
    if (!nombre) return respond({ ok: false, error: 'missing_nombre' });

    /* No se registra en Bitacora: esto corre en cada carga de página */
    if (estadoInvitado_(nombre) === 'no-invitado') {
      return respond({ ok: false, error: 'not_found' });
    }

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    if (!sheet) return respond({ ok: true, asistencia: null });

    var rows = sheet.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (normalizar_(rows[i][0]) === normalizar_(nombre)) {
        return respond({ ok: true, asistencia: rows[i][1] || null });
      }
    }
    return respond({ ok: true, asistencia: null });
  } catch (err) {
    return respond({ ok: false, error: err.message });
  }
}

function submitRsvp(e) {
  try {
    var nombre     = (e.parameter.nombre     || '').trim();
    var asistencia = (e.parameter.asistencia || '').trim();

    if (!nombre || !asistencia) {
      return respond({ ok: false, error: 'missing_params' });
    }

    /* ── Validación: sólo escriben los nombres de la hoja "Invitados" ── */
    if (estadoInvitado_(nombre) === 'no-invitado') {
      registrarBitacora_(nombre, asistencia, 'Rechazado — no está en Invitados');
      return respond({ ok: false, error: 'not_found' });
    }

    var ss   = SpreadsheetApp.getActiveSpreadsheet();
    var now  = new Date();

    /* ── Confirmaciones: keep only the latest per person ── */
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow(['Nombre', 'Asistencia', 'Última actualización']);
      sheet.getRange('1:1').setFontWeight('bold');
    }

    var rows    = sheet.getDataRange().getValues();
    var updated = false;

    for (var i = 1; i < rows.length; i++) {
      if (normalizar_(rows[i][0]) === normalizar_(nombre)) {
        sheet.getRange(i + 1, 2).setValue(asistencia);
        sheet.getRange(i + 1, 3).setValue(now);
        updated = true;
        break;
      }
    }
    if (!updated) sheet.appendRow([nombre, asistencia, now]);

    /* ── Vitacora: append every interaction ── */
    registrarBitacora_(nombre, asistencia, updated ? 'Actualización' : 'Nueva respuesta');

    return respond({ ok: true, updated: updated });
  } catch (err) {
    return respond({ ok: false, error: err.message });
  }
}

function respond(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
