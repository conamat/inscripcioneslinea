/**
 * Registro — Curso de Inducción a la Inteligencia Artificial
 * CONAMAT Plantel Hidalgo – San Cosme
 *
 * Recibe los registros que envía curso-ia.html y los guarda en la hoja
 * "Registros" de la Hoja de cálculo donde está pegado este script.
 * Opcional: manda un correo de confirmación al profesor.
 */

const HOJA = "Registros";
const ENVIAR_CONFIRMACION = true; // false para no mandar correo

const COLUMNAS = [
  ["fecha", "Fecha de registro"],
  ["folio", "Folio"],
  ["nombre", "Nombre(s)"],
  ["apellidos", "Apellidos"],
  ["correo", "Correo Gmail"],
  ["whatsapp", "WhatsApp"],
  ["plantel", "Secundaria"],
  ["puesto", "Puesto"],
  ["materias", "Materias"],
  ["niveles", "Grados"],
  ["experiencia", "Experiencia con IA"],
  ["herramientas", "Herramientas que conoce"],
  ["usos", "Para qué quiere usarla"],
  ["asistencia", "Confirma asistencia 18-oct"],
  ["laptop", "Llevará laptop"],
  ["correoIA", "Llevará Gmail y contraseña"],
  ["comentarios", "Comentarios"],
  ["aviso", "Aviso de privacidad"],
];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const d = JSON.parse(e.postData.contents);
    d.fecha = new Date();

    const hoja = obtenerHoja_();
    hoja.appendRow(COLUMNAS.map(([k]) => {
      const v = d[k] == null ? "" : d[k];
      // El apóstrofo evita que Sheets convierta el teléfono en número
      return k === "whatsapp" ? "'" + v : v;
    }));

    if (ENVIAR_CONFIRMACION && d.correo) enviarConfirmacion_(d);
    return ContentService.createTextOutput(JSON.stringify({ ok: true, folio: d.folio }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return ContentService.createTextOutput("Registro Curso de IA — activo");
}

function obtenerHoja_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let hoja = ss.getSheetByName(HOJA);
  if (!hoja) {
    hoja = ss.insertSheet(HOJA);
    hoja.appendRow(COLUMNAS.map(([, titulo]) => titulo));
    hoja.setFrozenRows(1);
    hoja.getRange(1, 1, 1, COLUMNAS.length)
      .setFontWeight("bold").setBackground("#14275a").setFontColor("#ffffff");
  }
  return hoja;
}

function enviarConfirmacion_(d) {
  const asunto = "Registro recibido · Curso de Inducción a la IA · CONAMAT";
  const html =
    '<div style="font-family:Arial,sans-serif;color:#15203d;max-width:520px">' +
    '<h2 style="color:#14275a">¡Hola, ' + escapar_(d.nombre) + '!</h2>' +
    '<p>Recibimos tu registro al <b>Curso de Inducción a la Inteligencia Artificial</b> ' +
    'de CONAMAT Plantel Hidalgo – San Cosme.</p>' +
    '<p>Tu folio es: <b style="font-size:18px">' + escapar_(d.folio) + '</b></p>' +
    '<p style="background:#14275a;color:#fff;padding:12px 16px;border-radius:10px;font-size:16px">' +
    '<b>Domingo 18 de octubre de 2026</b><br>De 9:00 a 13:00 h<br><br>' +
    '<b>CONAMAT Plantel San Cosme</b><br>Joaquín García Icazbalceta 57, col. San Rafael, ' +
    'alcaldía Cuauhtémoc, C.P. 06470, CDMX</p>' +
    '<p><a href="https://www.google.com/maps/search/?api=1&query=CONAMAT+San+Cosme+Joaqu%C3%ADn+Garc%C3%ADa+Icazbalceta+57+San+Rafael+06470">Ver ubicación en Google Maps</a></p>' +
    '<p><b>El cupo es limitado</b> y tu registro aparta un lugar. ¡Te esperamos!</p>' +
    '<p><b>No olvides traer:</b></p><ul>' +
    '<li>Tu equipo de cómputo (laptop cargada y su cargador).</li>' +
    '<li>Tu correo <b>@gmail.com</b> y su contraseña, para crear tus cuentas en las herramientas de IA.</li></ul>' +
    '<p style="color:#525e7a">Ing. Itzel Guerrero<br>Departamento de Vinculación<br>CONAMAT Hidalgo – San Cosme</p>' +
    '</div>';
  MailApp.sendEmail({ to: d.correo, subject: asunto, htmlBody: html, name: "CONAMAT Hidalgo – San Cosme" });
}

function escapar_(s) {
  return String(s || "").replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
}
