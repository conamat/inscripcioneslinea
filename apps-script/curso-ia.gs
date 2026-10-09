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
  const asunto = "Tu lugar en el Curso de IA · Domingo 18 de octubre · CONAMAT San Cosme";
  const mapa = "https://www.google.com/maps/search/?api=1&query=CONAMAT+San+Cosme+Joaqu%C3%ADn+Garc%C3%ADa+Icazbalceta+57+San+Rafael+06470";
  const logo = "https://conamat.github.io/inscripcioneslinea/logo-conamat.png";
  const F = "font-family:Arial,Helvetica,sans-serif;";
  const fila = (etq, val) =>
    '<tr><td style="' + F + 'padding:10px 0;border-bottom:1px solid #e3e8f2;width:110px;font-size:13px;color:#6b7591;vertical-align:top">' + etq + '</td>' +
    '<td style="' + F + 'padding:10px 0;border-bottom:1px solid #e3e8f2;font-size:15px;color:#15203d;vertical-align:top">' + val + '</td></tr>';
  const check = (txt) =>
    '<tr><td style="padding:6px 10px 6px 0;vertical-align:top;width:26px">' +
    '<div style="' + F + 'width:22px;height:22px;line-height:22px;border-radius:11px;background:#e6f6ed;color:#1f9d57;font-size:13px;font-weight:bold;text-align:center">✓</div></td>' +
    '<td style="' + F + 'padding:6px 0;font-size:15px;line-height:1.45;color:#15203d">' + txt + '</td></tr>';

  const html =
    '<div style="margin:0;padding:24px 12px;background:#f1f4f9">' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="max-width:560px;margin:0 auto;background:#ffffff;border-radius:14px;overflow:hidden;border:1px solid #dbe2ee">' +

    // Logo
    '<tr><td style="padding:22px 28px 18px;border-bottom:4px solid #14275a">' +
    '<img src="' + logo + '" width="300" alt="CONAMAT · Plantel Hidalgo – San Cosme" style="display:block;width:300px;max-width:100%;height:auto;border:0"></td></tr>' +

    // Saludo
    '<tr><td style="' + F + 'padding:28px 28px 8px">' +
    '<div style="font-size:13px;font-weight:bold;letter-spacing:1px;color:#1f9d57;text-transform:uppercase">Registro confirmado</div>' +
    '<h1 style="margin:8px 0 10px;font-size:24px;line-height:1.25;color:#14275a">¡Hola, ' + escapar_(d.nombre) + '! Ya tienes tu lugar.</h1>' +
    '<p style="margin:0;font-size:15px;line-height:1.55;color:#525e7a">Recibimos tu registro al <b style="color:#15203d">Curso de Inducción a la Inteligencia Artificial</b>. Guarda este correo: aquí tienes todos los datos.</p>' +
    '</td></tr>' +

    // Folio
    '<tr><td style="padding:18px 28px 6px">' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#eef3fc;border:1px dashed #2b50a8;border-radius:10px">' +
    '<tr><td style="' + F + 'padding:14px 18px;font-size:13px;color:#525e7a">Tu folio</td>' +
    '<td align="right" style="' + F + 'padding:14px 18px;font-size:22px;font-weight:bold;letter-spacing:2px;color:#14275a">' + escapar_(d.folio) + '</td></tr></table>' +
    '</td></tr>' +

    // Datos del curso
    '<tr><td style="padding:18px 28px 4px">' +
    '<div style="' + F + 'font-size:16px;font-weight:bold;color:#14275a;margin-bottom:4px">Datos del curso</div>' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0">' +
    fila("Fecha", "<b>Domingo 18 de octubre de 2026</b>") +
    fila("Horario", "De 9:00 a 13:00 h") +
    fila("Lugar", "<b>CONAMAT Plantel San Cosme</b><br>Joaquín García Icazbalceta 57, col. San Rafael, alcaldía Cuauhtémoc, C.P. 06470, CDMX") +
    fila("Costo", "Gratuito") +
    '</table></td></tr>' +

    // Botón mapa
    '<tr><td style="padding:18px 28px 6px">' +
    '<a href="' + mapa + '" style="' + F + 'display:inline-block;background:#14275a;color:#ffffff;text-decoration:none;font-size:15px;font-weight:bold;padding:13px 22px;border-radius:24px">Cómo llegar en Google Maps</a>' +
    '</td></tr>' +

    // Qué llevar
    '<tr><td style="padding:22px 28px 6px">' +
    '<div style="' + F + 'font-size:16px;font-weight:bold;color:#14275a;margin-bottom:6px">No olvides llevar</div>' +
    '<table role="presentation" cellpadding="0" cellspacing="0">' +
    check("<b>Tu equipo de cómputo</b> (laptop con batería cargada y su cargador). El curso es práctico.") +
    check("<b>Tu correo @gmail.com y su contraseña</b>, para crear tus cuentas en las herramientas de IA.") +
    '</table></td></tr>' +

    // Cupo
    '<tr><td style="padding:16px 28px 26px">' +
    '<div style="' + F + 'background:#fdeeec;border-radius:10px;padding:12px 16px;font-size:14px;line-height:1.5;color:#8f2a20"><b>El cupo es limitado.</b> ¡Te esperamos puntual!</div>' +
    '</td></tr>' +

    // Firma
    '<tr><td style="' + F + 'padding:18px 28px 24px;border-top:1px solid #e3e8f2;font-size:14px;line-height:1.5;color:#525e7a">' +
    '<b style="color:#15203d">CONAMAT Plantel Hidalgo – San Cosme</b>' +
    '</td></tr>' +

    '</table>' +
    '<p style="' + F + 'max-width:560px;margin:14px auto 0;text-align:center;font-size:12px;color:#8790a6">Recibiste este correo porque te registraste al Curso de Inducción a la IA de CONAMAT Hidalgo – San Cosme.</p>' +
    '</div>';

  const texto =
    "¡Hola, " + d.nombre + "! Ya tienes tu lugar en el Curso de Inducción a la Inteligencia Artificial.\n\n" +
    "Folio: " + d.folio + "\n" +
    "Fecha: domingo 18 de octubre de 2026, de 9:00 a 13:00 h\n" +
    "Lugar: CONAMAT Plantel San Cosme, Joaquín García Icazbalceta 57, col. San Rafael, C.P. 06470, CDMX\n" +
    "Mapa: " + mapa + "\n\n" +
    "No olvides llevar tu laptop con cargador y tu correo @gmail.com con su contraseña.\n" +
    "El cupo es limitado. ¡Te esperamos!\n\n" +
    "CONAMAT Plantel Hidalgo – San Cosme";

  MailApp.sendEmail({ to: d.correo, subject: asunto, body: texto, htmlBody: html, name: "CONAMAT Hidalgo – San Cosme" });
}

// Ejecuta esta función desde el editor para recibir un correo de muestra en tu propia cuenta
function probarCorreo() {
  enviarConfirmacion_({ nombre: "Itzel", folio: "IA-PRUEBA", correo: Session.getActiveUser().getEmail() });
}

function escapar_(s) {
  return String(s || "").replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
}
