// ============================================================
//  SINTROPÍA SOCIAL — Google Apps Script Backend v2
//  INSTRUCCIÓN: Reemplaza el contenido de Code.gs con esto,
//  luego vuelve a desplegar como Web App (nueva versión).
// ============================================================

var SHEET_ID        = '114sl6Mt-UhQQsv7zyicAAmsYzo3VDPoAvbT-0MakK94';
var SHEET_CITAS     = 'Hoja 1';
var SHEET_USUARIOS  = 'Usuarios';
var SHEET_PENDIENTES = 'Pendientes';

// Admins: email → SHA256 de contraseña
// Contraseña inicial de dsalgado: Sintropia2025!
var ADMINS = {
  'dsalgado@sintropiasocial.com': '41412db984c2db94df6515536ae3cdc10f5401914ba59a8436a1959346236d5d'
  // Para agregar otra admin:
  // ,'otra@sintropiasocial.com': 'SHA256_DE_SU_CONTRASENA'
};

// ── CORS ──
function makeResponse(data) {
  var output = ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
  return output;
}

// ── ROUTER — todo llega por GET ──
function doGet(e) {
  var p = e.parameter;
  var action = p.action || '';
  var result;

  try {
    if      (action === 'getCitas')          result = getCitas();
    else if (action === 'getUsuarios')       result = getUsuarios(p);
    else if (action === 'getPendientes')     result = getPendientes(p);
    else if (action === 'registrarUsuario')  result = registrarUsuario(p);
    else if (action === 'enviarCita')        result = enviarCita(p);
    else if (action === 'aprobarCita')       result = aprobarCita(p);
    else if (action === 'rechazarCita')      result = rechazarCita(p);
    else if (action === 'editarCita')        result = editarCita(p);
    else if (action === 'eliminarCita')      result = eliminarCita(p);
    else if (action === 'eliminarUsuario')   result = eliminarUsuario(p);
    else if (action === 'restablecerPass')   result = restablecerPass(p);
    else if (action === 'loginAdmin')        result = loginAdmin(p);
    else result = { ok: false, error: 'Accion no reconocida: ' + action };
  } catch(err) {
    result = { ok: false, error: err.toString() };
  }

  return makeResponse(result);
}

// ── CITAS ──
function getCitas() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(SHEET_CITAS);
  if (!sh) return { ok: false, error: 'No se encontró la hoja: ' + SHEET_CITAS };
  var data = sh.getDataRange().getValues();
  if (data.length < 2) return { ok: true, data: [] };
  var headers = data[0];
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      obj[String(headers[j]).trim()] = row[j];
    }
    obj._row = i + 1;
    if (obj['Cita'] && String(obj['Cita']).trim() !== '') {
      rows.push(obj);
    }
  }
  return { ok: true, data: rows };
}

function enviarCita(p) {
  ensureSheet(SHEET_PENDIENTES, ['No','Categoria','Indicador','Poblacion','Anio','Autor','Cita','Comentarios','Publicacion','Pagina','Cita Apa','Link','Usuario','Fecha','Estado']);
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(SHEET_PENDIENTES);
  sh.appendRow([
    '', p.categoria||'', p.indicador||'', p.poblacion||'',
    p.year||'', p.autor||'', p.cita||'', p.comentarios||'',
    p.publicacion||'', p.pagina||'', p.citaAPA||'', p.link||'',
    p.usuarioEmail||'', new Date().toISOString(), 'PENDIENTE'
  ]);
  return { ok: true, msg: 'Cita enviada para revision' };
}

function aprobarCita(p) {
  if (!verificarAdmin(p.adminToken)) return { ok: false, error: 'No autorizado' };
  var rowIndex = parseInt(p.rowIndex);
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var pend  = ss.getSheetByName(SHEET_PENDIENTES);
  var citas = ss.getSheetByName(SHEET_CITAS);
  var data  = pend.getDataRange().getValues();
  var row   = data[rowIndex + 1];
  if (!row) return { ok: false, error: 'Fila no encontrada' };
  var newId = 'C' + String(citas.getLastRow()).padStart(4,'0');
  citas.appendRow([newId, row[1], row[2], row[3], row[4], row[5], row[6], row[7], '', row[8], row[9], row[10], row[11]]);
  pend.getRange(rowIndex + 2, 15).setValue('APROBADA');
  return { ok: true, msg: 'Cita aprobada y publicada' };
}

function rechazarCita(p) {
  if (!verificarAdmin(p.adminToken)) return { ok: false, error: 'No autorizado' };
  var rowIndex = parseInt(p.rowIndex);
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var pend = ss.getSheetByName(SHEET_PENDIENTES);
  pend.getRange(rowIndex + 2, 15).setValue('RECHAZADA: ' + (p.motivo||''));
  return { ok: true, msg: 'Cita rechazada' };
}

function editarCita(p) {
  if (!verificarAdmin(p.adminToken)) return { ok: false, error: 'No autorizado' };
  var rowNum = parseInt(p.rowNum);
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(SHEET_CITAS);
  sh.getRange(rowNum, 2).setValue(p.categoria||'');
  sh.getRange(rowNum, 3).setValue(p.indicador||'');
  sh.getRange(rowNum, 4).setValue(p.poblacion||'');
  sh.getRange(rowNum, 5).setValue(p.year||'');
  sh.getRange(rowNum, 6).setValue(p.autor||'');
  sh.getRange(rowNum, 7).setValue(p.cita||'');
  sh.getRange(rowNum, 8).setValue(p.comentarios||'');
  sh.getRange(rowNum, 11).setValue(p.pagina||'');
  sh.getRange(rowNum, 12).setValue(p.citaAPA||'');
  sh.getRange(rowNum, 13).setValue(p.link||'');
  return { ok: true, msg: 'Cita actualizada' };
}

function eliminarCita(p) {
  if (!verificarAdmin(p.adminToken)) return { ok: false, error: 'No autorizado' };
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(SHEET_CITAS);
  sh.deleteRow(parseInt(p.rowNum));
  return { ok: true, msg: 'Cita eliminada' };
}

// ── USUARIOS ──
function registrarUsuario(p) {
  ensureSheet(SHEET_USUARIOS, ['ID','Nombre','Apellido','Email','Institucion','Area','Motivo','Fecha','Contrasena','Estado']);
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(SHEET_USUARIOS);
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][3]).toLowerCase() === String(p.email||'').toLowerCase()) {
      return { ok: false, error: 'Este correo ya esta registrado' };
    }
  }
  var id = 'U' + new Date().getTime();
  sh.appendRow([id, p.nombre||'', p.apellido||'', p.email||'',
    p.institucion||'', p.area||'', p.motivo||'',
    new Date().toISOString(), p.passHash||'', 'ACTIVO']);
  return { ok: true, id: id, msg: 'Usuario registrado' };
}

function getUsuarios(p) {
  if (!verificarAdmin(p.adminToken)) return { ok: false, error: 'No autorizado' };
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(SHEET_USUARIOS);
  if (!sh) return { ok: true, data: [] };
  var data = sh.getDataRange().getValues();
  var headers = data[0];
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      obj[String(headers[j]).trim()] = (j === 8) ? '***' : data[i][j];
    }
    obj._row = i + 1;
    rows.push(obj);
  }
  return { ok: true, data: rows };
}

function eliminarUsuario(p) {
  if (!verificarAdmin(p.adminToken)) return { ok: false, error: 'No autorizado' };
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(SHEET_USUARIOS);
  sh.deleteRow(parseInt(p.rowNum));
  return { ok: true, msg: 'Usuario eliminado' };
}

function restablecerPass(p) {
  if (!verificarAdmin(p.adminToken)) return { ok: false, error: 'No autorizado' };
  var chars = 'ABCDEFGHJKMNPQRSTUVWXYZabcdefghjkmnpqrstuvwxyz23456789';
  var tempPass = 'Tmp';
  for (var i = 0; i < 6; i++) {
    tempPass += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  tempPass += '!';
  var hash = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256, tempPass, Utilities.Charset.UTF_8
  ).map(function(b){ return (b < 0 ? b+256 : b).toString(16).padStart(2,'0'); }).join('');
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(SHEET_USUARIOS);
  sh.getRange(parseInt(p.rowNum), 9).setValue(hash);
  try {
    MailApp.sendEmail({
      to: p.email,
      subject: 'Sintropía Social — Contraseña restablecida',
      body: 'Hola,\n\nTu contraseña fue restablecida.\n\nContraseña temporal: ' + tempPass + '\n\nPor favor cámbiala al ingresar.\n\nSintropía Social\ncontacto@sintropiasocial.com'
    });
  } catch(err) {}
  return { ok: true, msg: 'Contrasena restablecida. Temporal: ' + tempPass + '. Email enviado a ' + p.email };
}

function getPendientes(p) {
  if (!verificarAdmin(p.adminToken)) return { ok: false, error: 'No autorizado' };
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(SHEET_PENDIENTES);
  if (!sh) return { ok: true, data: [] };
  var data = sh.getDataRange().getValues();
  var headers = data[0];
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      obj[String(headers[j]).trim()] = data[i][j];
    }
    obj._rowIndex = i - 1;
    if (obj['Estado'] === 'PENDIENTE') rows.push(obj);
  }
  return { ok: true, data: rows };
}

// ── ADMIN AUTH ──
function loginAdmin(p) {
  var email    = String(p.email||'').toLowerCase();
  var passHash = String(p.passHash||'');
  if (ADMINS[email] && ADMINS[email] === passHash) {
    var raw = email + passHash + 'sintropia_salt_2025';
    var token = Utilities.computeDigest(
      Utilities.DigestAlgorithm.SHA_256, raw, Utilities.Charset.UTF_8
    ).map(function(b){ return (b<0?b+256:b).toString(16).padStart(2,'0'); }).join('');
    PropertiesService.getScriptProperties().setProperty('adm_' + token, email);
    return { ok: true, token: token, email: email };
  }
  return { ok: false, error: 'Credenciales incorrectas' };
}

function verificarAdmin(token) {
  if (!token) return false;
  var val = PropertiesService.getScriptProperties().getProperty('adm_' + token);
  return !!val;
}

// ── HELPERS ──
function ensureSheet(name, headers) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }
  return sh;
}
