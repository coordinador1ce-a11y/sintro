// ============================================================
//  SINTROPÍA SOCIAL — Google Apps Script v3 (JSONP)
//  INSTRUCCIÓN: Pega este código en Apps Script,
//  guarda y despliega como NUEVA VERSIÓN.
// ============================================================

var SHEET_ID         = '114sl6Mt-UhQQsv7zyicAAmsYzo3VDPoAvbT-0MakK94';
var SHEET_CITAS      = 'Hoja 1';
var SHEET_USUARIOS   = 'Usuarios';
var SHEET_PENDIENTES = 'Pendientes';

// Admins: email → SHA256 de contraseña
// Contraseña inicial: [VER_SCRIPT_PROPERTIES]
// ── ADMIN CONFIG — Leer desde Script Properties (no código) ──
// Para configurar: Apps Script → Configuración del proyecto → Propiedades del script
// Agrega: ADMIN_EMAIL = dsalgado@sintropiasocial.com
//          ADMIN_HASH  = (SHA-256 de tu contraseña)
function getAdminConfig() {
  var props = PropertiesService.getScriptProperties();
  var email = props.getProperty('ADMIN_EMAIL') || 'dsalgado@sintropiasocial.com';
  var hash  = props.getProperty('ADMIN_HASH')  || '41412db984c2db94df6515536ae3cdc10f5401914ba59a8436a1959346236d5d';
  var admins = {};
  admins[email.toLowerCase()] = hash;
  return admins;
}

// ── RESPUESTA JSONP (resuelve CORS desde GitHub Pages) ──
function jsonpResponse(data, callback) {
  var json = JSON.stringify(data);
  var output;
  if (callback) {
    // JSONP: el browser ejecuta callback(data)
    output = ContentService.createTextOutput(callback + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  } else {
    output = ContentService.createTextOutput(json)
      .setMimeType(ContentService.MimeType.JSON);
  }
  return output;
}

// ── ROUTER ──
function doGet(e) {
  var p        = e.parameter;
  var action   = p.action   || '';
  var callback = p.callback || '';
  var result;

  try {
    if      (action === 'getCitas')         result = getCitas();
    else if (action === 'getUsuarios')      result = getUsuarios(p);
    else if (action === 'getPendientes')    result = getPendientes(p);
    else if (action === 'registrarUsuario') result = registrarUsuario(p);
    else if (action === 'enviarCita')       result = enviarCita(p);
    else if (action === 'aprobarCita')      result = aprobarCita(p);
    else if (action === 'rechazarCita')     result = rechazarCita(p);
    else if (action === 'editarCita')       result = editarCita(p);
    else if (action === 'eliminarCita')     result = eliminarCita(p);
    else if (action === 'eliminarUsuario')  result = eliminarUsuario(p);
    else if (action === 'restablecerPass')  result = restablecerPass(p);
    else if (action === 'registrarContribucion') result = registrarContribucion(p);
    else if (action === 'registrarSuscriptor')   result = registrarSuscriptor(p);
    else if (action === 'loginAdmin')       result = loginAdmin(p);
    else if (action === 'loginUsuario')     result = loginUsuario(p);
    else if (action === 'getCitasPublicas') result = getCitasPublicas();
    else if (action === 'registrarBlogEntry') result = registrarBlogEntry(p);
    else if (action === 'getSuscriptores')  result = getSuscriptores(p);
    else if (action === 'getContribuciones') result = getContribuciones(p);
    else if (action === 'getEstadisticas')  result = getEstadisticas(p);
    else if (action === 'ping')             result = { ok: true, msg: 'Apps Script funcionando correctamente' };
    else result = { ok: false, error: 'Accion no reconocida: ' + action };
  } catch(err) {
    result = { ok: false, error: err.toString() };
  }

  return jsonpResponse(result, callback);
}

// ── CITAS ──
function getCitas() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(SHEET_CITAS);
  if (!sh) return { ok: false, error: 'No se encontro la hoja: ' + SHEET_CITAS };
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
    var cita = String(obj['Cita'] || '').trim();
    if (cita !== '' && cita !== 'undefined') {
      rows.push(obj);
    }
  }
  return { ok: true, data: rows };
}

function enviarCita(p) {
  ensureSheet(SHEET_PENDIENTES, ['No','Categoria','Indicador','Poblacion','Anio','Autor',
    'Cita','Comentarios','Publicacion','Pagina','Cita Apa','Link','Usuario','Fecha','Estado']);
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(SHEET_PENDIENTES);
  sh.appendRow([
    '', p.categoria||'', p.indicador||'', p.poblacion||'', p.year||'',
    p.autor||'', p.cita||'', p.comentarios||'', p.publicacion||'',
    p.pagina||'', p.citaAPA||'', p.link||'',
    p.usuarioEmail||'', new Date().toISOString(), 'PENDIENTE'
  ]);
  return { ok: true, msg: 'Cita enviada para revision' };
}

function aprobarCita(p) {
  if (!verificarAdmin(p.adminToken)) return { ok: false, error: 'No autorizado' };
  var rowIndex = parseInt(p.rowIndex);
  var ss    = SpreadsheetApp.openById(SHEET_ID);
  var pend  = ss.getSheetByName(SHEET_PENDIENTES);
  var citas = ss.getSheetByName(SHEET_CITAS);
  var data  = pend.getDataRange().getValues();
  var row   = data[rowIndex + 1];
  if (!row) return { ok: false, error: 'Fila no encontrada' };
  var newId = 'C' + String(citas.getLastRow()).padStart(4, '0');
  citas.appendRow([newId, row[1], row[2], row[3], row[4], row[5],
    row[6], row[7], '', row[8], row[9], row[10], row[11]]);
  pend.getRange(rowIndex + 2, 15).setValue('APROBADA');
  return { ok: true, msg: 'Cita aprobada y publicada' };
}

function rechazarCita(p) {
  if (!verificarAdmin(p.adminToken)) return { ok: false, error: 'No autorizado' };
  var ss   = SpreadsheetApp.openById(SHEET_ID);
  var pend = ss.getSheetByName(SHEET_PENDIENTES);
  pend.getRange(parseInt(p.rowIndex) + 2, 15).setValue('RECHAZADA: ' + (p.motivo || ''));
  return { ok: true, msg: 'Cita rechazada' };
}

function editarCita(p) {
  if (!verificarAdmin(p.adminToken)) return { ok: false, error: 'No autorizado' };
  var rowNum = parseInt(p.rowNum);
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(SHEET_CITAS);
  sh.getRange(rowNum, 2).setValue(p.categoria   || '');
  sh.getRange(rowNum, 3).setValue(p.indicador   || '');
  sh.getRange(rowNum, 4).setValue(p.poblacion   || '');
  sh.getRange(rowNum, 5).setValue(p.year        || '');
  sh.getRange(rowNum, 6).setValue(p.autor       || '');
  sh.getRange(rowNum, 7).setValue(p.cita        || '');
  sh.getRange(rowNum, 8).setValue(p.comentarios || '');
  sh.getRange(rowNum, 11).setValue(p.pagina     || '');
  sh.getRange(rowNum, 12).setValue(p.citaAPA    || '');
  sh.getRange(rowNum, 13).setValue(p.link       || '');
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
  ensureSheet(SHEET_USUARIOS, ['ID','Nombre','Apellido','Email',
    'Institucion','Area','Motivo','Fecha','Contrasena','Estado']);
  var ss   = SpreadsheetApp.openById(SHEET_ID);
  var sh   = ss.getSheetByName(SHEET_USUARIOS);
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][3]).toLowerCase() === String(p.email || '').toLowerCase()) {
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
  var data    = sh.getDataRange().getValues();
  var headers = data[0];
  var rows    = [];
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
  var chars    = 'ABCDEFGHJKMNPQRSTUVWXYZabcdefghjkmnpqrstuvwxyz23456789';
  var tempPass = 'Tmp';
  for (var i = 0; i < 6; i++) {
    tempPass += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  tempPass += '!';
  var hash = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256, tempPass, Utilities.Charset.UTF_8
  ).map(function(b) {
    return (b < 0 ? b + 256 : b).toString(16).padStart(2, '0');
  }).join('');
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(SHEET_USUARIOS);
  sh.getRange(parseInt(p.rowNum), 9).setValue(hash);
  try {
    MailApp.sendEmail({
      to: p.email,
      subject: 'Sintropía Social — Contraseña restablecida',
      body: 'Hola,\n\nTu contraseña fue restablecida.\n\nContraseña temporal: ' + tempPass +
            '\n\nPor favor cámbiala al ingresar.\n\nSintropía Social\ncontacto@sintropiasocial.com'
    });
  } catch(err) {}
  return { ok: true, msg: 'Contrasena restablecida. Temporal: ' + tempPass };
}

function getPendientes(p) {
  if (!verificarAdmin(p.adminToken)) return { ok: false, error: 'No autorizado' };
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(SHEET_PENDIENTES);
  if (!sh) return { ok: true, data: [] };
  var data    = sh.getDataRange().getValues();
  var headers = data[0];
  var rows    = [];
  for (var i = 1; i < data.length; i++) {
    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      obj[String(headers[j]).trim()] = data[i][j];
    }
    obj._rowIndex = i - 1;
    if (String(obj['Estado']) === 'PENDIENTE') rows.push(obj);
  }
  return { ok: true, data: rows };
}

// ── ADMIN AUTH ──

// ── RATE LIMITING ──
var RATE_LIMIT_MAX = 10;  // max intentos por ventana
var RATE_LIMIT_WINDOW_MS = 15 * 60 * 1000; // 15 minutos

function checkRateLimit(ip_or_email) {
  var key = 'rl_' + Utilities.computeDigest(
    Utilities.DigestAlgorithm.MD5,
    String(ip_or_email),
    Utilities.Charset.UTF_8
  ).map(function(b){ return (b<0?b+256:b).toString(16).padStart(2,'0'); }).join('').substring(0,16);
  
  var props = PropertiesService.getScriptProperties();
  var raw = props.getProperty(key);
  var now = Date.now();
  var record = raw ? JSON.parse(raw) : { count: 0, window_start: now };
  
  // Reset window if expired
  if(now - record.window_start > RATE_LIMIT_WINDOW_MS) {
    record = { count: 0, window_start: now };
  }
  
  record.count++;
  props.setProperty(key, JSON.stringify(record));
  
  if(record.count > RATE_LIMIT_MAX) {
    var remaining = Math.ceil((RATE_LIMIT_WINDOW_MS - (now - record.window_start)) / 60000);
    return { limited: true, msg: 'Demasiados intentos. Espera ' + remaining + ' minutos.' };
  }
  return { limited: false };
}

function resetRateLimit(ip_or_email) {
  var key = 'rl_' + Utilities.computeDigest(
    Utilities.DigestAlgorithm.MD5,
    String(ip_or_email),
    Utilities.Charset.UTF_8
  ).map(function(b){ return (b<0?b+256:b).toString(16).padStart(2,'0'); }).join('').substring(0,16);
  PropertiesService.getScriptProperties().deleteProperty(key);
}

function loginAdmin(p) {
  // Rate limiting
  var email = String(p.email||'').toLowerCase();
  var rl = checkRateLimit('login_' + email);
  if(rl.limited) return { ok: false, error: rl.msg };

  var email    = String(p.email    || '').toLowerCase();
  var passHash = String(p.passHash || '');
  var ADMINS = getAdminConfig();
  if (ADMINS[email] && ADMINS[email] === passHash) {
    resetRateLimit('login_' + email);
    var raw   = email + passHash + 'sintropia_salt_2025';
    var token = Utilities.computeDigest(
      Utilities.DigestAlgorithm.SHA_256, raw, Utilities.Charset.UTF_8
    ).map(function(b) {
      return (b < 0 ? b + 256 : b).toString(16).padStart(2, '0');
    }).join('');
    var tokenData = JSON.stringify({ email: email, created: Date.now(), expires: Date.now() + 28800000 });
    PropertiesService.getScriptProperties().setProperty('adm_' + token, tokenData);
    return { ok: true, token: token, email: email };
  }
  return { ok: false, error: 'Credenciales incorrectas' };
}

function verificarAdmin(token) {
  if (!token) return false;
  try {
    var raw = PropertiesService.getScriptProperties().getProperty('adm_' + token);
    if (!raw) return false;
    try {
      var data = JSON.parse(raw);
      if (data && data.expires && Date.now() > data.expires) {
        PropertiesService.getScriptProperties().deleteProperty('adm_' + token);
        return false;
      }
    } catch(e) { /* old format = plain email string, still valid */ }
    return true;
  } catch(e) { return false; }
}

// ── HELPERS ──

function registrarContribucion(p) {
  ensureSheet('Contribuciones', ['ID','Titulo','Tipo','Resumen','Contenido','Archivo','Usuario','Fecha','Estado']);
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName('Contribuciones');
  var id = 'C' + new Date().getTime();
  sh.appendRow([id, p.titulo||'', p.tipo||'', p.resumen||'', p.contenido||'', p.archivo||'', p.usuario||'', p.fecha||new Date().toISOString(), 'pendiente']);
  return { ok: true, id: id, msg: 'Contribucion registrada' };
}

function registrarSuscriptor(p) {
  ensureSheet('Suscriptores', ['Email','Fecha','Estado']);
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName('Suscriptores');
  var data = sh.getDataRange().getValues();
  for(var i=1;i<data.length;i++){
    if(String(data[i][0]).toLowerCase()===String(p.email||'').toLowerCase())
      return { ok: true, msg: 'Ya suscrito' };
  }
  sh.appendRow([p.email||'', new Date().toISOString(), 'activo']);
  return { ok: true, msg: 'Suscriptor registrado' };
}

function loginUsuario(p) {
  var email    = String(p.email    || '').toLowerCase();
  var passHash = String(p.passHash || '');
  
  // Rate limiting
  var rl = checkRateLimit('user_' + email);
  if (rl.limited) return { ok: false, error: rl.msg };
  
  // Check if admin
  var ADMINS = getAdminConfig();
  if (ADMINS[email] && ADMINS[email] === passHash) {
    resetRateLimit('user_' + email);
    var token = Utilities.computeDigest(
      Utilities.DigestAlgorithm.SHA_256,
      email + passHash + 'sintropia2025',
      Utilities.Charset.UTF_8
    ).map(function(b){ return (b<0?b+256:b).toString(16).padStart(2,'0'); }).join('');
    var tokenData = JSON.stringify({ email: email, created: Date.now(), expires: Date.now() + 28800000 });
    PropertiesService.getScriptProperties().setProperty('adm_' + token, tokenData);
    return { ok: true, isAdmin: true, token: token, nombre: 'Admin', email: email };
  }
  
  // Check regular users
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName('Usuarios');
  if (!sh) return { ok: false, error: 'Sin usuarios registrados' };
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][3]).toLowerCase() === email && data[i][8] === passHash) {
      resetRateLimit('user_' + email);
      return { ok: true, id: data[i][0], nombre: data[i][1], apellido: data[i][2], email: data[i][3], area: data[i][5] };
    }
  }
  return { ok: false, error: 'Correo o contrasena incorrectos' };
}

function getCitasPublicas() {
  var r = getCitas();
  if (!r.ok) return r;
  var pub = r.data.slice(0, Math.ceil(r.data.length * 0.02));
  return { ok: true, data: pub, total: r.data.length };
}

function registrarBlogEntry(p) {
  ensureSheet('Blog', ['Titulo','Tag','Body','Link','Fecha','AdminEmail']);
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName('Blog');
  sh.appendRow([p.titulo||'', p.tag||'', p.body||'', p.link||'', p.fecha||new Date().toISOString(), p.adminEmail||'']);
  return { ok: true, msg: 'Entrada registrada' };
}

function getSuscriptores(p) {
  if (!verificarAdmin(p.adminToken)) return { ok: false, error: 'No autorizado' };
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName('Suscriptores');
  if (!sh) return { ok: true, data: [] };
  var data = sh.getDataRange().getValues();
  var headers = data[0];
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    var obj = {};
    for (var j = 0; j < headers.length; j++) obj[String(headers[j]).trim()] = data[i][j];
    rows.push(obj);
  }
  return { ok: true, data: rows };
}

function getContribuciones(p) {
  if (!verificarAdmin(p.adminToken)) return { ok: false, error: 'No autorizado' };
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName('Contribuciones');
  if (!sh) return { ok: true, data: [] };
  var data = sh.getDataRange().getValues();
  var headers = data[0];
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    var obj = {};
    for (var j = 0; j < headers.length; j++) obj[String(headers[j]).trim()] = data[i][j];
    rows.push(obj);
  }
  return { ok: true, data: rows };
}

function getEstadisticas(p) {
  if (!verificarAdmin(p.adminToken)) return { ok: false, error: 'No autorizado' };
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var citas = (ss.getSheetByName('Hoja 1') || ss.getSheetByName(SHEET_CITAS));
  var usuarios = ss.getSheetByName('Usuarios');
  var pendientes = ss.getSheetByName('Pendientes');
  var suscriptores = ss.getSheetByName('Suscriptores');
  return {
    ok: true,
    citas:       citas       ? Math.max(0, citas.getLastRow() - 1)       : 0,
    usuarios:    usuarios    ? Math.max(0, usuarios.getLastRow() - 1)    : 0,
    pendientes:  pendientes  ? Math.max(0, pendientes.getLastRow() - 1)  : 0,
    suscriptores:suscriptores? Math.max(0, suscriptores.getLastRow() - 1): 0
  };
}

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
