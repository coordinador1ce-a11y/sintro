var SHEET_ID = '114sl6Mt-UhQQsv7zyicAAmsYzo3VDPoAvbT-0MakK94';
var SHEET_CITAS = 'Hoja 1';
var SHEET_USUARIOS = 'Usuarios';
var SHEET_PENDIENTES = 'Pendientes';
var SHEET_DESCARGAS = 'Descargas';
var SHEET_BLOG = 'Blog';

var ADMINS = {
  'dsalgado@sintropiasocial.com': '41412db984c2db94df6515536ae3cdc10f5401914ba59a8436a1959346236d5d'
};

function makeResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  // Si no hay parámetros, devolver info básica
  if (!e || !e.parameter) {
    return makeResponse({ 
      ok: true, 
      message: 'API de Sintropía Social funcionando',
      version: '2.0'
    });
  }
  
  var p = e.parameter;
  var action = p.action || '';
  var result;

  try {
    if (action === 'getCitas') result = getCitas();
    else if (action === 'getCitasPublicas') result = getCitasPublicas();
    else if (action === 'getUsuarios') result = getUsuarios(p);
    else if (action === 'getPendientes') result = getPendientes(p);
    else if (action === 'getEstadisticas') result = getEstadisticas(p);
    else if (action === 'registrarUsuario') result = registrarUsuario(p);
    else if (action === 'loginUsuario') result = loginUsuario(p);
    else if (action === 'enviarCita') result = enviarCita(p);
    else if (action === 'aprobarCita') result = aprobarCita(p);
    else if (action === 'rechazarCita') result = rechazarCita(p);
    else if (action === 'editarCita') result = editarCita(p);
    else if (action === 'eliminarCita') result = eliminarCita(p);
    else if (action === 'eliminarUsuario') result = eliminarUsuario(p);
    else if (action === 'restablecerPass') result = restablecerPass(p);
    else if (action === 'loginAdmin') result = loginAdmin(p);
    else if (action === 'registrarDescarga') result = registrarDescarga(p);
    else if (action === 'getPerfilUsuario') result = getPerfilUsuario(p);
    else if (action === 'getBlogPosts') result = getBlogPosts(p);
    else if (action === 'getBlogPost') result = getBlogPost(p);
    else if (action === 'crearBlogPost') result = crearBlogPost(p);
    else if (action === 'editarBlogPost') result = editarBlogPost(p);
    else if (action === 'eliminarBlogPost') result = eliminarBlogPost(p);
    else result = { ok: false, error: 'Accion no reconocida: ' + action };
  } catch(err) {
    result = { ok: false, error: err.toString() };
  }

  return makeResponse(result);
}

// Handler para peticiones OPTIONS (CORS preflight)
function doOptions(e) {
  return makeResponse({ ok: true, message: 'CORS OK' });
}

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
    if (obj['Cita'] && String(obj['Cita']).trim() !== '') {
      rows.push(obj);
    }
  }
  return { ok: true, data: rows };
}

function getCitasPublicas() {
  // Devuelve todas las citas para usuarios públicos (invitados)
  // Es la misma función que getCitas pero con otro nombre para mayor claridad
  return getCitas();
}

function loginAdmin(p) {
  var email = String(p.email || '').toLowerCase();
  var passHash = String(p.passHash || '');
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

function loginUsuario(p) {
  var email = String(p.email || '').toLowerCase();
  var passHash = String(p.passHash || '');
  
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(SHEET_USUARIOS);
  if (!sh) return { ok: false, error: 'No se encontro la hoja de usuarios' };
  
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][3]).toLowerCase() === email && String(data[i][8]) === passHash) {
      var raw = email + passHash + 'user_salt_2025';
      var token = Utilities.computeDigest(
        Utilities.DigestAlgorithm.SHA_256, raw, Utilities.Charset.UTF_8
      ).map(function(b){ return (b<0?b+256:b).toString(16).padStart(2,'0'); }).join('');
      PropertiesService.getScriptProperties().setProperty('usr_' + token, email);
      
      return { 
        ok: true, 
        token: token, 
        user: {
          id: data[i][0],
          nombre: data[i][1],
          apellido: data[i][2],
          email: data[i][3],
          institucion: data[i][4],
          area: data[i][5]
        }
      };
    }
  }
  return { ok: false, error: 'Credenciales incorrectas' };
}

function verificarAdmin(token) {
  if (!token) return false;
  var val = PropertiesService.getScriptProperties().getProperty('adm_' + token);
  return !!val;
}

function verificarUsuario(token) {
  if (!token) return null;
  var email = PropertiesService.getScriptProperties().getProperty('usr_' + token);
  return email || null;
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

function registrarUsuario(p) {
  ensureSheet(SHEET_USUARIOS, ['ID','Nombre','Apellido','Email','Institucion','Area','Motivo','Fecha','Contrasena','Estado','TotalDescargas']);
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(SHEET_USUARIOS);
  var data = sh.getDataRange().getValues();
  
  var email = String(p.email || '').toLowerCase().trim();
  
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][3]).toLowerCase() === email) {
      return { ok: false, error: 'Este correo ya esta registrado' };
    }
  }
  
  var id = 'U' + new Date().getTime();
  var tempPassword = generarPasswordTemporal();
  var passHash = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256, tempPassword, Utilities.Charset.UTF_8
  ).map(function(b){ return (b < 0 ? b+256 : b).toString(16).padStart(2,'0'); }).join('');
  
  sh.appendRow([
    id, 
    p.nombre || '', 
    p.apellido || '', 
    email,
    p.institucion || '', 
    p.area || '', 
    p.motivo || '',
    new Date().toISOString(), 
    passHash, 
    'ACTIVO',
    0
  ]);
  
  // Enviar correo de bienvenida
  enviarCorreoBienvenida(email, p.nombre || '', tempPassword);
  
  return { ok: true, id: id, msg: 'Usuario registrado. Se envio un correo con tus credenciales.' };
}

function generarPasswordTemporal() {
  var chars = 'ABCDEFGHJKMNPQRSTUVWXYZabcdefghjkmnpqrstuvwxyz23456789';
  var password = 'SS';
  for (var i = 0; i < 6; i++) {
    password += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  password += '!';
  return password;
}

function enviarCorreoBienvenida(email, nombre, password) {
  try {
    var subject = 'Bienvenido a Sintropia Social - Tus credenciales de acceso';
    var body = 'Hola ' + nombre + ',\n\n' +
      'Bienvenido/a a Sintropia Social, tu repositorio bibliografico de ciencias sociales.\n\n' +
      'Tus credenciales de acceso son:\n' +
      '------------------------------------\n' +
      'Usuario (email): ' + email + '\n' +
      'Contrasena: ' + password + '\n' +
      '------------------------------------\n\n' +
      'Te recomendamos cambiar tu contrasena despues de iniciar sesion.\n\n' +
      'Accede al repositorio en: https://sintropiassociales-del.github.io/Investigacion-social/\n\n' +
      'Gracias por unirte a nuestra comunidad.\n\n' +
      'Atentamente,\n' +
      'Equipo Sintropia Social\n' +
      'contacto@sintropiasocial.com';
    
    var htmlBody = '<div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">' +
      '<div style="background: #1a3a2f; padding: 20px; text-align: center;">' +
      '<h1 style="color: #ffffff; margin: 0;">Sintropia Social</h1>' +
      '</div>' +
      '<div style="padding: 30px; background: #faf9f7;">' +
      '<h2 style="color: #1a3a2f;">Hola ' + nombre + ',</h2>' +
      '<p>Bienvenido/a a <strong>Sintropia Social</strong>, tu repositorio bibliografico de ciencias sociales.</p>' +
      '<div style="background: #ffffff; border: 1px solid #e0ddd5; border-radius: 8px; padding: 20px; margin: 20px 0;">' +
      '<h3 style="color: #1a3a2f; margin-top: 0;">Tus credenciales de acceso:</h3>' +
      '<p><strong>Usuario (email):</strong> ' + email + '</p>' +
      '<p><strong>Contrasena:</strong> <code style="background: #f0efe9; padding: 4px 8px; border-radius: 4px;">' + password + '</code></p>' +
      '</div>' +
      '<p style="color: #666;">Te recomendamos cambiar tu contrasena despues de iniciar sesion.</p>' +
      '<p style="text-align: center; margin-top: 30px;">' +
      '<a href="https://sintropiassociales-del.github.io/Investigacion-social/" style="background: #1a3a2f; color: #ffffff; padding: 12px 24px; text-decoration: none; border-radius: 6px;">Acceder al Repositorio</a>' +
      '</p>' +
      '</div>' +
      '<div style="background: #1a3a2f; padding: 15px; text-align: center;">' +
      '<p style="color: #ffffff; margin: 0; font-size: 12px;">Equipo Sintropia Social | contacto@sintropiasocial.com</p>' +
      '</div>' +
      '</div>';
    
    MailApp.sendEmail({
      to: email,
      subject: subject,
      body: body,
      htmlBody: htmlBody
    });
    
    return true;
  } catch(e) {
    Logger.log('Error enviando correo: ' + e.toString());
    return false;
  }
}

function registrarDescarga(p) {
  var userEmail = verificarUsuario(p.userToken);
  if (!userEmail) return { ok: false, error: 'Usuario no autenticado' };
  
  ensureSheet(SHEET_DESCARGAS, ['ID','Email','Fecha','TipoPago','Monto','Filtros','CantidadCitas']);
  
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var shDescargas = ss.getSheetByName(SHEET_DESCARGAS);
  var shUsuarios = ss.getSheetByName(SHEET_USUARIOS);
  
  // Registrar descarga
  var descargaId = 'D' + new Date().getTime();
  shDescargas.appendRow([
    descargaId,
    userEmail,
    new Date().toISOString(),
    p.tipoPago || 'unico',
    p.monto || '18',
    p.filtros || '',
    p.cantidadCitas || 0
  ]);
  
  // Actualizar contador en perfil de usuario
  var dataUsuarios = shUsuarios.getDataRange().getValues();
  for (var i = 1; i < dataUsuarios.length; i++) {
    if (String(dataUsuarios[i][3]).toLowerCase() === userEmail.toLowerCase()) {
      var totalActual = parseInt(dataUsuarios[i][10]) || 0;
      shUsuarios.getRange(i + 1, 11).setValue(totalActual + 1);
      break;
    }
  }
  
  return { ok: true, descargaId: descargaId, msg: 'Descarga registrada' };
}

function getPerfilUsuario(p) {
  var userEmail = verificarUsuario(p.userToken);
  if (!userEmail) return { ok: false, error: 'Usuario no autenticado' };
  
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var shUsuarios = ss.getSheetByName(SHEET_USUARIOS);
  var shDescargas = ss.getSheetByName(SHEET_DESCARGAS);
  
  // Obtener datos del usuario
  var dataUsuarios = shUsuarios.getDataRange().getValues();
  var usuario = null;
  for (var i = 1; i < dataUsuarios.length; i++) {
    if (String(dataUsuarios[i][3]).toLowerCase() === userEmail.toLowerCase()) {
      usuario = {
        id: dataUsuarios[i][0],
        nombre: dataUsuarios[i][1],
        apellido: dataUsuarios[i][2],
        email: dataUsuarios[i][3],
        institucion: dataUsuarios[i][4],
        area: dataUsuarios[i][5],
        fechaRegistro: dataUsuarios[i][7],
        totalDescargas: dataUsuarios[i][10] || 0
      };
      break;
    }
  }
  
  if (!usuario) return { ok: false, error: 'Usuario no encontrado' };
  
  // Obtener historial de descargas
  var descargas = [];
  if (shDescargas) {
    var dataDescargas = shDescargas.getDataRange().getValues();
    for (var j = 1; j < dataDescargas.length; j++) {
      if (String(dataDescargas[j][1]).toLowerCase() === userEmail.toLowerCase()) {
        descargas.push({
          id: dataDescargas[j][0],
          fecha: dataDescargas[j][2],
          tipoPago: dataDescargas[j][3],
          monto: dataDescargas[j][4],
          cantidadCitas: dataDescargas[j][6]
        });
      }
    }
  }
  
  return { ok: true, usuario: usuario, descargas: descargas };
}

function getEstadisticas(p) {
  if (!verificarAdmin(p.adminToken)) return { ok: false, error: 'No autorizado' };
  
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var shDescargas = ss.getSheetByName(SHEET_DESCARGAS);
  var shUsuarios = ss.getSheetByName(SHEET_USUARIOS);
  
  var totalDescargas = 0;
  var ingresoTotal = 0;
  var descargasPorMes = {};
  
  if (shDescargas) {
    var data = shDescargas.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      totalDescargas++;
      ingresoTotal += parseFloat(data[i][4]) || 0;
      
      var fecha = new Date(data[i][2]);
      var mes = fecha.getFullYear() + '-' + String(fecha.getMonth() + 1).padStart(2, '0');
      descargasPorMes[mes] = (descargasPorMes[mes] || 0) + 1;
    }
  }
  
  var totalUsuarios = 0;
  if (shUsuarios) {
    totalUsuarios = shUsuarios.getLastRow() - 1;
  }
  
  return {
    ok: true,
    estadisticas: {
      totalDescargas: totalDescargas,
      ingresoTotal: ingresoTotal,
      totalUsuarios: totalUsuarios,
      descargasPorMes: descargasPorMes
    }
  };
}

function enviarCita(p) {
  ensureSheet(SHEET_PENDIENTES, ['No','Categoria','Indicador','Poblacion','Anio','Autor','Cita','Comentarios','Publicacion','Pagina','Cita Apa','Link','Usuario','Fecha','Estado']);
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(SHEET_PENDIENTES);
  sh.appendRow([
    '', p.categoria || '', p.indicador || '', p.poblacion || '',
    p.year || '', p.autor || '', p.cita || '', p.comentarios || '',
    p.publicacion || '', p.pagina || '', p.citaAPA || '', p.link || '',
    p.usuarioEmail || '', new Date().toISOString(), 'PENDIENTE'
  ]);
  return { ok: true, msg: 'Cita enviada para revision' };
}

function aprobarCita(p) {
  if (!verificarAdmin(p.adminToken)) return { ok: false, error: 'No autorizado' };
  var rowIndex = parseInt(p.rowIndex);
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var pend = ss.getSheetByName(SHEET_PENDIENTES);
  var citas = ss.getSheetByName(SHEET_CITAS);
  var data = pend.getDataRange().getValues();
  var row = data[rowIndex + 1];
  if (!row) return { ok: false, error: 'Fila no encontrada' };
  var newId = 'C' + String(citas.getLastRow()).padStart(4, '0');
  citas.appendRow([newId, row[1], row[2], row[3], row[4], row[5], row[6], row[7], '', row[8], row[9], row[10], row[11]]);
  pend.getRange(rowIndex + 2, 15).setValue('APROBADA');
  return { ok: true, msg: 'Cita aprobada y publicada' };
}

function rechazarCita(p) {
  if (!verificarAdmin(p.adminToken)) return { ok: false, error: 'No autorizado' };
  var rowIndex = parseInt(p.rowIndex);
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var pend = ss.getSheetByName(SHEET_PENDIENTES);
  pend.getRange(rowIndex + 2, 15).setValue('RECHAZADA: ' + (p.motivo || ''));
  return { ok: true, msg: 'Cita rechazada' };
}

function editarCita(p) {
  if (!verificarAdmin(p.adminToken)) return { ok: false, error: 'No autorizado' };
  var rowNum = parseInt(p.rowNum);
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(SHEET_CITAS);
  sh.getRange(rowNum, 2).setValue(p.categoria || '');
  sh.getRange(rowNum, 3).setValue(p.indicador || '');
  sh.getRange(rowNum, 4).setValue(p.poblacion || '');
  sh.getRange(rowNum, 5).setValue(p.year || '');
  sh.getRange(rowNum, 6).setValue(p.autor || '');
  sh.getRange(rowNum, 7).setValue(p.cita || '');
  sh.getRange(rowNum, 8).setValue(p.comentarios || '');
  sh.getRange(rowNum, 11).setValue(p.pagina || '');
  sh.getRange(rowNum, 12).setValue(p.citaAPA || '');
  sh.getRange(rowNum, 13).setValue(p.link || '');
  return { ok: true, msg: 'Cita actualizada' };
}

function eliminarCita(p) {
  if (!verificarAdmin(p.adminToken)) return { ok: false, error: 'No autorizado' };
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(SHEET_CITAS);
  sh.deleteRow(parseInt(p.rowNum));
  return { ok: true, msg: 'Cita eliminada' };
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
  var tempPass = generarPasswordTemporal();
  var hash = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256, tempPass, Utilities.Charset.UTF_8
  ).map(function(b){ return (b < 0 ? b+256 : b).toString(16).padStart(2,'0'); }).join('');
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(SHEET_USUARIOS);
  sh.getRange(parseInt(p.rowNum), 9).setValue(hash);
  
  // Enviar correo con nueva contrasena
  if (p.email) {
    try {
      MailApp.sendEmail({
        to: p.email,
        subject: 'Sintropia Social - Nueva contrasena',
        body: 'Hola,\n\nTu contrasena ha sido restablecida.\n\nNueva contrasena: ' + tempPass + '\n\nPor favor cambiala despues de iniciar sesion.\n\nSaludos,\nEquipo Sintropia Social'
      });
    } catch(e) {}
  }
  
  return { ok: true, msg: 'Contrasena restablecida y enviada por correo' };
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


// ================== BLOG FUNCTIONS ==================

function getBlogPosts(p) {
  ensureSheet(SHEET_BLOG, ['ID','Titulo','Categoria','Contenido','Imagen','Autor','Fecha','Estado']);
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(SHEET_BLOG);
  var data = sh.getDataRange().getValues();
  
  if (data.length < 2) return { ok: true, data: [] };
  
  var headers = data[0];
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      obj[String(headers[j]).toLowerCase()] = data[i][j];
    }
    obj._row = i + 1;
    // Solo mostrar publicados para usuarios normales, todos para admin
    if (p.adminToken) {
      rows.push(obj);
    } else if (obj.estado === 'PUBLICADO') {
      rows.push(obj);
    }
  }
  return { ok: true, data: rows.reverse() };
}

function getBlogPost(p) {
  var postId = p.id;
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(SHEET_BLOG);
  if (!sh) return { ok: false, error: 'Blog no encontrado' };
  
  var data = sh.getDataRange().getValues();
  var headers = data[0];
  
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(postId)) {
      var obj = {};
      for (var j = 0; j < headers.length; j++) {
        obj[String(headers[j]).toLowerCase()] = data[i][j];
      }
      return { ok: true, post: obj };
    }
  }
  return { ok: false, error: 'Post no encontrado' };
}

function crearBlogPost(p) {
  if (!verificarAdmin(p.adminToken)) return { ok: false, error: 'No autorizado' };
  
  ensureSheet(SHEET_BLOG, ['ID','Titulo','Categoria','Contenido','Imagen','Autor','Fecha','Estado']);
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(SHEET_BLOG);
  
  sh.appendRow([
    p.id || 'B' + new Date().getTime(),
    p.titulo || '',
    p.categoria || 'General',
    p.contenido || '',
    p.imagen || '',
    p.autor || 'Admin',
    new Date().toISOString(),
    p.estado || 'BORRADOR'
  ]);
  
  return { ok: true, msg: 'Entrada de blog creada' };
}

function editarBlogPost(p) {
  if (!verificarAdmin(p.adminToken)) return { ok: false, error: 'No autorizado' };
  
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(SHEET_BLOG);
  if (!sh) return { ok: false, error: 'Blog no encontrado' };
  
  var data = sh.getDataRange().getValues();
  
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(p.id)) {
      sh.getRange(i + 1, 2).setValue(p.titulo || '');
      sh.getRange(i + 1, 3).setValue(p.categoria || 'General');
      sh.getRange(i + 1, 4).setValue(p.contenido || '');
      sh.getRange(i + 1, 5).setValue(p.imagen || '');
      sh.getRange(i + 1, 8).setValue(p.estado || 'BORRADOR');
      return { ok: true, msg: 'Entrada actualizada' };
    }
  }
  
  return { ok: false, error: 'Post no encontrado' };
}

function eliminarBlogPost(p) {
  if (!verificarAdmin(p.adminToken)) return { ok: false, error: 'No autorizado' };
  
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(SHEET_BLOG);
  
  if (p.rowNum) {
    sh.deleteRow(parseInt(p.rowNum));
    return { ok: true, msg: 'Entrada eliminada' };
  }
  
  return { ok: false, error: 'Fila no especificada' };
}
