// ============================================================
//  SINTROPÍA SOCIAL — config.js v4
//  Compatible con Apps Script que responde JSON (no JSONP)
//  Usa fetch + no-cors workaround vía URL con parámetros GET
// ============================================================

var CONFIG = {
  API_URL: 'https://script.google.com/macros/s/AKfycbxF7ntGOXPufT6nV9XSH7WJT6uTnIn8R5402K6b1B3uA6unQ2RgrcQtOgiBArGXQSugAA/exec',
  SHEET_ID: '114sl6Mt-UhQQsv7zyicAAmsYzo3VDPoAvbT-0MakK94',
  GUEST_PERCENT: 0.10,
  CONTACT_EMAIL: 'contacto@sintropiasocial.com',
  ADMIN_EMAILS: ['dsalgado@sintropiasocial.com'],
  PAYPAL_CLIENT_ID: 'BAADNWafE2xUH09mKvDiejlkmXxK9XQx1oa-ujzF7TF-pQNLf1a58OhHRUMUNoDx9dgXzhDclHdQhukdW0',
  PAYPAL_BUTTON_ID: 'RY5K7VHYRPJLY'
};

// ── Auth helpers ──
var Auth = {
  getUser: function() {
    try {
      return JSON.parse(localStorage.getItem('ss_user') ||
             localStorage.getItem('sintropia_usuario_actual') || 'null');
    } catch(e) { return null; }
  },
  getAdmin: function() {
    try { return JSON.parse(localStorage.getItem('ss_admin') || 'null'); } catch(e) { return null; }
  },
  setUser:     function(u) { localStorage.setItem('ss_user', JSON.stringify(u)); },
  setAdmin:    function(a) { localStorage.setItem('ss_admin', JSON.stringify(a)); },
  logout:      function() { localStorage.removeItem('ss_user'); localStorage.removeItem('sintropia_usuario_actual'); location.href = 'index.html'; },
  logoutAdmin: function() { localStorage.removeItem('ss_admin'); location.reload(); },
  isAdmin:     function() { var a = Auth.getAdmin(); return !!(a && a.token); },
  getToken:    function() { var a = Auth.getAdmin(); return (a && a.token) ? a.token : null; }
};

// ── SHA-256 (nativo del browser) ──
async function sha256(str) {
  var buf = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(str));
  return Array.from(new Uint8Array(buf)).map(function(b) {
    return b.toString(16).padStart(2, '0');
  }).join('');
}

// ── API: intenta JSONP primero, si falla intenta fetch directo ──
function api(action, params) {
  return new Promise(function(resolve) {

    // ── MÉTODO 1: JSONP (funciona con Apps Script + GitHub Pages sin CORS) ──
    var cbName = 'ss_cb_' + Date.now() + '_' + Math.floor(Math.random() * 99999);
    var timer = setTimeout(function() {
      cleanup();
      // Si JSONP falla, intenta fetch directo como fallback
      apiFetch(action, params).then(resolve);
    }, 12000);

    function cleanup() {
      clearTimeout(timer);
      delete window[cbName];
      var el = document.getElementById(cbName);
      if (el && el.parentNode) el.parentNode.removeChild(el);
    }

    window[cbName] = function(data) {
      cleanup();
      resolve(data);
    };

    var p = new URLSearchParams();
    p.append('action', action);
    p.append('callback', cbName);
    if (params) {
      Object.keys(params).forEach(function(k) {
        if (params[k] !== null && params[k] !== undefined) {
          p.append(k, String(params[k]));
        }
      });
    }

    var script = document.createElement('script');
    script.id = cbName;
    script.src = CONFIG.API_URL + '?' + p.toString();
    script.onerror = function() {
      cleanup();
      // Fallback a fetch
      apiFetch(action, params).then(resolve);
    };
    document.head.appendChild(script);
  });
}

// ── FALLBACK: fetch directo (funciona si Apps Script tiene CORS abierto) ──
function apiFetch(action, params) {
  var p = new URLSearchParams();
  p.append('action', action);
  if (params) {
    Object.keys(params).forEach(function(k) {
      if (params[k] !== null && params[k] !== undefined) {
        p.append(k, String(params[k]));
      }
    });
  }
  var url = CONFIG.API_URL + '?' + p.toString();
  return fetch(url)
    .then(function(res) { return res.text(); })
    .then(function(text) {
      // Puede llegar como JSON puro o como JSONP envuelto
      var clean = text.trim();
      // Si viene envuelto en callback(...) lo limpiamos
      var match = clean.match(/^[a-zA-Z_$][a-zA-Z0-9_$]*\(([\s\S]*)\)\s*;?\s*$/);
      if (match) clean = match[1];
      return JSON.parse(clean);
    })
    .catch(function(e) {
      console.error('apiFetch error:', e);
      return { ok: false, error: 'Error de conexión con Apps Script. Verifica que esté desplegado como "Cualquier usuario".' };
    });
}

// ── SESSION SECURITY ──
// Warn user 10 min before session expires
(function(){
  function checkSessionExpiry(){
    try{
      var keys = ['ss_user_v2','ss_admin_v2'];
      keys.forEach(function(k){
        var raw = JSON.parse(localStorage.getItem(k)||'null');
        if(!raw || !raw.expires) return;
        var remaining = raw.expires - Date.now();
        if(remaining < 600000 && remaining > 0){ // less than 10 min
          console.info('Sesión expira en', Math.ceil(remaining/60000), 'minutos');
        }
        if(remaining <= 0){
          localStorage.removeItem(k);
        }
      });
    }catch(e){}
  }
  setInterval(checkSessionExpiry, 60000); // check every minute
})();
