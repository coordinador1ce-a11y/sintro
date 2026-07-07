// ============================================================
//  SINTROPÍA SOCIAL — config.js v3 (JSONP para Apps Script)
// ============================================================

var CONFIG = {
  API_URL: 'https://script.google.com/macros/s/AKfycbxnBy1DKH14oWGEnzFHLDrD0XlgshVkBJrCL5b7zp8XKO0hmE4xriXEDXU9PZZl5KzovQ/exec',
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
    try { return JSON.parse(localStorage.getItem('ss_user')); } catch(e) { return null; }
  },
  getAdmin: function() {
    try { return JSON.parse(localStorage.getItem('ss_admin')); } catch(e) { return null; }
  },
  setUser: function(u) { localStorage.setItem('ss_user', JSON.stringify(u)); },
  setAdmin: function(a) { localStorage.setItem('ss_admin', JSON.stringify(a)); },
  logout: function() { localStorage.removeItem('ss_user'); location.href = 'index.html'; },
  logoutAdmin: function() { localStorage.removeItem('ss_admin'); location.reload(); },
  isAdmin: function() {
    var a = Auth.getAdmin();
    return !!(a && a.token);
  },
  getToken: function() {
    var a = Auth.getAdmin();
    return (a && a.token) ? a.token : null;
  }
};

// ── SHA-256 ──
async function sha256(str) {
  var buf = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(str));
  return Array.from(new Uint8Array(buf)).map(function(b) {
    return b.toString(16).padStart(2, '0');
  }).join('');
}

// ── API via JSONP — único método que funciona con Apps Script desde GitHub Pages ──
function api(action, params) {
  return new Promise(function(resolve) {
    // Nombre único para el callback
    var cbName = 'ss_cb_' + Date.now() + '_' + Math.floor(Math.random() * 10000);

    // Timeout de 15 segundos
    var timer = setTimeout(function() {
      cleanup();
      resolve({ ok: false, error: 'Tiempo de espera agotado. Verifica que Apps Script esté desplegado.' });
    }, 15000);

    function cleanup() {
      clearTimeout(timer);
      delete window[cbName];
      var el = document.getElementById(cbName);
      if (el) el.parentNode.removeChild(el);
    }

    // Registrar callback global
    window[cbName] = function(data) {
      cleanup();
      resolve(data);
    };

    // Construir URL con todos los parámetros + callback
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

    // Inyectar script tag (esto evita CORS)
    var script = document.createElement('script');
    script.id  = cbName;
    script.src = CONFIG.API_URL + '?' + p.toString();
    script.onerror = function() {
      cleanup();
      resolve({ ok: false, error: 'Error al conectar con Apps Script. Verifica la URL.' });
    };
    document.head.appendChild(script);
  });
}
