// ============================================================
//  SINTROPÍA SOCIAL — nav.js
//  Barra de navegación, modal de login y toasts compartidos.
//
//  Extraído de los 10 archivos HTML del sitio (la barra <nav>,
//  el modal de login, el helper de toasts y el objeto SS_AUTH
//  eran 100% idénticos byte a byte en las 10 páginas).
//
//  Depende de config.js (debe cargarse ANTES que este archivo):
//  usa CONFIG, Auth, sha256() y api() de config.js.
//
//  Uso en cada página:
//    <script src="config.js"></script>
//    <script src="nav.js"></script>
//    ...
//    <div id="sintro-nav"></div>
//    <script>SintroNav.render('#sintro-nav', 'inicio');</script>
// ============================================================

// ── AUTH SEGURO — TTL 8h ──
// (Este objeto SS_AUTH estaba duplicado, byte por byte, en las
// 10 páginas. Es el que de verdad controla el estado de sesión
// que se muestra en la barra de navegación y en el modal de
// login. El objeto `Auth` de config.js es un sistema aparte,
// más antiguo, que usan algunas páginas para su propia lógica
// de negocio (p.ej. admin.html) — no lo tocamos.)
var SS_AUTH = (function(){
  var TTL = 8 * 60 * 60 * 1000; // 8 horas en ms
  var USER_KEY = 'ss_user_v2';
  var ADMIN_KEY = 'ss_admin_v2';

  function _set(key, data){
    try{
      localStorage.setItem(key, JSON.stringify({
        data: data,
        expires: Date.now() + TTL
      }));
    }catch(e){}
  }

  function _get(key){
    try{
      var raw = JSON.parse(localStorage.getItem(key)||'null');
      if(!raw) return null;
      if(Date.now() > raw.expires){
        localStorage.removeItem(key);
        return null;
      }
      return raw.data;
    }catch(e){ return null; }
  }

  function _clear(key){ try{ localStorage.removeItem(key); }catch(e){} }

  return {
    setUser:  function(u){ _set(USER_KEY, u); },
    getUser:  function(){ return _get(USER_KEY); },
    setAdmin: function(a){ _set(ADMIN_KEY, a); },
    getAdmin: function(){ return _get(ADMIN_KEY); },
    isAdmin:  function(){ var a=_get(ADMIN_KEY); return !!(a&&a.token); },
    getToken: function(){ var a=_get(ADMIN_KEY); return (a&&a.token)||null; },
    logout:   function(){
      _clear(USER_KEY); _clear(ADMIN_KEY);
      // Also clear old keys for migration
      ['ss_user','ss_admin','sintropia_usuario_actual'].forEach(function(k){ try{ localStorage.removeItem(k); }catch(e){} });
      window.location.href = 'inicio.html';
    },
    // Migrate from old storage
    migrate: function(){
      try{
        var oldUser = JSON.parse(localStorage.getItem('ss_user')||localStorage.getItem('sintropia_usuario_actual')||'null');
        var oldAdmin = JSON.parse(localStorage.getItem('ss_admin')||'null');
        if(oldUser && !_get(USER_KEY)){ _set(USER_KEY, oldUser); localStorage.removeItem('ss_user'); localStorage.removeItem('sintropia_usuario_actual'); }
        if(oldAdmin && !_get(ADMIN_KEY)){ _set(ADMIN_KEY, oldAdmin); localStorage.removeItem('ss_admin'); }
      }catch(e){}
    }
  };
})();
// Run migration on load
SS_AUTH.migrate();

// ── Nav bar markup (idéntico en las 10 páginas) ──
// Se añade un ítem extra "Área interna" (id="ss-nav-internal-li"),
// oculto por defecto, que solo se muestra cuando Auth.isAdmin()
// (de config.js) es true — enlaza a /plataforma/, fuera del
// alcance de este trabajo.
var SINTRO_NAV_HTML = `
<nav class="ss-nav">
  <a href="inicio.html" class="ss-nav-logo">
    <img src="logo.png" alt="Sintropía Social" onerror="this.style.display='none'"/>
    Sintropía Social
  </a>
  <ul class="ss-nav-links" id="ss-nav-links">
    <li><a href="inicio.html" data-navkey="inicio">Inicio</a></li>
    <li><a href="index.html" data-navkey="index">Repositorio</a></li>
    <li>
      <div class="ss-dd" id="dd-herr">
        <button class="ss-dd-btn" onclick="ssDD('dd-herr')" data-navkey="herramientas">Herramientas
          <svg viewBox="0 0 10 6" fill="none"><path d="M1 1l4 4 4-4" stroke="currentColor" stroke-width="1.5" stroke-linecap="round"/></svg>
        </button>
        <div class="ss-dd-panel">
          <a href="diagnostico.html" data-navkey="diagnostico">Diagnostico Social</a>
          <a href="dashboard.html" data-navkey="dashboard">Dashboard Interactivo</a>
          <a href="pdfcitas.html" data-navkey="pdfcitas">Buscador de citas PDF</a>
        </div>
      </div>
    </li>
    <li><a href="servicios.html" data-navkey="servicios">Servicios</a></li>
    <li><a href="blog.html" data-navkey="blog">Newsletter</a></li>
    <li><a href="contribuir.html" data-navkey="contribuir">Contribuir</a></li>
    <li id="ss-nav-internal-li" style="display:none;"><a href="/plataforma/" data-navkey="admin-interna">Área interna</a></li>
  </ul>
  <div class="ss-nav-end">
    <span id="ss-nav-user">
      <a href="#" onclick="ssAbrirLogin();return false;" class="btn btn-ghost btn-sm">Entrar</a>
      <a href="registro.html" class="btn btn-primary btn-sm">Unete gratis</a>
    </span>
    <button class="ss-hamburger" onclick="ssMobileNav()" aria-label="Menu">
      <span></span><span></span><span></span>
    </button>
  </div>
</nav>
`;

// ── Login modal + toast markup (idéntico en las 10 páginas) ──
var SINTRO_MODAL_HTML = `
<div class="ss-modal-bg" id="ss-login-modal" onclick="if(event.target===this)ssCerrarLogin()">
  <div class="ss-modal">
    <button class="ss-modal-close" onclick="ssCerrarLogin()">x</button>
    <h3>Iniciar sesion</h3>
    <p>Accede a tu cuenta para ver el repositorio completo.</p>
    <div id="ss-li-err" style="display:none;background:var(--err-bg);color:var(--err);padding:.6rem .9rem;border-radius:var(--r-sm);margin-bottom:1rem;font-size:.83rem;"></div>
    <div class="campo" style="margin-bottom:.85rem;"><label>Correo</label><input type="email" id="ss-li-email" placeholder="correo@ejemplo.com" onkeydown="if(event.key==='Enter')ssDoLogin()"/></div>
    <div class="campo" style="margin-bottom:1.25rem;"><label>Contrasena</label><input type="password" id="ss-li-pass" placeholder="&bull;&bull;&bull;&bull;&bull;&bull;&bull;&bull;" onkeydown="if(event.key==='Enter')ssDoLogin()"/></div>
    <button id="ss-li-btn" class="btn btn-primary" style="width:100%;" onclick="ssDoLogin()">Entrar</button>
    <p style="text-align:center;font-size:.83rem;margin-top:1rem;color:var(--tx-light);">No tienes cuenta? <a href="registro.html">Registrate gratis</a></p>
  </div>
</div>
<div id="ss-toast" class="ss-toast"></div>
`;

// ── Dropdown + mobile toggle ──
function ssDD(id){
  document.querySelectorAll('.ss-dd').forEach(function(d){ if(d.id!==id) d.classList.remove('open'); });
  document.getElementById(id).classList.toggle('open');
}
function ssMobileNav(){ document.getElementById('ss-nav-links').classList.toggle('open'); }
document.addEventListener('click',function(e){ if(!e.target.closest('.ss-dd')) document.querySelectorAll('.ss-dd').forEach(function(d){d.classList.remove('open');}); });

// ── AUTH ──
function ssGetUser(){ return SS_AUTH.getUser() || (function(){ try{ return JSON.parse(localStorage.getItem("ss_user")||"null"); }catch(e){return null;} })(); }
function ssSetUser(u){ SS_AUTH.setUser(u); }
function ssLogout(){ SS_AUTH.logout(); }

// ── LOGIN MODAL ──
function ssAbrirLogin(){
  var m = document.getElementById('ss-login-modal');
  if(m) m.classList.add('open');
  else window.location.href='registro.html';
}
function ssCerrarLogin(){ var m=document.getElementById('ss-login-modal'); if(m) m.classList.remove('open'); }

// Tracks which page called SintroNav.render(), so ssDoLogin can
// reproduce the one real per-page difference found in the 10
// originals: on admin.html the login modal does NOT also try to
// fetch a separate admin token (admin.html already has its own
// dedicated admin-login flow elsewhere on the page); on every
// other page it does, to bridge the session into config.js's
// Auth (legacy 'ss_admin' key) so admin.html works right away
// if the user navigates there next.
var _sintroActivePage = null;

async function ssDoLogin(){
  var email = (document.getElementById('ss-li-email')||{}).value||'';
  var pass  = (document.getElementById('ss-li-pass')||{}).value||'';
  var err   = document.getElementById('ss-li-err');
  if(!email||!pass){ if(err){err.textContent='Completa todos los campos.';err.style.display='block';} return; }
  var btn = document.getElementById('ss-li-btn');
  if(btn){ btn.textContent='Entrando...'; btn.disabled=true; }
  try{
    var h = await sha256(pass);
    var r = await api('loginUsuario',{email:email,passHash:h});
    if(r&&r.ok){
      ssSetUser({ token:r.token||('u_'+Date.now()), nombre:r.nombre||'', email:r.email||email, isAdmin:r.isAdmin||false });
      if(_sintroActivePage === 'admin'){
        ssCerrarLogin();
        if(r.isAdmin) window.location.href='admin.html';
        else location.reload();
      } else {
        // Also try admin login to get admin token
        if(r.isAdmin || (r.token && r.token.indexOf('admin_')===0)){
          try{
            var ar = await api('loginAdmin',{email:email,passHash:h});
            if(ar&&ar.ok){
              localStorage.setItem('ss_admin', JSON.stringify({token:ar.token,email:ar.email}));
            }
          }catch(e){}
          ssCerrarLogin();
          window.location.href='admin.html';
        } else {
          ssCerrarLogin();
          location.reload();
        }
      }
    } else {
      if(err){ err.textContent=(r&&r.error)||'Correo o contrasena incorrectos.'; err.style.display='block'; }
      if(btn){ btn.textContent='Entrar'; btn.disabled=false; }
    }
  }catch(e){
    if(err){ err.textContent='Error de conexion.'; err.style.display='block'; }
    if(btn){ btn.textContent='Entrar'; btn.disabled=false; }
  }
}

// ── TOAST ──
function ssToast(msg,type){
  var t=document.getElementById('ss-toast');
  if(!t){ t=document.createElement('div'); t.id='ss-toast'; t.className='ss-toast'; document.body.appendChild(t); }
  t.textContent=msg; t.className='ss-toast show'+(type?' '+type:'');
  setTimeout(function(){ t.classList.remove('show'); },3200);
}

// ── Nav user-state rendering ──
// (Antes era una IIFE anónima inline; aquí es una función con
// nombre para poder llamarla explícitamente desde render().)
function _sintroRenderUserState(){
  var u = SS_AUTH.getUser() || ssGetUser();
  var el = document.getElementById('ss-nav-user');
  if(el && u && (u.nombre||u.email)){
    var label = u.nombre || u.email.split('@')[0];
    el.innerHTML = '<button onclick="ssLogout()" style="background:var(--morado);color:var(--blanco);padding:.42rem .95rem;border-radius:var(--r-pill);border:none;cursor:pointer;font-family:var(--font-body);font-size:.82rem;font-weight:600;">'+label+' · Salir</button>';
  }
  // "Área interna" — solo visible para administradores (Auth.isAdmin(),
  // el mecanismo real de admin del sitio, ya usado en admin.html).
  var internalLi = document.getElementById('ss-nav-internal-li');
  if(internalLi){
    var isAdm = false;
    try{ isAdm = (typeof Auth!=='undefined') && Auth.isAdmin(); }catch(e){}
    internalLi.style.display = isAdm ? '' : 'none';
  }
}

// ── Active-link marking ──
// No existía ningún marcado "active" en el nav original de las
// 10 páginas; se añade aquí como mejora nueva y no destructiva.
function _sintroMarkActive(key){
  if(!key) return;
  var links = document.querySelectorAll('[data-navkey="'+key+'"]');
  links.forEach(function(link){
    link.classList.add('active');
    if(link.tagName === 'BUTTON'){
      var dd = link.closest('.ss-dd');
      if(dd) dd.classList.add('active-parent');
    }
  });
}

// ── Punto de entrada público ──
var SintroNav = {
  render: function(mountSelector, activeKey){
    var mount = document.querySelector(mountSelector);
    if(!mount) return;
    _sintroActivePage = activeKey;
    mount.innerHTML = SINTRO_NAV_HTML;
    document.body.insertAdjacentHTML('beforeend', SINTRO_MODAL_HTML);
    _sintroMarkActive(activeKey);
    _sintroRenderUserState();
  }
};
