/* ═══════════════════════════════════════════════════════════════════
   Il Borgo · Dashboard — Autenticazione condivisa (login obbligatorio)

   - Gate d'ingresso full-screen: #app resta nascosto finché non c'è un token valido.
   - TokenManager: token in localStorage (condiviso tra index.html e affitti.html),
     rinnovo silenzioso (prompt:'') prima della scadenza, fallback al gate.
   - Nessun refresh token (impossibile senza backend): il gate è il fallback previsto.

   API pubblica → window.IlBorgoAuth
     .start()             avvio automatico al caricamento (chiamato da solo)
     .signIn()            → Promise<token>  (consenso interattivo, usato dal gate)
     .signOut()           revoca + torna al gate
     .ensureToken()       → Promise<token>  (cache o rinnovo silenzioso)
     .getToken()          token corrente o null (sincrono)
     .handleAuthError()   da chiamare su 401/403 di una API: rinnova o mostra il gate
     .onChange(cb)        cb(true|false) a ogni cambio di stato auth
     .isAuthed()
   ═══════════════════════════════════════════════════════════════════ */
(function (w, d) {
  'use strict';

  var CFG = w.ILBORGO_CONFIG;
  if (!CFG) { console.error('[auth] ILBORGO_CONFIG mancante — caricare config.js prima di auth.js'); return; }

  var SKEW_MS   = 5 * 60 * 1000;   // rinnova 5 min prima della scadenza
  var listeners = [];
  var tokenClient = null;
  var refreshTimer = null;
  var authed = false;

  var waiters  = [];               // richieste token in attesa del callback GIS
  var inFlight = false;

  // ── storage ───────────────────────────────────────────────────────
  function saveToken(tok, expiresInSec) {
    var exp = Date.now() + ((expiresInSec || 3600) * 1000);
    try {
      localStorage.setItem(CFG.STORAGE.token, tok);
      localStorage.setItem(CFG.STORAGE.tokenExp, String(exp));
    } catch (e) {}
    return exp;
  }
  function clearToken() {
    try {
      localStorage.removeItem(CFG.STORAGE.token);
      localStorage.removeItem(CFG.STORAGE.tokenExp);
    } catch (e) {}
  }
  function cachedToken() {
    try {
      var tok = localStorage.getItem(CFG.STORAGE.token);
      var exp = parseFloat(localStorage.getItem(CFG.STORAGE.tokenExp) || '0');
      if (tok && exp && Date.now() < exp - 10000) return { token: tok, exp: exp };
    } catch (e) {}
    return null;
  }

  // ── gapi wiring (se presente nella pagina) ────────────────────────
  function pushToGapi(tok) {
    if (w.gapi && w.gapi.client && typeof w.gapi.client.setToken === 'function') {
      try { w.gapi.client.setToken(tok ? { access_token: tok } : null); } catch (e) {}
    }
  }

  // ── stato / notifiche ─────────────────────────────────────────────
  function setAuthed(v) {
    authed = v;
    if (v) hideGate(); else showGate('login');
    listeners.forEach(function (cb) { try { cb(v); } catch (e) {} });
  }

  // ── GIS token client ──────────────────────────────────────────────
  function gisReady() {
    return !!(w.google && w.google.accounts && w.google.accounts.oauth2);
  }
  function waitForGis(cb) {
    if (gisReady()) return cb();
    var tries = 0;
    var iv = setInterval(function () {
      if (gisReady()) { clearInterval(iv); cb(); }
      else if (++tries > 80) { clearInterval(iv); settleWaiters(null, new Error('gis_unavailable')); }
    }, 150);
  }
  function ensureClient() {
    if (tokenClient) return;
    tokenClient = w.google.accounts.oauth2.initTokenClient({
      client_id: CFG.CLIENT_ID,
      scope: CFG.SCOPES,
      callback: function (resp) {
        if (resp && resp.access_token) {
          var exp = saveToken(resp.access_token, parseFloat(resp.expires_in));
          pushToGapi(resp.access_token);
          scheduleRefresh(exp);
          settleWaiters(resp.access_token);
          setAuthed(true);
        } else {
          settleWaiters(null, new Error((resp && resp.error) || 'no_token'));
        }
      },
      error_callback: function (err) {
        settleWaiters(null, new Error((err && err.type) || 'gis_error'));
      }
    });
  }
  function settleWaiters(token, err) {
    inFlight = false;
    var list = waiters; waiters = [];
    list.forEach(function (wt) { err ? wt.reject(err) : wt.resolve(token); });
  }

  // mode 'none'  → rinnovo silenzioso: nessun popup, error_callback se serve interazione
  // mode ''      → interattivo (bottone del gate): mostra account picker / consenso se necessario
  function requestToken(mode) {
    return new Promise(function (resolve, reject) {
      waiters.push({ resolve: resolve, reject: reject });
      if (inFlight) return;
      inFlight = true;
      waitForGis(function () {
        if (!inFlight) return;              // già risolto da timeout GIS
        try {
          ensureClient();
          tokenClient.requestAccessToken({ prompt: mode === 'none' ? 'none' : '' });
        } catch (e) {
          settleWaiters(null, e);
        }
      });
    });
  }

  function scheduleRefresh(exp) {
    if (refreshTimer) clearTimeout(refreshTimer);
    var delay = Math.max(15000, exp - Date.now() - SKEW_MS);
    refreshTimer = setTimeout(function () {
      requestToken('none').catch(function () {
        if (!cachedToken()) setAuthed(false);   // rinnovo fallito e token scaduto → gate
      });
    }, delay);
  }

  // ── avvio ─────────────────────────────────────────────────────────
  function start() {
    var c = cachedToken();
    if (c) {
      // ritorno rapido: mostra subito l'app, rinnova in background se vicino a scadenza
      pushToGapi(c.token);
      scheduleRefresh(c.exp);
      setAuthed(true);
      if (Date.now() > c.exp - SKEW_MS) requestToken('none').catch(function () {});
      return;
    }
    // Nessun token: mostra il gate senza alcuna richiesta automatica.
    // Il token client GIS (modello token) apre comunque una finestra quando non
    // c'è sessione, quindi la prima acquisizione deve partire da un gesto utente
    // (bottone del gate). Il rinnovo silenzioso (scheduleRefresh) parte solo
    // quando una sessione è già stata stabilita.
    showGate('login');
  }

  function signOut() {
    var c = cachedToken();
    try {
      if (c && gisReady()) w.google.accounts.oauth2.revoke(c.token, function () {});
    } catch (e) {}
    clearToken();
    pushToGapi(null);
    if (refreshTimer) clearTimeout(refreshTimer);
    setAuthed(false);
  }

  function handleAuthError() {
    clearToken();
    pushToGapi(null);
    return requestToken('none').catch(function (e) { setAuthed(false); throw e; });
  }

  // ── login gate ────────────────────────────────────────────────────
  var gateEl = null;
  function buildGate() {
    if (gateEl) return;
    var css = d.createElement('style');
    css.textContent =
      '#ilb-gate{position:fixed;inset:0;z-index:99999;background:linear-gradient(160deg,#1A1F2E,#0F1420);' +
      'display:flex;align-items:center;justify-content:center;padding:24px;' +
      "font-family:'DM Sans',system-ui,sans-serif;-webkit-font-smoothing:antialiased;}" +
      '#ilb-gate .box{max-width:340px;width:100%;text-align:center;}' +
      "#ilb-gate .eyebrow{font-family:'DM Mono',monospace;font-size:9px;letter-spacing:4px;" +
      'text-transform:uppercase;color:#8A92AA;margin-bottom:14px;}' +
      "#ilb-gate h1{font-family:'EB Garamond',Georgia,serif;font-weight:400;font-size:27px;" +
      'color:#E8EAF0;margin:0 0 8px;}' +
      '#ilb-gate p{font-size:13px;line-height:1.6;color:#8A92AA;margin:0 0 26px;}' +
      '#ilb-gate button{background:#B8791A;color:#fff;border:0;border-radius:8px;padding:12px 26px;' +
      "font-size:14px;font-weight:600;font-family:'DM Sans',sans-serif;cursor:pointer;transition:background .15s;}" +
      '#ilb-gate button:hover{background:#D4922A;}' +
      '#ilb-gate button:disabled{opacity:.5;cursor:default;}' +
      '#ilb-gate .err{color:#D45050;font-size:12px;margin-top:16px;min-height:16px;}' +
      "#ilb-gate .spin{margin-top:20px;color:#8A92AA;font-family:'DM Mono',monospace;font-size:11px;}";
    d.head.appendChild(css);

    gateEl = d.createElement('div');
    gateEl.id = 'ilb-gate';
    gateEl.innerHTML =
      '<div class="box">' +
        '<div class="eyebrow">Montedoro · Gestione Operativa</div>' +
        '<h1>Il Borgo</h1>' +
        '<p>Accesso riservato. Connetti il tuo account Google per continuare.</p>' +
        '<button type="button" id="ilb-gate-btn">Accedi con Google</button>' +
        '<div class="err" id="ilb-gate-err"></div>' +
        '<div class="spin" id="ilb-gate-spin" style="display:none">verifica sessione…</div>' +
      '</div>';
    (d.body || d.documentElement).appendChild(gateEl);

    d.getElementById('ilb-gate-btn').addEventListener('click', function () {
      var btn = this;
      btn.disabled = true;
      d.getElementById('ilb-gate-err').textContent = '';
      requestToken('').catch(function () {
        btn.disabled = false;
        d.getElementById('ilb-gate-err').textContent = 'Accesso non riuscito. Riprova.';
      });
    });
  }
  function showGate(mode) {
    setAppHidden(true);
    buildGate();
    if (!gateEl) return;
    gateEl.style.display = 'flex';
    var spin = d.getElementById('ilb-gate-spin');
    var btn  = d.getElementById('ilb-gate-btn');
    var loading = (mode === 'loading');
    if (spin) spin.style.display = loading ? 'block' : 'none';
    if (btn)  btn.disabled = loading;
  }
  function hideGate() {
    if (gateEl) gateEl.style.display = 'none';
    setAppHidden(false);
  }
  function setAppHidden(hidden) {
    // Nota: il CSS base imposta #app{visibility:hidden}; qui serve un valore
    // esplicito ('visible'), non stringa vuota — altrimenti lo stile inline
    // svuotato ricade sulla regola del foglio di stile e resta nascosto.
    var app = d.getElementById('app');
    if (app) app.style.visibility = hidden ? 'hidden' : 'visible';
    // fallback se la pagina non usa #app
    if (d.body) d.body.setAttribute('data-ilb-auth', hidden ? 'locked' : 'open');
  }

  // ── API pubblica ──────────────────────────────────────────────────
  w.IlBorgoAuth = {
    start: start,
    signIn: function () { return requestToken(''); },
    signOut: signOut,
    ensureToken: function () {
      var c = cachedToken();
      if (c) { pushToGapi(c.token); return Promise.resolve(c.token); }
      return requestToken('none');
    },
    getToken: function () { var c = cachedToken(); return c ? c.token : null; },
    handleAuthError: handleAuthError,
    onChange: function (cb) {
      if (typeof cb !== 'function') return;
      listeners.push(cb);
      if (authed) { try { cb(true); } catch (e) {} }
    },
    isAuthed: function () { return authed; }
  };

  if (d.readyState === 'loading') d.addEventListener('DOMContentLoaded', start);
  else start();
})(window, document);
