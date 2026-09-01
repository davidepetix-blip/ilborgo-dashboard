/* ═══════════════════════════════════════════════════════════════════
   Il Borgo · Dashboard — Google Calendar come DB primario (Parte 2)

   Modello: 1 attività/fornitore ⇄ 1 evento all-day nel calendario della
   sua area (ILBORGO_CONFIG.CALENDARS). I campi strutturati vivono in
   extendedProperties.private (ilb_type, ilb_id, ilb_area, + i campi
   specifici). Uno snapshot locale (id→{area,hash,eventId}) permette di
   fare push differenziale (insert/patch/move/delete) e riconoscere gli
   eventi già mappati durante il poll incrementale (syncToken per area).

   Eventi trovati in un calendario d'area SENZA i tag ilb_* (creati a mano
   in Google Calendar) vengono "adottati": diventano un'attività locale e
   il layer scrive i tag di default sull'evento.

   API pubblica → window.IlBorgoCal
     .isConfigured()          true se tutti e 3 i calendari sono configurati
     .push(tasks, forn)       → Promise<stats>  scrive le modifiche locali
     .poll()                  → Promise<{upserts:[{kind,item,eventId,adopt}], deletes:['t:id'|'f:id']}>
     .startAutoSync(onRemote) poll periodico (60s) + al ritorno sul tab
   ═══════════════════════════════════════════════════════════════════ */
(function (w, d) {
  'use strict';

  var CFG = w.ILBORGO_CONFIG;
  if (!CFG) { console.error('[gcal] ILBORGO_CONFIG mancante'); return; }

  var API = 'https://www.googleapis.com/calendar/v3';
  var SNAP_KEY = 'ilb_gcal_snap';
  var TOKEN_PREFIX = 'ilb_gcal_synctoken_';
  var ADOPT_KEY = 'ilb_adopt_external';
  var POLL_MS = 60000;

  // L'adozione di eventi NON creati dall'app è OPT-IN e spenta di default:
  // i calendari potrebbero non essere dedicati (eventi personali, ricorrenze…).
  function adoptEnabled() {
    try { return localStorage.getItem(ADOPT_KEY) === '1'; } catch (e) { return false; }
  }
  // Filtro prudente per gli eventi adottabili quando l'adozione è attiva.
  function adoptable(ev) {
    if (ev.recurringEventId || ev.recurrence) return false;   // niente ricorrenze
    if (ev.attendees && ev.attendees.length) return false;    // niente eventi con invitati
    if (ev.transparency === 'transparent') return false;      // niente eventi "libero"
    if (!ev.start || !ev.start.date) return false;            // solo all-day
    var t = new Date(ev.start.date + 'T00:00:00Z').getTime();
    var now = Date.now();
    if (t < now - 90 * 864e5 || t > now + 400 * 864e5) return false; // finestra ±
    return true;
  }

  // ── fetch autenticato, con un retry dopo rinnovo su 401/403 ─────────
  function apiFetch(url, opts, _retried) {
    opts = opts || {};
    var tok = w.IlBorgoAuth && w.IlBorgoAuth.getToken();
    if (!tok) return Promise.reject(new Error('no_token'));
    var headers = { Authorization: 'Bearer ' + tok };
    for (var k in (opts.headers || {})) headers[k] = opts.headers[k];
    return fetch(url, { method: opts.method, headers: headers, body: opts.body }).then(function (res) {
      // solo 401 = token non valido → rinnovo. I 403 (rate limit, API non
      // abilitata, permessi) NON devono buttare fuori l'utente.
      if (res.status === 401 && !_retried) {
        return w.IlBorgoAuth.handleAuthError().then(function () { return apiFetch(url, opts, true); });
      }
      if (res.status === 204) return {};
      return res.json().catch(function () { return {}; }).then(function (body) {
        if (!res.ok) {
          var e = new Error((body.error && body.error.message) || ('HTTP ' + res.status));
          e.status = res.status; e.body = body;
          throw e;
        }
        return body;
      });
    });
  }

  // ── snapshot (localStorage) ──────────────────────────────────────────
  function loadSnap() { try { return JSON.parse(localStorage.getItem(SNAP_KEY) || '{}'); } catch (e) { return {}; } }
  function saveSnap(s) { try { localStorage.setItem(SNAP_KEY, JSON.stringify(s)); } catch (e) {} }

  // ── helper ────────────────────────────────────────────────────────
  function calForArea(area) { return CFG.CALENDARS && CFG.CALENDARS[area]; }
  function calendarsReady() {
    var c = CFG.CALENDARS || {};
    return !!(c.borgo_admin && c.affitti && c.architettura);
  }
  function addDaysISO(iso, n) {
    // Costruzione/calcolo in UTC: `new Date(iso+'T00:00:00')` + toISOString()
    // soffre di off-by-one nei fusi orari diversi da UTC (mezzanotte locale
    // può cadere sul giorno prima/dopo in UTC). Date.UTC evita il problema.
    var p = iso.split('-').map(Number);
    var dt = new Date(Date.UTC(p[0], p[1] - 1, p[2]));
    dt.setUTCDate(dt.getUTCDate() + n);
    return dt.toISOString().split('T')[0];
  }
  function eur(n) { return '€' + Number(n || 0).toLocaleString('it-IT'); }

  // ── mapping item ⇄ evento ─────────────────────────────────────────
  function taskFields(t) {
    return {
      ilb_type: 'task', ilb_id: String(t.id), ilb_area: t.area,
      priorita: t.priorita || 'normale', ord: String(t.ord || 0),
      done: t.done ? '1' : '0', quad: t.quad || '', progetto: t.progetto || ''
    };
  }
  function fornFields(f) {
    return {
      ilb_type: 'forn', ilb_id: String(f.id), ilb_area: f.area,
      importo: String(f.importo || 0), stato: f.stato || 'da_pagare', ord: String(f.ord || 0)
    };
  }
  function taskToEventBody(t) {
    return {
      summary: (t.done ? '✓ ' : '') + t.testo,
      description: t.note || '',
      start: { date: t.scadenza }, end: { date: addDaysISO(t.scadenza, 1) },
      extendedProperties: { private: taskFields(t) }
    };
  }
  function fornToEventBody(f) {
    return {
      summary: (f.stato === 'pagato' ? '✓ ' : '💳 ') + f.nome + ' — ' + eur(f.importo),
      description: f.nota || '',
      start: { date: f.scadenza }, end: { date: addDaysISO(f.scadenza, 1) },
      extendedProperties: { private: fornFields(f) }
    };
  }
  function taskHash(t) { return JSON.stringify([t.testo, t.priorita, t.note, t.scadenza, !!t.done, t.ord || 0, t.quad || '', t.progetto || '', t.area]); }
  function fornHash(f) { return JSON.stringify([f.nome, f.importo, f.scadenza, f.stato, f.nota, f.ord || 0, f.area]); }

  function parseEventToTask(ev, area, localId) {
    var p = (ev.extendedProperties && ev.extendedProperties.private) || {};
    return {
      id: localId, area: area,
      testo: (ev.summary || '').replace(/^✓\s*/, ''),
      priorita: p.priorita || 'normale',
      note: ev.description || '',
      scadenza: (ev.start && ev.start.date) || '',
      done: p.done === '1' || /^✓/.test(ev.summary || ''),
      ord: parseInt(p.ord, 10) || 0,
      quad: p.quad || null,
      progetto: p.progetto || ''
    };
  }
  // riconosce gli eventi del VECCHIO sistema di sync (tag testuale nella description)
  function isLegacyManaged(ev) { return /ILBORGO_ID:/.test((ev && ev.description) || ''); }

  function parseEventToForn(ev, area, localId) {
    var p = (ev.extendedProperties && ev.extendedProperties.private) || {};
    return {
      id: localId, area: area,
      nome: (ev.summary || '').replace(/^(✓|💳)\s*/, '').replace(/\s*—\s*€[\d.,]+$/, ''),
      importo: parseFloat(p.importo) || 0,
      scadenza: (ev.start && ev.start.date) || '',
      stato: p.stato || 'da_pagare',
      nota: ev.description || '',
      ord: parseInt(p.ord, 10) || 0
    };
  }

  // ── operazioni Calendar API ──────────────────────────────────────
  function insertEvent(area, body) {
    var calId = calForArea(area);
    return apiFetch(API + '/calendars/' + encodeURIComponent(calId) + '/events', {
      method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body)
    });
  }
  function patchEvent(area, eventId, body) {
    var calId = calForArea(area);
    return apiFetch(API + '/calendars/' + encodeURIComponent(calId) + '/events/' + encodeURIComponent(eventId), {
      method: 'PATCH', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body)
    });
  }
  function deleteEvent(area, eventId) {
    var calId = calForArea(area);
    return apiFetch(API + '/calendars/' + encodeURIComponent(calId) + '/events/' + encodeURIComponent(eventId), { method: 'DELETE' })
      .catch(function (e) { if (e.status === 410 || e.status === 404) return null; throw e; });
  }
  function moveEvent(fromArea, toArea, eventId) {
    var fromCal = calForArea(fromArea), toCal = calForArea(toArea);
    return apiFetch(API + '/calendars/' + encodeURIComponent(fromCal) + '/events/' + encodeURIComponent(eventId) +
      '/move?destination=' + encodeURIComponent(toCal), { method: 'POST' });
  }
  function listEventsPage(calId, params) {
    var qs = Object.keys(params).map(function (k) { return k + '=' + encodeURIComponent(params[k]); }).join('&');
    return apiFetch(API + '/calendars/' + encodeURIComponent(calId) + '/events?' + qs);
  }

  // ── PUSH: diff stato locale vs snapshot ──────────────────────────
  function push(tasks, forn) {
    if (!calendarsReady()) return Promise.resolve({ skipped: true });
    var snap = loadSnap();
    var newSnap = {};
    var currentKeys = {};
    var chain = Promise.resolve();
    var stats = { created: 0, updated: 0, moved: 0, deleted: 0, errors: 0 };

    function queue(fn) { chain = chain.then(fn).catch(function (e) { stats.errors++; console.error('[gcal] push error', e); }); }

    function handleItem(key, item, hash, area, toBody) {
      currentKeys[key] = 1;
      var prev = snap[key];
      if (!prev) {
        queue(function () {
          return insertEvent(area, toBody(item)).then(function (ev) {
            newSnap[key] = { area: area, hash: hash, eventId: ev.id };
            stats.created++;
          });
        });
      } else if (prev.area !== area) {
        queue(function () {
          return moveEvent(prev.area, area, prev.eventId)
            .then(function () { return patchEvent(area, prev.eventId, toBody(item)); })
            .then(function () { newSnap[key] = { area: area, hash: hash, eventId: prev.eventId }; stats.moved++; });
        });
      } else if (prev.hash !== hash) {
        queue(function () {
          return patchEvent(area, prev.eventId, toBody(item)).then(function () {
            newSnap[key] = { area: area, hash: hash, eventId: prev.eventId }; stats.updated++;
          });
        });
      } else {
        newSnap[key] = prev; // invariato
      }
    }

    (tasks || []).forEach(function (t) {
      if (!t.scadenza) return; // difensivo: la UI rende la scadenza obbligatoria
      handleItem('t:' + t.id, t, taskHash(t), t.area, taskToEventBody);
    });
    (forn || []).forEach(function (f) {
      if (!f.scadenza) return;
      handleItem('f:' + f.id, f, fornHash(f), f.area, fornToEventBody);
    });

    // eliminazioni: presenti nello snapshot precedente ma non più nello stato locale
    Object.keys(snap).forEach(function (key) {
      if (currentKeys[key]) return;
      var s = snap[key];
      queue(function () { return deleteEvent(s.area, s.eventId).then(function () { stats.deleted++; }); });
    });

    return chain.then(function () { saveSnap(newSnap); return stats; });
  }

  // ── POLL: sync incrementale per area + adozione eventi esterni ──
  function pollArea(area) {
    var calId = calForArea(area);
    if (!calId) return Promise.resolve({ upserts: [], deletes: [] });
    var tokenKey = TOKEN_PREFIX + area;
    var syncToken = null;
    try { syncToken = localStorage.getItem(tokenKey) || null; } catch (e) {}

    var snap = loadSnap();
    function keyForEventId(evId) {
      for (var k in snap) { if (snap[k].eventId === evId) return k; }
      return null;
    }

    var upserts = [], deletes = [];

    function pageLoop(pageToken) {
      var params = { maxResults: 250, showDeleted: true, singleEvents: true };
      if (syncToken && !pageToken) params.syncToken = syncToken;
      if (pageToken) params.pageToken = pageToken;
      return listEventsPage(calId, params).then(function (data) {
        (data.items || []).forEach(function (ev) {
          var key = keyForEventId(ev.id);
          if (ev.status === 'cancelled') { if (key) deletes.push(key); return; }
          var p = (ev.extendedProperties && ev.extendedProperties.private) || {};
          if (p.ilb_type === 'task') {
            var id1 = key ? parseInt(key.slice(2), 10) : (parseInt(p.ilb_id, 10) || Date.now());
            upserts.push({ kind: 'task', item: parseEventToTask(ev, area, id1), eventId: ev.id });
          } else if (p.ilb_type === 'forn') {
            var id2 = key ? parseInt(key.slice(2), 10) : (parseInt(p.ilb_id, 10) || Date.now());
            upserts.push({ kind: 'forn', item: parseEventToForn(ev, area, id2), eventId: ev.id });
          } else if (adoptEnabled() && !key && !isLegacyManaged(ev) && adoptable(ev)) {
            // adozione (opt-in): evento creato a mano in un calendario d'area,
            // senza tag, non ricorrente, all-day, entro una finestra ragionevole.
            var newId = Date.now() + Math.floor(Math.random() * 1000);
            upserts.push({ kind: 'task', item: parseEventToTask(ev, area, newId), eventId: ev.id, adopt: true });
          }
        });
        if (data.nextPageToken) return pageLoop(data.nextPageToken);
        if (data.nextSyncToken) { try { localStorage.setItem(tokenKey, data.nextSyncToken); } catch (e) {} }
      });
    }

    return pageLoop(null)
      .catch(function (e) {
        if (e.status === 410) { // sync token non più valido: full resync
          try { localStorage.removeItem(tokenKey); } catch (err) {}
          syncToken = null;
          return pageLoop(null);
        }
        throw e;
      })
      .then(function () { return { upserts: upserts, deletes: deletes }; });
  }

  function poll() {
    if (!calendarsReady()) return Promise.resolve({ upserts: [], deletes: [] });
    var areas = Object.keys(CFG.CALENDARS);
    var allUpserts = [], allDeletes = [];
    var chain = Promise.resolve();
    areas.forEach(function (area) {
      chain = chain.then(function () {
        return pollArea(area).then(function (r) {
          allUpserts = allUpserts.concat(r.upserts);
          allDeletes = allDeletes.concat(r.deletes);
        }).catch(function (e) { console.error('[gcal] poll error area=' + area, e); });
      });
    });
    return chain.then(function () {
      var snap = loadSnap();
      var claimChain = Promise.resolve();
      allUpserts.forEach(function (u) {
        var key = (u.kind === 'task' ? 't:' : 'f:') + u.item.id;
        var hash = u.kind === 'task' ? taskHash(u.item) : fornHash(u.item);
        snap[key] = { area: u.item.area, hash: hash, eventId: u.eventId };
        if (u.adopt) {
          // scrive i tag ilb_* di default sull'evento adottato, per non riadottarlo ogni poll
          claimChain = claimChain.then(function () {
            return patchEvent(u.item.area, u.eventId, u.kind === 'task' ? taskToEventBody(u.item) : fornToEventBody(u.item))
              .catch(function (e) { console.error('[gcal] adopt patch error', e); });
          });
        }
      });
      allDeletes.forEach(function (key) { delete snap[key]; });
      return claimChain.then(function () { saveSnap(snap); return { upserts: allUpserts, deletes: allDeletes }; });
    });
  }

  // ── PURGE: rimuove tutti gli eventi gestiti dall'app (nuovi + legacy) ──
  function purgeCal(calId) {
    var ids = [];
    function collect(pageToken) {
      var params = { maxResults: 250, singleEvents: true };
      if (pageToken) params.pageToken = pageToken;
      return listEventsPage(calId, params).then(function (data) {
        (data.items || []).forEach(function (ev) {
          var p = (ev.extendedProperties && ev.extendedProperties.private) || {};
          if (p.ilb_type || isLegacyManaged(ev)) ids.push(ev.id);
        });
        if (data.nextPageToken) return collect(data.nextPageToken);
      });
    }
    return collect(null).then(function () {
      var c = Promise.resolve(), count = 0;
      ids.forEach(function (id) {
        c = c.then(function () {
          return apiFetch(API + '/calendars/' + encodeURIComponent(calId) + '/events/' + encodeURIComponent(id), { method: 'DELETE' })
            .then(function () { count++; })
            .catch(function (e) { if (e.status !== 404 && e.status !== 410) console.error('[gcal] purge', e); });
        });
      });
      return c.then(function () { return count; });
    });
  }
  function purgeAll() {
    if (!calendarsReady()) return Promise.resolve({ deleted: 0 });
    var areas = Object.keys(CFG.CALENDARS);
    var deleted = 0;
    var chain = Promise.resolve();
    areas.forEach(function (area) {
      chain = chain.then(function () { return purgeCal(calForArea(area)); }).then(function (n) { deleted += n; });
    });
    return chain.then(function () {
      try { localStorage.removeItem(SNAP_KEY); } catch (e) {}
      areas.forEach(function (a) { try { localStorage.removeItem(TOKEN_PREFIX + a); } catch (e) {} });
      return { deleted: deleted };
    });
  }

  // ── serializzazione: push e poll non si accavallano mai ──────────
  var syncChain = Promise.resolve();
  function serialized(fn) {
    var p = syncChain.then(fn, fn);
    syncChain = p.then(function () {}, function () {});
    return p;
  }

  w.IlBorgoCal = {
    isConfigured: calendarsReady,
    adoptEnabled: adoptEnabled,
    setAdopt: function (on) { try { localStorage.setItem(ADOPT_KEY, on ? '1' : '0'); } catch (e) {} },
    push: function (tasksArr, fornArr) { return serialized(function () { return push(tasksArr, fornArr); }); },
    poll: function () { return serialized(function () { return poll(); }); },
    purgeAll: function () { return serialized(function () { return purgeAll(); }); },
    startAutoSync: function (onRemote) {
      var last = 0;
      function tick(force) {
        if (!(w.IlBorgoAuth && w.IlBorgoAuth.isAuthed())) return;
        if (!force && Date.now() - last < 20000) return; // anti-raffica su visibilitychange
        last = Date.now();
        w.IlBorgoCal.poll().then(function (r) {
          if ((r.upserts && r.upserts.length) || (r.deletes && r.deletes.length)) onRemote(r);
        }).catch(function (e) { console.error('[gcal] autosync error', e); });
      }
      setInterval(function () { tick(true); }, POLL_MS);
      d.addEventListener('visibilitychange', function () { if (!d.hidden) tick(false); });
    }
  };
})(window, document);
