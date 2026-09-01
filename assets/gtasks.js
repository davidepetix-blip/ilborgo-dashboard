/* ═══════════════════════════════════════════════════════════════════
   Il Borgo · Dashboard — Google Tasks per i micro-passaggi (Parte 3)

   Keep API richiede Workspace → si usa Google Tasks (compatibile Gmail).
   Una lista Tasks per area ("Il Borgo · <Area>"), creata al bisogno.

   Un micro-passaggio = un task Google nella lista dell'area del task
   padre della dashboard:
     - title : "[<progetto o testo padre>] <testo passo>"  (grouping visibile
               anche aprendo Google Tasks)
     - due   : scadenza del task padre  → compare NATIVAMENTE in Google Calendar
     - notes : "ilb_parent:<idTaskLocale>"  → collegamento al task padre

   Spuntare in app  → tasks.patch status=completed
   Spuntare in Google Tasks → letto da poll()  (bidirezionale)

   API pubblica → window.IlBorgoTasks
     .ensureLists()                  → Promise  (popola ILBORGO_CONFIG.TASK_LISTS)
     .isConfigured()                 true se le 3 liste sono note
     .addStep(parentTask, testo)     → Promise<step>
     .toggleStep(step, done)         → Promise
     .deleteStep(step)               → Promise
     .syncStepsForTask(parentTask)   → Promise  (allinea due/lista dei passi al padre)
     .poll()                         → Promise<{steps:[{id,parentId,area,testo,done}]}>
     .startAutoSync(onSteps)         poll ogni 60s + focus tab
   ═══════════════════════════════════════════════════════════════════ */
(function (w, d) {
  'use strict';

  var CFG = w.ILBORGO_CONFIG;
  if (!CFG) { console.error('[gtasks] ILBORGO_CONFIG mancante'); return; }

  var API = 'https://tasks.googleapis.com/tasks/v1';
  var POLL_MS = 60000;
  var LIST_TITLES = {
    borgo_admin:  'Il Borgo · Amministrazione',
    affitti:      'Il Borgo · Affitti',
    architettura: 'Il Borgo · Architettura'
  };
  var AREAS = ['borgo_admin', 'affitti', 'architettura'];

  // ── fetch autenticato, un retry dopo rinnovo su 401/403 ─────────────
  function apiFetch(url, opts, _retried) {
    opts = opts || {};
    var tok = w.IlBorgoAuth && w.IlBorgoAuth.getToken();
    if (!tok) return Promise.reject(new Error('no_token'));
    var headers = { Authorization: 'Bearer ' + tok };
    for (var k in (opts.headers || {})) headers[k] = opts.headers[k];
    return fetch(url, { method: opts.method, headers: headers, body: opts.body }).then(function (res) {
      // solo 401 → rinnovo token. 403 (Tasks API non abilitata, rate limit) NON
      // deve buttare fuori l'utente.
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

  // ── liste per area ────────────────────────────────────────────────
  function loadListMap() {
    try { return JSON.parse(localStorage.getItem(CFG.STORAGE.taskLists) || '{}'); } catch (e) { return {}; }
  }
  function saveListMap(m) {
    try { localStorage.setItem(CFG.STORAGE.taskLists, JSON.stringify(m)); } catch (e) {}
    CFG.TASK_LISTS = m;
  }
  function isConfigured() {
    var m = CFG.TASK_LISTS || {};
    return !!(m.borgo_admin && m.affitti && m.architettura);
  }

  var listsPromise = null;
  function ensureLists() {
    if (listsPromise) return listsPromise;
    var cached = loadListMap();
    if (cached.borgo_admin && cached.affitti && cached.architettura) {
      CFG.TASK_LISTS = cached;
      listsPromise = Promise.resolve(cached);
      return listsPromise;
    }
    listsPromise = apiFetch(API + '/users/@me/lists?maxResults=100').then(function (data) {
      var byTitle = {};
      (data.items || []).forEach(function (l) { byTitle[l.title] = l.id; });
      var map = {};
      var chain = Promise.resolve();
      AREAS.forEach(function (area) {
        var title = LIST_TITLES[area];
        if (byTitle[title]) { map[area] = byTitle[title]; return; }
        chain = chain.then(function () {
          return apiFetch(API + '/users/@me/lists', {
            method: 'POST', headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ title: title })
          }).then(function (l) { map[area] = l.id; });
        });
      });
      return chain.then(function () { saveListMap(map); return map; });
    }).catch(function (e) {
      listsPromise = null; // riprova al prossimo giro
      throw e;
    });
    return listsPromise;
  }

  // ── mapping passo ⇄ task Google ───────────────────────────────────
  function stepLabel(parentTask) {
    return (parentTask.progetto && parentTask.progetto.trim()) || parentTask.testo || 'Attività';
  }
  function stepTitle(parentTask, testo) { return '[' + stepLabel(parentTask) + '] ' + testo; }
  function stripLabel(title) { return String(title || '').replace(/^\[[^\]]*\]\s*/, ''); }
  function dueFromScadenza(scad) { return scad ? (scad + 'T00:00:00.000Z') : null; }
  function parentIdFromNotes(notes) {
    var m = String(notes || '').match(/ilb_parent:(\d+)/);
    return m ? parseInt(m[1], 10) : null;
  }

  function listIdForArea(area) { return (CFG.TASK_LISTS || {})[area]; }

  // ── operazioni ───────────────────────────────────────────────────
  function addStep(parentTask, testo) {
    return ensureLists().then(function () {
      var listId = listIdForArea(parentTask.area);
      var body = {
        title: stepTitle(parentTask, testo),
        notes: 'ilb_parent:' + parentTask.id
      };
      var due = dueFromScadenza(parentTask.scadenza);
      if (due) body.due = due;
      return apiFetch(API + '/lists/' + encodeURIComponent(listId) + '/tasks', {
        method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body)
      }).then(function (t) {
        return { id: t.id, parentId: parentTask.id, area: parentTask.area, testo: testo, done: false };
      });
    });
  }

  function toggleStep(step, done) {
    var listId = listIdForArea(step.area);
    if (!listId) return Promise.reject(new Error('lista_sconosciuta'));
    return apiFetch(API + '/lists/' + encodeURIComponent(listId) + '/tasks/' + encodeURIComponent(step.id), {
      method: 'PATCH', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ status: done ? 'completed' : 'needsAction' })
    });
  }

  function deleteStep(step) {
    var listId = listIdForArea(step.area);
    if (!listId) return Promise.resolve();
    return apiFetch(API + '/lists/' + encodeURIComponent(listId) + '/tasks/' + encodeURIComponent(step.id),
      { method: 'DELETE' }).catch(function (e) { if (e.status === 404 || e.status === 410) return null; throw e; });
  }

  // allinea due (e lista, se l'area è cambiata) dei passi di un task padre
  function syncStepsForTask(parentTask) {
    return ensureLists().then(function () {
      return poll().then(function (r) {
        var mine = r.steps.filter(function (s) { return s.parentId === parentTask.id; });
        var due = dueFromScadenza(parentTask.scadenza);
        var chain = Promise.resolve();
        mine.forEach(function (s) {
          if (s.area === parentTask.area) {
            // stessa lista: aggiorna solo due + prefisso titolo
            var listId = listIdForArea(parentTask.area);
            chain = chain.then(function () {
              return apiFetch(API + '/lists/' + encodeURIComponent(listId) + '/tasks/' + encodeURIComponent(s.id), {
                method: 'PATCH', headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ title: stepTitle(parentTask, s.testo), due: due })
              }).catch(function (e) { console.error('[gtasks] sync step error', e); });
            });
          } else {
            // area cambiata: Tasks non sposta tra liste → ricrea nella nuova, elimina la vecchia
            chain = chain.then(function () {
              return addStep(parentTask, s.testo)
                .then(function (ns) { return s.done ? toggleStep(ns, true) : null; })
                .then(function () { return deleteStep(s); })
                .catch(function (e) { console.error('[gtasks] move step error', e); });
            });
          }
        });
        return chain;
      });
    });
  }

  // ── POLL: legge tutti i passi dalle 3 liste ──────────────────────
  function poll() {
    return ensureLists().then(function () {
      var steps = [];
      var chain = Promise.resolve();
      AREAS.forEach(function (area) {
        var listId = listIdForArea(area);
        if (!listId) return;
        chain = chain.then(function () {
          return apiFetch(API + '/lists/' + encodeURIComponent(listId) +
            '/tasks?showCompleted=true&showHidden=true&maxResults=100').then(function (data) {
            (data.items || []).forEach(function (t) {
              var pid = parentIdFromNotes(t.notes);
              if (pid == null) return; // non è un nostro micro-passaggio
              steps.push({
                id: t.id, parentId: pid, area: area,
                testo: stripLabel(t.title),
                done: t.status === 'completed'
              });
            });
          }).catch(function (e) { console.error('[gtasks] poll error area=' + area, e); });
        });
      });
      return chain.then(function () { return { steps: steps }; });
    }).catch(function (e) {
      console.error('[gtasks] poll/ensureLists error', e);
      return { steps: null }; // null = non toccare la cache locale
    });
  }

  // rimuove tutti i micro-passaggi (task con notes ilb_parent:) dalle 3 liste
  function purgeAll() {
    return ensureLists().then(function () {
      var count = 0;
      var chain = Promise.resolve();
      AREAS.forEach(function (area) {
        var listId = listIdForArea(area);
        if (!listId) return;
        chain = chain.then(function () {
          return apiFetch(API + '/lists/' + encodeURIComponent(listId) +
            '/tasks?showCompleted=true&showHidden=true&maxResults=100').then(function (data) {
            var ids = (data.items || []).filter(function (t) { return /ilb_parent:/.test(t.notes || ''); })
              .map(function (t) { return t.id; });
            var c = Promise.resolve();
            ids.forEach(function (id) {
              c = c.then(function () {
                return apiFetch(API + '/lists/' + encodeURIComponent(listId) + '/tasks/' + encodeURIComponent(id),
                  { method: 'DELETE' }).then(function () { count++; }).catch(function () {});
              });
            });
            return c;
          }).catch(function (e) { console.error('[gtasks] purge area=' + area, e); });
        });
      });
      return chain.then(function () { return { deleted: count }; });
    });
  }

  w.IlBorgoTasks = {
    ensureLists: ensureLists,
    isConfigured: isConfigured,
    purgeAll: purgeAll,
    addStep: addStep,
    toggleStep: toggleStep,
    deleteStep: deleteStep,
    syncStepsForTask: syncStepsForTask,
    poll: poll,
    startAutoSync: function (onSteps) {
      var last = 0;
      function tick(force) {
        if (!(w.IlBorgoAuth && w.IlBorgoAuth.isAuthed())) return;
        if (!force && Date.now() - last < 20000) return;
        last = Date.now();
        poll().then(function (r) { if (r.steps) onSteps(r); })
          .catch(function (e) { console.error('[gtasks] autosync error', e); });
      }
      setInterval(function () { tick(true); }, POLL_MS);
      d.addEventListener('visibilitychange', function () { if (!d.hidden) tick(false); });
    }
  };
})(window, document);
