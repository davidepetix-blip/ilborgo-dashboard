/* ═══════════════════════════════════════════════════════════════════
   Il Borgo · Dashboard — Configurazione condivisa
   Caricato da index.html e affitti.html PRIMA di auth.js e della logica inline.
   Nessun modulo ES: espone un singolo globale window.ILBORGO_CONFIG.
   ═══════════════════════════════════════════════════════════════════ */
(function (w) {
  'use strict';

  var CONFIG = {
    // OAuth client (Google Identity Services)
    CLIENT_ID: '850622821617-7l6ll27pfs5hkfng3savskud4pqjiqcn.apps.googleusercontent.com',

    // Scope unificati per tutte e tre le fasi.
    // Includiamo calendar + tasks fin da subito: così il consenso viene chiesto
    // una sola volta e i milestone 2/3 non richiedono un nuovo prompt.
    SCOPES: [
      'https://www.googleapis.com/auth/spreadsheets',
      'https://www.googleapis.com/auth/calendar',
      'https://www.googleapis.com/auth/tasks'
    ].join(' '),

    // Google Sheet primario della dashboard (index.html).
    // Dalla Parte 2 in poi retrocede a backup/log secondario.
    SHEET_ID_MAIN: '1tuO4m5MABDGmLb63_42sV3ikWXsO_eexLWg6Vx0eBf0',

    // Parte 2 — un calendario Google per area (sorgente primaria).
    // architettura coincide col vecchio CAL_ID_LEGACY.
    CALENDARS: {
      borgo_admin:  'p5bjseovvrcsfttn87e3drgh78@group.calendar.google.com',
      affitti:      'tasuud8409bhoqilm97eh7foj8@group.calendar.google.com',
      architettura: 'c5glbck4juqq4f9r5k6d13tbl4@group.calendar.google.com'
    },

    // Calendar ID singolo usato dalla vecchia syncCalendar() (deprecato in Parte 2).
    CAL_ID_LEGACY: 'c5glbck4juqq4f9r5k6d13tbl4@group.calendar.google.com',

    // Parte 3 — una lista Google Tasks per area. Popolate a runtime da
    // gtasks.ensureLists() e persistite in localStorage con la chiave qui sotto.
    TASK_LISTS: { borgo_admin: '', affitti: '', architettura: '' },

    // Chiavi localStorage condivise
    STORAGE: {
      token:     'ilb_token',
      tokenExp:  'ilb_token_exp',
      calendars: 'ilb_calendars',
      taskLists: 'ilb_tasklists'
    }
  };

  w.ILBORGO_CONFIG = CONFIG;
})(window);
