// =============================================================
// SCRIPT UNIFICATO — Prenotazioni + JSON_ANNUALE
// Versione: 2026-03-13
// =============================================================
// ISTRUZIONI:
//   1. Sostituisci TUTTO il contenuto dell'Apps Script con questo
//   2. Salva (Ctrl+S)
//   3. Menu "🏨 Script Prenotazioni" → "🔄 Rigenera JSON_ANNUALE"
//   4. Se il foglio è vuoto → "🔍 Debug JSON_ANNUALE" per diagnosticare
// =============================================================

// ── Costanti script prenotazioni ──
const EXCLUDED_SHEETS = [
  "Dati Centralizzati Realtime","Non toccare","Ricettività",
  "LOG COMPLESSIVO","PRENOTAZIONI","JSON_ANNUALE","Foglio212"
];
const YELLOW_BORDER_COLOR  = "#FFFF00";
const BLACK_BORDER_COLOR   = "#000000";
const ERROR_BORDER_COLOR   = "#FF0000";
const BLACK_BORDER_STYLE   = SpreadsheetApp.BorderStyle.SOLID_MEDIUM;
const SUNDAY_BORDER_STYLE  = SpreadsheetApp.BorderStyle.SOLID_MEDIUM;
const ERROR_BORDER_STYLE   = SpreadsheetApp.BorderStyle.SOLID_THICK;
const HEADER_RANGES        = ["B1:G1","H1:Q1","R1:V1","W1:AJ1","B2:G34","H2:Q34","R2:V34","W2:AJ34","B34:AJ34"];
const FIRST_DATA_ROW       = 3;
const HEADER_ROW_NUMBER    = 2;
const DATES_COLUMN         = 1;
const FIRST_CAMERA_COLUMN  = 2;
const OUTPUT_ROW           = 45;
const MONTH_NAMES = {
  "gen":0,"feb":1,"mar":2,"apr":3,"mag":4,"giu":5,"lug":6,"ago":7,"set":8,"ott":9,"nov":10,"dic":11,
  "gennaio":0,"febbraio":1,"marzo":2,"aprile":3,"maggio":4,"giugno":5,
  "luglio":6,"agosto":7,"settembre":8,"ottobre":9,"novembre":10,"dicembre":11
};
const VALID_BED_ARRANGEMENTS = ["1m/s","1m","2m","1s","2s","3s","4s","5s","6s","1c","1aff","ND"];
const PROCESSING_STATE_KEY   = 'jsonProcessingState';
const BATCH_TIME_LIMIT_MS    = 5 * 60 * 1000;

// ── Costanti JSON_ANNUALE ──
const JS_SHEET_NAME      = "JSON_ANNUALE";
const JS_TABLE_START_ROW = 4;
const JS_FIRST_CAM_COL   = 2;
const JS_FIRST_DATA_ROW  = 3;
const JS_HEADER_ROW      = 2;
const JS_DEBOUNCE_SEC    = 10;
const JS_SFONDO_NEUTRI   = ["#ffffff","#fffffe"]; // #fce5cd rimosso: è il colore degli affitti
const JS_MESI = {
  "gennaio":0,"febbraio":1,"marzo":2,"aprile":3,"maggio":4,"giugno":5,
  "luglio":6,"agosto":7,"settembre":8,"ottobre":9,"novembre":10,"dicembre":11
};
const JS_DISPO_RE = /\b(\d+\s*m\/s|\d+\s*ms|\d+\s*m(?![\/\w])|\d+\s*s(?!\w)|\d+\s*c(?!\w)|\d+\s*aff(?!\w)|nd)\b/gi;
const JS_SKIP_RE  = /^(dispo\b|2\s*cambi|1\s*cambio|magazzino|cp\b|\d+\s*cambi)/i;


// =============================================================
// MENU — unico onOpen
// =============================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🏨 Script Prenotazioni')
    .addItem('Vai a Oggi', 'goToToday')
    .addSeparator()
    .addItem('Applica Bordi e Formattazione', 'applySundayBordersToAllSheetsManually')
    .addItem('Aggiorna Tutti i JSON (Batch)', 'startBatchProcessing')
    .addSeparator()
    .addItem('🔄 Rigenera JSON_ANNUALE', 'aggiornaJSONAnnuale')
    .addItem('🔍 Debug JSON_ANNUALE', 'debugJSONAnnuale')
    .addSeparator()
    .addItem('⏱ Installa aggiornamento automatico (ogni 5 min)', 'installaTriggerAutomatico')
    .addItem('⏹ Rimuovi aggiornamento automatico', 'rimuoviTriggerAutomatico')
    .addToUi();
}


// =============================================================
// TRIGGER onEdit — unico per tutto
// =============================================================
function onEdit(e) {
  if (!e || e.user == null) return;
  const sheet = e.source.getActiveSheet();
  const col   = e.range.getColumn();
  const row   = e.range.getRow();

  if (!EXCLUDED_SHEETS.includes(sheet.getName())
      && col >= FIRST_CAMERA_COLUMN
      && row >= FIRST_DATA_ROW
      && row < OUTPUT_ROW) {
    processSingleColumnBookings(sheet, col);
  }

  // Segna modifica per il trigger time-based
  segnaModifica();
  aggiornaJSONAnnualeOnEdit(e);
}

/**
 * Chiamata dall'app via Sheets API dopo ogni scrittura di prenotazione.
 * Segna il foglio come modificato — il trigger lo rileverà entro 5 minuti.
 * Accessibile anche come Web App se deployato con accesso "chiunque".
 */
/**
 * Web App endpoint — chiamato dall'app dopo ogni scrittura di prenotazione.
 * Rigenera il JSON_ANNUALE immediatamente e risponde con il risultato.
 *
 * Deploy: Apps Script → Distribuisci → Nuova distribuzione → Tipo: App web
 *   - Esegui come: Me
 *   - Chi può accedere: Chiunque
 * Copia l'URL e incollalo nelle impostazioni dell'app (campo "Web App URL")
 */
function doGet(e) {
  return _rigenera(e);
}
function doPost(e) {
  return _rigenera(e);
}

function _rigenera(e) {
  const t0 = Date.now();
  try {
    const anno    = parseInt((e && e.parameter && e.parameter.anno) || new Date().getFullYear());
    const ss      = SpreadsheetApp.getActiveSpreadsheet();
    const segmenti= estraiSegmenti(ss, anno);
    const merged  = unisciMultiMese(segmenti);
    salvaJsonAnnuale(ss, merged, anno);
    const ms = Date.now() - t0;
    Logger.log('[WebApp] Rigenerato ' + merged.length + ' prenotazioni in ' + ms + 'ms');
    return ContentService
      .createTextOutput(JSON.stringify({ ok:true, prenotazioni:merged.length, ms }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    Logger.log('[WebApp] Errore: ' + err.message);
    return ContentService
      .createTextOutput(JSON.stringify({ ok:false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function notificaModificaDaApp(e) {
  segnaModifica();
  return _rigenera(e);
}


// =============================================================
// TRIGGER TIME-BASED — Rigenera JSON_ANNUALE automaticamente
// Installare una volta sola dal menu Script Prenotazioni
// =============================================================

/**
 * Installa un trigger che rigenera JSON_ANNUALE ogni 5 minuti.
 * Chiamato dal menu o manualmente una volta sola.
 */
function installaTriggerAutomatico() {
  // Rimuovi eventuali trigger esistenti sulle stesse funzioni
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'rigenera5min') {
      ScriptApp.deleteTrigger(t);
    }
  });
  // Installa trigger ogni 5 minuti
  ScriptApp.newTrigger('rigenera5min')
    .timeBased()
    .everyMinutes(5)
    .create();
  SpreadsheetApp.getUi().alert('✅ Trigger installato: JSON_ANNUALE si aggiornerà ogni 5 minuti automaticamente.');
}

/**
 * Rimuove il trigger automatico.
 */
function rimuoviTriggerAutomatico() {
  let rimossi = 0;
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'rigenera5min') {
      ScriptApp.deleteTrigger(t);
      rimossi++;
    }
  });
  SpreadsheetApp.getUi().alert(`Trigger rimosso (${rimossi} eliminati).`);
}

/**
 * Funzione chiamata dal trigger ogni 5 minuti.
 * Rigenera JSON_ANNUALE solo se il foglio è stato modificato di recente
 * (evita rigenarazioni inutili nelle ore di inattività).
 */
function rigenera5min() {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const anno  = new Date().getFullYear();
    const props = PropertiesService.getScriptProperties();
    const ora   = Date.now();

    // Controlla il flag: cella JSON_ANNUALE!A1 o Properties
    const jsSheet = ss.getSheetByName(JS_SHEET_NAME);
    let ultimaMod = parseInt(props.getProperty('ultima_modifica_ts') || '0');

    // Leggi anche dalla cella A1 (scritta dall'app via API)
    if (jsSheet) {
      const a1Val = jsSheet.getRange('A1').getValue();
      if (typeof a1Val === 'string' && a1Val.includes('app:')) {
        // Estrai timestamp dalla stringa "Ultimo aggiornamento app: 2026-..."
        const match = a1Val.match(/(\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2})/);
        if (match) {
          const tsApp = new Date(match[1]).getTime();
          if (tsApp > ultimaMod) ultimaMod = tsApp;
        }
      }
    }

    // Se l'ultima modifica è più vecchia di 10 minuti, non rigenerare
    if (ora - ultimaMod > 10 * 60 * 1000) {
      Logger.log('[Trigger 5min] Nessuna modifica recente, skip');
      return;
    }

    Logger.log('[Trigger 5min] Modifica rilevata (' + new Date(ultimaMod).toISOString() + '), rigenero...');
    const segmenti = estraiSegmenti(ss, anno);
    const merged   = unisciMultiMese(segmenti);
    salvaJsonAnnuale(ss, merged, anno);
    props.setProperty('ultima_regen_ts', String(ora));
    Logger.log('[Trigger 5min] ✓ JSON_ANNUALE: ' + merged.length + ' prenotazioni');
  } catch(e) {
    Logger.log('[Trigger 5min] Errore: ' + e.message);
  }
}

/**
 * Segna che il foglio è stato modificato (chiamato da onEdit e da scritture API).
 * Il trigger time-based usa questo flag per decidere se rigenerare.
 */
function segnaModifica() {
  PropertiesService.getScriptProperties().setProperty('ultima_modifica_ts', String(Date.now()));
}

// =============================================================
// NAVIGAZIONE
// =============================================================
function goToToday() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if (EXCLUDED_SHEETS.includes(sheet.getName())) return;
  const dates = sheet.getRange(1, DATES_COLUMN, sheet.getLastRow()).getValues();
  const today = new Date(); today.setHours(0,0,0,0);
  for (let i = FIRST_DATA_ROW-1; i < dates.length; i++) {
    const d = dates[i][0] instanceof Date ? dates[i][0] : new Date(dates[i][0]);
    if (!isNaN(d.getTime()) && d >= today) { sheet.getRange(i+1, DATES_COLUMN).activate(); return; }
  }
}


// =============================================================
// PRENOTAZIONI PER-COLONNA
// =============================================================
function processSingleColumnBookings(sheet, column) {
  const cameraHeader = String(sheet.getRange(HEADER_ROW_NUMBER, column).getValue()).trim();
  if (!cameraHeader) { sheet.getRange(OUTPUT_ROW, column).clearContent(); return; }
  const lastRow  = Math.min(sheet.getLastRow(), OUTPUT_ROW-1);
  const numRows  = lastRow - FIRST_DATA_ROW + 1;
  const dataRange = sheet.getRange(FIRST_DATA_ROW, column, numRows, 1);
  dataRange.setBorder(false,false,false,false,false,false);
  const values  = dataRange.getValues();
  const bgs     = dataRange.getBackgrounds();
  const uiNotes = dataRange.getNotes();
  const dates   = sheet.getRange(FIRST_DATA_ROW, DATES_COLUMN, numRows, 1).getValues();
  let bookings = [], currentRes = null, startRow = -1;
  for (let i = 0; i < values.length; i++) {
    const bg        = bgs[i][0];
    const isColored = bg && bg !== "#ffffff" && bg.toLowerCase() !== "white";
    const cellValue = String(values[i][0] || "").trim();
    const { dispositionString, remainder } = extractArrangements(cellValue);
    const { name, notes: textNotes }       = cleanAndExtractNameAndNotes(remainder);
    const parsedDate = parseDate(dates[i][0], sheet);
    const cellUINote = (uiNotes[i][0] || "").trim();
    if (isColored) {
      const isNew = !currentRes || bg !== currentRes.backgroundColor ||
                    (dispositionString && (name !== currentRes.nome || dispositionString !== currentRes.disposizione));
      if (isNew) {
        if (currentRes) { currentRes.al = parsedDate; validateAndPush(bookings, currentRes, sheet, startRow, column); }
        if (name || dispositionString) {
          let n = []; if (cellUINote) n.push(cellUINote); if (textNotes && textNotes !== name) n.push(textNotes);
          currentRes = { camera:cameraHeader, nome:name, dal:parsedDate, note:n.join(" - "),
                         backgroundColor:bg, disposizione:dispositionString, matrimoniali:0, singoli:0, culle:0 };
          startRow = FIRST_DATA_ROW + i;
        } else { currentRes = null; }
      } else if (currentRes) {
        let a = [];
        if (cellUINote && !currentRes.note.includes(cellUINote)) a.push(cellUINote);
        if (remainder && remainder !== currentRes.nome && !currentRes.note.includes(remainder)) a.push(remainder);
        if (a.length) currentRes.note = (currentRes.note ? currentRes.note + " - " : "") + a.join(" - ");
      }
    } else if (currentRes) {
      currentRes.al = parsedDate; validateAndPush(bookings, currentRes, sheet, startRow, column); currentRes = null;
    }
  }
  if (currentRes) {
    let lastD = parseDateToDateObject(parseDate(dates[dates.length-1][0], sheet));
    if (lastD) { lastD.setDate(lastD.getDate()+1); currentRes.al = Utilities.formatDate(lastD,"GMT+0100","dd/MM/yyyy"); }
    validateAndPush(bookings, currentRes, sheet, startRow, column);
  }
  sheet.getRange(OUTPUT_ROW, column).setValue(JSON.stringify(bookings));
  reapplySundayBordersToColumn(sheet, column, dates);
}

function validateAndPush(list, res, sheet, row, col) {
  if (res.nome && res.dal && res.al && res.disposizione) {
    calculateBedCounts(res); list.push(res);
  } else {
    sheet.getRange(row, col).setBorder(true,true,true,true,false,false,ERROR_BORDER_COLOR,ERROR_BORDER_STYLE);
  }
}
function calculateBedCounts(res) {
  const d = res.disposizione.toLowerCase();
  const m = d.match(/(\d+)m/); if (m) res.matrimoniali = parseInt(m[1]);
  const s = d.match(/(\d+)s/); if (s) res.singoli = parseInt(s[1]);
  const c = d.match(/(\d+)c/); if (c) res.culle = parseInt(c[1]);
}
function extractArrangements(text) {
  let found = [], temp = text.replace(/\+/g,' ');
  [...VALID_BED_ARRANGEMENTS].sort((a,b)=>b.length-a.length).forEach(arr => {
    const reg = new RegExp(`\\b${arr}\\b`,'gi');
    if (reg.test(temp)) { found.push(arr); temp = temp.replace(reg,' '); }
  });
  return { dispositionString: found.join(" "), remainder: temp.trim() };
}
function cleanAndExtractNameAndNotes(text) {
  if (!text) return { name:"", notes:"" };
  const isNotName = [/^\d+$/,/\d{1,2}[\/.-]\d{1,2}/,/^storno$/i].some(p=>p.test(text));
  return isNotName ? { name:"", notes:text } : { name:text, notes:"" };
}
function parseDate(val, sheet) {
  if (val instanceof Date) return Utilities.formatDate(val,"GMT+0100","dd/MM/yyyy");
  const s = String(val).trim();
  if (/^\d{1,2}$/.test(s)) {
    const sn = sheet.getName().toLowerCase();
    let m = 0, y = new Date().getFullYear();
    for (const [name,idx] of Object.entries(MONTH_NAMES)) { if (sn.includes(name)) { m=idx; break; } }
    const ym = sn.match(/\d{4}/); if (ym) y = parseInt(ym[0]);
    return Utilities.formatDate(new Date(Date.UTC(y,m,parseInt(s))),"GMT+0100","dd/MM/yyyy");
  }
  return s.includes('/') ? s : null;
}
function parseDateToDateObject(dmy) {
  if (!dmy) return null;
  const p = dmy.split('/'); return new Date(p[2],p[1]-1,p[0]);
}
function reapplySundayBordersToColumn(sheet, col, dates) {
  for (let i = 0; i < dates.length; i++) {
    const d = dates[i][0];
    if (d instanceof Date && d.getDay()===0)
      sheet.getRange(FIRST_DATA_ROW+i,col).setBorder(true,null,true,null,false,false,YELLOW_BORDER_COLOR,SUNDAY_BORDER_STYLE);
  }
}


// =============================================================
// BATCH PROCESSING
// =============================================================
function startBatchProcessing() {
  PropertiesService.getUserProperties().setProperty(PROCESSING_STATE_KEY, JSON.stringify({sheetIndex:0,columnIndex:FIRST_CAMERA_COLUMN}));
  ScriptApp.newTrigger('processNextBatch').timeBased().after(1000).create();
  SpreadsheetApp.getUi().alert("Batch avviato.");
}
function processNextBatch() {
  const t0 = new Date().getTime();
  const state = JSON.parse(PropertiesService.getUserProperties().getProperty(PROCESSING_STATE_KEY)||'{}');
  const sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  for (let si = state.sheetIndex; si < sheets.length; si++) {
    const s = sheets[si]; if (EXCLUDED_SHEETS.includes(s.getName())) continue;
    for (let col = (si===state.sheetIndex?state.columnIndex:FIRST_CAMERA_COLUMN); col<=s.getLastColumn(); col++) {
      processSingleColumnBookings(s,col);
      if (new Date().getTime()-t0 > BATCH_TIME_LIMIT_MS) {
        PropertiesService.getUserProperties().setProperty(PROCESSING_STATE_KEY,JSON.stringify({sheetIndex:si,columnIndex:col+1})); return;
      }
    }
  }
  PropertiesService.getUserProperties().deleteProperty(PROCESSING_STATE_KEY);
  ScriptApp.getProjectTriggers().forEach(t=>{ if(t.getHandlerFunction()==='processNextBatch') ScriptApp.deleteTrigger(t); });
}
function applySundayBordersToAllSheetsManually() {
  SpreadsheetApp.getActiveSpreadsheet().getSheets().forEach(s => {
    if (EXCLUDED_SHEETS.includes(s.getName())) return;
    s.getRange(1,1,s.getMaxRows(),s.getMaxColumns()).setBorder(false,false,false,false,false,false);
    HEADER_RANGES.forEach(r => s.getRange(r).setBorder(true,true,true,true,false,false,BLACK_BORDER_COLOR,BLACK_BORDER_STYLE));
    const dates = s.getRange(FIRST_DATA_ROW,DATES_COLUMN,s.getLastRow()-FIRST_DATA_ROW+1,1).getValues();
    reapplySundayBordersToColumn(s,1,dates);
    for (let c=FIRST_CAMERA_COLUMN;c<=s.getLastColumn();c++) reapplySundayBordersToColumn(s,c,dates);
  });
}


// =============================================================
// JSON_ANNUALE — Entry point e trigger
// =============================================================
function aggiornaJSONAnnuale() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log("[JSON_ANNUALE] Avvio...");
  const anno         = rilevaAnno(ss);
  const segmenti     = estraiSegmenti(ss, anno);
  const prenotazioni = unisciMultiMese(segmenti);
  salvaJsonAnnuale(ss, prenotazioni, anno);
  Logger.log("[JSON_ANNUALE] ✅ " + prenotazioni.length + " prenotazioni per " + anno);
}

function aggiornaJSONAnnualeOnEdit(e) {
  if (!e || !e.source) return;
  const name = e.source.getActiveSheet().getName();
  if (EXCLUDED_SHEETS.includes(name) || !isFoglioMensile(name)) return;
  const row = e.range.getRow(), col = e.range.getColumn();
  if (col < JS_FIRST_CAM_COL || row < JS_FIRST_DATA_ROW || row >= OUTPUT_ROW) return;
  const props = PropertiesService.getScriptProperties();
  const last  = parseInt(props.getProperty("jsonAnnuale_lastRun")||"0");
  if (Date.now()-last < JS_DEBOUNCE_SEC*1000) return;
  props.setProperty("jsonAnnuale_lastRun", String(Date.now()));
  aggiornaJSONAnnuale();
}

/**
 * Debug: mostra info sui fogli e celle — esegui se JSON_ANNUALE è vuoto.
 */
function debugJSONAnnuale() {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const anno = rilevaAnno(ss);
  const logs = ["=== DEBUG JSON_ANNUALE ===","Anno: "+anno,""];

  ss.getSheets().forEach(s => {
    const nome = s.getName();
    logs.push(nome + (isFoglioMensile(nome)?" ✅":" —") + (EXCLUDED_SHEETS.includes(nome)?" [escluso]":""));
  });

  // Controlla il foglio di gennaio come campione
  const gennaio = ss.getSheetByName("Gennaio "+anno);
  if (gennaio) {
    logs.push("","── Gennaio "+anno+" ──");
    const maxCol = gennaio.getLastColumn();
    const camRow = gennaio.getRange(JS_HEADER_ROW,1,1,maxCol).getValues()[0];
    logs.push("Camere (riga "+JS_HEADER_ROW+"):");
    for (let c=JS_FIRST_CAM_COL-1;c<camRow.length;c++) {
      if (camRow[c]) logs.push("  col"+(c+1)+": "+camRow[c]);
    }
    // Prime 5 celle della prima camera
    const vals = gennaio.getRange(JS_FIRST_DATA_ROW,JS_FIRST_CAM_COL,5,1).getValues();
    const bgs  = gennaio.getRange(JS_FIRST_DATA_ROW,JS_FIRST_CAM_COL,5,1).getBackgrounds();
    const date = gennaio.getRange(JS_FIRST_DATA_ROW,1,5,1).getValues();
    logs.push("","Prime 5 celle cam. "+camRow[JS_FIRST_CAM_COL-1]+":");
    for (let i=0;i<5;i++) {
      const bg = bgs[i][0];
      logs.push("  riga"+(JS_FIRST_DATA_ROW+i)+": data="+date[i][0]+" val='"+vals[i][0]+"' bg='"+bg+"' neutro="+isNeutro(bg));
    }
  } else {
    logs.push("","⚠ 'Gennaio "+anno+"' non trovato!");
    logs.push("Nomi attesi: Gennaio "+anno+", Febbraio "+anno+", ...");
    logs.push("Controlla che i fogli abbiano ESATTAMENTE questo formato.");
  }

  const msg = logs.join("\n");
  Logger.log(msg);
  SpreadsheetApp.getUi().alert(msg);
}


// =============================================================
// STEP 1 — Rileva anno
// =============================================================
function rilevaAnno(ss) {
  for (const s of ss.getSheets()) {
    if (EXCLUDED_SHEETS.includes(s.getName())) continue;
    const m = s.getName().match(/\b(\d{4})\b/);
    if (m) return parseInt(m[1]);
  }
  return new Date().getFullYear();
}


// =============================================================
// STEP 2 — Estrai segmenti colorati
// =============================================================
/**
 * Estrae la disposizione letti da un testo cella (es. "MWT srl 2s" → "2s")
 * Restituisce null se nessuna disposizione trovata.
 */
function estraiDisposizione(testo) {
  if (!testo) return null;
  JS_DISPO_RE.lastIndex = 0;
  const found = [];
  let m;
  while ((m = JS_DISPO_RE.exec(testo)) !== null) {
    found.push(m[0].replace(/\s+/g,'').toLowerCase());
  }
  return found.length > 0 ? found.join(' ') : null;
}

function estraiSegmenti(ss, anno) {
  const segmenti = [];
  const ordine   = Object.keys(JS_MESI).map(m => m.charAt(0).toUpperCase()+m.slice(1)+" "+anno);

  for (const sheetName of ordine) {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) { Logger.log("[JS] Mancante: "+sheetName); continue; }

    const maxCol = sheet.getLastColumn();
    const maxRow = sheet.getLastRow();
    if (maxCol < JS_FIRST_CAM_COL || maxRow < JS_FIRST_DATA_ROW) continue;

    // Mappa colonna → camera
    const camRow = sheet.getRange(JS_HEADER_ROW,1,1,maxCol).getValues()[0];
    const camMap = {};
    for (let c=JS_FIRST_CAM_COL-1;c<camRow.length;c++) {
      const v = camRow[c];
      if (v!==null && v!=="" && v!==undefined)
        camMap[c+1] = typeof v==="number" ? String(Math.round(v)) : String(v).trim();
    }
    if (Object.keys(camMap).length===0) { Logger.log("[JS] Nessuna camera in "+sheetName); continue; }

    // Mappa riga → Date (colonna A)
    const dataFine = Math.min(maxRow, OUTPUT_ROW-1);
    if (dataFine < JS_FIRST_DATA_ROW) continue;
    const nRighe   = dataFine - JS_FIRST_DATA_ROW + 1;
    const dateVals = sheet.getRange(JS_FIRST_DATA_ROW,1,nRighe,1).getValues();
    const dateMap  = {};
    for (let i=0;i<dateVals.length;i++) {
      const v = dateVals[i][0];
      if (v instanceof Date) dateMap[JS_FIRST_DATA_ROW+i] = new Date(v);
    }
    const rows = Object.keys(dateMap).map(Number).sort((a,b)=>a-b);
    if (rows.length===0) { Logger.log("[JS] Nessuna data in "+sheetName); continue; }

    // Leggi blocco dati completo
    const firstRow  = rows[0], lastRow = rows[rows.length-1];
    const nDataRows = lastRow-firstRow+1;
    const firstCol  = JS_FIRST_CAM_COL;
    const nCols     = maxCol-firstCol+1;
    if (nCols<=0||nDataRows<=0) continue;

    const blockRange = sheet.getRange(firstRow,firstCol,nDataRows,nCols);
    const allVals    = blockRange.getValues();
    const allBgs     = blockRange.getBackgrounds();

    for (const col of Object.keys(camMap).map(Number)) {
      const camName  = camMap[col];
      const blockCol = col-firstCol;
      let cur = null;

      for (let ri=0;ri<rows.length;ri++) {
        const row      = rows[ri];
        const blockRow = row-firstRow;
        if (blockRow<0||blockRow>=nDataRows) continue;

        const bg    = normalizzaColore(allBgs[blockRow][blockCol]);
        const val   = allVals[blockRow][blockCol];
        const testo = (val!==null&&val!==undefined) ? String(val).trim() : "";
        const d     = dateMap[row];

        if (bg) {
          // Estrai disposizione dal testo della cella corrente
          const dispoCorrente = estraiDisposizione(testo);
          // Stessa cella: stesso colore E stessa disposizione (o nessuna disposizione esplicita)
          // Se la disposizione cambia → nuovo segmento (es. 2s → 1s stesso colore)
          const stessaDispo = !dispoCorrente || !cur || !cur.dispoIniziale || dispoCorrente === cur.dispoIniziale;
          if (cur && cur.colore===bg && stessaDispo) {
            cur.end = new Date(d);
            if (testo && !cur.testi.includes(testo)) cur.testi.push(testo);
          } else {
            if (cur) segmenti.push(cur);
            cur = { camera:camName, colore:bg, sheetName, start:new Date(d), end:new Date(d),
                    testi:testo?[testo]:[], dispoIniziale:dispoCorrente||null };
          }
        } else {
          if (cur) { segmenti.push(cur); cur=null; }
        }
      }
      if (cur) { segmenti.push(cur); cur=null; }
    }
    Logger.log("[JS] "+sheetName+": "+segmenti.length+" seg. totali finora");
  }
  return segmenti;
}


// =============================================================
// STEP 3 — Unisci multi-mese
// =============================================================
function unisciMultiMese(segmenti) {
  segmenti.sort((a,b) => {
    if (a.camera!==b.camera) return a.camera.localeCompare(b.camera,"it",{numeric:true});
    if (a.colore!==b.colore) return a.colore.localeCompare(b.colore);
    return a.start-b.start;
  });
  const merged=[], used=new Set();
  for (let i=0;i<segmenti.length;i++) {
    if (used.has(i)) continue;
    const s=segmenti[i];
    const base={camera:s.camera,colore:s.colore,start:new Date(s.start),end:new Date(s.end),testi:[...s.testi]};
    for (let j=i+1;j<segmenti.length;j++) {
      if (used.has(j)) continue;
      const t=segmenti[j];
      if (t.camera!==base.camera||t.colore!==base.colore) break;
      if (Math.round((t.start-base.end)/86400000)>2) break;
      // Non unire se la disposizione è diversa (cambio a cavallo di mese)
      const dispBase = estraiDisposizione(base.testi.join(' '));
      const dispT    = estraiDisposizione(t.testi.join(' '));
      if (dispBase && dispT && dispBase !== dispT) break;
      if (t.start>=base.start) {
        base.end=new Date(t.end);
        t.testi.forEach(tx=>{if(tx&&!base.testi.includes(tx))base.testi.push(tx);});
        used.add(j);
      }
    }
    const checkout=new Date(base.end); checkout.setDate(checkout.getDate()+1);
    const {nome,disposizione,note}=parsaTesti(base.testi);
    const letti=calcolaLetti(disposizione);
    merged.push({
      camera:base.camera, nome, dal:formatData(base.start), al:formatData(checkout),
      disposizione, note, backgroundColor:base.colore,
      matrimoniali:letti.m, singoli:letti.s, culle:letti.c, matrimonialiUS:letti.ms
    });
  }
  merged.sort((a,b)=>{
    const da=parseDataStr(a.dal),db=parseDataStr(b.dal);
    return da-db||a.camera.localeCompare(b.camera,"it",{numeric:true});
  });
  return merged;
}


// =============================================================
// STEP 4 — Salva foglio JSON_ANNUALE
// =============================================================
function salvaJsonAnnuale(ss, prenotazioni, anno) {
  let js = ss.getSheetByName(JS_SHEET_NAME);
  if (!js) { js=ss.insertSheet(JS_SHEET_NAME); ss.moveActiveSheet(ss.getNumSheets()); }

  // Pulisci il foglio in modo sicuro
  js.clear();

  // Riga 1: metadati
  js.getRange(1,1,1,5).setValues([["Anno:",anno,"Aggiornato:",
    Utilities.formatDate(new Date(),Session.getScriptTimeZone(),"dd/MM/yyyy HH:mm:ss"),
    "Prenotazioni: "+prenotazioni.length]]);
  js.getRange(1,1,1,5).setFontWeight("bold");

  // ── JSON spezzato per mese (righe 2-13) ──
  // Google Sheets ha limite 50.000 char/cella → scriviamo una riga per mese.
  // L'app web concatena le 12 righe per ricostruire l'array completo.
  const MESI_NOMI = ["Gennaio","Febbraio","Marzo","Aprile","Maggio","Giugno",
                     "Luglio","Agosto","Settembre","Ottobre","Novembre","Dicembre"];
  const perMese = {};
  MESI_NOMI.forEach(m => { perMese[m] = []; });
  prenotazioni.forEach(p => {
    const parts = (p.dal||"").split("/");
    if (parts.length === 3) {
      const mIdx = parseInt(parts[1]) - 1;
      if (mIdx >= 0 && mIdx < 12) perMese[MESI_NOMI[mIdx]].push(p);
    }
  });

  // Scrivi righe 2-13: una per mese
  MESI_NOMI.forEach(function(m, i) {
    const chunk = JSON.stringify(perMese[m]);
    js.getRange(2+i, 1).setValue(chunk);
    js.getRange(2+i, 1).setFontFamily("Courier New").setFontSize(9);
    js.getRange(2+i, 2).setValue(m + " (" + perMese[m].length + " pren.)");
  });

  // Riga 14: separatore
  js.getRange(14,1).setValue("— JSON per mese (righe 2-13) | Tabella leggibile (riga 15+) —");
  js.getRange(14,1).setFontColor("#888888").setFontStyle("italic");

  if (prenotazioni.length===0) {
    js.getRange(15,1).setValue("⚠ 0 prenotazioni. Usa 'Debug JSON_ANNUALE' per diagnosticare.");
    SpreadsheetApp.flush(); return;
  }

  // Tabella leggibile dalla riga 15
  const TABLE_ROW = 15;
  const cols=["camera","nome","dal","al","disposizione","matrimoniali","singoli","culle","matrimonialiUS","backgroundColor","note"];
  js.getRange(TABLE_ROW,1,1,cols.length).setValues([cols]);
  js.getRange(TABLE_ROW,1,1,cols.length).setFontWeight("bold").setBackground("#eeeeee");

  const righe=prenotazioni.map(function(p){return cols.map(function(c){return p[c]!==undefined&&p[c]!==null?p[c]:"";});});
  js.getRange(TABLE_ROW+1,1,righe.length,cols.length).setValues(righe);

  const colBg=cols.indexOf("backgroundColor")+1;
  for (var i=0;i<prenotazioni.length;i++) {
    var bg=String(prenotazioni[i].backgroundColor||"").trim();
    if (bg&&bg.startsWith("#")&&!isNeutro(bg)) {
      try { js.getRange(TABLE_ROW+1+i,colBg).setBackground(bg); } catch(e) {}
    }
  }
  try { js.autoResizeColumns(1,cols.length); } catch(e) {}
  SpreadsheetApp.flush();
}


// =============================================================
// HELPERS — Parsing testi
// =============================================================
function parsaTesti(testi) {
  let nome="", disposizione="", noteArr=[];
  for (const t of testi) {
    const clean=(t||"").trim();
    if (!clean||JS_SKIP_RE.test(clean)) continue;
    JS_DISPO_RE.lastIndex=0;
    const found=[]; let m;
    while ((m=JS_DISPO_RE.exec(clean))!==null) found.push(m[0].replace(/\s+/g,"").toLowerCase());
    if (found.length&&!disposizione) disposizione=found.join(" ");
    JS_DISPO_RE.lastIndex=0;
    const nomePart=clean.replace(JS_DISPO_RE,"").replace(/\s+/g," ").trim().replace(/^[-\/\s]+|[-\/\s]+$/g,"").trim();
    if (nomePart&&!nome&&!JS_SKIP_RE.test(nomePart)) nome=nomePart;
    else if (nomePart&&nomePart!==nome&&nomePart!==disposizione) noteArr.push(nomePart);
  }
  return { nome:nome||"???", disposizione:disposizione||"ND", note:noteArr.filter(n=>n).join("; ") };
}

function calcolaLetti(d) {
  const l={m:0,s:0,c:0,ms:0};
  if (!d||d==="ND") return l;
  d.split(/\s+/).forEach(p=>{
    const n=parseInt(p.match(/\d+/)?.[0]||"1");
    if (/m\/s$/i.test(p)||/ms$/i.test(p)) l.ms+=n;
    else if (/m$/i.test(p)) l.m+=n;
    else if (/s$/i.test(p)) l.s+=n;
    else if (/c$/i.test(p)) l.c+=n;
  });
  return l;
}


// =============================================================
// HELPERS — Date e colori
// =============================================================
function formatData(d) {
  if (!(d instanceof Date)) return "";
  return Utilities.formatDate(d,Session.getScriptTimeZone(),"dd/MM/yyyy");
}
function parseDataStr(s) {
  if (!s) return new Date(0);
  const [d,m,y]=s.split("/").map(Number); return new Date(y,m-1,d);
}
function normalizzaColore(hex) {
  if (!hex||hex==="") return null;
  const n=hex.toLowerCase().trim();
  return isNeutro(n) ? null : n;
}
function isNeutro(hex) {
  if (!hex||hex==="") return true;
  const n=hex.toLowerCase().trim();
  if (JS_SFONDO_NEUTRI.includes(n)) return true;
  if (n==="#000000"||n==="#ffffffff") return true;
  // Varianti di bianco: #fff, #ffffff, #fffffe ecc.
  if (/^#f{3,}$/i.test(n)) return true;
  return false;
}
function isFoglioMensile(nome) {
  if (EXCLUDED_SHEETS.includes(nome)) return false;
  const m=nome.match(/^([A-Za-zÀ-ÖØ-öø-ÿ]+)\s+(\d{4})$/i);
  return m ? (m[1].toLowerCase() in JS_MESI) : false;
}
