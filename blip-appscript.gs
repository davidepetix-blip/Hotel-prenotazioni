// =============================================================
// SCRIPT UNIFICATO — Prenotazioni + JSON_ANNUALE + Bridge Blip
// Versione: 2026-05-03
// =============================================================
// ISTRUZIONI:
//   1. Sostituisci TUTTO il contenuto dell'Apps Script con questo
//   2. Salva (Ctrl+S)
//   3. Menu "🏨 Script Prenotazioni" → "🔄 Rigenera JSON_ANNUALE"
//   4. Se il foglio è vuoto → "🔍 Debug JSON_ANNUALE" per diagnosticare
//   5. Ridistribuisci la Web App (stessa URL, stesso deploy) dopo ogni modifica
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
const BLIP_ID_ROW          = 46; // ← NUOVO: riga dove vengono scritti i BLIP_ID per colonna
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

// ── Nomi mesi in ordine (usato da bridge e JSON_ANNUALE) ──
const MESI_NOMI_ARR = [
  "Gennaio","Febbraio","Marzo","Aprile","Maggio","Giugno",
  "Luglio","Agosto","Settembre","Ottobre","Novembre","Dicembre"
];


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
    .addSeparator()
    .addItem('🔗 Test Bridge: leggi ultimo log', 'testBridgeLog')
    .addSeparator()
    .addItem('📧 Installa trigger email Sicily Divide (ogni 10 min)', 'installaSicilyDivideTrigger')
    .addItem('📧 Mostra trigger esistenti', 'mostraTriggerEsistenti')
    .addItem('📧 Elimina TUTTI i trigger', 'eliminaTuttiITrigger')
    .addItem('📧 Rimuovi trigger email Sicily Divide', 'rimuoviSicilyDivideTrigger')
    .addItem('📧 Test elaborazione email (manuale)', 'testSicilyDivideEmail')
    .addItem('📧 Salva DATABASE_SHEET_ID', 'salvaDbSheetIdEmail')
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


// =============================================================
// WEB APP — routing GET / POST
// =============================================================

/**
 * GET → routing:
 *   • action "scrivi" | "cancella"  → bridge scrittura foglio grafico (da Blip)
 *   • altrimenti                    → rigenera JSON_ANNUALE (retrocompatibile)
 *
 * Il GET è usato dal browser (CORS-compatibile con Apps Script Web App).
 * Il POST rimane disponibile per chiamate server-side future.
 *
 * Parametri GET per scrivi/cancella:
 *   action, blipId, camera, dal, al, nome, disposizione, colore, note,
 *   vecchioDal?, vecchioAl?, vecchiaCamera?
 */
function doGet(e) {
  const action = e && e.parameter && e.parameter.action;
  if (action === 'scrivi' || action === 'cancella') {
    const p = e.parameter;
    const payload = {
      action:        action,
      blipId:        p.blipId        || '',
      camera:        p.camera        || '',
      dal:           p.dal           || '',
      al:            p.al            || '',
      nome:          p.nome          || '',
      disposizione:  p.disposizione  || '',
      colore:        p.colore        || '',
      note:          p.note          || '',
      vecchioDal:    p.vecchioDal    || '',
      vecchioAl:     p.vecchioAl    || '',
      vecchiaCamera: p.vecchiaCamera || '',
    };
    // Salva log ultima chiamata per diagnostica dal menu
    try {
      PropertiesService.getScriptProperties().setProperty(
        'bridge_last_log',
        new Date().toISOString() + '\n' +
        'action=' + action + ' blipId=' + payload.blipId +
        ' camera=' + payload.camera + ' dal=' + payload.dal + ' al=' + payload.al
      );
    } catch(e2) {}
    const result = scriviPrenotazioneSuFoglio(payload);
    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  }
  return _rigenera(e);
}

/**
 * POST → routing:
 *   • action "scrivi" | "cancella"  → bridge scrittura foglio grafico
 *   • altrimenti                    → rigenera JSON_ANNUALE (retrocompatibile)
 *
 * Payload scrivi/cancella:
 * {
 *   action:        "scrivi" | "cancella",
 *   blipId:        "PRE-2026-CAM3-001",
 *   camera:        "Camera 3",           // esatto come intestazione foglio
 *   dal:           "10/07/2026",         // dd/MM/yyyy — prima notte
 *   al:            "15/07/2026",         // dd/MM/yyyy — giorno checkout (non colorato)
 *   nome:          "Mario Rossi",
 *   disposizione:  "1m",
 *   colore:        "#ea9999",
 *   note:          "",
 *   vecchioDal:    "08/07/2026",         // solo per modifica date
 *   vecchioAl:     "13/07/2026",
 *   vecchiaCamera: "Camera 2"            // solo per modifica camera
 * }
 */
function doPost(e) {
  if (e && e.postData && e.postData.contents) {
    try {
      const p = JSON.parse(e.postData.contents);
      if (p.action === 'scrivi' || p.action === 'cancella') {
        const result = scriviPrenotazioneSuFoglio(p);
        return ContentService
          .createTextOutput(JSON.stringify(result))
          .setMimeType(ContentService.MimeType.JSON);
      }
    } catch(err) {
      // JSON non valido o action non riconosciuta → fallback a rigenera
      Logger.log('[doPost] Fallback rigenera: ' + err.message);
    }
  }
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
      .createTextOutput(JSON.stringify({ ok:true, action:'rigenera', prenotazioni:merged.length, ms }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    Logger.log('[WebApp] Errore: ' + err.message);
    return ContentService
      .createTextOutput(JSON.stringify({ ok:false, action:'rigenera', error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function notificaModificaDaApp(e) {
  segnaModifica();
  return _rigenera(e);
}


// =============================================================
// BRIDGE — scrittura foglio grafico da Blip
// =============================================================

/**
 * Entry point principale del bridge.
 * Gestisce scrivi, modifica (= cancella vecchio + scrivi nuovo) e cancella.
 * Chiamato da doPost dopo parsing e routing del payload.
 *
 * Ritorna { ok, log[], prenotazioni? } per la risposta HTTP.
 */
function scriviPrenotazioneSuFoglio(payload) {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const log = [];

  try {
    // ── Validazione ──────────────────────────────────────────
    if (!payload.blipId) throw new Error('blipId obbligatorio');
    if (!payload.camera) throw new Error('camera obbligatoria');
    if (!payload.dal)    throw new Error('dal obbligatorio (dd/MM/yyyy)');
    if (payload.action === 'scrivi' && !payload.al) throw new Error('al obbligatorio per azione scrivi');

    const dalDate = _parseDataGS(payload.dal);
    const alDate  = payload.al ? _parseDataGS(payload.al) : null;

    if (!dalDate) throw new Error('dal non valido: ' + payload.dal);
    if (payload.action === 'scrivi' && !alDate) throw new Error('al non valido: ' + payload.al);
    if (alDate && alDate <= dalDate) throw new Error('al deve essere successivo a dal');

    // ── Cancella vecchio range (modifica o cancellazione) ─────
    const haCambioDate   = payload.vecchioDal && payload.vecchioAl;
    const haCambioCamera = payload.vecchiaCamera && payload.vecchiaCamera !== payload.camera;

    if (payload.action === 'cancella' || haCambioDate || haCambioCamera) {
      const vCam = haCambioCamera ? payload.vecchiaCamera : payload.camera;
      const vDal = haCambioDate   ? _parseDataGS(payload.vecchioDal) : dalDate;
      const vAl  = haCambioDate   ? _parseDataGS(payload.vecchioAl)  : alDate;

      if (vDal && vAl) {
        log.push('── Cancellazione vecchio range (' + vCam + ' ' + payload.vecchioDal + '→' + payload.vecchioAl + ')');
        _cancellaRangeFoglio(ss, vCam, payload.blipId, vDal, vAl, log);
      }
    }

    // ── Solo cancellazione: aggiornamento chirurgico e termina ──
    if (payload.action === 'cancella') {
      SpreadsheetApp.flush();
      const annoCancella = dalDate.getFullYear();
      const aggiornato = _rimuoviDaJSON(ss, payload, annoCancella, log);
      if (!aggiornato) { aggiornaJSONAnnuale(); } // fallback
      segnaModifica();
      log.push('✓ Cancellazione completata');
      return { ok: true, log };
    }

    // ── Scrivi nuovo range ────────────────────────────────────
    log.push('── Scrittura nuovo range (' + payload.camera + ' ' + payload.dal + '→' + payload.al + ')');
    var celleScritteCount = _scriviRangeFoglio(ss, payload, dalDate, alDate, log);

    // ── Verifica che qualcosa sia stato scritto ───────────────
    if (celleScritteCount === 0) {
      const warnings = log.filter(function(l) { return l.indexOf('⚠') !== -1; }).join(' | ');
      throw new Error('Nessuna cella scritta sul foglio. ' + warnings);
    }

    // ── Aggiornamento chirurgico JSON_ANNUALE ─────────────────
    // OTTIMIZZAZIONE PERFORMANCE:
    // Invece di rileggere tutti i 12 fogli mensili (lento: 15-40s),
    // modifichiamo solo il/i record interessato/i nel JSON già esistente.
    // Il trigger rigenera5min rimane come safety net per aggiornamenti completi.
    SpreadsheetApp.flush();
    const anno = dalDate.getFullYear();
    const aggiornato = _aggiornamentoChirurgicoJSON(ss, payload, dalDate, alDate, anno, log);
    segnaModifica();

    if (!aggiornato) {
      // Fallback: rigenerazione completa (solo se l'aggiornamento chirurgico non riesce)
      log.push('⟳ Fallback: rigenerazione completa JSON_ANNUALE');
      const segmenti = estraiSegmenti(ss, anno);
      const merged   = unisciMultiMese(segmenti);
      salvaJsonAnnuale(ss, merged, anno);
      log.push('✓ JSON_ANNUALE rigenerato (' + merged.length + ' prenotazioni)');
      return { ok: true, action: 'scrivi', written: celleScritteCount, log, prenotazioni: merged.length };
    }

    log.push('✓ JSON_ANNUALE aggiornato (' + aggiornato.totale + ' prenotazioni)');
    return { ok: true, action: 'scrivi', written: celleScritteCount, log, prenotazioni: aggiornato.totale };

  } catch(err) {
    log.push('✗ Errore: ' + err.message);
    Logger.log('[Bridge] ' + err.message + '\n' + (err.stack || ''));
    return { ok: false, error: err.message, log };
  }
}

/**
 * Cancella le celle colorate di una prenotazione nel foglio grafico.
 * Gestisce prenotazioni multi-mese.
 * Usa BLIP_ID_ROW come guardia: non cancella celle che appartengono ad altri booking.
 */
function _cancellaRangeFoglio(ss, camera, blipId, dalDate, alDate, log) {
  const mesi = _mesiCoperti(dalDate, alDate);

  mesi.forEach(function(mese) {
    const sheet = ss.getSheetByName(mese.sheetName);
    if (!sheet) { log.push('  ⚠ Foglio non trovato: ' + mese.sheetName); return; }

    const col = _trovaCameraColonna(sheet, camera);
    if (!col) { log.push('  ⚠ Camera "' + camera + '" non trovata in ' + mese.sheetName); return; }

    // ── Guardia: verifica che le celle appartengano a questo blipId ──
    // Legge riga BLIP_ID_ROW per questa colonna e controlla la mappa.
    const idMap = _leggiBlipIdMap(sheet, col);
    if (Object.keys(idMap).length > 0 && !idMap[blipId]) {
      log.push('  ⚠ Skip ' + mese.sheetName + ': celle appartengono ad altro booking (non ' + blipId + ')');
      return;
    }

    const rows = _trovaRigheDate(sheet, mese.firstDay, mese.lastDay);
    if (rows.length === 0) { log.push('  ⚠ Nessuna riga date in ' + mese.sheetName); return; }

    const firstRow = rows[0];
    const nRows    = rows[rows.length - 1] - firstRow + 1;
    const range    = sheet.getRange(firstRow, col, nRows, 1);
    range.clearContent();
    range.setBackground(null); // reset a bianco/default

    // Rimuovi blipId dalla mappa riga BLIP_ID_ROW
    _aggiornaBlipIdRow46(sheet, col, blipId, null, null, true);

    log.push('  ✓ ' + mese.sheetName + ': ' + rows.length + ' celle cancellate (col ' + col + ')');
  });
}

/**
 * Scrive le celle colorate di una nuova prenotazione nel foglio grafico.
 * Gestisce prenotazioni multi-mese.
 * Prima cella del primo mese: "nome disposizione"
 * Celle successive: testo vuoto, stesso colore
 */
function _scriviRangeFoglio(ss, payload, dalDate, alDate, log) {
  const mesi = _mesiCoperti(dalDate, alDate);
  if (mesi.length === 0) { log.push('  ⚠ Nessun mese da colorare'); return 0; }

  var totaleRigheScritite = 0;

  const testoIniziale = [payload.nome, payload.disposizione]
    .map(function(s) { return (s || '').trim(); })
    .filter(Boolean)
    .join(' ');

  mesi.forEach(function(mese, idx) {
    const sheet = ss.getSheetByName(mese.sheetName);
    if (!sheet) { log.push('  ⚠ Foglio non trovato: ' + mese.sheetName); return; }

    const col = _trovaCameraColonna(sheet, payload.camera);
    if (!col) {
      // Aggiungi diagnostica: mostra le intestazioni trovate per aiutare il debug
      const maxCol = sheet.getLastColumn();
      var intestazioni = [];
      if (maxCol >= FIRST_CAMERA_COLUMN) {
        var hdr = sheet.getRange(HEADER_ROW_NUMBER, 1, 1, Math.min(maxCol, 40)).getValues()[0];
        for (var ci = FIRST_CAMERA_COLUMN - 1; ci < hdr.length; ci++) {
          if (hdr[ci] !== null && hdr[ci] !== '') intestazioni.push(String(hdr[ci]).trim());
        }
      }
      log.push('  ⚠ Camera "' + payload.camera + '" non trovata in ' + mese.sheetName +
               '. Intestazioni rilevate: [' + intestazioni.slice(0,10).join(', ') + ']');
      return;
    }

    const rows = _trovaRigheDate(sheet, mese.firstDay, mese.lastDay);
    if (rows.length === 0) {
      log.push('  ⚠ Nessuna riga date in ' + mese.sheetName +
               ' per ' + mese.firstDay.toLocaleDateString('it-IT') +
               '→' + mese.lastDay.toLocaleDateString('it-IT'));
      return;
    }

    const idMapEsistente = _leggiBlipIdMap(sheet, col);
    const altriBid = Object.keys(idMapEsistente).filter(function(k) { return k !== payload.blipId; });
    if (altriBid.length > 0) {
      log.push('  ⚠ Sovrapposizione: ' + payload.camera + ' in ' + mese.sheetName +
               ' ha già: ' + altriBid.join(', '));
    }

    const firstRow = rows[0];
    const nRows    = rows[rows.length - 1] - firstRow + 1;

    const valori = [];
    for (var i = 0; i < nRows; i++) {
      valori.push([(i === 0 && idx === 0) ? testoIniziale : '']);
    }

    const range = sheet.getRange(firstRow, col, nRows, 1);
    range.setValues(valori);
    range.setBackground(payload.colore || '#ea9999');

    if (payload.note && idx === 0) {
      sheet.getRange(firstRow, col).setNote(payload.note);
    }

    _aggiornaBlipIdRow46(sheet, col, payload.blipId, payload.dal, payload.al, false);

    const nRighe = Math.min(sheet.getLastRow(), OUTPUT_ROW - 1) - FIRST_DATA_ROW + 1;
    if (nRighe > 0) {
      const dates = sheet.getRange(FIRST_DATA_ROW, DATES_COLUMN, nRighe, 1).getValues();
      reapplySundayBordersToColumn(sheet, col, dates);
    }

    totaleRigheScritite += rows.length;
    log.push('  ✓ ' + mese.sheetName + ': ' + rows.length + ' celle colorate (col ' + col + ')');
  });

  return totaleRigheScritite;
}


// =============================================================
// BRIDGE — Helpers interni
// =============================================================

/**
 * Restituisce l'array dei mesi coperti dalla prenotazione,
 * con firstDay e lastDay da colorare in quel mese.
 *
 * NOTA: alDate è il giorno di checkout (NON colorato).
 *       L'ultimo giorno colorato è alDate - 1.
 *
 * Ogni elemento: { sheetName, firstDay, lastDay, isFirst }
 */
function _mesiCoperti(dalDate, alDate) {
  const result = [];

  // Ultimo giorno da colorare = checkout - 1
  const ultimoColorato = new Date(alDate);
  ultimoColorato.setDate(ultimoColorato.getDate() - 1);
  ultimoColorato.setHours(0, 0, 0, 0);

  if (ultimoColorato < dalDate) return result; // booking di 0 notti (impossibile ma sicuro)

  const cur  = new Date(dalDate.getFullYear(), dalDate.getMonth(), 1);
  const fine = new Date(ultimoColorato.getFullYear(), ultimoColorato.getMonth(), 1);

  while (cur <= fine) {
    const anno    = cur.getFullYear();
    const meseIdx = cur.getMonth();
    const sheetName = MESI_NOMI_ARR[meseIdx] + ' ' + anno;

    const firstDay = (meseIdx === dalDate.getMonth() && anno === dalDate.getFullYear())
      ? new Date(dalDate.getFullYear(), dalDate.getMonth(), dalDate.getDate())
      : new Date(anno, meseIdx, 1);

    const ultimoDiMese = new Date(anno, meseIdx + 1, 0);
    const lastDay = (meseIdx === ultimoColorato.getMonth() && anno === ultimoColorato.getFullYear())
      ? new Date(ultimoColorato.getFullYear(), ultimoColorato.getMonth(), ultimoColorato.getDate())
      : ultimoDiMese;

    firstDay.setHours(0, 0, 0, 0);
    lastDay.setHours(0, 0, 0, 0);

    result.push({ sheetName, firstDay, lastDay, isFirst: result.length === 0 });
    cur.setMonth(cur.getMonth() + 1);
  }

  return result;
}

/**
 * Trova la colonna (1-based) corrispondente al nome camera nell'intestazione del foglio.
 * Confronto case-insensitive e normalizzazione numeri (es. "3.0" → "3").
 * Ritorna null se non trovata.
 */
function _trovaCameraColonna(sheet, camera) {
  const maxCol  = sheet.getLastColumn();
  if (maxCol < FIRST_CAMERA_COLUMN) return null;
  const headers = sheet.getRange(HEADER_ROW_NUMBER, 1, 1, maxCol).getValues()[0];
  const camNorm = String(camera).trim().replace(/\.0$/, '').toLowerCase();

  for (var c = FIRST_CAMERA_COLUMN - 1; c < headers.length; c++) {
    const h = headers[c];
    if (h === null || h === '' || h === undefined) continue;
    const hStr = (typeof h === 'number' ? String(Math.round(h)) : String(h).trim())
                  .replace(/\.0$/, '').toLowerCase();
    if (hStr === camNorm) return c + 1; // 1-based
  }
  return null;
}

/**
 * Trova le righe (1-based) del foglio corrispondenti al range di date [firstDay, lastDay].
 * Legge la colonna date (DATES_COLUMN) dalle righe dati fino a OUTPUT_ROW - 1.
 * Ritorna array di numeri riga ordinati.
 */
function _trovaRigheDate(sheet, firstDay, lastDay) {
  const maxRow  = sheet.getLastRow();
  if (maxRow < FIRST_DATA_ROW) return [];

  const endRow  = Math.min(maxRow, OUTPUT_ROW - 1);
  const nRows   = endRow - FIRST_DATA_ROW + 1;
  if (nRows <= 0) return [];

  const dateVals = sheet.getRange(FIRST_DATA_ROW, DATES_COLUMN, nRows, 1).getValues();
  const result   = [];

  const f = new Date(firstDay); f.setHours(0, 0, 0, 0);
  const l = new Date(lastDay);  l.setHours(0, 0, 0, 0);

  for (var i = 0; i < dateVals.length; i++) {
    const v = dateVals[i][0];
    if (!(v instanceof Date)) continue;
    const d = new Date(v); d.setHours(0, 0, 0, 0);
    if (d >= f && d <= l) result.push(FIRST_DATA_ROW + i);
  }

  return result;
}

/**
 * Legge la mappa BLIP_ID dalla riga BLIP_ID_ROW per una colonna.
 * Formato cella: JSON {"PRE-2026-XXX": ["10/07/2026","15/07/2026"], ...}
 * Ritorna oggetto vuoto se la cella è vuota o non parsificabile.
 */
function _leggiBlipIdMap(sheet, col) {
  try {
    const val = String(sheet.getRange(BLIP_ID_ROW, col).getValue() || '').trim();
    if (!val) return {};
    // Rimuovi eventuale apostrofo iniziale (Google Sheets lo aggiunge a volte)
    const clean = val.startsWith("'") ? val.slice(1) : val;
    const parsed = JSON.parse(clean);
    return (parsed && typeof parsed === 'object' && !Array.isArray(parsed)) ? parsed : {};
  } catch(e) {
    return {};
  }
}

/**
 * Aggiorna la mappa BLIP_ID nella riga BLIP_ID_ROW per una colonna.
 *
 * rimuovi = false → aggiunge o aggiorna blipId con [dal, al]
 * rimuovi = true  → rimuove blipId dalla mappa
 *
 * Se la mappa risultante è vuota, svuota la cella.
 */
function _aggiornaBlipIdRow46(sheet, col, blipId, dal, al, rimuovi) {
  const idMap = _leggiBlipIdMap(sheet, col);

  if (rimuovi) {
    delete idMap[blipId];
  } else {
    idMap[blipId] = [dal, al];
  }

  const cell = sheet.getRange(BLIP_ID_ROW, col);
  if (Object.keys(idMap).length === 0) {
    cell.clearContent();
  } else {
    // Prefisso apostrofo per evitare che Google Sheets interpreti il JSON come formula
    cell.setValue("'" + JSON.stringify(idMap));
  }
}

/**
 * Parsa una data in formato "dd/MM/yyyy" → oggetto Date.
 * Alternativa a parseDataStr (che usa lo stesso formato ma è già definita sotto).
 * Gestisce anche oggetti Date passati direttamente.
 */
function _parseDataGS(val) {
  if (val instanceof Date) return new Date(val);
  if (!val) return null;
  const s = String(val).trim();
  if (!s) return null;
  const parts = s.split('/');
  if (parts.length !== 3) return null;
  const d = parseInt(parts[0]), m = parseInt(parts[1]) - 1, y = parseInt(parts[2]);
  if (isNaN(d) || isNaN(m) || isNaN(y)) return null;
  const result = new Date(y, m, d);
  result.setHours(0, 0, 0, 0);
  return result;
}

/**
 * Test bridge dal menu: mostra l'ultimo log del bridge nelle Properties.
 */
function testBridgeLog() {
  const props = PropertiesService.getScriptProperties();
  const log   = props.getProperty('bridge_last_log') || '(nessun log disponibile)';
  SpreadsheetApp.getUi().alert('📋 Ultimo log Bridge:\n\n' + log);
}


// =============================================================
// TRIGGER TIME-BASED — Rigenera JSON_ANNUALE automaticamente
// =============================================================

function installaTriggerAutomatico() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'rigenera5min') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('rigenera5min').timeBased().everyMinutes(5).create();
  SpreadsheetApp.getUi().alert('✅ Trigger installato: JSON_ANNUALE si aggiornerà ogni 5 minuti automaticamente.');
}

function rimuoviTriggerAutomatico() {
  var rimossi = 0;
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'rigenera5min') { ScriptApp.deleteTrigger(t); rimossi++; }
  });
  SpreadsheetApp.getUi().alert('Trigger rimosso (' + rimossi + ' eliminati).');
}

function rigenera5min() {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const anno  = new Date().getFullYear();
    const props = PropertiesService.getScriptProperties();
    const ora   = Date.now();

    const jsSheet    = ss.getSheetByName(JS_SHEET_NAME);
    var ultimaMod    = parseInt(props.getProperty('ultima_modifica_ts') || '0');

    if (jsSheet) {
      const a1Val = jsSheet.getRange('A1').getValue();
      if (typeof a1Val === 'string' && a1Val.includes('app:')) {
        const match = a1Val.match(/(\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2})/);
        if (match) {
          const tsApp = new Date(match[1]).getTime();
          if (tsApp > ultimaMod) ultimaMod = tsApp;
        }
      }
    }

    if (ora - ultimaMod > 10 * 60 * 1000) { Logger.log('[Trigger 5min] Nessuna modifica recente, skip'); return; }

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
  for (var i = FIRST_DATA_ROW-1; i < dates.length; i++) {
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
  var bookings = [], currentRes = null, startRow = -1;
  for (var i = 0; i < values.length; i++) {
    const bg        = bgs[i][0];
    const isColored = bg && bg !== "#ffffff" && bg.toLowerCase() !== "white";
    const cellValue = String(values[i][0] || "").trim();
    const extracted = extractArrangements(cellValue);
    const dispositionString = extracted.dispositionString;
    const remainder = extracted.remainder;
    const cleaned = cleanAndExtractNameAndNotes(remainder);
    const name = cleaned.name;
    const textNotes = cleaned.notes;
    const parsedDate = parseDate(dates[i][0], sheet);
    const cellUINote = (uiNotes[i][0] || "").trim();
    if (isColored) {
      const isNew = !currentRes || bg !== currentRes.backgroundColor ||
                    (dispositionString && (name !== currentRes.nome || dispositionString !== currentRes.disposizione));
      if (isNew) {
        if (currentRes) { currentRes.al = parsedDate; validateAndPush(bookings, currentRes, sheet, startRow, column); }
        if (name || dispositionString) {
          var n = []; if (cellUINote) n.push(cellUINote); if (textNotes && textNotes !== name) n.push(textNotes);
          currentRes = { camera:cameraHeader, nome:name, dal:parsedDate, note:n.join(" - "),
                         backgroundColor:bg, disposizione:dispositionString, matrimoniali:0, singoli:0, culle:0 };
          startRow = FIRST_DATA_ROW + i;
        } else { currentRes = null; }
      } else if (currentRes) {
        var a = [];
        if (cellUINote && !currentRes.note.includes(cellUINote)) a.push(cellUINote);
        if (remainder && remainder !== currentRes.nome && !currentRes.note.includes(remainder)) a.push(remainder);
        if (a.length) currentRes.note = (currentRes.note ? currentRes.note + " - " : "") + a.join(" - ");
      }
    } else if (currentRes) {
      currentRes.al = parsedDate; validateAndPush(bookings, currentRes, sheet, startRow, column); currentRes = null;
    }
  }
  if (currentRes) {
    var lastD = parseDateToDateObject(parseDate(dates[dates.length-1][0], sheet));
    if (lastD) { lastD.setDate(lastD.getDate()+1); currentRes.al = Utilities.formatDate(lastD,Session.getScriptTimeZone(),"dd/MM/yyyy"); }
    validateAndPush(bookings, currentRes, sheet, startRow, column);
  }
  // Scrivi JSON del booking in riga OUTPUT_ROW (45)
  // GUARDIA: non scrivere mai nella riga BLIP_ID_ROW (46) — contiene i BLIP_IDs
  const outputRow = Math.min(OUTPUT_ROW, BLIP_ID_ROW - 1); // mai superare riga 45
  sheet.getRange(outputRow, column).setValue(JSON.stringify(bookings));
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
  var found = [], temp = text.replace(/\+/g,' ');
  [...VALID_BED_ARRANGEMENTS].sort(function(a,b){return b.length-a.length;}).forEach(function(arr) {
    const reg = new RegExp('\\b'+arr+'\\b','gi');
    if (reg.test(temp)) { found.push(arr); temp = temp.replace(reg,' '); }
  });
  return { dispositionString: found.join(" "), remainder: temp.trim() };
}
function cleanAndExtractNameAndNotes(text) {
  if (!text) return { name:"", notes:"" };
  const isNotName = [/^\d+$/,/\d{1,2}[\/.-]\d{1,2}/,/^storno$/i].some(function(p){return p.test(text);});
  return isNotName ? { name:"", notes:text } : { name:text, notes:"" };
}
function parseDate(val, sheet) {
  if (val instanceof Date) return Utilities.formatDate(val,Session.getScriptTimeZone(),"dd/MM/yyyy");
  const s = String(val).trim();
  if (/^\d{1,2}$/.test(s)) {
    const sn = sheet.getName().toLowerCase();
    var m = 0, y = new Date().getFullYear();
    for (const name in MONTH_NAMES) { if (sn.includes(name)) { m=MONTH_NAMES[name]; break; } }
    const ym = sn.match(/\d{4}/); if (ym) y = parseInt(ym[0]);
    return Utilities.formatDate(new Date(y,m,parseInt(s)),Session.getScriptTimeZone(),"dd/MM/yyyy");
  }
  return s.includes('/') ? s : null;
}
function parseDateToDateObject(dmy) {
  if (!dmy) return null;
  const p = dmy.split('/'); return new Date(p[2],p[1]-1,p[0]);
}
function reapplySundayBordersToColumn(sheet, col, dates) {
  for (var i = 0; i < dates.length; i++) {
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
  for (var si = state.sheetIndex; si < sheets.length; si++) {
    const s = sheets[si]; if (EXCLUDED_SHEETS.includes(s.getName())) continue;
    for (var col = (si===state.sheetIndex?state.columnIndex:FIRST_CAMERA_COLUMN); col<=s.getLastColumn(); col++) {
      processSingleColumnBookings(s,col);
      if (new Date().getTime()-t0 > BATCH_TIME_LIMIT_MS) {
        PropertiesService.getUserProperties().setProperty(PROCESSING_STATE_KEY,JSON.stringify({sheetIndex:si,columnIndex:col+1})); return;
      }
    }
  }
  PropertiesService.getUserProperties().deleteProperty(PROCESSING_STATE_KEY);
  ScriptApp.getProjectTriggers().forEach(function(t){if(t.getHandlerFunction()==='processNextBatch') ScriptApp.deleteTrigger(t);});
}
function applySundayBordersToAllSheetsManually() {
  SpreadsheetApp.getActiveSpreadsheet().getSheets().forEach(function(s) {
    if (EXCLUDED_SHEETS.includes(s.getName())) return;
    s.getRange(1,1,s.getMaxRows(),s.getMaxColumns()).setBorder(false,false,false,false,false,false);
    HEADER_RANGES.forEach(function(r) { s.getRange(r).setBorder(true,true,true,true,false,false,BLACK_BORDER_COLOR,BLACK_BORDER_STYLE); });
    const dates = s.getRange(FIRST_DATA_ROW,DATES_COLUMN,s.getLastRow()-FIRST_DATA_ROW+1,1).getValues();
    reapplySundayBordersToColumn(s,1,dates);
    for (var c=FIRST_CAMERA_COLUMN;c<=s.getLastColumn();c++) reapplySundayBordersToColumn(s,c,dates);
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

function debugJSONAnnuale() {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const anno = rilevaAnno(ss);
  const logs = ["=== DEBUG JSON_ANNUALE ===","Anno: "+anno,""];

  ss.getSheets().forEach(function(s) {
    const nome = s.getName();
    logs.push(nome + (isFoglioMensile(nome)?" ✅":" —") + (EXCLUDED_SHEETS.includes(nome)?" [escluso]":""));
  });

  const gennaio = ss.getSheetByName("Gennaio "+anno);
  if (gennaio) {
    logs.push("","── Gennaio "+anno+" ──");
    const maxCol = gennaio.getLastColumn();
    const camRow = gennaio.getRange(JS_HEADER_ROW,1,1,maxCol).getValues()[0];
    logs.push("Camere (riga "+JS_HEADER_ROW+"):");
    for (var c=JS_FIRST_CAM_COL-1;c<camRow.length;c++) {
      if (camRow[c]) logs.push("  col"+(c+1)+": "+camRow[c]);
    }
    const vals = gennaio.getRange(JS_FIRST_DATA_ROW,JS_FIRST_CAM_COL,5,1).getValues();
    const bgs  = gennaio.getRange(JS_FIRST_DATA_ROW,JS_FIRST_CAM_COL,5,1).getBackgrounds();
    const date = gennaio.getRange(JS_FIRST_DATA_ROW,1,5,1).getValues();
    logs.push("","Prime 5 celle cam. "+camRow[JS_FIRST_CAM_COL-1]+":");
    for (var i=0;i<5;i++) {
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
  const sheets = ss.getSheets();
  for (var i=0; i<sheets.length; i++) {
    const s = sheets[i];
    if (EXCLUDED_SHEETS.includes(s.getName())) continue;
    const m = s.getName().match(/\b(\d{4})\b/);
    if (m) return parseInt(m[1]);
  }
  return new Date().getFullYear();
}


// =============================================================
// STEP 2 — Estrai segmenti colorati
// =============================================================
function estraiDisposizione(testo) {
  if (!testo) return null;
  JS_DISPO_RE.lastIndex = 0;
  const found = [];
  var m;
  while ((m = JS_DISPO_RE.exec(testo)) !== null) {
    found.push(m[0].replace(/\s+/g,'').toLowerCase());
  }
  return found.length > 0 ? found.join(' ') : null;
}

function estraiSegmenti(ss, anno) {
  const segmenti = [];
  const ordine   = Object.keys(JS_MESI).map(function(m) { return m.charAt(0).toUpperCase()+m.slice(1)+" "+anno; });

  for (var oi=0; oi<ordine.length; oi++) {
    const sheetName = ordine[oi];
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) { Logger.log("[JS] Mancante: "+sheetName); continue; }

    const maxCol = sheet.getLastColumn();
    const maxRow = sheet.getLastRow();
    if (maxCol < JS_FIRST_CAM_COL || maxRow < JS_FIRST_DATA_ROW) continue;

    const camRow = sheet.getRange(JS_HEADER_ROW,1,1,maxCol).getValues()[0];
    const camMap = {};
    for (var c=JS_FIRST_CAM_COL-1;c<camRow.length;c++) {
      const v = camRow[c];
      if (v!==null && v!=="" && v!==undefined)
        camMap[c+1] = typeof v==="number" ? String(Math.round(v)) : String(v).trim();
    }
    if (Object.keys(camMap).length===0) { Logger.log("[JS] Nessuna camera in "+sheetName); continue; }

    const dataFine = Math.min(maxRow, OUTPUT_ROW-1);
    if (dataFine < JS_FIRST_DATA_ROW) continue;
    const nRighe   = dataFine - JS_FIRST_DATA_ROW + 1;
    const dateVals = sheet.getRange(JS_FIRST_DATA_ROW,1,nRighe,1).getValues();
    const dateMap  = {};
    for (var i=0;i<dateVals.length;i++) {
      const v = dateVals[i][0];
      if (v instanceof Date) dateMap[JS_FIRST_DATA_ROW+i] = new Date(v);
    }
    const rows = Object.keys(dateMap).map(Number).sort(function(a,b){return a-b;});
    if (rows.length===0) { Logger.log("[JS] Nessuna data in "+sheetName); continue; }

    const firstRow  = rows[0], lastRow = rows[rows.length-1];
    const nDataRows = lastRow-firstRow+1;
    const firstCol  = JS_FIRST_CAM_COL;
    const nCols     = maxCol-firstCol+1;
    if (nCols<=0||nDataRows<=0) continue;

    const blockRange = sheet.getRange(firstRow,firstCol,nDataRows,nCols);
    const allVals    = blockRange.getValues();
    const allBgs     = blockRange.getBackgrounds();

    const camCols = Object.keys(camMap).map(Number);
    for (var ci=0; ci<camCols.length; ci++) {
      const col     = camCols[ci];
      const camName = camMap[col];
      const blockCol = col-firstCol;
      var cur = null;

      for (var ri=0;ri<rows.length;ri++) {
        const row      = rows[ri];
        const blockRow = row-firstRow;
        if (blockRow<0||blockRow>=nDataRows) continue;

        const bg    = normalizzaColore(allBgs[blockRow][blockCol]);
        const val   = allVals[blockRow][blockCol];
        const testo = (val!==null&&val!==undefined) ? String(val).trim() : "";
        const d     = dateMap[row];

        if (bg) {
          const dispoCorrente = estraiDisposizione(testo);
          const stessaDispo = !dispoCorrente || !cur || !cur.dispoIniziale || dispoCorrente === cur.dispoIniziale;
          if (cur && cur.colore===bg && stessaDispo) {
            cur.end = new Date(d);
            if (testo && !cur.testi.includes(testo)) cur.testi.push(testo);
          } else {
            if (cur) segmenti.push(cur);
            cur = { camera:camName, colore:bg, sheetName:sheetName, start:new Date(d), end:new Date(d),
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
  segmenti.sort(function(a,b) {
    if (a.camera!==b.camera) return a.camera.localeCompare(b.camera,"it",{numeric:true});
    if (a.colore!==b.colore) return a.colore.localeCompare(b.colore);
    return a.start-b.start;
  });
  const merged=[], used=new Set();
  for (var i=0;i<segmenti.length;i++) {
    if (used.has(i)) continue;
    const s=segmenti[i];
    const base={camera:s.camera,colore:s.colore,start:new Date(s.start),end:new Date(s.end),testi:[...s.testi]};
    for (var j=i+1;j<segmenti.length;j++) {
      if (used.has(j)) continue;
      const t=segmenti[j];
      if (t.camera!==base.camera||t.colore!==base.colore) break;
      if (Math.round((t.start-base.end)/86400000)>2) break;
      const dispBase = estraiDisposizione(base.testi.join(' '));
      const dispT    = estraiDisposizione(t.testi.join(' '));
      if (dispBase && dispT && dispBase !== dispT) break;
      if (t.start>=base.start) {
        base.end=new Date(t.end);
        t.testi.forEach(function(tx){if(tx&&!base.testi.includes(tx))base.testi.push(tx);});
        used.add(j);
      }
    }
    const checkout=new Date(base.end); checkout.setDate(checkout.getDate()+1);
    const parsed=parsaTesti(base.testi);
    const letti=calcolaLetti(parsed.disposizione);
    merged.push({
      camera:base.camera, nome:parsed.nome, dal:formatData(base.start), al:formatData(checkout),
      disposizione:parsed.disposizione, note:parsed.note, backgroundColor:base.colore,
      matrimoniali:letti.m, singoli:letti.s, culle:letti.c, matrimonialiUS:letti.ms
    });
  }
  merged.sort(function(a,b){
    const da=parseDataStr(a.dal),db=parseDataStr(b.dal);
    return da-db||a.camera.localeCompare(b.camera,"it",{numeric:true});
  });
  return merged;
}


// =============================================================
// STEP 4 — Salva foglio JSON_ANNUALE
// =============================================================
function salvaJsonAnnuale(ss, prenotazioni, anno) {
  var js = ss.getSheetByName(JS_SHEET_NAME);
  if (!js) { js=ss.insertSheet(JS_SHEET_NAME); ss.moveActiveSheet(ss.getNumSheets()); }

  js.clear();

  js.getRange(1,1,1,5).setValues([["Anno:",anno,"Aggiornato:",
    Utilities.formatDate(new Date(),Session.getScriptTimeZone(),"dd/MM/yyyy HH:mm:ss"),
    "Prenotazioni: "+prenotazioni.length]]);
  js.getRange(1,1,1,5).setFontWeight("bold");

  const perMese = {};
  MESI_NOMI_ARR.forEach(function(m) { perMese[m] = []; });
  prenotazioni.forEach(function(p) {
    const parts = (p.dal||"").split("/");
    if (parts.length === 3) {
      const mIdx = parseInt(parts[1]) - 1;
      if (mIdx >= 0 && mIdx < 12) perMese[MESI_NOMI_ARR[mIdx]].push(p);
    }
  });

  MESI_NOMI_ARR.forEach(function(m, i) {
    const chunk = JSON.stringify(perMese[m]);
    js.getRange(2+i, 1).setValue(chunk);
    js.getRange(2+i, 1).setFontFamily("Courier New").setFontSize(9);
    js.getRange(2+i, 2).setValue(m + " (" + perMese[m].length + " pren.)");
  });

  const fpVals = MESI_NOMI_ARR.map(function(m) {
    const pren = perMese[m];
    const json = JSON.stringify(pren);
    var hash = 0;
    for (var k = 0; k < json.length; k++) {
      hash = (hash * 31 + json.charCodeAt(k)) & 0xFFFFFF;
    }
    return [pren.length + ':' + hash.toString(16)];
  });
  js.getRange(2, 15, 12, 1).setValues(fpVals);

  js.getRange(14,1).setValue("— JSON per mese (righe 2-13) | Fingerprint col. O | Tabella leggibile (riga 15+) —");
  js.getRange(14,1).setFontColor("#888888").setFontStyle("italic");

  if (prenotazioni.length===0) {
    js.getRange(15,1).setValue("⚠ 0 prenotazioni. Usa 'Debug JSON_ANNUALE' per diagnosticare.");
    SpreadsheetApp.flush(); return;
  }

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
  Logger.log("[JSON_ANNUALE] Fingerprint scritti in O2:O13 (" + fpVals.length + " mesi)");
}


// =============================================================
// HELPERS — Parsing testi
// =============================================================
function parsaTesti(testi) {
  var nome="", disposizione="", noteArr=[];
  for (var ti=0; ti<testi.length; ti++) {
    const clean=(testi[ti]||"").trim();
    if (!clean||JS_SKIP_RE.test(clean)) continue;
    JS_DISPO_RE.lastIndex=0;
    const found=[]; var m;
    while ((m=JS_DISPO_RE.exec(clean))!==null) found.push(m[0].replace(/\s+/g,"").toLowerCase());
    if (found.length&&!disposizione) disposizione=found.join(" ");
    JS_DISPO_RE.lastIndex=0;
    const nomePart=clean.replace(JS_DISPO_RE,"").replace(/\s+/g," ").trim().replace(/^[-\/\s]+|[-\/\s]+$/g,"").trim();
    if (nomePart&&!nome&&!JS_SKIP_RE.test(nomePart)) nome=nomePart;
    else if (nomePart&&nomePart!==nome&&nomePart!==disposizione) noteArr.push(nomePart);
  }
  return { nome:nome||"???", disposizione:disposizione||"ND", note:noteArr.filter(function(n){return n;}).join("; ") };
}

function calcolaLetti(d) {
  const l={m:0,s:0,c:0,ms:0};
  if (!d||d==="ND") return l;
  d.split(/\s+/).forEach(function(p){
    const n=parseInt((p.match(/\d+/)||["1"])[0]);
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
  const parts=s.split("/").map(Number); return new Date(parts[2],parts[1]-1,parts[0]);
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
  if (/^#f{3,}$/i.test(n)) return true;
  return false;
}
function isFoglioMensile(nome) {
  if (EXCLUDED_SHEETS.includes(nome)) return false;
  const m=nome.match(/^([A-Za-zÀ-ÖØ-öø-ÿ]+)\s+(\d{4})$/i);
  return m ? (m[1].toLowerCase() in JS_MESI) : false;
}


// =============================================================
// AGGIORNAMENTO CHIRURGICO JSON_ANNUALE
// =============================================================
// Invece di rileggere tutti i 12 fogli mensili (15-40s per foglio
// con molte prenotazioni), legge il JSON esistente da B2:B13,
// rimuove i record della camera/periodo modificato, estrae solo
// il mese interessato dal foglio, e riscrive il JSON aggiornato.
//
// Tempo: ~1-3s invece di 15-40s.
// Safety net: trigger rigenera5min fa la rigenerazione completa.
// =============================================================

/**
 * Aggiorna chirurgicamente JSON_ANNUALE dopo una scrittura.
 * Legge il JSON corrente, sostituisce i record della prenotazione
 * modificata con i dati freschi letti dal solo foglio mensile.
 *
 * @returns {{ totale: number }} oppure null se fallisce
 */
function _aggiornamentoChirurgicoJSON(ss, payload, dalDate, alDate, anno, log) {
  try {
    var t0 = Date.now();
    var js = ss.getSheetByName(JS_SHEET_NAME);
    if (!js) return null;

    // ── 1. Leggi JSON corrente da colonna B, righe 2-13 ─────
    // IMPORTANTE: la colonna A contiene i label (nomi mesi), il JSON è in colonna B
    var jsonValues = js.getRange(2, 1, 12, 1).getValues(); // colonna A riga 2-13 (JSON)
    var perMese = {};
    var totaleOrig = 0;
    MESI_NOMI_ARR.forEach(function(m, i) {
      try {
        perMese[m] = JSON.parse(jsonValues[i][0] || '[]');
        totaleOrig += perMese[m].length;
      } catch(e) { perMese[m] = []; }
    });

    // ── 2. Determina i mesi interessati dalla prenotazione ───
    var mesiDaAggiornare = _getMesiRange(dalDate, alDate);

    // ── 3. Per ogni mese interessato: rileggi dal foglio ─────
    // Rimuovi i record con la stessa camera e blipId dal JSON corrente,
    // poi aggiungi i record freschi letti dal foglio mensile.
    var camera = payload.camera;
    var blipId = payload.blipId;

    for (var mi = 0; mi < mesiDaAggiornare.length; mi++) {
      var meseNome = mesiDaAggiornare[mi];
      // Rimuovi record esistenti per questa camera+blipId nel mese
      perMese[meseNome] = (perMese[meseNome] || []).filter(function(p) {
        return !(p.camera === camera ||
                 (p.blipId && p.blipId === blipId));
      });
      // Leggi segmenti freschi dal foglio mensile per questa camera
      var nuoviRecord = _estraiSegmentiCamera(ss, meseNome + ' ' + anno, camera, blipId);
      perMese[meseNome] = perMese[meseNome].concat(nuoviRecord);
      // Ordina per data
      perMese[meseNome].sort(function(a,b) {
        return parseDataStr(a.dal) - parseDataStr(b.dal);
      });
    }

    // ── 4. Riscrivi solo le righe dei mesi modificati ────────
    var aggiornamenti = [];
    var totaleFin = 0;
    var fpValsNew = js.getRange(2, 15, 12, 1).getValues(); // leggi fingerprint esistenti

    MESI_NOMI_ARR.forEach(function(m, i) {
      totaleFin += perMese[m].length;
      var isMeseModificato = mesiDaAggiornare.indexOf(m) !== -1;
      if (isMeseModificato) {
        var chunk = JSON.stringify(perMese[m]);
        aggiornamenti.push({ row: 2 + i, json: chunk, mese: m, n: perMese[m].length });
        // Ricalcola fingerprint per questo mese
        var hash = 0;
        for (var k = 0; k < chunk.length; k++) {
          hash = (hash * 31 + chunk.charCodeAt(k)) & 0xFFFFFF;
        }
        fpValsNew[i][0] = perMese[m].length + ':' + hash.toString(16);
      }
    });

    // Scrivi JSON aggiornato (solo i mesi modificati) — colonna A
    for (var ai = 0; ai < aggiornamenti.length; ai++) {
      var agg = aggiornamenti[ai];
      js.getRange(agg.row, 1).setValue(agg.json);
      js.getRange(agg.row, 2).setValue(agg.mese + ' (' + agg.n + ' pren.)');
    }
    js.getRange(2, 15, 12, 1).setValues(fpValsNew);
    // Aggiorna intestazione con timestamp
    js.getRange(1, 4).setValue(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss'));
    js.getRange(1, 5).setValue('Prenotazioni: ' + totaleFin);

    SpreadsheetApp.flush();
    var ms = Date.now() - t0;
    log.push('⚡ Aggiornamento chirurgico: ' + aggiornamenti.length + ' mes' + (aggiornamenti.length===1?'e':'i') + ' aggiornati in ' + ms + 'ms');
    return { totale: totaleFin };

  } catch(e) {
    log.push('⚠ Aggiornamento chirurgico fallito (' + e.message + ') — fallback a rigenerazione completa');
    Logger.log('[Chirurgico] ' + e.message + '\n' + (e.stack || ''));
    return null;
  }
}

/**
 * Rimuove una prenotazione dal JSON_ANNUALE senza rileggere i fogli.
 */
function _rimuoviDaJSON(ss, payload, anno, log) {
  try {
    var js = ss.getSheetByName(JS_SHEET_NAME);
    if (!js) return false;
    var jsonValues = js.getRange(2, 1, 12, 1).getValues(); // colonna A (JSON)
    var camera = payload.camera;
    var blipId = payload.blipId;
    var dalDate = _parseDataGS(payload.vecchioDal || payload.dal);
    var alDate  = _parseDataGS(payload.vecchioAl  || payload.al);
    if (!dalDate) return false;
    var mesiDaAggiornare = _getMesiRange(dalDate, alDate || dalDate);
    var fpVals = js.getRange(2, 15, 12, 1).getValues();
    var totale = 0;

    MESI_NOMI_ARR.forEach(function(m, i) {
      var arr;
      try { arr = JSON.parse(jsonValues[i][0] || '[]'); } catch(e) { arr = []; }
      totale += arr.length;
      if (mesiDaAggiornare.indexOf(m) === -1) return;
      var before = arr.length;
      arr = arr.filter(function(p) {
        return !(p.camera === camera || (p.blipId && p.blipId === blipId));
      });
      if (arr.length === before) return; // nessuna modifica
      totale -= (before - arr.length);
      var chunk = JSON.stringify(arr);
      js.getRange(2 + i, 1).setValue(chunk);
      js.getRange(2 + i, 2).setValue(m + ' (' + arr.length + ' pren.)');
      var hash = 0;
      for (var k = 0; k < chunk.length; k++) hash = (hash * 31 + chunk.charCodeAt(k)) & 0xFFFFFF;
      fpVals[i][0] = arr.length + ':' + hash.toString(16);
    });

    js.getRange(2, 15, 12, 1).setValues(fpVals);
    js.getRange(1, 4).setValue(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss'));
    SpreadsheetApp.flush();
    log.push('⚡ Rimozione chirurgica da JSON_ANNUALE');
    return true;
  } catch(e) {
    log.push('⚠ Rimozione chirurgica fallita: ' + e.message);
    return false;
  }
}

/**
 * Estrae i record di una singola camera dal foglio mensile.
 * Molto più veloce di estraiSegmenti() che legge tutte le camere.
 */
function _estraiSegmentiCamera(ss, sheetName, cameraNome, blipId) {
  try {
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) return [];
    var maxCol = sheet.getLastColumn();
    var maxRow = sheet.getLastRow();
    if (maxCol < JS_FIRST_CAM_COL || maxRow < JS_FIRST_DATA_ROW) return [];

    // Trova la colonna della camera
    var camRow = sheet.getRange(JS_HEADER_ROW, 1, 1, maxCol).getValues()[0];
    var colIdx = -1;
    for (var c = JS_FIRST_CAM_COL - 1; c < camRow.length; c++) {
      var v = camRow[c];
      var nome = typeof v === 'number' ? String(Math.round(v)) : String(v).trim();
      if (nome === cameraNome) { colIdx = c + 1; break; }
    }
    if (colIdx === -1) return [];

    // Leggi BLIP_ID dalla riga 46 per questa colonna
    var blipIdInCell = '';
    try {
      blipIdInCell = sheet.getRange(BLIP_ID_ROW, colIdx).getValue() || '';
      if (blipIdInCell) {
        try {
          var map = JSON.parse(blipIdInCell);
          var keys = Object.keys(map);
          if (keys.length > 0) blipIdInCell = keys[0]; // prendi il primo BLIP_ID
        } catch(e) {}
      }
    } catch(e) {}

    // Leggi solo la colonna della camera
    var dataFine = Math.min(maxRow, OUTPUT_ROW - 1);
    var nRighe = dataFine - JS_FIRST_DATA_ROW + 1;
    if (nRighe <= 0) return [];

    var dateVals = sheet.getRange(JS_FIRST_DATA_ROW, 1, nRighe, 1).getValues();
    var colVals  = sheet.getRange(JS_FIRST_DATA_ROW, colIdx, nRighe, 1).getValues();
    var bgVals   = sheet.getRange(JS_FIRST_DATA_ROW, colIdx, nRighe, 1).getBackgrounds();

    // Ricostruisci segmenti per questa camera
    var segmenti = [];
    var cur = null;
    for (var i = 0; i < nRighe; i++) {
      var bg = (bgVals[i][0] || '').trim().toLowerCase();
      var txt = String(colVals[i][0] || '').trim();
      var dataCell = dateVals[i][0];
      if (!dataCell) continue;
      var dataRow = new Date(dataCell); dataRow.setHours(12, 0, 0, 0);
      var neutro = !bg || bg === '#ffffff' || bg === '#fffffe' || bg === 'white';
      if (neutro || !txt) {
        if (cur) { segmenti.push(cur); cur = null; }
        continue;
      }
      if (!cur || cur.colore !== bg) {
        if (cur) segmenti.push(cur);
        cur = { camera: cameraNome, colore: bg, start: dataRow, end: dataRow, testi: [txt] };
      } else {
        cur.end = dataRow;
        if (txt && cur.testi.indexOf(txt) === -1) cur.testi.push(txt);
      }
    }
    if (cur) segmenti.push(cur);

    // Converti in formato JSON_ANNUALE
    return segmenti.map(function(s) {
      var checkout = new Date(s.end); checkout.setDate(checkout.getDate() + 1);
      var parsed = parsaTesti(s.testi);
      var letti = calcolaLetti(parsed.disposizione);
      return {
        camera: s.camera, nome: parsed.nome,
        dal: formatData(s.start), al: formatData(checkout),
        s: s.start.toISOString(), e: checkout.toISOString(), // per Blip
        disposizione: parsed.disposizione, note: parsed.note,
        backgroundColor: s.colore,
        matrimoniali: letti.m, singoli: letti.s, culle: letti.c, matrimonialiUS: letti.ms,
        blipId: blipId || blipIdInCell || ''
      };
    });
  } catch(e) {
    Logger.log('[_estraiSegmentiCamera] ' + e.message);
    return [];
  }
}

/**
 * Restituisce i nomi dei mesi coperti da un range di date.
 * Ex: 28/04 → 03/05 → ['Aprile', 'Maggio']
 */
function _getMesiRange(dalDate, alDate) {
  var mesi = [];
  var cur = new Date(dalDate.getFullYear(), dalDate.getMonth(), 1);
  var fine = new Date(alDate.getFullYear(), alDate.getMonth(), 1);
  while (cur <= fine) {
    var nome = MESI_NOMI_ARR[cur.getMonth()];
    if (mesi.indexOf(nome) === -1) mesi.push(nome);
    cur.setMonth(cur.getMonth() + 1);
  }
  return mesi;
}


// =============================================================
// EMAIL RICHIESTE — Sicily Divide + form sito web
// =============================================================
// Flusso per ogni email non letta che corrisponde alla query:
//
//   Se DISPONIBILE:
//     → Crea pre-prenotazione grigia sul Gantt
//     → Invia email di notifica a Davide (stesso account)
//     → NON risponde al cliente (lo fa Davide personalmente)
//     → Aggiunge label "Blip-DaRispondere"
//
//   Se NON DISPONIBILE:
//     → Invia risposta automatica al cliente
//     → Aggiunge label "Blip-NonDisponibile"
//
//   In entrambi i casi:
//     → Marca come letta
//     → Logga su foglio EMAIL_LOG
//
// Account: davide.petix@gmail.com (stesso del login Blip)
// Trigger: ogni 10 minuti via ScriptApp
// =============================================================

const SD_SEARCH_QUERY   = 'is:unread from:info@sicilydevide.com OR from:booking@sicilydevide.com OR (is:unread subject:"Richiesta di preventivo soggiorno")';
const SD_LABEL_RISPONDERE  = 'Blip-DaRispondere';
const SD_LABEL_NON_DISP    = 'Blip-NonDisponibile';
const SD_LABEL_ELABORATA   = 'Blip-Elaborata';
const SD_LOG_SHEET         = 'EMAIL_LOG';
const SD_MAX_PER_RUN       = 10;
const SD_PRE_COLOR         = '#D9D9D9'; // grigio

// ── Entry point del trigger ───────────────────────────────────
function processEmailRequestsTrigger() {
  try { _processSicilyDivideEmails(); }
  catch(e) { Logger.log('❌ processEmailRequestsTrigger: ' + e.message); }
}

function _processSicilyDivideEmails() {
  var cfg     = _sdLoadConfig();
  var query   = cfg.emailSearchQuery || SD_SEARCH_QUERY;
  Logger.log('📧 Email: ricerca — ' + query);

  var threads = GmailApp.search(query, 0, SD_MAX_PER_RUN);
  if (!threads.length) { Logger.log('📧 Email: nessuna nuova richiesta'); return; }
  Logger.log('📧 Email: ' + threads.length + ' da elaborare');

  var labelRisp  = _sdEnsureLabel(SD_LABEL_RISPONDERE);
  var labelNonD  = _sdEnsureLabel(SD_LABEL_NON_DISP);
  var labelElab  = _sdEnsureLabel(SD_LABEL_ELABORATA);

  for (var i = 0; i < threads.length; i++) {
    try {
      _sdProcessThread(threads[i], labelRisp, labelNonD, labelElab, cfg);
      Utilities.sleep(600);
    } catch(e) {
      Logger.log('⚠ Thread ' + threads[i].getId() + ': ' + e.message);
    }
  }
}

function _sdProcessThread(thread, labelRisp, labelNonD, labelElab, cfg) {
  var msg     = thread.getMessages()[0];
  var body    = msg.getPlainBody() || msg.getBody().replace(/<[^>]+>/g,' ').replace(/&nbsp;/g,' ');
  var from    = msg.getFrom();
  var subject = msg.getSubject();
  var toEmail = (from.match(/[\w.+\-]+@[\w.\-]+\.\w+/) || [from])[0];

  Logger.log('📧 Elaboro: "' + subject + '" da ' + toEmail);

  // Parsa il formato Sicily Divide (e fallback formato Aruba)
  var parsed = _sdParseEmail(body);
  if (!parsed.valida) {
    Logger.log('⚠ Campi mancanti — marca letta e skip');
    thread.markRead();
    if (labelElab) labelElab.addToThread(thread);
    return;
  }

  var fmtD = function(d) { return d ? Utilities.formatDate(d, 'Europe/Rome', 'dd/MM/yyyy') : '?'; };
  Logger.log('📧 ' + parsed.nome + ' ' + fmtD(parsed.checkin) + '→' + fmtD(parsed.checkout) + ' ' + parsed.persone + ' ospiti');

  // Verifica disponibilità
  var avail = _checkAvailability(parsed.checkin, parsed.checkout);
  Logger.log('📧 Disponibilità: ' + avail.camereDisponibili.length + ' camere libere');

  var preBlipId = null;
  var roomSuggerita = _matchBestRoom(parsed, avail.camereDisponibili);

  if (avail.camereDisponibili.length > 0) {
    // ── DISPONIBILE: pre-prenotazione + notifica a Davide ────
    preBlipId = _createPreBookingOnSheet(parsed, roomSuggerita, cfg);
    _sdNotificaDavide(parsed, avail, roomSuggerita, preBlipId, subject, toEmail, from, cfg);
    if (labelRisp)  labelRisp.addToThread(thread);
    Logger.log('✅ Pre-prenotazione creata, notifica inviata a Davide');
  } else {
    // ── NON DISPONIBILE: risposta automatica al cliente ──────
    var risposta = _sdRispostaNoDisponibilita(parsed, cfg);
    GmailApp.sendEmail(toEmail, 'Re: ' + subject, '', {
      htmlBody: risposta,
      name:     cfg.hotelName || 'Il Borgo Montedoro',
      replyTo:  Session.getActiveUser().getEmail()
    });
    if (labelNonD)  labelNonD.addToThread(thread);
    Logger.log('✅ Risposta "non disponibile" inviata a ' + toEmail);
  }

  // Marca elaborata in entrambi i casi
  thread.markRead();
  if (labelElab) labelElab.addToThread(thread);

  // Log su foglio
  _sdLogEmail({
    data:        new Date(),
    mittente:    toEmail,
    nome:        parsed.nome,
    checkin:     fmtD(parsed.checkin),
    checkout:    fmtD(parsed.checkout),
    persone:     parsed.persone,
    disponibile: avail.camereDisponibili.length > 0,
    camera:      roomSuggerita ? roomSuggerita.nome : '',
    preBlipId:   preBlipId || '',
    stato:       avail.camereDisponibili.length > 0 ? 'notifica-davide' : 'risposta-no-disp'
  });
}

// ── Parser email Sicily Divide ────────────────────────────────
// Supporta anche il formato Aruba come fallback
function _sdParseEmail(body) {
  var lines  = body.replace(/\r\n/g,'\n').split('\n');
  var fields = {};

  for (var i = 0; i < lines.length; i++) {
    var line = lines[i];
    // Formato Sicily Divide: "Key: Value" (separatore ": ")
    var sepSD = line.indexOf(': ');
    // Formato Aruba: "Key : Value" (separatore " : ")
    var sepAR = line.indexOf(' : ');

    var sep = -1, offset = 2;
    if (sepSD >= 0 && (sepAR < 0 || sepSD <= sepAR)) { sep = sepSD; offset = 2; }
    else if (sepAR >= 0) { sep = sepAR; offset = 3; }
    if (sep < 0) continue;

    var key = line.slice(0, sep).trim().toLowerCase()
      .replace(/[àá]/g,'a').replace(/[èé]/g,'e').replace(/[ìí]/g,'i')
      .replace(/[òó]/g,'o').replace(/[ùú]/g,'u');
    fields[key] = line.slice(sep + offset).trim();
  }

  var get = function() {
    for (var k = 0; k < arguments.length; k++) {
      if (fields[arguments[k]]) return fields[arguments[k]];
    }
    return '';
  };

  // Date: supporta YYYY-MM-DD (Sicily Divide) e DD/MM/YYYY (Aruba)
  var parseDate = function(str) {
    if (!str) return null;
    // YYYY-MM-DD
    var m1 = str.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (m1) return new Date(+m1[1], +m1[2]-1, +m1[3], 12, 0, 0);
    // DD/MM/YYYY
    var m2 = str.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (m2) return new Date(+m2[3], +m2[2]-1, +m2[1], 12, 0, 0);
    return null;
  };

  var nome = get('name','nome e cognome','nome','nominativo');
  var cognome = get('surname','cognome');
  if (cognome && nome && !nome.includes(cognome)) nome = nome + ' ' + cognome;

  var checkin  = parseDate(get('check-in date','check-in','data check-in','arrivo','data di check-in'));
  var checkout = parseDate(get('check-out date','check-out','data check-out','partenza','data di check-out'));

  return {
    nome:     nome,
    email:    get('e-mail address','e-mail','email','indirizzo email'),
    telefono: get('phone','telefono','tel','cellulare'),
    persone:  parseInt(get('number of guests','numero di persone','persone','ospiti') || '0') || 1,
    rooms:    parseInt(get('number of rooms','numero camere','camere') || '1') || 1,
    messaggio: get('message','messaggio','richieste','note'),
    checkin, checkout,
    valida:   !!(checkin && checkout && nome),
  };
}

// ── Notifica interna a Davide (quando c'è disponibilità) ─────
function _sdNotificaDavide(parsed, avail, room, preBlipId, oggOriginal, emailCliente, fromOrig, cfg) {
  var fmtD = function(d) { return d ? Utilities.formatDate(d, 'Europe/Rome', 'dd/MM/yyyy') : '?'; };
  var notti = avail.notti || Math.round((parsed.checkout - parsed.checkin) / 86400000);
  var hn    = cfg.hotelName || 'Il Borgo Montedoro';
  var me    = Session.getActiveUser().getEmail();

  var htmlBody =
    '<div style="font-family:sans-serif;max-width:600px;border:2px solid #2d6a4f;border-radius:8px;overflow:hidden">' +
    '<div style="background:#2d6a4f;color:#fff;padding:14px 18px">' +
    '<b style="font-size:16px">📋 Nuova richiesta — DA RISPONDERE</b>' +
    '</div>' +
    '<div style="padding:16px 18px">' +
    '<table style="width:100%;border-collapse:collapse;font-size:14px">' +
    '<tr><td style="padding:4px 0;color:#666;width:140px">Cliente</td><td><b>' + parsed.nome + '</b></td></tr>' +
    '<tr><td style="padding:4px 0;color:#666">Email</td><td><a href="mailto:' + parsed.email + '">' + parsed.email + '</a></td></tr>' +
    '<tr><td style="padding:4px 0;color:#666">Telefono</td><td>' + (parsed.telefono || '—') + '</td></tr>' +
    '<tr><td style="padding:4px 0;color:#666">Check-in</td><td><b>' + fmtD(parsed.checkin) + '</b></td></tr>' +
    '<tr><td style="padding:4px 0;color:#666">Check-out</td><td><b>' + fmtD(parsed.checkout) + '</b> (' + notti + ' notti)</td></tr>' +
    '<tr><td style="padding:4px 0;color:#666">Ospiti</td><td>' + parsed.persone + (parsed.rooms > 1 ? ' · ' + parsed.rooms + ' camere' : '') + '</td></tr>' +
    (parsed.messaggio ? '<tr><td style="padding:4px 0;color:#666;vertical-align:top">Messaggio</td><td style="font-style:italic">"' + parsed.messaggio + '"</td></tr>' : '') +
    '</table>' +
    (room ? '<div style="margin-top:12px;padding:10px;background:#f0fdf4;border-radius:6px;font-size:13px">' +
    '✅ <b>Pre-prenotazione creata:</b> Camera ' + room.nome +
    (preBlipId ? ' · ID: <code>' + preBlipId + '</code>' : '') + '</div>' : '') +
    '<div style="margin-top:16px;padding:10px;background:#fef9c3;border-radius:6px;font-size:13px">' +
    '⚠ <b>Ricorda di rispondere al cliente</b> — la pre-prenotazione è in grigio sul Gantt.' +
    '</div>' +
    '<div style="margin-top:12px">' +
    '<a href="mailto:' + parsed.email + '?subject=Re: ' + oggOriginal + '" ' +
    'style="background:#2d6a4f;color:#fff;padding:8px 18px;border-radius:6px;text-decoration:none;font-size:13px;font-weight:bold">' +
    '✉ Rispondi al cliente</a>' +
    '</div>' +
    '</div>' +
    '</div>';

  GmailApp.sendEmail(me,
    '📋 [Blip] Da rispondere: ' + parsed.nome + ' ' + fmtD(parsed.checkin) + '→' + fmtD(parsed.checkout),
    'Nuova richiesta da rispondere. Vedi email HTML.',
    { htmlBody: htmlBody, name: 'Blip — ' + hn }
  );
}

// ── Risposta automatica al cliente (non disponibile) ─────────
function _sdRispostaNoDisponibilita(parsed, cfg) {
  var fmtD = function(d) { return d ? Utilities.formatDate(d, 'Europe/Rome', 'dd/MM/yyyy') : '?'; };
  var nome = (parsed.nome || '').split(' ')[0];
  var hn   = cfg.hotelName || 'Il Borgo Montedoro';
  var tel  = cfg.hotelTel  || '';

  return (
    'Gentile ' + nome + ',<br><br>' +
    'la ringraziamo per la sua richiesta di soggiorno presso <b>' + hn + '</b>.<br><br>' +
    'Purtroppo per il periodo richiesto ' +
    '(<b>' + fmtD(parsed.checkin) + ' – ' + fmtD(parsed.checkout) + '</b>) ' +
    'non abbiamo disponibilità.<br><br>' +
    'La invitiamo a contattarci per verificare date alternative:<br>' +
    (tel ? '📞 ' + tel + '<br>' : '') +
    '📧 ' + Session.getActiveUser().getEmail() + '<br><br>' +
    'Saremo lieti di accoglierla in un altro periodo.<br><br>' +
    'Cordiali saluti,<br>' +
    '<b>' + hn + '</b>'
  );
}

// ── Helpers ───────────────────────────────────────────────────
function _sdEnsureLabel(name) {
  try {
    var labels = GmailApp.getUserLabels();
    for (var i = 0; i < labels.length; i++) {
      if (labels[i].getName() === name) return labels[i];
    }
    return GmailApp.createLabel(name);
  } catch(e) { return null; }
}

function _sdLoadConfig() {
  try {
    var props  = PropertiesService.getScriptProperties();
    var dbId   = props.getProperty('DATABASE_SHEET_ID');
    if (!dbId) return {};
    var dbSS   = SpreadsheetApp.openById(dbId);
    var imp    = dbSS.getSheetByName('IMPOSTAZIONI');
    if (!imp) return {};
    var rows   = imp.getDataRange().getValues();
    var map    = {};
    rows.forEach(function(r) { if (r[0]) map[r[0]] = r[1]; });
    var bs = {};
    try { bs = JSON.parse(map['billSettings'] || '{}'); } catch(e) {}
    return {
      hotelName:        bs.hotelName    || map['hotelName']    || 'Il Borgo Montedoro',
      hotelTel:         bs.hotelTel     || map['hotelTel']     || '',
      emailSearchQuery: bs.emailSearchQuery || map['emailSearchQuery'] || '',
      geminiApiKey:     bs.geminiApiKey || map['geminiApiKey'] || '',
    };
  } catch(e) { return {}; }
}

function _sdLogEmail(entry) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var log = ss.getSheetByName(SD_LOG_SHEET);
  if (!log) {
    log = ss.insertSheet(SD_LOG_SHEET);
    log.appendRow(['Data','Mittente','Nome','Check-in','Check-out','Ospiti','Disponibile','Camera','BLIP_ID','Stato']);
    log.getRange(1,1,1,10).setFontWeight('bold').setBackground('#f3f4f6');
    log.setFrozenRows(1);
  }
  log.appendRow([
    entry.data, entry.mittente, entry.nome,
    entry.checkin, entry.checkout, entry.persone,
    entry.disponibile ? '✅' : '❌',
    entry.camera, entry.preBlipId, entry.stato
  ]);
}

// ── Setup trigger ─────────────────────────────────────────────
function installaSicilyDivideTrigger() {
  var ui = SpreadsheetApp.getUi();

  // Rimuovi solo eventuali duplicati del trigger email (non toccare gli altri)
  var duplicati = 0;
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'processEmailRequestsTrigger') {
      ScriptApp.deleteTrigger(t);
      duplicati++;
    }
  });

  // Installa il trigger email
  ScriptApp.newTrigger('processEmailRequestsTrigger')
    .timeBased().everyMinutes(10).create();

  var totale = ScriptApp.getProjectTriggers().length;
  ui.alert(
    '✅ Trigger email installato!\n\n' +
    'Trigger attivi ora: ' + totale + '/20\n' +
    (duplicati > 0 ? '(' + duplicati + ' duplicati rimossi)\n\n' : '\n') +
    'Le richieste da Sicily Divide saranno elaborate ogni 10 minuti.\n\n' +
    'Cosa succede:\n' +
    '• Disponibile → pre-prenotazione grigia sul Gantt + email di notifica a te\n' +
    '• Non disponibile → risposta automatica al cliente\n\n' +
    'Controlla il foglio EMAIL_LOG per lo storico.'
  );
}

function mostraTriggerEsistenti() {
  var triggers = ScriptApp.getProjectTriggers();
  var elenco = triggers.length === 0
    ? 'Nessun trigger installato.'
    : triggers.map(function(t, i) {
        return (i+1) + '. ' + t.getHandlerFunction() +
               ' — ogni ' + (t.getTriggerSource() === ScriptApp.TriggerSource.CLOCK ? 'X min' : 'evento');
      }).join('\n');
  SpreadsheetApp.getUi().alert('Trigger esistenti (' + triggers.length + '/20):\n\n' + elenco);
}

function eliminaTuttiITrigger() {
  var ui = SpreadsheetApp.getUi();
  var ok = ui.alert('⚠ Eliminare TUTTI i trigger?\nDovrai reinstallarli manualmente.', ui.ButtonSet.YES_NO);
  if (ok !== ui.Button.YES) return;
  var n = ScriptApp.getProjectTriggers().length;
  ScriptApp.getProjectTriggers().forEach(function(t) { ScriptApp.deleteTrigger(t); });
  ui.alert('✅ Eliminati ' + n + ' trigger.');
}

function rimuoviSicilyDivideTrigger() {
  var n = 0;
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'processEmailRequestsTrigger') { ScriptApp.deleteTrigger(t); n++; }
  });
  SpreadsheetApp.getUi().alert(n > 0 ? '✅ Trigger rimosso.' : 'Nessun trigger da rimuovere.');
}

function testSicilyDivideEmail() {
  _processSicilyDivideEmails();
  SpreadsheetApp.getUi().alert('✅ Elaborazione completata. Controlla:\n• Foglio EMAIL_LOG\n• La tua casella (label Blip-DaRispondere)\n• Il Gantt per le pre-prenotazioni');
}

function salvaDbSheetIdEmail() {
  var ui   = SpreadsheetApp.getUi();
  var resp = ui.prompt('Inserisci il DATABASE_SHEET_ID di Blip\n(in Blip → DevTools console → digita: DATABASE_SHEET_ID)');
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  PropertiesService.getScriptProperties().setProperty('DATABASE_SHEET_ID', resp.getResponseText().trim());
  ui.alert('✅ DATABASE_SHEET_ID salvato.');
}


// ── Pre-prenotazione nel DB (NON sul foglio grafico) ────────────
// Le pre-prenotazioni vivono solo nel foglio PRENOTAZIONI del DB.
// Non toccano il foglio grafico → non influenzano JSON_ANNUALE.
// Il motore di disponibilità le legge dal DB → camera risulta occupata.
// Quando Davide conferma dal drawer Blip, il bridge scrive sul foglio grafico.
function _createPreBookingOnSheet(parsed, room, cfg) {
  if (!room) return null;

  var props = PropertiesService.getScriptProperties();
  var dbId  = props.getProperty('DATABASE_SHEET_ID');
  if (!dbId) {
    Logger.log('⚠ Pre-prenotazione: DATABASE_SHEET_ID non configurato — usa menu "Salva DATABASE_SHEET_ID"');
    return null;
  }

  var blipId = 'PRE-' + new Date().getFullYear() + '-' + _randomHash();
  var fmtD   = function(d) { return Utilities.formatDate(d, 'Europe/Rome', 'dd/MM/yyyy'); };
  var anno   = parsed.checkin.getFullYear();
  var disp   = parsed.persone + 's';

  // Schema colonne PRENOTAZIONI (A:O = 15 colonne)
  // A:ID  B:CAMERA  C:NOME  D:DAL  E:AL  F:DISP  G:NOTE  H:COLORE
  // I:ANNO  J:FONTE  K:TS  L:DELETED  M:CLIENTE_ID  N:STATO_PREN  O:FONTE2
  var row = [
    blipId,
    room.nome,
    parsed.nome,
    fmtD(parsed.checkin),
    fmtD(parsed.checkout),
    disp,
    '⏳ PRE-PREN. form web — ' + (parsed.email || ''),
    '#D9D9D9',           // grigio — cambierà a verde quando Davide conferma
    String(anno),
    'form-web',
    new Date().toISOString(),
    '',                  // DELETED
    '',                  // CLIENTE_ID
    'pre',               // STATO_PRENOTAZIONE → Blip lo mostra tratteggiato
    'form-web',          // FONTE extra
  ];

  try {
    var dbSS   = SpreadsheetApp.openById(dbId);
    var prenSh = dbSS.getSheetByName('PRENOTAZIONI');
    if (!prenSh) {
      Logger.log('⚠ Foglio PRENOTAZIONI non trovato nel DB');
      return null;
    }
    prenSh.appendRow(row);
    Logger.log('✅ Pre-prenotazione nel DB: ' + blipId + ' cam.' + room.nome);
    return blipId;
  } catch(e) {
    Logger.log('⚠ _createPreBookingOnSheet: ' + e.message);
    return null;
  }
}

function _randomHash() {
  return Math.random().toString(36).slice(2, 6).toUpperCase() + '-' +
         Math.random().toString(36).slice(2, 6).toUpperCase();
}

function _matchBestRoom(parsed, camereDisponibili) {
  if (!camereDisponibili || !camereDisponibili.length) return null;
  // Filtra per numero ospiti
  var fits = camereDisponibili.filter(function(c) {
    return !c.maxGuests || c.maxGuests >= (parsed.persone || 1);
  });
  var pool = fits.length ? fits : camereDisponibili;
  // Preferisci camere numeriche (Scuola) per soggiorni brevi
  var scuola = pool.filter(function(c) { return /^\d+$/.test(c.nome) && parseInt(c.nome) <= 104; });
  pool = scuola.length ? scuola : pool;
  pool.sort(function(a, b) { return (a.maxGuests || 99) - (b.maxGuests || 99); });
  return pool[0] || null;
}
