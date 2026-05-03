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

    // ── Solo cancellazione: rigenera e termina ─────────────────
    if (payload.action === 'cancella') {
      SpreadsheetApp.flush();
      aggiornaJSONAnnuale();
      segnaModifica();
      log.push('✓ Cancellazione completata');
      return { ok: true, log };
    }

    // ── Scrivi nuovo range ────────────────────────────────────
    log.push('── Scrittura nuovo range (' + payload.camera + ' ' + payload.dal + '→' + payload.al + ')');
    _scriviRangeFoglio(ss, payload, dalDate, alDate, log);

    // ── Rigenera JSON_ANNUALE (legge i nuovi colori) ──────────
    SpreadsheetApp.flush(); // forza scrittura prima della lettura
    const anno     = dalDate.getFullYear();
    const segmenti = estraiSegmenti(ss, anno);
    const merged   = unisciMultiMese(segmenti);
    salvaJsonAnnuale(ss, merged, anno);
    segnaModifica();

    log.push('✓ JSON_ANNUALE aggiornato (' + merged.length + ' prenotazioni)');
    return { ok: true, log, prenotazioni: merged.length };

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
  if (mesi.length === 0) { log.push('  ⚠ Nessun mese da colorare'); return; }

  // Testo prima cella: "Mario Rossi 1m"  (o solo nome se manca disposizione)
  const testoIniziale = [payload.nome, payload.disposizione]
    .map(function(s) { return (s || '').trim(); })
    .filter(Boolean)
    .join(' ');

  mesi.forEach(function(mese, idx) {
    const sheet = ss.getSheetByName(mese.sheetName);
    if (!sheet) { log.push('  ⚠ Foglio non trovato: ' + mese.sheetName); return; }

    const col = _trovaCameraColonna(sheet, payload.camera);
    if (!col) { log.push('  ⚠ Camera "' + payload.camera + '" non trovata in ' + mese.sheetName); return; }

    const rows = _trovaRigheDate(sheet, mese.firstDay, mese.lastDay);
    if (rows.length === 0) { log.push('  ⚠ Nessuna riga date in ' + mese.sheetName); return; }

    // ── Verifica sovrapposizione con prenotazioni esistenti ─────
    // Controlla i colori esistenti: se una cella ha già un colore non neutro
    // che appartiene a un ALTRO blipId, segnala avviso senza bloccare.
    const idMapEsistente = _leggiBlipIdMap(sheet, col);
    const altriBid = Object.keys(idMapEsistente).filter(function(k) { return k !== payload.blipId; });
    if (altriBid.length > 0) {
      log.push('  ⚠ ATTENZIONE: colonna ' + payload.camera + ' in ' + mese.sheetName +
               ' ha già ' + altriBid.length + ' altri booking: ' + altriBid.join(', '));
    }

    const firstRow = rows[0];
    const nRows    = rows[rows.length - 1] - firstRow + 1;

    // ── Prepara valori: testo solo nella prima cella del primo mese ──
    const valori = [];
    for (var i = 0; i < nRows; i++) {
      valori.push([(i === 0 && idx === 0) ? testoIniziale : '']);
    }

    const range = sheet.getRange(firstRow, col, nRows, 1);
    range.setValues(valori);
    range.setBackground(payload.colore || '#ea9999');

    // ── Note (se presenti) → nella prima cella ──
    if (payload.note && idx === 0) {
      sheet.getRange(firstRow, col).setNote(payload.note);
    }

    // ── Aggiorna BLIP_ID in riga BLIP_ID_ROW ────────────────────
    _aggiornaBlipIdRow46(sheet, col, payload.blipId, payload.dal, payload.al, false);

    // ── Riapplica bordi domenica sulla colonna modificata ────────
    const nRighe = Math.min(sheet.getLastRow(), OUTPUT_ROW - 1) - FIRST_DATA_ROW + 1;
    if (nRighe > 0) {
      const dates = sheet.getRange(FIRST_DATA_ROW, DATES_COLUMN, nRighe, 1).getValues();
      reapplySundayBordersToColumn(sheet, col, dates);
    }

    log.push('  ✓ ' + mese.sheetName + ': ' + rows.length + ' celle colorate (col ' + col + ')');
  });
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
