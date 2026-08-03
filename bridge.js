// =============================================================
// bridge.js — Blip ↔ Foglio Grafico Google
// Build: 18.10.4.8.45.27.23.33 | Bridge v1
// =============================================================
// Sostituisce: writeBookingToSheet, clearBookingFromSheet,
//              segnalaModificaAdAppsScript, writeFragment,
//              clearFragment, getSheetIdMap, triggerAppsScriptUpdate
//
// Tutte le scritture sul foglio grafico avvengono via GET
// all'Apps Script Web App (CORS-compatibile, con fallback no-cors).
// Apps Script aggiorna JSON_ANNUALE prima di rispondere →
// nessun polling dopo la risposta CORS.
//
// API pubblica (usata da gantt.js):
//   bridgeSalva(newB, oldB?)   → nuova prenotazione o modifica
//   bridgeCancella(b)          → cancellazione
// =============================================================

const BLIP_VER_BRIDGE = '4';

// ─────────────────────────────────────────────────────────────────
// HELPERS INTERNI
// ─────────────────────────────────────────────────────────────────

/**
 * Converte un oggetto Date → "dd/MM/yyyy" per il foglio grafico.
 * Il foglio usa sempre questo formato — non ISO, non locale.
 */
function _dateToDMY(d) {
  if (!(d instanceof Date) || isNaN(d)) return '';
  const dd   = String(d.getDate()).padStart(2, '0');
  const mm   = String(d.getMonth() + 1).padStart(2, '0');
  const yyyy = d.getFullYear();
  return `${dd}/${mm}/${yyyy}`;
}

/**
 * Costruisce l'URL GET verso l'Apps Script Web App.
 * Legge webAppUrl da billSettings (campo già usato in precedenza).
 * Lancia eccezione se l'URL non è configurato.
 *
 * params: oggetto chiave→valore, vengono omessi i valori vuoti.
 */
function _bridgeUrl(params) {
  // Priorità: global caricato al login dal DB → localStorage → errore
  // window._blipWebAppUrl è impostato da loadBillSettingsDB() in onLoginSuccess
  // ed è condiviso su tutti i dispositivi senza configurazione manuale.
  const base = (window._blipWebAppUrl || '')
    || (typeof loadBillSettings === 'function' ? (loadBillSettings().webAppUrl || '') : '');
  const trimmed = base.trim();
  if (!trimmed) throw new Error('URL Web App non configurato. Vai in ⚙ Tariffe → "Web App URL" e salva — verrà condiviso automaticamente su tutti i dispositivi.');
  const qs = Object.entries(params)
    .filter(([, v]) => v !== null && v !== undefined && String(v).trim() !== '')
    .map(([k, v]) => `${encodeURIComponent(k)}=${encodeURIComponent(v)}`)
    .join('&');
  return `${trimmed}?${qs}&_ts=${Date.now()}`;
}

/**
 * Chiama l'Apps Script via GET.
 * - Primo tentativo: mode:'cors'  → leggiamo la risposta JSON
 * - Fallback:        mode:'no-cors' → fire-and-forget, risposta non leggibile
 *
 * Ritorna:
 *   { ok, log[], prenotazioni? }   se CORS riesce
 *   null                           se fallback no-cors (non sappiamo l'esito)
 *
 * Lancia Error se la chiamata fallisce completamente.
 */
async function _callBridge(url) {
  // ── Primo tentativo: CORS ──────────────────────────────────────
  try {
    const r = await fetch(url, { method: 'GET', mode: 'cors' });
    if (!r.ok) throw new Error(`HTTP ${r.status}`);
    const d = await r.json();
    return d; // { ok, log[], prenotazioni? }
  } catch (corsErr) {
    // CORS non disponibile (es. redirect OAuth) → no-cors fire-and-forget
    syncLog('📡 Bridge: risposta non leggibile (no-cors), polling attivo', 'syn');
  }

  // ── Fallback: no-cors ──────────────────────────────────────────
  try {
    await fetch(url, { method: 'GET', mode: 'no-cors' });
    return null; // chiamata partita ma risposta non leggibile
  } catch (e) {
    throw new Error('Apps Script irraggiungibile: ' + e.message);
  }
}

/**
 * Ricarica bookings[] da JSON_ANNUALE dopo una chiamata bridge.
 *
 * polling = false (CORS): Apps Script ha già aggiornato JSON_ANNUALE
 *   prima di rispondere → rileggi subito.
 *
 * polling = true (no-cors): non sappiamo quando Apps Script finisce
 *   → poll ogni 8s per max 3 tentativi (24s totali).
 */
async function _bridgeReload(polling, verify = null) {
  // ── Usa il foglio dell'anno corrente — JSON_ANNUALE esiste solo lì ──
  // annualSheets può contenere più anni (es. 2025, 2026).
  // find(e => e.sheetId) restituisce il primo, che potrebbe non avere JSON_ANNUALE.
  const currentYear = new Date().getFullYear();
  const entry = (typeof annualSheets !== 'undefined')
    ? (annualSheets.find(e => e.sheetId && e.year === currentYear)
       || annualSheets.find(e => e.sheetId))
    : null;
  if (!entry?.sheetId) {
    syncLog('⚠ Bridge reload: nessun foglio annuale disponibile', 'wrn');
    return false;
  }

  // Se il chiamante passa "verify", il valore di ritorno di _bridgeReload
  // riflette se la modifica attesa e' REALMENTE presente nei dati appena
  // riletti dal foglio (fresh, prima della guardia anti-flicker qui sotto)
  // — non solo se il reload in generale e' andato a buon fine. Senza questo
  // controllo, un salvataggio silenziosamente fallito poteva risultare
  // "confermato" solo perche' la stessa guardia lo ripreservava in bookings[]
  // in attesa che il foglio si aggiornasse (entro 2 ore) — mascherando il
  // fallimento reale invece di segnalarlo.
  let _verified = true;

  const _apply = async (fresh) => {
    // NOTA: verify() va valutata ANCHE quando fresh è vuoto/assente — es. la
    // verifica di una cancellazione è per definizione "la riga non c'è più",
    // quindi un array vuoto (o senza il nostro dbId) è un ESITO VALIDO da
    // verificare, non un errore di lettura da ignorare.
    _verified = verify ? verify(fresh || []) : true;
    if (!fresh?.length) return false;

    // ── Preserva prenotazioni inserite via app non ancora nel foglio ──
    // Caso tipico: bridge ha chiamato Apps Script, Apps Script ha risposto OK,
    // ma JSON_ANNUALE non è ancora stato rigenerato (latenza 2-5s) oppure
    // la scrittura è avvenuta ma il reload legge una versione cached.
    // Senza questa guardia, la prenotazione sparisce dal Gantt 5 secondi
    // dopo il salvataggio, per poi ricomparire al bgSync successivo.
    const PROTEZIONE_APP_MS = 2 * 60 * 60 * 1000; // 2 ore
    const appLocali = (typeof bookings !== 'undefined' ? bookings : []).filter(b => {
      if (b.fonte !== 'app' || !b.ts) return false;
      return (Date.now() - new Date(b.ts).getTime()) < PROTEZIONE_APP_MS;
    });

    let freshMerged = mergeMultiMonthBookings(
      fresh.filter(b => !isDeletedLocally(b.dbId))
    );

    // Reinserisce le prenotazioni app recenti assenti dal foglio
    for (const app of appLocali) {
      const giaPresente = freshMerged.some(f =>
        f.dbId === app.dbId ||
        (f.cameraName === app.cameraName &&
         Math.abs((f.s?.getTime?.() || 0) - (app.s?.getTime?.() || 0)) < 86400000)
      );
      if (!giaPresente) {
        freshMerged.push({ ...app, pending: true });
        syncLog(`🛡 Bridge reload: preservata "${app.n}" (app ${Math.round((Date.now()-new Date(app.ts).getTime())/60000)}min fa)`, 'syn');
      }
    }

    bookings = freshMerged;
    saveDbCache(bookings);
    render();
    return true;
  };

  if (!polling) {
    // CORS: rileggi immediatamente
    try {
      const fresh = await readJSONAnnuale(entry.sheetId);
      if (await _apply(fresh)) {
        syncLog(`✓ Bridge: Gantt aggiornato (${bookings.length} prenotazioni)`, 'ok');
      }
    } catch(e) {
      syncLog('⚠ Bridge reload: ' + e.message, 'wrn');
      return false;
    }
    return _verified;
  }

  // no-cors: polling
  for (let t = 1; t <= 3; t++) {
    await new Promise(r => setTimeout(r, 8000));
    try {
      const fresh = await readJSONAnnuale(entry.sheetId);
      const applied = await _apply(fresh); // side-effect: aggiorna bookings[] se ci sono dati
      if (applied) {
        syncLog(`✓ Bridge: Gantt aggiornato (${bookings.length} pren., poll ${t}/3)`, 'ok');
      }
      // _verified è indipendente da "applied" — è il vero segnale di successo
      // (va controllato anche quando fresh è vuoto, es. cancellazioni).
      if (_verified) return true;
    } catch(e) {
      syncLog(`⏳ Bridge poll ${t}/3: ${e.message}`, 'syn');
    }
  }

  syncLog('⚠ Bridge: JSON_ANNUALE non aggiornato entro 24s — premi 🔄', 'wrn');
  return false;
}


// ─────────────────────────────────────────────────────────────────
// API PUBBLICA
// ─────────────────────────────────────────────────────────────────

/**
 * Salva una prenotazione sul foglio grafico.
 * Usata sia per nuove prenotazioni sia per modifiche.
 *
 * Per le modifiche, passare il booking ORIGINALE come oldB:
 *   - se la camera è cambiata  → cancella dal vecchio foglio + scrivi nel nuovo
 *   - se le date sono cambiate → cancella il vecchio range + scrivi il nuovo
 *   - se solo il nome cambia   → sovrascrive senza cancellare (stesso range)
 *
 * La funzione gestisce anche il DB (dbUpsert) dopo il bridge.
 *
 * @param {Object} newB  - prenotazione da scrivere
 * @param {Object} [oldB] - prenotazione originale (solo per modifiche)
 */
async function bridgeSalva(newB, oldB = null) {
  // ── Garantisci dbId prima di chiamare il bridge ─────────────────
  if (!newB.dbId) {
    newB.dbId = genBookingId(newB.s.getFullYear());
  }

  const params = {
    action:       'scrivi',
    blipId:       newB.dbId,
    camera:       newB.cameraName,
    dal:          _dateToDMY(newB.s),
    al:           _dateToDMY(newB.e),
    nome:         newB.n,
    disposizione: newB.d,
    colore:       newB.c,
    note:         newB.note || '',
  };

  // ── Per modifica: includi dati vecchi se camera o date sono cambiate ──
  if (oldB) {
    const cambiataCamera = oldB.cameraName !== newB.cameraName;
    const cambiateDal    = Math.abs((oldB.s?.getTime() || 0) - newB.s.getTime()) > 43200000; // > 12h
    const cambiateAl     = Math.abs((oldB.e?.getTime() || 0) - newB.e.getTime()) > 43200000;

    if (cambiataCamera || cambiateDal || cambiateAl) {
      params.vecchioDal    = _dateToDMY(oldB.s);
      params.vecchioAl     = _dateToDMY(oldB.e);
      params.vecchiaCamera = cambiataCamera ? oldB.cameraName : '';
    }
  }

  syncLog(
    `📡 Bridge → scrivi: ${newB.n} cam.${newB.cameraName}` +
    ` ${params.dal}→${params.al}` +
    (params.vecchioDal ? ` (era ${params.vecchioDal}→${params.vecchioAl})` : ''),
    'syn'
  );

  // ── Blocca bgSync mentre il bridge è attivo (evita sovrascritture) ──
  if (typeof _lastFullSyncTs !== 'undefined') _lastFullSyncTs = Date.now();

  // Verifica che LA NOSTRA modifica specifica sia davvero presente nei dati
  // grezzi riletti dal foglio (per dbId + checkout finale, tollerante ai
  // frammenti multi-mese: nel frammento finale l'"al" coincide col checkout
  // vero). Usata per non dichiarare mai un salvataggio riuscito quando in
  // realtà non lo sappiamo con certezza (caso no-cors).
  const _verificaScrittura = (fresh) => (fresh || []).some(f =>
    f.dbId === newB.dbId &&
    Math.abs((f.e?.getTime?.() || 0) - newB.e.getTime()) < 43200000
  );

  // ── Chiama Apps Script ─────────────────────────────────────────
  const url  = _bridgeUrl(params);
  const resp = await _callBridge(url);

  if (resp === null) {
    // no-cors: non sappiamo l'esito dalla risposta HTTP → polling con verifica
    // esplicita sui dati riletti. FIX: prima si chiamava _bridgeReload(true)
    // e si proseguiva SEMPRE come se fosse andato tutto bene (nessun throw),
    // quindi il chiamante mostrava "✓ salvato sul foglio Google" anche quando
    // la scrittura non era mai realmente avvenuta (es. token OAuth scaduto a
    // metà sessione) — l'unico avviso reale restava sepolto nel pannello LOG,
    // mai mostrato come toast. Ora, se dopo 24s di polling la modifica non
    // risulta confermata sul foglio, solleviamo un errore esplicito.
    const confermato = await _bridgeReload(true, _verificaScrittura);
    if (!confermato) {
      throw new Error('scrittura non confermata sul foglio entro 24s (sessione scaduta o connessione instabile?) — verifica manualmente sul foglio Google prima di procedere');
    }
  } else if (resp.ok && resp.action === 'rigenera') {
    // Apps Script ha risposto ma ha eseguito _rigenera invece di scriviPrenotazioneSuFoglio.
    // Questo accade quando i parametri GET si perdono nel redirect OAuth di Google.
    // La cella NON è stata scritta sul foglio.
    syncLog('⚠ Bridge: Apps Script non ha ricevuto action=scrivi (redirect OAuth?) — cella NON scritta', 'wrn');
    // FIX: prima qui si mostrava solo un toast e si CONTINUAVA senza throw —
    // il chiamante procedeva comunque a mostrare "✓ salvato", sovrascrivendo
    // (stesso elemento <div id="toast">) l'avviso appena mostrato prima che
    // fosse leggibile. Ora solleviamo un errore: il chiamante mostrerà UN
    // solo toast, quello reale.
    throw new Error('sessione Google scaduta durante il salvataggio — la prenotazione NON è stata scritta sul foglio, riprova');
  } else if (resp.ok && resp.written > 0) {
    syncLog(
      `✓ Bridge scrivi OK — ${resp.written} celle scritte, ${resp.prenotazioni ?? '?'} pren.` +
      (resp.log?.length ? ' | ' + resp.log[resp.log.length - 1] : ''),
      'ok'
    );
    await _bridgeReload(false);
  } else if (resp.ok && !resp.written && resp.log) {
    // scriviPrenotazioneSuFoglio è partita ma non ha scritto celle (ok:false avrebbe lanciato eccezione,
    // ma gestiamo anche il caso in cui written=0 con ok:true per retrocompatibilità)
    const detail = (resp.log || []).filter(l => l.includes('⚠')).join('\n');
    syncLog('⚠ Bridge: nessuna cella scritta — ' + (detail || 'causa sconosciuta'), 'wrn');
    // FIX: vedi commento sopra — non dichiarare mai successo se non c'è stata
    // nessuna scrittura reale sul foglio.
    throw new Error('nessuna cella scritta sul foglio — ' + (detail || 'verifica il nome camera'));
  } else {
    // Apps Script ha risposto con ok:false → errore esplicito con dettagli
    const detail = (resp.log || []).join('\n');
    throw new Error('Bridge: ' + (resp.error || 'errore sconosciuto') + (detail ? '\n' + detail : ''));
  }

  // ── Aggiorna BLIP-DB (fire-and-forget, non blocca UI) ─────────
  // NOTA: si arriva qui SOLO se il blocco sopra non ha sollevato eccezioni,
  // cioè solo quando la scrittura sul foglio grafico è confermata riuscita.
  if (typeof DATABASE_SHEET_ID !== 'undefined' && DATABASE_SHEET_ID) {
    newB.ts = nowISO(); newB.fonte = 'app'; newB.fromSheet = true;
    dbUpsert(newB, 'app').catch(e => syncLog('⚠ DB upsert: ' + e.message, 'wrn'));
  }
}

/**
 * Cancella una prenotazione dal foglio grafico.
 * Richiede che b.dbId sia presente (chiave BLIP_ID nella riga 46).
 * Se dbId manca, logga un avviso e non chiama il bridge (la cella
 * non è collegata a Blip → non possiamo identificarla con certezza).
 *
 * @param {Object} b - prenotazione da eliminare
 */
async function bridgeCancella(b) {
  if (!b.dbId) {
    syncLog('⚠ Bridge cancella: nessun dbId — cella foglio non rimossa', 'wrn');
    showToast('⚠ Prenotazione non agganciata al foglio Google — rimuovi la cella manualmente', 'warning');
    return;
  }

  const params = {
    action: 'cancella',
    blipId: b.dbId,
    camera: b.cameraName,
    dal:    _dateToDMY(b.s),
    al:     _dateToDMY(b.e),
  };

  syncLog(`📡 Bridge → cancella: ${b.n} cam.${b.cameraName} ${params.dal}→${params.al}`, 'syn');

  if (typeof _lastFullSyncTs !== 'undefined') _lastFullSyncTs = Date.now();

  // Verifica che la cella sia REALMENTE sparita dal foglio (nessun frammento
  // con questo dbId e range ancora presente nei dati riletti).
  const _verificaCancellazione = (fresh) => !(fresh || []).some(f =>
    f.dbId === b.dbId &&
    Math.abs((f.s?.getTime?.() || 0) - b.s.getTime()) < 43200000
  );

  const url  = _bridgeUrl(params);
  const resp = await _callBridge(url);

  if (resp === null) {
    // Stesso fix di bridgeSalva: non dichiarare mai successo (nessun throw)
    // se dopo il polling non risulta confermato che la cella sia sparita.
    const confermato = await _bridgeReload(true, _verificaCancellazione);
    if (!confermato) {
      throw new Error('cancellazione non confermata sul foglio entro 24s (sessione scaduta o connessione instabile?) — verifica manualmente');
    }
  } else if (resp.ok) {
    syncLog(
      `✓ Bridge cancella OK` +
      (resp.log?.length ? ' | ' + resp.log[resp.log.length - 1] : ''),
      'ok'
    );
    await _bridgeReload(false);
  } else {
    throw new Error('Bridge: ' + (resp.error || 'errore cancellazione'));
  }
}
