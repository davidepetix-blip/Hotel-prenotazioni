// ═══════════════════════════════════════════════════════════════════
// sync.js — Build 18.12.01 (Optimized Sync Engine)
// ═══════════════════════════════════════════════════════════════════

const BLIP_VER_SYNC = '15';

/**
 * 1 & 2. Funzione potenziata per il salvataggio: 
 * Legge la griglia, confronta con riga 45/46 e corregge in background.
 */
async function syncGridToDatabase(sheetName) {
  dbg(`[Sync] Analisi foglio: ${sheetName}...`);
  try {
    // Legge l'intero foglio (Griglia + Metadati righe 45 e 46)
    const response = await gapi.client.sheets.spreadsheets.values.get({
      spreadsheetId: DATABASE_SHEET_ID,
      range: `${sheetName}!A1:AJ46`,
    });

    const rows = response.result.values;
    if (!rows) return;

    const roomsInSheet = rows[1].slice(7, 36); // Righe intestazione camere (H2:AJ2)
    const currentJsonRow = rows[44] || []; // Riga 45 (Index 44)
    const currentIdRow = rows[45] || [];   // Riga 46 (Index 45)

    let updateNeeded = false;
    const newJsonRow = [...currentJsonRow];
    const newIdRow = [...currentIdRow];

    // Iteriamo sulle colonne delle camere (dalla H alla AJ)
    for (let colIdx = 0; colIdx < roomsInSheet.length; colIdx++) {
      const realColIdx = colIdx + 7; // Offset per colonna H
      const roomName = roomsInSheet[colIdx];
      
      // Estraiamo la prenotazione visiva dalla colonna
      const extractedData = extractBookingFromColumn(rows, realColIdx);
      const generatedId = generateBookingId(roomName, extractedData);

      // Verifica Coerenza Riga 45 (JSON)
      const stringified = JSON.stringify(extractedData);
      if (currentJsonRow[realColIdx] !== stringified) {
        newJsonRow[realColIdx] = stringified;
        updateNeeded = true;
      }

      // Verifica Coerenza Riga 46 (ID)
      if (currentIdRow[realColIdx] !== generatedId) {
        newIdRow[realColIdx] = generatedId;
        updateNeeded = true;
      }
    }

    if (updateNeeded) {
      dbg(`[Sync] Rilevate discrepanze nelle righe 45/46. Correzione in corso...`);
      await gapi.client.sheets.spreadsheets.values.update({
        spreadsheetId: DATABASE_SHEET_ID,
        range: `${sheetName}!A45:AJ46`,
        valueInputOption: 'RAW',
        resource: { values: [newJsonRow, newIdRow] }
      });
      dbg(`[Sync] Righe 45 e 46 allineate con successo.`);
    }

    // 4. Confronto con Database (Tab PRENOTAZIONI)
    await reconcileWithDatabase(newJsonRow, newIdRow);

  } catch (e) {
    dbg(`[Sync Error] ${e.message}`, true);
  }
}

/**
 * 4. Carica e confronta i dati: aggiunge solo ciò che è nuovo o modificato
 */
async function reconcileWithDatabase(jsonRow, idRow) {
  dbg("[Sync] Sincronizzazione con Database centrale...");
  const dbResponse = await gapi.client.sheets.spreadsheets.values.get({
    spreadsheetId: DATABASE_SHEET_ID,
    range: 'PRENOTAZIONI!A:L',
  });
  
  const dbEntries = dbResponse.result.values || [];
  const dbIds = new Set(dbEntries.map(r => r[0])); // Set di ID esistenti

  const newEntries = [];
  for (let i = 7; i < idRow.length; i++) {
    const id = idRow[i];
    if (id && !dbIds.has(id)) {
      const data = JSON.parse(jsonRow[i]);
      // A=ID, B=Camera, C=Nome, D=Dal, E=Al, F=Disp, G=Note, H=Colore, I=Anno, J=Fonte, K=TS
      newEntries.push([
        id, data.camera, data.nome, data.dal, data.al, 
        data.disposizione, data.note, data.backgroundColor, 
        new Date().getFullYear(), 'sync_auto', new Date().toISOString()
      ]);
    }
  }

  if (newEntries.length > 0) {
    await gapi.client.sheets.spreadsheets.values.append({
      spreadsheetId: DATABASE_SHEET_ID,
      range: 'PRENOTAZIONI!A:A',
      valueInputOption: 'RAW',
      resource: { values: newEntries }
    });
    dbg(`[Sync] Aggiunte ${newEntries.length} nuove prenotazioni al DB.`);
  }
}

function generateBookingId(room, data) {
  if (!data || !data.nome) return "";
  const hash = btoa(room + data.dal + data.nome).substring(0, 8).toUpperCase();
  return `PRE-2026-${hash}`;
}
  accessToken = token;
  onLoginSuccess();
  return true;
}

function initGoogleAuth() {
  handleOAuthRedirect();
}

async function onLoginSuccess() {
  document.getElementById('loginScreen').style.display = 'none';
  try {
    const r = await fetch('https://www.googleapis.com/oauth2/v3/userinfo', {
      headers: { Authorization: 'Bearer ' + accessToken }
    });
    const u = await r.json();
    if (u.picture) {
      const av = document.getElementById('userAvatar');
      av.src = u.picture; av.style.display = 'block';
      av.title = `${u.name} — Clicca per uscire`;
    }
  } catch(e) {}
  if (DATABASE_SHEET_ID || loadDbSheetId()) {
    DATABASE_SHEET_ID = DATABASE_SHEET_ID || loadDbSheetId();
    loadBillSettingsDB().then(s => {
      if (s) localStorage.setItem(BILL_SETTINGS_KEY, JSON.stringify(s));
    }).catch(()=>{});
  }
  await loadFromSheets();
}

// Rigenerazione manuale JSON_ANNUALE via Web App Apps Script
async function rigenera() {
  const cfg = loadBillSettings();
  const url = (cfg.webAppUrl||'').trim();
  if (!url) {
    showToast('URL Web App non configurato — vai in ⚙ Tariffe', 'error');
    return;
  }
  showToast('📡 Chiamata Web App…', 'info');
  try {
    await fetch(`${url}?anno=${new Date().getFullYear()}&ts=${Date.now()}`, {method:'GET',mode:'no-cors'}).catch(()=>{});
    showToast('⏳ Attendi 7 secondi…', 'info');
    await new Promise(r => setTimeout(r, 7000));
    // Verifica che il JSON_ANNUALE sia stato effettivamente aggiornato
    try {
      const sheetEntry = loadAnnualSheets().find(e => e.sheetId);
      if (sheetEntry) {
        const fresh = await readJSONAnnuale(sheetEntry.sheetId);
        if (fresh && fresh.length > 0) {
          annualSheets = loadAnnualSheets();
          await loadFromSheets();
          showToast('✓ Calendario rigenerato', 'success');
          return;
        }
      }
    } catch(e2) {}
    // Se arriviamo qui, la Web App non ha risposto correttamente
    showToast('⚠ Web App non ha risposto — verifica il deploy in Apps Script (Distribuisci → Gestisci distribuzioni → Accesso: Chiunque)', 'warning');
  } catch(e) {
    console.warn('[rigenera] errore:', e.message);
    showToast('❌ Errore Web App: ' + e.message, 'error');
  }
}

function forceSync() {
  loadFromSheets._forceNext = true;
  invalidateDbCache();
  stopBgSync();
  loadFromSheets();
}

function logout() {
  if (!confirm('Vuoi uscire?')) return;
  accessToken = null;
  bookings = [];
  Object.keys(_sheetIdCaches).forEach(k => delete _sheetIdCaches[k]);
  invalidateDbCache();
  stopBgSync();
  document.getElementById('loginScreen').style.display = 'flex';
  document.getElementById('userAvatar').style.display = 'none';
}

// ═══════════════════════════════════════════════════════════════════
// TOKEN SCADUTO — Re-login silenzioso automatico
// ═══════════════════════════════════════════════════════════════════

let _reAuthInProgress = false;

async function apiFetch(url, options = {}) {
  if (!options.headers) options.headers = {};
  options.headers['Authorization'] = 'Bearer ' + accessToken;

  let r = await fetch(url, options);
  if (r.status !== 401) return r;

  const newToken = await trySilentReAuth();
  if (!newToken) {
    showSessionExpiredBanner();
    throw new Error('Sessione scaduta. Fai clic su "Riconnetti" per continuare.');
  }

  options.headers['Authorization'] = 'Bearer ' + newToken;
  return fetch(url, options);
}

function trySilentReAuth() {
  if (_reAuthInProgress) {
    return new Promise(res => {
      const poll = setInterval(() => {
        if (!_reAuthInProgress) { clearInterval(poll); res(accessToken); }
      }, 200);
      setTimeout(() => { clearInterval(poll); res(null); }, 10000);
    });
  }

  _reAuthInProgress = true;
  return new Promise(resolve => {
    const state = randomState();
    const params = new URLSearchParams({
      client_id:     CLIENT_ID,
      redirect_uri:  getRedirectUri(),
      response_type: 'token',
      scope:         SCOPES,
      state,
      prompt:        'none',
      include_granted_scopes: 'true',
    });

    const iframe = document.createElement('iframe');
    iframe.style.cssText = 'display:none;width:1px;height:1px;position:fixed;top:-9999px';
    iframe.src = 'https://accounts.google.com/o/oauth2/v2/auth?' + params.toString();
    document.body.appendChild(iframe);

    const timeout = setTimeout(() => { cleanup(null); }, 8000);

    function cleanup(token) {
      clearTimeout(timeout);
      try { document.body.removeChild(iframe); } catch(e) {}
      _reAuthInProgress = false;
      if (token) { accessToken = token; hideSessionExpiredBanner(); }
      resolve(token);
    }

    iframe.onload = () => {
      try {
        const hash = iframe.contentWindow.location.hash;
        if (hash) {
          const p = new URLSearchParams(hash.slice(1));
          const token = p.get('access_token');
          if (token) { cleanup(token); return; }
        }
      } catch(e) {}
    };
  });
}

function showSessionExpiredBanner() {
  let b = document.getElementById('sessionExpiredBanner');
  if (!b) {
    b = document.createElement('div');
    b.id = 'sessionExpiredBanner';
    b.style.cssText = 'position:fixed;bottom:0;left:0;right:0;z-index:9999;background:#1a1a2e;color:#fff;padding:14px 20px;display:flex;align-items:center;justify-content:space-between;gap:12px;font-size:13px;box-shadow:0 -2px 12px rgba(0,0,0,.4)';
    b.innerHTML = '<span>⏰ <b>Sessione scaduta</b> — Il token Google è scaduto dopo 1 ora.</span>' +
      '<button onclick="startLogin()" style="background:#4285f4;color:#fff;border:none;padding:8px 18px;border-radius:6px;cursor:pointer;font-weight:600;white-space:nowrap">🔑 Riconnetti</button>';
    document.body.appendChild(b);
  }
  b.style.display = 'flex';
}

function hideSessionExpiredBanner() {
  const b = document.getElementById('sessionExpiredBanner');
  if (b) b.style.display = 'none';
}

// ═══════════════════════════════════════════════════════════════════
// GOOGLE SHEETS API — helpers generici
// ═══════════════════════════════════════════════════════════════════

function getDefaultSheetId() {
  return annualSheets.find(e => e.sheetId)?.sheetId || '';
}

async function sheetsGet(range, spreadsheetId) {
  const sid = spreadsheetId || getDefaultSheetId();
  if (!sid) throw new Error('Nessun foglio annuale configurato.');
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${sid}/values/${encodeURIComponent(range)}?valueRenderOption=FORMATTED_VALUE`;
  for (let attempt = 0; attempt < 3; attempt++) {
    const r = await apiFetch(url);
    if (r.status === 429) { await new Promise(res => setTimeout(res, (attempt+1)*2000)); continue; }
    if (!r.ok) throw new Error(`Sheets API error ${r.status}: ${await r.text()}`);
    return r.json();
  }
  throw new Error('RATE_LIMIT: troppo richieste. Riprova tra qualche secondo.');
}

async function sheetsGetWithFormats(range, spreadsheetId) {
  const sid = spreadsheetId || getDefaultSheetId();
  if (!sid) throw new Error('Nessun foglio annuale configurato.');
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${sid}?ranges=${encodeURIComponent(range)}&fields=sheets(data(rowData(values(userEnteredValue,userEnteredFormat/backgroundColor,note))))`;
  for (let attempt = 0; attempt < 3; attempt++) {
    const r = await apiFetch(url);
    if (r.status === 429) { await new Promise(res => setTimeout(res, (attempt+1)*2000)); continue; }
    if (!r.ok) throw new Error(`Sheets API error ${r.status}: ${await r.text()}`);
    return r.json();
  }
  throw new Error('RATE_LIMIT: troppo richieste. Riprova tra qualche secondo.');
}

// ═══════════════════════════════════════════════════════════════════
// CRUD FOGLIO DATABASE
// ═══════════════════════════════════════════════════════════════════

async function dbGet(range) {
  const id = DATABASE_SHEET_ID;
  if (!id) throw new Error('DATABASE_SHEET_ID non configurato. Vai in Impostazioni.');
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${id}/values/${encodeURIComponent(range)}?valueRenderOption=FORMATTED_VALUE`;
  const ctrl = new AbortController();
  const tid  = setTimeout(() => ctrl.abort(), 30000);
  let r;
  try {
    r = await apiFetch(url, { signal: ctrl.signal });
  } catch(e) {
    if (e.name === 'AbortError') throw new Error('DB GET timeout (>30s) — controlla la connessione');
    throw e;
  } finally { clearTimeout(tid); }
  if (!r.ok) throw new Error(`DB GET error ${r.status}: ${await r.text()}`);
  return r.json();
}

async function dbBatchUpdate(requests) {
  const id = DATABASE_SHEET_ID;
  if (!id) throw new Error('DATABASE_SHEET_ID non configurato.');
  const r = await apiFetch(`https://sheets.googleapis.com/v4/spreadsheets/${id}:batchUpdate`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ requests })
  });
  if (!r.ok) throw new Error(`DB batchUpdate error ${r.status}: ${await r.text()}`);
  return r.json();
}

async function dbAppendRow(values) {
  return dbBatchAppendRows([values]);
}

async function dbBatchAppendRows(rowsArray) {
  const id = DATABASE_SHEET_ID;
  if (!id) throw new Error('DATABASE_SHEET_ID non configurato.');
  if (!rowsArray.length) return null;
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${id}/values/${encodeURIComponent(DB_SHEET_NAME)}:append?valueInputOption=RAW&insertDataOption=INSERT_ROWS`;
  const r = await apiFetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ values: rowsArray })
  });
  if (!r.ok) throw new Error(`DB batch append error ${r.status}: ${await r.text()}`);
  return r.json();
}

async function dbUpdateRow(rowNum, values) {
  const id = DATABASE_SHEET_ID;
  if (!id) throw new Error('DATABASE_SHEET_ID non configurato.');
  const lastCol = values.length >= 13 ? 'N' : 'L';
  const range = `${DB_SHEET_NAME}!A${rowNum}:${lastCol}${rowNum}`;
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${id}/values/${encodeURIComponent(range)}?valueInputOption=RAW`;
  const r = await apiFetch(url, {
    method: 'PUT',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ values: [values] })
  });
  if (!r.ok) throw new Error(`DB update error ${r.status}: ${await r.text()}`);
  return r.json();
}

function bookingToDbRow(b, fonte = 'app') {
  const dal = b.s instanceof Date
    ? `${String(b.s.getDate()).padStart(2,'0')}/${String(b.s.getMonth()+1).padStart(2,'0')}/${b.s.getFullYear()}`
    : (b.dal || '');
  const al  = b.e instanceof Date
    ? `${String(b.e.getDate()).padStart(2,'0')}/${String(b.e.getMonth()+1).padStart(2,'0')}/${b.e.getFullYear()}`
    : (b.al || '');
  const anno = b.s instanceof Date ? b.s.getFullYear() : (b.anno || new Date().getFullYear());
  const arr = new Array(12).fill('');
  arr[DB_COLS.ID-1]         = b.dbId || b.id || genBookingId(anno);
  arr[DB_COLS.CAMERA-1]     = b.cameraName || roomName(b.r) || '';
  arr[DB_COLS.NOME-1]       = b.n || '';
  arr[DB_COLS.DAL-1]        = dal;
  arr[DB_COLS.AL-1]         = al;
  arr[DB_COLS.DISP-1]       = b.d || '';
  arr[DB_COLS.NOTE-1]       = b.note || '';
  arr[DB_COLS.COLORE-1]     = b.c || '#D9D9D9';
  arr[DB_COLS.ANNO-1]       = String(anno);
  arr[DB_COLS.FONTE-1]      = fonte;
  arr[DB_COLS.TS-1]         = b.ts || nowISO();
  arr[DB_COLS.DELETED-1]    = b.deleted ? 'true' : '';
  arr[12]                   = b.deletedAt || '';
  arr[13]                   = b.deleteReason || '';
  arr[DB_COLS.CLIENTE_ID-1] = b.clienteId || '';
  return arr;
}

function dbRowToBooking(row, rowNum) {
  const get = (col) => (row[col-1] || '').toString().trim();
  const parseD = (s) => {
    if (!s) return null;
    const p = s.split('/');
    if (p.length === 3) return new Date(+p[2], +p[1]-1, +p[0], 12);
    return null;
  };
  const sd = parseD(get(DB_COLS.DAL)), ed = parseD(get(DB_COLS.AL));
  if (!sd || !ed) return null;
  const camName = get(DB_COLS.CAMERA);
  const room = ROOMS.find(r => r.name === camName || r.name === camName.replace(/\.0$/,''));
  if (!room) return null;
  return {
    id:         rowNum * 10000 + 1,
    dbId:       get(DB_COLS.ID),
    dbRow:      rowNum,
    r:          room.id,
    cameraName: camName,
    n:          get(DB_COLS.NOME),
    d:          get(DB_COLS.DISP),
    c:          get(DB_COLS.COLORE) || '#D9D9D9',
    s:          sd,
    e:          ed,
    note:       get(DB_COLS.NOTE),
    anno:       parseInt(get(DB_COLS.ANNO)) || sd.getFullYear(),
    fonte:      get(DB_COLS.FONTE),
    ts:         get(DB_COLS.TS),
    deleted:    get(DB_COLS.DELETED) === 'true',
    clienteId:  get(DB_COLS.CLIENTE_ID) || null,
    fromSheet:  true,
    fromDb:     true,
    sheetName:  null,
  };
}

// ═══════════════════════════════════════════════════════════════════
// SCHEDA CAMERE (stati operativi + configurazione)
// ═══════════════════════════════════════════════════════════════════

const ROOMS_CACHE_KEY = 'hotelRoomsCache';

async function ensureRoomsSheet() {
  const id = DATABASE_SHEET_ID;
  if (!id) return;
  try {
    const d = await dbGet(`${ROOMS_SHEET_NAME}!A1:G1`);
    if (d.values?.[0]?.[0] === 'CAMERA') return;
  } catch(e) {
    try {
      const addSheet = await fetch(
        `https://sheets.googleapis.com/v4/spreadsheets/${id}:batchUpdate`,
        { method:'POST', headers:{Authorization:'Bearer '+accessToken,'Content-Type':'application/json'},
          body: JSON.stringify({ requests:[{ addSheet:{ properties:{ title:ROOMS_SHEET_NAME } } }] }) }
      );
      if (!addSheet.ok) {
        const err = await addSheet.text();
        if (!err.includes('already exists') && !err.includes('ALREADY_EXISTS')) throw new Error(err);
      }
    } catch(e2) { if (!String(e2.message).includes('already')) throw e2; }
  }
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${id}/values/${encodeURIComponent(ROOMS_SHEET_NAME+'!A1:G1')}?valueInputOption=RAW`;
  await fetch(url, {
    method:'PUT', headers:{Authorization:'Bearer '+accessToken,'Content-Type':'application/json'},
    body: JSON.stringify({ values:[['CAMERA','MAX_OSPITI','LETTI_AMMESSI','PULIZIA','CONFIGURAZIONE','NOTE_OPS','TS_AGGIORNAMENTO']] })
  });
}

async function readRoomsSheet() {
  const data = await dbGet(`${ROOMS_SHEET_NAME}!A2:G9999`);
  const rows = data.values || [];
  const result = {};
  rows.forEach((row, i) => {
    const cam = (row[RCOLS.CAMERA-1] || '').trim();
    if (!cam) return;
    const room = ROOMS.find(r => r.name === cam);
    if (!room) return;
    result[room.id] = {
      dbRow:          i + 2,
      maxGuests:      parseInt(row[RCOLS.MAX_OSPITI-1]) || 2,
      allowedBeds:    (row[RCOLS.LETTI_AMMESSI-1] || 'm,s').split(',').map(s=>s.trim()).filter(Boolean),
      pulizia:        row[RCOLS.PULIZIA-1]        || 'da-pulire',
      configurazione: row[RCOLS.CONFIGURAZIONE-1] || '',
      noteOps:        row[RCOLS.NOTE_OPS-1]       || '',
      ts:             row[RCOLS.TS-1]             || '',
    };
  });
  return result;
}

async function writeRoomsSheet(statesMap) {
  const id = DATABASE_SHEET_ID;
  if (!id) return;
  const rows = ROOMS.map(room => {
    const s = statesMap[room.id] || {};
    return [
      room.name,
      s.maxGuests     || ROOM_DEFAULTS[room.id]?.maxGuests || 2,
      (s.allowedBeds  || ROOM_DEFAULTS[room.id]?.allowedBeds || ['m','s']).join(','),
      s.pulizia       || 'da-pulire',
      s.configurazione|| '',
      s.noteOps       || '',
      nowISO(),
    ];
  });
  const range = `${ROOMS_SHEET_NAME}!A2:G${1+rows.length}`;
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${id}/values/${encodeURIComponent(range)}?valueInputOption=RAW`;
  await fetch(url, {
    method:'PUT', headers:{Authorization:'Bearer '+accessToken,'Content-Type':'application/json'},
    body: JSON.stringify({ values: rows })
  });
  try { localStorage.setItem(ROOMS_CACHE_KEY, JSON.stringify({ ts:Date.now(), data:statesMap })); } catch(e){}
}

async function updateSingleRoomState(roomId, patch) {
  const id = DATABASE_SHEET_ID;
  const room = ROOMS.find(r => r.id === roomId);
  if (!room || !id) return;
  await ensureRoomsSheet();
  roomStates[roomId] = { ...(roomStates[roomId]||{}), ...patch, ts: nowISO() };
  const s = roomStates[roomId];
  const values = [[
    room.name,
    s.maxGuests     || ROOM_DEFAULTS[roomId]?.maxGuests || 2,
    (s.allowedBeds  || ROOM_DEFAULTS[roomId]?.allowedBeds || ['m','s']).join(','),
    s.pulizia       || 'da-pulire',
    s.configurazione|| '',
    s.noteOps       || '',
    s.ts,
  ]];
  if (s.dbRow) {
    const range = `${ROOMS_SHEET_NAME}!A${s.dbRow}:G${s.dbRow}`;
    const url = `https://sheets.googleapis.com/v4/spreadsheets/${id}/values/${encodeURIComponent(range)}?valueInputOption=RAW`;
    await fetch(url, { method:'PUT', headers:{Authorization:'Bearer '+accessToken,'Content-Type':'application/json'}, body:JSON.stringify({values}) });
  } else {
    const url = `https://sheets.googleapis.com/v4/spreadsheets/${id}/values/${encodeURIComponent(ROOMS_SHEET_NAME)}:append?valueInputOption=RAW&insertDataOption=INSERT_ROWS`;
    const resp = await fetch(url, { method:'POST', headers:{Authorization:'Bearer '+accessToken,'Content-Type':'application/json'}, body:JSON.stringify({values}) });
    const r2 = await resp.json();
    const m = (r2.updates?.updatedRange||'').match(/(\d+):/);
    if (m) roomStates[roomId].dbRow = parseInt(m[1]);
  }
  try { localStorage.setItem(ROOMS_CACHE_KEY, JSON.stringify({ ts:Date.now(), data:roomStates })); } catch(e){}
  syncRoomSettingsFromStates();
}

function loadRoomsCache() {
  try {
    const raw = localStorage.getItem(ROOMS_CACHE_KEY);
    if (!raw) return null;
    const p = JSON.parse(raw);
    if (Date.now() - p.ts > 60*60*1000) return null;
    return p.data;
  } catch(e) { return null; }
}

async function loadRoomStates() {
  if (!DATABASE_SHEET_ID) return;
  const cached = loadRoomsCache();
  if (cached) { roomStates = cached; syncRoomSettingsFromStates(); return; }
  try {
    await ensureRoomsSheet();
    const fromDb = await readRoomsSheet();
    if (Object.keys(fromDb).length > 0) {
      roomStates = fromDb;
    } else {
      ROOMS.forEach(r => {
        roomStates[r.id] = {
          maxGuests:      ROOM_DEFAULTS[r.id]?.maxGuests || 2,
          allowedBeds:    ROOM_DEFAULTS[r.id]?.allowedBeds || ['m','s'],
          pulizia:        'da-pulire',
          configurazione: '',
          noteOps:        '',
          ts:             '',
        };
      });
      await writeRoomsSheet(roomStates);
    }
    syncRoomSettingsFromStates();
    try { localStorage.setItem(ROOMS_CACHE_KEY, JSON.stringify({ ts:Date.now(), data:roomStates })); } catch(e){}
  } catch(e) { console.warn('loadRoomStates error:', e.message); }
}

function syncRoomSettingsFromStates() {
  ROOMS.forEach(r => {
    const s = roomStates[r.id];
    if (!s) return;
    roomSettings[r.id] = {
      maxGuests:   s.maxGuests   || ROOM_DEFAULTS[r.id]?.maxGuests   || 2,
      allowedBeds: s.allowedBeds || ROOM_DEFAULTS[r.id]?.allowedBeds || ['m','s'],
    };
  });
}

// ═══════════════════════════════════════════════════════════════════
// DATABASE CRUD
// ═══════════════════════════════════════════════════════════════════

async function readDatabase() {
  const data = await dbGet(`${DB_SHEET_NAME}!A${DB_FIRST_ROW}:O3000`); // A-O: 15 col (incl. CLIENTE_ID)
  const rows = data.values || [];
  dbRowCache = [];
  const result = [];
  rows.forEach((row, i) => {
    const rowNum = DB_FIRST_ROW + i;
    const b = dbRowToBooking(row, rowNum);
    dbRowCache.push({ rowNum, raw: row });
    if (b && !b.deleted) result.push(b);
  });
  return result;
}

async function ensureDbHeaders() {
  try {
    const d = await dbGet(`${DB_SHEET_NAME}!A1:L1`);
    const row = d.values?.[0] || [];
    if (row[0] === 'ID') return;
  } catch(e) {}
  const id = DATABASE_SHEET_ID;
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${id}/values/${encodeURIComponent(DB_SHEET_NAME+'!A1:L1')}?valueInputOption=RAW`;
  await fetch(url, {
    method: 'PUT',
    headers: { Authorization: 'Bearer ' + accessToken, 'Content-Type': 'application/json' },
    body: JSON.stringify({ values: [['ID','CAMERA','NOME','DAL','AL','DISPOSIZIONE','NOTE','COLORE','ANNO','FONTE','TS_MODIFICA','DELETED']] })
  });
}

async function dbUpsert(b, fonte = 'app') {
  b.ts = nowISO();
  const row = bookingToDbRow(b, fonte);
  if (b.dbRow) {
    await dbUpdateRow(b.dbRow, row);
  } else {
    const resp = await dbAppendRow(row);
    const updatedRange = resp.updates?.updatedRange || '';
    const m = updatedRange.match(/(\d+):/);
    if (m) b.dbRow = parseInt(m[1]);
    if (!b.dbId) b.dbId = row[DB_COLS.ID-1];
  }
}

async function dbDelete(b, reason = 'cancellata dal foglio') {
  // Se manca dbRow ma abbiamo dbId, recupera la riga dal DB prima di archiviare
  let target = b;
  if (!b.dbRow && b.dbId && DATABASE_SHEET_ID) {
    try {
      const all = await readDatabase();
      const found = all.find(r => r.dbId === b.dbId);
      if (found) target = found;
    } catch(e) { console.warn('[dbDelete] lookup dbId:', e.message); }
  }
  if (!target.dbRow) {
    // Nessuna riga DB trovata — segna comunque deleted=true via update diretto per dbId
    console.warn('[dbDelete] dbRow mancante, impossibile archiviare:', b.dbId || b.id);
    return;
  }
  try { await archiviaInCestino([target], reason); } catch(e) { console.warn('[dbDelete]:', e.message); }
}

async function archiviaInCestino(lista, reason) {
  const id = DATABASE_SHEET_ID;
  if (!id || lista.length === 0) return;
  const ts = nowISO();
  await ensureCestinoHeaders();

  const righe = lista.map(b => {
    const row = bookingToDbRow(b, b.fonte || 'app');
    row[DB_COLS.DELETED-1] = 'true';
    row[12] = ts; row[13] = reason || 'motivo non specificato'; row[14] = b.dbRow || '';
    return row;
  });
  const urlAppend = `https://sheets.googleapis.com/v4/spreadsheets/${id}/values/${encodeURIComponent(CESTINO_SHEET_NAME)}:append?valueInputOption=RAW&insertDataOption=INSERT_ROWS`;
  await apiFetch(urlAppend, { method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({values:righe}) });

  // Elimina fisicamente le righe dal foglio PRENOTAZIONI
  // (invece di aggiornarle con DELETED=true, che le fa accumulare infinitamente)
  const conRiga = lista.filter(b => b.dbRow).sort((a,b) => b.dbRow - a.dbRow); // ordine decrescente!
  if (conRiga.length > 0) {
    // Ottieni lo sheetId numerico del foglio PRENOTAZIONI
    let sheetNumId = null;
    try {
      const meta = await apiFetch(`https://sheets.googleapis.com/v4/spreadsheets/${id}?fields=sheets.properties`);
      const metaJson = await meta.json();
      const sheet = (metaJson.sheets||[]).find(s => s.properties.title === DB_SHEET_NAME);
      sheetNumId = sheet?.properties?.sheetId ?? null;
    } catch(e) { console.warn('[CESTINO] sheetId lookup:', e.message); }

    if (sheetNumId !== null) {
      // Raggruppa le righe adiacenti in range per minimizzare le chiamate API
      // Ordine decrescente per evitare che la cancellazione sposti gli indici
      const deleteRequests = conRiga.map(b => ({
        deleteDimension: {
          range: { sheetId: sheetNumId, dimension: 'ROWS',
                   startIndex: b.dbRow - 1, endIndex: b.dbRow }
        }
      }));
      for (let i = 0; i < deleteRequests.length; i += 500) {
        const chunk = deleteRequests.slice(i, i+500);
        await apiFetch(`https://sheets.googleapis.com/v4/spreadsheets/${id}:batchUpdate`, {
          method:'POST', headers:{'Content-Type':'application/json'},
          body: JSON.stringify({ requests: chunk })
        });
      }
    } else {
      // Fallback: aggiorna con DELETED=true se non riusciamo a ottenere lo sheetId
      const data = conRiga.map(b => {
        const row = bookingToDbRow(b, b.fonte || 'app');
        row[DB_COLS.DELETED-1] = 'true'; row[DB_COLS.TS-1] = ts; row[12] = ts; row[13] = reason;
        const lastCol = String.fromCharCode(64 + row.length);
        return { range: `${DB_SHEET_NAME}!A${b.dbRow}:${lastCol}${b.dbRow}`, values: [row] };
      });
      for (let i = 0; i < data.length; i += 1000) {
        await apiFetch(`https://sheets.googleapis.com/v4/spreadsheets/${id}/values:batchUpdate`, {
          method:'POST', headers:{'Content-Type':'application/json'},
          body: JSON.stringify({ valueInputOption:'RAW', data: data.slice(i, i+1000) })
        });
      }
    }
  }
  console.log(`[CESTINO] ${lista.length} righe archiviate e rimosse dal DB`);
}

let _cestinoHeadersChecked = false;
async function ensureCestinoHeaders() {
  if (_cestinoHeadersChecked) return;
  _cestinoHeadersChecked = true;
  const id = DATABASE_SHEET_ID;
  if (!id) return;
  let exists = false;
  try {
    const meta = await apiFetch(`https://sheets.googleapis.com/v4/spreadsheets/${id}?fields=sheets.properties.title`);
    const metaJson = await meta.json();
    exists = (metaJson.sheets||[]).some(s => s.properties.title === CESTINO_SHEET_NAME);
  } catch(e) {}
  if (!exists) {
    try {
      await apiFetch(`https://sheets.googleapis.com/v4/spreadsheets/${id}:batchUpdate`, {
        method:'POST', headers:{'Content-Type':'application/json'},
        body: JSON.stringify({ requests:[{ addSheet:{ properties:{ title:CESTINO_SHEET_NAME } } }] })
      });
    } catch(e) { return; }
  }
  try {
    const d = await dbGet(`${CESTINO_SHEET_NAME}!A1:O1`);
    if (d.values?.[0]?.length > 0) return;
  } catch(e) {}
  try {
    const url = `https://sheets.googleapis.com/v4/spreadsheets/${id}/values/${encodeURIComponent(CESTINO_SHEET_NAME+'!A1:O1')}?valueInputOption=RAW`;
    await apiFetch(url, { method:'PUT', headers:{'Content-Type':'application/json'},
      body: JSON.stringify({ values:[['ID','CAMERA','NOME','DAL','AL','DISPOSIZIONE','NOTE','COLORE','ANNO','FONTE','TS_MODIFICA','DELETED','DELETED_AT','REASON','RIGA_ORIGINALE']] })
    });
  } catch(e) {}
}

// ═══════════════════════════════════════════════════════════════════
// LETTURA FOGLI ANNUALI
// ═══════════════════════════════════════════════════════════════════

async function readAnnualSheet(entry) {
  const { sheetId } = entry;
  const result = [];
  const metaR = await fetch(
    `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}?fields=sheets.properties.title`,
    { headers: { Authorization: 'Bearer ' + accessToken } }
  );
  if (!metaR.ok) return result;
  const meta = await metaR.json();
  const sheetNames = meta.sheets
    .map(s => s.properties.title)
    .filter(n => !EXCLUDED_SHEETS.includes(n) && /^[A-Za-z\u00C0-\u00D6\u00D8-\u00F6\u00F8-\u00FF]+\s+\d{4}$/i.test(n));

  for (const sName of sheetNames) {
    try {
      const base = `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values/`;
      const enc  = s => encodeURIComponent(s);
      const [hR, jR, idR] = await Promise.all([
        fetch(base+enc(`'${sName}'!B${HEADER_ROW}:AJ${HEADER_ROW}`)+'?valueRenderOption=FORMATTED_VALUE', {headers:{Authorization:'Bearer '+accessToken}}),
        fetch(base+enc(`'${sName}'!B${OUTPUT_ROW}:AJ${OUTPUT_ROW}`)+'?valueRenderOption=FORMATTED_VALUE', {headers:{Authorization:'Bearer '+accessToken}}),
        fetch(base+enc(`'${sName}'!B${BLIP_ID_ROW}:AJ${BLIP_ID_ROW}`)+'?valueRenderOption=FORMATTED_VALUE', {headers:{Authorization:'Bearer '+accessToken}}),
      ]);
      const headers = (await hR.json()).values?.[0] || [];
      const jsonRow = (await jR.json()).values?.[0] || [];
      const idRow   = (await idR.json()).values?.[0] || []; // BLIP IDs riga 46

      sheetColumnMap[sName] = {};
      headers.forEach((h, i) => {
        if (!h) return;
        const raw = String(h).trim(), norm = raw.replace(/\.0$/, '');
        sheetColumnMap[sName][raw] = i + 2;
        if (norm !== raw) sheetColumnMap[sName][norm] = i + 2;
      });

      jsonRow.forEach(cell => {
        if (!cell || !cell.trim()) return;
        try {
          let s = cell.trim();
          if (s.startsWith("'")) s = s.slice(1);
          const bList = Array.isArray(JSON.parse(s)) ? JSON.parse(s) : [JSON.parse(s)];
          bList.forEach(b => {
            if (!b.dal || !b.al || !b.camera) return;
            const [dd,mm,yy] = b.dal.split('/').map(Number);
            const [de,me,ye] = b.al.split('/').map(Number);
            const camNorm = String(b.camera).trim().replace(/\.0$/,'');
            const room = ROOMS.find(r => r.name===camNorm || r.name.toLowerCase()===camNorm.toLowerCase());
            if (!room) { console.warn('Camera non riconosciuta:', b.camera); return; }
            const colorHex = (b.backgroundColor||'#D9D9D9').startsWith('#') ? (b.backgroundColor||'#D9D9D9') : '#'+b.backgroundColor;
            // Legge il BLIP_ID dalla riga 46 per la colonna corrispondente
            const colIdx   = sheetColumnMap[sName]?.[String(b.camera).trim()] || null;
            const blipColI = colIdx ? colIdx - 2 : null; // riga 46 inizia da col B (indice 0)
            const blipId   = (blipColI !== null && idRow[blipColI]) ? String(idRow[blipColI]).trim() : null;
            result.push({
              id:nid++, r:room.id, n:b.nome||'—', d:b.disposizione||'',
              c:colorHex, s:new Date(yy,mm-1,dd,12), e:new Date(ye,me-1,de,12),
              note:b.note||'', fromSheet:true, sheetName:sName,
              cameraName:room.name, sheetId,
              dbId:blipId||null, dbRow:null, ts:null, fonte:'manuale',
              _sheetCol:colIdx, // colonna nel foglio, usata per scrivere riga 46
            });
          });
        } catch(e) {}
      });
    } catch(e) { console.warn(`Errore "${sName}":`, e.message); }
  }
  return result;
}

async function readJSONAnnuale(sheetId) {
  const TAB   = 'JSON_ANNUALE';
  const range = encodeURIComponent("'" + TAB + "'!A2:A13");
  const url   = "https://sheets.googleapis.com/v4/spreadsheets/" + sheetId + "/values/" + range + "?valueRenderOption=FORMATTED_VALUE";

  for (let attempt = 0; attempt < 3; attempt++) {
    const r = await fetch(url, { headers: { Authorization: 'Bearer ' + accessToken } });
    if (r.status === 429) { await new Promise(res => setTimeout(res, (attempt+1)*2000)); continue; }
    if (!r.ok) throw new Error(TAB + ' non disponibile (HTTP ' + r.status + ')');
    const data = await r.json();
    const rows = data.values;
    if (!rows || rows.length === 0) throw new Error(TAB + ' vuoto — esegui "Rigenera JSON_ANNUALE" dal menu del foglio');

    const allPren = [];
    let parseErrors = 0;
    for (const row of rows) {
      const raw = row?.[0];
      if (!raw || raw.trim()==='' || raw.startsWith('—')) continue;
      try {
        const chunk = JSON.parse(raw);
        if (Array.isArray(chunk)) chunk.forEach(p => allPren.push(p));
      } catch(e) { parseErrors++; console.warn('[JSON_ANNUALE] Riga non parsificabile:', raw.substring(0,80)); }
    }
    if (allPren.length === 0) {
      if (parseErrors > 0) throw new Error(TAB + ': JSON non valido in ' + parseErrors + ' righe');
      throw new Error(TAB + ': nessuna prenotazione — rigenera dal menu del foglio');
    }
    return _parseJSONAnnualeBookings(allPren, sheetId, TAB);
  }
  throw new Error(TAB + ': rate limit persistente, riprova');
}

function _parseJSONAnnualeBookings(parsed, sheetId, tabName) {
  const result = [];
  let skipped = 0;
  for (const b of parsed) {
    if (!b.dal || !b.al || !b.camera) { skipped++; continue; }
    const [dd,mm,yy] = b.dal.split('/').map(Number);
    const [de,me,ye] = b.al.split('/').map(Number);
    if (!dd||!mm||!yy||!de||!me||!ye) { skipped++; continue; }
    const camNorm = String(b.camera).trim().replace(/\.0$/,'');
    const room = ROOMS.find(r => r.name===camNorm || r.name.toLowerCase()===camNorm.toLowerCase());
    if (!room) { console.warn(`[${tabName}] Camera non riconosciuta: "${b.camera}"`); skipped++; continue; }
    const colorHex = String(b.backgroundColor||'#D9D9D9').trim();
    const color = colorHex.startsWith('#') ? colorHex : '#'+colorHex;
    const sName = sheetName(yy, mm-1);
    if (!sheetColumnMap[sName]) sheetColumnMap[sName] = {};
    result.push({
      id: nid++, r: room.id, n: b.nome||'—', d: b.disposizione||'',
      c: color, s: new Date(yy,mm-1,dd,12), e: new Date(ye,me-1,de,12),
      note: b.note||'', fromSheet:true, fromJSONAnnuale:true,
      sheetName:sName, sheetId, cameraName:room.name,
      dbId:null, dbRow:null, ts:null, fonte:'manuale',
    });
  }
  if (skipped > 0) console.warn(`[${tabName}] ${skipped} prenotazioni saltate`);
  console.log(`[${tabName}] ✓ ${result.length} prenotazioni caricate`);
  return result;
}

// ═══════════════════════════════════════════════════════════════════
// CACHE DB LOCALE
// ═══════════════════════════════════════════════════════════════════

const DB_CACHE_KEY    = 'hotelDbCache';
const DB_CACHE_TTL_MS = 8 * 60 * 60 * 1000; // 8 ore — ricalcola solo se la cache è davvero vecchia
const BATCH_SIZE      = 50;
const DAY_MS          = 86400000;

function saveDbCache(data) {
  try {
    localStorage.setItem(DB_CACHE_KEY, JSON.stringify({
      ts: Date.now(),
      data: data.map(b => ({ ...b, s: b.s?.toISOString(), e: b.e?.toISOString() }))
    }));
  } catch(e) { console.warn('Cache non salvata:', e.message); }
}
function loadDbCache() {
  try {
    const p = JSON.parse(localStorage.getItem(DB_CACHE_KEY) || 'null');
    if (!p || Date.now() - p.ts > DB_CACHE_TTL_MS) return null;
    return p.data.map(b => ({ ...b, s: new Date(b.s), e: new Date(b.e) }));
  } catch(e) { return null; }
}
function invalidateDbCache() { localStorage.removeItem(DB_CACHE_KEY); }

// ═══════════════════════════════════════════════════════════════════
// SYNC ENGINE
// ═══════════════════════════════════════════════════════════════════

function findMatch(target, list) {
  // PRIORITÀ 1: match per BLIP_ID dalla riga 46 — match perfetto, immune a spostamenti
  if (target.dbId) {
    const byId = list.find(b => b.dbId === target.dbId);
    if (byId) return byId;
  }

  const camT  = (target.cameraName || roomName(target.r) || '').toLowerCase().trim();
  const nomT  = (target.n || '').toLowerCase().trim();
  const dayT  = Math.round((target.s?.getTime?.() || 0) / DAY_MS);
  const dispT = (target.d || '').trim().toLowerCase();

  // PRIORITÀ 2: match esatto per nome + camera + data
  let m = list.find(b => {
    if ((b.n||'').toLowerCase().trim() !== nomT) return false;
    if ((b.cameraName||roomName(b.r)||'').toLowerCase().trim() !== camT) return false;
    return Math.round((b.s?.getTime?.() || 0) / DAY_MS) === dayT;
  });
  if (m) return m;

  // PRIORITÀ 3: fuzzy — nome + camera + data ±1 giorno
  m = list.find(b => {
    if ((b.n||'').toLowerCase().trim() !== nomT) return false;
    if ((b.cameraName||roomName(b.r)||'').toLowerCase().trim() !== camT) return false;
    if (dispT && (b.d||'').trim().toLowerCase() !== dispT) return false;
    return Math.abs(Math.round((b.s?.getTime?.() || 0) / DAY_MS) - dayT) <= 1;
  });
  return m || null;
}

async function cleanupDeletedFromDb(dbRows) {
  // Con il nuovo archiviaInCestino che elimina fisicamente le righe dal DB,
  // non dovrebbero più esistere righe DELETED nel foglio PRENOTAZIONI.
  // Questa funzione è mantenuta come safety net ma non fa più niente di default.
  const alreadyDeleted = dbRows.filter(b => b.deleted && b.dbRow);
  if (alreadyDeleted.length === 0) return 0;
  // Se troviamo righe DELETED residue (es. da prima del fix), le eliminiamo
  syncLog(`⚠ Trovate ${alreadyDeleted.length} righe DELETED residue — pulizia…`, 'wrn');
  try { await archiviaInCestino(alreadyDeleted, 'Pulizia residui DELETED'); } catch(e) {}
  return alreadyDeleted.length;
}

async function syncWithDatabase(sheetBookings, forceFullSync = false) {
  if (!DATABASE_SHEET_ID) return sheetBookings;

  showLoading('Lettura database…');
  const allDbRows = await readDatabase();
  const cleanedN  = await cleanupDeletedFromDb(allDbRows);
  if (cleanedN > 0) showLoading(`Cestino: ${cleanedN} righe spostate…`);

  const dbActive      = allDbRows.filter(b => !b.deleted);
  const result        = [];
  const toAddToDB     = [];
  const toUpdateInDB  = [];
  const toArchive     = [];
  const seenDbIds     = new Set();

  // GUARD: se il DB è quasi vuoto rispetto al foglio (es. dopo pulizia manuale),
  // blocca il re-import automatico — è quasi sempre distruttivo e lento.
  // Il forceSync (🔄) bypassa questo guard e reimporta tutto intenzionalmente.
  const dbWasEmpty = dbActive.length < sheetBookings.length * 0.1 && sheetBookings.length > 50;
  if (dbWasEmpty && !forceFullSync) {
    syncLog(`⚠ DB ha solo ${dbActive.length} righe vs ${sheetBookings.length} nel foglio — import bloccato. Premi 🔄 per reimportare tutto.`, 'wrn');
    showToast(`⚠ DB quasi vuoto (${dbActive.length} righe vs ${sheetBookings.length} nel foglio). Premi 🔄 per reimportare.`, 'warning');
    return dbActive;
  }
  if (dbWasEmpty && forceFullSync) {
    syncLog(`🔄 Force sync: reimport completo (${sheetBookings.length} prenotazioni dal foglio)`, 'syn');
  }

  // FASE 1: Foglio → verità assoluta
  for (const sheet of sheetBookings) {
    if (!sheet.n || sheet.n === '???' || sheet.n.trim() === '') continue;
    // Salta prenotazioni cancellate localmente (blacklist anti-ghost)
    if (isDeletedLocally(sheet.dbId)) { syncLog(`🗑 Skip blacklist: ${sheet.n}`, 'wrn'); continue; }
    const match = findMatch(sheet, dbActive);
    if (!match) {
      sheet.dbId  = genBookingId(sheet.s.getFullYear());
      sheet.ts    = nowISO();
      sheet.fonte = 'manuale';
      toAddToDB.push(sheet);
      result.push(sheet);
    } else {
      seenDbIds.add(match.dbId);
      const changed =
        match.d !== sheet.d || match.c !== sheet.c || match.note !== sheet.note ||
        Math.abs(match.e.getTime() - sheet.e.getTime()) > DAY_MS/2 ||
        Math.abs(match.s.getTime() - sheet.s.getTime()) > DAY_MS/2;
      if (changed) {
        match.d = sheet.d; match.c = sheet.c; match.note = sheet.note;
        match.s = sheet.s; match.e = sheet.e; match.ts = nowISO(); match.fonte = 'manuale';
        toUpdateInDB.push(match);
      }
      result.push(match);
    }
  }

  // FASE 2: DB → gestione non-matchati
  for (const db of dbActive) {
    if (seenDbIds.has(db.dbId)) continue;
    // Salta prenotazioni cancellate localmente (blacklist anti-ghost)
    if (isDeletedLocally(db.dbId)) { syncLog(`🗑 Skip blacklist DB: ${db.n}`, 'wrn'); continue; }
    if (!db.n || db.n === '???' || db.n.trim() === '') { result.push(db); continue; }

    const tsModifica = db.ts ? new Date(db.ts).getTime() : 0;
    const etaMs      = Date.now() - tsModifica;
    const isRecenteApp = db.fonte === 'app' && etaMs < 15 * 60 * 1000;
    if (isRecenteApp) {
      syncLog(`🛡 Protetta ${db.n} cam.${db.cameraName||db.r} (inserita ${Math.round(etaMs/1000)}s fa)`, 'wrn');
      result.push(db); continue;
    }

    // GUARD re-import massiccio: non cestinare se il foglio porta >50% nuove
    if (toAddToDB.length > sheetBookings.length * 0.5 && sheetBookings.length > 20) {
      result.push(db); continue;
    }

    // GUARD mesi futuri: il JSON_ANNUALE potrebbe non coprire mesi oltre l'anno corrente
    // o mesi per cui la Web App non ha ancora rigenerato il JSON.
    // Non cestinare prenotazioni future se il foglio non ha dati per quel mese.
    const dbMonth = db.s ? new Date(db.s).getMonth() : -1;
    const dbYear  = db.s ? new Date(db.s).getFullYear() : 0;
    const sheetHasMonth = sheetBookings.some(s =>
      s.s && new Date(s.s).getFullYear() === dbYear && new Date(s.s).getMonth() === dbMonth
    );
    if (!sheetHasMonth && db.s && db.s > new Date()) {
      // Il foglio non ha nessuna prenotazione in quel mese futuro — probabilmente
      // il JSON_ANNUALE non copre ancora quel mese. Teniamo la prenotazione.
      result.push(db); continue;
    }

    db.deleted      = true;
    db.deleteReason = 'Rimossa dal foglio Gantt · sync del ' + new Date().toLocaleDateString('it-IT');
    db.deletedAt    = nowISO();
    toArchive.push(db);
  }

  // FASE 3: Scrittura DB
  if (toAddToDB.length > 0) {
    const total = toAddToDB.length;
    for (let i = 0; i < total; i += BATCH_SIZE) {
      const chunk = toAddToDB.slice(i, i+BATCH_SIZE);
      showLoading(`Importazione ${Math.min(i+BATCH_SIZE,total)}/${total}…`);
      try {
        const resp = await dbBatchAppendRows(chunk.map(b => bookingToDbRow(b, 'manuale')));
        const m = (resp?.updates?.updatedRange||'').match(/(\d+):/);
        if (m) { const startRow = parseInt(m[1])-chunk.length+1; chunk.forEach((b,idx) => { b.dbRow = startRow+idx; }); }
      } catch(e) { console.warn('[DB] Errore import batch:', e.message); }
    }
    syncLog(`✚ Importate ${total} nuove prenotazioni nel DB`, 'db');
  }
  if (toUpdateInDB.length > 0) {
    showLoading(`Aggiornamento ${toUpdateInDB.length} prenotazioni…`);
    for (const b of toUpdateInDB) {
      try { await dbUpsert(b, b.fonte); } catch(e) { console.warn('[DB] Update error:', e.message); }
    }
  }
  if (toArchive.length > 0) {
    showLoading(`Cestino: ${toArchive.length} prenotazioni…`);
    try { await archiviaInCestino(toArchive, 'Rimossa dal foglio Gantt · ' + new Date().toLocaleDateString('it-IT')); } catch(e) {}
    const rimossi = toArchive.filter(b=>b.dbRow).length;
    syncLog(`🗑 ${rimossi} → CESTINO`, 'wrn');
  }

  const dbActiveCount = dbActive.length;
  syncLog(`DB: ${result.length} prenotazioni attive, rimossi ${toArchive.length}`, 'db');

  // Scrivi BLIP_ID nella riga 46 per le prenotazioni nuove (fire & forget)
  const toWriteRow46 = toAddToDB.filter(b => b.dbId && b._sheetCol && b.sheetName && b.sheetId);
  if (toWriteRow46.length > 0) {
    writeBlipIdsToRow46(toWriteRow46).catch(e =>
      syncLog('⚠ Scrittura riga 46: ' + e.message, 'wrn')
    );
  }

  return result;
}

// Scrive i BLIP_ID nella riga 46 del foglio visivo
async function writeBlipIdsToRow46(bookings) {
  const bySheet = {};
  for (const b of bookings) {
    const key = b.sheetId + '|' + b.sheetName;
    if (!bySheet[key]) bySheet[key] = { sheetId:b.sheetId, sName:b.sheetName, updates:[] };
    bySheet[key].updates.push({ col:b._sheetCol, id:b.dbId });
  }
  for (const { sheetId, sName, updates } of Object.values(bySheet)) {
    if (!updates.length) continue;
    try {
      const url = 'https://sheets.googleapis.com/v4/spreadsheets/' + sheetId + '/values:batchUpdate';
      const resp = await apiFetch(url, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          valueInputOption: 'RAW',
          data: updates.map(({ col, id }) => ({
            range: "'" + sName + "'!" + columnLetter(col) + BLIP_ID_ROW,
            values: [[id]]
          }))
        })
      });
      if (resp.ok) syncLog('✓ BLIP_ID scritti riga 46 in ' + sName + ': ' + updates.length, 'ok');
      else syncLog('⚠ Errore scrittura riga 46 in ' + sName, 'wrn');
    } catch(e) { syncLog('⚠ writeBlipIdsToRow46: ' + e.message, 'wrn'); }
  }
}

// ── Backfill riga 46: scrive BLIP_ID per tutte le prenotazioni già in memoria ──
// Da chiamare una volta sola per popolare la riga 46 su tutte le colonne
async function backfillRow46() {
  if (!DATABASE_SHEET_ID || !bookings.length) {
    showToast('Carica prima le prenotazioni', 'error'); return;
  }
  showLoading('Lettura intestazioni fogli…');

  // Raccogli sheetId univoci per le prenotazioni dal foglio
  const sheetIds = [...new Set(bookings.filter(b => b.fromSheet && b.sheetId).map(b => b.sheetId))];
  if (!sheetIds.length) { hideLoading(); showToast('Nessun foglio fonte trovato', 'error'); return; }

  // Per ogni sheetId, leggi i nomi dei fogli mensili e le loro intestazioni
  for (const sid of sheetIds) {
    try {
      const metaR = await apiFetch(
        'https://sheets.googleapis.com/v4/spreadsheets/' + sid + '?fields=sheets.properties.title'
      );
      const meta = await metaR.json();
      const monthSheets = (meta.sheets || [])
        .map(s => s.properties.title)
        .filter(n => !EXCLUDED_SHEETS.includes(n) && /^[A-Za-z\u00C0-\u00FF]+\s+\d{4}$/i.test(n));

      for (const sName of monthSheets) {
        if (sheetColumnMap[sName] && Object.keys(sheetColumnMap[sName]).length > 0) continue; // già noto
        try {
          const enc = s => encodeURIComponent(s);
          const hR = await apiFetch(
            'https://sheets.googleapis.com/v4/spreadsheets/' + sid + '/values/' +
            enc("'" + sName + "'!B" + HEADER_ROW + ":AJ" + HEADER_ROW) +
            '?valueRenderOption=FORMATTED_VALUE'
          );
          const headers = (await hR.json()).values?.[0] || [];
          if (!sheetColumnMap[sName]) sheetColumnMap[sName] = {};
          headers.forEach((h, i) => {
            if (!h) return;
            const raw = String(h).trim(), norm = raw.replace(/\.0$/, '');
            sheetColumnMap[sName][raw] = i + 2;
            if (norm !== raw) sheetColumnMap[sName][norm] = i + 2;
          });
          syncLog('✓ Intestazioni lette: ' + sName + ' (' + headers.filter(Boolean).length + ' camere)', 'ok');
        } catch(e) { syncLog('⚠ Intestazioni ' + sName + ': ' + e.message, 'wrn'); }
      }
    } catch(e) { syncLog('⚠ Meta foglio: ' + e.message, 'wrn'); }
  }

  // Ora assegna _sheetCol a ogni prenotazione
  let rebuilt = 0;
  for (const b of bookings.filter(b2 => b2.fromSheet && b2.dbId && !b2._sheetCol)) {
    const sName = b.sheetName;
    const cam   = b.cameraName || roomName(b.r) || '';
    const colMap = sheetColumnMap[sName] || {};
    const col = colMap[cam] || colMap[cam.trim()] || colMap[String(cam).replace(/\.0$/, '')];
    if (col && b.sheetId) { b._sheetCol = col; rebuilt++; }
  }
  syncLog('Colonne ricostruite: ' + rebuilt, 'ok');

  const toWrite = bookings.filter(b => b.fromSheet && b.dbId && b._sheetCol && b.sheetName && b.sheetId);
  if (toWrite.length === 0) {
    hideLoading();
    showToast('Nessuna prenotazione abbinabile alle colonne del foglio', 'error');
    return;
  }
  showLoading('Scrittura riga 46 (' + toWrite.length + ' prenotazioni)…');
  await writeBlipIdsToRow46(toWrite);
  hideLoading();
  showToast('✓ Riga 46 compilata: ' + toWrite.length + ' celle scritte', 'success');
  syncLog('✓ Backfill riga 46 completato: ' + toWrite.length, 'ok');
}

// Converte numero colonna (1-based) → lettera Excel (A, B, ..., Z, AA, ...)
function columnLetter(col) {
  let s = '';
  while (col > 0) { col--; s = String.fromCharCode(65 + col % 26) + s; col = Math.floor(col / 26); }
  return s;
}

// ═══════════════════════════════════════════════════════════════════
// DIAGNOSTICA — Confronto riga 45 (Apps Script) vs DB
// ═══════════════════════════════════════════════════════════════════

// Confronta le prenotazioni del foglio con quelle nel DB
// e logga le discrepanze → aiuta a capire se Apps Script funziona
function diagnosticaAppScript(sheetBookings, dbBookings) {
  const now = new Date().toISOString().slice(0,10);
  let ok = 0, discrepanze = 0;
  for (const sb of sheetBookings) {
    const match = dbBookings.find(db => db.dbId === sb.dbId || (
      (db.n||'').toLowerCase() === (sb.n||'').toLowerCase() &&
      (db.cameraName||'') === (sb.cameraName||'') &&
      Math.abs((db.s?.getTime()||0) - (sb.s?.getTime()||0)) < 86400000
    ));
    if (!match) { discrepanze++; syncLog('⚠ AppsScript: "' + sb.n + '" cam.' + sb.cameraName + ' ' + (sb.s?.toISOString()?.slice(0,10)||'?') + ' — non in DB', 'wrn'); }
    else ok++;
  }
  syncLog('📊 Diagnostica: ' + ok + ' OK, ' + discrepanze + ' discrepanze foglio↔DB', discrepanze > 0 ? 'wrn' : 'ok');
  return discrepanze;
}

// ═══════════════════════════════════════════════════════════════════
// SYNC LOG
// ═══════════════════════════════════════════════════════════════════

let _logEntries = [];
const MAX_LOG   = 200;

function syncLog(msg, type='inf') {
  const now  = new Date();
  const time = now.toTimeString().slice(0,8);
  _logEntries.unshift({ time, msg, type });
  if (_logEntries.length > MAX_LOG) _logEntries.pop();
  const el = document.getElementById('syncLogEntries');
  if (el && document.getElementById('syncLog')?.classList.contains('open')) _renderLog();
  const btn = document.getElementById('syncLogBtn');
  if (btn) btn.textContent = `📋 LOG (${_logEntries.length})`;
  if (type === 'err') showBgToast('❌ ' + msg.slice(0,60), 4000);
}

function _renderLog() {
  const el = document.getElementById('syncLogEntries');
  if (!el) return;
  el.innerHTML = _logEntries.map(e =>
    `<div class="sle ${e.type}"><span class="slt">${e.time}</span><span class="slm">${e.msg}</span></div>`
  ).join('');
}

function toggleSyncLog() {
  const log = document.getElementById('syncLog');
  if (!log) return;
  const wasOpen = log.classList.contains('open');
  log.classList.toggle('open');
  if (!wasOpen) _renderLog();
}

function clearSyncLog(e) {
  e && e.stopPropagation();
  _logEntries = [];
  const el = document.getElementById('syncLogEntries');
  if (el) el.innerHTML = '';
  const btn = document.getElementById('syncLogBtn');
  if (btn) btn.textContent = '📋 LOG (0)';
}

// ═══════════════════════════════════════════════════════════════════
// BGSYNC — sincronizzazione periodica in background
// ═══════════════════════════════════════════════════════════════════

let _bgSyncTimer   = null;
let _bgSyncRunning = false;
// Blacklist locale: dbId cancellati localmente, ignorati dal bgSync
// per evitare il "ghost reappearance" prima che il DB si aggiorni
const _deletedLocally = new Map(); // dbId → timestamp cancellazione
const _DELETED_TTL = 60 * 60 * 1000; // 1 ora

function markDeletedLocally(dbId) {
  if (dbId) _deletedLocally.set(String(dbId), Date.now());
}
function isDeletedLocally(dbId) {
  if (!dbId) return false;
  const t = _deletedLocally.get(String(dbId));
  if (!t) return false;
  if (Date.now() - t > _DELETED_TTL) { _deletedLocally.delete(String(dbId)); return false; }
  return true;
}

function startBgSync() {
  if (_bgSyncTimer) clearInterval(_bgSyncTimer);
  _bgSyncTimer = setInterval(bgSync, 2 * 60 * 1000);
}
function stopBgSync() {
  if (_bgSyncTimer) { clearInterval(_bgSyncTimer); _bgSyncTimer = null; }
}
function showBgToast(msg, ms = 2500) {
  const el = document.getElementById('bgSyncToast');
  if (!el) return;
  el.textContent = msg; el.classList.add('show');
  setTimeout(() => el.classList.remove('show'), ms);
}
function setSyncPulsing(on) {
  const dot = document.getElementById('syncDot');
  if (!dot) return;
  dot.classList.toggle('pulsing', on);
  if (!on) dot.classList.remove('show');
}

async function bgSync() {
  if (!accessToken || _bgSyncRunning) return;
  _bgSyncRunning = true;
  setSyncPulsing(true);
  syncLog('⟳ bgSync avviato', 'syn');
  try {
    if (!DATABASE_SHEET_ID) DATABASE_SHEET_ID = loadDbSheetId();
    const dbFresh = await readDatabase();
    const active  = dbFresh.filter(b => !b.deleted && !isDeletedLocally(b.dbId));

    const prevCount = bookings.length;
    let changed = active.length !== prevCount;
    if (!changed) {
      const localMap = new Map(bookings.map(b => [b.dbId, b.ts]).filter(([k]) => k));
      for (const db of active) {
        const localTs = localMap.get(db.dbId);
        if (!localTs || (db.ts && db.ts > localTs)) { changed = true; break; }
      }
    }
    const rimossi = bookings.filter(b => b.dbId && !active.find(d => d.dbId === b.dbId)).length;
    syncLog(`DB: ${active.length} prenotazioni attive, rimossi ${rimossi}`, 'db');

    // GUARD: se il DB ha molto meno prenotazioni di quelle in memoria, salta il render
    if (active.length < prevCount * 0.7 && prevCount > 20) {
      syncLog(`⚠ bgSync: DB ha ${active.length} vs ${prevCount} in memoria — skip render`, 'wrn');
      await loadRoomStates();
      return;
    }

    if (changed || rimossi > 0) {
      bookings = mergeMultiMonthBookings(active);
      saveDbCache(active);
      render();
      const diff = active.length - prevCount;
      showBgToast(
        rimossi > 0 ? `↻ ${rimossi} rimoss${rimossi===1?'a':'e'}` :
        diff    > 0 ? `↻ +${diff} nuove` : '↻ Aggiornato'
      );
    }
    await loadRoomStates();
    if (document.getElementById('roomDashPage')?.classList.contains('open')) renderRoomDash();
  } catch(e) {
    console.warn('[bgSync] Errore:', e.message);
  } finally {
    setSyncPulsing(false);
    _bgSyncRunning = false;
  }
}

// ═══════════════════════════════════════════════════════════════════
// CARICAMENTO INIZIALE
// ═══════════════════════════════════════════════════════════════════

async function loadFromSheets() {
  if (!accessToken) return;
  document.getElementById('syncDot')?.classList.remove('show', 'pulsing');

  if (!DATABASE_SHEET_ID) DATABASE_SHEET_ID = loadDbSheetId();
  annualSheets = loadAnnualSheets();
  const forcing = loadFromSheets._forceNext === true;
  loadFromSheets._forceNext = false;
  stopBgSync();

  try {
    bookings = [];
    sheetColumnMap = {};

    // FAST PATH: cache locale valida
    if (DATABASE_SHEET_ID && !forcing) {
      const cached = loadDbCache();
      if (cached && cached.length > 0) {
        bookings = mergeMultiMonthBookings(cached);
        await loadRoomStates();
        hideLoading();
        render();
        showToast(`✓ ${bookings.length} prenotazioni (cache)`, 'success');
        setTimeout(bgSync, 2000);
        startBgSync();
        return;
      }
    }

    // FULL PATH
    const sheetEntry = annualSheets.find(e => e.sheetId);
    let sheetBookings = [];

    if (sheetEntry) {
      syncLog('📖 Lettura JSON_ANNUALE da foglio…', 'syn');
      showLoading('Lettura JSON_ANNUALE…');
      try {
        sheetBookings = await readJSONAnnuale(sheetEntry.sheetId);
        showLoading(`Foglio: ${sheetBookings.length} prenotazioni`);
      } catch(err) {
        console.warn('[JSON_ANNUALE] Fallback fogli mensili:', err.message);
        for (let i = 0; i < annualSheets.length; i++) {
          const entry = annualSheets[i];
          if (!entry.sheetId) continue;
          showLoading(`Lettura ${entry.label} (${i+1}/${annualSheets.length})…`);
          try { sheetBookings.push(...(await readAnnualSheet(entry))); } catch(e2) {}
          if (i < annualSheets.length-1) await new Promise(r => setTimeout(r, 300));
        }
      }
    }

    if (DATABASE_SHEET_ID) {
      syncLog('Sincronizzazione DB…', 'syn');
      showLoading('Sincronizzazione database…');
      await ensureDbHeaders();
      sheetBookings = await syncWithDatabase(sheetBookings, true);
    }

    await loadRoomStates();
    bookings = mergeMultiMonthBookings(sheetBookings);
    hideLoading();
    render();

    const fromJSON = sheetBookings.some(b => b.fromJSONAnnuale);
    const badge = document.getElementById('jsonSourceBadge');
    if (badge) {
      badge.textContent = fromJSON ? 'JSON' : '12f';
      badge.title = fromJSON ? 'Dati da JSON_ANNUALE (1 chiamata API)' : 'Dati dai 12 fogli mensili (fallback)';
      badge.style.display = 'inline';
      badge.style.color   = fromJSON ? '#2d6a4f' : '#e67e22';
    }
    showToast(`✓ ${bookings.length} prenotazioni`, 'success');
    syncLog(`✓ ${bookings.length} prenotazioni caricate`, 'ok');
    saveDbCache(DATABASE_SHEET_ID ? sheetBookings : bookings);
    startBgSync();
  } catch(e) {
    hideLoading();
    showToast('Errore caricamento: ' + e.message, 'error');
    syncLog('❌ ' + e.message, 'err');
    dbg('Errore loadFromSheets: ' + e.message, true);
  }
}

// ═══════════════════════════════════════════════════════════════════
// SCRITTURA / CANCELLAZIONE SUL FOGLIO GOOGLE
// ═══════════════════════════════════════════════════════════════════

function splitBookingByMonth(b) {
  const fragments = [];
  let cur = new Date(b.s);
  const end = new Date(b.e);
  while (cur < end) {
    const y = cur.getFullYear(), m = cur.getMonth();
    const actualEnd = end <= new Date(y, m+1, 0, 12) ? end : new Date(y, m+1, 0, 12);
    fragments.push({
      sName:      sheetName(y, m),
      startDay:   cur.getDate(),
      endDay:     actualEnd.getDate(),
      isLastFrag: actualEnd >= end,
      y, m,
    });
    cur = new Date(y, m+1, 1, 12);
  }
  return fragments;
}

let _sheetIdCache = null;
const _sheetIdCaches = {};

async function getSheetIdMap(spreadsheetId) {
  if (!spreadsheetId) {
    const first = annualSheets.find(e => e.sheetId);
    spreadsheetId = first?.sheetId || '';
  }
  if (!spreadsheetId) throw new Error('Nessun foglio annuale configurato in Impostazioni.');
  if (_sheetIdCaches[spreadsheetId]) return _sheetIdCaches[spreadsheetId];
  const r = await fetch(
    `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}?fields=sheets(properties(title,sheetId))`,
    { headers: { Authorization: 'Bearer ' + accessToken } }
  );
  const d = await r.json();
  _sheetIdCaches[spreadsheetId] = {};
  d.sheets.forEach(s => { _sheetIdCaches[spreadsheetId][s.properties.title] = s.properties.sheetId; });
  return _sheetIdCaches[spreadsheetId];
}

async function writeFragment(sName, cameraName, startDay, endDay, firstCellText, note, color, sheetIdMap, sid) {
  let colIdx = sheetColumnMap[sName]?.[cameraName];
  if (!colIdx) {
    try {
      const hd = await sheetsGet(`'${sName}'!B${HEADER_ROW}:AJ${HEADER_ROW}`);
      sheetColumnMap[sName] = {};
      (hd.values?.[0]||[]).forEach((h,i) => {
        if(!h) return;
        const raw=String(h).trim(), norm=raw.replace(/\.0$/,'');
        sheetColumnMap[sName][raw]=i+2;
        if(norm!==raw) sheetColumnMap[sName][norm]=i+2;
      });
      colIdx = sheetColumnMap[sName]?.[cameraName];
    } catch(e) {}
  }
  if (!colIdx) throw new Error(`Camera "${cameraName}" non trovata nel foglio "${sName}"`);
  const sheetId = sheetIdMap[sName];
  if (sheetId === undefined) throw new Error(`Foglio "${sName}" non trovato`);

  const startRow = FIRST_DATA_ROW + startDay - 1;
  const numRows  = endDay - startDay + 1;
  if (numRows <= 0) return;

  const sheetsColor = hexToSheetsColor(color);
  const requests = [
    { updateCells: { range:{sheetId,startRowIndex:startRow-1,endRowIndex:startRow,startColumnIndex:colIdx-1,endColumnIndex:colIdx}, rows:[{values:[{userEnteredValue:{stringValue:firstCellText}}]}], fields:'userEnteredValue' } },
    { repeatCell:  { range:{sheetId,startRowIndex:startRow-1,endRowIndex:startRow-1+numRows,startColumnIndex:colIdx-1,endColumnIndex:colIdx}, cell:{userEnteredFormat:{backgroundColor:sheetsColor}}, fields:'userEnteredFormat.backgroundColor' } }
  ];
  if (note) {
    requests.push({ updateCells: { range:{sheetId,startRowIndex:startRow-1,endRowIndex:startRow,startColumnIndex:colIdx-1,endColumnIndex:colIdx}, rows:[{values:[{note}]}], fields:'note' } });
  }

  const spreadsheetId = sid || annualSheets.find(e=>e.sheetId)?.sheetId || '';
  const resp = await apiFetch(`https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}:batchUpdate`, {
    method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({requests})
  });
  if (!resp.ok) throw new Error(`Scrittura fallita (${resp.status}): ${await resp.text()}`);
}

async function writeBookingToSheet(b) {
  const firstCellText = `${b.n} ${b.d}`.trim();
  const fragments     = splitBookingByMonth(b);

  for (const frag of fragments) {
    const colorEnd = frag.isLastFrag ? frag.endDay - 1 : frag.endDay;
    if (colorEnd < frag.startDay) continue;
    const annEntry = annualSheets.find(e => e.year === frag.y);
    if (!annEntry?.sheetId) { console.warn('Nessun sheetId per anno ' + frag.y); continue; }
    const sheetIdMap = await getSheetIdMap(annEntry.sheetId);
    await writeFragment(frag.sName, b.cameraName, frag.startDay, colorEnd, firstCellText, b.note, b.c, sheetIdMap, annEntry.sheetId);
    await triggerAppsScriptUpdate(frag.sName, b.cameraName, annEntry.sheetId);
  }

  syncLog(`✏ Prenotazione scritta: ${b.n} cam.${b.cameraName} ${b.s?.toLocaleDateString('it-IT')||''}→${b.e?.toLocaleDateString('it-IT')||''} (${fragments.length} mesi)`, 'ok');
  segnalaModificaAdAppsScript(annualSheets[0]?.sheetId).catch(e => console.warn('[segnaModifica]:', e.message));

  if (DATABASE_SHEET_ID) {
    b.dbId = b.dbId || genBookingId(b.s.getFullYear());
    b.ts   = nowISO(); b.fonte = 'app'; b.fromSheet = true;
    await dbUpsert(b, 'app');
  }
}

async function triggerAppsScriptUpdate(sName, cameraName, spreadsheetId) {
  // no-op: la rigenerazione JSON_ANNUALE è delegata alla Web App
  console.log(`[triggerAppsScriptUpdate] skip — gestito da WebApp`);
}

async function segnalaModificaAdAppsScript(sheetId) {
  const cfg = loadBillSettings();
  const webAppUrl = (cfg.webAppUrl || '').trim();
  if (!webAppUrl) {
    if (!window._webAppWarnShown) {
      window._webAppWarnShown = true;
      showToast('⚠ Configura URL Web App in ⚙ Tariffe per aggiornamento immediato del calendario', 'warning');
    }
    return;
  }
  try {
    const anno = new Date().getFullYear();
    const url  = `${webAppUrl}?anno=${anno}&ts=${Date.now()}`;
    syncLog(`📡 Chiamata Web App: ${url.slice(0,60)}…`, 'syn');

    // no-cors: non possiamo leggere la risposta, ma la richiesta parte
    // Usiamo un Image trick come fallback diagnostico
    await fetch(url, { method:'GET', mode:'no-cors' }).catch(() => {});

    // Aspetta che Apps Script elabori (3-8 sec tipicamente)
    await new Promise(r => setTimeout(r, 7000));

    // Rileggi il JSON_ANNUALE per verificare se è stato aggiornato
    try {
      const sheetEntry = annualSheets.find(e => e.sheetId);
      if (!sheetEntry) return;
      const freshBookings = await readJSONAnnuale(sheetEntry.sheetId);
      if (freshBookings && freshBookings.length > 0) {
        syncLog(`✓ Web App OK — JSON_ANNUALE aggiornato (${freshBookings.length} prenotazioni)`, 'ok');
        // Aggiorna bookings solo se non c'è già un sync in corso
        if (!_bgSyncRunning) {
          bookings = mergeMultiMonthBookings(freshBookings.filter(b => !isDeletedLocally(b.dbId)));
          render();
        }
      } else {
        syncLog('⚠ Web App: JSON_ANNUALE vuoto dopo aggiornamento — verifica deploy', 'wrn');
        showToast('⚠ Web App non risponde correttamente — verifica il deploy in Apps Script', 'warning');
      }
    } catch(e2) {
      syncLog(`⚠ Web App: verifica fallita (${e2.message}) — premi 🔄 per aggiornare manualmente`, 'wrn');
    }
  } catch(e) {
    console.warn('[WebApp] chiamata fallita:', e.message);
    syncLog(`❌ Web App errore: ${e.message}`, 'err');
  }
}

async function clearFragment(sName, cameraName, startDay, endDay, sheetIdMap, spreadsheetId) {
  let colIdx = sheetColumnMap[sName]?.[cameraName];
  if (!colIdx) {
    // Carica la mappa colonne on-demand (può essere vuota dopo forceSync/riavvio)
    try {
      const hd = await sheetsGet(`'${sName}'!B${HEADER_ROW}:AJ${HEADER_ROW}`);
      sheetColumnMap[sName] = {};
      (hd.values?.[0]||[]).forEach((h,i) => {
        if (!h) return;
        const raw = String(h).trim(), norm = raw.replace(/\.0$/,'');
        sheetColumnMap[sName][raw] = i + 2;
        if (norm !== raw) sheetColumnMap[sName][norm] = i + 2;
      });
      colIdx = sheetColumnMap[sName]?.[cameraName];
    } catch(e) { console.warn('[clearFragment] caricamento mappa colonne:', e.message); }
  }
  if (!colIdx) { console.warn(`[clearFragment] Camera "${cameraName}" non trovata in "${sName}"`); return; }
  const sheetId = sheetIdMap[sName];
  if (sheetId === undefined) return;
  const startRow = FIRST_DATA_ROW + startDay - 1;
  const numRows  = endDay - startDay + 1;
  if (numRows <= 0) return;
  const requests = [
    { repeatCell:  { range:{sheetId,startRowIndex:startRow-1,endRowIndex:startRow-1+numRows,startColumnIndex:colIdx-1,endColumnIndex:colIdx}, cell:{userEnteredFormat:{backgroundColor:{red:1,green:1,blue:1}}}, fields:'userEnteredFormat.backgroundColor' } },
    { updateCells: { range:{sheetId,startRowIndex:startRow-1,endRowIndex:startRow,startColumnIndex:colIdx-1,endColumnIndex:colIdx}, rows:[{values:[{userEnteredValue:{stringValue:''}}]}], fields:'userEnteredValue' } }
  ];
  const cSid = spreadsheetId || annualSheets.find(e=>e.sheetId)?.sheetId || '';
  const cr = await apiFetch(`https://sheets.googleapis.com/v4/spreadsheets/${cSid}:batchUpdate`, {
    method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({requests})
  });
  if (!cr.ok) console.warn(`clearFragment HTTP ${cr.status}`);
}

async function clearBookingFromSheet(b) {
  const fragments = splitBookingByMonth(b);
  for (const frag of fragments) {
    const colorEnd = frag.isLastFrag ? frag.endDay - 1 : frag.endDay;
    if (colorEnd < frag.startDay) continue;
    const annEntry = annualSheets.find(e => e.year === frag.y);
    const sid = annEntry?.sheetId || b.sheetId || annualSheets.find(e=>e.sheetId)?.sheetId || '';
    if (!sid) continue;
    try {
      const sheetIdMap = await getSheetIdMap(sid);
      await clearFragment(frag.sName, b.cameraName, frag.startDay, colorEnd, sheetIdMap, sid);
      await triggerAppsScriptUpdate(frag.sName, b.cameraName, sid);
    } catch(e) { console.warn(`Errore pulizia frammento ${frag.sName}:`, e.message); }
  }
  segnalaModificaAdAppsScript(annualSheets[0]?.sheetId)
    .catch(e => console.warn('[clearBooking→WebApp]:', e.message));
}

// ═══════════════════════════════════════════════════════════════════
// MERGE MULTI-MESE
// ═══════════════════════════════════════════════════════════════════

function mergeMultiMonthBookings(list) {
  const sorted = [...list].sort((a,b) => {
    if (a.r !== b.r) return a.r.localeCompare(b.r);
    return a.s - b.s;
  });
  const merged = [], used = new Set();
  for (let i = 0; i < sorted.length; i++) {
    if (used.has(i)) continue;
    let base = sorted[i];
    for (let j = i+1; j < sorted.length; j++) {
      if (used.has(j)) continue;
      const next = sorted[j];
      if (base.r !== next.r) break;
      if (base.n !== next.n || base.c !== next.c || base.d !== next.d) continue;
      const baseLastDay = new Date(base.e.getFullYear(), base.e.getMonth()+1, 0).getDate();
      const baseEndsOnLast = base.e.getDate() === baseLastDay;
      const nextStartsOnFirst = next.s.getDate() === 1;
      const nextIsFollowing =
        (next.s.getFullYear()===base.e.getFullYear() && next.s.getMonth()===base.e.getMonth()+1) ||
        (base.e.getMonth()===11 && next.s.getMonth()===0 && next.s.getFullYear()===base.e.getFullYear()+1);
      if (baseEndsOnLast && nextStartsOnFirst && nextIsFollowing) {
        base = { ...base, e: next.e };
        used.add(j);
      }
    }
    merged.push(base);
  }
  return merged;
}
