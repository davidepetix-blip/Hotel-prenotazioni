// ═══════════════════════════════════════════════════════════════════
// store.js — DB CRUD, Sync Engine, bgSync, CESTINO, readJSONAnnuale
// Blip Hotel Management — build 18.10.4
//
// Responsabilità:
//   • CRUD su foglio DATABASE (PRENOTAZIONI, CONTI, CAMERE, CESTINO)
//   • syncWithDatabase: merge foglio grafico ↔ DB
//   • bgSync: polling periodico con fingerprint change detection
//   • readJSONAnnuale: lettura ottimizzata dati foglio grafico
//   • loadFromSheets: entry point principale post-login
//   • splitBookingByMonth, mergeMultiMonthBookings: utilità multi-mese
//   • Log sessione e diagnostica
//
// Dipende da: core.js, api.js, auth.js
// Caricato PRIMA di: clienti.js, gantt.js, checkin.js, billing.js, bridge.js
// ═══════════════════════════════════════════════════════════════════

const BLIP_VER_STORE = '1'; // ← incrementa ad ogni modifica (era BLIP_VER_SYNC)


// ─────────────────────────────────────────────────────────────────
// CESTINO BLACKLIST — set in-memory degli ID cestinati
//
// IMPORTANTE: mette in blacklist SOLO le cancellazioni esplicite utente.
// Le cancellazioni automatiche da sync ("Rimossa dal foglio Gantt")
// NON entrano in blacklist: se la prenotazione torna nel foglio grafico
// deve poter essere reimportata normalmente.
// Questo evita che booking finiti nel CESTINO per errore durante una
// sync anomala (es. il caso dei 103 booking) blocchino per sempre
// il reimport di prenotazioni valide.
// ─────────────────────────────────────────────────────────────────
let _cestinoBlacklist = null;       // null = non ancora caricata
let _cestinoBlacklistTs = 0;        // timestamp ultimo caricamento
const CESTINO_BLACKLIST_TTL = 10 * 60 * 1000; // 10 minuti

// Reason che indicano cancellazione AUTOMATICA da sync — NON vanno in blacklist
const CESTINO_AUTO_REASONS = [
  'Rimossa dal foglio Gantt',
  'Rimossa dal foglio',
  'sync del',
  'Pulizia residui DELETED',
];
function _isCestinoAutoSync(reason) {
  if (!reason) return false;
  const r = String(reason).toLowerCase();
  return CESTINO_AUTO_REASONS.some(k => r.includes(k.toLowerCase()));
}

// Set dei BLIP_ID trovati nella riga 46 del foglio grafico durante readAnnualSheet.
// Usato da syncWithDatabase per non cestinare prenotazioni che sono presenti
// nel foglio (riga 46 aggiornata) ma mancano dalla riga 45 (Apps Script non girato).
// Reset ad ogni loadFromSheets completo.
let _row46BlipIds = new Set();
// Mappa per assegnare dbId alle prenotazioni lette da JSON_ANNUALE (che non hanno dbId).
// Formato: "sheetName|cameraName|dal" → dbId (BLIP_ID)
// Costruita durante la batchGet in readJSONAnnuale leggendo la riga 46.
let _row46BookingMap = {};

// ─────────────────────────────────────────────────────────────────
// FINGERPRINT — riga 47 del foglio grafico
//
// Ogni colonna del foglio mensile ha una formula in riga 47:
//   =SUMPRODUCT(LEN(B3:B44)*ROW(B3:B44))+COUNTA(B3:B44)*100000
//
// Questa formula si ricalcola automaticamente ogni volta che il
// contenuto della colonna cambia — indipendentemente da chi ha
// fatto la modifica (utente, Apps Script, o Blip via API).
//
// bgSync legge solo la riga 47 (1 chiamata leggera), confronta
// con i valori memorizzati, e rilegge il JSON_ANNUALE SOLO se
// almeno un foglio mensile è cambiato.
// ─────────────────────────────────────────────────────────────────
// _sheetFingerprints[sheetName] = "val1|val2|...|valN" (stringify della riga 47)
let _sheetFingerprints = {};

// ─────────────────────────────────────────────────────────────────
// CACHE METADATA FOGLI — evita chiamate duplicate ?fields=sheets.properties.title
//
// Ogni chiamata a spreadsheets/{id}?fields=sheets.properties.title costa 1 quota.
// readJSONAnnuale la eseguiva 2 volte (una per il JSON, una per le intestazioni).
// Questa cache (TTL 30 min) la riduce a 1 per sessione di sync.
// ─────────────────────────────────────────────────────────────────
const _metaCacheTitles = {}; // spreadsheetId → { ts, titles: string[] }
const _META_CACHE_TTL  = 30 * 60 * 1000; // 30 minuti

async function _getMonthSheetTitles(spreadsheetId) {
  const cached = _metaCacheTitles[spreadsheetId];
  if (cached && Date.now() - cached.ts < _META_CACHE_TTL) return cached.titles;
  const r = await apiFetch(
    'https://sheets.googleapis.com/v4/spreadsheets/' + spreadsheetId + '?fields=sheets.properties.title'
  );
  if (!r.ok) return [];
  const j = await r.json();
  const titles = (j.sheets || [])
    .map(s => s.properties.title)
    .filter(n => !EXCLUDED_SHEETS.includes(n) && /^[A-Za-zÀ-ÿ]+\s+\d{4}$/i.test(n));
  _metaCacheTitles[spreadsheetId] = { ts: Date.now(), titles };
  return titles;
}

async function loadCestinoBlacklist(force = false) {
  const age = Date.now() - _cestinoBlacklistTs;
  if (!force && _cestinoBlacklist && age < CESTINO_BLACKLIST_TTL) return _cestinoBlacklist;
  if (!DATABASE_SHEET_ID) { _cestinoBlacklist = new Set(); return _cestinoBlacklist; }
  try {
    // Legge solo ID(A) e REASON(N) con batchGet per minimizzare il trasferimento.
    // Con 4872 righe leggere A:N era il principale collo di bottiglia (~3s).
    // Ora leggiamo solo le due colonne necessarie in una batchGet.
    const enc = s => encodeURIComponent(s);
    const batchUrl = 'https://sheets.googleapis.com/v4/spreadsheets/' + DATABASE_SHEET_ID
      + '/values:batchGet?ranges=' + enc(`${CESTINO_SHEET_NAME}!A2:A9999`)
      + '&ranges=' + enc(`${CESTINO_SHEET_NAME}!N2:N9999`)
      + '&valueRenderOption=FORMATTED_VALUE';
    const br = await apiFetch(batchUrl);
    const bj = br.ok ? await br.json() : { valueRanges: [] };
    const idsCol    = bj.valueRanges?.[0]?.values || [];
    const reasonCol = bj.valueRanges?.[1]?.values || [];
    const blacklisted = [];
    const autoSync    = [];
    idsCol.forEach((row, i) => {
      const id     = (row[0]  || '').trim();
      const reason = ((reasonCol[i] || [])[0] || '').trim();
      if (!id) return;
      if (_isCestinoAutoSync(reason)) {
        autoSync.push(id);
      } else {
        blacklisted.push(id);
      }
    });
    _cestinoBlacklist = new Set(blacklisted);
    _cestinoBlacklistTs = Date.now();
    syncLog(`🗑 Blacklist CESTINO: ${blacklisted.length} ID (utente) + ${autoSync.length} da sync (esclusi)`, 'syn');
  } catch(e) {
    _cestinoBlacklist = new Set();
    _cestinoBlacklistTs = Date.now();
  }
  return _cestinoBlacklist;
}

function isInCestinoBlacklist(dbId) {
  if (!dbId || !_cestinoBlacklist) return false;
  return _cestinoBlacklist.has(String(dbId));
}

// Aggiorna la blacklist quando una prenotazione viene cestinata
function addToCestinoBlacklist(dbId) {
  if (!dbId) return;
  if (!_cestinoBlacklist) _cestinoBlacklist = new Set();
  _cestinoBlacklist.add(String(dbId));
}


// ── apiFetch, trySilentReAuth, token bucket → api.js
// ── randomState, handleOAuthRedirect, onLoginSuccess, logout → auth.js

// ═══════════════════════════════════════════════════════════════════
// GOOGLE SHEETS API — helpers generici
// ═══════════════════════════════════════════════════════════════════

// ── Helper: restituisce l'entry annualSheets dell'anno corrente ──
// annualSheets può contenere più anni (es. 2025 e 2026).
// JSON_ANNUALE esiste solo sul foglio dell'anno in corso — usare sempre
// questo helper invece di .find(e => e.sheetId) che restituisce il primo
// entry disponibile (potrebbe essere un anno passato senza JSON_ANNUALE).
function _currentYearSheetEntry() {
  const y = new Date().getFullYear();
  return annualSheets.find(e => e.sheetId && e.year === y)
      || annualSheets.find(e => e.sheetId)
      || null;
}

function getDefaultSheetId() {
  return _currentYearSheetEntry()?.sheetId || '';
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

  // PRENOTAZIONI ha 13 colonne (A-M): ID,CAMERA,NOME,DAL,AL,DISP,NOTE,COLORE,ANNO,FONTE,TS,DELETED,CLIENTE_ID
  // deletedAt e deleteReason sono usati SOLO dal foglio CESTINO (colonne N-O) e vengono
  // aggiunti separatamente da archiviaInCestino — NON fanno parte di questo array base.
  const arr = new Array(13).fill('');
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
    n:          get(DB_COLS.NOME).replace(/\s*\+\s*$/, '').trim(),
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
      const addSheet = await apiFetch(
        `https://sheets.googleapis.com/v4/spreadsheets/${id}:batchUpdate`,
        { method:'POST', headers:{'Content-Type':'application/json'},
          body: JSON.stringify({ requests:[{ addSheet:{ properties:{ title:ROOMS_SHEET_NAME } } }] }) }
      );
      if (!addSheet.ok) {
        const err = await addSheet.text();
        if (!err.includes('already exists') && !err.includes('ALREADY_EXISTS')) throw new Error(err);
      }
    } catch(e2) { if (!String(e2.message).includes('already')) throw e2; }
  }
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${id}/values/${encodeURIComponent(ROOMS_SHEET_NAME+'!A1:G1')}?valueInputOption=RAW`;
  await apiFetch(url, {
    method:'PUT', headers:{'Content-Type':'application/json'},
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
  await apiFetch(url, {
    method:'PUT', headers:{'Content-Type':'application/json'},
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
    await apiFetch(url, { method:'PUT', headers:{'Content-Type':'application/json'}, body:JSON.stringify({values}) });
  } else {
    const url = `https://sheets.googleapis.com/v4/spreadsheets/${id}/values/${encodeURIComponent(ROOMS_SHEET_NAME)}:append?valueInputOption=RAW&insertDataOption=INSERT_ROWS`;
    const resp = await apiFetch(url, { method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({values}) });
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

let _dbHeadersChecked = false;
async function ensureDbHeaders() {
  if (_dbHeadersChecked) return; // già verificato in questa sessione
  try {
    const d = await dbGet(`${DB_SHEET_NAME}!A1:L1`);
    const row = d.values?.[0] || [];
    if (row[0] === 'ID') { _dbHeadersChecked = true; return; }
  } catch(e) {}
  const id = DATABASE_SHEET_ID;
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${id}/values/${encodeURIComponent(DB_SHEET_NAME+'!A1:L1')}?valueInputOption=RAW`;
  await apiFetch(url, {
    method: 'PUT',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ values: [['ID','CAMERA','NOME','DAL','AL','DISPOSIZIONE','NOTE','COLORE','ANNO','FONTE','TS_MODIFICA','DELETED']] })
  });
  _dbHeadersChecked = true;
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

// Cache sheetId numerico del foglio PRENOTAZIONI (per batchUpdate delete)
let _dbSheetIdCache = null;
async function _getDbSheetId(spreadsheetId) {
  if (_dbSheetIdCache) return _dbSheetIdCache;
  try {
    const r = await apiFetch(`https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}?fields=sheets.properties`);
    const j = await r.json();
    const sh = (j.sheets||[]).find(s => s.properties.title === DB_SHEET_NAME);
    _dbSheetIdCache = sh?.properties?.sheetId ?? 0;
  } catch(e) { _dbSheetIdCache = 0; }
  return _dbSheetIdCache;
}

async function archiviaInCestino(lista, reason) {
  const id = DATABASE_SHEET_ID;
  if (!id || lista.length === 0) return;
  const ts = nowISO();
  await ensureCestinoHeaders();

  const righe = lista.map(b => {
    // CESTINO ha 15 colonne: le prime 12 (ID→DELETED) coincidono con PRENOTAZIONI,
    // poi DELETED_AT(13), REASON(14), RIGA_ORIGINALE(15) — senza CLIENTE_ID.
    // Prendiamo solo i primi 12 elementi da bookingToDbRow (slice esclude CLIENTE_ID),
    // poi aggiungiamo le 3 colonne specifiche del CESTINO.
    const base = bookingToDbRow(b, b.fonte || 'app').slice(0, 12);
    base[DB_COLS.DELETED-1] = 'true'; // indice 11 = DELETED
    base.push(ts);                     // indice 12 = DELETED_AT
    base.push(reason || 'motivo non specificato'); // indice 13 = REASON
    base.push(b.dbRow || '');          // indice 14 = RIGA_ORIGINALE
    return base;
  });

  // Auto-sync: non scrive nel CESTINO (evita 4000+ righe inutili)
  // MA deve comunque eliminare fisicamente le righe dal foglio PRENOTAZIONI,
  // altrimenti al prossimo fast-read le rilegge e il ciclo è infinito.
  const isAutoSync = _isCestinoAutoSync(reason);
  if (isAutoSync) {
    syncLog(`🗑 Auto-sync: ${lista.length} pren. rimossa/e da DB senza scrivere nel CESTINO`, 'syn');
    // FALL-THROUGH: esegue comunque la delete fisica sotto ↓
  } else {
    // Scrivi nel CESTINO solo per cancellazioni esplicite utente
    await apiFetch(
      `https://sheets.googleapis.com/v4/spreadsheets/${id}/values/${encodeURIComponent(CESTINO_SHEET_NAME)}:append?valueInputOption=RAW&insertDataOption=INSERT_ROWS`,
      { method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({values:righe}) }
    );
  }

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
      // Scrive solo le colonne PRENOTAZIONI (A-M), senza aggiungere DELETED_AT/REASON
      const data = conRiga.map(b => {
        const row = bookingToDbRow(b, b.fonte || 'app');
        row[DB_COLS.DELETED-1] = 'true';
        row[DB_COLS.TS-1]      = ts;
        // CLIENTE_ID rimane in posizione 12 — non aggiungiamo campi del CESTINO
        const lastCol = String.fromCharCode(64 + row.length); // 'M' per 13 colonne
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
  // Aggiorna la blacklist in-memory solo per cancellazioni esplicite utente.
  // Le cancellazioni da sync automatica NON entrano in blacklist: se la prenotazione
  // torna nel foglio grafico (es. dopo una sync anomala) deve poter essere reimportata.
  // Non scrivere nel foglio CESTINO le cancellazioni automatiche da sync:
  // ingombrano il foglio (4872+ righe!) senza utilità — la protezione
  // è garantita da _row46BlipIds e dai guard in syncWithDatabase.
  if (!_isCestinoAutoSync(reason)) {
    lista.forEach(b => { if (b.dbId) addToCestinoBlacklist(b.dbId); });
  }
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
  // _getMonthSheetTitles usa cache TTL 30min → 0 chiamate aggiuntive se già letta da readJSONAnnuale
  const sheetNames = await _getMonthSheetTitles(sheetId);
  if (!sheetNames.length) return result;

  for (const sName of sheetNames) {
    try {
      const base = `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values/`;
      const enc  = s => encodeURIComponent(s);
      const [hR, jR, idR] = await Promise.all([
        apiFetch(base+enc(`'${sName}'!B${HEADER_ROW}:AJ${HEADER_ROW}`)+'?valueRenderOption=FORMATTED_VALUE'),
        apiFetch(base+enc(`'${sName}'!B${OUTPUT_ROW}:AJ${OUTPUT_ROW}`)+'?valueRenderOption=FORMATTED_VALUE'),
        apiFetch(base+enc(`'${sName}'!B${BLIP_ID_ROW}:AJ${BLIP_ID_ROW}`)+'?valueRenderOption=FORMATTED_VALUE'),
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

      // ── Scansiona TUTTA la riga 46 per raccogliere i BLIP_ID ──────
      // Anche le colonne con riga 45 vuota (Apps Script non girato):
      // così syncWithDatabase non cestina prenotazioni presenti nel foglio.
      idRow.forEach(cell => {
        if (!cell || !cell.trim()) return;
        try {
          const idMap = JSON.parse(String(cell).trim());
          if (typeof idMap === 'object' && !Array.isArray(idMap)) {
            Object.keys(idMap).forEach(k => { if (k.startsWith('PRE-')) _row46BlipIds.add(k); });
          }
        } catch(e) {
          const s = String(cell).trim();
          if (s.startsWith('PRE-')) _row46BlipIds.add(s);
        }
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
            // Legge la mappa BLIP_ID dalla riga 46 ({dbId:[dal,al],...})
            const colIdx   = sheetColumnMap[sName]?.[String(b.camera).trim()] || null;
            const blipColI = colIdx ? colIdx - 2 : null;
            const rawId46  = (blipColI !== null && idRow[blipColI]) ? String(idRow[blipColI]).trim() : null;
            let blipId = null;
            if (rawId46) {
              try {
                // Formato nuovo: JSON map {dbId:[dal,al]}
                const idMap = JSON.parse(rawId46);
                // Registra TUTTI i BLIP_ID della colonna nel set globale —
                // anche quelli non matchati da riga 45 (Apps Script non aggiornata)
                Object.keys(idMap).forEach(k => { if (k.startsWith('PRE-')) _row46BlipIds.add(k); });
                // Cerca la prenotazione per dal+al
                blipId = Object.keys(idMap).find(k => {
                  const [mapDal, mapAl] = idMap[k];
                  return mapDal === b.dal && mapAl === b.al;
                }) || null;
              } catch(e) {
                // Formato vecchio: singolo ID (retrocompatibilità)
                blipId = rawId46.startsWith('PRE-') ? rawId46 : null;
                if (blipId) _row46BlipIds.add(blipId);
              }
            }
            result.push({
              id:nid++, r:room.id, n:(b.nome||'—').replace(/\s*\+\s*$/, '').trim(), d:b.disposizione||'',
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
  // Legge A2:O13: colonna A = JSON mensile, colonna O = fingerprint mensile
  // In una sola chiamata otteniamo sia i dati che il fingerprint — zero costo aggiuntivo.
  // NOTA: il range va nel path URL (non in query param) → encodeURIComponent codifica
  // anche '!' e ':' producendo HTTP 400 su Chrome desktop. Usiamo solo replace degli spazi.
  const range = TAB.replace(/ /g, '%20') + "!A2:O13";
  const url   = "https://sheets.googleapis.com/v4/spreadsheets/" + sheetId + "/values/" + range + "?valueRenderOption=FORMATTED_VALUE";

  for (let attempt = 0; attempt < 3; attempt++) {
    const r = await apiFetch(url);
    if (r.status === 429) { await new Promise(res => setTimeout(res, (attempt+1)*2000)); continue; }
    if (!r.ok) throw new Error(TAB + ' non disponibile (HTTP ' + r.status + ')');
    const data = await r.json();
    const rows = data.values;
    if (!rows || rows.length === 0) throw new Error(TAB + ' vuoto — esegui "Rigenera JSON_ANNUALE" dal menu del foglio');

    // Colonna A (indice 0) = JSON, colonna O (indice 14) = fingerprint
    // Aggiorna i fingerprint in memoria (usati da bgSync per skip intelligente)
    const MESI_NOMI = ['Gennaio','Febbraio','Marzo','Aprile','Maggio','Giugno',
                       'Luglio','Agosto','Settembre','Ottobre','Novembre','Dicembre'];
    rows.forEach((row, i) => {
      const fp = row?.[14]; // colonna O = indice 14
      if (fp) {
        const anno = new Date().getFullYear();
        const sName = (MESI_NOMI[i] || '') + ' ' + anno;
        _sheetFingerprints[sName] = String(fp).trim();
      }
    });

    const allPren = [];
    let parseErrors = 0;
    for (const row of rows) {
      const raw = row?.[0]; // colonna A
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
    // Popola sheetColumnMap leggendo le intestazioni di tutti i fogli mensili
    // in una sola chiamata batchGet — necessario per assegnare _sheetCol ai booking.
    // _getMonthSheetTitles usa la cache (TTL 30min) → 0 chiamate aggiuntive se già letta.
    try {
      const monthSheets = await _getMonthSheetTitles(sheetId);
      {
        if (monthSheets.length > 0) {
          // Legge in una sola chiamata: intestazioni (riga HEADER_ROW) + riga 46 (BLIP_ID_ROW)
          // per tutti i fogli mensili. I fingerprint ora sono in JSON_ANNUALE!O2:O13 — letti
          // da readJSONAnnuale senza costo aggiuntivo.
          const headerRanges = monthSheets
            .map(sn => encodeURIComponent("'" + sn + "'!B" + HEADER_ROW + ":AJ" + HEADER_ROW));
          const blipIdRanges = monthSheets
            .map(sn => encodeURIComponent("'" + sn + "'!B" + BLIP_ID_ROW + ":AJ" + BLIP_ID_ROW));
          const allRanges = [...headerRanges, ...blipIdRanges].join('&ranges=');
          const bR = await apiFetch(
            'https://sheets.googleapis.com/v4/spreadsheets/' + sheetId +
            '/values:batchGet?ranges=' + allRanges + '&valueRenderOption=FORMATTED_VALUE'
          );
          if (bR.ok) {
            const bData = await bR.json();
            const vRanges = bData.valueRanges || [];
            const n = monthSheets.length;
            // Prime n ranges = intestazioni
            vRanges.slice(0, n).forEach((vr, idx) => {
              const sn = monthSheets[idx];
              const headers = vr.values?.[0] || [];
              sheetColumnMap[sn] = {};
              headers.forEach((h, i) => {
                if (!h) return;
                const raw = String(h).trim(), norm = raw.replace(/\.0$/, '');
                sheetColumnMap[sn][raw] = i + 2;
                if (norm !== raw) sheetColumnMap[sn][norm] = i + 2;
              });
            });
            // Ultime n ranges = riga 46 → popola _row46BlipIds e _row46BookingMap
            vRanges.slice(n).forEach((vr, idx) => {
              const sn     = monthSheets[idx];
              const idRow  = vr.values?.[0] || [];
              const hdrs   = sheetColumnMap[sn] || {};
              // Costruisce mappa inversa: colIdx → cameraName
              const colToCamera = {};
              Object.entries(hdrs).forEach(([cam, col]) => { colToCamera[col - 2] = cam; });

              idRow.forEach((cell, colI) => {
                if (!cell || !cell.trim()) return;
                try {
                  const idMap = JSON.parse(String(cell).trim());
                  if (typeof idMap === 'object' && !Array.isArray(idMap)) {
                    const camName = colToCamera[colI] || '';
                    Object.entries(idMap).forEach(([k, v]) => {
                      if (!k.startsWith('PRE-')) return;
                      _row46BlipIds.add(k);
                      // v = [dal, al] es. ["24/07/2026","07/08/2026"]
                      const dal = Array.isArray(v) ? v[0] : '';
                      if (camName && dal) {
                        _row46BookingMap[sn + '|' + camName + '|' + dal] = k;
                      }
                    });
                  }
                } catch(e) {
                  const s = String(cell).trim();
                  if (s.startsWith('PRE-')) _row46BlipIds.add(s);
                }
              });
            });
            console.log('[JSON_ANNUALE] sheetColumnMap + riga46 popolati per ' + monthSheets.length + ' fogli, _row46BlipIds: ' + _row46BlipIds.size + ', bookingMap: ' + Object.keys(_row46BookingMap).length);
          }
        }
      }
    } catch(hErr) { console.warn('[JSON_ANNUALE] intestazioni:', hErr.message); }

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
    // Assegna _sheetCol se sheetColumnMap è già popolato per questo foglio
    const _colMap = sheetColumnMap[sName] || {};
    const _sheetCol = _colMap[room.name] || _colMap[String(room.name).trim()] || null;
    // Cerca il BLIP_ID dalla riga 46 tramite la mappa sheetName|camera|dal
    // Questo permette a findMatch di usare la PRIORITÀ 1 (ID) invece della fuzzy
    // evitando duplicati quando il nome in DB differisce leggermente (es. trailing +)
    const _mapKey = sName + '|' + room.name + '|' + b.dal;
    const _dbId   = _row46BookingMap[_mapKey] || null;
    result.push({
      id: nid++, r: room.id, n: (b.nome||'—').replace(/\s*\+\s*$/, '').trim(), d: b.disposizione||'',
      c: color, s: new Date(yy,mm-1,dd,12), e: new Date(ye,me-1,de,12),
      note: b.note||'', fromSheet:true, fromJSONAnnuale:true,
      sheetName:sName, sheetId, cameraName:room.name,
      dbId: _dbId, dbRow:null, ts:null, fonte:'manuale',
      _sheetCol: _sheetCol || undefined,
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

// Normalizza il nome per il confronto in findMatch:
// - lowercase e trim
// - rimuove '+' isolati finali (lasciati da Apps Script come separatore)
// - normalizza spazi multipli interni
function _normName(s) {
  return (s||'').toLowerCase().trim()
    .replace(/\s*\+\s*$/, '')   // trailing + con spazi opzionali
    .replace(/\s+/g, ' ')       // spazi multipli interni
    .trim();
}

function findMatch(target, list) {
  // PRIORITÀ 1: match per BLIP_ID dalla riga 46 — match perfetto, immune a spostamenti
  if (target.dbId) {
    const byId = list.find(b => b.dbId === target.dbId);
    if (byId) return byId;
  }

  const camT  = (target.cameraName || roomName(target.r) || '').toLowerCase().trim();
  const nomT  = _normName(target.n);
  const dayT  = Math.round((target.s?.getTime?.() || 0) / DAY_MS);
  const dispT = (target.d || '').trim().toLowerCase();

  // PRIORITÀ 2: match esatto per nome normalizzato + camera + data
  let m = list.find(b => {
    if (_normName(b.n) !== nomT) return false;
    if ((b.cameraName||roomName(b.r)||'').toLowerCase().trim() !== camT) return false;
    return Math.round((b.s?.getTime?.() || 0) / DAY_MS) === dayT;
  });
  if (m) return m;

  // PRIORITÀ 3: fuzzy — nome normalizzato + camera + data ±1 giorno
  m = list.find(b => {
    if (_normName(b.n) !== nomT) return false;
    if ((b.cameraName||roomName(b.r)||'').toLowerCase().trim() !== camT) return false;
    if (dispT && (b.d||'').trim().toLowerCase() !== dispT) return false;
    return Math.abs(Math.round((b.s?.getTime?.() || 0) / DAY_MS) - dayT) <= 1;
  });
  if (m) return m;

  // PRIORITÀ 4: overlap — frammenti di prenotazioni multi-mese
  // ─────────────────────────────────────────────────────────────
  // Caso: il DB conserva la data di inizio reale (es. 15 aprile),
  // ma JSON_ANNUALE presenta il mese di maggio come frammento con
  // s = 1 maggio. Le priorità 2 e 3 non trovano il match perché
  // cercano la data di inizio esatta.
  // Soluzione: stessa camera + stesso nome normalizzato + date che
  // si sovrappongono (overlap standard: sDb < eTarget && eDb > sTarget).
  // Anno uguale come guard extra per non confondere anni diversi.
  // Questa priorità è più "larga" ma è sicura perché richiede sia
  // nome che camera identici — l'unico elemento variabile è la data.
  const sT = target.s?.getTime?.() || 0;
  const eT = target.e?.getTime?.() || 0;
  const yT = target.s ? new Date(target.s).getFullYear() : 0;
  m = list.find(b => {
    if (_normName(b.n) !== nomT) return false;
    if ((b.cameraName||roomName(b.r)||'').toLowerCase().trim() !== camT) return false;
    const sDb = b.s?.getTime?.() || 0;
    const eDb = b.e?.getTime?.() || 0;
    const yDb = b.s ? new Date(b.s).getFullYear() : 0;
    if (yDb !== yT) return false;
    // ±1 giorno di tolleranza per gestire frammenti adiacenti a cambio mese:
    // es. frammento aprile termina 30/04 12:00, frammento maggio inizia 01/05 12:00
    // → eDb (30/04) + DAY_MS = 01/05 > sT (01/05): adiacenti = match
    return sDb < eT + DAY_MS && eDb + DAY_MS > sT;
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

// ─────────────────────────────────────────────────────────────────
// readDatabaseAndCestino — batchGet unico per PRENOTAZIONI + CESTINO
//
// Fonde readDatabase() + loadCestinoBlacklist() in 1 sola chiamata API
// invece di 2. Risparmio critico in syncWithDatabase, che è il momento
// in cui il token bucket è già consumato dal caricamento iniziale.
//
// Restituisce { dbRows, blacklist } dove:
//   dbRows    = array di booking (già filtrati !deleted) come readDatabase()
//   blacklist = Set<string> di BLIP_ID cestinati dall'utente (non da sync auto)
// ─────────────────────────────────────────────────────────────────
async function readDatabaseAndCestino() {
  const id = DATABASE_SHEET_ID;
  if (!id) return { dbRows: [], blacklist: new Set() };

  const enc = s => encodeURIComponent(s);
  const batchUrl = 'https://sheets.googleapis.com/v4/spreadsheets/' + id
    + '/values:batchGet?ranges=' + enc(`${DB_SHEET_NAME}!A${DB_FIRST_ROW}:O3000`)
    + '&ranges=' + enc(`${CESTINO_SHEET_NAME}!A2:A9999`)
    + '&ranges=' + enc(`${CESTINO_SHEET_NAME}!N2:N9999`)
    + '&valueRenderOption=FORMATTED_VALUE';

  const ctrl = new AbortController();
  const tid  = setTimeout(() => ctrl.abort(), 30000);
  let batchResp;
  try {
    batchResp = await apiFetch(batchUrl, { signal: ctrl.signal });
  } catch(e) {
    if (e.name === 'AbortError') throw new Error('DB+CESTINO batchGet timeout (>30s)');
    throw e;
  } finally { clearTimeout(tid); }

  if (!batchResp.ok) throw new Error(`DB+CESTINO batchGet error ${batchResp.status}: ${await batchResp.text()}`);
  const bj = await batchResp.json();
  const vr = bj.valueRanges || [];

  // ── Parsing PRENOTAZIONI ──────────────────────────────────────
  const dbRawRows = vr[0]?.values || [];
  dbRowCache = [];
  const dbRows = [];
  dbRawRows.forEach((row, i) => {
    const rowNum = DB_FIRST_ROW + i;
    const b = dbRowToBooking(row, rowNum);
    dbRowCache.push({ rowNum, raw: row });
    if (b && !b.deleted) dbRows.push(b);
  });

  // ── Parsing CESTINO blacklist ─────────────────────────────────
  const idsCol    = vr[1]?.values || [];
  const reasonCol = vr[2]?.values || [];
  const blacklisted = [];
  idsCol.forEach((row, i) => {
    const id_     = (row[0]  || '').trim();
    const reason  = ((reasonCol[i] || [])[0] || '').trim();
    if (!id_) return;
    if (!_isCestinoAutoSync(reason)) blacklisted.push(id_);
  });
  const blacklist = new Set(blacklisted);

  // Aggiorna la blacklist globale (TTL reset)
  _cestinoBlacklist    = blacklist;
  _cestinoBlacklistTs  = Date.now();
  syncLog(`🗑 Blacklist CESTINO: ${blacklisted.length} ID (utente) [batchGet fusione DB+CESTINO]`, 'syn');

  return { dbRows, blacklist };
}

// fromFallback=true quando la fonte è il fallback dei 12 fogli mensili (JSON_ANNUALE non disponibile).
// In modalità fallback la lettura può essere incompleta (429 su singoli fogli, fogli mancanti,
// nomi non standard) → non è affidabile come fonte di verità per le CANCELLAZIONI.
// Con fromFallback=true: la FASE 2 preserva tutti i booking DB non-matchati (nessuna cestinazione),
// MAX_ARCHIVE_PER_SYNC scende a 0 (sicurezza assoluta).
async function syncWithDatabase(sheetBookings, forceFullSync = false, fromFallback = false) {
  if (!DATABASE_SHEET_ID) return sheetBookings;

  // ── 1 sola chiamata batchGet per PRENOTAZIONI + CESTINO blacklist ──
  // (sostituisce readDatabase() + loadCestinoBlacklist() separati = -1 chiamata API)
  const { dbRows: allDbRows } = await readDatabaseAndCestino();
  const cleanedN  = await cleanupDeletedFromDb(allDbRows);
  if (cleanedN > 0) {
    showLoading(`Cestino: ${cleanedN} righe spostate…`);
    // Righe DELETED residue trovate: ricarica blacklist aggiornata (1 chiamata extra, rara)
    await loadCestinoBlacklist(true);
  }
  // _cestinoBlacklist già aggiornata da readDatabaseAndCestino — nessuna chiamata aggiuntiva

  if (fromFallback) {
    syncLog('⚠ syncWithDatabase: fonte fallback (12 fogli) — cestinazione disabilitata per sicurezza', 'wrn');
  }

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
    if (isDeletedLocally(sheet.dbId)) { syncLog(`🗑 Skip blacklist locale: ${sheet.n}`, 'wrn'); continue; }
    // Salta prenotazioni già nel CESTINO remoto — blocca il reimport di prenotazioni
    // cancellate manualmente dal DB o tramite app (anche dopo cancellazione manuale)
    if (isInCestinoBlacklist(sheet.dbId)) {
      syncLog(`🗑 Skip CESTINO: ${sheet.n} (${sheet.dbId})`, 'wrn');
      continue;
    }
    const match = findMatch(sheet, dbActive);
    if (!match) {
      sheet.dbId  = genBookingId(sheet.s.getFullYear());
      sheet.ts    = nowISO();
      sheet.fonte = 'manuale';
      toAddToDB.push(sheet);
      result.push(sheet);
    } else {
      seenDbIds.add(match.dbId);

      // ── Deduplicazione frammenti multi-mese ──────────────────────────
      // Quando findMatch trova il frammento "principale" (di solito il primo
      // mese), altri frammenti dello stesso booking rimangono nel DB come
      // fantasmi — stessa camera e nome, date adiacenti/sovrapposte.
      // Li identifichiamo e li aggiungiamo a toArchive per pulizia automatica.
      // Condizioni: stesso cameraName, nome normalizzato uguale, date che si
      // sovrappongono o sono adiacenti (± 2 giorni) al booking del foglio,
      // BLIP_ID diverso dal match principale.
      const camMatch  = (match.cameraName || '').toLowerCase().trim();
      const nomMatch  = _normName(match.n);
      const sSheet    = sheet.s?.getTime?.() || 0;
      const eSheet    = sheet.e?.getTime?.() || 0;
      for (const db of dbActive) {
        if (db.dbId === match.dbId) continue;           // il match principale
        if (seenDbIds.has(db.dbId)) continue;           // già processato
        if (isDeletedLocally(db.dbId)) continue;        // già cancellato
        if (isInCestinoBlacklist(db.dbId)) continue;   // già nel CESTINO
        if ((db.cameraName || '').toLowerCase().trim() !== camMatch) continue;
        if (_normName(db.n) !== nomMatch) continue;
        const sDb = db.s?.getTime?.() || 0;
        const eDb = db.e?.getTime?.() || 0;
        // Sovrapposto o adiacente (entro 2 giorni)
        const sovrapposto = sDb <= eSheet + 2*DAY_MS && eDb >= sSheet - 2*DAY_MS;
        if (!sovrapposto) continue;
        // È un frammento duplicato — marca per rimozione
        seenDbIds.add(db.dbId); // non finirà nella fase 2
        db.deleted      = true;
        db.deleteReason = `Frammento duplicato di ${match.dbId} · dedup del ${new Date().toLocaleDateString('it-IT')}`;
        db.deletedAt    = nowISO();
        toArchive.push(db);
        syncLog(`🧹 Frammento duplicato rimosso: ${db.n} (${db.dbId}) → unificato con ${match.dbId}`, 'syn');
      }

      // Propaga info del foglio all'oggetto DB (necessario per backfill riga 46)
      if (sheet.sheetId)    match.sheetId    = sheet.sheetId;
      if (sheet.sheetName)  match.sheetName  = sheet.sheetName;
      if (sheet.cameraName) match.cameraName = sheet.cameraName;
      if (sheet._sheetCol)  match._sheetCol  = sheet._sheetCol;
      match.fromSheet = true;
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
    // 2 ore: tempo sufficiente perché il bridge scriva sul foglio e il trigger
    // Apps Script rigeneri JSON_ANNUALE. 15 min era troppo poco — se il bridge
    // è lento o la prenotazione non è ancora apparsa nel foglio, veniva cestinata.
    const isRecenteApp = db.fonte === 'app' && etaMs < 2 * 60 * 60 * 1000;
    if (isRecenteApp) {
      syncLog(`🛡 Protetta (app ${Math.round(etaMs/60000)}min fa): ${db.n} cam.${db.cameraName||db.r}`, 'wrn');
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

    // GUARD anni non coperti: se il foglio non ha NESSUNA prenotazione nell'anno
    // della prenotazione DB, probabilmente stiamo leggendo un anno diverso.
    // Non cestinare prenotazioni di anni non rappresentati nel foglio corrente.
    const sheetHasYear = sheetBookings.some(s =>
      s.s && new Date(s.s).getFullYear() === dbYear
    );
    if (!sheetHasYear) {
      result.push(db); continue;
    }

    // GUARD riga 46: se il BLIP_ID è presente nella riga 46 del foglio grafico
    // significa che la prenotazione esiste fisicamente nel foglio ma Apps Script
    // non ha aggiornato la riga 45 (JSON). Non cestinare — è visibile nel foglio.
    if (db.dbId && _row46BlipIds.has(db.dbId)) {
      syncLog(`🛡 Protetta da riga 46: ${db.n} (${db.dbId})`, 'wrn');
      result.push(db); continue;
    }

    // GUARD merge-mese: Apps Script unisce prenotazioni multi-mese con nomi leggermente
    // diversi (es. "Erasmus" + "Erasmus 24s" → un solo booking in JSON_ANNUALE).
    // Se il DB ha un entry con nome simile sulla stessa camera nello stesso mese,
    // è quasi certamente un frammento del booking unito — non cestinare.
    if (db.dbId && db.s && db.cameraName) {
      const dbNomNorm = _normName(db.n);
      const dbCam     = (db.cameraName || '').toLowerCase().trim();
      const dbMese    = new Date(db.s).getMonth();
      const dbAnno    = new Date(db.s).getFullYear();
      const hasSimilarInSheet = sheetBookings.some(s => {
        if (!s.s || !s.cameraName) return false;
        if ((s.cameraName||'').toLowerCase().trim() !== dbCam) return false;
        if (new Date(s.s).getFullYear() !== dbAnno) return false;
        // Stessa camera, stesso anno: se il nome del foglio CONTIENE il nome DB o viceversa
        const sNom = _normName(s.n);
        return sNom.includes(dbNomNorm) || dbNomNorm.includes(sNom);
      });
      if (hasSimilarInSheet) {
        syncLog(`🛡 Protetta (merge-mese): ${db.n} (${db.dbId})`, 'wrn');
        result.push(db); continue;
      }
    }

    // GUARD fallback: in modalità fallback (12 fogli mensili) la lettura può essere
    // incompleta — non cestinare MAI in questo caso. La prenotazione rimane visibile.
    if (fromFallback) {
      result.push(db); continue;
    }

    // GUARD recente-DB: non cestinare prenotazioni create/aggiornate negli ultimi 30 giorni
    // che non appaiono nel foglio. Potrebbe essere latenza bridge (Apps Script non ha
    // ancora scritto nel foglio) o una prenotazione inserita direttamente nel DB.
    // Dopo 30 giorni, se ancora non è nel foglio, è sicuramente un residuo.
    // forceFullSync bypassa anche questo guard (🔄 esplicito dell'utente).
    if (!forceFullSync && db.ts) {
      const etaCreazione = Date.now() - new Date(db.ts).getTime();
      if (etaCreazione < 30 * 24 * 60 * 60 * 1000) {
        syncLog('🛡 Protetta (recente ' + Math.round(etaCreazione/86400000) + 'gg): ' + db.n + ' (' + (db.dbId||'no-id') + ')', 'wrn');
        result.push(db); continue;
      }
    }

    db.deleted      = true;
    db.deleteReason = 'Rimossa dal foglio Gantt · sync del ' + new Date().toLocaleDateString('it-IT');
    db.deletedAt    = nowISO();
    toArchive.push(db);
  }

  // FASE 3: Scrittura DB
  // Separa i frammenti duplicati (certi) dai candidati normali (incerti)
  // I frammenti hanno deleteReason che inizia con "Frammento duplicato"
  const toArchiveFragmenti = toArchive.filter(b => b.deleteReason?.startsWith('Frammento duplicato'));
  const toArchiveNormali   = toArchive.filter(b => !b.deleteReason?.startsWith('Frammento duplicato'));

  // ── GUARD ASSOLUTO anti-cestinazione massiva ──────────────────────
  // Si applica solo ai candidati normali — i frammenti duplicati sono certi e sicuri.
  const MAX_ARCHIVE_PER_SYNC = fromFallback ? 0 : 20;
  if (toArchiveNormali.length > MAX_ARCHIVE_PER_SYNC && !forceFullSync) {
    syncLog(`🛑 STOP: ${toArchiveNormali.length} prenotazioni da cestinare — limite sicurezza (${MAX_ARCHIVE_PER_SYNC}) superato. Usa 🔄 per forzare.`, 'err');
    showToast(`⚠ ${toArchiveNormali.length} prenotazioni da cestinare — operazione bloccata per sicurezza. Usa 🔄 se intenzionale.`, 'error');
    toArchiveNormali.forEach(b => { b.deleted = false; result.push(b); });
    toArchiveNormali.length = 0;
    // Ricostruisce toArchive con soli i frammenti (questi procedono sempre)
    toArchive.length = 0;
    toArchiveFragmenti.forEach(b => toArchive.push(b));
  }

  if (toArchiveFragmenti.length > 0) {
    syncLog(`🧹 ${toArchiveFragmenti.length} frammenti duplicati da rimuovere`, 'syn');
  }

  // ── GUARD anti-import massivo ─────────────────────────────────────
  // Se ci sono molte prenotazioni nuove da importare, le importiamo a tranche
  // (MAX_ADD_PER_SYNC per sessione). Le restanti verranno importate al prossimo bgSync.
  // Questo evita che il sync si blocchi per minuti su mobile o connessioni lente.
  // forceFullSync bypassa il limite.
  const MAX_ADD_PER_SYNC = forceFullSync ? Infinity : 50;
  let addSkipped = 0;
  if (toAddToDB.length > MAX_ADD_PER_SYNC) {
    addSkipped = toAddToDB.length - MAX_ADD_PER_SYNC;
    toAddToDB.splice(MAX_ADD_PER_SYNC); // tronca — le restanti arrivano nel result dal foglio
    syncLog(`⚠ Import parziale: ${MAX_ADD_PER_SYNC}/${MAX_ADD_PER_SYNC + addSkipped} nuove pren. (le altre al prossimo sync)`, 'wrn');
  }

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
    // LOG DIAGNOSTICO: stampa i candidati alla cestinazione prima di procedere
    toArchive.forEach(b => {
      const anno = b.s ? new Date(b.s).getFullYear() : '?';
      syncLog(`🗑 Candidato cestino: ${b.n} (${b.dbId||'no-id'}) anno=${anno} dbRow=${b.dbRow||'?'}`, 'wrn');
    });
    showLoading(`Cestino: ${toArchive.length} prenotazioni…`);
    try { await archiviaInCestino(toArchive, 'Rimossa dal foglio Gantt · ' + new Date().toLocaleDateString('it-IT')); } catch(e) {}
    const rimossi = toArchive.filter(b=>b.dbRow).length;
    syncLog(`🗑 ${rimossi} → CESTINO`, 'wrn');
  }

  const dbActiveCount = dbActive.length;
  syncLog(`DB: ${result.length} prenotazioni attive, rimossi ${toArchive.length}`, 'db');

  // Scrivi BLIP_ID nella riga 46 per le prenotazioni nuove (fire & forget)
  // Skip se l'import era grande: il token bucket è già sotto pressione e
  // writeBlipIdsToRow46 farebbe scattare 429. Il bridge gestisce riga 46 lato Apps Script.
  const toWriteRow46 = toAddToDB.filter(b => b.dbId && b._sheetCol && b.sheetName && b.sheetId);
  if (toWriteRow46.length > 0 && toWriteRow46.length <= 10 && addSkipped === 0) {
    writeBlipIdsToRow46(toWriteRow46).catch(e =>
      syncLog('⚠ Scrittura riga 46: ' + e.message, 'wrn')
    );
  } else if (toWriteRow46.length > 10) {
    syncLog(`⏭ Scrittura riga 46 rinviata (${toWriteRow46.length} celle — evita 429)`, 'syn');
  }

  return result;
}

// Scrive i BLIP_ID nella riga 46 del foglio visivo
async function writeBlipIdsToRow46(bookings) {
  // ═══════════════════════════════════════════════════════════════
  // REGOLA FONDAMENTALE: non sovrascrivere mai una cella riga 46
  // che contiene già un BLIP_ID valido (formato {"PRE-...":...}).
  // In caso di cella già popolata → MERGE: aggiunge solo ID nuovi,
  // non rimuove quelli esistenti. Garantisce che backfillRow46 e
  // le scritture automatiche da sync non distruggano le protezioni.
  // ═══════════════════════════════════════════════════════════════

  const fmtDate = d => {
    if (!d) return '';
    const dt = d instanceof Date ? d : new Date(d);
    return dt.toLocaleDateString('it-IT', {day:'2-digit',month:'2-digit',year:'numeric'});
  };

  // Raggruppa per sheet+colonna: ogni cella contiene un JSON map {dbId:[dal,al],...}
  const byCell = {}; // key = sheetId|sName|col
  for (const b of bookings) {
    if (!b.dbId || !b._sheetCol || !b.sheetName || !b.sheetId) continue;
    const key = b.sheetId + '|' + b.sheetName + '|' + b._sheetCol;
    if (!byCell[key]) byCell[key] = { sheetId:b.sheetId, sName:b.sheetName, col:b._sheetCol, map:{} };
    byCell[key].map[b.dbId] = [fmtDate(b.s), fmtDate(b.e)];
  }

  // Raggruppa per foglio
  const bySheet = {}; // key = sheetId|sName
  for (const { sheetId, sName, col, map } of Object.values(byCell)) {
    const sk = sheetId + '|' + sName;
    if (!bySheet[sk]) bySheet[sk] = { sheetId, sName, cells:[] };
    bySheet[sk].cells.push({ col, map });
  }

  for (const { sheetId, sName, cells } of Object.values(bySheet)) {
    if (!cells.length) continue;

    // ── STEP 1: leggi riga 46 esistente per l'intero range delle colonne ──
    const cols = cells.map(c => c.col);
    const minCol = Math.min(...cols);
    const maxCol = Math.max(...cols);
    const existingRow46 = {}; // col → stringa JSON attuale
    try {
      const range = "'" + sName + "'!" + columnLetter(minCol) + BLIP_ID_ROW + ':' + columnLetter(maxCol) + BLIP_ID_ROW;
      const r = await apiFetch(
        'https://sheets.googleapis.com/v4/spreadsheets/' + sheetId +
        '/values/' + encodeURIComponent(range) + '?valueRenderOption=FORMATTED_VALUE'
      );
      if (r.ok) {
        const j = await r.json();
        const row = j.values?.[0] || [];
        row.forEach((val, idx) => { existingRow46[minCol + idx] = val || ''; });
      }
    } catch(e) { syncLog('⚠ Lettura pre-write riga 46 (' + sName + '): ' + e.message, 'wrn'); }

    // ── STEP 2: per ogni cella, decide se scrivere e cosa scrivere ──
    const data = [];
    let skipped = 0;
    for (const { col, map } of cells) {
      const existing = existingRow46[col] || '';
      if (existing && existing.includes('"PRE-')) {
        // Cella già popolata — MERGE: aggiunge solo ID nuovi, mai rimuove
        try {
          const existingMap = JSON.parse(existing);
          const mergedMap   = { ...existingMap }; // parte dagli esistenti
          let added = 0;
          for (const [id, dates] of Object.entries(map)) {
            if (!mergedMap[id]) { mergedMap[id] = dates; added++; } // solo ID nuovi
          }
          if (added > 0) {
            data.push({
              range:  "'" + sName + "'!" + columnLetter(col) + BLIP_ID_ROW,
              values: [[JSON.stringify(mergedMap)]]
            });
          } else {
            skipped++; // cella già aggiornata — skip
          }
        } catch(e) {
          // JSON corrotto — sovrascrive con i nuovi dati
          data.push({
            range:  "'" + sName + "'!" + columnLetter(col) + BLIP_ID_ROW,
            values: [[JSON.stringify(map)]]
          });
        }
      } else {
        // Cella vuota o senza PRE- → scrivi normalmente
        data.push({
          range:  "'" + sName + "'!" + columnLetter(col) + BLIP_ID_ROW,
          values: [[JSON.stringify(map)]]
        });
      }
    }

    if (skipped > 0) syncLog('✓ Riga 46 ' + sName + ': ' + skipped + ' celle già aggiornate (skip)', 'ok');
    if (!data.length) continue;

    // ── STEP 3: batchUpdate solo delle celle che richiedono scrittura ──
    try {
      const url = 'https://sheets.googleapis.com/v4/spreadsheets/' + sheetId + '/values:batchUpdate';
      const resp = await fetch(url, {
        method: 'POST',
        headers: { Authorization: 'Bearer ' + accessToken, 'Content-Type': 'application/json' },
        body: JSON.stringify({ valueInputOption: 'RAW', data })
      });
      if (resp.ok) syncLog('✓ BLIP_ID scritti riga 46 in ' + sName + ': ' + data.length + ' celle', 'ok');
      else syncLog('⚠ Errore scrittura riga 46 in ' + sName + ' (HTTP ' + resp.status + ')', 'wrn');
    } catch(e) { syncLog('⚠ writeBlipIdsToRow46: ' + e.message, 'wrn'); }
  }
}

// ── Backfill riga 46: usa JSON_ANNUALE come fonte di verità ──────────────────
//
// REGOLA: la riga 46 deve rispecchiare esattamente il contenuto del foglio
// grafico (JSON_ANNUALE). Blip non inventa ID — li ricava dal foglio.
//
// Flusso:
//   1. Rilegge JSON_ANNUALE fresco (non usa bookings in memoria = misto DB+foglio)
//   2. Per ogni prenotazione senza dbId (riga 46 vuota), cerca il BLIP_ID nel DB
//   3. Chiama writeBlipIdsToRow46 con guard anti-sovrascrittura (merge)
//
// Può essere chiamato dalla console: backfillRow46()
// ─────────────────────────────────────────────────────────────────────────────
async function backfillRow46() {
  const sheetEntry = _currentYearSheetEntry();
  if (!sheetEntry) {
    showToast('Nessun foglio anno corrente configurato', 'error'); return;
  }

  showLoading('Lettura JSON_ANNUALE per backfill riga 46…');
  syncLog('📖 Backfill riga 46: rilettura JSON_ANNUALE (fonte di verità)…', 'syn');

  let sheetBookings = [];
  try {
    sheetBookings = await readJSONAnnuale(sheetEntry.sheetId);
  } catch(e) {
    hideLoading();
    showToast('Errore lettura JSON_ANNUALE: ' + e.message, 'error');
    syncLog('❌ Backfill: ' + e.message, 'err');
    return;
  }
  syncLog('📖 Backfill: ' + sheetBookings.length + ' prenotazioni da JSON_ANNUALE', 'syn');

  // Per le prenotazioni senza dbId (riga 46 vuota), cerca il BLIP_ID nel DB locale.
  // Usa findMatch contro bookings in memoria (che hanno dbId dal DB).
  // Questo è il passaggio che "associa" le prenotazioni del foglio ai record DB.
  const dbWithIds = bookings.filter(b => b.dbId);
  let assigned = 0, alreadyHad = 0;
  for (const sb of sheetBookings) {
    if (sb.dbId) { alreadyHad++; continue; } // già aveva ID da riga 46
    const match = findMatch(sb, dbWithIds);
    if (match) { sb.dbId = match.dbId; assigned++; }
  }
  syncLog('🔗 Backfill: ' + assigned + ' BLIP_ID assegnati da DB, ' + alreadyHad + ' già presenti', 'syn');

  // Filtra: solo prenotazioni con coordinate foglio complete
  const toWrite = sheetBookings.filter(b => b.dbId && b._sheetCol && b.sheetName && b.sheetId);
  const senzaCol = sheetBookings.filter(b => b.dbId && !b._sheetCol).length;
  if (senzaCol > 0) syncLog('⚠ Backfill: ' + senzaCol + ' pren. con dbId ma senza coordinata colonna (ignorati)', 'wrn');

  if (!toWrite.length) {
    hideLoading();
    showToast('Nessuna prenotazione con coordinate foglio trovata', 'warning');
    return;
  }

  showLoading('Scrittura riga 46 (' + toWrite.length + ' celle — guard anti-sovrascrittura attiva)…');
  await writeBlipIdsToRow46(toWrite);
  hideLoading();
  showToast('✓ Riga 46 aggiornata: ' + toWrite.length + ' celle (JSON_ANNUALE → fonte)', 'success');
  syncLog('✓ Backfill riga 46 completato: ' + toWrite.length + ' celle scritte', 'ok');
}

// Converte numero colonna (1-based) → lettera Excel (A, B, ..., Z, AA, ...)
function columnLetter(col) {
  let s = '';
  while (col > 0) { col--; s = String.fromCharCode(65 + col % 26) + s; col = Math.floor(col / 26); }
  return s;
}


// ═══════════════════════════════════════════════════════════════════
// RIPARAZIONE DATABASE — aggiorna legami BLIP_ID in tutti i fogli
// ═══════════════════════════════════════════════════════════════════

async function riparaDatabase() {
  if (!DATABASE_SHEET_ID) { showToast('DATABASE_SHEET_ID non configurato', 'error'); return; }
  if (!bookings.length)   { showToast('Carica prima le prenotazioni (↻)', 'error'); return; }

  syncLog('🔧 Avvio riparazione database…', 'syn');
  showLoading('Riparazione database…');
  let totFixed = 0, totWarn = 0;

  // ── Indici di ricerca dai bookings in memoria ──
  // (camera|dal) → dbId  e  (cameraName|nome8) → dbId (fuzzy)
  const camDalMap  = {};   // "cam|dd/MM/yyyy" → dbId
  const camNomeMap = {};   // "cam|nome8"       → dbId  (fallback)
  for (const b of bookings) {
    if (!b.dbId) continue;
    const cam = (b.cameraName || roomName(b.r) || '').trim();
    if (b.s) {
      const d  = b.s;
      const dal = String(d.getDate()).padStart(2,'0') + '/'
                + String(d.getMonth()+1).padStart(2,'0') + '/'
                + d.getFullYear();
      camDalMap[cam + '|' + dal] = b.dbId;
    }
    const n8 = (b.n || '').toLowerCase().trim().slice(0, 10);
    if (n8) camNomeMap[cam + '|' + n8] = b.dbId;
  }

  function blipFromContoJSON(c) {
    if (!c) return null;
    const cam = String(c.camera || '').trim();
    // da ISO a dd/MM/yyyy
    try {
      const dt  = new Date(c.checkin || '');
      const dal = String(dt.getDate()).padStart(2,'0') + '/'
                + String(dt.getMonth()+1).padStart(2,'0') + '/'
                + dt.getFullYear();
      const byDal = camDalMap[cam + '|' + dal];
      if (byDal) return byDal;
    } catch(e) {}
    // fallback fuzzy nome
    const n8 = (c.nome || '').toLowerCase().trim().slice(0, 10);
    return camNomeMap[cam + '|' + n8] || null;
  }

  // ── STEP 1: Header CLIENTE_ID in PRENOTAZIONI (col M = 13) ──
  try {
    showLoading('Verifica schema PRENOTAZIONI…');
    const hR  = await dbGet(DB_SHEET_NAME + '!A1:M1');
    const hdr = hR.values?.[0] || [];
    if (!hdr[12] || String(hdr[12]).trim() !== 'CLIENTE_ID') {
      const url = 'https://sheets.googleapis.com/v4/spreadsheets/' + DATABASE_SHEET_ID
                + '/values/' + encodeURIComponent(DB_SHEET_NAME + '!M1')
                + '?valueInputOption=RAW';
      await apiFetch(url, {
        method: 'PUT', headers: {'Content-Type':'application/json'},
        body: JSON.stringify({ values: [['CLIENTE_ID']] })
      });
      syncLog('✓ Header CLIENTE_ID aggiunto (col M)', 'ok');
      totFixed++;
    } else {
      syncLog('✓ PRENOTAZIONI.CLIENTE_ID: già presente', 'ok');
    }
  } catch(e) { syncLog('⚠ Header PRENOTAZIONI: ' + e.message, 'wrn'); totWarn++; }

  // ── STEP 2: CONTI — migra BOOKING_ID numerici → BLIP_ID ──
  const contoIdToBlip = {};   // C17xxxxx → PRE-2026-xxx  (per collegare i pagamenti)
  try {
    showLoading('Riparazione CONTI…');
    if (typeof ensureContiSheet === 'function') await ensureContiSheet();
    const cr   = await dbGet('CONTI!A2:F9999');
    const rows = cr.values || [];
    const upd  = [];

    for (let i = 0; i < rows.length; i++) {
      const row      = rows[i];
      const bid      = String(row[0] || '').trim();
      const contoStr = String(row[4] || '').trim();
      const rowNum   = i + 2;

      // Registra mapping id-conto → bookingId (serve per PAGAMENTI)
      if (contoStr) {
        try {
          const c = JSON.parse(contoStr);
          if (c.id) contoIdToBlip[c.id] = String(c.bookingId || bid);
        } catch(e) {}
      }

      if (!bid || bid.startsWith('PRE-')) continue;   // già OK

      if (contoStr) {
        try {
          const c      = JSON.parse(contoStr);
          const blipId = blipFromContoJSON(c);
          if (blipId) {
            c.bookingId = blipId;
            if (c.id) contoIdToBlip[c.id] = blipId;
            upd.push({ range: 'CONTI!A' + rowNum, values: [[blipId]] });
            upd.push({ range: 'CONTI!E' + rowNum, values: [[JSON.stringify(c)]] });
            // aggiorna cache in-memory billing
            if (typeof _contiDatiCache !== 'undefined') {
              const dati = _contiDatiCache[bid];
              if (dati) { _contiDatiCache[blipId] = { ...dati, contoEmesso: c }; }
            }
            syncLog('✓ CONTI: ' + bid + ' → ' + blipId + ' (' + (c.nome||'?') + ')', 'ok');
            totFixed++;
          } else {
            syncLog('⚠ CONTI: ' + bid + ' (' + (c.nome||'?') + ' cam' + (c.camera||'?') + ') — match non trovato', 'wrn');
            totWarn++;
          }
        } catch(e) { syncLog('⚠ CONTI row ' + rowNum + ': ' + e.message, 'wrn'); totWarn++; }
      } else {
        // Riga con solo EXTRA/OVERRIDE: il bid è solo la chiave di cache, non critico
        syncLog('⚠ CONTI: ' + bid + ' ha solo extra/override, nessun conto emesso — skip', 'inf');
      }
    }

    if (upd.length > 0) {
      await apiFetch(
        'https://sheets.googleapis.com/v4/spreadsheets/' + DATABASE_SHEET_ID + '/values:batchUpdate',
        { method:'POST', headers:{'Content-Type':'application/json'},
          body: JSON.stringify({ valueInputOption:'RAW', data: upd }) }
      );
      syncLog('✓ CONTI: ' + (upd.length/2|0) + ' righe scritte', 'ok');
    }
  } catch(e) { syncLog('⚠ Fix CONTI: ' + e.message, 'wrn'); totWarn++; }

  // ── STEP 3: PAGAMENTI — migra BOOKING_ID numerici → BLIP_ID ──
  try {
    showLoading('Riparazione PAGAMENTI…');
    if (typeof ensurePagamentiSheet === 'function') await ensurePagamentiSheet();
    const pr   = await dbGet('PAGAMENTI!A2:K9999');
    const rows = pr.values || [];
    const upd  = [];

    for (let i = 0; i < rows.length; i++) {
      const row       = rows[i];
      const pagId     = String(row[0] || '').trim();
      const contoId   = String(row[1] || '').trim();
      const bookingId = String(row[2] || '').trim();
      const rowNum    = i + 2;

      if (!pagId || bookingId.startsWith('PRE-')) continue;

      // Cerca BLIP_ID tramite contoId (bridge affidabile)
      let blipId = contoIdToBlip[contoId] || null;

      // Fallback: cerca nei conti in cache per booking id numerico
      if (!blipId && typeof _contiDatiCache !== 'undefined') {
        const entry = Object.values(_contiDatiCache).find(d =>
          d.contoEmesso && String(d.contoEmesso.bookingId) === bookingId
        );
        if (entry?.contoEmesso?.bookingId?.startsWith?.('PRE-')) {
          blipId = entry.contoEmesso.bookingId;
        }
      }

      if (blipId && blipId.startsWith('PRE-')) {
        upd.push({ range: 'PAGAMENTI!C' + rowNum, values: [[blipId]] });
        if (typeof _pagamentiCache !== 'undefined' && _pagamentiCache) {
          const p = _pagamentiCache.find(p => p.id === pagId);
          if (p) p.bookingId = blipId;
        }
        syncLog('✓ PAGAMENTI: ' + pagId + ' → ' + blipId, 'ok');
        totFixed++;
      } else {
        syncLog('⚠ PAGAMENTI: ' + pagId + ' (bid=' + bookingId + ') — BLIP_ID non trovato', 'wrn');
        totWarn++;
      }
    }

    if (upd.length > 0) {
      await apiFetch(
        'https://sheets.googleapis.com/v4/spreadsheets/' + DATABASE_SHEET_ID + '/values:batchUpdate',
        { method:'POST', headers:{'Content-Type':'application/json'},
          body: JSON.stringify({ valueInputOption:'RAW', data: upd }) }
      );
      syncLog('✓ PAGAMENTI: ' + upd.length + ' righe scritte', 'ok');
    }
  } catch(e) { syncLog('⚠ Fix PAGAMENTI: ' + e.message, 'wrn'); totWarn++; }

  // ── STEP 4: CHECK-IN — re-linka orfani (puntano a prenotazioni in CESTINO) ──
  try {
    showLoading('Riparazione CHECK-IN…');
    const activeIds = new Set(bookings.map(b => b.dbId).filter(Boolean));
    const ciR  = await dbGet('CHECK-IN!A2:H9999');
    const rows = ciR.values || [];
    const upd  = [];

    for (let i = 0; i < rows.length; i++) {
      const row    = rows[i];
      const ciId   = String(row[0] || '').trim();
      const preId  = String(row[1] || '').trim();
      const camera = String(row[2] || '').trim();
      const data   = String(row[3] || '').trim();   // YYYY-MM-DD
      const rowNum = i + 2;

      if (!ciId || !preId || activeIds.has(preId)) continue;

      // Cerca booking attivo che copre camera+data
      let dt;
      try { dt = new Date(data + 'T12:00:00'); } catch(e) { continue; }
      const match = bookings.find(b => {
        const bCam = (b.cameraName || roomName(b.r) || '').trim();
        return bCam === camera && b.dbId && b.s && b.e && dt >= b.s && dt < b.e;
      });

      if (match?.dbId) {
        upd.push({ range: 'CHECK-IN!B' + rowNum, values: [[match.dbId]] });
        syncLog('✓ CI: ' + ciId + ' cam' + camera + ' ' + data + ' → ' + match.dbId + ' (' + match.n + ')', 'ok');
        totFixed++;
      } else {
        syncLog('⚠ CI: ' + ciId + ' cam' + camera + ' ' + data + ' — booking attivo non trovato', 'wrn');
        totWarn++;
      }
    }

    if (upd.length > 0) {
      await apiFetch(
        'https://sheets.googleapis.com/v4/spreadsheets/' + DATABASE_SHEET_ID + '/values:batchUpdate',
        { method:'POST', headers:{'Content-Type':'application/json'},
          body: JSON.stringify({ valueInputOption:'RAW', data: upd }) }
      );
      syncLog('✓ CHECK-IN: ' + upd.length + ' righe aggiornate', 'ok');
    }
  } catch(e) { syncLog('⚠ Fix CHECK-IN: ' + e.message, 'wrn'); totWarn++; }

  // ── STEP 5: CONTI — migra anche le righe con solo extra/override ──
  // (bid numerico senza conto emesso — usa matching per booking in memoria)
  try {
    showLoading('Riparazione chiavi CONTI extra/override…');
    const cr   = await dbGet('CONTI!A2:D9999');
    const rows = cr.values || [];
    const upd  = [];

    for (let i = 0; i < rows.length; i++) {
      const row    = rows[i];
      const bid    = String(row[0] || '').trim();
      const rowNum = i + 2;
      if (!bid || bid.startsWith('PRE-')) continue;
      // Cerca nel cache in-memory il bid numerico e trova il BLIP_ID dall'indice
      // Questi record non hanno conto emesso quindi non possiamo fare match per nome+camera
      // Segnaliamo solo come avviso
      syncLog('⚠ CONTI extra-only: ' + bid + ' — rimuovi manualmente o re-inserisci il conto', 'wrn');
      totWarn++;
    }
  } catch(e) {}

  hideLoading();
  const msg = '🔧 Riparazione: ' + totFixed + ' corretti, ' + totWarn + ' avvisi';
  syncLog(msg, totWarn > 0 ? 'wrn' : 'ok');
  showToast(msg, totWarn > 0 ? 'warning' : 'success');
  // Invalida cache e aggiorna il render
  if (typeof render === 'function') render();
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
const MAX_LOG   = 500; // aumentato per avere storico completo nel file di log

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

// ─────────────────────────────────────────────────────────────────
// ESPORTA LOG SESSIONE — genera un file .txt da inviare al supporto
// Contiene: versioni, stato sistema, log completo + prompt di
// continuazione per iniziare una nuova chat AI già contestualizzata.
// ─────────────────────────────────────────────────────────────────
function esportaLogSessione() {
  const now    = new Date();
  const ts     = now.toISOString().replace('T',' ').slice(0,19);
  const tsFile = now.toISOString().slice(0,16).replace(/[:T]/g,'-');

  // ── Raccoglie versioni ─────────────────────────────────────────
  const vers = [
    typeof BLIP_BUILD        !== 'undefined' ? 'BLIP_BUILD='       + BLIP_BUILD        : '',
    typeof BLIP_VER_CORE     !== 'undefined' ? 'BLIP_VER_CORE='    + BLIP_VER_CORE     : '',
    typeof BLIP_VER_SYNC     !== 'undefined' ? 'BLIP_VER_SYNC='    + BLIP_VER_SYNC     : '',
    typeof BLIP_VER_GANTT    !== 'undefined' ? 'BLIP_VER_GANTT='   + BLIP_VER_GANTT    : '',
    typeof BLIP_VER_CHECKIN  !== 'undefined' ? 'BLIP_VER_CHECKIN=' + BLIP_VER_CHECKIN  : '',
    typeof BLIP_VER_BILLING  !== 'undefined' ? 'BLIP_VER_BILLING=' + BLIP_VER_BILLING  : '',
    typeof BLIP_VER_CLIENTI  !== 'undefined' ? 'BLIP_VER_CLIENTI=' + BLIP_VER_CLIENTI  : '',
    typeof BLIP_VER_BRIDGE   !== 'undefined' ? 'BLIP_VER_BRIDGE='  + BLIP_VER_BRIDGE   : '',
  ].filter(Boolean).join('  ');

  // ── Raccoglie stato sistema ────────────────────────────────────
  const stato = [
    'Prenotazioni in memoria : ' + (typeof bookings !== 'undefined' ? bookings.length : '?'),
    'DATABASE_SHEET_ID       : ' + (typeof DATABASE_SHEET_ID !== 'undefined' && DATABASE_SHEET_ID ? DATABASE_SHEET_ID.slice(0,12) + '…' : 'non configurato'),
    'ciData.keys             : ' + (typeof ciData !== 'undefined' ? Object.keys(ciData).length : '?'),
    '_pagamentiCache         : ' + (typeof _pagamentiCache !== 'undefined' && _pagamentiCache ? _pagamentiCache.length + ' record' : 'null'),
    '_cestinoBlacklist       : ' + (typeof _cestinoBlacklist !== 'undefined' && _cestinoBlacklist ? _cestinoBlacklist.size + ' ID' : 'null'),
    '_row46BlipIds           : ' + (typeof _row46BlipIds !== 'undefined' ? _row46BlipIds.size + ' ID' : '?'),
    '_tbTokens               : ' + (typeof _tbTokens !== 'undefined' ? _tbTokens : '?'),
    'annualSheets            : ' + (typeof annualSheets !== 'undefined' ? annualSheets.map(s => s.label||s.year).join(', ') : '?'),
  ].join('\n');

  // ── Log entries (dal più recente) ─────────────────────────────
  const logTxt = [..._logEntries].map(e =>
    `${e.time}  [${(e.type||'inf').toUpperCase().padEnd(3)}]  ${e.msg}`
  ).join('\n');

  // ── Prompt di continuazione ───────────────────────────────────
  const prompt = `
═══════════════════════════════════════════════════════════════
PROMPT DI CONTINUAZIONE — incolla questo testo in una nuova chat
═══════════════════════════════════════════════════════════════

Sto sviluppando Blip, un'app di gestione prenotazioni hotel in
vanilla JavaScript con Google Sheets come DB remoto (GitHub Pages).

Repository: https://github.com/davidepetix-blip/Hotel-prenotazioni
App:        https://davidepetix-blip.github.io/Hotel-prenotazioni/

Versioni attuali:
${vers}

Architettura file (ordine caricamento):
core.js → sync.js → clienti.js → gantt.js → alloggiati-data.js → checkin.js → billing.js → bridge.js

File principali e responsabilità:
- core.js    : costanti ROOMS, helpers puri (date, colori, bed parsing)
- sync.js    : OAuth, apiFetch + token bucket, Sheets API, DB CRUD, bgSync,
               CESTINO blacklist, _row46BlipIds guard, readJSONAnnuale, loadFromSheets
- gantt.js   : render Gantt, modal prenotazione, drawer (usa bridgeSalva/bridgeCancella)
- billing.js : conti, pagamenti, PDF — _ck(bid) normalizza BLIP_ID, calcolo tariffe
- checkin.js : check-in, Alloggiati Web export, OCR Gemini
- clienti.js : anagrafica clienti
- bridge.js  : scrittura/cancellazione sul foglio grafico via GET Apps Script Web App
               sostituisce writeBookingToSheet, clearBookingFromSheet, segnalaModificaAdAppsScript

Apps Script (blip-appscript.gs):
- doGet action=scrivi|cancella → scriviPrenotazioneSuFoglio() → colora celle + riga 46 + JSON_ANNUALE
- doGet default → rigenera JSON_ANNUALE
- onEdit → processSingleColumnBookings + aggiornaJSONAnnuale
- rigenera5min → trigger time-based ogni 5 min

Decisioni architetturali chiave:
- bookingId è sempre stringa BLIP_ID (mai parseInt)
- stato conto calcolato da getStatoContoCalcolato(), mai persistito
- apiFetch ha token bucket (45 token, 900ms/token) + retry 429
- syncWithDatabase: MAX_ARCHIVE_PER_SYNC=20 guard + _row46BlipIds protezione
- bgSync cooldown 3min dopo ogni loadFromSheets completo
- CESTINO blacklist caricata ad ogni sync (TTL 10min)
- _currentYearSheetEntry() usato ovunque per evitare HTTP 400 su foglio anno passato
- bridge.js usa GET verso Apps Script (CORS-compatibile, no-cors fallback)
- JSON_ANNUALE esiste solo sul foglio anno corrente — mai usare .find(e=>e.sheetId) raw

Problema / richiesta attuale:
[DESCRIVI QUI IL PROBLEMA O LA FUNZIONALITÀ DA IMPLEMENTARE]
════════════════════════════════════════════════════════════════`.trim();

  // ── Assembla il file ──────────────────────────────────────────
  const contenuto = [
    '═══════════════════════════════════════════════════════════════',
    'BLIP — LOG DI SESSIONE',
    `Data/ora : ${ts}`,
    `UserAgent: ${navigator.userAgent.slice(0,80)}`,
    '═══════════════════════════════════════════════════════════════',
    '',
    '── VERSIONI ──',
    vers,
    '',
    '── STATO SISTEMA ──',
    stato,
    '',
    '── LOG EVENTI (' + _logEntries.length + ' voci, più recenti prima) ──',
    logTxt || '(nessun evento registrato)',
    '',
    prompt,
  ].join('\n');

  // ── Download ──────────────────────────────────────────────────
  const blob = new Blob([contenuto], { type: 'text/plain;charset=utf-8' });
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement('a');
  a.href     = url;
  a.download = `blip-log-${tsFile}.txt`;
  a.click();
  URL.revokeObjectURL(url);
  syncLog('📥 Log sessione esportato: blip-log-' + tsFile + '.txt', 'ok');
}

// ═══════════════════════════════════════════════════════════════════
// BGSYNC — sincronizzazione periodica in background
// ═══════════════════════════════════════════════════════════════════

let _bgSyncTimer   = null;
let _bgSyncRunning = false;
// Timestamp dell'ultima sync completa (loadFromSheets full path).
// bgSync aspetta almeno BGSYNC_COOLDOWN_MS dopo una sync completa
// per evitare di scatenare un nuovo burst di chiamate API.
let _lastFullSyncTs = 0;
const BGSYNC_COOLDOWN_MS = 6 * 60 * 1000; // 6 minuti (anti-429: sync completa può durare 60s+ più scritture riga 46)
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

// Legge JSON_ANNUALE!O2:O13 (12 celle — 1 per mese) e confronta con i fingerprint
// memorizzati. Costo API: 1 range su 1 foglio — il minimo possibile.
// Restituisce { changed: bool, changedMonths: string[] } dove changedMonths
// sono i nomi dei mesi (es. ["Luglio 2026"]) dove le prenotazioni sono cambiate.
async function checkFingerprintsChanged(sheetId) {
  if (!sheetId) return { changed: true, changedMonths: [] };

  try {
    const range = "JSON_ANNUALE!O2:O13";
    const r = await apiFetch(
      'https://sheets.googleapis.com/v4/spreadsheets/' + sheetId +
      '/values/' + range + '?valueRenderOption=FORMATTED_VALUE'
    );
    if (!r.ok) return { changed: true, changedMonths: [] };

    const data = await r.json();
    const fpRows = data.values || [];
    const MESI_NOMI = ['Gennaio','Febbraio','Marzo','Aprile','Maggio','Giugno',
                       'Luglio','Agosto','Settembre','Ottobre','Novembre','Dicembre'];
    const anno = new Date().getFullYear();
    const changedMonths = [];

    fpRows.forEach((row, i) => {
      const sName  = (MESI_NOMI[i] || '') + ' ' + anno;
      const fpNew  = String(row?.[0] || '').trim();
      const fpOld  = _sheetFingerprints[sName] || '';
      if (!fpNew) return; // formula non ancora installata → ignora
      if (fpNew !== fpOld) {
        changedMonths.push(sName);
        _sheetFingerprints[sName] = fpNew;
      }
    });

    // Se non ci sono fingerprint in memoria (primo avvio, no formule installate)
    // → considera tutto cambiato per sicurezza
    const hasSomeFingerprint = Object.keys(_sheetFingerprints).length > 0;
    if (!hasSomeFingerprint) return { changed: true, changedMonths: [] };

    return { changed: changedMonths.length > 0, changedMonths };
  } catch(e) {
    console.warn('[fingerprint] errore:', e.message);
    return { changed: true, changedMonths: [] };
  }
}

async function bgSync() {
  if (!accessToken || _bgSyncRunning) return;
  // Non partire se una sync completa è avvenuta da meno di BGSYNC_COOLDOWN_MS:
  // evita burst API subito dopo il caricamento iniziale o un forceSync
  if (Date.now() - _lastFullSyncTs < BGSYNC_COOLDOWN_MS) {
    syncLog('⟳ bgSync rimandato — sync recente, attendo cooldown', 'syn');
    return;
  }
  _bgSyncRunning = true;
  setSyncPulsing(true);
  syncLog('⟳ bgSync avviato', 'syn');
  try {
    if (!DATABASE_SHEET_ID) DATABASE_SHEET_ID = loadDbSheetId();

    // ══════════════════════════════════════════════════════════════
    // bgSync — strategia a 3 livelli per minimizzare le chiamate API
    //
    // LIVELLO 1 — fingerprint (1 chiamata leggera, ~100 byte):
    //   Legge JSON_ANNUALE!O2:O13 (12 valori hash, 1 per mese).
    //   Se TUTTI invariati → salta la rilettura del foglio grafico.
    //
    // LIVELLO 2 — DB check (1 chiamata readDatabase):
    //   Solo se il DB potrebbe essere cambiato (scritture recenti o
    //   fingerprint cambiato). Se il DB è identico → nessun render.
    //
    // LIVELLO 3 — render:
    //   Solo se DB effettivamente cambiato.
    //
    // Caso ottimale (nessuna modifica): 1 sola chiamata API totale.
    // ══════════════════════════════════════════════════════════════

    const sheetEntry = _currentYearSheetEntry();
    const fpResult = sheetEntry
      ? await checkFingerprintsChanged(sheetEntry.sheetId)
      : { changed: true, changedMonths: [] };

    // Ci sono scritture dell'app negli ultimi 5 minuti?
    const RECENTE_APP_MS = 5 * 60 * 1000;
    const recenteApp = bookings.filter(b =>
      b.fonte === 'app' && b.ts && (Date.now() - new Date(b.ts).getTime()) < RECENTE_APP_MS
    );
    const hasPendingWrites = recenteApp.length > 0;

    if (!fpResult.changed && !hasPendingWrites) {
      // ── LIVELLO 1: nessuna modifica rilevata, 0 chiamate aggiuntive ──
      syncLog('⟳ bgSync: invariato (fingerprint ok) — skip', 'syn');
      await loadRoomStates();
      return;
    }

    if (fpResult.changedMonths?.length > 0) {
      syncLog('⟳ bgSync: modifiche in ' + fpResult.changedMonths.join(', '), 'syn');
    } else if (hasPendingWrites) {
      syncLog('⟳ bgSync: scritture recenti — verifico DB', 'syn');
    }

    // ── LIVELLO 2: rileggi DB ──────────────────────────────────────
    const dbFresh = await readDatabase();
    const active  = dbFresh.filter(b => !b.deleted && !isDeletedLocally(b.dbId));

    const prevCount = bookings.length;

    // GUARD: se il DB ha molto meno prenotazioni di quelle in memoria, salta il render
    if (active.length < prevCount * 0.7 && prevCount > 20) {
      syncLog(`⚠ bgSync: DB ha ${active.length} vs ${prevCount} in memoria — skip render`, 'wrn');
      await loadRoomStates();
      return;
    }

    // Controlla se il DB è effettivamente cambiato rispetto alla memoria
    const rimossi = bookings.filter(b => b.dbId && !active.find(d => d.dbId === b.dbId)).length;
    let dbChanged = active.length !== prevCount || rimossi > 0;
    if (!dbChanged) {
      const localMap = new Map(bookings.map(b => [b.dbId, b.ts]).filter(([k]) => k));
      for (const db of active) {
        const localTs = localMap.get(db.dbId);
        if (!localTs || (db.ts && db.ts > localTs)) { dbChanged = true; break; }
      }
    }
    syncLog(`DB: ${active.length} prenotazioni attive, rimossi ${rimossi}`, 'db');

    // ── LIVELLO 3: render solo se DB cambiato ─────────────────────
    if (dbChanged) {
      let merged = mergeMultiMonthBookings(active);
      // Preserva prenotazioni locali recenti non ancora nel DB
      // (latenza scrittura API + Apps Script 20-30s)
      if (hasPendingWrites) {
        recenteApp.forEach(localB => {
          const inDb = merged.find(db =>
            db.dbId === localB.dbId ||
            (db.n === localB.n && db.r === localB.r && Math.abs(db.s - localB.s) < 86400000)
          );
          if (!inDb) {
            syncLog(`🛡 bgSync: preservata prenotazione locale recente: ${localB.n} (${localB.dbId||'no dbId'})`, 'wrn');
            merged.push(localB);
          }
        });
      }
      bookings = merged;
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
        // Carica la blacklist CESTINO in background — senza await per non ritardare il render.
        // In fast path syncWithDatabase non viene chiamata, quindi senza questo
        // _cestinoBlacklist resterebbe null per tutta la sessione.
        loadCestinoBlacklist().catch(() => {});
        // Aspetta 90s prima del primo bgSync: preloadContoDati e CI data
        // caricano in questo intervallo — partire prima causa burst 429
        _lastFullSyncTs = Date.now() - BGSYNC_COOLDOWN_MS + 90 * 1000;
        setTimeout(bgSync, 90 * 1000);
        startBgSync();
        return;
      }
    }

    // ── DB-FIRST FULL PATH ──────────────────────────────────────────
    // STEP 1: legge il DB locale → render immediato (~1s)
    // STEP 2: in background legge JSON_ANNUALE + sincronizza → re-render silenzioso
    // Se non c'è DATABASE_SHEET_ID, cade direttamente al path JSON_ANNUALE.
    const _t0 = Date.now();
    _row46BlipIds    = new Set();
    _row46BookingMap = {};

    if (DATABASE_SHEET_ID) {
      // ── STEP 1: DB → render veloce ──────────────────────────────
      showLoading('Lettura database…');
      await ensureDbHeaders();
      const _t1db = Date.now();
      const dbRowsFast = await readDatabase();
      const dbActiveFast = dbRowsFast.filter(b => !b.deleted);
      syncLog(`⏱ DB fast-read: ${dbActiveFast.length} pren. in ${Date.now()-_t1db}ms`, 'syn');

      bookings = mergeMultiMonthBookings(dbActiveFast);
      await loadRoomStates();
      hideLoading();
      render();
      showToast(`✓ ${bookings.length} prenotazioni (DB)`, 'success');
      syncLog(`✓ ${bookings.length} prenotazioni caricate dal DB`, 'ok');

      // Imposta una cache veloce con i dati DB per il fast-path al prossimo avvio
      saveDbCache(dbActiveFast);

      // Mostra indicatore sync in corso sul syncDot
      setSyncPulsing(true);
      syncLog('📖 Aggiornamento da JSON_ANNUALE in background…', 'syn');

      // ── STEP 2: JSON_ANNUALE + sync in background ───────────────
      // Non blocca l'UI — eseguito dopo che l'utente vede già il calendario
      ;(async () => {
        try {
          const _t2 = Date.now();
          const sheetEntry = _currentYearSheetEntry();
          let sheetBookings = [];

          if (sheetEntry) {
            try {
              sheetBookings = await readJSONAnnuale(sheetEntry.sheetId);
              syncLog(`⏱ JSON_ANNUALE: ${sheetBookings.length} pren. in ${Date.now()-_t2}ms`, 'syn');
            } catch(err) {
              syncLog('⚠ JSON_ANNUALE fallback fogli mensili: ' + err.message, 'wrn');
              for (let i = 0; i < annualSheets.length; i++) {
                const entry = annualSheets[i];
                if (!entry.sheetId) continue;
                try { sheetBookings.push(...(await readAnnualSheet(entry))); } catch(e2) {}
                if (i < annualSheets.length-1) await new Promise(r => setTimeout(r, 300));
              }
            }
          }

          // fromFallback=true se non siamo riusciti a leggere JSON_ANNUALE
          const _fromFallback = sheetBookings.length > 0 && !sheetBookings.some(b => b.fromJSONAnnuale);
          const _t3 = Date.now();
          syncLog('Sincronizzazione DB…', 'syn');
          sheetBookings = await syncWithDatabase(sheetBookings, forcing, _fromFallback);
          syncLog(`⏱ Sync DB: ${Date.now()-_t3}ms`, 'syn');

          await loadRoomStates();
          const updated = mergeMultiMonthBookings(sheetBookings);

          // Re-render solo se i dati sono cambiati rispetto a quelli già mostrati
          const prevIds = new Set(bookings.map(b => b.dbId).filter(Boolean));
          const newIds  = new Set(updated.map(b => b.dbId).filter(Boolean));
          const hasChanges = updated.length !== bookings.length
            || updated.some(b => b.dbId && !prevIds.has(b.dbId))
            || bookings.some(b => b.dbId && !newIds.has(b.dbId));

          if (hasChanges) {
            bookings = updated;
            render();
            const diff = updated.length - dbActiveFast.length;
            if (diff !== 0) {
              showBgToast(diff > 0 ? `↻ +${diff} nuove da foglio` : `↻ ${Math.abs(diff)} rimosse da foglio`);
            }
            syncLog(`✓ ${bookings.length} prenotazioni dopo sync foglio`, 'ok');
          } else {
            syncLog('✓ Nessuna modifica rilevata dal foglio', 'ok');
          }

          // Usa il flag calcolato PRIMA del sync — dopo syncWithDatabase
          // i booking del DB (senza fromJSONAnnuale) diluiscono il check.
          const fromJSON = !_fromFallback;
          syncLog(fromJSON ? 'Fonte: JSON_ANNUALE' : 'Fonte: 12 fogli mensili (fallback)', 'inf');
          saveDbCache(DATABASE_SHEET_ID ? sheetBookings : bookings);

        } catch(e2) {
          syncLog('⚠ Aggiornamento foglio: ' + e2.message, 'wrn');
        } finally {
          hideLoading();
          setSyncPulsing(false);
          _lastFullSyncTs = Date.now();
          startBgSync();
        }
      })();

      return; // STEP 1 completato — STEP 2 gira in background
    }

    // ── PATH senza DB: solo JSON_ANNUALE (come prima) ────────────
    const sheetEntry = _currentYearSheetEntry();
    let sheetBookings = [];
    if (sheetEntry) {
      syncLog('📖 Lettura JSON_ANNUALE da foglio…', 'syn');
      showLoading('Lettura JSON_ANNUALE…');
      try {
        sheetBookings = await readJSONAnnuale(sheetEntry.sheetId);
        syncLog(`⏱ JSON_ANNUALE: ${sheetBookings.length} pren. in ${Date.now()-_t0}ms`, 'syn');
      } catch(err) {
        for (let i = 0; i < annualSheets.length; i++) {
          const entry = annualSheets[i];
          if (!entry.sheetId) continue;
          showLoading(`Lettura ${entry.label} (${i+1}/${annualSheets.length})…`);
          try { sheetBookings.push(...(await readAnnualSheet(entry))); } catch(e2) {}
          if (i < annualSheets.length-1) await new Promise(r => setTimeout(r, 300));
        }
      }
    }
    await loadRoomStates();
    bookings = mergeMultiMonthBookings(sheetBookings);
    hideLoading();
    render();
    showToast(`✓ ${bookings.length} prenotazioni`, 'success');
    syncLog(`✓ ${bookings.length} prenotazioni caricate`, 'ok');
    saveDbCache(bookings);
    _lastFullSyncTs = Date.now();
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

// getSheetIdMap, writeFragment, writeBookingToSheet, clearFragment,
// clearBookingFromSheet, segnalaModificaAdAppsScript rimossi in build .4.3
// → sostituiti da bridge.js (chiamata GET all'Apps Script Web App)

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
