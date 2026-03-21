// ═══════════════════════════════════════════════════════════════════
// sync.js — Auth OAuth, Google Sheets API, DB, Sync Engine
// Blip Hotel Management — build 18.7.x
// Dipende da: core.js (deve essere caricato prima)
// ═══════════════════════════════════════════════════════════════════

// ═══════════════════════════════════════════════════════════════════
// AUTH — OAuth 2.0 redirect flow
// ═══════════════════════════════════════════════════════════════════


const BLIP_VER_SYNC = '1'; // ← incrementa ad ogni modifica

function randomState() {
  return Array.from(crypto.getRandomValues(new Uint8Array(16)))
    .map(b => b.toString(16).padStart(2,'0')).join('');
}

function getRedirectUri() {
  const path = location.pathname.replace(/\/index\.html$/, '/').replace(/([^/])$/, '$1/');
  return location.origin + path;
}

function startLogin() {
  dbg('▶ startLogin');
  document.getElementById('loginErr').textContent = '';
  const state = randomState();
  sessionStorage.setItem('oauth_state', state);
  const params = new URLSearchParams({
    client_id:     CLIENT_ID,
    redirect_uri:  getRedirectUri(),
    response_type: 'token',
    scope:         SCOPES,
    state,
    include_granted_scopes: 'true',
  });
  location.href = 'https://accounts.google.com/o/oauth2/v2/auth?' + params.toString();
}

function handleOAuthRedirect() {
  dbg('▶ handleOAuthRedirect hash='+(location.hash?'si':'no'));
  const hash = location.hash.slice(1);
  if (!hash) return false;
  const params = new URLSearchParams(hash);
  const token  = params.get('access_token');
  const error  = params.get('error');
  const state  = params.get('state');

  history.replaceState(null, '', location.pathname);

  if (error) {
    document.getElementById('loginErr').textContent = 'Accesso negato: ' + error;
    return true;
  }
  if (!token) return false;

  const saved = sessionStorage.getItem('oauth_state');
  sessionStorage.removeItem('oauth_state');
  if (saved && state !== saved) {
    document.getElementById('loginErr').textContent = 'Errore sicurezza. Riprova.';
    return true;
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
  showToast('📡 Rigenerazione in corso…', 'info');
  try {
    await fetch(`${url}?anno=${new Date().getFullYear()}&ts=${Date.now()}`, {method:'GET',mode:'no-cors'});
    showToast('⏳ Attendi 5 secondi poi il calendario si aggiorna…', 'info');
    setTimeout(async () => {
      try {
        annualSheets = loadAnnualSheets();
        await loadFromSheets();
        showToast('✓ Calendario rigenerato', 'success');
      } catch(e2) { showToast('⚠ Ricarica la pagina manualmente', 'warning'); }
    }, 5000);
  } catch(e) {
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
  arr[DB_COLS.ID-1]      = b.dbId || b.id || genBookingId(anno);
  arr[DB_COLS.CAMERA-1]  = b.cameraName || roomName(b.r) || '';
  arr[DB_COLS.NOME-1]    = b.n || '';
  arr[DB_COLS.DAL-1]     = dal;
  arr[DB_COLS.AL-1]      = al;
  arr[DB_COLS.DISP-1]    = b.d || '';
  arr[DB_COLS.NOTE-1]    = b.note || '';
  arr[DB_COLS.COLORE-1]  = b.c || '#D9D9D9';
  arr[DB_COLS.ANNO-1]    = String(anno);
  arr[DB_COLS.FONTE-1]   = fonte;
  arr[DB_COLS.TS-1]      = b.ts || nowISO();
  arr[DB_COLS.DELETED-1] = b.deleted ? 'true' : '';
  arr[12] = b.deletedAt || '';
  arr[13] = b.deleteReason || '';
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
  const data = await dbGet(`${DB_SHEET_NAME}!A${DB_FIRST_ROW}:N9999`);
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
  if (!b.dbRow) return;
  try { await archiviaInCestino([b], reason); } catch(e) { console.warn('[dbDelete]:', e.message); }
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

  const conRiga = lista.filter(b => b.dbRow);
  if (conRiga.length > 0) {
    const data = conRiga.map(b => {
      const row = bookingToDbRow(b, b.fonte || 'app');
      row[DB_COLS.DELETED-1] = 'true'; row[DB_COLS.TS-1] = ts; row[12] = ts; row[13] = reason;
      const lastCol = String.fromCharCode(64 + row.length);
      return { range: `${DB_SHEET_NAME}!A${b.dbRow}:${lastCol}${b.dbRow}`, values: [row] };
    });
    for (let i = 0; i < data.length; i += 1000) {
      const chunk = data.slice(i, i+1000);
      await apiFetch(`https://sheets.googleapis.com/v4/spreadsheets/${id}/values:batchUpdate`, {
        method:'POST', headers:{'Content-Type':'application/json'},
        body: JSON.stringify({ valueInputOption:'RAW', data:chunk })
      });
    }
  }
  console.log(`[CESTINO] ${lista.length} righe archiviate`);
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
      const [hR, jR] = await Promise.all([
        fetch(base+enc(`'${sName}'!B${HEADER_ROW}:AJ${HEADER_ROW}`)+'?valueRenderOption=FORMATTED_VALUE', {headers:{Authorization:'Bearer '+accessToken}}),
        fetch(base+enc(`'${sName}'!B${OUTPUT_ROW}:AJ${OUTPUT_ROW}`)+'?valueRenderOption=FORMATTED_VALUE', {headers:{Authorization:'Bearer '+accessToken}}),
      ]);
      const headers = (await hR.json()).values?.[0] || [];
      const jsonRow = (await jR.json()).values?.[0] || [];

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
            result.push({
              id:nid++, r:room.id, n:b.nome||'—', d:b.disposizione||'',
              c:colorHex, s:new Date(yy,mm-1,dd,12), e:new Date(ye,me-1,de,12),
              note:b.note||'', fromSheet:true, sheetName:sName,
              cameraName:room.name, sheetId,
              dbId:null, dbRow:null, ts:null, fonte:'manuale',
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
const DB_CACHE_TTL_MS = 60 * 60 * 1000;
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
  const camT  = (target.cameraName || roomName(target.r) || '').toLowerCase().trim();
  const nomT  = (target.n || '').toLowerCase().trim();
  const dayT  = Math.round((target.s?.getTime?.() || 0) / DAY_MS);
  const dispT = (target.d || '').trim().toLowerCase();

  let m = list.find(b => {
    if ((b.n||'').toLowerCase().trim() !== nomT) return false;
    if ((b.cameraName||roomName(b.r)||'').toLowerCase().trim() !== camT) return false;
    return Math.round((b.s?.getTime?.() || 0) / DAY_MS) === dayT;
  });
  if (m) return m;

  m = list.find(b => {
    if ((b.n||'').toLowerCase().trim() !== nomT) return false;
    if ((b.cameraName||roomName(b.r)||'').toLowerCase().trim() !== camT) return false;
    if (dispT && (b.d||'').trim().toLowerCase() !== dispT) return false;
    return Math.abs(Math.round((b.s?.getTime?.() || 0) / DAY_MS) - dayT) <= 1;
  });
  return m || null;
}

async function cleanupDeletedFromDb(dbRows) {
  const alreadyDeleted = dbRows.filter(b => b.deleted && b.dbRow);
  if (alreadyDeleted.length === 0) return 0;
  try { await archiviaInCestino(alreadyDeleted, 'Pulizia avvio — era già marcata DELETED'); } catch(e) {}
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

  // FASE 1: Foglio → verità assoluta
  for (const sheet of sheetBookings) {
    if (!sheet.n || sheet.n === '???' || sheet.n.trim() === '') continue;
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
  return result;
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
    const active  = dbFresh.filter(b => !b.deleted);

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
    showToast('🔄 Aggiornamento calendario in corso…', 'info');
    const anno = new Date().getFullYear();
    const url  = `${webAppUrl}?anno=${anno}&ts=${Date.now()}`;
    syncLog(`📡 Chiamata Web App: ${url.slice(0,60)}…`, 'syn');
    await fetch(url, { method:'GET', mode:'no-cors' });
    setTimeout(async () => {
      try {
        annualSheets = loadAnnualSheets();
        await loadFromSheets();
        syncLog('✓ JSON_ANNUALE rigenerato, calendario aggiornato', 'ok');
        showToast('✓ Calendario aggiornato', 'success');
      } catch(e2) { console.warn('[WebApp] reload fallito:', e2.message); }
    }, 6000);
  } catch(e) {
    console.warn('[WebApp] chiamata fallita:', e.message);
    showToast('⚠ Aggiornamento calendario fallito — ricarica manualmente con 🔄', 'warning');
  }
}

async function clearFragment(sName, cameraName, startDay, endDay, sheetIdMap, spreadsheetId) {
  let colIdx = sheetColumnMap[sName]?.[cameraName];
  if (!colIdx) return;
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
