// ═══════════════════════════════════════════════════════════════════
// core.js — Stato globale e Utilities
// Blip Hotel Management — build 18.11.51
// ═══════════════════════════════════════════════════════════════════

const BLIP_VER_CORE = '6';

function dbg(msg, isErr) {
  console.log(msg);
  const box = document.getElementById('dbgLog');
  if (box) {
    box.style.display = 'block';
    const l = document.createElement('div');
    l.style.color = isErr ? '#ff4d4d' : '#00ff00';
    l.textContent = new Date().toLocaleTimeString() + ' ' + msg;
    box.appendChild(l);
    box.scrollTop = box.scrollHeight;
  }
  if (isErr) {
    const e = document.getElementById('loginErr');
    if (e) e.textContent = msg;
  }
}

window.onerror = (m, s, l) => { dbg(`❌ Errore riga ${l}: ${m}`, true); return false; };

const SCRIPT_ID = 'AKfycbzL7QO5o3Xm1Ld60E_3p5I39_57Lw5v0SIn8-sVj8w51VpB16A-iMvPIdfA_FjV2S8'; 
const BASE_URL  = `https://script.google.com/macros/s/${SCRIPT_ID}/exec`;

const DB_COLS = {
  PRENOTAZIONI: { ID:0, CAMERA:1, NOME:2, DAL:3, AL:4, DISP:5, NOTE:6, COLORE:7, ANNO:8, FONTE:9, TS:10, DELETED:11, CLIENTE_ID:12 },
  CONTI:        { BOOKING_ID:0, EXTRA_JSON:1, OVERRIDE_JSON:2, APPART_MODE:3, CONTO_EMESSO_JSON:4, TS:5 },
  PAGAMENTI:    { PAG_ID:0, CONTO_ID:1, BOOKING_ID:2, DATA:3, IMPORTO:4, TIPO:5, METODO:6, RIF:7, CON_DOC:8, NOTE:9, TS:10 },
  CAMERE:       { CAMERA:0, MAX:1, LETTI:2, PULIZIA:3, CONFIG:4, NOTE:5, TS:6 }
};

const DB_SHEETS = {
  PRENOTAZIONI: 'PRENOTAZIONI',
  CONTI: 'CONTI',
  PAGAMENTI: 'PAGAMENTI'
};

let bookings = [];
let gUserToken = null;

function nowISO() { return new Date().toISOString(); }

function diffDays(a, b) {
  const d1 = new Date(a); d1.setHours(12,0,0,0);
  const d2 = new Date(b); d2.setHours(12,0,0,0);
  return Math.round((d2 - d1) / (1000*60*60*24));
}

function genPagamentoId() {
  return 'PAG-' + new Date().getFullYear() + '-' + Math.random().toString(36).substring(2, 9).toUpperCase();
}

// Network con Token Bucket
let _tokens = 45;
setInterval(() => { if (_tokens < 45) _tokens++; }, 1500);

async function apiFetch(action, params = {}, retry = 0) {
  if (!gUserToken && action !== 'login') {
    throw new Error("Effettuare il login prima di procedere");
  }

  if (_tokens <= 0) {
    await new Promise(r => setTimeout(r, 2000));
    return apiFetch(action, params, retry);
  }

  _tokens--;
  const url = `${BASE_URL}?action=${action}${gUserToken ? '&token='+gUserToken : ''}`;

  try {
    const r = await fetch(url, { method: 'POST', body: JSON.stringify(params) });
    const d = await r.json();
    if (d.error) throw new Error(d.error);
    return d.result;
  } catch (e) {
    if (retry < 2) return apiFetch(action, params, retry + 1);
    throw e;
  }
}

async function fetchSheet(sheet) {
  return await apiFetch('readSheet', { sheet });
}
  PAGAMENTI: 'PAGAMENTI',
  CLIENTI: 'CLIENTI',
  IMPOSTAZIONI: 'IMPOSTAZIONI',
  CAMERE: 'CAMERE'
};

// Configurazione Camere (Ridotta per brevità, caricata dinamicamente in sync.js)
let ROOMS = []; 

// ═══════════════════════════════════════════════════════════════════
// STATO GLOBALE
// ═══════════════════════════════════════════════════════════════════
let bookings = [];
let gUserToken = null;
let curM = new Date().getMonth();
let curY = new Date().getFullYear();

// ═══════════════════════════════════════════════════════════════════
// UTILITIES — DATE & STRINGHE
// ═══════════════════════════════════════════════════════════════════

function isoToDate(iso) { return new Date(iso); }
function dateToIso(d) { return d.toISOString(); }
function nowISO() { return new Date().toISOString(); }

function parseDateIT(str) {
  if (!str) return null;
  const parts = str.split('/');
  if (parts.length!==3) return new Date(str); // fall back
  return new Date(parts[2], parts[1]-1, parts[0], 12, 0, 0);
}

function formatDateIT(d) {
  if (!d) return '';
  const day = String(d.getDate()).padStart(2,'0');
  const month = String(d.getMonth()+1).padStart(2,'0');
  return `${day}/${month}/${d.getFullYear()}`;
}

function diffDays(a, b) {
  const d1 = new Date(a); d1.setHours(12,0,0,0);
  const d2 = new Date(b); d2.setHours(12,0,0,0);
  return Math.round((d2 - d1) / (1000*60*60*24));
}

function genBookingId() {
  const prefix = 'PRE-' + new Date().getFullYear();
  const random = Math.random().toString(36).substring(2, 8).toUpperCase();
  const time   = Date.now().toString(36).toUpperCase().slice(-4);
  return `${prefix}-${random}-${time}`;
}

function genPagamentoId() {
  return 'PAG-' + new Date().getFullYear() + '-' + Math.random().toString(36).substring(2, 8).toUpperCase();
}

// ═══════════════════════════════════════════════════════════════════
// NETWORK — apiFetch con Token Bucket & Retry
// ═══════════════════════════════════════════════════════════════════

let _tokens = 45;
const _maxTokens = 45;
setInterval(() => { if (_tokens < _maxTokens) _tokens++; }, 1400);

async function apiFetch(action, params = {}, retryCount = 0) {
  if (_tokens <= 0) {
    const wait = 2000 + (retryCount * 2000);
    dbg(`⏳ Rate limit, attendo ${wait}ms...`);
    await new Promise(r => setTimeout(r, wait));
    return apiFetch(action, params, retryCount + 1);
  }

  _tokens--;
  const url = new URL(BASE_URL);
  url.searchParams.set('action', action);
  if (gUserToken) url.searchParams.set('token', gUserToken);

  const options = {
    method: 'POST',
    mode: 'cors',
    body: JSON.stringify(params)
  };

  try {
    const resp = await fetch(url.toString(), options);
    if (resp.status === 429) throw new Error('429');
    const data = await resp.json();
    if (data.error) throw new Error(data.error);
    return data.result;
  } catch (err) {
    if (err.message === '429' && retryCount < 3) {
      const backoff = [5000, 10000, 20000][retryCount];
      dbg(`⚠️ 429 Too Many Requests. Retry ${retryCount+1} in ${backoff}ms`);
      await new Promise(r => setTimeout(r, backoff));
      return apiFetch(action, params, retryCount + 1);
    }
    dbg(`❌ API Error (${action}): ` + err.message, true);
    throw err;
  }
}

// Helper per scaricare un intero foglio
async function fetchSheet(sheetName) {
  return await apiFetch('readSheet', { sheet: sheetName });
}

// ═══════════════════════════════════════════════════════════════════
// HELPERS — UI (toast, loading)
// ═══════════════════════════════════════════════════════════════════
function showToast(msg, dur=3000) {
  let t = document.getElementById('blipToast');
  if(!t) {
    t = document.createElement('div'); t.id='blipToast';
    Object.assign(t.style, {
      position:'fixed', bottom:'24px', left:'50%', transform:'translateX(-50%)',
      background:'#333', color:'#fff', padding:'12px 20px', borderRadius:'8px',
      fontSize:'14px', zIndex:'9999', transition:'opacity 0.3s', opacity:'0'
    });
    document.body.appendChild(t);
  }
  t.textContent = msg; t.style.display='block'; setTimeout(()=>t.style.opacity='1',10);
  setTimeout(()=>{ t.style.opacity='0'; setTimeout(()=>t.style.display='none',300); }, dur);
}
(e) {}
  return JSON.parse(JSON.stringify(DEFAULT_ANNUAL_SHEETS));
}
function saveAnnualSheetsLS(arr) {
  localStorage.setItem('hotelAnnualSheets', JSON.stringify(arr));
}
function loadDbSheetId() {
  return localStorage.getItem('hotelDbSheetId') || '';
}
function saveDbSheetIdLS(id) {
  localStorage.setItem('hotelDbSheetId', id);
  DATABASE_SHEET_ID = id;
}

// Inizializza annualSheets dopo che loadAnnualSheets è definita
annualSheets = loadAnnualSheets();

// ═══════════════════════════════════════════════════════════════════
// ═══════════════════════════════════════════════════════════════════
// HELPERS — generatori ID e formatters
// ═══════════════════════════════════════════════════════════════════

function _randSuffix(n = 4) {
  const chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  return Array.from(crypto.getRandomValues(new Uint8Array(n)))
    .map(b => chars[b % chars.length]).join('');
}
function _tsBase36() {
  return Math.floor(Date.now() / 1000).toString(36).toUpperCase();
}
// PRE-{anno}-{ts6}-{rand4}  es. PRE-2026-MN4H0T-X7K2
function genBookingId(year) {
  return `PRE-${year}-${_tsBase36()}-${_randSuffix(4)}`;
}
// CLI-{anno}-{ts6}-{rand4}
function genClienteId(year) {
  year = year || new Date().getFullYear();
  return `CLI-${year}-${_tsBase36()}-${_randSuffix(4)}`;
}
// CON-{anno}-{ts6}-{rand4}
function genContoId(year) {
  year = year || new Date().getFullYear();
  return `CON-${year}-${_tsBase36()}-${_randSuffix(4)}`;
}
// PAG-{anno}-{ts6}-{rand4}
function genPagamentoId(year) {
  year = year || new Date().getFullYear();
  return `PAG-${year}-${_tsBase36()}-${_randSuffix(4)}`;
}

function nowISO() { return new Date().toISOString(); }

function buildBedString(counts) {
  const parts = [];
  if (counts.m   > 0) parts.push(`${counts.m}m`);
  if (counts.ms  > 0) parts.push(`${counts.ms}m/s`);
  if (counts.s   > 0) parts.push(`${counts.s}s`);
  if (counts.c   > 0) parts.push(`${counts.c}c`);
  if (counts.aff > 0) parts.push(`${counts.aff}aff`);
  return parts.join(' + ') || 'ND';
}

function parseBedString(str) {
  const counts = { m:0, ms:0, s:0, c:0, aff:0 };
  if (!str || str === 'ND') return counts;
  const norm   = str.replace(/\s*\+\s*/g, '+').replace(/\s*,\s*/g, '+').trim();
  const normMS = norm.replace(/\s/g,'').toLowerCase();
  if (/^(\d*)m\/s$/.test(normMS)) { counts.ms = parseInt(normMS) || 1; return counts; }
  const parts = norm.split('+');
  for (const part of parts) {
    const p = part.trim();
    const nms = p.match(/^(\d+)ms$/i); if (nms) { counts.ms  = parseInt(nms[1]); continue; }
    const nm  = p.match(/^(\d+)m$/);   if (nm)  { counts.m   = parseInt(nm[1]);  continue; }
    const ns  = p.match(/^(\d+)s$/);   if (ns)  { counts.s   = parseInt(ns[1]);  continue; }
    const nc  = p.match(/^(\d+)c$/);   if (nc)  { counts.c   = parseInt(nc[1]);  continue; }
    const na  = p.match(/^(\d+)aff$/); if (na)  { counts.aff = parseInt(na[1]);  continue; }
    if (p === 'ms')  { counts.ms=1;  continue; }
    if (p === 'm')   { counts.m=1;   continue; }
    if (p === 's')   { counts.s=1;   continue; }
    if (p === 'c')   { counts.c=1;   continue; }
    if (p === 'aff') { counts.aff=1; continue; }
    if (p.match(/^(\d+)?m\/s$/i)) { counts.ms = parseInt(p) || 1; }
  }
  return counts;
}

// ═══════════════════════════════════════════════════════════════════
// HELPERS — puri (nessun DOM, nessuna chiamata API)
// ═══════════════════════════════════════════════════════════════════
const dim       = (y,m) => new Date(y,m+1,0).getDate();
const fmt       = d => d.toLocaleDateString('it-IT');
const nights    = (a,b) => Math.round((new Date(b.getFullYear(),b.getMonth(),b.getDate()) - new Date(a.getFullYear(),a.getMonth(),a.getDate()))/86400000);
const light     = h => { if(!h||h.length<7) return true; const r=parseInt(h.slice(1,3),16),g=parseInt(h.slice(3,5),16),b=parseInt(h.slice(5,7),16); return (r*299+g*587+b*114)/1000>150; };
// Normalizza qualsiasi colore hex verso il pastello mescolandolo con il bianco.
// Preserva l'HUE (riconoscibilità), uniforma luminosità e saturazione.
// mix=0 → colore originale, mix=1 → bianco puro. Default 0.68 → pastello deciso.
const pastello  = (h, mix=0.68) => { if(!h||h.length<7) return h||'#D9D9D9'; const r=parseInt(h.slice(1,3),16),g=parseInt(h.slice(3,5),16),b=parseInt(h.slice(5,7),16); const nr=Math.round(r+(255-r)*mix),ng=Math.round(g+(255-g)*mix),nb=Math.round(b+(255-b)*mix); return '#'+[nr,ng,nb].map(v=>v.toString(16).padStart(2,'0').toUpperCase()).join(''); };
const overlaps  = (b,s,e,xid) => b.id!==xid && b.s<e && b.e>s;
const roomName  = rid => ROOMS.find(r=>r.id===rid)?.name || rid;
const roomGroup = rid => ROOMS.find(r=>r.id===rid)?.g || '';
const sheetName = (y,m) => `${MONTHS_IT[m]} ${y}`;
const numToCol  = n => { let s=''; while(n>0){const r=(n-1)%26; s=String.fromCharCode(65+r)+s; n=Math.floor((n-1)/26);} return s; };

const adjConflict = b => bookings.filter(o =>
  o.id !== b.id &&
  o.r  === b.r  &&
  o.c  === b.c  &&
  (o.e.getTime() === b.s.getTime() || o.s.getTime() === b.e.getTime()) &&
  !(o.n === b.n && o.d !== b.d)
);
const colorsUsed = (rid,xid) => {
  const ms=new Date(curY,curM,1), me=new Date(curY,curM+1,0);
  return bookings.filter(b=>b.id!==xid&&b.r===rid&&b.s<=me&&b.e>=ms).map(b=>b.c);
};

function hexToSheetsColor(hex) {
  if (!hex || hex.length < 7) return {red:1,green:1,blue:1};
  return {
    red:   parseInt(hex.slice(1,3),16)/255,
    green: parseInt(hex.slice(3,5),16)/255,
    blue:  parseInt(hex.slice(5,7),16)/255,
  };
}
function sheetsColorToHex(c) {
  if (!c) return '#ffffff';
  const r = Math.round((c.red||0)*255).toString(16).padStart(2,'0');
  const g = Math.round((c.green||0)*255).toString(16).padStart(2,'0');
  const b = Math.round((c.blue||0)*255).toString(16).padStart(2,'0');
  return '#'+r+g+b;
}

// ═══════════════════════════════════════════════════════════════════
// HELPERS — UI (toast, loading)
// ═══════════════════════════════════════════════════════════════════
function showToast(msg, type='', ms=3000) {
  const t = document.getElementById('toast');
  t.textContent = msg; t.className = 'toast show ' + type;
  setTimeout(() => { t.className = 'toast'; }, ms);
}
function showLoading(msg='Caricamento…') {
  if (typeof _bgSyncRunning !== 'undefined' && _bgSyncRunning) return;
  document.getElementById('loadMsg').textContent = msg;
  document.getElementById('loadingOverlay').classList.add('show');
}
function hideLoading() {
  if (typeof _bgSyncRunning !== 'undefined' && _bgSyncRunning) return;
  document.getElementById('loadingOverlay').classList.remove('show');
}
