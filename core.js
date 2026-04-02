// ═══════════════════════════════════════════════════════════════════
// core.js — Stato globale, costanti, ROOMS, utilities condivise
// Blip Hotel Management — build 18.7.x
// Caricato PRIMA di tutti gli altri moduli.
// ═══════════════════════════════════════════════════════════════════

// ── Error handler globale per debug mobile ──

const BLIP_VER_CORE = '5'; // ← incrementa ad ogni modifica

function dbg(msg, isErr) {
  console.log(msg);
  const box = document.getElementById('dbgLog');
  if (box) { box.style.display='block'; const l=document.createElement('div'); l.style.color=isErr?'#c0392b':'#555'; l.textContent=new Date().toLocaleTimeString()+' '+msg; box.appendChild(l); box.scrollTop=box.scrollHeight; }
  if (isErr) { const e=document.getElementById('loginErr'); if(e) e.textContent=msg; }
}
window.onerror = function(msg, src, line) { dbg('❌ JS riga '+line+': '+msg, true); return false; };
window.addEventListener('unhandledrejection', function(e) { dbg('❌ Promise: '+(e.reason?.message||String(e.reason)), true); });

// ═══════════════════════════════════════════════════════════════════
// CONFIG
// ═══════════════════════════════════════════════════════════════════
const CLIENT_ID = '13060466249-bk4s31a1vanhnd6j0qhequ3d3ptd2b2g.apps.googleusercontent.com';
const SCOPES    = 'https://www.googleapis.com/auth/spreadsheets';

const DEFAULT_ANNUAL_SHEETS = [
  { year: 2025, sheetId: '', label: '2025' },
  { year: 2026, sheetId: '1XTalxQBUFywBW4DL3JwSJvKkPlOhstAbRMH4DWx-eGI', label: '2026' },
];

let DATABASE_SHEET_ID = '';

const DB_COLS = {
  ID:         1,  // A
  CAMERA:     2,  // B
  NOME:       3,  // C
  DAL:        4,  // D — dd/MM/yyyy
  AL:         5,  // E — dd/MM/yyyy
  DISP:       6,  // F
  NOTE:       7,  // G
  COLORE:     8,  // H
  ANNO:       9,  // I
  FONTE:      10, // J — "app" | "manuale"
  TS:         11, // K — timestamp ISO
  DELETED:    12, // L — "true" se eliminata
  CLIENTE_ID: 13, // M — collegamento anagrafica CLI-xxxx
};
const DB_SHEET_NAME      = 'PRENOTAZIONI';
const CESTINO_SHEET_NAME = 'CESTINO';
const DB_HEADER_ROW      = 1;
const DB_FIRST_ROW       = 2;

const ROOMS_SHEET_NAME = 'CAMERE';
const RCOLS = {
  CAMERA: 1, MAX_OSPITI: 2, LETTI_AMMESSI: 3,
  PULIZIA: 4, CONFIGURAZIONE: 5, NOTE_OPS: 6, TS: 7
};
const PULIZIA_STATI = [
  { id:'pulita',         label:'Pulita',                cls:'s-pulita'         },
  { id:'da-pulire',      label:'Da pulire',             cls:'s-da-pulire'      },
  { id:'in-corso',       label:'Controllare/Rassettare',cls:'s-in-corso'       },
  { id:'fuori-servizio', label:'Fuori servizio',        cls:'s-fuori-servizio' },
];

// Costanti foglio annuale — devono corrispondere all'Apps Script
const FIRST_DATA_ROW  = 3;
const HEADER_ROW      = 2;
const OUTPUT_ROW      = 45;
const BLIP_ID_ROW     = 46; // Riga dove Blip scrive i propri ID per ogni camera
const DATES_COL       = 1;
const FIRST_ROOM_COL  = 2;
const EXCLUDED_SHEETS = ['Dati Centralizzati Realtime','Non toccare','Ricettività','LOG COMPLESSIVO','PRENOTAZIONI'];

const MONTHS_IT = ['Gennaio','Febbraio','Marzo','Aprile','Maggio','Giugno',
                   'Luglio','Agosto','Settembre','Ottobre','Novembre','Dicembre'];
const MONTHS_S  = ['GEN','FEB','MAR','APR','MAG','GIU','LUG','AGO','SET','OTT','NOV','DIC'];
const DAYS_IT   = ['Do','Lu','Ma','Me','Gi','Ve','Sa'];

const ROOMS = [
  {id:'r1',  name:'1',   g:'Scuola'}, {id:'r2',  name:'2',   g:'Scuola'},
  {id:'r3',  name:'3',   g:'Scuola'}, {id:'r4',  name:'4',   g:'Scuola'},
  {id:'r5',  name:'5',   g:'Scuola'}, {id:'r6',  name:'6',   g:'Scuola'},
  {id:'r7',  name:'7',   g:'Scuola'}, {id:'r8',  name:'8',   g:'Scuola'},
  {id:'r9',  name:'9',   g:'Scuola'}, {id:'r10', name:'10',  g:'Scuola'},
  {id:'r100',name:'100', g:'Scuola'}, {id:'r101',name:'101', g:'Scuola'},
  {id:'r102',name:'102', g:'Scuola'}, {id:'r103',name:'103', g:'Scuola'},
  {id:'r104',name:'104', g:'Scuola'},
  {id:'r21', name:'21',  g:'Largo Roma'}, {id:'r22', name:'22',  g:'Largo Roma'},
  {id:'r23', name:'23',  g:'Largo Roma'}, {id:'r24', name:'24',  g:'Largo Roma'},
  {id:'r25', name:'25',  g:'Largo Roma'}, {id:'r31', name:'31',  g:'Largo Roma'},
  {id:'marasanta', name:'marasanta',       g:'Appartamenti'},
  {id:'margher',   name:'margherita p.t.', g:'Appartamenti'},
  {id:'marg1',     name:'marg.ta 1 p.',   g:'Appartamenti'},
  {id:'sole',      name:'sole',            g:'Appartamenti'},
  {id:'giove',     name:'giove 4/6',       g:'Appartamenti'},
  {id:'saturno',   name:'saturno 3',       g:'Appartamenti'},
  {id:'nettuno',   name:'nettuno 2',       g:'Appartamenti'},
  {id:'marte',     name:'marte 3/4',       g:'Appartamenti'},
  {id:'venere',    name:'venere 2',        g:'Appartamenti'},
  {id:'mercurio',  name:'mercurio 3',      g:'Appartamenti'},
  {id:'giuliaB',   name:'Giulia B',        g:'Appartamenti'},
  {id:'giuliaC',   name:'Giulia C',        g:'Appartamenti'},
  {id:'giuliaD',   name:'Giulia D',        g:'Appartamenti'},
  {id:'giuliaE',   name:'Giulia E',        g:'Appartamenti'},
];

const PALETTE = [
  {h:'#D9D9D9',n:'Grigio'},      {h:'#FCE5CD',n:'Pesca'},
  {h:'#B6D7A8',n:'Verde salvia'},{h:'#CFE2F3',n:'Azzurro'},
  {h:'#FFE599',n:'Giallo'},      {h:'#EA9999',n:'Rosato'},
  {h:'#F9CB9C',n:'Arancio'},     {h:'#B4A7D6',n:'Lavanda'},
  {h:'#A2C4C9',n:'Turchese'},    {h:'#D5A6BD',n:'Rosa antico'},
  {h:'#C9DAF8',n:'Celeste'},     {h:'#D9EAD3',n:'Menta'},
  {h:'#FFF2CC',n:'Crema'},       {h:'#F4CCCC',n:'Salmone'},
  {h:'#EAD1DC',n:'Cipria'},      {h:'#93C47D',n:'Verde'},
  {h:'#76D7EA',n:'Ciano'},       {h:'#76A5AF',n:'Petrolio'},
];

const BED_TYPES = [
  { id:'m',   label:'Matrimoniale',             short:'m'   },
  { id:'ms',  label:'Matrimoniale uso singolo', short:'ms'  },
  { id:'s',   label:'Singolo',                  short:'s'   },
  { id:'c',   label:'Culla',                    short:'c'   },
  { id:'aff', label:'Affollato/Extra',           short:'aff' },
];

const ROOM_DEFAULTS = {
  r1:{maxGuests:2,allowedBeds:['m','ms','s']},    r2:{maxGuests:2,allowedBeds:['m','ms','s']},
  r3:{maxGuests:2,allowedBeds:['m','ms','s']},    r4:{maxGuests:2,allowedBeds:['m','ms','s']},
  r5:{maxGuests:2,allowedBeds:['m','ms','s']},    r6:{maxGuests:2,allowedBeds:['m','ms','s']},
  r7:{maxGuests:2,allowedBeds:['m','ms','s']},    r8:{maxGuests:2,allowedBeds:['m','ms','s']},
  r9:{maxGuests:2,allowedBeds:['m','ms','s']},    r10:{maxGuests:2,allowedBeds:['m','ms','s']},
  r100:{maxGuests:2,allowedBeds:['m','ms','s','c']}, r101:{maxGuests:2,allowedBeds:['m','ms','s','c']},
  r102:{maxGuests:2,allowedBeds:['m','ms','s','c']}, r103:{maxGuests:2,allowedBeds:['m','ms','s','c']},
  r104:{maxGuests:2,allowedBeds:['m','ms','s','c']},
  r21:{maxGuests:3,allowedBeds:['m','ms','s','c']}, r22:{maxGuests:3,allowedBeds:['m','ms','s','c']},
  r23:{maxGuests:3,allowedBeds:['m','ms','s','c']}, r24:{maxGuests:3,allowedBeds:['m','ms','s','c']},
  r25:{maxGuests:3,allowedBeds:['m','ms','s','c']}, r31:{maxGuests:4,allowedBeds:['m','ms','s','c']},
  marasanta:{maxGuests:6,allowedBeds:['m','ms','s','c','aff']},
  margher:  {maxGuests:4,allowedBeds:['m','ms','s','c']},
  marg1:    {maxGuests:4,allowedBeds:['m','ms','s','c']},
  sole:     {maxGuests:4,allowedBeds:['m','ms','s','c']},
  giove:    {maxGuests:6,allowedBeds:['m','ms','s','c','aff']},
  saturno:  {maxGuests:3,allowedBeds:['m','ms','s','c']},
  nettuno:  {maxGuests:2,allowedBeds:['m','ms','s']},
  marte:    {maxGuests:4,allowedBeds:['m','ms','s','c']},
  venere:   {maxGuests:2,allowedBeds:['m','ms','s']},
  mercurio: {maxGuests:3,allowedBeds:['m','ms','s','c']},
  giuliaB:  {maxGuests:4,allowedBeds:['m','ms','s','c']},
  giuliaC:  {maxGuests:4,allowedBeds:['m','ms','s','c']},
  giuliaD:  {maxGuests:4,allowedBeds:['m','ms','s','c']},
  giuliaE:  {maxGuests:4,allowedBeds:['m','ms','s','c']},
};

// ═══════════════════════════════════════════════════════════════════
// ROOM SETTINGS — localStorage
// ═══════════════════════════════════════════════════════════════════
function loadRoomSettings() {
  try {
    const saved = localStorage.getItem('hotelRoomSettings');
    if (saved) {
      const settings = JSON.parse(saved);
      let migrated = false;
      for (const rid of Object.keys(settings)) {
        const beds = settings[rid].allowedBeds || [];
        if (beds.includes('m') && !beds.includes('ms')) {
          const idx = beds.indexOf('m');
          beds.splice(idx + 1, 0, 'ms');
          settings[rid].allowedBeds = beds;
          migrated = true;
        }
      }
      if (migrated) localStorage.setItem('hotelRoomSettings', JSON.stringify(settings));
      return settings;
    }
  } catch(e) {}
  return JSON.parse(JSON.stringify(ROOM_DEFAULTS));
}
function saveRoomSettingsLS(settings) {
  localStorage.setItem('hotelRoomSettings', JSON.stringify(settings));
}

// ═══════════════════════════════════════════════════════════════════
// STATO APPLICAZIONE — variabili globali condivise tra i moduli
// ═══════════════════════════════════════════════════════════════════
let accessToken  = null;
let curM         = new Date().getMonth();
let curY         = new Date().getFullYear();
let selColor     = '#D9D9D9';
let editId       = null;
let nid          = 1000;
let bedCounts    = { m:0, ms:0, s:0, c:0, aff:0 };
let roomSettings = loadRoomSettings();
let roomStates   = {};          // { roomId: { pulizia, configurazione, noteOps, ts, dbRow } }
let bookings     = [];
let sheetColumnMap = {};        // sheetName → { roomName → colIdx }
let annualSheets = [];          // inizializzato dopo la definizione di loadAnnualSheets()
let dbRowCache   = [];
let _rdashFilter      = 'tutti';
let _rstateEditRoom   = null;

// ═══════════════════════════════════════════════════════════════════
// HELPERS — configurazione localStorage
// ═══════════════════════════════════════════════════════════════════
function loadAnnualSheets() {
  try {
    const s = localStorage.getItem('hotelAnnualSheets');
    if (s) return JSON.parse(s);
  } catch(e) {}
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
