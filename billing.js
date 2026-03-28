// ═══════════════════════════════════════════════════════════════════
// billing.js — Conti, tariffe, PDF, XML FatturaPA, listino
// Blip Hotel Management — build 18.7.x
// Dipende da: core.js, sync.js
// ═══════════════════════════════════════════════════════════════════



const BLIP_VER_BILLING = '12'; // ← incrementa ad ogni modifica

const BILL_SETTINGS_KEY = 'hotelBillSettings';
const BILL_CONTI_KEY    = 'hotelConti';

// ─────────────────────────────────────────────────────────────────
// STRUTTURA TARIFFE
//
// Logica di calcolo notte per camera albergo:
//   1. Calcola tariffa dalla DISPOSIZIONE LETTI:
//        singolo:                  T.s (es. 35€)
//        matrimoniale uso singolo: T.ms (es. 38€)
//        matrimoniale (2 pers):    T.m  (es. 45€)
//        singolo aggiunto:         +T.ag a persona (es. +15€)
//        combinazione m+s:         T.m + singoli×T.ag
//   2. Se la camera ha override giornaliero > 0 → usa quello
//   3. Applica moltiplicatore stagionale (media sui giorni)
//   4. Applica convenzione cliente (sconto%)
//   5. Oppure sconto durata (se no convenzione)
//
// Appartamenti: tariffa override per camera (giornaliero o mensile)
// ─────────────────────────────────────────────────────────────────

function billSettingsDefault() {
  return {
    hotelName:    'Il mio Hotel',
    hotelAddress: '',
    hotelTel:     '',

    // ── Tariffe base per DISPOSIZIONE ──
    tariffe: {
      s:  35,   // camera singola (1 letto singolo)
      ms: 38,   // matrimoniale uso singolo
      m:  45,   // matrimoniale (2 persone)
      ag: 15,   // aggiunta singolo in camera matrimoniale (per persona)
    },

    // ── Override per singola camera (0 = usa disposizione) ──
    // { [roomId]: { giornaliera: 0, mensile: 0 } }
    tariffeCamere: {},

    // ── Stagionalità: periodi con moltiplicatore ──
    stagioni: [
      { nome:'Alta stagione',  dal:'06-15', al:'09-15', molt:1.5 },
      { nome:'Media stagione', dal:'03-15', al:'06-14', molt:1.2 },
      { nome:'Bassa stagione', dal:'09-16', al:'03-14', molt:1.0 },
    ],

    // ── Convenzioni clienti (sconto % sul totale) ──
    convenzioni: [
      { nome:'mammana',       sconto:20 },
      { nome:'geoambiente',   sconto:20 },
      { nome:'Sicily Divide', sconto:15 },
    ],

    // ── IVA ──
    aliquotaIVA: 10,   // % IVA applicata al conto (default 10% alloggio)

    // ── Sconti durata ──
    scontiDurata: [
      { soglia:30, sconto:20 },
      { soglia:7,  sconto:10 },
    ],

    // ── Extra: { label, prezzo, unita } ──
    // unita: 'notte' | 'persona' | 'volta' | 'kwh' | 'mc'
    extra: [
      { id:'colazione',      label:'🍳 Colazione',          prezzo:0,  unita:'persona' },
      { id:'pranzo',         label:'🍽 Pranzo',             prezzo:0,  unita:'persona' },
      { id:'cena',           label:'🌙 Cena',               prezzo:0,  unita:'persona' },
      { id:'piscina',        label:'🏊 Piscina',            prezzo:0,  unita:'persona' },
      { id:'pulizie',        label:'🧹 Pulizie extra',      prezzo:0,  unita:'volta'   },
      { id:'cambioLenzuola', label:'🛏 Cambio lenzuola',   prezzo:0,  unita:'volta'   },
      { id:'luce',           label:'⚡ Consumo elettrico',  prezzo:0,  unita:'kwh'     },
      { id:'acqua',          label:'💧 Consumo idrico',     prezzo:0,  unita:'mc'      },
    ],
  };
}

// ═══════════════════════════════════════════════════════════════════
// BILLING DB LAYER
// ═══════════════════════════════════════════════════════════════════
// Tutte le informazioni operative di fatturazione vivono sul DATABASE
// Google (foglio condiviso), non in localStorage.
// Il localStorage è usato SOLO come cache read-through con TTL 5min.
//
// Schede nel foglio DATABASE:
//   IMPOSTAZIONI  — tariffe, convenzioni, extra (riga 2 = JSON)
//   CONTI         — una riga per prenotazione con extra, override, modo
//
// Pattern di accesso:
//   leggi → prova cache → se scaduta leggi DB → aggiorna cache
//   scrivi → scrivi DB → invalida cache
// ═══════════════════════════════════════════════════════════════════

const IMPOSTAZIONI_SHEET = 'IMPOSTAZIONI';
const CONTI_SHEET        = 'CONTI';
const PAGAMENTI_SHEET    = 'PAGAMENTI';
const BILL_DB_TTL        = 5 * 60 * 1000; // cache 5 minuti

// Colonne foglio PAGAMENTI
const PAG_COLS = {
  ID:           1,  // A  PAG-2026-XXXXXX
  CONTO_ID:     2,  // B
  BOOKING_ID:   3,  // C
  DATA:         4,  // D  gg/mm/aaaa
  IMPORTO:      5,  // E  numero
  TIPO:         6,  // F  acconto/saldo/extra
  METODO:       7,  // G  contanti/carta/bonifico/assegno/altro
  RIFERIMENTO:  8,  // H  es. "Visa *4521"
  CON_DOCUMENTO:9,  // I  true/false
  NOTE:         10, // J
  TS:           11, // K
};

// Cache pagamenti in-memory
let _pagamentiCache = null;
let _pagamentiCacheTs = 0;
let _pagamentiSheetReady = false;

// ── Cache keys (localStorage, solo temporanea) ──
const _C_SETTINGS = 'hotelBillSettingsCache';
const _C_CONTI    = 'hotelContiCache';
const _C_DATI     = 'hotelContoDatiCache'; // extra+override+modo per bid

// ─────────────────────────────────────────────────────────────────
// IMPOSTAZIONI (tariffe, convenzioni, extra)
// ─────────────────────────────────────────────────────────────────

async function ensureImpostazioniSheet() {
  if (!DATABASE_SHEET_ID) return;
  // Verifica se la scheda esiste
  try {
    const meta = await apiFetch(
      `https://sheets.googleapis.com/v4/spreadsheets/${DATABASE_SHEET_ID}?fields=sheets.properties.title`
    );
    const mj = await meta.json();
    const exists = (mj.sheets||[]).some(s=>s.properties.title===IMPOSTAZIONI_SHEET);
    if (!exists) {
      await apiFetch(`https://sheets.googleapis.com/v4/spreadsheets/${DATABASE_SHEET_ID}:batchUpdate`,{
        method:'POST', headers:{'Content-Type':'application/json'},
        body: JSON.stringify({ requests:[{ addSheet:{ properties:{ title:IMPOSTAZIONI_SHEET } } }] })
      });
    }
    // Scrivi intestazione riga 1
    const u = `https://sheets.googleapis.com/v4/spreadsheets/${DATABASE_SHEET_ID}/values/${encodeURIComponent(IMPOSTAZIONI_SHEET+'!A1:B1')}?valueInputOption=RAW`;
    const d = await dbGet(`${IMPOSTAZIONI_SHEET}!A1:A1`);
    if (!d.values?.[0]?.[0]) {
      await apiFetch(u,{ method:'PUT', headers:{'Content-Type':'application/json'},
        body: JSON.stringify({ values:[['CHIAVE','VALORE']] }) });
    }
  } catch(e) { console.warn('[IMPOSTAZIONI] ensure:', e.message); }
}

async function loadBillSettingsDB() {
  // 1. Prova cache
  try {
    const raw = localStorage.getItem(_C_SETTINGS);
    if (raw) {
      const p = JSON.parse(raw);
      if (Date.now() - p.ts < BILL_DB_TTL) return mergeSettings(p.data);
    }
  } catch(e) {}

  // 2. Leggi dal DB
  if (!DATABASE_SHEET_ID) return loadBillSettingsLocal();
  try {
    const d = await dbGet(`${IMPOSTAZIONI_SHEET}!A2:B99`);
    const rows = d.values || [];
    const map = {};
    rows.forEach(r => { if(r[0]) map[r[0]] = r[1]; });
    if (map['billSettings']) {
      const saved = JSON.parse(map['billSettings']);
      localStorage.setItem(_C_SETTINGS, JSON.stringify({ ts:Date.now(), data:saved }));
      return mergeSettings(saved);
    }
  } catch(e) { console.warn('[IMPOSTAZIONI] load:', e.message); }

  return loadBillSettingsLocal(); // fallback
}

async function saveBillSettingsDB(s) {
  // Invalida cache
  localStorage.removeItem(_C_SETTINGS);
  // Salva anche in localStorage per compatibilità
  localStorage.setItem(BILL_SETTINGS_KEY, JSON.stringify(s));

  if (!DATABASE_SHEET_ID) return;
  try {
    await ensureImpostazioniSheet();
    // Leggi righe esistenti per trovare/aggiornare la riga billSettings
    const d = await dbGet(`${IMPOSTAZIONI_SHEET}!A2:B99`);
    const rows = d.values || [];
    const idx  = rows.findIndex(r=>r[0]==='billSettings');
    const rowNum = idx >= 0 ? idx + 2 : rows.length + 2;
    const url = `https://sheets.googleapis.com/v4/spreadsheets/${DATABASE_SHEET_ID}/values/${encodeURIComponent(`${IMPOSTAZIONI_SHEET}!A${rowNum}:B${rowNum}`)}?valueInputOption=RAW`;
    await apiFetch(url,{ method:'PUT', headers:{'Content-Type':'application/json'},
      body: JSON.stringify({ values:[['billSettings', JSON.stringify(s)]] }) });
    // Aggiorna cache
    localStorage.setItem(_C_SETTINGS, JSON.stringify({ ts:Date.now(), data:s }));
  } catch(e) { console.warn('[IMPOSTAZIONI] save:', e.message); }
}

// ─────────────────────────────────────────────────────────────────
// DATI CONTO PER PRENOTAZIONE (extra, override, modo appartamento)
// Una riga per prenotazione nella scheda CONTI
// Colonne: BOOKING_ID | EXTRA_JSON | OVERRIDE_JSON | APPART_MODE | TS
// ─────────────────────────────────────────────────────────────────

let _contiSheetReady = false;

async function ensureContiSheet() {
  if (_contiSheetReady || !DATABASE_SHEET_ID) return;
  _contiSheetReady = true;
  try {
    const meta = await apiFetch(
      `https://sheets.googleapis.com/v4/spreadsheets/${DATABASE_SHEET_ID}?fields=sheets.properties.title`
    );
    const mj = await meta.json();
    if (!(mj.sheets||[]).some(s=>s.properties.title===CONTI_SHEET)) {
      await apiFetch(`https://sheets.googleapis.com/v4/spreadsheets/${DATABASE_SHEET_ID}:batchUpdate`,{
        method:'POST', headers:{'Content-Type':'application/json'},
        body: JSON.stringify({ requests:[{ addSheet:{ properties:{ title:CONTI_SHEET } } }] })
      });
    }
    const hd = await dbGet(`${CONTI_SHEET}!A1:F1`);
    if (!hd.values?.[0]?.[0]) {
      const u = `https://sheets.googleapis.com/v4/spreadsheets/${DATABASE_SHEET_ID}/values/${encodeURIComponent(CONTI_SHEET+'!A1:F1')}?valueInputOption=RAW`;
      await apiFetch(u,{ method:'PUT', headers:{'Content-Type':'application/json'},
        body: JSON.stringify({ values:[['BOOKING_ID','EXTRA_JSON','OVERRIDE_JSON','APPART_MODE','CONTO_EMESSO_JSON','TS']] }) });
    }
  } catch(e) { console.warn('[CONTI] ensure:', e.message); }
}

// ─────────────────────────────────────────────────────────────────
// FOGLIO PAGAMENTI — setup + CRUD
// ─────────────────────────────────────────────────────────────────
async function ensurePagamentiSheet() {
  if (_pagamentiSheetReady || !DATABASE_SHEET_ID) return;
  _pagamentiSheetReady = true;
  try {
    const meta = await apiFetch(`https://sheets.googleapis.com/v4/spreadsheets/${DATABASE_SHEET_ID}?fields=sheets.properties.title`);
    const mj = await meta.json();
    if (!(mj.sheets||[]).some(s => s.properties.title === PAGAMENTI_SHEET)) {
      await apiFetch(`https://sheets.googleapis.com/v4/spreadsheets/${DATABASE_SHEET_ID}:batchUpdate`, {
        method:'POST', headers:{'Content-Type':'application/json'},
        body: JSON.stringify({ requests:[{ addSheet:{ properties:{ title:PAGAMENTI_SHEET } } }] })
      });
    }
    const hd = await dbGet(`${PAGAMENTI_SHEET}!A1:K1`);
    if (!hd.values?.[0]?.[0]) {
      const u = `https://sheets.googleapis.com/v4/spreadsheets/${DATABASE_SHEET_ID}/values/${encodeURIComponent(PAGAMENTI_SHEET+'!A1:K1')}?valueInputOption=RAW`;
      await apiFetch(u, { method:'PUT', headers:{'Content-Type':'application/json'},
        body: JSON.stringify({ values:[['PAG_ID','CONTO_ID','BOOKING_ID','DATA','IMPORTO','TIPO','METODO','RIFERIMENTO','CON_DOCUMENTO','NOTE','TS']] })
      });
    }
  } catch(e) { console.warn('[PAGAMENTI] ensure:', e.message); }
}

async function loadPagamentiPerBooking(bid) {
  if (!DATABASE_SHEET_ID) return [];
  // Usa cache se fresca (5 min)
  if (_pagamentiCache && Date.now() - _pagamentiCacheTs < BILL_DB_TTL) {
    return (_pagamentiCache || []).filter(p => p.bookingId === bid);
  }
  try {
    await ensurePagamentiSheet();
    const d = await dbGet(`${PAGAMENTI_SHEET}!A2:K9999`);
    const rows = d.values || [];
    _pagamentiCache = rows.map((row, i) => ({
      id:           (row[PAG_COLS.ID-1]||'').trim(),
      contoId:      (row[PAG_COLS.CONTO_ID-1]||'').trim(),
      bookingId:    parseInt(row[PAG_COLS.BOOKING_ID-1])||0,
      data:         (row[PAG_COLS.DATA-1]||'').trim(),
      importo:      parseFloat(row[PAG_COLS.IMPORTO-1])||0,
      tipo:         (row[PAG_COLS.TIPO-1]||'saldo').trim(),
      metodo:       (row[PAG_COLS.METODO-1]||'Contanti').trim(),
      riferimento:  (row[PAG_COLS.RIFERIMENTO-1]||'').trim(),
      conDocumento: (row[PAG_COLS.CON_DOCUMENTO-1]||'').trim() === 'true',
      note:         (row[PAG_COLS.NOTE-1]||'').trim(),
      ts:           (row[PAG_COLS.TS-1]||'').trim(),
      dbRow:        i + 2,
    })).filter(p => p.id && p.importo > 0);
    _pagamentiCacheTs = Date.now();
    return _pagamentiCache.filter(p => p.bookingId === bid);
  } catch(e) {
    console.warn('[PAGAMENTI] load:', e.message);
    return [];
  }
}

function getPagamentiPerBookingSync(bid) {
  if (!_pagamentiCache) return [];
  return _pagamentiCache.filter(p => p.bookingId === bid);
}

function getTotalePagatoPerBooking(bid) {
  return getPagamentiPerBookingSync(bid).reduce((acc, p) => acc + p.importo, 0);
}

async function registraPagamento(pag) {
  if (!DATABASE_SHEET_ID) throw new Error('DATABASE_SHEET_ID non configurato');
  await ensurePagamentiSheet();
  const anno = new Date().getFullYear();
  const id = pag.id || genPagamentoId(anno);
  const oggi = new Date().toLocaleDateString('it-IT');
  const row = [
    id,
    pag.contoId    || '',
    String(pag.bookingId || ''),
    pag.data       || oggi,
    String(pag.importo   || 0),
    pag.tipo       || 'saldo',
    pag.metodo     || 'Contanti',
    pag.riferimento|| '',
    pag.conDocumento ? 'true' : 'false',
    pag.note       || '',
    nowISO(),
  ];
  const u = `https://sheets.googleapis.com/v4/spreadsheets/${DATABASE_SHEET_ID}/values/${encodeURIComponent(PAGAMENTI_SHEET)}:append?valueInputOption=RAW&insertDataOption=INSERT_ROWS`;
  const r = await apiFetch(u, { method:'POST', headers:{'Content-Type':'application/json'}, body: JSON.stringify({ values:[row] }) });
  if (!r.ok) throw new Error(`Registrazione pagamento fallita (${r.status})`);
  // Aggiorna cache
  if (!_pagamentiCache) _pagamentiCache = [];
  _pagamentiCache.push({ ...pag, id, dbRow: _pagamentiCache.length + 2 });
  _pagamentiCacheTs = Date.now();
  return id;
}

function eliminaPagamentoUI(pagId, bid) {
  if (!confirm('Eliminare questo pagamento?')) return;
  eliminaPagamento(pagId).then(() => {
    // Ricarica la sezione pagamenti nel drawer
    if (typeof refreshBillTab === 'function') refreshBillTab(bid);
    // Se il conto era 'pagato' e abbiamo rimosso un pagamento, torna a 'emesso'
    const conti = loadConti();
    const conto = conti.find(c => c.bookingId === bid);
    if (conto && conto.status === 'pagato') {
      const rimasto = getTotalePagatoPerBooking(bid);
      if (rimasto < (conto.totale||0) - 0.01) {
        const idx = conti.findIndex(c => c.bookingId === bid);
        if (idx >= 0) { conti[idx] = { ...conti[idx], status:'emesso', ts:nowISO() }; saveConti(conti); }
        if (typeof render === 'function') render();
      }
    }
  }).catch(e => showToast('Errore eliminazione: '+e.message, 'error'));
}

async function eliminaPagamento(pagId) {
  if (!_pagamentiCache) return;
  const idx = _pagamentiCache.findIndex(p => p.id === pagId);
  if (idx < 0) return;
  const pag = _pagamentiCache[idx];
  // Elimina riga dal foglio
  if (pag.dbRow && DATABASE_SHEET_ID) {
    try {
      const meta = await apiFetch(`https://sheets.googleapis.com/v4/spreadsheets/${DATABASE_SHEET_ID}?fields=sheets.properties`);
      const mj = await meta.json();
      const sheet = (mj.sheets||[]).find(s => s.properties.title === PAGAMENTI_SHEET);
      const sheetNumId = sheet?.properties?.sheetId ?? null;
      if (sheetNumId !== null) {
        await apiFetch(`https://sheets.googleapis.com/v4/spreadsheets/${DATABASE_SHEET_ID}:batchUpdate`, {
          method:'POST', headers:{'Content-Type':'application/json'},
          body: JSON.stringify({ requests:[{ deleteDimension: { range: { sheetId:sheetNumId, dimension:'ROWS', startIndex:pag.dbRow-1, endIndex:pag.dbRow } } }] })
        });
      }
    } catch(e) { console.warn('[PAGAMENTI] elimina:', e.message); }
  }
  _pagamentiCache.splice(idx, 1);
  // Riaggiusta dbRow nei pagamenti successivi
  _pagamentiCache.forEach((p, i) => { if (p.dbRow > pag.dbRow) p.dbRow--; });
}

// Cache in-memory per la sessione (riduce letture DB durante la stessa sessione)
const _contiDatiCache = {}; // bid → { extra, override, appartMode, contoEmesso, dbRow, ts }

async function loadContoDati(bid) {
  // 1. Cache in-memory
  if (_contiDatiCache[bid]) return _contiDatiCache[bid];

  // 2. Cache localStorage (5min)
  try {
    const raw = localStorage.getItem(`${_C_DATI}_${bid}`);
    if (raw) {
      const p = JSON.parse(raw);
      if (Date.now() - p.ts < BILL_DB_TTL) {
        _contiDatiCache[bid] = p.data;
        return p.data;
      }
    }
  } catch(e) {}

  // 3. Leggi dal DB (fallback su localStorage legacy)
  const def = { extra:[], override:null, appartMode:null, contoEmesso:null, dbRow:null };

  if (!DATABASE_SHEET_ID) return _migraDatiLocali(bid, def);

  try {
    await ensureContiSheet();
    const d = await dbGet(`${CONTI_SHEET}!A2:F9999`);
    const rows = d.values || [];
    // Cerca tutte le prenotazioni e metti in cache (1 lettura per tutte)
    rows.forEach((row, i) => {
      const id = String(row[0]||'').trim();
      if (!id) return;
      const dati = {
        extra:        _jsonParse(row[1], []),
        override:     _jsonParse(row[2], null),
        appartMode:   row[3] || null,
        contoEmesso:  _jsonParse(row[4], null),
        dbRow:        i + 2,
        ts:           row[5] || ''
      };
      _contiDatiCache[id] = dati;
      localStorage.setItem(`${_C_DATI}_${id}`, JSON.stringify({ ts:Date.now(), data:dati }));
    });
    return _contiDatiCache[bid] || _migraDatiLocali(bid, def);
  } catch(e) {
    console.warn('[CONTI] load:', e.message);
    return _migraDatiLocali(bid, def);
  }
}

function _migraDatiLocali(bid, def) {
  // Migra dati legacy da localStorage se presenti
  const extra    = _jsonParse(localStorage.getItem(`billExtra_${bid}`), []);
  const override = _jsonParse(localStorage.getItem(`billOv_${bid}`), null);
  const appartM  = localStorage.getItem(`appartMode_${bid}`) || null;
  return { ...def, extra, override, appartMode:appartM };
}

async function saveContoDati(bid, patch) {
  // Aggiorna cache in-memory
  const cur = _contiDatiCache[bid] || { extra:[], override:null, appartMode:null, contoEmesso:null, dbRow:null };
  const next = { ...cur, ...patch, ts: nowISO() };
  _contiDatiCache[bid] = next;
  // Invalida localStorage cache
  localStorage.removeItem(`${_C_DATI}_${bid}`);
  // Rimuovi dati legacy
  localStorage.removeItem(`billExtra_${bid}`);
  localStorage.removeItem(`billOv_${bid}`);
  localStorage.removeItem(`appartMode_${bid}`);

  if (!DATABASE_SHEET_ID) return;
  try {
    await ensureContiSheet();
    const row = [
      String(bid),
      JSON.stringify(next.extra || []),
      next.override ? JSON.stringify(next.override) : '',
      next.appartMode || '',
      next.contoEmesso ? JSON.stringify(next.contoEmesso) : '',
      next.ts
    ];
    if (next.dbRow) {
      const u = `https://sheets.googleapis.com/v4/spreadsheets/${DATABASE_SHEET_ID}/values/${encodeURIComponent(`${CONTI_SHEET}!A${next.dbRow}:F${next.dbRow}`)}?valueInputOption=RAW`;
      await apiFetch(u,{ method:'PUT', headers:{'Content-Type':'application/json'},
        body: JSON.stringify({ values:[row] }) });
    } else {
      const u = `https://sheets.googleapis.com/v4/spreadsheets/${DATABASE_SHEET_ID}/values/${encodeURIComponent(CONTI_SHEET)}:append?valueInputOption=RAW&insertDataOption=INSERT_ROWS`;
      const resp = await apiFetch(u,{ method:'POST', headers:{'Content-Type':'application/json'},
        body: JSON.stringify({ values:[row] }) });
      const rj = await resp.json();
      const m  = (rj.updates?.updatedRange||'').match(/(\d+):/);
      if (m) { next.dbRow = parseInt(m[1]); _contiDatiCache[bid] = next; }
    }
  } catch(e) { console.warn('[CONTI] save:', e.message); }
}

function _jsonParse(s, fallback) {
  if (!s || s === '') return fallback;
  try { return JSON.parse(s); } catch(e) { return fallback; }
}

// ─────────────────────────────────────────────────────────────────
// API PUBBLICA — sostituisce le vecchie funzioni localStorage
// ─────────────────────────────────────────────────────────────────

// SINCRONO (usa cache in-memory, parte da vuoto se non ancora caricato)
function getExtraForBooking(bid) {
  return _contiDatiCache[bid]?.extra ?? _jsonParse(localStorage.getItem(`billExtra_${bid}`), []);
}
function setExtraForBooking(bid, extras) {
  if (_contiDatiCache[bid]) _contiDatiCache[bid].extra = extras;
  else _contiDatiCache[bid] = { extra:extras, override:null, appartMode:null, contoEmesso:null, dbRow:null };
  saveContoDati(bid, { extra: extras }); // fire-and-forget
}

function getContoOverrides(bid) {
  return _contiDatiCache[bid]?.override ?? _jsonParse(localStorage.getItem(`billOv_${bid}`), null);
}
function setContoOverrides(bid, righe) {
  if (_contiDatiCache[bid]) _contiDatiCache[bid].override = righe;
  else _contiDatiCache[bid] = { extra:[], override:righe, appartMode:null, contoEmesso:null, dbRow:null };
  saveContoDati(bid, { override: righe }); // fire-and-forget
}

function getAppartMode(bid, notti) {
  const m = _contiDatiCache[bid]?.appartMode ?? localStorage.getItem(`appartMode_${bid}`);
  if (m === 'giornaliera') return 'giornaliera';
  if (m === 'mensile')     return 'mensile';
  if (m === 'standard')    return 'standard';
  return notti >= 20 ? 'mensile' : 'giornaliera';
}
function toggleAppartMode(bid) {
  const b = bookings.find(x=>x.id===bid); if(!b) return;
  const cur = getAppartMode(bid, nights(b.s,b.e));
  const ciclo = { 'giornaliera':'mensile', 'mensile':'standard', 'standard':'giornaliera' };
  const next = ciclo[cur] || 'giornaliera';
  if (_contiDatiCache[bid]) _contiDatiCache[bid].appartMode = next;
  else _contiDatiCache[bid] = { extra:[], override:null, appartMode:next, contoEmesso:null, dbRow:null };
  saveContoDati(bid, { appartMode: next });
  refreshBillTab(bid);
}

// Conti emessi: salvati nella riga del conto specifica
function loadConti() {
  // Raccoglie tutti i contoEmesso dalla cache in-memory
  return Object.entries(_contiDatiCache)
    .filter(([,d]) => d.contoEmesso)
    .map(([,d]) => d.contoEmesso)
    .sort((a,b) => (b.emessoIl||'').localeCompare(a.emessoIl||''));
}
function saveConti(contiArray) {
  // Salva ogni conto nella riga della prenotazione corrispondente
  contiArray.forEach(doc => {
    const bid = doc.bookingId;
    if (!bid) return;
    if (_contiDatiCache[bid]) _contiDatiCache[bid].contoEmesso = doc;
    else _contiDatiCache[bid] = { extra:[], override:null, appartMode:null, contoEmesso:doc, dbRow:null };
    saveContoDati(bid, { contoEmesso: doc });
  });
}

// ─────────────────────────────────────────────────────────────────
// HELPER: stato conto per una prenotazione (usato dal Gantt)
// Legge solo la cache in-memory — nessuna chiamata API, sempre sincrono.
// Ritorna: null | 'bozza' | 'emesso' | 'fatturato' | 'pagato'
// ─────────────────────────────────────────────────────────────────
function getBillingStatusForBooking(bid) {
  const dati = _contiDatiCache[bid];
  if (!dati?.contoEmesso) return null;
  return dati.contoEmesso.status || 'bozza';
}

// Colore bordo sinistro Gantt per stato conto
function billingBorderColor(bid) {
  const stato = getBillingStatusForBooking(bid);
  if (stato === 'pagato')    return '#34a853'; // verde
  if (stato === 'fatturato') return '#4285f4'; // blu
  if (stato === 'emesso')    return '#fa7b17'; // arancione
  if (stato === 'bozza')     return '#9e9e9e'; // grigio
  return null; // nessun conto → nessun bordo
}

// Settings billing (async)
function loadBillSettings() { return loadBillSettingsLocal(); } // sincrono per compatibilità render
function loadBillSettingsLocal() {
  try {
    const raw = localStorage.getItem(BILL_SETTINGS_KEY);
    if (!raw) return billSettingsDefault();
    return mergeSettings(JSON.parse(raw));
  } catch(e) { return billSettingsDefault(); }
}
function mergeSettings(saved) {
  const def = billSettingsDefault();
  return { ...def, ...saved, tariffe:{ ...def.tariffe,...(saved.tariffe||{}) }, tariffeCamere:{...(saved.tariffeCamere||{})}, aliquotaIVA: saved.aliquotaIVA ?? def.aliquotaIVA };
}
function saveBillSettings(s) {
  localStorage.setItem(BILL_SETTINGS_KEY, JSON.stringify(s)); // sincrono immediato
  saveBillSettingsDB(s); // asincrono al DB
}

// ── Precarica i dati di conto all'avvio (dopo loadFromSheets) ──
async function preloadContoDati() {
  if (!DATABASE_SHEET_ID) return;
  try {
    await ensureContiSheet();
    // Legge tutta la scheda CONTI in una chiamata e popola la cache
    const d = await dbGet(`${CONTI_SHEET}!A2:F9999`);
    const rows = d.values || [];
    rows.forEach((row, i) => {
      const id = String(row[0]||'').trim();
      if (!id) return;
      const dati = {
        extra:       _jsonParse(row[1], []),
        override:    _jsonParse(row[2], null),
        appartMode:  row[3] || null,
        contoEmesso: _jsonParse(row[4], null),
        dbRow:       i + 2,
        ts:          row[5] || ''
      };
      _contiDatiCache[id] = dati;
    });
    console.log(`[CONTI] Precaricati ${rows.length} record`);
    // Re-render Gantt per mostrare i bordi stato conto sulle barre
    if (typeof render === 'function') render();
    // Precarica anche i pagamenti in background
    ensurePagamentiSheet().then(() =>
      dbGet(`${PAGAMENTI_SHEET}!A2:K9999`).then(d => {
        const rows = d.values || [];
        _pagamentiCache = rows.map((row, i) => ({
          id: (row[0]||'').trim(), contoId:(row[1]||'').trim(),
          bookingId:parseInt(row[2])||0, data:(row[3]||'').trim(),
          importo:parseFloat(row[4])||0, tipo:(row[5]||'saldo').trim(),
          metodo:(row[6]||'Contanti').trim(), riferimento:(row[7]||'').trim(),
          conDocumento:(row[8]||'').trim()==='true', note:(row[9]||'').trim(),
          ts:(row[10]||'').trim(), dbRow:i+2,
        })).filter(p => p.id && p.importo > 0);
        _pagamentiCacheTs = Date.now();
      }).catch(() => {})
    ).catch(() => {});
  } catch(e) { console.warn('[CONTI] preload:', e.message); }
}


// ─────────────────────────────────────────────────────────────────
// MOTORE DI CALCOLO
// ─────────────────────────────────────────────────────────────────

/**
 * Calcola la tariffa notte da disposizione letti.
 * letti: { m, ms, s, c } (output di parseBedString)
 * tariffe: cfg.tariffe
 */
function tariffaDaDisposizione(letti, tariffe) {
  const m  = letti.m  || 0;
  const ms = letti.ms || 0;
  const s  = letti.s  || 0;

  // Matrimoniale uso singolo
  if (ms > 0 && m === 0 && s === 0) return tariffe.ms * ms;

  // Matrimoniale (con o senza singoli aggiuntivi)
  if (m > 0 && s === 0) return tariffe.m * m + (ms > 0 ? tariffe.ms * ms : 0);
  if (m > 0 && s >  0)  return tariffe.m * m + s * tariffe.ag;

  // Solo singoli:
  //   1 singolo → tariffa singola (35€)
  //   2+ singoli nella stessa stanza → equivale a matrimoniale (45€)
  //   ogni singolo ulteriore oltre i 2 → aggiunta (tariffe.ag per persona)
  if (s === 1) return tariffe.s;
  if (s === 2) return tariffe.m;                         // 2s = tariffa matrimoniale
  if (s >  2)  return tariffe.m + (s - 2) * tariffe.ag; // 3s+ = matrimoniale + aggiunte

  // fallback
  return tariffe.s;
}

/**
 * Descrizione leggibile della tariffa usata
 */
function descrizioneTariffa(letti, tariffe, cfg, roomId) {
  const ov = cfg.tariffeCamere?.[roomId];
  if (ov?.giornaliera > 0) return `camera ${roomName(roomId)}: ${ov.giornaliera}€/notte (tariffa camera)`;

  const m  = letti.m  || 0;
  const ms = letti.ms || 0;
  const s  = letti.s  || 0;
  const parts = [];
  if (m  > 0) parts.push(`${m} matrim. ×${tariffe.m}€`);
  if (ms > 0) parts.push(`${ms} m/s ×${tariffe.ms}€`);
  if (m  > 0 && s > 0) parts.push(`${s} sing. ×${tariffe.ag}€`);
  else if (s === 1) parts.push(`1 sing. ×${tariffe.s}€`);
  else if (s === 2) parts.push(`2 sing. = matrimoniale ×${tariffe.m}€`);
  else if (s >  2)  parts.push(`2 sing.=matrim. ×${tariffe.m}€ + ${s-2} ×${tariffe.ag}€`);
  return parts.join(' + ') || `${tariffe.s}€`;
}

/**
 * Moltiplicatore stagionale: media pesata sui giorni del soggiorno.
 */
function calcolaMoltiplicatoreStagionale(s, e, stagioni) {
  if (!stagioni?.length) return 1;
  const n = Math.max(1, Math.round((e - s) / 86400000));
  let sum = 0;
  for (let i = 0; i < n; i++) {
    const d  = new Date(s.getTime() + i * 86400000);
    const mm = String(d.getMonth()+1).padStart(2,'0');
    const dd = String(d.getDate()).padStart(2,'0');
    const k  = `${mm}-${dd}`;
    let mx = 1;
    for (const st of stagioni) {
      if (inStagione(k, st.dal, st.al)) mx = Math.max(mx, parseFloat(st.molt)||1);
    }
    sum += mx;
  }
  return parseFloat((sum/n).toFixed(3));
}
function inStagione(key, dal, al) {
  if (dal <= al) return key >= dal && key <= al;
  return key >= dal || key <= al;
}

/**
 * Calcolo conto alberghiero completo.
 */
function calcolaConto(booking, extraRows = []) {
  const cfg         = loadBillSettings();
  const notti       = nights(booking.s, booking.e);
  let letti       = parseBedString(booking.d);
  // Fallback: usa i campi numerici del booking se parseBedString non trova nulla
  if (!letti.m && !letti.ms && !letti.s && !letti.c) {
    if (booking.matrimonialiUS > 0) letti.ms = booking.matrimonialiUS;
    else if (booking.matrimoniali > 0) letti.m = booking.matrimoniali;
    if (booking.singoli > 0 && letti.m > 0) letti.s = booking.singoli;
    else if (booking.singoli > 0) letti.s = booking.singoli;
  }
  const nomeCliente = booking.n.toLowerCase();

  // 1. Tariffa/notte: override camera > disposizione letti
  const ov = cfg.tariffeCamere?.[booking.r];
  const tariffaNotte = (ov?.giornaliera > 0)
    ? ov.giornaliera
    : tariffaDaDisposizione(letti, cfg.tariffe);

  // 2. Stagionalità
  const molt       = calcolaMoltiplicatoreStagionale(booking.s, booking.e, cfg.stagioni);
  const prezzoN    = parseFloat((tariffaNotte * molt).toFixed(2));
  const subtBase   = parseFloat((prezzoN * notti).toFixed(2));

  // 3. Convenzione
  let convenzione = null;
  for (const cv of (cfg.convenzioni||[])) {
    if (nomeCliente.includes(cv.nome.toLowerCase())) { convenzione = cv; break; }
  }

  // 4. Sconto durata (solo se no convenzione)
  let scontoDurata = 0;
  if (!convenzione) {
    const ord = [...(cfg.scontiDurata||[])].sort((a,b)=>b.soglia-a.soglia);
    for (const sd of ord) { if (notti>=sd.soglia) { scontoDurata=sd.sconto; break; } }
  }

  // ── Righe ──
  const righe = [];
  const stagLabel = molt !== 1 ? ` ×${molt.toFixed(2)}` : '';
  const desc = descrizioneTariffa(letti, cfg.tariffe, cfg, booking.r);

  righe.push({
    label: `Pernottamento — ${desc}${stagLabel}`,
    qty: notti, unitPrice: prezzoN, total: subtBase,
    tipo: 'base',
    badge: molt !== 1 ? `stagione ×${molt.toFixed(2)}` : null
  });

  let subtotale = subtBase;

  if (convenzione) {
    const sc = parseFloat((subtotale * convenzione.sconto / 100).toFixed(2));
    righe.push({ label:`Conv. "${convenzione.nome}" -${convenzione.sconto}%`, qty:null, unitPrice:null, total:-sc, tipo:'sconto', badge:'conv' });
    subtotale -= sc;
  } else if (scontoDurata > 0) {
    const sc = parseFloat((subtotale * scontoDurata / 100).toFixed(2));
    righe.push({ label:`Sconto lunga durata (${notti} notti) -${scontoDurata}%`, qty:null, unitPrice:null, total:-sc, tipo:'sconto', badge:'lunga' });
    subtotale -= sc;
  }

  for (const ex of extraRows) {
    const tot = parseFloat((ex.qty * ex.unitPrice).toFixed(2));
    righe.push({ label:ex.label, qty:ex.qty, unitPrice:ex.unitPrice, total:tot, tipo:'extra' });
    subtotale += tot;
  }

  return { righe, totale:parseFloat(subtotale.toFixed(2)), notti, molt, convenzione, scontoDurata, tariffaNotte, letti, cfg, aliquotaIVA: cfg.aliquotaIVA||10 };
}

/**
 * Calcolo conto appartamento.
 * Usa override mensile o giornaliero per camera.
 */
function calcolaContoAppart(b, room, extras) {
  const cfg    = loadBillSettings();
  const notti  = nights(b.s, b.e);
  const ov     = cfg.tariffeCamere?.[room.id] || {};
  const modo   = getAppartMode(b.id, notti); // 'giornaliera' | 'mensile' | 'standard'

  // Modalità standard: usa disposizione letti come un normale albergo
  if (modo === 'standard') {
    return calcolaConto(b, extras); // delega al calcolo albergo
  }

  const usaM   = modo === 'mensile';
  const canone = usaM ? (ov.mensile||0) : (ov.giornaliera||0);
  const qty    = usaM ? parseFloat((notti/30).toFixed(2)) : notti;
  const unit   = usaM ? 'mese' : 'notte';
  const total  = parseFloat((canone * qty).toFixed(2));

  const righe = [{
    label:`${room.name} — ${qty} ${unit} × ${canone}€`,
    qty, unitPrice:canone, total, tipo:'base', badge:'appart'
  }];
  let totale = total;

  // Sconto lungo periodo
  if (!usaM && total > 0) {
    const sconti = (cfg.scontiLungoPeriodo || [])
      .filter(s => s.minNotti > 0 && s.percSconto > 0)
      .sort((a,b) => b.minNotti - a.minNotti);
    const sconto = sconti.find(s => notti >= s.minNotti);
    if (sconto) {
      const imp = parseFloat((total * sconto.percSconto / 100).toFixed(2));
      const lbl = sconto.label || `Sconto lungo periodo (${sconto.percSconto}% ≥${sconto.minNotti}gg)`;
      righe.push({ label:lbl, qty:null, unitPrice:null, total:-imp, tipo:'sconto', badge:'lunga' });
      totale -= imp;
    }
  }

  for (const ex of extras) {
    const t = parseFloat((ex.qty * ex.unitPrice).toFixed(2));
    righe.push({ label:ex.label, qty:ex.qty, unitPrice:ex.unitPrice, total:t, tipo:'extra' });
    totale += t;
  }
  return { righe, totale:parseFloat(totale.toFixed(2)), aliquotaIVA: loadBillSettings().aliquotaIVA||10 };
}

// getAppartMode + toggleAppartMode → vedi billing DB layer

// ─────────────────────────────────────────────────────────────────
// RENDER — TAB CONTO NEL DRAWER
// ─────────────────────────────────────────────────────────────────

// getContoOverrides + setContoOverrides → vedi billing DB layer

// Apri editor inline per una riga del conto (override)
function editRigaConto(bid, idx) {
  const b = bookings.find(x=>x.id===bid); if(!b) return;
  const room = ROOMS.find(r=>r.id===b.r);
  const isA  = room?.g==='Appartamenti';
  const ext  = getExtraForBooking(bid);
  const base = isA ? calcolaContoAppart(b,room,ext) : calcolaConto(b,ext);
  const ovs  = getContoOverrides(bid) || base.righe.map(r=>({...r}));
  const r    = ovs[idx]; if(!r) return;

  const overlay = document.createElement('div');
  overlay.style.cssText='position:fixed;inset:0;background:rgba(0,0,0,.5);z-index:9999;display:flex;align-items:flex-end;justify-content:center';
  overlay.innerHTML=`
    <div style="background:var(--surface);border-radius:16px 16px 0 0;padding:20px;width:100%;max-width:460px;box-sizing:border-box">
      <div style="font-size:13px;font-weight:600;margin-bottom:14px;display:flex;justify-content:space-between">
        <span>✏️ Modifica voce</span>
        <button onclick="this.closest('[style*=fixed]').remove()" style="border:none;background:none;font-size:18px;cursor:pointer;color:var(--text2)">✕</button>
      </div>
      <div style="display:flex;flex-direction:column;gap:10px">
        <label style="font-size:11px;color:var(--text3)">Descrizione</label>
        <input id="_ovLabel" value="${r.label.replace(/"/g,'&quot;')}" style="border:1px solid var(--border);border-radius:6px;padding:8px;font-size:13px;width:100%;box-sizing:border-box">
        <div style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:8px">
          <div>
            <label style="font-size:11px;color:var(--text3)">Qtà</label>
            <input id="_ovQty" type="number" step="0.01" value="${r.qty??''}" placeholder="—" style="border:1px solid var(--border);border-radius:6px;padding:8px;font-size:13px;width:100%;box-sizing:border-box">
          </div>
          <div>
            <label style="font-size:11px;color:var(--text3)">Prezzo unitario €</label>
            <input id="_ovPrice" type="number" step="0.01" value="${r.unitPrice??''}" placeholder="—" style="border:1px solid var(--border);border-radius:6px;padding:8px;font-size:13px;width:100%;box-sizing:border-box">
          </div>
          <div>
            <label style="font-size:11px;color:var(--text3)">Totale €</label>
            <input id="_ovTotal" type="number" step="0.01" value="${r.total.toFixed(2)}" style="border:1px solid var(--border);border-radius:6px;padding:8px;font-size:13px;width:100%;box-sizing:border-box">
          </div>
        </div>
        <div style="font-size:10px;color:var(--text3)">💡 Modifica il totale direttamente oppure qtà × prezzo. Il totale ha la precedenza se compilato.</div>
        <div style="display:flex;gap:8px;margin-top:4px">
          <button onclick="salvaOverrideRiga(${bid},${idx},this)" class="btn primary" style="flex:1;justify-content:center">✓ Applica</button>
          <button onclick="rimuoviOverrideRiga(${bid},${idx},this)" class="btn" style="flex:1;justify-content:center;color:var(--danger)">✕ Rimuovi voce</button>
          <button onclick="resetOverride(${bid},this)" class="btn" style="flex:1;justify-content:center">↺ Reset</button>
        </div>
      </div>
    </div>`;
  document.body.appendChild(overlay);
  overlay.addEventListener('click', e=>{ if(e.target===overlay) overlay.remove(); });
}

function salvaOverrideRiga(bid, idx, btn) {
  const b    = bookings.find(x=>x.id===bid); if(!b) return;
  const room = ROOMS.find(r=>r.id===b.r);
  const isA  = room?.g==='Appartamenti';
  const ext  = getExtraForBooking(bid);
  const base = isA ? calcolaContoAppart(b,room,ext) : calcolaConto(b,ext);
  const ovs  = getContoOverrides(bid) || base.righe.map(r=>({...r}));

  const label    = document.getElementById('_ovLabel')?.value.trim() || ovs[idx].label;
  const qty      = parseFloat(document.getElementById('_ovQty')?.value);
  const price    = parseFloat(document.getElementById('_ovPrice')?.value);
  const totalRaw = parseFloat(document.getElementById('_ovTotal')?.value);
  // Priorità: qty×price se entrambi validi (ignora totalRaw che parte da 0)
  // Solo se qty o price mancano → usa totalRaw, altrimenti valore originale
  let total;
  if (!isNaN(qty) && !isNaN(price)) {
    total = parseFloat((qty * price).toFixed(2));
  } else if (!isNaN(totalRaw) && totalRaw !== 0) {
    total = totalRaw;
  } else {
    total = ovs[idx].total;
  }
  const qtyFinal   = !isNaN(qty)   ? qty   : null;
  const priceFinal = !isNaN(price) ? price : null;

  ovs[idx] = { ...ovs[idx], label, qty:qtyFinal, unitPrice:priceFinal, total: parseFloat(total.toFixed(2)) };
  setContoOverrides(bid, ovs);
  btn.closest('[style*=fixed]').remove();
  refreshBillTab(bid);
}

function rimuoviOverrideRiga(bid, idx, btn) {
  const b    = bookings.find(x=>x.id===bid); if(!b) return;
  const room = ROOMS.find(r=>r.id===b.r);
  const isA  = room?.g==='Appartamenti';
  const ext  = getExtraForBooking(bid);
  const base = isA ? calcolaContoAppart(b,room,ext) : calcolaConto(b,ext);
  const ovs  = getContoOverrides(bid) || base.righe.map(r=>({...r}));
  ovs.splice(idx, 1);
  setContoOverrides(bid, ovs);
  btn.closest('[style*=fixed]').remove();
  refreshBillTab(bid);
}

function resetOverride(bid, btn) {
  setContoOverrides(bid, null);
  btn.closest('[style*=fixed]').remove();
  refreshBillTab(bid);
}

function renderDrawerBill(b) {
  const room     = ROOMS.find(r=>r.id===b.r);
  const isAppart = room?.g === 'Appartamenti';
  const ext      = getExtraForBooking(b.id);
  const base     = isAppart ? calcolaContoAppart(b, room, ext) : calcolaConto(b, ext);
  // Applica override se presenti
  const ovs      = getContoOverrides(b.id);
  // Se ci sono ovs, includi anche gli extras non già presenti negli ovs
  let righe;
  if (ovs) {
    const extrasInBase = base.righe.filter(r => r.tipo === 'extra' || r.tipo === 'sconto');
    const ovsLabels = new Set(ovs.map(r => r.label));
    const extrasExtra = extrasInBase.filter(r => !ovsLabels.has(r.label));
    righe = [...ovs, ...extrasExtra];
  } else {
    righe = base.righe;
  }
  const totale   = parseFloat(righe.reduce((s,r)=>s+r.total,0).toFixed(2));
  const hasOv    = ovs !== null;

  const rigaHtml = (r, idx) => {
    const neg = r.total < 0;
    const tot = `<span style="font-weight:600;${neg?'color:var(--danger)':''}">${neg?'':'+'}${r.total.toFixed(2)}€</span>`;
    return `<div class="bill-row ${r.tipo==='sconto'?'discount':''}" style="cursor:pointer" onclick="editRigaConto(${b.id},${idx})" title="Clicca per modificare">
      <span class="bill-row-label">${r.label}${r.badge?` <span class="rate-badge ${r.badge==='conv'?'conv':r.badge==='appart'?'appart':r.badge.startsWith('stagione')?'stagionale':'lunga'}">${r.badge}</span>`:''} <span style="font-size:9px;color:var(--text3);margin-left:3px">✏️</span></span>
      <span style="color:var(--text3);font-size:11px;white-space:nowrap">${r.qty!=null?r.qty+'×':''}</span>
      <span style="color:var(--text2);font-size:11px;white-space:nowrap">${r.unitPrice!=null?r.unitPrice.toFixed(2)+'€':''}</span>
      ${tot}
    </div>`;
  };

  // Per appartamenti: mostra selettore modalità (giornaliero/mensile/standard)
  let toggleHtml = '';
  if (isAppart) {
    const notti  = nights(b.s, b.e);
    const modo   = getAppartMode(b.id, notti);
    const cfg    = loadBillSettings();
    const ov     = cfg.tariffeCamere?.[b.r]||{};
    const modoLabel = { giornaliera:'Giornaliero', mensile:'Mensile', standard:'Std. letti' };
    const modoNext  = { giornaliera:'mensile', mensile:'standard', standard:'giornaliera' };
    const nextLabel = { giornaliera:`Mensile (${ov.mensile||0}€/mese)`, mensile:`Std. letti (tariffe camera)`, standard:`Giornaliero (${ov.giornaliera||0}€/notte)` };
    const noTariffa = modo!=='standard' && !ov.giornaliera && !ov.mensile;
    toggleHtml = `<div style="font-size:11px;color:var(--text2);margin-bottom:10px;display:flex;align-items:center;gap:8px;flex-wrap:wrap">
      <span>Modo: <b>${modoLabel[modo]||modo}</b></span>
      <button onclick="toggleAppartMode(${b.id})" style="border:1px solid var(--border);background:var(--surface2);padding:3px 8px;border-radius:4px;font-size:11px;cursor:pointer">
        ↕ Usa ${nextLabel[modo]||''}
      </button>
      ${noTariffa?'<span style="color:var(--danger);font-size:10px">⚠ Tariffa non impostata</span>':''}
    </div>`;
  }

  // Extra adder
  const cfg = loadBillSettings();
  const extraOptions = (cfg.extra||[]).map(e=>
    `<option value="${e.id}" data-price="${e.prezzo}" data-unita="${e.unita}">${e.label} ${e.prezzo>0?'('+e.prezzo+'€)':''}</option>`
  ).join('');

  const extraAdder = isAppart
    ? `<div class="extra-adder">
        <input type="text"   id="extraLabel_${b.id}" placeholder="Descrizione" style="flex:2">
        <input type="number" id="extraQty_${b.id}"   value="1" min="0.01" step="0.01" style="width:60px" placeholder="Qty">
        <input type="number" id="extraPrice_${b.id}" placeholder="€" step="0.01" style="width:70px">
        <button class="btn" onclick="addExtraLibero(${b.id})">+</button>
      </div>`
    : `<div class="extra-adder">
        <select id="extraType_${b.id}" onchange="prefillExtraPrice(${b.id})">${extraOptions}</select>
        <input type="number" id="extraQty_${b.id}"   value="1" min="0.01" step="0.01" style="width:55px" placeholder="Qty">
        <input type="number" id="extraPrice_${b.id}" placeholder="€" step="0.01" style="width:70px">
        <button class="btn" onclick="addExtraVoce(${b.id})">+</button>
      </div>`;

  // Consumo elettrico/idrico (solo se extra configurato con unita kwh/mc)
  const consumiHtml = !isAppart ? '' : `
    <div class="conti-section-title" style="margin-top:10px">Consumi (lettura contatori)</div>
    <div style="display:grid;grid-template-columns:1fr 1fr 1fr 1fr;gap:6px;align-items:end;margin-bottom:8px">
      <div><label style="font-size:10px;color:var(--text3);display:block">⚡ Inizio kWh</label><input type="number" id="kwh_start_${b.id}" step="0.1" placeholder="0" style="width:100%;padding:5px;border:1px solid var(--border);border-radius:4px;font-size:12px"></div>
      <div><label style="font-size:10px;color:var(--text3);display:block">⚡ Fine kWh</label><input type="number" id="kwh_end_${b.id}" step="0.1" placeholder="0" style="width:100%;padding:5px;border:1px solid var(--border);border-radius:4px;font-size:12px"></div>
      <div><label style="font-size:10px;color:var(--text3);display:block">💧 Inizio m³</label><input type="number" id="mc_start_${b.id}" step="0.001" placeholder="0" style="width:100%;padding:5px;border:1px solid var(--border);border-radius:4px;font-size:12px"></div>
      <div><label style="font-size:10px;color:var(--text3);display:block">💧 Fine m³</label><input type="number" id="mc_end_${b.id}" step="0.001" placeholder="0" style="width:100%;padding:5px;border:1px solid var(--border);border-radius:4px;font-size:12px"></div>
    </div>
    <button class="btn" style="width:100%;justify-content:center;font-size:11px" onclick="aggiungiConsumi(${b.id})">⚡💧 Aggiungi consumi al conto</button>`;

  // Calcola stato conto UNA VOLTA sola — usato sia nel badge che nel bottone
  const _contiOuter = loadConti();
  const _ceOuter = _contiOuter.find(x => x.bookingId === b.id) || null;
  const _isEmesso = _ceOuter && _ceOuter.status && _ceOuter.status !== 'bozza';

  return `
    <div style="display:flex;justify-content:space-between;align-items:baseline;margin-bottom:2px">
      <span style="font-size:11px;font-weight:600;color:var(--text3);text-transform:uppercase">Conto preventivo${hasOv?' <span style="font-size:9px;color:var(--accent);margin-left:4px">● modificato</span>':''}</span>
      <span style="font-size:20px;font-weight:700;color:var(--accent);font-family:'Playfair Display',serif">${totale.toFixed(2)}€</span>
    </div>
    <div style="font-size:10px;color:var(--text3);text-align:right;margin-bottom:10px">
      IVA ${base.aliquotaIVA||10}% inclusa · imponibile ${(totale/(1+(base.aliquotaIVA||10)/100)).toFixed(2)}€
    </div>
    ${toggleHtml}
    ${righe.map((r,i)=>rigaHtml(r,i)).join('')}
    ${renderExtraRows(b.id)}
    <div class="conti-section-title" style="margin-top:14px">Aggiungi voce</div>
    ${extraAdder}
    ${consumiHtml}
    ${(()=>{
      const _ce = _ceOuter;
      const _si = _ce ? (STATO_CFG[_ce.status]||STATO_CFG.bozza) : null;
      return _si ? `<div style="display:flex;align-items:center;gap:6px;margin:10px 0 6px;padding:8px 10px;background:var(--surface2);border-radius:8px;font-size:12px">
        <span class="stato-pill ${_ce.status}">${_si.icon} ${_si.label}</span>
        <span style="color:var(--text3);flex:1">Salvato · ${_ce.totale?.toFixed(2)}€</span>
        ${_ce.numDoc?`<span style="font-size:10px;color:var(--text3)">📄 ${_ce.numDoc}</span>`:''}
      </div>` : '';
    })()}
    ${(()=>{
      const _pags = (typeof getPagamentiPerBookingSync==='function') ? getPagamentiPerBookingSync(b.id) : [];
      const _ce2 = _ceOuter;
      const _tot  = _ce2 ? (_ce2.totale||0) : 0;
      const _pagato = _pags.reduce((s,p)=>s+p.importo,0);
      const _residuo = Math.max(0,_tot-_pagato);
      if (_pags.length===0 && !_ce2) return '';
      const tipoClass = t => ({'acconto':'pag-tipo-acconto','saldo':'pag-tipo-saldo','extra':'pag-tipo-extra'}[t]||'');
      return `<div class="pagamenti-section">
        <div class="pag-header">
          <span class="pag-title">💳 Pagamenti</span>
          ${_tot>0?`<span class="pag-riepilogo">Tot €${_tot.toFixed(2)} · Pagato €${_pagato.toFixed(2)}${_residuo>0.01?` · <b class="pag-residuo">Residuo €${_residuo.toFixed(2)}</b>`:' <b class=\'pag-saldato\'>✓ Saldato</b>'}</span>`:''}
        </div>
        ${_pags.map(p=>`
          <div class="pag-row">
            <span class="pag-data">${p.data}</span>
            <span class="pag-tipo ${tipoClass(p.tipo)}">${p.tipo}</span>
            <span class="pag-importo">€${p.importo.toFixed(2)}</span>
            <span class="pag-metodo">${p.metodo}</span>
            ${p.conDocumento?'<span class="pag-doc" title="Con documento">📄</span>':''}
            ${p.riferimento?`<span class="pag-rif">${p.riferimento}</span>`:''}
            <button class="pag-del" onclick="eliminaPagamentoUI('${p.id}',${b.id})" title="Elimina">✕</button>
          </div>
        `).join('')}
        <button class="btn pag-add-btn" onclick="apriDialogPagamento('${_ce2?.id||''}',${_residuo.toFixed(2)})">+ Pagamento</button>
      </div>`;
    })()}
    <div style="display:flex;gap:8px;margin-top:10px">
      ${_isEmesso
        ? `<button class="btn" onclick="emettiConto(${b.id})" style="flex:1;justify-content:center;opacity:0.7;">✎ Modifica conto</button>`
        : `<button class="btn primary" onclick="emettiConto(${b.id})" style="flex:1;justify-content:center;">📄 Emetti conto</button>`
      }
      <button class="btn" onclick="apriPdf(${b.id})" style="flex:1;justify-content:center;">👁 PDF</button>
    </div>`;
}

// ─────────────────────────────────────────────────────────────────
// EXTRA
// ─────────────────────────────────────────────────────────────────

// getExtraForBooking + setExtraForBooking → vedi billing DB layer

function renderExtraRows(bid) {
  const extras = getExtraForBooking(bid);
  if (!extras.length) return '';
  return `<div style="margin-top:6px;border-top:1px solid var(--border);padding-top:6px">${extras.map((ex,i)=>`
    <div class="bill-row" style="font-size:12px">
      <span class="bill-row-label">${ex.label}</span>
      <span style="color:var(--text3);font-size:11px">${ex.qty}${ex.unita?'×'+ex.unita:''}</span>
      <span style="color:var(--text2);font-size:11px">${ex.unitPrice.toFixed(2)}€</span>
      <span style="font-weight:600">+${(ex.qty*ex.unitPrice).toFixed(2)}€</span>
      <button onclick="removeExtra(${bid},${i})" style="border:none;background:none;color:var(--danger);cursor:pointer;padding:0 4px;font-size:14px">✕</button>
    </div>`).join('')}</div>`;
}

function prefillExtraPrice(bid) {
  const sel = document.getElementById(`extraType_${bid}`);
  const opt = sel?.selectedOptions?.[0];
  if (opt) {
    const p = parseFloat(opt.dataset.price||0);
    const inp = document.getElementById(`extraPrice_${bid}`);
    if (inp && p > 0) inp.value = p;
  }
}

// Aggiorna il tab Conto in-place senza ricostruire l'intero drawer
// Preserva il tab attivo e lo scroll position del drawer
function refreshBillTab(bid) {
  const b = bookings.find(x=>x.id===bid);
  if (!b) return;
  const tabEl = document.getElementById('drTabBill');
  if (tabEl) {
    const scrollTop = tabEl.scrollTop;
    tabEl.innerHTML = renderDrawerBill(b);
    tabEl.scrollTop = scrollTop;
    // Assicurati che il tab Conto sia visibile
    tabEl.style.display = '';
    document.getElementById('drTabInfo') && (document.getElementById('drTabInfo').style.display = 'none');
    // Aggiorna label tab attivo
    document.querySelectorAll('.dr-bill-tab').forEach(t => {
      t.classList.toggle('active', t.textContent.includes('Conto'));
    });
  } else {
    // Drawer non aperto — ricostruisci normalmente
    if (typeof selBook === 'function') selBook(bid, null);
  }
}

function addExtraVoce(bid) {
  const cfg   = loadBillSettings();
  const sel   = document.getElementById(`extraType_${bid}`);
  const opt   = sel?.selectedOptions?.[0];
  const id    = sel?.value;
  const def   = (cfg.extra||[]).find(e=>e.id===id);
  const qty   = parseFloat(document.getElementById(`extraQty_${bid}`)?.value||1);
  const price = parseFloat(document.getElementById(`extraPrice_${bid}`)?.value||(def?.prezzo||0));
  const label = def?.label || id;
  const unita = def?.unita || '';
  const extras = getExtraForBooking(bid);
  extras.push({ label, qty, unitPrice:price, unita });
  setExtraForBooking(bid, extras);
  refreshBillTab(bid);
}

function addExtraLibero(bid) {
  const label = document.getElementById(`extraLabel_${bid}`)?.value?.trim();
  const qty   = parseFloat(document.getElementById(`extraQty_${bid}`)?.value||1);
  const price = parseFloat(document.getElementById(`extraPrice_${bid}`)?.value||0);
  if (!label) return;
  const extras = getExtraForBooking(bid);
  extras.push({ label, qty, unitPrice:price, unita:'' });
  setExtraForBooking(bid, extras);
  refreshBillTab(bid);
}

function aggiungiConsumi(bid) {
  const cfg    = loadBillSettings();
  const extras = getExtraForBooking(bid);
  const kwhS   = parseFloat(document.getElementById(`kwh_start_${bid}`)?.value||0);
  const kwhE   = parseFloat(document.getElementById(`kwh_end_${bid}`)?.value||0);
  const mcS    = parseFloat(document.getElementById(`mc_start_${bid}`)?.value||0);
  const mcE    = parseFloat(document.getElementById(`mc_end_${bid}`)?.value||0);

  const defLuce  = (cfg.extra||[]).find(e=>e.id==='luce');
  const defAcqua = (cfg.extra||[]).find(e=>e.id==='acqua');

  if (kwhE > kwhS) {
    const diff  = parseFloat((kwhE-kwhS).toFixed(2));
    const price = defLuce?.prezzo||0;
    extras.push({ label:'⚡ Consumo elettrico', qty:diff, unitPrice:price, unita:'kWh' });
  }
  if (mcE > mcS) {
    const diff  = parseFloat((mcE-mcS).toFixed(3));
    const price = defAcqua?.prezzo||0;
    extras.push({ label:'💧 Consumo idrico', qty:diff, unitPrice:price, unita:'m³' });
  }
  setExtraForBooking(bid, extras);
  refreshBillTab(bid);
}

function removeExtra(bid, idx) {
  const extras = getExtraForBooking(bid);
  extras.splice(idx, 1);
  setExtraForBooking(bid, extras);
  refreshBillTab(bid);
}

// ── Tab switcher nel drawer ──
function drTab(el, showId) {
  el.closest('#drbody').querySelectorAll('.dr-bill-tab').forEach(t=>t.classList.remove('active'));
  el.classList.add('active');
  ['drTabInfo','drTabBill','drTabCI'].forEach(id=>{
    const d=document.getElementById(id);
    if(d) d.style.display = id===showId?'':'none';
  });
}

// ─────────────────────────────────────────────────────────────────
// EMETTI CONTO + PDF
// ─────────────────────────────────────────────────────────────────

function getContoEffettivo(bid) {
  const b    = bookings.find(x=>x.id===bid); if(!b) return null;
  const room = ROOMS.find(r=>r.id===b.r);
  const isA  = room?.g==='Appartamenti';
  const ext  = getExtraForBooking(bid);
  const base = isA ? calcolaContoAppart(b,room,ext) : calcolaConto(b,ext);
  const ovs  = getContoOverrides(bid);
  if (ovs) {
    // Gli extras (sconto, voci libere) vengono da ext e sono già in base.righe
    // ma NON sono in ovs (che contiene solo le righe modificate manualmente).
    // Li aggiungiamo separatamente per non perderli nel totale.
    const extrasInBase = base.righe.filter(r => r.tipo === 'extra' || r.tipo === 'sconto');
    // Evita duplicati: non aggiungere extras già presenti in ovs
    const ovsLabels = new Set(ovs.map(r => r.label));
    const extrasExtra = extrasInBase.filter(r => !ovsLabels.has(r.label));
    const righeFinali = [...ovs, ...extrasExtra];
    const totale = parseFloat(righeFinali.reduce((s,r) => s + r.total, 0).toFixed(2));
    return { ...base, righe: righeFinali, totale };
  }
  return base;
}

function emettiConto(bid) {
  const b    = bookings.find(x=>x.id===bid); if(!b) return;
  const room = ROOMS.find(r=>r.id===b.r);
  const isA  = room?.g==='Appartamenti';
  const conto= getContoEffettivo(bid);
  const conti= loadConti();
  const idx  = conti.findIndex(c=>c.bookingId===bid);
  const ora  = new Date().toISOString();
  const esistente = idx >= 0 ? conti[idx] : null;
  const doc  = {
    id:         esistente?.id || ('C'+Date.now()),
    bookingId:  bid,
    groupId:    esistente?.groupId || null,
    nome:       b.n,
    camera:     room?.name||roomName(b.r),
    checkin:    b.s.toISOString(),
    checkout:   b.e.toISOString(),
    righe:      conto.righe,
    totale:     conto.totale,
    status:     esistente?.status || 'emesso',
    emessoIl:   esistente?.emessoIl || ora,
    fatturatoIl:  esistente?.fatturatoIl  || null,
    tipoDoc:      esistente?.tipoDoc      || null,
    numDoc:       esistente?.numDoc       || null,
    pagatoIl:     esistente?.pagatoIl     || null,
    modalitaPag:  esistente?.modalitaPag  || null,
    isAppart:   isA,
    ts:         ora
  };
  if(idx>=0) conti[idx]=doc; else conti.unshift(doc);
  saveConti(conti);
  apriPdf(bid);
  showToast('✓ Conto salvato e emesso','success');
}

let _currentPdfBid = null;
function apriPdf(bid) {
  _currentPdfBid = bid;
  const b=bookings.find(x=>x.id===bid); if(!b) return;
  const room=ROOMS.find(r=>r.id===b.r);
  const isA=room?.g==='Appartamenti';
  const conto=getContoEffettivo(bid);
  const cfg=loadBillSettings();
  document.getElementById('pdfTitle').textContent=`Conto — ${b.n}`;
  document.getElementById('printDoc').innerHTML=buildPdfDoc(b,room,conto,cfg,isA);
  document.getElementById('pdfOverlay').classList.add('open');
}

function buildPdfDoc(b, room, conto, cfg, isAppart) {
  const oggi = new Date().toLocaleDateString('it-IT');
  const tipo = isAppart ? 'Rendiconto Affitto' : 'Conto Soggiorno';
  const aliqIVA     = conto.aliquotaIVA || cfg.aliquotaIVA || 10;
  const imponibilePdf = parseFloat((conto.totale / (1 + aliqIVA/100)).toFixed(2));
  const ivaPdf        = parseFloat((conto.totale - imponibilePdf).toFixed(2));
  const righeHtml = conto.righe.filter(r=>r.total!==0).map(r=>`
    <tr>
      <td>${r.label}</td>
      <td style="text-align:center">${r.qty!=null?r.qty:'—'}</td>
      <td style="text-align:right">${r.unitPrice!=null?r.unitPrice.toFixed(2)+'€':'—'}</td>
      <td style="text-align:right;color:${r.total<0?'#c0392b':'inherit'}">${r.total>=0?'+':''}${r.total.toFixed(2)}€</td>
    </tr>`).join('');
  return `
    <div class="doc-header">
      <div class="doc-hotel-name">${cfg.hotelName}</div>
      ${cfg.hotelAddress?`<div class="doc-hotel-sub">${cfg.hotelAddress}</div>`:''}
      ${cfg.hotelTel?`<div class="doc-hotel-sub">Tel: ${cfg.hotelTel}</div>`:''}
      <div class="doc-type">${tipo}</div>
    </div>
    <div class="doc-section">
      <div class="doc-section-title">Dati cliente</div>
      <div class="doc-info-grid">
        <div class="doc-info-item"><label>Nome</label><span>${b.n}</span></div>
        <div class="doc-info-item"><label>Camera / Unità</label><span>${room?.name||roomName(b.r)}</span></div>
        <div class="doc-info-item"><label>Check-in</label><span>${fmt(b.s)}</span></div>
        <div class="doc-info-item"><label>Check-out</label><span>${fmt(b.e)}</span></div>
        <div class="doc-info-item"><label>${isAppart?'Periodo':'Notti'}</label><span>${isAppart?(nights(b.s,b.e)/30).toFixed(1)+' mesi':nights(b.s,b.e)+' notti'}</span></div>
        ${!isAppart?`<div class="doc-info-item"><label>Disposizione</label><span>${b.d||'—'}</span></div>`:''}
      </div>
    </div>
    <div class="doc-section">
      <div class="doc-section-title">Dettaglio</div>
      <table class="doc-table">
        <thead><tr><th>Descrizione</th><th style="text-align:center">Qtà</th><th style="text-align:right">Prezzo</th><th style="text-align:right">Importo</th></tr></thead>
        <tbody>${righeHtml}</tbody>
        <tfoot>
          <tr style="font-size:11px;color:#666"><td colspan="3">Imponibile (IVA ${aliqIVA}% esclusa)</td><td style="text-align:right">${imponibilePdf.toFixed(2)} €</td></tr>
          <tr style="font-size:11px;color:#666"><td colspan="3">IVA ${aliqIVA}%</td><td style="text-align:right">${ivaPdf.toFixed(2)} €</td></tr>
          <tr class="doc-total-row"><td colspan="3"><strong>TOTALE IVA inclusa</strong></td><td><strong>${conto.totale.toFixed(2)} €</strong></td></tr>
        </tfoot>
      </table>
    </div>
    <div class="doc-footer">${cfg.hotelName} · Documento del ${oggi}${cfg.hotelTel?' · Tel: '+cfg.hotelTel:''}</div>`;
}

function closePdfOverlay() { document.getElementById('pdfOverlay').classList.remove('open'); }

function shareDocWhatsApp() {
  const t = document.getElementById('printDoc').innerText.split('\n').filter(l=>l.trim()).slice(0,25).join('\n');
  window.open(`https://wa.me/?text=${encodeURIComponent(t)}`,'_blank');
}
function shareDocEmail() {
  const s=encodeURIComponent('Conto — '+document.getElementById('pdfTitle').textContent);
  const b=encodeURIComponent(document.getElementById('printDoc').innerText);
  window.open(`mailto:?subject=${s}&body=${b}`,'_blank');
}

function esportaXML() {
  // Tutto in un unico try — le variabili devono essere nello stesso scope del codice che le usa
  try {
  const bid   = _currentPdfBid;
  if (!bid) { showToast('Apri prima un conto', 'error'); return; }
  const b     = bookings.find(x=>x.id===bid); if(!b) { showToast('Prenotazione non trovata','error'); return; }
  const room  = ROOMS.find(r=>r.id===b.r);
  const cfg   = loadBillSettings();
  const conto = getContoEffettivo(bid);
  if (!conto) { showToast('Impossibile calcolare il conto','error'); return; }
  const oggi  = new Date();

  // Formatta data ISO
  const fmtD = d => d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0')+'-'+String(d.getDate()).padStart(2,'0');

  // Escape XML + rimuove caratteri non Latin1 (schema FatturaPA accetta solo Latin)
  // Sostituisce caratteri speciali comuni con equivalenti ASCII
  const toLatin = s => String(s||'')
    .replace(/[·•]/g, '-')
    .replace(/[→⇒►]/g, '->')
    .replace(/[×x✕]/g, 'x')
    .replace(/[àáâã]/g, 'a').replace(/[èéêë]/g, 'e')
    .replace(/[ìíîï]/g, 'i').replace(/[òóôõ]/g, 'o')
    .replace(/[ùúûü]/g, 'u').replace(/[ÀÁ]/g, 'A')
    .replace(/[ÈÉ]/g, 'E').replace(/[ÙÚ]/g, 'U')
    .replace(/[^\u0000-\u00FF]/g, '?'); // tutto il resto → ?
  const esc = s => toLatin(s)
    .replace(/&/g,'&amp;').replace(/</g,'&lt;')
    .replace(/>/g,'&gt;').replace(/"/g,'&quot;');

  // Pulisce un campo per uso in elementi stringa FatturaPA (no caratteri di controllo)
  const clean = s => esc(s).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g,'').trim();

  // P.IVA hotel: letta da impostazioni, default placeholder valido per schema
  const pivaHotel = clean(cfg.pivaHotel || cfg.piva || '00000000000');
  const hotelAddr = clean(cfg.hotelAddress || 'Via da definire');
  const hotelCAP  = clean(cfg.hotelCAP    || '00000');
  const hotelProv = clean(cfg.hotelProv   || 'XX');

  // Split nome cliente in nome/cognome (primo token = nome, resto = cognome)
  // Se è una sola parola o un'azienda usiamo Denominazione
  const nomeParti = (b.n||'Cliente').trim().split(/\s+/);
  const usaDenominazione = nomeParti.length === 1 || b.n.includes('srl') || b.n.includes('spa') || b.n.includes('snc') || b.n.includes('SS') || b.n.length > 30;
  const clienteAnag = usaDenominazione
    ? `<Denominazione>${clean(b.n)}</Denominazione>`
    : `<Nome>${clean(nomeParti[0])}</Nome><Cognome>${clean(nomeParti.slice(1).join(' '))}</Cognome>`;

  // Numero fattura progressivo (anno + timestamp per unicità)
  // ProgressivoInvio: max 10 char, solo lettere e cifre (schema FatturaPA)
  const numFattura = String(oggi.getFullYear()).slice(-2)
    + String(oggi.getMonth()+1).padStart(2,'0')
    + String(oggi.getDate()).padStart(2,'0')
    + String(bid).slice(-2); // es. "260315" + "01" = 8 char

  // Causale: solo ASCII, max 200 char
  const causale = clean('Soggiorno ' + b.n + ' cam. ' + (room?.name||roomName(b.r)) + ' ' + fmtD(b.s) + ' - ' + fmtD(b.e)).slice(0, 200);

  // aliqXML deve essere definita PRIMA del map che la usa
  const aliqXML    = parseFloat(conto.aliquotaIVA || cfg.aliquotaIVA || 10);
  const aliqFactor = 1 + aliqXML / 100;

  // ── Linee dettaglio ──
  // Le righe di tipo 'sconto' (total < 0) NON vanno come DettaglioLinee separate
  // ma come ScontoMaggiorazione sulla riga base precedente.
  // Strategia: raggruppa ogni riga base con gli sconti che la seguono.
  // Sconti globali (convenzione/durata) → ScontoMaggiorazione sulla prima riga base.


  // Normalizza la descrizione per XML FatturaPA:
  // descrizioni generiche, senza numero camera
  const normalizzaDescXML = label => {
    const l = (label||'').toLowerCase();
    // Pernottamento / camera
    if (l.includes('pernottamento') || l.includes('camera') || l.includes('matrim') || l.includes('singol') || l.includes('nott')) {
      // Determina tipo camera dalla disposizione del booking
      const disp = (b.d||'').toLowerCase();
      const ms   = disp.match(/(\d+)\s*m\/s|ms/);
      const m    = disp.match(/(\d+)\s*m\b/);
      const s    = disp.match(/(\d+)\s*s\b/);
      const nS   = s ? parseInt(s[1]) : 0;
      if (ms)                      return 'Camera matrimoniale uso singolo';
      if (m || nS >= 2)            return 'Camera doppia/matrimoniale';
      return 'Camera singola';
    }
    // Convenzione / sconto → non appare come riga (già gestito come ScontoMaggiorazione)
    if (l.includes('conv') || l.includes('sconto'))   return null;
    // Colazione
    if (l.includes('colazion'))                        return 'Colazione';
    // Pranzo
    if (l.includes('pranzo'))                          return 'Pranzo a prezzo fisso';
    // Cena
    if (l.includes('cena'))                            return 'Cena a prezzo fisso';
    // Piscina
    if (l.includes('piscin'))                          return 'Accesso piscina';
    // Pulizie extra
    if (l.includes('puliz'))                           return 'Servizio pulizie extra';
    // Cambio lenzuola
    if (l.includes('lenzuol') || l.includes('cambio')) return 'Cambio biancheria';
    // Consumo elettrico
    if (l.includes('elettr') || l.includes('kwh'))     return 'Consumo elettrico';
    // Consumo idrico
    if (l.includes('idric') || l.includes('acqua') || l.includes('mc')) return 'Consumo idrico';
    // Qualsiasi altro extra: usa label pulita ma senza numero camera
    return clean(label.replace(/cam\.?\s*\d+|camera\s*\d+|stanza\s*\d+/gi, '').trim()).slice(0,100) || 'Servizio';
  };
  const righeBase   = conto.righe.filter(r => r.total > 0);
  const righeSconto = conto.righe.filter(r => r.total < 0);

  // Associa gli sconti alla prima riga base (comportamento standard per sconti globali)
  let lineaNum = 0;
  const linee = righeBase.map(r => {
    lineaNum++;
    const desc      = normalizzaDescXML(r.label) || clean(r.label).slice(0,100);
    const totIvaInc = parseFloat(r.total.toFixed(2));
    const totImpon  = parseFloat((totIvaInc / aliqFactor).toFixed(2));
    const qty       = r.qty != null ? parseFloat(r.qty).toFixed(2) : null;
    const prezzoUnit= r.unitPrice != null
      ? parseFloat((Math.abs(r.unitPrice) / aliqFactor).toFixed(2))
      : parseFloat((totImpon / (qty ? parseFloat(qty) : 1)).toFixed(2));

    // Gli sconti globali li attacchiamo solo alla prima riga base (lineaNum === 1)
    const scontiTag = lineaNum === 1 ? righeSconto.map(s => {
      const importoSconto = parseFloat((Math.abs(s.total) / aliqFactor).toFixed(2));
      // Determina se è una percentuale o un importo fisso
      const percTag = s.tipo === 'sconto' && s.label && s.label.match(/(\d+)%/)
        ? `<Percentuale>${s.label.match(/(\d+)%/)[1]}</Percentuale>`
        : `<Importo>${importoSconto.toFixed(2)}</Importo>`;
      return `
        <ScontoMaggiorazione>
          <Tipo>SC</Tipo>
          ${percTag}
          <Importo>${importoSconto.toFixed(2)}</Importo>
        </ScontoMaggiorazione>`;
    }).join('') : '';

    // PrezzoTotale = prezzo lordo - sconti applicati a questa riga
    const scontiImporto = lineaNum === 1
      ? parseFloat(righeSconto.reduce((s,r) => s + Math.abs(r.total) / aliqFactor, 0).toFixed(2))
      : 0;
    const prezzoTotale = parseFloat((totImpon - scontiImporto).toFixed(2));

    return `
    <DettaglioLinee>
      <NumeroLinea>${lineaNum}</NumeroLinea>
      <Descrizione>${desc||'Servizio'}</Descrizione>
      ${qty ? '<Quantita>'+qty+'</Quantita>' : ''}
      <PrezzoUnitario>${prezzoUnit.toFixed(2)}</PrezzoUnitario>
      ${scontiTag}
      <PrezzoTotale>${prezzoTotale.toFixed(2)}</PrezzoTotale>
      <AliquotaIVA>${aliqXML.toFixed(2)}</AliquotaIVA>
    </DettaglioLinee>`;
  }).join('');

  // DatiRiepilogo: ImponibileImporto = somma netta (righe positive - sconti) / aliqFactor
  const totaleNetto  = conto.righe.reduce((s,r) => s + r.total, 0);
  const imponibile   = parseFloat((totaleNetto / aliqFactor).toFixed(2));
  const iva          = parseFloat((imponibile * aliqXML / 100).toFixed(2));
  const totDocumento = parseFloat((imponibile + iva).toFixed(2));

  const xml = `<?xml version="1.0" encoding="UTF-8"?>
<p:FatturaElettronica versione="FPR12"
  xmlns:p="http://ivaservizi.agenziaentrate.gov.it/docs/xsd/fatture/v1.2"
  xmlns:ds="http://www.w3.org/2000/09/xmldsig#"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">

  <FatturaElettronicaHeader>
    <DatiTrasmissione>
      <IdTrasmittente>
        <IdPaese>IT</IdPaese>
        <IdCodice>${pivaHotel}</IdCodice>
      </IdTrasmittente>
      <ProgressivoInvio>${numFattura}</ProgressivoInvio>
      <FormatoTrasmissione>FPR12</FormatoTrasmissione>
      <CodiceDestinatario>0000000</CodiceDestinatario>
    </DatiTrasmissione>

    <CedentePrestatore>
      <DatiAnagrafici>
        <IdFiscaleIVA>
          <IdPaese>IT</IdPaese>
          <IdCodice>${pivaHotel}</IdCodice>
        </IdFiscaleIVA>
        <Anagrafica>
          <Denominazione>${clean(cfg.hotelName)}</Denominazione>
        </Anagrafica>
        <RegimeFiscale>RF01</RegimeFiscale>
      </DatiAnagrafici>
      <Sede>
        <Indirizzo>${hotelAddr}</Indirizzo>
        <CAP>${hotelCAP}</CAP>
        <Comune>${clean(cfg.hotelComune||'Da definire')}</Comune>
        <Provincia>${hotelProv}</Provincia>
        <Nazione>IT</Nazione>
      </Sede>
    </CedentePrestatore>

    <CessionarioCommittente>
      <DatiAnagrafici>
        <CodiceFiscale>RSSMRA00A00H501U</CodiceFiscale>
        <Anagrafica>
          ${clienteAnag}
        </Anagrafica>
      </DatiAnagrafici>
      <Sede>
        <Indirizzo>Da definire</Indirizzo>
        <CAP>00000</CAP>
        <Comune>Da definire</Comune>
        <Nazione>IT</Nazione>
      </Sede>
    </CessionarioCommittente>
  </FatturaElettronicaHeader>

  <FatturaElettronicaBody>
    <DatiGenerali>
      <DatiGeneraliDocumento>
        <TipoDocumento>TD01</TipoDocumento>
        <Divisa>EUR</Divisa>
        <Data>${fmtD(oggi)}</Data>
        <Numero>${numFattura}</Numero>
        <ImportoTotaleDocumento>${totDocumento.toFixed(2)}</ImportoTotaleDocumento>
        <Causale>${causale}</Causale>
      </DatiGeneraliDocumento>
    </DatiGenerali>

    <DatiBeniServizi>
      ${linee}
      <DatiRiepilogo>
        <AliquotaIVA>${aliqXML.toFixed(2)}</AliquotaIVA>
        <ImponibileImporto>${imponibile.toFixed(2)}</ImponibileImporto>
        <Imposta>${iva.toFixed(2)}</Imposta>
        <EsigibilitaIVA>I</EsigibilitaIVA>
      </DatiRiepilogo>
    </DatiBeniServizi>

    <DatiPagamento>
      <CondizioniPagamento>TP02</CondizioniPagamento>
      <DettaglioPagamento>
        <ModalitaPagamento>MP01</ModalitaPagamento>
        <ImportoPagamento>${totDocumento.toFixed(2)}</ImportoPagamento>
      </DettaglioPagamento>
    </DatiPagamento>
  </FatturaElettronicaBody>
</p:FatturaElettronica>`;

  // ── Mostra overlay con link download (dentro il try, stessa scope delle variabili) ──
  const nomeFile = 'fattura_' + b.n.replace(/\s+/g,'_') + '_' + fmtD(oggi) + '.xml';
  const blob     = new Blob([xml], { type:'application/xml;charset=utf-8' });
  const url      = URL.createObjectURL(blob);

  const ov = document.createElement('div');
  ov.id    = 'xmlDlOverlay';
  ov.style.cssText = 'position:fixed;inset:0;background:rgba(0,0,0,.6);z-index:9999;display:flex;align-items:center;justify-content:center;padding:20px;box-sizing:border-box';
  ov.innerHTML = `
    <div style="background:var(--surface);border-radius:16px;padding:24px;max-width:360px;width:100%;text-align:center;box-shadow:0 8px 32px rgba(0,0,0,.3)">
      <div style="font-size:36px;margin-bottom:10px">📄</div>
      <div style="font-weight:700;font-size:15px;margin-bottom:4px">XML FatturaPA pronto</div>
      <div style="font-size:11px;color:var(--text3);margin-bottom:20px;word-break:break-all">${nomeFile}</div>
      <a id="xmlDownloadLink" href="${url}" download="${nomeFile}"
         style="display:block;background:var(--accent);color:#fff;padding:14px 20px;border-radius:10px;text-decoration:none;font-weight:700;font-size:15px;margin-bottom:10px">
        ⬇ Scarica XML
      </a>
      <button onclick="document.getElementById('xmlDlOverlay').remove()"
        style="background:none;border:1px solid var(--border);padding:10px 20px;border-radius:8px;cursor:pointer;font-size:13px;color:var(--text2);width:100%">
        Chiudi
      </button>
    </div>`;
  document.body.appendChild(ov);
  ov.addEventListener('click', e => { if(e.target===ov) ov.remove(); });
  // Pulizia URL al click del link
  document.getElementById('xmlDownloadLink').addEventListener('click', () => {
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  });

  } catch(err) {
    showToast('Errore XML: ' + err.message, 'error');
    console.error('[esportaXML]', err);
  }
}

// ─────────────────────────────────────────────────────────────────
// PAGINA CONTI
// ─────────────────────────────────────────────────────────────────

let _contiTab = 'lista';

function openConti()  { document.getElementById('contiScreen').classList.add('open'); renderContiTab(_contiTab); }
function closeConti() { document.getElementById('contiScreen').classList.remove('open'); }

function switchContiTab(tab) {
  _contiTab = tab;
  document.querySelectorAll('.conti-tab').forEach(t=>{
    t.classList.toggle('active', t.textContent.includes(
      tab==='lista'?'Conti':tab==='gruppo'?'Gruppo':tab==='appart'?'Appartamenti':'Listino'
    ));
  });
  renderContiTab(tab);
}
function renderContiTab(tab) {
  const body = document.getElementById('contiBody');
  if (tab==='lista')   body.innerHTML = renderContiLista();
  if (tab==='gruppo')  { body.innerHTML = renderContoGruppo(); bindContoGruppo(); }
  if (tab==='appart')  body.innerHTML = renderContiAppart();
  if (tab==='tariffe') body.innerHTML = renderListino();
}

// Labels e icone stati
const STATO_CFG = {
  bozza:     { icon:'📝', label:'Bozza',     next:'emesso',    nextLabel:'Segna Emesso' },
  emesso:    { icon:'📤', label:'Emesso',    next:'fatturato', nextLabel:'Segna Fatturato/Scontrino' },
  fatturato: { icon:'🧾', label:'Fatturato', next:'pagato',    nextLabel:'Segna Pagato' },
  pagato:    { icon:'✅', label:'Pagato',    next:null,        nextLabel:null }
};

function renderContiLista() {
  const conti = loadConti();
  if (!conti.length) return `<div class="empty" style="padding:40px 0">
    <div class="emptyicon">📋</div>
    <div style="font-size:12px;color:var(--text3)">Nessun conto emesso.<br>Apri una prenotazione → tab 💶 Conto.</div>
  </div>`;

  const ordine = ['pagato','fatturato','emesso','bozza'];
  const titoli = { pagato:'✅ Pagati', fatturato:'🧾 Fatturati', emesso:'📤 Emessi', bozza:'📝 Bozze' };
  const byS = { emesso:[], pagato:[], bozza:[], fatturato:[] };
  conti.forEach(c => (byS[c.status] || byS.bozza).push(c));

  return ordine.map(stato => {
    const lista = byS[stato];
    if (!lista.length) return '';
    const totSez = lista.reduce((s,c)=>s+(c.totale||0),0);
    return `<div class="conti-section">
      <div class="conti-section-title" style="display:flex;justify-content:space-between">
        <span>${titoli[stato]} (${lista.length})</span>
        <span style="font-weight:600;color:var(--accent)">${totSez.toFixed(2)}€</span>
      </div>
      ${lista.map(c => {
        const ci  = new Date(c.checkin), co = new Date(c.checkout);
        const dot = pastello(bookings.find(b=>b.id===c.bookingId)?.c || '#ccc');
        const cfg = STATO_CFG[c.status] || STATO_CFG.bozza;
        const isGruppo = c.groupId != null;
        const metaExtra = [
          c.numDoc  ? `📄 ${c.numDoc}` : '',
          c.tipoDoc ? c.tipoDoc : '',
          c.pagatoIl ? `💶 ${new Date(c.pagatoIl).toLocaleDateString('it-IT')}` : '',
          c.modalitaPag ? c.modalitaPag : '',
        ].filter(Boolean).join(' · ');
        return `<div class="bill-list-item">
          <div class="bill-list-dot" style="background:${dot}"></div>
          <div class="bill-list-info" onclick="riapriFoglio(${c.bookingId})" style="cursor:pointer;flex:1">
            <div class="bill-list-name">${c.nome} ${isGruppo?'<span style="font-size:9px;background:#e8f4fd;color:#1a6fa8;padding:1px 5px;border-radius:8px;margin-left:4px">GRUPPO</span>':''}</div>
            <div class="bill-list-sub">Cam. ${c.camera} · ${fmt(ci)} → ${fmt(co)}</div>
            ${metaExtra?`<div style="font-size:10px;color:var(--text3);margin-top:2px">${metaExtra}</div>`:''}
          </div>
          <div style="display:flex;flex-direction:column;align-items:flex-end;gap:4px;min-width:80px">
            <div class="bill-list-total">${(c.totale||0).toFixed(2)}€</div>
            <div class="stato-pill ${c.status}" onclick="avanzaStatoConto('${c.id}',event)">
              ${cfg.icon} ${cfg.label} ▸
            </div>
          </div>
        </div>`;
      }).join('')}
    </div>`;
  }).join('');
}

function avanzaStatoConto(contoId, e) {
  e && e.stopPropagation();
  const conti = loadConti();
  const idx = conti.findIndex(c => c.id === contoId);
  if (idx < 0) return;
  const c   = conti[idx];
  const cfg = STATO_CFG[c.status] || STATO_CFG.bozza;
  if (!cfg.next) { showToast('Conto già pagato','info'); return; }

  // Se avanza a fatturato: chiede tipo doc e numero
  if (cfg.next === 'fatturato') {
    apriDialogFatturazione(contoId);
    return;
  }
  // Se avanza a pagato: chiede modalità pagamento
  if (cfg.next === 'pagato') {
    apriDialogPagamento(contoId);
    return;
  }
  // emesso: avanza direttamente
  conti[idx] = { ...c, status: cfg.next, emessoIl: c.emessoIl || new Date().toISOString(), ts: new Date().toISOString() };
  saveConti(conti);
  showToast(`✓ Stato aggiornato: ${STATO_CFG[cfg.next].label}`, 'success');
  renderContiTab('lista');
}

function apriDialogFatturazione(contoId) {
  const conti = loadConti();
  const c = conti.find(x => x.id === contoId);
  if (!c) return;
  const ov = document.createElement('div');
  ov.style.cssText = 'position:fixed;inset:0;background:rgba(0,0,0,.5);z-index:9999;display:flex;align-items:center;justify-content:center;padding:20px;box-sizing:border-box';
  ov.innerHTML = `<div style="background:var(--surface);border-radius:16px;padding:24px;max-width:340px;width:100%">
    <div style="font-weight:700;font-size:15px;margin-bottom:16px">🧾 Registra documento fiscale</div>
    <div style="margin-bottom:10px">
      <label style="font-size:11px;color:var(--text3);text-transform:uppercase;display:block;margin-bottom:4px">Tipo documento</label>
      <select id="_dTipo" style="width:100%;border:1px solid var(--border);border-radius:8px;padding:9px;font-size:13px;background:var(--surface2);color:var(--text)">
        <option value="Fattura">Fattura</option>
        <option value="Scontrino">Scontrino</option>
        <option value="Ricevuta">Ricevuta</option>
        <option value="Fattura PA">Fattura PA</option>
      </select>
    </div>
    <div style="margin-bottom:16px">
      <label style="font-size:11px;color:var(--text3);text-transform:uppercase;display:block;margin-bottom:4px">Numero / riferimento</label>
      <input id="_dNum" placeholder="es. 2026/042" style="width:100%;box-sizing:border-box;border:1px solid var(--border);border-radius:8px;padding:9px;font-size:13px;background:var(--surface2);color:var(--text)">
    </div>
    <div style="display:flex;gap:8px">
      <button onclick="
        const tipo=document.getElementById('_dTipo').value;
        const num=document.getElementById('_dNum').value.trim();
        confermaFatturazione('${contoId}',tipo,num);
        this.closest('div[style*=fixed]').remove()
      " style="flex:1;background:var(--accent);color:#fff;border:none;border-radius:8px;padding:10px;font-weight:600;cursor:pointer;font-size:13px">
        ✓ Conferma
      </button>
      <button onclick="this.closest('div[style*=fixed]').remove()"
        style="background:none;border:1px solid var(--border);border-radius:8px;padding:10px;cursor:pointer;font-size:13px;color:var(--text2)">
        Annulla
      </button>
    </div>
  </div>`;
  document.body.appendChild(ov);
  ov.addEventListener('click', e => { if(e.target===ov) ov.remove(); });
  setTimeout(() => document.getElementById('_dNum')?.focus(), 100);
}

function confermaFatturazione(contoId, tipoDoc, numDoc) {
  const conti = loadConti();
  const idx = conti.findIndex(c => c.id === contoId);
  if (idx < 0) return;
  conti[idx] = { ...conti[idx], status:'fatturato', tipoDoc, numDoc: numDoc||'—',
    fatturatoIl: new Date().toISOString(), ts: new Date().toISOString() };
  saveConti(conti);
  showToast(`✓ ${tipoDoc} registrata`, 'success');
  renderContiTab('lista');
}

function apriDialogPagamento(contoId, prefillImporto) {
  const conti = loadConti();
  const conto = conti.find(c => c.id === contoId);
  if (!conto) return;
  const totaleConto = conto.totale || 0;
  const giaPagato   = getTotalePagatoPerBooking(conto.bookingId);
  const residuo     = Math.max(0, totaleConto - giaPagato).toFixed(2);
  const importoSuggerito = prefillImporto !== undefined ? prefillImporto : residuo;

  const ov = document.createElement('div');
  ov.style.cssText = 'position:fixed;inset:0;background:rgba(0,0,0,.5);z-index:9999;display:flex;align-items:center;justify-content:center;padding:20px;box-sizing:border-box';
  ov.innerHTML = `<div style="background:var(--surface);border-radius:16px;padding:24px;max-width:360px;width:100%">
    <div style="font-weight:700;font-size:15px;margin-bottom:4px">💶 Registra pagamento</div>
    ${giaPagato > 0 ? `<div style="font-size:12px;color:var(--text3);margin-bottom:14px">Già pagato: €${giaPagato.toFixed(2)} · Residuo: €${residuo}</div>` : '<div style="margin-bottom:14px"></div>'}
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-bottom:10px">
      <div>
        <label style="font-size:11px;color:var(--text3);text-transform:uppercase;display:block;margin-bottom:4px">Importo €</label>
        <input id="_pImporto" type="number" step="0.01" min="0.01" value="${importoSuggerito}"
          style="width:100%;box-sizing:border-box;border:1.5px solid var(--accent);border-radius:8px;padding:9px;font-size:14px;font-weight:600;background:var(--surface2);color:var(--text)">
      </div>
      <div>
        <label style="font-size:11px;color:var(--text3);text-transform:uppercase;display:block;margin-bottom:4px">Tipo</label>
        <select id="_pTipo" style="width:100%;border:1px solid var(--border);border-radius:8px;padding:9px;font-size:13px;background:var(--surface2);color:var(--text)">
          <option value="saldo">Saldo</option>
          <option value="acconto">Acconto</option>
          <option value="extra">Extra</option>
        </select>
      </div>
    </div>
    <div style="margin-bottom:10px">
      <label style="font-size:11px;color:var(--text3);text-transform:uppercase;display:block;margin-bottom:4px">Metodo</label>
      <select id="_pMod" style="width:100%;border:1px solid var(--border);border-radius:8px;padding:9px;font-size:13px;background:var(--surface2);color:var(--text)">
        <option value="Contanti">Contanti</option>
        <option value="Carta di credito">Carta di credito</option>
        <option value="Bancomat">Bancomat</option>
        <option value="Bonifico">Bonifico</option>
        <option value="Assegno">Assegno</option>
        <option value="Altro">Altro</option>
      </select>
    </div>
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-bottom:10px">
      <div>
        <label style="font-size:11px;color:var(--text3);text-transform:uppercase;display:block;margin-bottom:4px">Data</label>
        <input id="_pData" type="date" value="${new Date().toISOString().slice(0,10)}"
          style="width:100%;box-sizing:border-box;border:1px solid var(--border);border-radius:8px;padding:9px;font-size:13px;background:var(--surface2);color:var(--text)">
      </div>
      <div>
        <label style="font-size:11px;color:var(--text3);text-transform:uppercase;display:block;margin-bottom:4px">Riferimento</label>
        <input id="_pRif" placeholder="es. Visa *4521" style="width:100%;box-sizing:border-box;border:1px solid var(--border);border-radius:8px;padding:9px;font-size:13px;background:var(--surface2);color:var(--text)">
      </div>
    </div>
    <div style="margin-bottom:16px;display:flex;align-items:center;gap:8px">
      <input type="checkbox" id="_pDoc" style="width:16px;height:16px;cursor:pointer">
      <label for="_pDoc" style="font-size:13px;color:var(--text2);cursor:pointer">Con documento fiscale (scontrino/fattura)</label>
    </div>
    <div style="display:flex;gap:8px">
      <button onclick="
        const imp=parseFloat(document.getElementById('_pImporto').value)||0;
        const tipo=document.getElementById('_pTipo').value;
        const mod=document.getElementById('_pMod').value;
        const data=document.getElementById('_pData').value;
        const rif=document.getElementById('_pRif').value.trim();
        const doc=document.getElementById('_pDoc').checked;
        if(imp<=0){alert('Inserisci un importo valido');return;}
        confermaPagamento('${contoId}',imp,tipo,mod,data,rif,doc);
        this.closest('div[style*=fixed]').remove();
      " style="flex:1;background:#1e8449;color:#fff;border:none;border-radius:8px;padding:11px;font-weight:600;cursor:pointer;font-size:13px">
        ✓ Registra pagamento
      </button>
      <button onclick="this.closest('div[style*=fixed]').remove()"
        style="background:none;border:1px solid var(--border);border-radius:8px;padding:11px;cursor:pointer;font-size:13px;color:var(--text2)">
        Annulla
      </button>
    </div>
  </div>`;
  document.body.appendChild(ov);
  ov.addEventListener('click', e => { if(e.target===ov) ov.remove(); });
  setTimeout(() => document.getElementById('_pImporto')?.select(), 100);
}

function confermaPagamento(contoId, importo, tipo, modalita, dataStr, riferimento, conDocumento) {
  const conti = loadConti();
  const idx = conti.findIndex(c => c.id === contoId);
  if (idx < 0) return;
  const conto = conti[idx];
  const totaleConto = conto.totale || 0;
  const giaPagato   = getTotalePagatoPerBooking(conto.bookingId);
  const nuovoTotale = giaPagato + importo;

  // Registra il pagamento nel foglio PAGAMENTI
  const dataFmt = dataStr
    ? new Date(dataStr).toLocaleDateString('it-IT')
    : new Date().toLocaleDateString('it-IT');
  registraPagamento({
    contoId, bookingId: conto.bookingId,
    data: dataFmt, importo, tipo, metodo: modalita,
    riferimento, conDocumento,
  }).catch(e => console.warn('[PAGAMENTI] registra:', e.message));

  // Aggiorna stato del conto
  const nuovoStato = nuovoTotale >= totaleConto - 0.01 ? 'pagato' : 'emesso';
  conti[idx] = {
    ...conto, status: nuovoStato,
    modalitaPag: modalita,
    pagatoIl: dataStr ? new Date(dataStr).toISOString() : new Date().toISOString(),
    ts: new Date().toISOString(),
  };
  saveConti(conti);

  if (nuovoStato === 'pagato') {
    showToast('✓ Pagamento completato — conto saldato', 'success');
  } else {
    showToast(`✓ Acconto €${importo.toFixed(2)} registrato — residuo €${(totaleConto-nuovoTotale).toFixed(2)}`, 'success');
  }
  renderContiTab('lista');
  if (typeof render === 'function') render();
}

function riapriFoglio(bid) {
  closeConti();
  if (bookings.find(x=>x.id===bid)) showBookingDetail(bid);
}

// ═══════════════════════════════════════════════════════════════════
// CONTO DI GRUPPO — fattura riepilogativa per più prenotazioni
// ═══════════════════════════════════════════════════════════════════

function renderContoGruppo() {
  const mesPad = n => String(n).padStart(2,'0');

  // Genera shortcut mesi: mese corrente + 5 precedenti
  const mesiShortcut = [];
  for (let i = 0; i < 6; i++) {
    let m = curM - i, y = curY;
    if (m < 0) { m += 12; y--; }
    mesiShortcut.push({ y, m, label: MONTHS_S[m] + ' ' + y });
  }
  const shortcutHtml = mesiShortcut.map(s =>
    `<span class="sf-chip" onclick="setGrpMese(${s.y},${s.m})">${s.label}</span>`
  ).join('');

  const dalDef = `${curY}-${mesPad(curM+1)}-01`;
  const alDef  = `${curY}-${mesPad(curM+1)}-${new Date(curY,curM+1,0).getDate()}`;

  // Autocomplete nomi clienti
  const nomiUnici = [...new Set(bookings.map(b=>b.n).filter(Boolean))].sort();
  const datalist  = nomiUnici.map(n=>`<option value="${n.replace(/"/g,'&quot;')}">`).join('');

  return `
    <div class="conti-section">
      <div class="conti-section-title">Cerca prenotazioni per cliente</div>
      <datalist id="dlNomiGruppo">${datalist}</datalist>
      <div style="display:flex;flex-direction:column;gap:10px">

        <div>
          <label style="font-size:10px;color:var(--text3);text-transform:uppercase;letter-spacing:.06em;display:block;margin-bottom:4px">Nome cliente / gruppo</label>
          <input id="grpNome" list="dlNomiGruppo" placeholder="es. mammana"
            style="width:100%;box-sizing:border-box;border:1px solid var(--border);border-radius:8px;padding:9px 12px;font-size:14px;background:var(--surface2);color:var(--text)">
        </div>

        <div>
          <label style="font-size:10px;color:var(--text3);text-transform:uppercase;letter-spacing:.06em;display:block;margin-bottom:6px">Periodo rapido</label>
          <div style="display:flex;gap:6px;flex-wrap:wrap">${shortcutHtml}</div>
        </div>

        <div>
          <label style="font-size:10px;color:var(--text3);text-transform:uppercase;letter-spacing:.06em;display:block;margin-bottom:6px">Oppure date libere</label>
          <div style="display:grid;grid-template-columns:1fr 1fr;gap:8px">
            <div>
              <label style="font-size:10px;color:var(--text3);display:block;margin-bottom:3px">Dal</label>
              <input id="grpDal" type="date" value="${dalDef}"
                style="width:100%;box-sizing:border-box;border:1px solid var(--border);border-radius:8px;padding:9px 12px;font-size:13px;background:var(--surface2);color:var(--text)">
            </div>
            <div>
              <label style="font-size:10px;color:var(--text3);display:block;margin-bottom:3px">Al</label>
              <input id="grpAl" type="date" value="${alDef}"
                style="width:100%;box-sizing:border-box;border:1px solid var(--border);border-radius:8px;padding:9px 12px;font-size:13px;background:var(--surface2);color:var(--text)">
            </div>
          </div>
        </div>

        <div style="display:flex;gap:8px">
          <button class="btn primary" id="grpCercaBtn" style="flex:1;justify-content:center">🔍 Cerca prenotazioni</button>
          <button class="btn" onclick="setGrpTutto()" style="justify-content:center" title="Cerca in tutto l'anno">Tutti</button>
        </div>
      </div>
    </div>
    <div id="grpRisultati"></div>`;
}

function setGrpMese(y, m) {
  const mesPad = n => String(n).padStart(2,'0');
  const dal = document.getElementById('grpDal');
  const al  = document.getElementById('grpAl');
  if (dal) dal.value = `${y}-${mesPad(m+1)}-01`;
  if (al)  al.value  = `${y}-${mesPad(m+1)}-${new Date(y,m+1,0).getDate()}`;
  // Evidenzia il chip selezionato
  document.querySelectorAll('#contiBody .sf-chip').forEach(c => {
    c.classList.toggle('active', c.textContent === MONTHS_S[m]+' '+y);
  });
}

function setGrpTutto() {
  const dal = document.getElementById('grpDal');
  const al  = document.getElementById('grpAl');
  if (dal) dal.value = `${curY}-01-01`;
  if (al)  al.value  = `${curY}-12-31`;
  document.querySelectorAll('#contiBody .sf-chip').forEach(c => c.classList.remove('active'));
}

function bindContoGruppo() {
  const btn = document.getElementById('grpCercaBtn');
  if (btn) btn.onclick = cercaContoGruppo;
  const inp = document.getElementById('grpNome');
  if (inp) inp.addEventListener('keydown', e => { if(e.key==='Enter') cercaContoGruppo(); });
}

function cercaContoGruppo() {
  const nome = (document.getElementById('grpNome')?.value||'').trim().toLowerCase();
  const dal  = document.getElementById('grpDal')?.value;
  const al   = document.getElementById('grpAl')?.value;
  if (!nome) { showToast('Inserisci il nome del cliente','error'); return; }

  const dalD = dal ? new Date(dal+'T00:00:00') : null;
  const alD  = al  ? new Date(al +'T23:59:59') : null;

  // Cerca prenotazioni che contengono il nome e si sovrappongono al periodo
  const trovate = bookings.filter(b => {
    if (!b.n.toLowerCase().includes(nome)) return false;
    if (dalD && b.e < dalD) return false;
    if (alD  && b.s > alD)  return false;
    return true;
  }).sort((a,b)=>a.s-b.s);

  const el = document.getElementById('grpRisultati');
  if (!trovate.length) {
    el.innerHTML = `<div class="search-empty" style="padding:30px 0">
      <div class="emptyicon">🔍</div>
      <div style="font-size:12px;color:var(--text3)">Nessuna prenotazione trovata per "<b>${nome}</b>"</div>
    </div>`;
    return;
  }

  renderRisultatiGruppo(trovate, nome, dalD, alD, el);
}

function renderRisultatiGruppo(trovate, nome, dalD, alD, el) {
  const cfg  = loadBillSettings();
  let righeGruppo = [];
  let totaleGruppo = 0;

  const righeHtml = trovate.map(b => {
    const room   = ROOMS.find(r=>r.id===b.r);
    const n      = nights(b.s, b.e);
    // Ricalcola il conto usando i dati ATTUALI del booking (disposizione dal foglio)
    // invece di getContoEffettivo che potrebbe avere dati obsoleti dal DB
    const contoFresco = calcolaConto(b, getExtraForBooking(b.id));
    const totB   = contoFresco?.totale || 0;
    const prezzoN= n > 0 ? parseFloat((totB/n).toFixed(2)) : totB;
    totaleGruppo += totB;

    // Aggiungi riga per il conto di gruppo
    righeGruppo.push({
      label:     `${room?.name||roomName(b.r)} · ${fmt(b.s)} → ${fmt(b.e)} (${n} notti) · ${b.d||''}`,
      qty:       n,
      unitPrice: prezzoN,
      total:     totB,
      tipo:      'base',
      bookingId: b.id
    });

    return `<div class="sr-item" style="cursor:default">
      <div class="sr-dot" style="background:${pastello(b.c)}"></div>
      <div class="sr-body">
        <div class="sr-name">${room?.name||roomName(b.r)}</div>
        <div class="sr-meta">
          <span class="sr-badge">${fmt(b.s)} → ${fmt(b.e)}</span>
          <span class="sr-badge">${n} notti</span>
          ${b.d?`<span class="sr-badge">${b.d}</span>`:''}
          <span class="sr-badge" style="background:var(--accent);color:#fff">${prezzoN.toFixed(2)}€/n</span>
        </div>
      </div>
      <div style="font-weight:700;font-size:13px;color:var(--accent);padding-top:2px;flex-shrink:0">${totB.toFixed(2)}€</div>
    </div>`;
  }).join('');

  // Applica eventuale convenzione al totale di gruppo
  const nomeCliente = nome;
  const convenzione = (cfg.convenzioni||[]).find(cv =>
    nomeCliente.includes(cv.nome.toLowerCase())
  );
  let scontoHtml = '';
  if (convenzione) {
    const sc = parseFloat((totaleGruppo * convenzione.sconto / 100).toFixed(2));
    totaleGruppo -= sc;
    righeGruppo.push({
      label: `Convenzione "${convenzione.nome}" -${convenzione.sconto}%`,
      qty: null, unitPrice: null, total: -sc, tipo: 'sconto'
    });
    scontoHtml = `<div class="bill-row discount" style="font-size:12px">
      <span class="bill-row-label">Conv. ${convenzione.nome} -${convenzione.sconto}%</span>
      <span style="color:var(--danger);font-weight:600">-${sc.toFixed(2)}€</span>
    </div>`;
  }

  const aliqIVA     = cfg.aliquotaIVA || 10;
  const imponibile  = parseFloat((totaleGruppo / (1+aliqIVA/100)).toFixed(2));
  const iva         = parseFloat((totaleGruppo - imponibile).toFixed(2));

  el.innerHTML = `
    <div class="conti-section">
      <div class="conti-section-title">${trovate.length} prenotazioni trovate</div>
      ${righeHtml}
    </div>
    <div class="conti-section">
      <div class="conti-section-title">Extra di gruppo</div>
      <div id="grpExtraList" style="margin-bottom:8px"></div>
      <div style="display:flex;gap:6px;align-items:center;flex-wrap:wrap">
        <input id="grpExtraLabel" placeholder="Descrizione (es. Cene)" 
          style="flex:2;min-width:100px;border:1px solid var(--border);border-radius:6px;padding:7px 10px;font-size:13px;background:var(--surface2);color:var(--text)">
        <input id="grpExtraQty" type="number" placeholder="Qtà" min="1" value="1"
          style="width:54px;border:1px solid var(--border);border-radius:6px;padding:7px 8px;font-size:13px;background:var(--surface2);color:var(--text)">
        <input id="grpExtraPrezzo" type="number" placeholder="€" min="0" step="0.01"
          style="width:64px;border:1px solid var(--border);border-radius:6px;padding:7px 8px;font-size:13px;background:var(--surface2);color:var(--text)">
        <button class="btn" onclick="aggiungiExtraGruppo()" style="padding:7px 12px;font-size:12px">+ Aggiungi</button>
      </div>
    </div>
    <div class="conti-section" id="grpRiepilogoSection">
      <div class="conti-section-title">Riepilogo conto di gruppo</div>
      ${scontoHtml}
      <div id="grpRigheExtra"></div>
      <div class="bill-row" id="grpRowImponibile" style="font-size:12px">
        <span class="bill-row-label">Imponibile (IVA ${aliqIVA}% esclusa)</span>
        <span style="font-weight:600">${imponibile.toFixed(2)}€</span>
      </div>
      <div class="bill-row" id="grpRowIva" style="font-size:12px">
        <span class="bill-row-label">IVA ${aliqIVA}%</span>
        <span style="font-weight:600">${iva.toFixed(2)}€</span>
      </div>
      <div class="bill-row total-row" id="grpRowTotale">
        <span class="bill-row-label"><b>TOTALE IVA inclusa</b></span>
        <span style="font-size:18px;font-weight:700;color:var(--accent)">${totaleGruppo.toFixed(2)}€</span>
      </div>
      <div style="display:flex;gap:8px;margin-top:14px;flex-wrap:wrap">
        <button class="btn primary" onclick="emettiContoGruppo()" style="flex:1;justify-content:center;min-width:80px">📄 Conto</button>
        <button class="btn" onclick="apriPdfGruppo()" style="flex:1;justify-content:center;min-width:80px">👁 PDF</button>
        <button class="btn" onclick="esportaXMLGruppo()" style="flex:1;justify-content:center;min-width:80px">📄 XML</button>
        <button class="btn" onclick="apriReportPresenze()" style="flex:1;justify-content:center;min-width:80px;background:var(--surface2)">📊 Report</button>
        <button class="btn" onclick="esportaReportCSV()" style="flex:1;justify-content:center;min-width:80px;background:#e8f5e9;color:#2d6a4f">📋 CSV</button>
      </div>
    </div>`;

  // Salva in variabile globale per PDF/XML
  window._gruppoCorrente = {
    nome:      trovate[0]?.n || nome,
    righe:     righeGruppo,
    totaleBase:parseFloat(totaleGruppo.toFixed(2)),
    totale:    parseFloat(totaleGruppo.toFixed(2)),
    aliquotaIVA: aliqIVA,
    bookings:  trovate,
    extraGruppo: [],
    dalD, alD
  };
}

function aggiungiExtraGruppo() {
  const label  = document.getElementById('grpExtraLabel')?.value.trim();
  const qty    = parseFloat(document.getElementById('grpExtraQty')?.value)  || 1;
  const prezzo = parseFloat(document.getElementById('grpExtraPrezzo')?.value)|| 0;
  if (!label || prezzo <= 0) { showToast('Inserisci descrizione e prezzo','error'); return; }

  const tot = parseFloat((qty * prezzo).toFixed(2));
  const g   = window._gruppoCorrente;
  if (!g) return;

  // Aggiungi alle righe
  g.extraGruppo.push({ label, qty, unitPrice: prezzo, total: tot, tipo:'extra' });
  g.righe.push({ label, qty, unitPrice: prezzo, total: tot, tipo:'extra' });
  g.totale = parseFloat((g.totale + tot).toFixed(2));

  // Aggiorna UI
  const listEl = document.getElementById('grpExtraList');
  if (listEl) {
    const div = document.createElement('div');
    div.style.cssText='display:flex;justify-content:space-between;align-items:center;padding:5px 0;font-size:12px;border-bottom:1px solid var(--border)';
    div.innerHTML=`<span>${label} × ${qty}</span><span style="font-weight:600">${tot.toFixed(2)}€
      <button onclick="rimuoviExtraGruppo(this,${tot},'${label.replace(/'/g,"\'")}',${qty},${prezzo})"
        style="margin-left:8px;background:none;border:none;color:var(--danger);cursor:pointer;font-size:13px">✕</button>
    </span>`;
    listEl.appendChild(div);
  }

  // Ricalcola riepilogo
  aggiornaRiepilogoGruppo();

  // Pulisci campi
  document.getElementById('grpExtraLabel').value='';
  document.getElementById('grpExtraQty').value='1';
  document.getElementById('grpExtraPrezzo').value='';
}

function rimuoviExtraGruppo(btn, tot, label, qty, prezzo) {
  const g = window._gruppoCorrente;
  if (!g) return;
  g.righe = g.righe.filter(r => !(r.tipo==='extra' && r.label===label && r.qty===qty && r.unitPrice===prezzo));
  g.extraGruppo = g.extraGruppo.filter(r => !(r.label===label && r.qty===qty && r.unitPrice===prezzo));
  g.totale = parseFloat((g.totale - tot).toFixed(2));
  btn.closest('div').remove();
  aggiornaRiepilogoGruppo();
}

function aggiornaRiepilogoGruppo() {
  const g = window._gruppoCorrente;
  if (!g) return;
  const aliq = g.aliquotaIVA || 10;
  const imp  = parseFloat((g.totale / (1+aliq/100)).toFixed(2));
  const iva  = parseFloat((g.totale - imp).toFixed(2));

  const riEl = document.getElementById('grpRowImponibile');
  const ivEl = document.getElementById('grpRowIva');
  const toEl = document.getElementById('grpRowTotale');
  if (riEl) riEl.querySelector('span:last-child').textContent = imp.toFixed(2)+'€';
  if (ivEl) ivEl.querySelector('span:last-child').textContent = iva.toFixed(2)+'€';
  if (toEl) toEl.querySelector('span:last-child').textContent = g.totale.toFixed(2)+'€';
}

function emettiContoGruppo() {
  const g = window._gruppoCorrente;
  if (!g || !g.bookings?.length) return;
  const conti  = loadConti();
  const groupId = 'G' + Date.now();
  const ora    = new Date().toISOString();

  // Crea una riga per ogni prenotazione del gruppo, collegate dallo stesso groupId
  g.bookings.forEach(b => {
    const room = ROOMS.find(r=>r.id===b.r);
    const n    = nights(b.s, b.e);
    const conto= calcolaConto(b, getExtraForBooking(b.id));
    const idx  = conti.findIndex(x => x.bookingId === b.id);
    const doc  = {
      id:          idx >= 0 ? conti[idx].id : ('C'+Date.now()+b.id),
      bookingId:   b.id,
      groupId:     groupId,
      nome:        b.n,
      camera:      room?.name || roomName(b.r),
      checkin:     b.s.toISOString(),
      checkout:    b.e.toISOString(),
      righe:       conto.righe,
      totale:      conto.totale,
      status:      'emesso',
      emessoIl:    ora,
      fatturatoIl: null, tipoDoc:null, numDoc:null,
      pagatoIl:    null, modalitaPag:null,
      isAppart:    false, ts: ora
    };
    if (idx >= 0) conti[idx] = doc; else conti.unshift(doc);
  });

  // Se ci sono extra di gruppo, aggiungi una riga extra con groupId
  if (g.extraGruppo?.length) {
    const extraDoc = {
      id: 'CE' + Date.now(),
      bookingId: null,
      groupId: groupId,
      nome: g.nome + ' (extra gruppo)',
      camera: '—',
      checkin: (g.dalD||new Date()).toISOString(),
      checkout: (g.alD||new Date()).toISOString(),
      righe: g.extraGruppo,
      totale: g.extraGruppo.reduce((s,r)=>s+r.total, 0),
      status: 'emesso', emessoIl: ora,
      fatturatoIl:null, tipoDoc:null, numDoc:null,
      pagatoIl:null, modalitaPag:null, ts:ora
    };
    conti.unshift(extraDoc);
  }

  saveConti(conti);
  apriPdfGruppo();
  showToast(`✓ Conto gruppo salvato — ${g.bookings.length} prenotazioni`, 'success');
}

function apriPdfGruppo() {
  const g = window._gruppoCorrente;
  if (!g) return;
  const cfg  = loadBillSettings();
  const oggi = new Date();

  const aliqIVA    = g.aliquotaIVA || 10;
  const imponibile = parseFloat((g.totale / (1+aliqIVA/100)).toFixed(2));
  const iva        = parseFloat((g.totale - imponibile).toFixed(2));

  const righeHtml = g.righe.map(r => `
    <tr>
      <td>${r.label}</td>
      <td style="text-align:center">${r.qty!=null?r.qty:'—'}</td>
      <td style="text-align:right">${r.unitPrice!=null?r.unitPrice.toFixed(2)+'€':'—'}</td>
      <td style="text-align:right;color:${r.total<0?'#c0392b':'inherit'}">${r.total>=0?'+':''}${r.total.toFixed(2)}€</td>
    </tr>`).join('');

  const periodoLabel = g.dalD && g.alD
    ? `${g.dalD.toLocaleDateString('it-IT')} — ${g.alD.toLocaleDateString('it-IT')}`
    : '';

  document.getElementById('pdfTitle').textContent = `Conto gruppo — ${g.nome}`;
  document.getElementById('printDoc').innerHTML = `
    <div class="doc-header">
      <div class="doc-hotel-name">${cfg.hotelName}</div>
      ${cfg.hotelAddress?`<div class="doc-hotel-sub">${cfg.hotelAddress}</div>`:''}
      <div class="doc-type">Conto di Gruppo</div>
    </div>
    <div class="doc-section">
      <div class="doc-section-title">Cliente</div>
      <div class="doc-info-grid">
        <div class="doc-info-item"><label>Nome</label><span>${g.nome}</span></div>
        <div class="doc-info-item"><label>Periodo</label><span>${periodoLabel}</span></div>
        <div class="doc-info-item"><label>Soggiorni</label><span>${g.bookings.length}</span></div>
      </div>
    </div>
    <div class="doc-section">
      <div class="doc-section-title">Dettaglio soggiorni</div>
      <table class="doc-table">
        <thead><tr><th>Descrizione</th><th style="text-align:center">Qtà</th><th style="text-align:right">Prezzo</th><th style="text-align:right">Importo</th></tr></thead>
        <tbody>${righeHtml}</tbody>
        <tfoot>
          <tr style="font-size:11px;color:#666"><td colspan="3">Imponibile (IVA ${aliqIVA}% esclusa)</td><td style="text-align:right">${imponibile.toFixed(2)} €</td></tr>
          <tr style="font-size:11px;color:#666"><td colspan="3">IVA ${aliqIVA}%</td><td style="text-align:right">${iva.toFixed(2)} €</td></tr>
          <tr class="doc-total-row"><td colspan="3"><strong>TOTALE IVA inclusa</strong></td><td><strong>${g.totale.toFixed(2)} €</strong></td></tr>
        </tfoot>
      </table>
    </div>
    <div class="doc-footer">${cfg.hotelName} · ${oggi.toLocaleDateString('it-IT')}</div>`;

  document.getElementById('pdfOverlay').classList.add('open');
  // Imposta bid fittizio per XML gruppo
  _currentPdfBid = '__gruppo__';
}

function esportaXMLGruppo() {
  apriPdfGruppo();
  // Dopo che il PDF overlay è aperto, triggera esportaXML in modo speciale
  setTimeout(() => _esportaXMLGruppo(), 100);
}

function _esportaXMLGruppo() {
  const g = window._gruppoCorrente;
  if (!g) return;
  try {
    const cfg    = loadBillSettings();
    const oggi   = new Date();
    const fmtD   = d => d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0')+'-'+String(d.getDate()).padStart(2,'0');
    const toLatin= s => String(s||'').replace(/[·•]/g,'-').replace(/[→⇒►]/g,'->').replace(/[×✕]/g,'x')
      .replace(/[àáâã]/g,'a').replace(/[èéêë]/g,'e').replace(/[ìíîï]/g,'i')
      .replace(/[òóôõ]/g,'o').replace(/[ùúûü]/g,'u').replace(/[ÀÁ]/g,'A')
      .replace(/[ÈÉ]/g,'E').replace(/[ÙÚ]/g,'U').replace(/[^\u0000-\u00FF]/g,'?');
    const esc    = s => toLatin(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
    const clean  = s => esc(s).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g,'').trim();

    const pivaHotel = clean(cfg.pivaHotel || '00000000000');
    const aliqXML   = parseFloat(g.aliquotaIVA || 10);
    const aliqFactor= 1 + aliqXML / 100;
    const numFattura= String(oggi.getFullYear()).slice(-2)+String(oggi.getMonth()+1).padStart(2,'0')+String(oggi.getDate()).padStart(2,'0')+'GR';

    const nomeParti = g.nome.trim().split(/\s+/);
    const usaDenom  = nomeParti.length===1 || g.nome.length>30;
    const clienteAnag = usaDenom
      ? `<Denominazione>${clean(g.nome)}</Denominazione>`
      : `<Nome>${clean(nomeParti[0])}</Nome><Cognome>${clean(nomeParti.slice(1).join(' '))}</Cognome>`;

    const righeBase   = g.righe.filter(r=>r.total>0);
    const righeSconto = g.righe.filter(r=>r.total<0);
    let lineaNum = 0;

    const linee = righeBase.map(r => {
      lineaNum++;
      // Normalizzazione generica per gruppo
      const _lbl = (r.label||'').toLowerCase();
      let desc;
      if (_lbl.includes('pernottamento')||_lbl.includes('nott')||_lbl.includes('camera')) {
        desc = 'Soggiorno';
      } else if (_lbl.includes('colazion')) { desc = 'Colazione';
      } else if (_lbl.includes('pranzo'))   { desc = 'Pranzo a prezzo fisso';
      } else if (_lbl.includes('cena'))     { desc = 'Cena a prezzo fisso';
      } else if (_lbl.includes('puliz'))    { desc = 'Servizio pulizie extra';
      } else if (_lbl.includes('lenzuol')||_lbl.includes('cambio')) { desc = 'Cambio biancheria';
      } else { desc = clean(r.label).slice(0,100)||'Servizio'; }
      const totImpon  = parseFloat((r.total / aliqFactor).toFixed(2));
      const qty       = r.qty != null ? parseFloat(r.qty).toFixed(2) : null;
      const prezzoUnit= r.unitPrice != null ? parseFloat((r.unitPrice/aliqFactor).toFixed(2)) : totImpon;
      const scontiTag = lineaNum===1 ? righeSconto.map(s=>{
        const imp = parseFloat((Math.abs(s.total)/aliqFactor).toFixed(2));
        return `<ScontoMaggiorazione><Tipo>SC</Tipo><Importo>${imp.toFixed(2)}</Importo></ScontoMaggiorazione>`;
      }).join('') : '';
      const scontiImp = lineaNum===1 ? parseFloat(righeSconto.reduce((s,r)=>s+Math.abs(r.total)/aliqFactor,0).toFixed(2)) : 0;
      const prezzoTot = parseFloat((totImpon - scontiImp).toFixed(2));
      return `
    <DettaglioLinee>
      <NumeroLinea>${lineaNum}</NumeroLinea>
      <Descrizione>${desc||'Soggiorno'}</Descrizione>
      ${qty?'<Quantita>'+qty+'</Quantita>':''}
      <PrezzoUnitario>${prezzoUnit.toFixed(2)}</PrezzoUnitario>
      ${scontiTag}
      <PrezzoTotale>${prezzoTot.toFixed(2)}</PrezzoTotale>
      <AliquotaIVA>${aliqXML.toFixed(2)}</AliquotaIVA>
    </DettaglioLinee>`;
    }).join('');

    const totaleNetto = g.righe.reduce((s,r)=>s+r.total,0);
    const imponibile  = parseFloat((totaleNetto/aliqFactor).toFixed(2));
    const iva         = parseFloat((imponibile*aliqXML/100).toFixed(2));
    const totDoc      = parseFloat((imponibile+iva).toFixed(2));
    const causale     = clean('Conto gruppo '+g.nome+(g.dalD?' dal '+fmtD(g.dalD):'')+(g.alD?' al '+fmtD(g.alD):'')).slice(0,200);

    const xml = `<?xml version="1.0" encoding="UTF-8"?>
<p:FatturaElettronica versione="FPR12"
  xmlns:p="http://ivaservizi.agenziaentrate.gov.it/docs/xsd/fatture/v1.2"
  xmlns:ds="http://www.w3.org/2000/09/xmldsig#"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <FatturaElettronicaHeader>
    <DatiTrasmissione>
      <IdTrasmittente><IdPaese>IT</IdPaese><IdCodice>${pivaHotel}</IdCodice></IdTrasmittente>
      <ProgressivoInvio>${numFattura}</ProgressivoInvio>
      <FormatoTrasmissione>FPR12</FormatoTrasmissione>
      <CodiceDestinatario>0000000</CodiceDestinatario>
    </DatiTrasmissione>
    <CedentePrestatore>
      <DatiAnagrafici>
        <IdFiscaleIVA><IdPaese>IT</IdPaese><IdCodice>${pivaHotel}</IdCodice></IdFiscaleIVA>
        <Anagrafica><Denominazione>${clean(cfg.hotelName)}</Denominazione></Anagrafica>
        <RegimeFiscale>RF01</RegimeFiscale>
      </DatiAnagrafici>
      <Sede>
        <Indirizzo>${clean(cfg.hotelAddress||'Da definire')}</Indirizzo>
        <CAP>${clean(cfg.hotelCAP||'00000')}</CAP>
        <Comune>${clean(cfg.hotelComune||'Da definire')}</Comune>
        <Provincia>${clean(cfg.hotelProv||'XX')}</Provincia>
        <Nazione>IT</Nazione>
      </Sede>
    </CedentePrestatore>
    <CessionarioCommittente>
      <DatiAnagrafici>
        <CodiceFiscale>RSSMRA00A00H501U</CodiceFiscale>
        <Anagrafica>${clienteAnag}</Anagrafica>
      </DatiAnagrafici>
      <Sede><Indirizzo>Da definire</Indirizzo><CAP>00000</CAP><Comune>Da definire</Comune><Nazione>IT</Nazione></Sede>
    </CessionarioCommittente>
  </FatturaElettronicaHeader>
  <FatturaElettronicaBody>
    <DatiGenerali>
      <DatiGeneraliDocumento>
        <TipoDocumento>TD01</TipoDocumento><Divisa>EUR</Divisa>
        <Data>${fmtD(oggi)}</Data>
        <Numero>${numFattura}</Numero>
        <ImportoTotaleDocumento>${totDoc.toFixed(2)}</ImportoTotaleDocumento>
        <Causale>${causale}</Causale>
      </DatiGeneraliDocumento>
    </DatiGenerali>
    <DatiBeniServizi>
      ${linee}
      <DatiRiepilogo>
        <AliquotaIVA>${aliqXML.toFixed(2)}</AliquotaIVA>
        <ImponibileImporto>${imponibile.toFixed(2)}</ImponibileImporto>
        <Imposta>${iva.toFixed(2)}</Imposta>
        <EsigibilitaIVA>I</EsigibilitaIVA>
      </DatiRiepilogo>
    </DatiBeniServizi>
    <DatiPagamento>
      <CondizioniPagamento>TP02</CondizioniPagamento>
      <DettaglioPagamento>
        <ModalitaPagamento>MP01</ModalitaPagamento>
        <ImportoPagamento>${totDoc.toFixed(2)}</ImportoPagamento>
      </DettaglioPagamento>
    </DatiPagamento>
  </FatturaElettronicaBody>
</p:FatturaElettronica>`;

    const nomeFile = 'fattura_gruppo_'+g.nome.replace(/\s+/g,'_')+'_'+fmtD(oggi)+'.xml';
    const blob     = new Blob([xml],{type:'application/xml;charset=utf-8'});
    const url      = URL.createObjectURL(blob);
    const ov       = document.createElement('div');
    ov.style.cssText='position:fixed;inset:0;background:rgba(0,0,0,.6);z-index:9999;display:flex;align-items:center;justify-content:center;padding:20px;box-sizing:border-box';
    ov.innerHTML=`<div style="background:var(--surface);border-radius:16px;padding:24px;max-width:360px;width:100%;text-align:center">
      <div style="font-size:36px;margin-bottom:10px">📄</div>
      <div style="font-weight:700;font-size:15px;margin-bottom:4px">XML Conto Gruppo pronto</div>
      <div style="font-size:11px;color:var(--text3);margin-bottom:20px;word-break:break-all">${nomeFile}</div>
      <a href="${url}" download="${nomeFile}" id="_grpXmlLink"
        style="display:block;background:var(--accent);color:#fff;padding:14px;border-radius:10px;text-decoration:none;font-weight:700;font-size:15px;margin-bottom:10px">
        ⬇ Scarica XML
      </a>
      <button onclick="this.closest('div[style]').remove()"
        style="background:none;border:1px solid var(--border);padding:10px;border-radius:8px;cursor:pointer;font-size:13px;color:var(--text2);width:100%">Chiudi</button>
    </div>`;
    document.body.appendChild(ov);
    ov.querySelector('#_grpXmlLink').addEventListener('click',()=>setTimeout(()=>{URL.revokeObjectURL(url);ov.remove();},1000));
    ov.addEventListener('click',e=>{if(e.target===ov)ov.remove();});

  } catch(err) {
    showToast('Errore XML gruppo: '+err.message,'error');
    console.error('[XML gruppo]',err);
  }
}

function renderContiAppart() {
  const oggi   = new Date();
  const apIds  = ROOMS.filter(r=>r.g==='Appartamenti').map(r=>r.id);
  const attivi = bookings.filter(b=>apIds.includes(b.r) && b.s<=oggi && b.e>=oggi);
  if (!attivi.length) return `<div class="empty" style="padding:40px 0"><div class="emptyicon">🏠</div><div style="font-size:12px;color:var(--text3)">Nessun appartamento occupato oggi.</div></div>`;
  const cfg = loadBillSettings();
  return `<div class="conti-section"><div class="conti-section-title">Appartamenti occupati — ${oggi.toLocaleDateString('it-IT',{month:'long',year:'numeric'})}</div>
    ${attivi.map(b=>{
      const room=ROOMS.find(r=>r.id===b.r);
      const ov=cfg.tariffeCamere?.[b.r]||{};
      const notti=nights(b.s,b.e);
      const usaM=getAppartMode(b.id,notti);
      const canone=usaM?(ov.mensile||0):(ov.giornaliera||0);
      const qty=usaM?parseFloat((notti/30).toFixed(2)):notti;
      const tot=parseFloat((canone*qty).toFixed(2));
      return `<div class="bill-card">
        <div class="bill-card-hdr">
          <div><div class="bill-card-name">${b.n}</div><div class="bill-card-room">${room?.name}</div></div>
          <div><div class="bill-total">${tot.toFixed(2)}€</div>
            <button class="btn" style="font-size:11px;margin-top:4px" onclick="closeConti();showBookingDetail(${b.id})">→ Conto</button>
          </div>
        </div>
        <div class="bill-row"><span class="bill-row-label">Periodo</span><span class="bill-row-total">${fmt(b.s)} → ${fmt(b.e)}</span></div>
        <div class="bill-row"><span class="bill-row-label">Tariffa</span><span class="bill-row-total">${canone}€/${usaM?'mese':'notte'} × ${qty}</span></div>
      </div>`;
    }).join('')}</div>`;
}

function renderListino() {
  const cfg = loadBillSettings();
  const t   = cfg.tariffe||{};
  const groups = {};
  ROOMS.forEach(r=>{ if(!groups[r.g]) groups[r.g]=[]; groups[r.g].push(r); });

  const camereHtml = Object.entries(groups).map(([g,rooms])=>`
    <div class="conti-section">
      <div class="conti-section-title">${g}</div>
      <div style="display:grid;grid-template-columns:1fr 70px 80px;gap:3px 8px;font-size:10px;font-weight:600;color:var(--text3);padding-bottom:4px">
        <span>Camera</span><span style="text-align:right">Notte</span><span style="text-align:right">Mensile</span>
      </div>
      ${rooms.map(r=>{
        const ov=cfg.tariffeCamere?.[r.id]||{};
        const gv=ov.giornaliera>0?ov.giornaliera+'€':'—';
        const mv=ov.mensile>0?ov.mensile+'€':'—';
        return `<div style="display:grid;grid-template-columns:1fr 70px 80px;gap:3px 8px;padding:5px 0;border-bottom:1px solid var(--border);font-size:12px">
          <span>${r.name}</span>
          <span style="text-align:right;font-weight:600;color:${ov.giornaliera>0?'var(--text)':'var(--text3)'}">${gv}</span>
          <span style="text-align:right;font-weight:600;color:${ov.mensile>0?'var(--text)':'var(--text3)'}">${mv}</span>
        </div>`;
      }).join('')}
    </div>`).join('');

  return `
    <div class="conti-section">
      <div class="conti-section-title">Tariffe base per disposizione letti</div>
      <div class="bill-row"><span class="bill-row-label">Camera singola (1 letto singolo)</span><span class="bill-row-total">${t.s||0}€/notte</span></div>
      <div class="bill-row"><span class="bill-row-label">Matrimoniale uso singolo</span><span class="bill-row-total">${t.ms||0}€/notte</span></div>
      <div class="bill-row"><span class="bill-row-label">Matrimoniale (2 persone)</span><span class="bill-row-total">${t.m||0}€/notte</span></div>
      <div class="bill-row"><span class="bill-row-label">Letto singolo aggiunto (+p.p.)</span><span class="bill-row-total">+${t.ag||0}€/notte</span></div>
    </div>
    ${camereHtml}
    <div class="conti-section">
      <div class="conti-section-title">Stagionalità</div>
      ${cfg.stagioni.map(s=>`<div class="bill-row"><span class="bill-row-label">${s.nome} (${s.dal} → ${s.al})</span><span class="rate-badge stagionale">×${s.molt}</span></div>`).join('')}
    </div>
    <div class="conti-section">
      <div class="conti-section-title">Convenzioni</div>
      ${cfg.convenzioni.map(c=>`<div class="bill-row"><span class="bill-row-label">${c.nome}</span><span class="rate-badge conv">-${c.sconto}%</span></div>`).join('')}
    </div>
    <div class="conti-section">
      <div class="conti-section-title">Extra configurati</div>
      ${(cfg.extra||[]).map(e=>`<div class="bill-row"><span class="bill-row-label">${e.label}</span><span class="bill-row-total">${e.prezzo>0?e.prezzo+'€/'+e.unita:'da definire'}</span></div>`).join('')}
    </div>
    <button class="btn primary" onclick="openContiSettings()" style="width:100%;justify-content:center;margin-top:8px">⚙ Modifica tariffe</button>`;
}

// ─────────────────────────────────────────────────────────────────
// SETTINGS TARIFFE
// ─────────────────────────────────────────────────────────────────

function openContiSettings() {
  document.getElementById('contiSettingsBody').innerHTML = buildContiSettingsForm(loadBillSettings());
  document.getElementById('contiSettingsOverlay').style.display = 'flex';
}
function closeContiSettings() { document.getElementById('contiSettingsOverlay').style.display='none'; }

function buildContiSettingsForm(cfg) {
  const t = cfg.tariffe||{};

  const stagRow = (s,i) => `
    <div class="rate-season-row" id="stagRow_${i}">
      <input type="text"   id="stagNome_${i}"   value="${s.nome}" placeholder="Nome">
      <input type="text"   id="stagPeriod_${i}" value="${s.dal}–${s.al}" placeholder="MM-DD–MM-DD">
      <input type="number" id="stagMolt_${i}"   value="${s.molt}" step="0.1" min="0.5" max="3">
      <button class="btn-icon-sm" onclick="this.closest('.rate-season-row').remove()">✕</button>
    </div>`;

  const convRow = (c,i) => `
    <div class="rate-conv-row" id="convRow_${i}">
      <input type="text"   id="convNome_${i}" value="${c.nome}" placeholder="Nome cliente">
      <input type="number" id="convVal_${i}"  value="${c.sconto}" step="1" min="0" max="100" placeholder="%">
      <button class="btn-icon-sm" onclick="this.closest('.rate-conv-row').remove()">✕</button>
    </div>`;

  const slpRow = (s,i) => `
    <div id="slpRow_${i}" style="display:grid;grid-template-columns:80px 60px 1fr 28px;gap:6px;align-items:center;margin-bottom:6px">
      <input type="number" id="slpNotti_${i}" value="${s.minNotti||0}" min="1" style="padding:6px;border:1px solid var(--border);border-radius:4px;font-size:12px">
      <input type="number" id="slpPerc_${i}" value="${s.percSconto||0}" min="0" max="100" step="0.5" style="padding:6px;border:1px solid var(--border);border-radius:4px;font-size:12px">
      <input type="text" id="slpLabel_${i}" value="${s.label||''}" placeholder="es. Sconto 7+ notti" style="padding:6px;border:1px solid var(--border);border-radius:4px;font-size:12px">
      <button class="btn-icon-sm" onclick="document.getElementById('slpRow_${i}').remove()">✕</button>
    </div>`;

  const durRow = (d,i) => `
    <div class="rate-season-row" id="durRow_${i}" style="grid-template-columns:1fr 1fr 36px">
      <input type="number" id="durSoglia_${i}" value="${d.soglia}" placeholder="Notti">
      <input type="number" id="durSconto_${i}" value="${d.sconto}" placeholder="%">
      <button class="btn-icon-sm" onclick="this.closest('.rate-season-row').remove()">✕</button>
    </div>`;

  const extraRow = (e,i) => `
    <div style="display:grid;grid-template-columns:1fr 24px 80px 90px 30px;gap:6px;align-items:center;margin-bottom:6px" id="extraRow_${i}">
      <input type="text"   id="eLabel_${i}"  value="${e.label}"  placeholder="Nome voce">
      <span style="font-size:11px;color:var(--text3);text-align:center">€</span>
      <input type="number" id="ePrice_${i}"  value="${e.prezzo}" step="0.5" min="0" placeholder="0">
      <select id="eUnita_${i}" style="padding:5px;border:1px solid var(--border);border-radius:4px;font-size:11px">
        ${['persona','notte','volta','kwh','mc'].map(u=>`<option value="${u}" ${e.unita===u?'selected':''}>${u}</option>`).join('')}
      </select>
      <button class="btn-icon-sm" onclick="this.closest('[id^=extraRow]').remove()">✕</button>
    </div>`;

  // Griglia override tariffe per camera
  const groups={};
  ROOMS.forEach(r=>{ if(!groups[r.g]) groups[r.g]=[]; groups[r.g].push(r); });
  const camGrpHtml = Object.entries(groups).map(([g,rooms])=>`
    <div class="conti-section">
      <div class="conti-section-title">${g} — override tariffa camera</div>
      <div style="font-size:10px;color:var(--text3);margin-bottom:8px">Lascia 0 per usare le tariffe da disposizione letti (solo albergo)</div>
      <div style="display:grid;grid-template-columns:1fr 90px 90px;gap:6px 10px;align-items:center">
        <span style="font-size:10px;font-weight:600;color:var(--text3)">Camera</span>
        <span style="font-size:10px;font-weight:600;color:var(--text3);text-align:center">€/notte</span>
        <span style="font-size:10px;font-weight:600;color:var(--text3);text-align:center">€/mese</span>
        ${rooms.map(r=>{
          const ov=cfg.tariffeCamere?.[r.id]||{};
          return `<label style="font-size:12px">${r.name}</label>
            <input type="number" id="tg_${r.id}" value="${ov.giornaliera||''}" placeholder="0" min="0" step="1"
              style="padding:5px;border:1px solid var(--border);border-radius:4px;font-size:12px;text-align:center;width:100%">
            <input type="number" id="tm_${r.id}" value="${ov.mensile||''}"     placeholder="0" min="0" step="10"
              style="padding:5px;border:1px solid var(--border);border-radius:4px;font-size:12px;text-align:center;width:100%">`;
        }).join('')}
      </div>
    </div>`).join('');

  return `
    <div class="conti-section">
      <div class="conti-section-title">Struttura</div>
      <div class="rate-grid">
        <div class="rate-field" style="grid-column:1/-1"><label>Nome / Ragione sociale</label><input id="cfgHotelName" value="${cfg.hotelName}"></div>
        <div class="rate-field" style="grid-column:1/-1"><label>Indirizzo (via e numero)</label><input id="cfgHotelAddr" value="${cfg.hotelAddress||''}"></div>
        <div class="rate-field"><label>CAP</label><input id="cfgHotelCAP" value="${cfg.hotelCAP||''}" maxlength="5" placeholder="es. 90100"></div>
        <div class="rate-field"><label>Comune</label><input id="cfgHotelComune" value="${cfg.hotelComune||''}" placeholder="es. Palermo"></div>
        <div class="rate-field"><label>Provincia (sigla)</label><input id="cfgHotelProv" value="${cfg.hotelProv||''}" maxlength="2" placeholder="es. PA"></div>
        <div class="rate-field"><label>P.IVA / Codice fiscale</label><input id="cfgHotelPiva" value="${cfg.pivaHotel||''}" placeholder="11 cifre senza IT"></div>
        <div class="rate-field"><label>Telefono</label><input id="cfgHotelTel" value="${cfg.hotelTel||''}"></div>
        <div class="rate-field" style="grid-column:1/-1">
          <label>URL Web App Apps Script (per aggiornamento immediato JSON)</label>
          <input id="cfgWebAppUrl" value="${cfg.webAppUrl||''}" placeholder="https://script.google.com/macros/s/xxx/exec" style="font-size:11px">
          <div style="font-size:10px;color:var(--text3);margin-top:3px">Apps Script → Distribuisci → App web → copia URL</div>
        </div>
      </div>
    </div>
    <div class="conti-section">
      <div class="conti-section-title">Tariffe base per disposizione letti</div>
      <div class="rate-grid">
        <div class="rate-field" style="grid-column:1/-1">
          <label>Aliquota IVA % (default 10% per alloggio, 22% extra, 4% agriturismo)</label>
          <input type="number" id="cfgAliquotaIVA" value="${cfg.aliquotaIVA||10}" min="0" max="100" step="1" style="max-width:120px">
        </div>
        <div class="rate-field"><label>Singola (1 letto singolo) €/notte</label><input type="number" id="tarS"  value="${t.s||35}"  step="0.5"></div>
        <div class="rate-field"><label>Matrimoniale uso singolo €/notte</label><input type="number" id="tarMS" value="${t.ms||38}" step="0.5"></div>
        <div class="rate-field"><label>Matrimoniale (2 pers.) €/notte</label><input type="number"   id="tarM"  value="${t.m||45}"  step="0.5"></div>
        <div class="rate-field"><label>Letto singolo aggiunto +€/pers.</label><input type="number"  id="tarAG" value="${t.ag||15}" step="0.5"></div>
      </div>
    </div>
    <div class="conti-section">
      <div class="conti-section-title">Stagionalità</div>
      <div style="font-size:10px;color:var(--text3);margin-bottom:8px">Formato: MM-DD–MM-DD</div>
      <div id="stagRows">${cfg.stagioni.map(stagRow).join('')}</div>
      <button class="btn-icon-add" onclick="addStagRow()" style="margin-top:6px">+</button>
    </div>
    <div class="conti-section">
      <div class="conti-section-title">Convenzioni (sconto %)</div>
      <div id="convRows">${cfg.convenzioni.map(convRow).join('')}</div>
      <button class="btn-icon-add" onclick="addConvRow()" style="margin-top:6px">+</button>
    </div>
    <div class="conti-section">
      <div class="conti-section-title">Sconti durata</div>
      <div id="durRows">${cfg.scontiDurata.map(durRow).join('')}</div>
      <button class="btn-icon-add" onclick="addDurRow()" style="margin-top:6px">+</button>
    </div>
    <div class="conti-section">
      <div class="conti-section-title">Sconti lungo periodo (appartamenti)</div>
      <div style="font-size:10px;color:var(--text3);margin-bottom:8px">Sconto % automatico sul pernottamento in modalità giornaliera</div>
      <div style="display:grid;grid-template-columns:80px 60px 1fr 28px;gap:6px;font-size:10px;font-weight:600;color:var(--text3);margin-bottom:4px">
        <span>Notti min</span><span>Sconto %</span><span>Etichetta</span><span></span>
      </div>
      <div id="slpRows">${(cfg.scontiLungoPeriodo||[]).map(slpRow).join('')}</div>
      <button class="btn-icon-add" onclick="addSlpRow()" style="margin-top:6px">+</button>
    </div>
    <div class="conti-section">
      <div class="conti-section-title">Voci extra</div>
      <div style="display:grid;grid-template-columns:1fr 24px 80px 90px 30px;gap:4px;font-size:10px;font-weight:600;color:var(--text3);margin-bottom:6px">
        <span>Voce</span><span></span><span>Prezzo</span><span>Unità</span><span></span>
      </div>
      <div id="extraRows">${(cfg.extra||[]).map(extraRow).join('')}</div>
      <button class="btn-icon-add" onclick="addExtraRow()" style="margin-top:6px">+</button>
    </div>
    ${camGrpHtml}`;
}

// ── Aggiungi righe dinamiche ──
function addStagRow() {
  const rows=document.getElementById('stagRows');
  const i=rows.children.length;
  rows.insertAdjacentHTML('beforeend',`<div class="rate-season-row" id="stagRow_${i}">
    <input type="text"   id="stagNome_${i}"   placeholder="Nome">
    <input type="text"   id="stagPeriod_${i}" placeholder="MM-DD–MM-DD">
    <input type="number" id="stagMolt_${i}"   value="1.0" step="0.1">
    <button class="btn-icon-sm" onclick="this.closest('.rate-season-row').remove()">✕</button>
  </div>`);
}
function addConvRow() {
  const rows=document.getElementById('convRows');
  const i=rows.children.length;
  rows.insertAdjacentHTML('beforeend',`<div class="rate-conv-row" id="convRow_${i}">
    <input type="text"   id="convNome_${i}" placeholder="Nome cliente">
    <input type="number" id="convVal_${i}"  placeholder="%" step="1" min="0" max="100">
    <button class="btn-icon-sm" onclick="this.closest('.rate-conv-row').remove()">✕</button>
  </div>`);
}
function addSlpRow() {
  const list = document.getElementById('slpRows');
  if (!list) return;
  const i = list.children.length;
  list.insertAdjacentHTML('beforeend', `<div id="slpRow_${i}" style="display:grid;grid-template-columns:80px 60px 1fr 28px;gap:6px;align-items:center;margin-bottom:6px">
    <input type="number" id="slpNotti_${i}" value="7" min="1" style="padding:6px;border:1px solid var(--border);border-radius:4px;font-size:12px">
    <input type="number" id="slpPerc_${i}" value="10" min="0" max="100" step="0.5" style="padding:6px;border:1px solid var(--border);border-radius:4px;font-size:12px">
    <input type="text" id="slpLabel_${i}" placeholder="es. Sconto 7+ notti" style="padding:6px;border:1px solid var(--border);border-radius:4px;font-size:12px">
    <button class="btn-icon-sm" onclick="this.parentElement.remove()">✕</button>
  </div>`);
}

function addDurRow() {
  const rows=document.getElementById('durRows');
  const i=rows.children.length;
  rows.insertAdjacentHTML('beforeend',`<div class="rate-season-row" id="durRow_${i}" style="grid-template-columns:1fr 1fr 36px">
    <input type="number" id="durSoglia_${i}" placeholder="Notti">
    <input type="number" id="durSconto_${i}" placeholder="%">
    <button class="btn-icon-sm" onclick="this.closest('.rate-season-row').remove()">✕</button>
  </div>`);
}
function addExtraRow() {
  const rows=document.getElementById('extraRows');
  const i=rows.children.length;
  rows.insertAdjacentHTML('beforeend',`<div style="display:grid;grid-template-columns:1fr 24px 80px 90px 30px;gap:6px;align-items:center;margin-bottom:6px" id="extraRow_${i}">
    <input type="text"   id="eLabel_${i}"  placeholder="Es. Parcheggio">
    <span style="font-size:11px;color:var(--text3);text-align:center">€</span>
    <input type="number" id="ePrice_${i}"  placeholder="0" step="0.5" min="0">
    <select id="eUnita_${i}" style="padding:5px;border:1px solid var(--border);border-radius:4px;font-size:11px">
      <option>persona</option><option>notte</option><option>volta</option><option>kwh</option><option>mc</option>
    </select>
    <button class="btn-icon-sm" onclick="this.closest('[id^=extraRow]').remove()">✕</button>
  </div>`);
}
function removeStagRow(i){ document.getElementById(`stagRow_${i}`)?.remove(); }
function removeConvRow(i){ document.getElementById(`convRow_${i}`)?.remove(); }
function removeDurRow(i) { document.getElementById(`durRow_${i}`)?.remove(); }

function saveContiSettings() {
  const cfg = loadBillSettings();
  cfg.aliquotaIVA  = parseFloat(document.getElementById('cfgAliquotaIVA')?.value || 10);
  cfg.hotelName    = document.getElementById('cfgHotelName')?.value  || cfg.hotelName;
  cfg.hotelAddress = document.getElementById('cfgHotelAddr')?.value  || '';
  cfg.hotelCAP     = document.getElementById('cfgHotelCAP')?.value   || '';
  cfg.hotelComune  = document.getElementById('cfgHotelComune')?.value || '';
  cfg.hotelProv    = document.getElementById('cfgHotelProv')?.value  || '';
  cfg.pivaHotel    = document.getElementById('cfgHotelPiva')?.value  || '';
  cfg.hotelTel     = document.getElementById('cfgHotelTel')?.value   || '';
  cfg.webAppUrl    = document.getElementById('cfgWebAppUrl')?.value  || '';

  // Tariffe disposizione
  cfg.tariffe = {
    s:  parseFloat(document.getElementById('tarS')?.value  || 35),
    ms: parseFloat(document.getElementById('tarMS')?.value || 38),
    m:  parseFloat(document.getElementById('tarM')?.value  || 45),
    ag: parseFloat(document.getElementById('tarAG')?.value || 15),
  };

  // Override per camera
  if (!cfg.tariffeCamere) cfg.tariffeCamere = {};
  ROOMS.forEach(r => {
    const g = parseFloat(document.getElementById(`tg_${r.id}`)?.value || 0);
    const m = parseFloat(document.getElementById(`tm_${r.id}`)?.value || 0);
    if (g > 0 || m > 0) cfg.tariffeCamere[r.id] = { giornaliera:g, mensile:m };
    else delete cfg.tariffeCamere[r.id];
  });

  // Stagioni
  cfg.stagioni = [];
  document.querySelectorAll('[id^=stagRow_]').forEach(row => {
    const i=row.id.split('_')[1];
    const nom=document.getElementById(`stagNome_${i}`)?.value?.trim();
    const per=document.getElementById(`stagPeriod_${i}`)?.value?.trim();
    const mol=parseFloat(document.getElementById(`stagMolt_${i}`)?.value||1);
    if (nom && per && per.includes('–')) {
      const [dal,al]=per.split('–');
      cfg.stagioni.push({ nome:nom, dal:dal.trim(), al:al.trim(), molt:mol });
    }
  });

  // Convenzioni
  cfg.convenzioni = [];
  document.querySelectorAll('[id^=convRow_]').forEach(row => {
    const i=row.id.split('_')[1];
    const nom=document.getElementById(`convNome_${i}`)?.value?.trim();
    const val=parseFloat(document.getElementById(`convVal_${i}`)?.value||0);
    if (nom && val > 0) cfg.convenzioni.push({ nome:nom, sconto:val });
  });

  // Sconti durata
  cfg.scontiDurata = [];
  document.querySelectorAll('[id^=durRow_]').forEach(row => {
    const i=row.id.split('_')[1];
    const sog=parseInt(document.getElementById(`durSoglia_${i}`)?.value||0);
    const sco=parseFloat(document.getElementById(`durSconto_${i}`)?.value||0);
    if (sog>0 && sco>0) cfg.scontiDurata.push({ soglia:sog, sconto:sco });
  });
  cfg.scontiDurata.sort((a,b)=>b.soglia-a.soglia);

  // Sconti lungo periodo (appartamenti)
  cfg.scontiLungoPeriodo = [];
  document.querySelectorAll('[id^=slpRow_]').forEach(row => {
    const i = row.id.split('_')[1];
    const n = parseInt(document.getElementById(`slpNotti_${i}`)?.value || 0);
    const p = parseFloat(document.getElementById(`slpPerc_${i}`)?.value || 0);
    const l = document.getElementById(`slpLabel_${i}`)?.value?.trim() || '';
    if (n > 0 && p > 0) cfg.scontiLungoPeriodo.push({ minNotti:n, percSconto:p, label:l });
  });
  cfg.scontiLungoPeriodo.sort((a,b) => b.minNotti - a.minNotti);

  // Extra voci
  cfg.extra = [];
  document.querySelectorAll('[id^=extraRow_]').forEach(row => {
    const i=row.id.split('_')[1];
    const label=document.getElementById(`eLabel_${i}`)?.value?.trim();
    const prezzo=parseFloat(document.getElementById(`ePrice_${i}`)?.value||0);
    const unita=document.getElementById(`eUnita_${i}`)?.value||'volta';
    if (label) cfg.extra.push({ id:`extra_${i}`, label, prezzo, unita });
  });

  saveBillSettings(cfg);
  closeContiSettings();
  renderContiTab(_contiTab);
  showToast('✓ Tariffe salvate','success');
}
