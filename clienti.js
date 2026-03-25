// ═══════════════════════════════════════════════════════════════════
// clienti.js — Anagrafica clienti Blip
// Dipende da: core.js, sync.js (DATABASE_SHEET_ID, apiFetch, dbGet)
// ═══════════════════════════════════════════════════════════════════

const BLIP_VER_CLIENTI = '1';

const CLIENTI_SHEET = 'CLIENTI';

// Colonne foglio CLIENTI
const CLI_COLS = {
  ID:          1,  // A  CLI-2026-XXXXXX
  NOME:        2,  // B
  EMAIL:       3,  // C
  TELEFONO:    4,  // D
  DOC_TIPO:    5,  // E  CI/Passaporto/Patente/Altro
  DOC_NUM:     6,  // F
  NAZIONALITA: 7,  // G
  DATA_NASCITA:8,  // H  gg/mm/aaaa
  NOTE:        9,  // I
  PRIMA_VISITA:10, // J  gg/mm/aaaa
  N_SOGGIORNI: 11, // K
  TS_CREAZIONE:12, // L
  TS_AGG:      13, // M
};

// Cache in-memory per la sessione
let _clientiCache = null;       // array di oggetti cliente
let _clientiCacheTs = 0;
const CLIENTI_TTL = 5 * 60 * 1000; // 5 minuti

// ─────────────────────────────────────────────────────────────────
// SETUP FOGLIO
// ─────────────────────────────────────────────────────────────────
let _clientiSheetReady = false;

async function ensureClientiSheet() {
  if (_clientiSheetReady || !DATABASE_SHEET_ID) return;
  _clientiSheetReady = true;
  try {
    const meta = await apiFetch(
      `https://sheets.googleapis.com/v4/spreadsheets/${DATABASE_SHEET_ID}?fields=sheets.properties.title`
    );
    const mj = await meta.json();
    const exists = (mj.sheets||[]).some(s => s.properties.title === CLIENTI_SHEET);
    if (!exists) {
      await apiFetch(`https://sheets.googleapis.com/v4/spreadsheets/${DATABASE_SHEET_ID}:batchUpdate`, {
        method: 'POST', headers: {'Content-Type':'application/json'},
        body: JSON.stringify({ requests: [{ addSheet: { properties: { title: CLIENTI_SHEET } } }] })
      });
    }
    const hd = await dbGet(`${CLIENTI_SHEET}!A1:M1`);
    if (!hd.values?.[0]?.[0]) {
      const url = `https://sheets.googleapis.com/v4/spreadsheets/${DATABASE_SHEET_ID}/values/${encodeURIComponent(CLIENTI_SHEET+'!A1:M1')}?valueInputOption=RAW`;
      await apiFetch(url, {
        method: 'PUT', headers: {'Content-Type':'application/json'},
        body: JSON.stringify({ values: [['CLIENTE_ID','NOME','EMAIL','TELEFONO','DOC_TIPO','DOC_NUM','NAZIONALITA','DATA_NASCITA','NOTE','PRIMA_VISITA','N_SOGGIORNI','TS_CREAZIONE','TS_AGG']] })
      });
    }
  } catch(e) { console.warn('[CLIENTI] ensure:', e.message); }
}

// ─────────────────────────────────────────────────────────────────
// LETTURA
// ─────────────────────────────────────────────────────────────────
function _rowToCliente(row, rowNum) {
  const get = i => (row[i] || '').toString().trim();
  return {
    id:          get(CLI_COLS.ID - 1),
    nome:        get(CLI_COLS.NOME - 1),
    email:       get(CLI_COLS.EMAIL - 1),
    telefono:    get(CLI_COLS.TELEFONO - 1),
    docTipo:     get(CLI_COLS.DOC_TIPO - 1),
    docNum:      get(CLI_COLS.DOC_NUM - 1),
    nazionalita: get(CLI_COLS.NAZIONALITA - 1),
    dataNascita: get(CLI_COLS.DATA_NASCITA - 1),
    note:        get(CLI_COLS.NOTE - 1),
    primaVisita: get(CLI_COLS.PRIMA_VISITA - 1),
    nSoggiorni:  parseInt(get(CLI_COLS.N_SOGGIORNI - 1)) || 0,
    tsCreazione: get(CLI_COLS.TS_CREAZIONE - 1),
    tsAgg:       get(CLI_COLS.TS_AGG - 1),
    dbRow:       rowNum,
  };
}

async function loadClienti(forceRefresh = false) {
  if (!forceRefresh && _clientiCache && Date.now() - _clientiCacheTs < CLIENTI_TTL) {
    return _clientiCache;
  }
  if (!DATABASE_SHEET_ID) return [];
  try {
    await ensureClientiSheet();
    const d = await dbGet(`${CLIENTI_SHEET}!A2:M9999`);
    const rows = d.values || [];
    _clientiCache = rows
      .map((row, i) => _rowToCliente(row, i + 2))
      .filter(c => c.id && c.nome);
    _clientiCacheTs = Date.now();
    return _clientiCache;
  } catch(e) {
    console.warn('[CLIENTI] load:', e.message);
    return _clientiCache || [];
  }
}

// ─────────────────────────────────────────────────────────────────
// RICERCA (usata dal modal prenotazione)
// ─────────────────────────────────────────────────────────────────
function cercaClienti(query) {
  if (!query || query.length < 2 || !_clientiCache) return [];
  const q = query.toLowerCase().trim();
  return _clientiCache
    .filter(c =>
      c.nome.toLowerCase().includes(q) ||
      c.email.toLowerCase().includes(q) ||
      c.telefono.replace(/\s/g,'').includes(q.replace(/\s/g,''))
    )
    .slice(0, 6);
}

// ─────────────────────────────────────────────────────────────────
// SCRITTURA
// ─────────────────────────────────────────────────────────────────
function _clienteToRow(c) {
  const now = nowISO();
  const oggi = new Date().toLocaleDateString('it-IT');
  return [
    c.id          || '',
    c.nome        || '',
    c.email       || '',
    c.telefono    || '',
    c.docTipo     || '',
    c.docNum      || '',
    c.nazionalita || 'IT',
    c.dataNascita || '',
    c.note        || '',
    c.primaVisita || oggi,
    String(c.nSoggiorni || 1),
    c.tsCreazione || now,
    now,
  ];
}

async function creaCliente(dati) {
  if (!DATABASE_SHEET_ID) throw new Error('DATABASE_SHEET_ID non configurato');
  await ensureClientiSheet();
  const anno = new Date().getFullYear();
  const cliente = {
    ...dati,
    id: genClienteId(anno),
    nSoggiorni: 1,
    primaVisita: new Date().toLocaleDateString('it-IT'),
    tsCreazione: nowISO(),
  };
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${DATABASE_SHEET_ID}/values/${encodeURIComponent(CLIENTI_SHEET)}:append?valueInputOption=RAW&insertDataOption=INSERT_ROWS`;
  const r = await apiFetch(url, {
    method: 'POST', headers: {'Content-Type':'application/json'},
    body: JSON.stringify({ values: [_clienteToRow(cliente)] })
  });
  if (!r.ok) throw new Error(`Creazione cliente fallita (${r.status})`);
  // Aggiorna cache
  const resp = await r.json();
  const m = (resp.updates?.updatedRange||'').match(/(\d+):/);
  if (m) cliente.dbRow = parseInt(m[1]);
  if (!_clientiCache) _clientiCache = [];
  _clientiCache.push(cliente);
  return cliente;
}

async function aggiornaCliente(cliente) {
  if (!DATABASE_SHEET_ID || !cliente.dbRow) throw new Error('dbRow mancante');
  const row = _clienteToRow(cliente);
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${DATABASE_SHEET_ID}/values/${encodeURIComponent(CLIENTI_SHEET+'!A'+cliente.dbRow+':M'+cliente.dbRow)}?valueInputOption=RAW`;
  const r = await apiFetch(url, {
    method: 'PUT', headers: {'Content-Type':'application/json'},
    body: JSON.stringify({ values: [row] })
  });
  if (!r.ok) throw new Error(`Aggiornamento cliente fallito (${r.status})`);
  // Aggiorna cache
  if (_clientiCache) {
    const idx = _clientiCache.findIndex(c => c.id === cliente.id);
    if (idx >= 0) _clientiCache[idx] = { ...cliente, tsAgg: nowISO() };
  }
}

async function incrementaSoggiorni(clienteId) {
  if (!_clientiCache) return;
  const c = _clientiCache.find(x => x.id === clienteId);
  if (!c || !c.dbRow) return;
  c.nSoggiorni = (c.nSoggiorni || 0) + 1;
  c.tsAgg = nowISO();
  try { await aggiornaCliente(c); } catch(e) { console.warn('[CLIENTI] incrementa:', e.message); }
}

// ─────────────────────────────────────────────────────────────────
// UI — Dropdown autocomplete nel modal prenotazione
// ─────────────────────────────────────────────────────────────────

// Stato corrente del modal
let _modalClienteId = null;      // id del cliente selezionato (null = nessuno / nuovo)
let _modalClienteOriginale = null; // snapshot al momento della selezione

function initAnagraficaModal() {
  // Precarica clienti in cache
  loadClienti().catch(() => {});
}

function onFNameInput(e) {
  const val = e.target.value;
  _modalClienteId = null; // deseleziona se l'utente modifica manualmente
  if (val.length < 2) { chiudiDropdownClienti(); return; }
  const risultati = cercaClienti(val);
  mostraDropdownClienti(risultati, val);
}

function mostraDropdownClienti(clienti, query) {
  let dd = document.getElementById('clientiDropdown');
  if (!dd) {
    dd = document.createElement('div');
    dd.id = 'clientiDropdown';
    dd.className = 'clienti-dropdown';
    const fNameEl = document.getElementById('fName');
    fNameEl.parentNode.style.position = 'relative';
    fNameEl.parentNode.appendChild(dd);
  }
  if (!clienti.length) {
    dd.innerHTML = `<div class="cdd-empty">Nessun cliente trovato — verrà creato nuovo</div>`;
    dd.style.display = 'block';
    return;
  }
  dd.innerHTML = clienti.map(c => `
    <div class="cdd-item" onclick="selezionaCliente('${c.id}')">
      <div class="cdd-nome">${escHtml(c.nome)}</div>
      <div class="cdd-meta">${escHtml(c.email||'')}${c.email&&c.telefono?' · ':''}${escHtml(c.telefono||'')}${c.nSoggiorni>1?` · ${c.nSoggiorni} soggiorni`:''}</div>
    </div>
  `).join('');
  dd.style.display = 'block';
}

function chiudiDropdownClienti() {
  const dd = document.getElementById('clientiDropdown');
  if (dd) dd.style.display = 'none';
}

function selezionaCliente(clienteId) {
  const c = _clientiCache?.find(x => x.id === clienteId);
  if (!c) return;
  _modalClienteId = clienteId;
  _modalClienteOriginale = { ...c };
  document.getElementById('fName').value = c.nome;
  chiudiDropdownClienti();
  // Mostra badge cliente selezionato
  aggiornaClienteBadge(c);
}

function aggiornaClienteBadge(c) {
  let badge = document.getElementById('clienteBadge');
  if (!badge) {
    badge = document.createElement('div');
    badge.id = 'clienteBadge';
    badge.className = 'cliente-badge';
    document.getElementById('fName').parentNode.appendChild(badge);
  }
  badge.innerHTML = `
    <span class="cb-icon">👤</span>
    <span class="cb-info">
      <b>${escHtml(c.nome)}</b>
      ${c.email ? `<span>${escHtml(c.email)}</span>` : ''}
      ${c.telefono ? `<span>${escHtml(c.telefono)}</span>` : ''}
      ${c.nSoggiorni > 1 ? `<span>${c.nSoggiorni} soggiorni</span>` : ''}
    </span>
    <button class="cb-clear" onclick="deselezionaCliente()" title="Rimuovi selezione">✕</button>
  `;
  badge.style.display = 'flex';
}

function deselezionaCliente() {
  _modalClienteId = null;
  _modalClienteOriginale = null;
  const badge = document.getElementById('clienteBadge');
  if (badge) badge.style.display = 'none';
}

function resetAnagraficaModal() {
  _modalClienteId = null;
  _modalClienteOriginale = null;
  chiudiDropdownClienti();
  const badge = document.getElementById('clienteBadge');
  if (badge) badge.style.display = 'none';
}

function preimpostaClienteModal(clienteId) {
  if (!clienteId || !_clientiCache) return;
  const c = _clientiCache.find(x => x.id === clienteId);
  if (c) selezionaCliente(c.id);
}

// ─────────────────────────────────────────────────────────────────
// GESTIONE SALVATAGGIO — chiamata da saveBooking in gantt.js
// ─────────────────────────────────────────────────────────────────
async function gestisciClienteAlSalvataggio(nomeDalModal) {
  // Caso 1: cliente già selezionato dal dropdown
  if (_modalClienteId) {
    const c = _clientiCache?.find(x => x.id === _modalClienteId);
    // Controlla se il nome nel modal è cambiato rispetto al cliente selezionato
    const nomeCambiato = c && nomeDalModal !== c.nome;
    if (nomeCambiato) {
      const scelta = await mostraDialogoCliente(nomeDalModal, c);
      if (scelta === 'nuovo') {
        return await _creaClienteMinimo(nomeDalModal);
      } else if (scelta === 'aggiorna') {
        const aggiornato = { ...c, nome: nomeDalModal };
        await aggiornaCliente(aggiornato);
        await incrementaSoggiorni(_modalClienteId);
        return _modalClienteId;
      } else {
        // ignora → usa cliente esistente senza modifiche
        await incrementaSoggiorni(_modalClienteId);
        return _modalClienteId;
      }
    } else {
      await incrementaSoggiorni(_modalClienteId);
      return _modalClienteId;
    }
  }

  // Caso 2: nessun cliente selezionato — cerca per nome esatto
  if (_clientiCache) {
    const existing = _clientiCache.find(c =>
      c.nome.toLowerCase() === nomeDalModal.toLowerCase()
    );
    if (existing) {
      // Trovato cliente con stesso nome — chiede se usare quello esistente
      const usa = confirm(`Trovato cliente esistente: "${existing.nome}"\nEmail: ${existing.email||'—'}  Soggiorni: ${existing.nSoggiorni}\n\nUsare questo cliente?`);
      if (usa) {
        await incrementaSoggiorni(existing.id);
        return existing.id;
      }
    }
  }

  // Caso 3: nuovo cliente (solo nome per ora)
  return await _creaClienteMinimo(nomeDalModal);
}

async function _creaClienteMinimo(nome) {
  try {
    const c = await creaCliente({ nome });
    return c.id;
  } catch(e) {
    console.warn('[CLIENTI] creazione minima:', e.message);
    return null;
  }
}

function mostraDialogoCliente(nuovoNome, clienteEsistente) {
  return new Promise(resolve => {
    const ov = document.createElement('div');
    ov.style.cssText = 'position:fixed;inset:0;z-index:9999;background:rgba(0,0,0,.5);display:flex;align-items:center;justify-content:center;padding:20px';
    ov.innerHTML = `
      <div style="background:var(--surface);border-radius:16px;padding:24px;max-width:360px;width:100%">
        <div style="font-weight:700;font-size:15px;margin-bottom:8px">Nome modificato</div>
        <div style="font-size:13px;color:var(--text2);margin-bottom:18px">
          Hai cambiato il nome da <b>${escHtml(clienteEsistente.nome)}</b> a <b>${escHtml(nuovoNome)}</b>.<br>
          Come vuoi procedere?
        </div>
        <div style="display:flex;flex-direction:column;gap:8px">
          <button class="btn primary" id="btnDlgAggiorna">Aggiorna cliente esistente</button>
          <button class="btn" id="btnDlgNuovo">Salva come nuovo cliente</button>
          <button class="btn" id="btnDlgIgnora" style="color:var(--text3)">Ignora modifica</button>
        </div>
      </div>`;
    document.body.appendChild(ov);
    const cleanup = (val) => { document.body.removeChild(ov); resolve(val); };
    ov.querySelector('#btnDlgAggiorna').onclick = () => cleanup('aggiorna');
    ov.querySelector('#btnDlgNuovo').onclick    = () => cleanup('nuovo');
    ov.querySelector('#btnDlgIgnora').onclick   = () => cleanup('ignora');
  });
}
