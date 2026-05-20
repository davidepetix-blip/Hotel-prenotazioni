// ═══════════════════════════════════════════════════════════════════
// clienti-panel.js — Anagrafica clienti con deduplicazione
// Blip Hotel Management — build 18.10.5
// Usa le variabili CSS del tema Blip (--bg, --surface, --accent, ecc.)
// ═══════════════════════════════════════════════════════════════════

const BLIP_VER_CP = '1';

let _cpClienti   = [];   // clienti dal foglio CLIENTI
let _cpFiltered  = [];
let _cpSelected  = null;
let _cpMergeSet  = new Set();
let _cpSort      = { key: 'nSoggiorni', dir: -1 };
let _cpFilter    = 'tutti';
let _cpEditMode  = null;  // id del cliente in modifica
let _cpNewClientFor = null; // id 'UNCENSED:...' per cui stiamo creando
let _cpQuery     = '';

// ── Inizializza ───────────────────────────────────────────────────
async function cpInit() {
  const app = document.getElementById('cp-app');
  if (!app) return;
  app.innerHTML = '<div style="padding:40px;text-align:center;color:var(--text3)">⏳ Caricamento clienti...</div>';

  // 1. Clienti censiti dal foglio CLIENTI
  let censiti = [];
  try { censiti = await loadClienti(true); } catch(e) { censiti = []; }

  // 2. Clienti "da censire" = prenotazioni senza clienteId, con nome unico
  const daCensire = _cpBuildUncensed(censiti);

  // 3. Unisci in lista unica
  _cpClienti = [
    ...censiti.map(c => ({ ...c, _tipo: 'censito' })),
    ...daCensire.map(c => ({ ...c, _tipo: 'daCensire' })),
  ];

  _cpFindDuplicates();
  _cpApplyFilter();
  cpRender();
}

// ── Costruisce lista "da censire" da bookings ─────────────────────
function _cpBuildUncensed(censiti) {
  if (typeof bookings === 'undefined') return [];
  const cNomi = new Set(censiti.map(c => _normNome(c.nome)));
  const seen  = new Set();
  const out   = [];
  bookings.forEach(b => {
    if (b.deleted || b.clienteId || !b.n) return;
    const k = _normNome(b.n);
    if (seen.has(k) || cNomi.has(k)) return;
    seen.add(k);
    // Aggrega tutte le prenotazioni con questo nome
    const bsGruppo = bookings.filter(x => !x.deleted && _normNome(x.n) === k);
    out.push({
      id:          'UNCENSED:' + k,
      nome:        b.n,
      email:       '',
      telefono:    '',
      docTipo:     '',
      docNum:      '',
      nazionalita: '',
      dataNascita: '',
      note:        '',
      primaVisita: _cpMinDate(bsGruppo),
      nSoggiorni:  bsGruppo.length,
      _bookingIds: bsGruppo.map(x => x.id),
      dbRow:       null,
    });
  });
  return out;
}

function _normNome(s) {
  return (s || '').toLowerCase().trim()
    .replace(/\s+/g, ' ')
    .replace(/[àáâ]/g,'a').replace(/[èéê]/g,'e')
    .replace(/[ìíî]/g,'i').replace(/[òóô]/g,'o').replace(/[ùúû]/g,'u');
}

function _cpMinDate(bs) {
  const dates = bs.map(b => b.s).filter(Boolean).sort();
  if (!dates.length) return '';
  const d = dates[0] instanceof Date ? dates[0] : new Date(dates[0]);
  return d.toLocaleDateString('it-IT');
}

// ── Trova duplicati per nome normalizzato ─────────────────────────
function _cpFindDuplicates() {
  const byName = {};
  _cpClienti.forEach(c => {
    const k = _normNome(c.nome);
    if (!byName[k]) byName[k] = [];
    byName[k].push(c.id);
  });
  const dupIds = new Set();
  Object.values(byName).filter(g => g.length > 1).forEach(g => g.forEach(id => dupIds.add(id)));
  _cpClienti.forEach(c => { c._isDup = dupIds.has(c.id); });
}

// ── Filtra e ordina ───────────────────────────────────────────────
function _cpApplyFilter() {
  let list = [..._cpClienti];
  if (_cpQuery) {
    const q = _cpQuery.toLowerCase();
    list = list.filter(c =>
      (c.nome||'').toLowerCase().includes(q) ||
      (c.email||'').toLowerCase().includes(q) ||
      (c.telefono||'').includes(q) ||
      (c.docNum||'').toLowerCase().includes(q)
    );
  }
  if (_cpFilter === 'censiti')    list = list.filter(c => c._tipo === 'censito');
  if (_cpFilter === 'daCensire')  list = list.filter(c => c._tipo === 'daCensire');
  if (_cpFilter === 'duplicati')  list = list.filter(c => c._isDup);
  if (_cpFilter === 'incompleti') list = list.filter(c => c._tipo === 'censito' && (!c.email || !c.docNum));

  list.sort((a, b) => {
    let va = a[_cpSort.key] ?? 0, vb = b[_cpSort.key] ?? 0;
    if (typeof va === 'string') va = va.toLowerCase();
    if (typeof vb === 'string') vb = vb.toLowerCase();
    return va < vb ? -_cpSort.dir : va > vb ? _cpSort.dir : 0;
  });
  _cpFiltered = list;
}

// ── Prenotazioni del cliente ──────────────────────────────────────
function _cpGetBookings(c) {
  if (typeof bookings === 'undefined') return [];
  if (c._tipo === 'daCensire') {
    const k = _normNome(c.nome);
    return bookings.filter(b => !b.deleted && _normNome(b.n) === k);
  }
  return bookings.filter(b => !b.deleted && b.clienteId === c.id);
}

// ── Render principale ─────────────────────────────────────────────
function cpRender() {
  const app = document.getElementById('cp-app');
  if (!app) return;

  const totCensiti   = _cpClienti.filter(c => c._tipo === 'censito').length;
  const totDaCensire = _cpClienti.filter(c => c._tipo === 'daCensire').length;
  const totDup       = _cpClienti.filter(c => c._isDup).length;
  const totIncompl   = _cpClienti.filter(c => c._tipo === 'censito' && (!c.email || !c.docNum)).length;

  const filterBtn = (id, label, count, color) => {
    const active = _cpFilter === id;
    return `<button onclick="cpSetFilter('${id}')" style="
      padding:5px 12px;border-radius:20px;font-size:11px;font-weight:600;cursor:pointer;
      border:1px solid ${active ? 'var(--accent)' : 'var(--border)'};
      background:${active ? 'var(--accent)' : 'var(--surface)'};
      color:${active ? '#fff' : (color || 'var(--text2)')};
    ">${label} <span style="opacity:.7">${count}</span></button>`;
  };

  const thStyle = 'padding:8px 12px;background:var(--surface2);border-bottom:1px solid var(--border);font-size:10px;font-weight:600;color:var(--text3);text-transform:uppercase;letter-spacing:.06em;text-align:left;white-space:nowrap;cursor:pointer;user-select:none;';
  const th = (key, label) => {
    const active = _cpSort.key === key;
    return `<th style="${thStyle}${active?'color:var(--accent)':''}" onclick="cpSort('${key}')">${label}${active?(_cpSort.dir===1?' ↑':' ↓'):''}</th>`;
  };

  app.innerHTML = `
    <div style="display:flex;align-items:center;gap:10px;padding:12px 16px;border-bottom:1px solid var(--border);background:var(--surface)">
      <span style="font-size:14px;font-weight:700;color:var(--text)">👥 Anagrafica Clienti</span>
      <span style="font-size:11px;color:var(--text3)">${totCensiti} censiti · ${totDaCensire} da censire</span>
      <div style="flex:1"></div>
      <button class="btn" onclick="cpEsportaCSV()">↓ CSV</button>
      <button class="btn" onclick="cpInit()">↺</button>
      <button class="btn" onclick="closeClientiPanel()">✕ Chiudi</button>
    </div>

    <div style="display:flex;gap:6px;align-items:center;padding:10px 16px;border-bottom:1px solid var(--border);background:var(--surface);flex-wrap:wrap">
      <div style="position:relative;flex:1;max-width:300px">
        <span style="position:absolute;left:8px;top:50%;transform:translateY(-50%);color:var(--text3);font-size:12px">🔍</span>
        <input id="cp-search" type="text" placeholder="Cerca nome, email, documento…"
          value="${_cpQuery.replace(/"/g,'&quot;')}"
          oninput="cpSearch(this.value)"
          style="width:100%;padding:6px 8px 6px 28px;border:1px solid var(--border);border-radius:var(--radius);font-size:12px;background:var(--bg);color:var(--text);outline:none">
      </div>
      ${filterBtn('tutti','Tutti',_cpClienti.length)}
      ${filterBtn('censiti','✓ Censiti',totCensiti)}
      ${filterBtn('daCensire','⚠ Da censire',totDaCensire,'var(--danger)')}
      ${filterBtn('duplicati','≡ Duplicati',totDup,'#e67e22')}
      ${filterBtn('incompleti','? Incompleti',totIncompl)}
    </div>

    <div style="display:flex;height:calc(100vh - 110px)">
      <div style="flex:1;overflow:auto">
        ${_cpFiltered.length === 0 ? `
          <div style="padding:50px;text-align:center;color:var(--text3)">
            <div style="font-size:32px;margin-bottom:10px">👤</div>
            <div style="font-size:13px;color:var(--text2)">Nessun cliente trovato</div>
          </div>
        ` : `
        <table style="width:100%;border-collapse:collapse">
          <thead><tr>
            <th style="${thStyle}width:28px"><input type="checkbox" onchange="cpSelectAll(this.checked)"></th>
            ${th('nome','Cliente')}
            ${th('_tipo','Stato')}
            ${th('nSoggiorni','Soggiorni')}
            ${th('primaVisita','Prima visita')}
            <th style="${thStyle}">Dati</th>
            <th style="${thStyle}">Prenotazioni</th>
            <th style="${thStyle}"></th>
          </tr></thead>
          <tbody>
            ${_cpFiltered.map(c => `
              <tr onclick="cpSelectClient('${c.id}')" style="border-bottom:1px solid var(--border);cursor:pointer;background:${_cpSelected?.id===c.id?'var(--accent-light)':c._isDup?'#fff8f0':''}">
                <td style="padding:8px 12px" onclick="event.stopPropagation()">
                  <input type="checkbox" ${_cpMergeSet.has(c.id)?'checked':''} onchange="cpToggleMerge('${c.id}',this.checked)">
                </td>
                <td style="padding:8px 12px">
                  <div style="font-weight:600;font-size:13px;color:var(--text)">${c.nome}${c._isDup?'<span style="margin-left:5px;padding:1px 6px;border-radius:4px;font-size:9px;font-weight:700;background:#fff3cd;color:#856404">DUP</span>':''}</div>
                  ${c.email?`<div style="font-size:11px;color:var(--text3)">${c.email}</div>`:''}
                </td>
                <td style="padding:8px 12px">
                  ${c._tipo==='censito'
                    ? `<span style="padding:2px 8px;border-radius:20px;font-size:10px;font-weight:600;background:var(--success-light);color:var(--success)">✓ Censito</span>`
                    : `<span style="padding:2px 8px;border-radius:20px;font-size:10px;font-weight:600;background:var(--danger-light);color:var(--danger)">⚠ Da censire</span>`}
                </td>
                <td style="padding:8px 12px;font-size:12px;font-weight:600;color:var(--accent)">${c.nSoggiorni}</td>
                <td style="padding:8px 12px;font-size:11px;color:var(--text3)">${c.primaVisita||'—'}</td>
                <td style="padding:8px 12px">
                  ${c._tipo==='censito'
                    ? (() => {
                        const n = [c.email,c.telefono,c.docNum,c.dataNascita].filter(Boolean).length;
                        return `<span style="padding:2px 8px;border-radius:20px;font-size:10px;background:${n===4?'var(--success-light)':'var(--surface2)'};color:${n===4?'var(--success)':'var(--text3)'}">${n}/4</span>`;
                      })()
                    : `<span style="color:var(--text3);font-size:11px">—</span>`}
                </td>
                <td style="padding:8px 12px">
                  ${_cpGetBookings(c).length > 0
                    ? `<span style="padding:2px 8px;border-radius:20px;font-size:10px;background:var(--surface2);color:var(--text2)">${_cpGetBookings(c).length}</span>`
                    : `<span style="color:var(--text3);font-size:11px">—</span>`}
                </td>
                <td style="padding:8px 12px">
                  ${c._tipo==='daCensire'
                    ? `<button class="btn" style="font-size:11px;padding:3px 10px;height:auto;background:var(--accent);color:#fff;border-color:var(--accent)" onclick="event.stopPropagation();cpCensisci('${c.id}')">+ Censisci</button>`
                    : `<button class="btn" style="font-size:11px;padding:3px 10px;height:auto" onclick="event.stopPropagation();cpEditClient('${c.id}')">✎</button>`}
                </td>
              </tr>
            `).join('')}
          </tbody>
        </table>
        `}
      </div>

      ${_cpSelected ? `
        <div id="cp-detail-panel" class="cp-detail-panel">
          ${cpDetailHTML(_cpSelected)}
        </div>
      ` : ''}
    </div>

    ${_cpMergeSet.size >= 2 ? `
      <div style="position:fixed;bottom:0;left:0;right:0;background:var(--surface);border-top:2px solid var(--accent);padding:10px 16px;display:flex;align-items:center;gap:10px;z-index:100;box-shadow:var(--shadow-md)">
        <span style="background:var(--accent);color:#fff;border-radius:20px;padding:3px 10px;font-size:11px;font-weight:700">${_cpMergeSet.size}</span>
        <span style="font-size:12px;color:var(--text2);flex:1">clienti selezionati per unificazione</span>
        <button class="btn" onclick="cpClearMerge()">✕ Annulla</button>
        <button class="btn" style="background:var(--accent);color:#fff;border-color:var(--accent)" onclick="cpMerge()">⚡ Unifica</button>
      </div>
    ` : ''}
  `;
}

// ── Detail panel ──────────────────────────────────────────────────
function cpDetailHTML(c) {
  if (_cpEditMode === c.id) return cpEditFormHTML(c);
  const bs = _cpGetBookings(c);
  const row = (k, v) => v ? `<div style="display:flex;justify-content:space-between;margin-bottom:6px"><span style="font-size:11px;color:var(--text3)">${k}</span><span style="font-size:11px;color:var(--text);font-weight:500;text-align:right;max-width:200px">${v}</span></div>` : '';

  return `
    <div style="padding:14px 16px;border-bottom:1px solid var(--border);display:flex;justify-content:space-between;align-items:start">
      <div>
        <div style="font-size:15px;font-weight:700;color:var(--text)">${c.nome}</div>
        <div style="font-size:10px;color:var(--text3);font-family:monospace;margin-top:2px">${c.id}</div>
      </div>
      <button class="btn" style="padding:2px 8px;height:auto;font-size:11px" onclick="_cpSelected=null;cpRender()">✕</button>
    </div>

    ${c._tipo === 'daCensire' ? `
      <div style="margin:12px 16px;padding:10px;background:var(--danger-light);border-radius:var(--radius);border:1px solid rgba(192,57,43,.15)">
        <div style="font-size:12px;font-weight:600;color:var(--danger);margin-bottom:4px">⚠ Cliente non censito</div>
        <div style="font-size:11px;color:var(--text2);margin-bottom:8px">Questo nome appare nelle prenotazioni ma non ha una scheda anagrafica.</div>
        <button class="btn" style="background:var(--accent);color:#fff;border-color:var(--accent);font-size:11px;padding:4px 12px;height:auto;width:100%"
          onclick="cpCensisci('${c.id}')">+ Crea scheda anagrafica</button>
      </div>
    ` : ''}

    ${c._isDup ? `
      <div style="margin:12px 16px;padding:10px;background:#fff8f0;border-radius:var(--radius);border:1px solid #fde8c8">
        <div style="font-size:11px;font-weight:600;color:#856404;margin-bottom:6px">≡ Possibile duplicato</div>
        <button class="btn" style="background:#f59e0b;color:#fff;border-color:#f59e0b;font-size:11px;padding:4px 12px;height:auto;width:100%"
          onclick="cpPrepareMergeGroup('${c.id}')">⚡ Prepara unificazione</button>
      </div>
    ` : ''}

    <div style="padding:12px 16px;border-bottom:1px solid var(--border)">
      <div style="font-size:10px;font-weight:600;color:var(--text3);text-transform:uppercase;letter-spacing:.06em;margin-bottom:8px">Dati anagrafici</div>
      ${row('Email', c.email)}
      ${row('Telefono', c.telefono)}
      ${row('Documento', c.docTipo ? c.docTipo+' '+c.docNum : c.docNum)}
      ${row('Nazionalità', c.nazionalita)}
      ${row('Data nascita', c.dataNascita)}
      ${row('Note', c.note)}
      ${!c.email && !c.docNum && c._tipo === 'censito' ? '<div style="font-size:11px;color:var(--text3)">Dati mancanti — modifica per completare</div>' : ''}
    </div>

    <div style="padding:12px 16px;border-bottom:1px solid var(--border)">
      <div style="font-size:10px;font-weight:600;color:var(--text3);text-transform:uppercase;letter-spacing:.06em;margin-bottom:8px">Soggiorni (${bs.length})</div>
      ${bs.length === 0
        ? '<div style="font-size:11px;color:var(--text3)">Nessuna prenotazione collegata</div>'
        : bs.sort((a,b)=>new Date(b.s)-new Date(a.s)).slice(0,8).map(b => `
          <div onclick="selBook(${b.id},null);closeClientiPanel()" style="padding:7px 8px;border-radius:var(--radius);margin-bottom:4px;background:var(--surface2);border:1px solid var(--border);cursor:pointer;transition:border-color .1s" onmouseover="this.style.borderColor='var(--accent)'" onmouseout="this.style.borderColor='var(--border)'">
            <div style="display:flex;justify-content:space-between">
              <span style="font-size:11px;font-weight:600">${fmt ? fmt(b.s) : b.s} → ${fmt ? fmt(b.e) : b.e}</span>
              <span style="font-size:10px;color:var(--text3)">${Math.round((new Date(b.e)-new Date(b.s))/86400000)}n</span>
            </div>
            <div style="font-size:10px;color:var(--text3);margin-top:2px">Camera ${typeof roomName==='function'?roomName(b.r):b.r} · ${b.d||'—'}</div>
          </div>
        `).join('')}
    </div>

    <div style="padding:12px 16px;display:flex;gap:6px;flex-wrap:wrap">
      ${c._tipo==='censito'
        ? `<button class="btn" onclick="cpEditClient('${c.id}')">✎ Modifica</button>`
        : `<button class="btn" style="background:var(--accent);color:#fff;border-color:var(--accent)" onclick="cpCensisci('${c.id}')">+ Censisci</button>`}
      <button class="btn" onclick="cpToggleMerge('${c.id}',true)">+ Unifica</button>
    </div>
  `;
}

// ── Actions ───────────────────────────────────────────────────────
function cpSearch(q)    { _cpQuery = q; _cpApplyFilter(); cpRender(); }
function cpSetFilter(f) { _cpFilter = f; _cpApplyFilter(); cpRender(); }
function cpSort(key)    { _cpSort = _cpSort.key===key?{key,dir:-_cpSort.dir}:{key,dir:-1}; _cpApplyFilter(); cpRender(); }

function cpSelectClient(id) {
  _cpSelected = _cpClienti.find(c => c.id === id) || null;
  cpRender();
  // Su mobile scrolla il panel in cima
  if (window.innerWidth < 700) {
    setTimeout(() => document.getElementById('cp-detail-panel')?.scrollIntoView(), 50);
  }
}
function cpToggleMerge(id, on) { if(on) _cpMergeSet.add(id); else _cpMergeSet.delete(id); cpRender(); }
function cpSelectAll(on)       { if(on) _cpFiltered.forEach(c=>_cpMergeSet.add(c.id)); else _cpMergeSet.clear(); cpRender(); }
function cpClearMerge()        { _cpMergeSet.clear(); cpRender(); }

function cpPrepareMergeGroup(id) {
  const c = _cpClienti.find(x=>x.id===id);
  if (!c) return;
  const k = _normNome(c.nome);
  _cpClienti.filter(x=>_normNome(x.nome)===k).forEach(x=>_cpMergeSet.add(x.id));
  cpRender();
}

function cpCensisci(id) {
  _cpSelected = _cpClienti.find(c => c.id === id) || null;
  _cpNewClientFor = id; // indica che stiamo creando ex-novo per questo "da censire"
  _cpEditMode = id;
  cpRender();
  setTimeout(() => {
    document.getElementById('cp-edit-nome')?.focus();
    if (window.innerWidth < 700) document.getElementById('cp-detail-panel')?.scrollIntoView({ behavior:'smooth' });
  }, 80);
}

function cpEditClient(id) {
  _cpSelected = _cpClienti.find(c => c.id === id) || null;
  _cpEditMode = id;
  cpRender();
  setTimeout(() => {
    document.getElementById('cp-edit-nome')?.focus();
    if (window.innerWidth < 700) document.getElementById('cp-detail-panel')?.scrollIntoView();
  }, 50);
}

async function cpSaveClient() {
  const btn = document.getElementById('cp-save-btn');
  if (btn) { btn.disabled = true; btn.textContent = '⏳ Salvataggio...'; }
  try {
    const nome      = document.getElementById('cp-edit-nome')?.value?.trim() || '';
    const email     = document.getElementById('cp-edit-email')?.value?.trim() || '';
    const telefono  = document.getElementById('cp-edit-tel')?.value?.trim() || '';
    const docTipo   = document.getElementById('cp-edit-docTipo')?.value || '';
    const docNum    = document.getElementById('cp-edit-docNum')?.value?.trim() || '';
    const naz       = document.getElementById('cp-edit-naz')?.value?.trim() || '';
    const dataN     = document.getElementById('cp-edit-dataN')?.value?.trim() || '';
    const note      = document.getElementById('cp-edit-note')?.value?.trim() || '';

    if (!nome) { alert('Il nome è obbligatorio.'); if(btn){btn.disabled=false;btn.textContent='💾 Salva';} return; }

    if (_cpNewClientFor) {
      // Crea nuovo cliente
      const created = await creaCliente({ nome, email, telefono, docTipo, docNum, nazionalita:naz, dataNascita:dataN, note });
      // Collega tutte le prenotazioni del "da censire" al nuovo cliente
      const k = _normNome(nome);
      const uncBs = typeof bookings !== 'undefined'
        ? bookings.filter(b => !b.deleted && !b.clienteId && _normNome(b.n) === k)
        : [];
      for (const b of uncBs) {
        b.clienteId = created.id;
        if (typeof dbUpdateRow === 'function') {
          await dbUpdateRow(b.dbRow, bookingToDbRow(b, b.fonte || 'app'));
        }
      }
      if (typeof showToast === 'function') showToast('✅ Scheda anagrafica creata — ' + uncBs.length + ' prenotaz. collegate', 'success');
    } else {
      // Aggiorna cliente esistente
      const c = _cpClienti.find(x => x.id === _cpEditMode);
      if (!c) return;
      Object.assign(c, { nome, email, telefono, docTipo, docNum, nazionalita:naz, dataNascita:dataN, note });
      if (typeof aggiornaCliente === 'function') await aggiornaCliente(c);
      if (typeof showToast === 'function') showToast('✅ Cliente aggiornato', 'success');
    }

    _cpEditMode = null;
    _cpNewClientFor = null;
    await cpInit();
  } catch(e) {
    if (typeof showToast === 'function') showToast('❌ ' + e.message, 'error');
    if (btn) { btn.disabled = false; btn.textContent = '💾 Salva'; }
  }
}

function cpCancelEdit() {
  _cpEditMode = null;
  _cpNewClientFor = null;
  cpRender();
}

async function cpMerge() {
  if (_cpMergeSet.size < 2) return;
  const ids     = [..._cpMergeSet];
  const clienti = ids.map(id=>_cpClienti.find(c=>c.id===id)).filter(c=>c?._tipo==='censito');
  if (clienti.length < 1) { alert('Seleziona almeno un cliente censito come destinazione.'); return; }
  const master  = clienti.reduce((b,c)=>c.nSoggiorni>b.nSoggiorni?c:b);
  const toMerge = clienti.filter(c=>c.id!==master.id);
  if (!confirm(`Unificare in "${master.nome}"?\nLe prenotazioni di ${toMerge.map(c=>c.nome).join(', ')} verranno attribuite a ${master.nome}.`)) return;
  try {
    if (typeof showLoading==='function') showLoading('Unificazione...');
    for (const dup of toMerge) {
      const dupBs = typeof bookings!=='undefined' ? bookings.filter(b=>b.clienteId===dup.id) : [];
      for (const b of dupBs) {
        b.clienteId = master.id;
        if (typeof dbUpdateRow==='function') await dbUpdateRow(b.dbRow, bookingToDbRow(b, b.fonte||'app'));
      }
    }
    _cpMergeSet.clear();
    if (typeof hideLoading==='function') hideLoading();
    if (typeof showToast==='function') showToast('✅ Clienti unificati','success');
    await cpInit();
  } catch(e) {
    if (typeof hideLoading==='function') hideLoading();
    if (typeof showToast==='function') showToast('❌ '+e.message,'error');
  }
}

function cpEsportaCSV() {
  const header = 'ID,Nome,Stato,Email,Telefono,Doc,Nazionalità,Data nascita,Prima visita,Soggiorni';
  const rows = _cpFiltered.map(c =>
    [c.id,c.nome,c._tipo==='censito'?'Censito':'Da censire',c.email,c.telefono,
     (c.docTipo||'')+' '+(c.docNum||''),c.nazionalita,c.dataNascita,c.primaVisita,c.nSoggiorni]
    .map(v=>`"${(v||'').toString().replace(/"/g,'""')}"`).join(',')
  );
  const a = document.createElement('a');
  a.href = 'data:text/csv;charset=utf-8,' + encodeURIComponent([header,...rows].join('\n'));
  a.download = `clienti_${new Date().toISOString().slice(0,10)}.csv`;
  a.click();
}

// ── Panel mount ───────────────────────────────────────────────────
function openClientiPanel() {
  let panel = document.getElementById('cp-panel');
  if (!panel) {
    panel = document.createElement('div');
    panel.id = 'cp-panel';
    panel.style.cssText = 'position:fixed;inset:0;z-index:300;background:var(--bg);overflow:hidden;display:flex;flex-direction:column';
    panel.innerHTML = '<div id="cp-app" style="flex:1;overflow:hidden;display:flex;flex-direction:column"></div>';
    document.body.appendChild(panel);
    document.addEventListener('keydown', e=>{ if(e.key==='Escape') closeClientiPanel(); }, {once:true});
  }
  panel.style.display = 'flex';
  cpInit();
}

function closeClientiPanel() {
  const p = document.getElementById('cp-panel');
  if (p) p.remove();
  // Rimuovi il parametro dall'URL senza ricaricare la pagina
  const url = new URL(window.location.href);
  if (url.searchParams.has('clienti')) {
    url.searchParams.delete('clienti');
    history.replaceState(null, '', url.toString());
  }
}

// ── Auto-open da URL ──────────────────────────────────────────────
// Supporta: ?clienti oppure #clienti
// Esempio: https://davidepetix-blip.github.io/Hotel-prenotazioni/?clienti
function _cpCheckUrlAutoOpen() {
  const hasParam = new URL(window.location.href).searchParams.has('clienti');
  const hasHash  = window.location.hash === '#clienti';
  if (!hasParam && !hasHash) return;

  // Aspetta che l'app sia pronta (login completato)
  const tryOpen = () => {
    // Se l'utente è loggato (accessToken disponibile)
    if (typeof accessToken !== 'undefined' && accessToken) {
      openClientiPanel();
    } else {
      // Riprova ogni 500ms finché non è loggato
      setTimeout(tryOpen, 500);
    }
  };
  setTimeout(tryOpen, 300);
}

// Esegui al caricamento del file
if (typeof window !== 'undefined') {
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', _cpCheckUrlAutoOpen);
  } else {
    _cpCheckUrlAutoOpen();
  }
}
