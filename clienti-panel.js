// ═══════════════════════════════════════════════════════════════════
// clienti-panel.js — Anagrafica clienti con deduplicazione
// Blip Hotel Management — build 18.10.5
// ═══════════════════════════════════════════════════════════════════


// ─── STANDALONE CLIENTI PANEL ─────────────────────────────────────
// Usa apiFetch, loadClienti, bookings, tutti già definiti in Blip.
// Se incluso come file autonomo per test, usa mock data.
// ──────────────────────────────────────────────────────────────────

// State
let _cpClienti = [];
let _cpFiltered = [];
let _cpSelected = null;
let _cpMergeSet = new Set();
let _cpSort = { key: 'nSoggiorni', dir: -1 };
let _cpFilter = 'tutti';
let _cpQuery = '';
let _cpLoading = false;
let _cpDuplicateGroups = [];

// ── Inizializza ───────────────────────────────────────────────────
async function cpInit() {
  _cpLoading = true;
  cpRender();
  try {
    _cpClienti = await loadClienti(true);
    _cpFindDuplicates();
    _cpApplyFilter();
    cpRender();
  } catch(e) {
    console.error('[CP]', e);
  }
  _cpLoading = false;
  cpRender();
}

// ── Trova duplicati per nome normalizzato ─────────────────────────
function _cpFindDuplicates() {
  const norm = s => s.toLowerCase().trim()
    .replace(/\s+/g, ' ')
    .replace(/[àáâ]/g,'a').replace(/[èéê]/g,'e')
    .replace(/[ìíî]/g,'i').replace(/[òóô]/g,'o').replace(/[ùúû]/g,'u');

  const byName = {};
  _cpClienti.forEach(c => {
    const k = norm(c.nome);
    if (!byName[k]) byName[k] = [];
    byName[k].push(c.id);
  });

  _cpDuplicateGroups = Object.values(byName).filter(g => g.length > 1);
  const dupIds = new Set(_cpDuplicateGroups.flat());

  _cpClienti.forEach(c => {
    c._isDup = dupIds.has(c.id);
    c._dupGroup = _cpDuplicateGroups.find(g => g.includes(c.id)) || null;
  });
}

// ── Filtra e ordina ───────────────────────────────────────────────
function _cpApplyFilter() {
  let list = [..._cpClienti];
  if (_cpQuery) {
    const q = _cpQuery.toLowerCase();
    list = list.filter(c =>
      c.nome.toLowerCase().includes(q) ||
      c.email.toLowerCase().includes(q) ||
      c.telefono.includes(q) ||
      c.docNum.toLowerCase().includes(q)
    );
  }
  if (_cpFilter === 'duplicati') list = list.filter(c => c._isDup);
  if (_cpFilter === 'completi')  list = list.filter(c => c.email && c.telefono && c.docNum);
  if (_cpFilter === 'incompleti')list = list.filter(c => !c.email || !c.telefono || !c.docNum);

  list.sort((a, b) => {
    let va = a[_cpSort.key] || 0, vb = b[_cpSort.key] || 0;
    if (typeof va === 'string') va = va.toLowerCase();
    if (typeof vb === 'string') vb = vb.toLowerCase();
    return va < vb ? -_cpSort.dir : va > vb ? _cpSort.dir : 0;
  });

  _cpFiltered = list;
}

// ── Prenotazioni del cliente ──────────────────────────────────────
function _cpGetBookings(clienteId) {
  if (typeof bookings === 'undefined') return [];
  return bookings.filter(b => b.clienteId === clienteId && !b.deleted);
}

function _cpGetBookingsAnyDup(c) {
  if (!c._dupGroup) return _cpGetBookings(c.id);
  return c._dupGroup.flatMap(id => _cpGetBookings(id));
}

// ── Calcola totale stimato ────────────────────────────────────────
function _cpTotal(clienteId) {
  const bs = _cpGetBookings(clienteId);
  if (!bs.length) return 0;
  return bs.reduce((s, b) => {
    const notti = Math.round((new Date(b.e) - new Date(b.s)) / 86400000);
    const cfg = typeof loadBillSettings === 'function' ? loadBillSettings() : {};
    const tariffa = cfg.tariffe?.m || 0;
    return s + notti * tariffa;
  }, 0);
}

// ── Render principale ─────────────────────────────────────────────
function cpRender() {
  const app = document.getElementById('cp-app');
  if (!app) return;

  const totClienti = _cpClienti.length;
  const totDup = _cpDuplicateGroups.length;
  const totIncompleti = _cpClienti.filter(c => !c.email || !c.docNum).length;
  const totSoggiorni = _cpClienti.reduce((s, c) => s + c.nSoggiorni, 0);

  app.innerHTML = `
    ${_cpLoading ? '<div class="loading-bar"></div>' : ''}
    <div class="header">
      <div>
        <div class="header h1" style="font-size:15px;font-weight:700">👥 Anagrafica Clienti</div>
        <div class="sub">${totClienti} clienti · ${totSoggiorni} soggiorni totali</div>
      </div>
      <div class="spacer"></div>
      <button class="btn btn-ghost btn-sm" onclick="cpInit()">↺ Ricarica</button>
      <button class="btn btn-primary btn-sm" onclick="cpEsportaCSV()">↓ CSV</button>
    </div>

    <div class="stats">
      <div class="stat"><div class="val">${totClienti}</div><div class="lbl">Clienti</div></div>
      <div class="stat"><div class="val" style="color:var(--danger)">${totDup}</div><div class="lbl">Gruppi dup.</div></div>
      <div class="stat"><div class="val" style="color:var(--warning)">${totIncompleti}</div><div class="lbl">Incompleti</div></div>
      <div class="stat"><div class="val" style="color:var(--success)">${totSoggiorni}</div><div class="lbl">Soggiorni</div></div>
    </div>

    <div class="toolbar">
      <div class="search-wrap">
        <span class="search-icon">🔍</span>
        <input id="cp-search" type="text" placeholder="Cerca per nome, email, documento..."
          value="${_cpQuery}" oninput="cpSearch(this.value)">
      </div>
      <button class="filter-btn ${_cpFilter==='tutti'?'active':''}" onclick="cpSetFilter('tutti')">Tutti (${_cpClienti.length})</button>
      <button class="filter-btn ${_cpFilter==='duplicati'?'active':''}" onclick="cpSetFilter('duplicati')" style="color:${_cpFilter!=='duplicati'?'var(--danger)':''}">
        ⚠ Duplicati (${_cpClienti.filter(c=>c._isDup).length})
      </button>
      <button class="filter-btn ${_cpFilter==='incompleti'?'active':''}" onclick="cpSetFilter('incompleti')">Incompleti (${totIncompleti})</button>
      <button class="filter-btn ${_cpFilter==='completi'?'active':''}" onclick="cpSetFilter('completi')">Completi</button>
    </div>

    <div class="split">
      <div class="list-side">
        ${_cpFiltered.length === 0 ? `
          <div class="empty">
            <div class="empty-icon">👤</div>
            <div class="empty-title">${_cpLoading ? 'Caricamento...' : 'Nessun cliente trovato'}</div>
          </div>
        ` : `
        <table>
          <thead>
            <tr>
              <th style="width:28px"><input type="checkbox" class="checkbox" onchange="cpSelectAll(this.checked)"></th>
              ${cpTh('nome','Cliente')}
              ${cpTh('nSoggiorni','Soggiorni')}
              ${cpTh('primaVisita','Prima visita')}
              <th>Dati</th>
              <th>Prenotazioni</th>
              <th></th>
            </tr>
          </thead>
          <tbody>
            ${_cpFiltered.map(c => cpRow(c)).join('')}
          </tbody>
        </table>
        `}
      </div>
      <div class="detail-side ${_cpSelected ? 'open' : ''}" id="cp-detail">
        ${_cpSelected ? cpDetail(_cpSelected) : ''}
      </div>
    </div>

    <div class="merge-bar ${_cpMergeSet.size >= 2 ? 'visible' : ''}" id="cp-merge-bar">
      <span class="merge-count">${_cpMergeSet.size}</span>
      <span class="merge-text">clienti selezionati per unificazione</span>
      <button class="btn btn-ghost btn-sm" onclick="cpClearMerge()">✕ Annulla</button>
      <button class="btn btn-primary btn-sm" onclick="cpMerge()">⚡ Unifica</button>
    </div>
  `;
}

function cpTh(key, label) {
  const active = _cpSort.key === key;
  const arrow = active ? (_cpSort.dir === 1 ? ' ↑' : ' ↓') : '';
  return `<th class="${active?'sorted':''}" onclick="cpSort('${key}')">${label}<span class="sort-arrow">${arrow}</span></th>`;
}

function cpRow(c) {
  const bs = _cpGetBookings(c.id);
  const isSelected = _cpSelected?.id === c.id;
  const inMerge = _cpMergeSet.has(c.id);
  const completeness = [c.email, c.telefono, c.docNum, c.dataNascita].filter(Boolean).length;
  const completePill = completeness === 4
    ? `<span class="pill pill-green">✓ Completo</span>`
    : `<span class="pill pill-warn">${completeness}/4</span>`;

  return `<tr class="${isSelected?'selected':''} ${c._isDup?'duplicate':''}"
    onclick="cpSelectClient('${c.id}')">
    <td onclick="event.stopPropagation()">
      <input type="checkbox" class="checkbox" ${inMerge?'checked':''}
        onchange="cpToggleMerge('${c.id}', this.checked)">
    </td>
    <td>
      <div class="cliente-nome">${c.nome}${c._isDup?'<span class="dup-badge">DUP</span>':''}</div>
      <div class="cliente-email">${c.email || '<span style="color:var(--text3)">—</span>'}</div>
    </td>
    <td><span class="pill pill-blue">${c.nSoggiorni}×</span></td>
    <td style="color:var(--text2);font-size:11px">${c.primaVisita || '—'}</td>
    <td>${completePill}</td>
    <td>
      ${bs.length > 0
        ? `<span class="pill pill-gray">${bs.length} prenotaz.</span>`
        : `<span style="color:var(--text3);font-size:11px">—</span>`}
    </td>
    <td>
      <button class="btn btn-ghost btn-sm" onclick="event.stopPropagation();cpEditClient('${c.id}')">✎</button>
    </td>
  </tr>`;
}

function cpDetail(c) {
  const bs = _cpGetBookingsAnyDup(c);
  const dupGroup = c._dupGroup ? c._dupGroup.filter(id => id !== c.id) : [];

  return `
    <div class="detail-header">
      <div style="display:flex;align-items:flex-start;justify-content:space-between">
        <div>
          <div class="detail-name">${c.nome}</div>
          <div class="detail-id">${c.id}</div>
        </div>
        <button class="btn btn-ghost btn-sm" onclick="_cpSelected=null;cpRender()">✕</button>
      </div>
      ${c._isDup ? `
        <div style="margin-top:8px;padding:8px;background:rgba(244,63,94,.1);border-radius:6px;border:1px solid rgba(244,63,94,.2)">
          <div style="font-size:11px;font-weight:600;color:var(--danger);margin-bottom:4px">⚠ Possibile duplicato</div>
          <div style="font-size:10px;color:var(--text3)">Trovati ${c._dupGroup.length} clienti con nome simile</div>
          <button class="btn btn-danger btn-sm" style="margin-top:6px;width:100%"
            onclick="cpPrepareMergeGroup('${c.id}')">⚡ Unifica con i duplicati</button>
        </div>` : ''}
    </div>

    <div class="detail-section">
      <div class="detail-section-title">📋 Dati anagrafici</div>
      ${cpDetailRow('Email', c.email)}
      ${cpDetailRow('Telefono', c.telefono)}
      ${cpDetailRow('Documento', c.docTipo ? `${c.docTipo} ${c.docNum}` : c.docNum)}
      ${cpDetailRow('Nazionalità', c.nazionalita)}
      ${cpDetailRow('Data nascita', c.dataNascita)}
      ${cpDetailRow('Note', c.note)}
    </div>

    <div class="detail-section">
      <div class="detail-section-title">📅 Storico soggiorni (${bs.length})</div>
      ${bs.length === 0
        ? `<div style="color:var(--text3);font-size:11px">Nessuna prenotazione collegata</div>`
        : bs.sort((a,b) => new Date(b.s)-new Date(a.s)).map(b => `
          <div class="booking-item" onclick="selBook(${b.id},null);closeClientiPanel()">
            <div style="display:flex;justify-content:space-between;align-items:start">
              <div>
                <div class="b-date">${fmt(b.s)} → ${fmt(b.e)}</div>
                <div class="b-room">Camera ${typeof roomName==='function'?roomName(b.r):b.r} · ${b.d || '—'}</div>
              </div>
              <div>
                <span class="pill pill-${b.c?'':'gray'}" style="background:${b.c};color:#1a1a2e;font-size:10px">
                  ${Math.round((new Date(b.e)-new Date(b.s))/86400000)}n
                </span>
              </div>
            </div>
          </div>
        `).join('')
      }
    </div>

    <div class="detail-section" style="border-bottom:none">
      <div style="display:flex;gap:6px;flex-wrap:wrap">
        <button class="btn btn-primary btn-sm" onclick="cpEditClient('${c.id}')">✎ Modifica</button>
        <button class="btn btn-ghost btn-sm" onclick="cpToggleMerge('${c.id}',true);document.querySelector('.merge-bar').scrollIntoView()">+ Unifica</button>
        <button class="btn btn-ghost btn-sm" style="color:var(--danger);border-color:rgba(244,63,94,.3)"
          onclick="cpDeleteClient('${c.id}')">🗑 Elimina</button>
      </div>
    </div>
  `;
}

function cpDetailRow(key, val) {
  if (!val) return `<div class="detail-row"><span class="detail-key">${key}</span><span style="color:var(--text3);font-size:11px">—</span></div>`;
  return `<div class="detail-row"><span class="detail-key">${key}</span><span class="detail-val">${val}</span></div>`;
}

// ── Actions ───────────────────────────────────────────────────────
function cpSearch(q) { _cpQuery = q; _cpApplyFilter(); cpRender(); }
function cpSetFilter(f) { _cpFilter = f; _cpApplyFilter(); cpRender(); }
function cpSort(key) {
  _cpSort = _cpSort.key === key ? { key, dir: -_cpSort.dir } : { key, dir: -1 };
  _cpApplyFilter(); cpRender();
}

function cpSelectClient(id) {
  _cpSelected = _cpClienti.find(c => c.id === id) || null;
  cpRender();
}

function cpToggleMerge(id, on) {
  if (on) _cpMergeSet.add(id); else _cpMergeSet.delete(id);
  cpRender();
}

function cpSelectAll(on) {
  if (on) _cpFiltered.forEach(c => _cpMergeSet.add(c.id));
  else _cpMergeSet.clear();
  cpRender();
}

function cpClearMerge() { _cpMergeSet.clear(); cpRender(); }

function cpPrepareMergeGroup(id) {
  const c = _cpClienti.find(x => x.id === id);
  if (!c?._dupGroup) return;
  c._dupGroup.forEach(gid => _cpMergeSet.add(gid));
  cpRender();
  setTimeout(() => document.getElementById('cp-merge-bar')?.scrollIntoView({ behavior: 'smooth' }), 100);
}

async function cpMerge() {
  if (_cpMergeSet.size < 2) return;
  const ids = [..._cpMergeSet];
  const clienti = ids.map(id => _cpClienti.find(c => c.id === id)).filter(Boolean);
  const names = clienti.map(c => c.nome).join(', ');

  if (!confirm(`Unificare questi ${ids.length} clienti?\n${names}\n\nTutte le prenotazioni verranno attribuite al primo cliente della lista.`)) return;

  // Il cliente "master" è quello con più soggiorni
  const master = clienti.reduce((best, c) => c.nSoggiorni > best.nSoggiorni ? c : best);
  const toMerge = clienti.filter(c => c.id !== master.id);

  try {
    showLoading?.('Unificazione clienti...');
    for (const dup of toMerge) {
      // Aggiorna tutte le prenotazioni del duplicato al master
      const dupBookings = _cpGetBookings(dup.id);
      for (const b of dupBookings) {
        b.clienteId = master.id;
        if (typeof dbUpdateRow === 'function') {
          await dbUpdateRow(b.dbRow, bookingToDbRow(b, b.fonte || 'app'));
        }
      }
      // Aggiorna soggiorni master
      master.nSoggiorni += dup.nSoggiorni;
      if (typeof aggiornaCliente === 'function') await aggiornaCliente(master);
      // Segna duplicato come eliminato nel foglio CLIENTI
      if (typeof dbGet === 'function' && DATABASE_SHEET_ID) {
        await apiFetch(
          `https://sheets.googleapis.com/v4/spreadsheets/${DATABASE_SHEET_ID}/values/${encodeURIComponent('CLIENTI!A'+dup.dbRow+':M'+dup.dbRow)}?valueInputOption=RAW`,
          { method:'PUT', body: JSON.stringify({ values:[[dup.id, '[UNIFICATO→'+master.id+']', ...Array(11).fill('')]] }), headers:{'Content-Type':'application/json'} }
        );
      }
    }
    _cpMergeSet.clear();
    hideLoading?.();
    showToast?.('✅ Clienti unificati', 'success');
    await cpInit();
  } catch(e) {
    hideLoading?.();
    showToast?.('❌ Errore: ' + e.message, 'error');
  }
}

function cpEditClient(id) {
  // Apre modal anagrafica Blip se disponibile
  if (typeof preimpostaClienteModal === 'function') preimpostaClienteModal(id);
}

async function cpDeleteClient(id) {
  const c = _cpClienti.find(x => x.id === id);
  if (!c) return;
  const bs = _cpGetBookings(id);
  if (bs.length > 0) {
    alert(`Impossibile eliminare: il cliente ha ${bs.length} prenotazioni collegate.\nScollegale prima dalla prenotazione.`);
    return;
  }
  if (!confirm(`Eliminare il cliente "${c.nome}"?`)) return;
  // Marca come eliminato nel DB
  try {
    await apiFetch(
      `https://sheets.googleapis.com/v4/spreadsheets/${DATABASE_SHEET_ID}/values/${encodeURIComponent('CLIENTI!A'+c.dbRow+':M'+c.dbRow)}?valueInputOption=RAW`,
      { method:'PUT', body: JSON.stringify({ values:[[c.id, '[ELIMINATO]', ...Array(11).fill('')]] }), headers:{'Content-Type':'application/json'} }
    );
    showToast?.('Cliente eliminato', 'info');
    _cpSelected = null;
    await cpInit();
  } catch(e) {
    showToast?.('❌ ' + e.message, 'error');
  }
}

function cpEsportaCSV() {
  const header = 'ID,Nome,Email,Telefono,Doc tipo,Doc num,Naz,Data nascita,Note,Prima visita,N soggiorni';
  const rows = _cpFiltered.map(c =>
    [c.id, c.nome, c.email, c.telefono, c.docTipo, c.docNum, c.nazionalita,
     c.dataNascita, c.note, c.primaVisita, c.nSoggiorni]
    .map(v => `"${(v||'').toString().replace(/"/g,'""')}"`).join(',')
  );
  const csv = [header, ...rows].join('\n');
  const a = document.createElement('a');
  a.href = 'data:text/csv;charset=utf-8,' + encodeURIComponent(csv);
  a.download = `clienti_${new Date().toISOString().slice(0,10)}.csv`;
  a.click();
}

// ── Esponi come pannello globale per Blip ─────────────────────────
function openClientiPanel() {
  let panel = document.getElementById('cp-panel');
  if (!panel) {
    panel = document.createElement('div');
    panel.id = 'cp-panel';
    panel.style.cssText = 'position:fixed;inset:0;z-index:300;background:var(--bg,#0f1117);overflow:auto';
    panel.innerHTML = '<div id="cp-app" style="min-height:100vh"></div>';
    document.body.appendChild(panel);
    // Close on Escape
    document.addEventListener('keydown', e => { if(e.key==='Escape') closeClientiPanel(); }, {once:true});
  }
  panel.style.display = 'block';
  cpInit();
}

function closeClientiPanel() {
  document.getElementById('cp-panel')?.remove();
}

// ── Avvio se standalone ───────────────────────────────────────────
if (document.getElementById('cp-app')) cpInit();
