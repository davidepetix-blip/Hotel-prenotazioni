// ═══════════════════════════════════════════════════════════════════
// checkin.js — Check-in operativo, Alloggiati Web, OCR Gemini
// Blip Hotel Management — build 18.7.x
// Dipende da: core.js, sync.js
// ═══════════════════════════════════════════════════════════════════


const BLIP_VER_CHECKIN = '11'; // ← incrementa ad ogni modifica

const CI_SHEET_NAME  = 'CHECK-IN';
const CI_CACHE_KEY   = 'hotelCiCache';
const CI_CACHE_TTL   = 60 * 60 * 1000; // 1 ora

// Colonne scheda CHECK-IN
// A=ID_CHECKIN B=ID_PRENOTAZIONE C=CAMERA D=DATA_CHECKIN E=NUM_OSPITI
// F=OSPITI_JSON G=TS_INSERIMENTO H=UTENTE

// Tipi documento Alloggiati Web
const TIPI_DOC = [
  {cod:'IDENT', label:'Carta d\'identità'},
  {cod:'PASOR', label:'Passaporto'},
  {cod:'PATEN', label:'Patente di guida'},
  {cod:'ALTRO', label:'Altro documento'},
];

// Mappa sessi
const SESSI = [{cod:'M',label:'Maschio'},{cod:'F',label:'Femmina'}];

// Stato check-in in memoria: { prenotazioneDbId: { guests:[], ciId, ciRow, ts } }
let ciData = {};
let _ciTab = 0;
let _ciEditBookingId = null; // dbId della prenotazione in editing
let _ciEditGuests = [];      // array ospiti nel form

// ── Apertura / chiusura pagina ──
function openCheckin() {
  document.getElementById('checkinPage').classList.add('open');
  loadCiData().then(() => renderCiTab());
}
function closeCheckin() {
  document.getElementById('checkinPage').classList.remove('open');
}
function switchCiTab(n) {
  _ciTab = n;
  document.querySelectorAll('.ci-tab').forEach((t,i) => t.classList.toggle('active', i===n));
  renderCiTab();
}
function renderCiTab() {
  if (_ciTab === 0) renderCiToday();
  else              renderCiHistory();
}

// ── Scheda CHECK-IN nel DATABASE ──
async function ensureCiSheet() {
  const id = DATABASE_SHEET_ID;
  if (!id) return;
  try {
    const d = await dbGet(`${CI_SHEET_NAME}!A1:H1`);
    if (d.values?.[0]?.[0] === 'ID_CHECKIN') return;
  } catch(e) {
    // Crea la scheda
    try {
      const r = await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${id}:batchUpdate`, {
        method:'POST',
        headers:{Authorization:'Bearer '+accessToken,'Content-Type':'application/json'},
        body: JSON.stringify({requests:[{addSheet:{properties:{title:CI_SHEET_NAME}}}]})
      });
      if (!r.ok) { const t=await r.text(); if (!t.includes('already')) throw new Error(t); }
    } catch(e2) { if (!String(e2.message).includes('already')) throw e2; }
  }
  // Scrivi intestazioni
  await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${id}/values/${encodeURIComponent(CI_SHEET_NAME+'!A1:H1')}?valueInputOption=RAW`, {
    method:'PUT',
    headers:{Authorization:'Bearer '+accessToken,'Content-Type':'application/json'},
    body: JSON.stringify({values:[['ID_CHECKIN','ID_PRENOTAZIONE','CAMERA','DATA_CHECKIN','NUM_OSPITI','OSPITI_JSON','TS_INSERIMENTO','UTENTE']]})
  });
}

async function readCiSheet() {
  try {
    const d = await dbGet(`${CI_SHEET_NAME}!A2:H9999`);
    const rows = d.values || [];
    const result = {};
    rows.forEach((row, i) => {
      const preId = (row[1] || '').trim();
      const ciId  = (row[0] || '').trim();
      // Usa preId come chiave principale; fallback su ciId per righe senza preId
      const key = preId || ciId;
      if (!key) return;
      try {
        const rec = {
          ciId,
          preId,
          camera:    row[2] || '',
          data:      row[3] || '',
          numOspiti: parseInt(row[4]) || 0,
          guests:    JSON.parse(row[5] || '[]'),
          ts:        row[6] || '',
          utente:    row[7] || '',
          ciRow:     i + 2,
        };
        result[key] = rec;
        // NON indicizziamo per ciId separatamente: causerebbe duplicati in Object.values()
      } catch(e) {}
    });
    return result;
  } catch(e) { return {}; }
}

async function writeCiRow(preId, data) {
  const id = DATABASE_SHEET_ID;
  if (!id) throw new Error('DATABASE non configurato');
  await ensureCiSheet();
  const row = [
    data.ciId, preId, data.camera, data.data,
    data.numOspiti, JSON.stringify(data.guests), data.ts, data.utente
  ];
  if (data.ciRow) {
    const range = `${CI_SHEET_NAME}!A${data.ciRow}:H${data.ciRow}`;
    await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${id}/values/${encodeURIComponent(range)}?valueInputOption=RAW`, {
      method:'PUT', headers:{Authorization:'Bearer '+accessToken,'Content-Type':'application/json'},
      body: JSON.stringify({values:[row]})
    });
  } else {
    const url = `https://sheets.googleapis.com/v4/spreadsheets/${id}/values/${encodeURIComponent(CI_SHEET_NAME)}:append?valueInputOption=RAW&insertDataOption=INSERT_ROWS`;
    const resp = await fetch(url, {
      method:'POST', headers:{Authorization:'Bearer '+accessToken,'Content-Type':'application/json'},
      body: JSON.stringify({values:[row]})
    });
    const r2 = await resp.json();
    const m = (r2.updates?.updatedRange||'').match(/(\d+):/);
    if (m) data.ciRow = parseInt(m[1]);
  }
}

async function loadCiData(force=false) {
  if (!force) {
    try {
      const raw = localStorage.getItem(CI_CACHE_KEY);
      if (raw) {
        const p = JSON.parse(raw);
        if (Date.now() - p.ts < CI_CACHE_TTL) { ciData = p.data; return; }
      }
    } catch(e) {}
  }
  if (!DATABASE_SHEET_ID) return;
  try {
    await ensureCiSheet();
    ciData = await readCiSheet();
    localStorage.setItem(CI_CACHE_KEY, JSON.stringify({ts:Date.now(), data:ciData}));
  } catch(e) { console.warn('loadCiData:', e.message); }
}

function invalidateCiCache() { localStorage.removeItem(CI_CACHE_KEY); }

// Invalida cache se la versione del modulo è cambiata
(function() {
  const cacheVerKey = 'hotelCiCacheVer';
  const currentVer = typeof BLIP_VER_CHECKIN !== 'undefined' ? BLIP_VER_CHECKIN : '0';
  const cachedVer = localStorage.getItem(cacheVerKey);
  if (cachedVer !== currentVer) {
    localStorage.removeItem(CI_CACHE_KEY);
    localStorage.setItem(cacheVerKey, currentVer);
  }
})();

// ─────────────────────────────────────────────────────────────────
// RICONCILIAZIONE — collega check-in orfani alle prenotazioni
// Legge le righe CHECK-IN senza ID_PRENOTAZIONE e tenta di trovarla
// per camera + data_checkin nel foglio PRENOTAZIONI
// ─────────────────────────────────────────────────────────────────
async function riconciliaCheckin() {
  if (!DATABASE_SHEET_ID) { showToast('DB non configurato', 'error'); return; }
  showLoading('Riconciliazione check-in…');
  try {
    const d = await dbGet(`${CI_SHEET_NAME}!A2:H9999`);
    const rows = d.values || [];
    const orfani = rows
      .map((row, i) => ({ row, rowNum: i + 2 }))
      .filter(({ row }) => !(row[1] || '').trim() && (row[0] || '').trim()); // preId vuoto, ciId presente

    if (orfani.length === 0) {
      hideLoading();
      showToast('Nessun check-in orfano trovato', 'success');
      return;
    }

    let fixed = 0, notFound = 0;
    const updates = [];

    for (const { row, rowNum } of orfani) {
      const camera   = (row[2] || '').trim();
      const dataCI   = (row[3] || '').trim(); // formato YYYY-MM-DD
      if (!camera || !dataCI) { notFound++; continue; }

      // Cerca prenotazione con stessa camera e data di arrivo
      const booking = bookings.find(b => {
        const camName = b.cameraName || roomName(b.r) || '';
        const arrivo  = b.s instanceof Date ? b.s.toISOString().slice(0,10) : '';
        return camName === camera && arrivo === dataCI;
      });

      if (!booking || !booking.dbId) { notFound++; continue; }

      // Aggiorna colonna B con il dbId trovato
      updates.push({
        range: `${CI_SHEET_NAME}!B${rowNum}`,
        values: [[booking.dbId]]
      });
      fixed++;
    }

    if (updates.length > 0) {
      const url = `https://sheets.googleapis.com/v4/spreadsheets/${DATABASE_SHEET_ID}/values:batchUpdate`;
      await apiFetch(url, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ valueInputOption: 'RAW', data: updates })
      });
      // Ricarica ciData
      await loadCiData(true);
    }

    hideLoading();
    const msg = `Riconciliati ${fixed} check-in` + (notFound > 0 ? `, ${notFound} non trovati` : '');
    showToast(msg, fixed > 0 ? 'success' : 'warning');
    syncLog(msg, fixed > 0 ? 'ok' : 'wrn');
    renderCiTab();
  } catch(e) {
    hideLoading();
    showToast('Errore riconciliazione: ' + e.message, 'error');
    console.error('[riconcilia]', e);
  }
}

// ── Render tab Oggi ──
function renderCiToday() {
  const body = document.getElementById('ciBody');
  const today = new Date(); today.setHours(0,0,0,0);
  const yesterday = new Date(today); yesterday.setDate(yesterday.getDate() - 1);

  // Arrivi di oggi e ieri
  // De-duplica per (camera + data arrivo) — evita doppioni da multi-month o import
  const _dedup = (arr) => {
    const seen = new Set();
    return arr.filter(b => {
      const key = (b.r || '') + '|' + (b.s?.toISOString()?.slice(0,10) || '');
      if (seen.has(key)) return false;
      seen.add(key);
      return true;
    });
  };
  const arriviOggi = _dedup(bookings.filter(b => {
    const s = new Date(b.s); s.setHours(0,0,0,0);
    return s.getTime() === today.getTime();
  }));
  const arriviIeri = _dedup(bookings.filter(b => {
    const s = new Date(b.s); s.setHours(0,0,0,0);
    const key = b.dbId || b.id;
    return s.getTime() === yesterday.getTime() && !ciData[key]; // ieri solo se non ancora registrati
  }));
  const tuttiArrivi = [...arriviOggi, ...arriviIeri];

  // Contatori riepilogo (solo oggi per i contatori principali)
  const totArrivi   = arriviOggi.length;
  const completati  = arriviOggi.filter(b => getCiForBooking(b)).length;
  const daFare      = totArrivi - completati;
  const totOspiti   = arriviOggi.reduce((sum,b) => {
    const ci = getCiForBooking(b);
    return sum + (ci ? ci.numOspiti : 0);
  }, 0);

  let html = `
    <div class="ci-summary-box">
      <div><div class="ci-summary-num">${totArrivi}</div><div class="ci-summary-label">Arrivi oggi</div></div>
      <div><div class="ci-summary-num" style="color:#e67e22">${daFare}</div><div class="ci-summary-label">Da fare</div></div>
      <div><div class="ci-summary-num">${totOspiti}</div><div class="ci-summary-label">Ospiti reg.</div></div>
    </div>`;

  if (tuttiArrivi.length === 0) {
    html += `<div style="text-align:center;padding:40px 20px;color:var(--text3);">Nessun arrivo previsto oggi</div>`;
  } else {
    // Raggruppa: prima ieri (da fare), poi oggi (da fare), poi oggi (completati)
    const ieriDaFare  = arriviIeri; // già filtrati solo non registrati
    const oggiDaFare  = arriviOggi.filter(b => !getCiForBooking(b));
    const oggiDone    = arriviOggi.filter(b =>  getCiForBooking(b));

    const renderCard = (b, isYesterday) => {
      const ci      = getCiForBooking(b);
      const done    = !!ci;
      const room    = ROOMS.find(r => r.id === b.r);
      const roomName= room?.name || b.cameraName || '—';
      const ospiti  = ci ? `${ci.numOspiti} ospite${ci.numOspiti!==1?'i':''}` : (b.d || '');
      const tag     = isYesterday ? '<span style="font-size:10px;background:#f0ad4e;color:#fff;padding:1px 6px;border-radius:4px;margin-left:6px;">ieri</span>' : '';
      return `
        <div class="ci-arrival-card ${done?'done':'todo'}" onclick="openCiModal('${b.dbId||''}')">
          <span class="ci-status-dot" style="background:${done?'#2d6a4f':'#e67e22'}"></span>
          <div class="ci-arrival-info">
            <div class="ci-arrival-name">${b.n || '—'}${tag}</div>
            <div class="ci-arrival-sub">${done ? '✓ Registrato · '+ospiti : 'Da registrare · '+(b.d||'')}
              ${done && ci.guests[0] ? ' · '+ci.guests[0].cognome+' '+ci.guests[0].nome : ''}
            </div>
          </div>
          <div class="ci-arrival-cam">${roomName}</div>
        </div>`;
    };

    if (ieriDaFare.length > 0) {
      html += `<div style="font-size:11px;font-weight:600;color:#e67e22;padding:8px 0 4px;letter-spacing:.04em;">⚠ IN RITARDO — ARRIVI DI IERI</div>`;
      ieriDaFare.forEach(b => { html += renderCard(b, true); });
      html += `<div style="font-size:11px;color:var(--text3);padding:8px 0 4px;letter-spacing:.04em;">ARRIVI DI OGGI</div>`;
    }
    oggiDaFare.forEach(b => { html += renderCard(b, false); });
    oggiDone.forEach(b  => { html += renderCard(b, false); });
  }

  // Export bar
  const completatiTot = tuttiArrivi.filter(b => ciData[b.dbId]).length;
  if (completatiTot > 0) {
    html += `
      <div class="ci-export-bar">
        <div class="ci-export-info">
          <strong>${completatiTot} check-in</strong> pronti per Alloggiati Web<br>
          Genera il file .txt da caricare sul portale della Polizia di Stato
        </div>
        <button class="btn primary" style="flex-shrink:0" onclick="exportAlloggiati('today')">⬇ Genera .txt</button>
      </div>`;
  }

  body.innerHTML = html;
}

// ── Render tab Storico ──
let _ciHistSearch = '';
let _exportDialogItems = []; // items temporanei per il dialog export
function renderCiHistory() {
  const body = document.getElementById('ciBody');
  // De-duplica per ciId prima di mostrare lo storico
  const _seenIds = new Set();
  const allCi = Object.values(ciData).filter(ci => {
    const k = ci.ciId || ci.preId;
    if (!k || _seenIds.has(k)) return false;
    _seenIds.add(k);
    return true;
  }).sort((a,b) => b.data.localeCompare(a.data));
  const filtered = _ciHistSearch
    ? allCi.filter(ci => {
        const q = _ciHistSearch.toLowerCase();
        return ci.camera.toLowerCase().includes(q) ||
               (ci.guests[0] && (ci.guests[0].cognome+' '+ci.guests[0].nome).toLowerCase().includes(q)) ||
               ci.data.includes(q);
      })
    : allCi;

  // Raggruppa per data
  const byDate = {};
  filtered.forEach(ci => {
    const d = ci.data || '—';
    if (!byDate[d]) byDate[d] = [];
    byDate[d].push(ci);
  });

  let html = `
    <div class="ci-search-bar">
      <input type="text" placeholder="Cerca per nome, camera, data…"
             value="${_ciHistSearch}"
             oninput="_ciHistSearch=this.value; renderCiHistory()"
             autocomplete="off">
      ${_ciHistSearch ? `<button class="btn" onclick="_ciHistSearch=''; renderCiHistory()">✕</button>` : ''}
    </div>`;

  if (filtered.length === 0) {
    html += `<div style="text-align:center;padding:40px 20px;color:var(--text3);">Nessun check-in trovato</div>`;
  } else {
    Object.keys(byDate).sort((a,b)=>b.localeCompare(a)).forEach(date => {
      const items = byDate[date];
      const d = date.split('-');
      const dataLabel = d.length===3 ? `${d[2]}/${d[1]}/${d[0]}` : date;
      html += `<div style="font-size:10px;font-weight:700;color:var(--text3);text-transform:uppercase;letter-spacing:.08em;margin:14px 0 8px;">${dataLabel} · ${items.length} arriv${items.length===1?'o':'i'}</div>`;
      items.forEach(ci => {
        const cap = ci.guests[0] || {};
        const nomeCapo = cap.cognome ? cap.cognome+' '+cap.nome : '—';
        html += `
          <div class="ci-history-row" onclick="openCiModalFromCi('${ci.preId}')">
            <div>
              <div class="ci-history-name">${nomeCapo}</div>
              <div class="ci-history-sub">Camera ${ci.camera} · ${ci.numOspiti} ospite${ci.numOspiti!==1?'i':''} · ${ci.preId ? 'Prenotazione: '+ci.preId.slice(-5) : ''}</div>
            </div>
            <div class="ci-history-cam">${ci.camera}</div>
          </div>`;
      });
    });
  }

  // Export tutto + riconcilia
  html += `
    <div class="ci-export-bar" style="margin-top:16px;">
      <div class="ci-export-info">Genera file Alloggiati Web${_ciHistSearch?' (filtrati)':''}</div>
      <div style="display:flex;gap:8px;flex-shrink:0">
        ${filtered.length > 0 ? `<button class="btn primary" onclick="exportAlloggiati('filtered')">⬇ .txt</button>` : ''}
        <button class="btn" onclick="riconciliaCheckin()" title="Collega check-in orfani alle prenotazioni">🔗 Riconcilia</button>
      </div>
    </div>`;

  body.innerHTML = html;
}

// ── Modale check-in ──
// Helper: cerca il check-in per una prenotazione provando più chiavi
function getCiForBooking(b) {
  if (!b) return null;
  // 1. Cerca per dbId (es. PRE-2026-XXXXXX)
  if (b.dbId && ciData[b.dbId]) return ciData[b.dbId];
  // 2. Cerca per id numerico come stringa
  if (ciData[String(b.id)]) return ciData[String(b.id)];
  // 3. Fallback: cerca per camera + data arrivo (recupera check-in orfani senza preId)
  const camName = b.cameraName || (typeof roomName === 'function' ? roomName(b.r) : '') || '';
  const arrivo  = b.s instanceof Date ? b.s.toISOString().slice(0,10) : '';
  if (camName && arrivo) {
    const found = Object.values(ciData).find(ci =>
      ci.camera === camName && ci.data === arrivo
    );
    if (found) {
      if (typeof syncLog === 'function') syncLog('CI trovato per camera+data: ' + camName + ' ' + arrivo, 'ok');
      return found;
    }
  }
  return null;
}

function openCiModal(bookingDbId) {
  let b = bookings.find(b => b.dbId === bookingDbId);
  if (!b && bookingDbId) {
    const numId = parseInt(bookingDbId);
    if (!isNaN(numId)) b = bookings.find(x => x.id === numId);
  }
  if (!b) { showToast('Prenotazione non trovata', 'error'); return; }
  openCiModalWithBooking(b);
}

function openCiModalWithBooking(b) {
  _ciEditBookingId = b.dbId || String(b.id);
  const room = ROOMS.find(r => r.id === b.r);
  document.getElementById('ciPanelTitle').textContent = `Check-in · Camera ${room?.name||b.cameraName||''}`;
  document.getElementById('ciPanelSub').textContent = `${b.n} · ${fmtDate(b.s)} → ${fmtDate(b.e)}`;

  // Carica ospiti: cerca con getCiForBooking per trovare anche check-in orfani
  const existing = getCiForBooking(b);
  _ciEditGuests = existing ? JSON.parse(JSON.stringify(existing.guests)) : [emptyGuest(true)];
  renderCiGuests();
  document.getElementById('ciOverlay').classList.add('open');
  document.getElementById('ciPanel').scrollTop = 0;
}

function openCiModalFromCi(preId) {
  // Cerca prima nel modo standard
  let b = bookings.find(b => b.dbId === preId);
  if (!b && preId) {
    const numId = parseInt(preId);
    if (!isNaN(numId)) b = bookings.find(x => x.id === numId);
  }
  // Fallback: cerca il ci in ciData e prova camera+data
  if (!b) {
    const ci = ciData[preId];
    if (ci) {
      b = bookings.find(bk => {
        const camName = bk.cameraName || (typeof roomName==='function'?roomName(bk.r):'') || '';
        const arrivo  = bk.s instanceof Date ? bk.s.toISOString().slice(0,10) : '';
        return camName === ci.camera && arrivo === ci.data;
      });
    }
  }
  if (!b) { showToast('Prenotazione non trovata', 'error'); return; }
  openCiModalWithBooking(b);
}

function closeCiModal() {
  document.getElementById('ciOverlay').classList.remove('open');
  _ciEditBookingId = null;
  _ciEditGuests = [];
}

function ciOverlayClick(e) {
  if (e.target === document.getElementById('ciOverlay')) closeCiModal();
}

function emptyGuest(isCapo=false) {
  return { isCapo, nome:'', cognome:'', dataNascita:'', sesso:'M', cittadinanza:'IT',
           luogoNascita:'', provNascita:'', statoEsteroNascita:'',
           tipoDoc: isCapo ? 'IDENT' : '', numDoc: isCapo ? '' : undefined,
           luogoRilascio: isCapo ? '' : undefined };
}

function fmtDate(d) {
  if (!d) return '—';
  const dt = d instanceof Date ? d : new Date(d);
  return dt.toLocaleDateString('it-IT',{day:'2-digit',month:'2-digit',year:'numeric'});
}

function renderCiGuests() {
  const container = document.getElementById('ciGuestsContainer');
  container.innerHTML = '';
  _ciEditGuests.forEach((g, idx) => {
    const isCapo   = idx === 0;
    const titleTxt = isCapo ? '👤 Capogruppo' : `👤 Accompagnatore ${idx}`;
    const block    = document.createElement('div');
    block.className= 'ci-ospite-block';
    block.innerHTML = `
      <div class="ci-ospite-header">
        <span class="ci-ospite-num">${titleTxt}</span>
        ${!isCapo ? `<button class="ci-remove-btn" onclick="ciRemoveGuest(${idx})">✕</button>` : ''}
      </div>
      ${isCapo ? `
      <button class="ci-scan-btn" id="ciScanBtn" onclick="ciTriggerScan(${idx})">
        📷 Scansiona documento d'identità
      </button>
      <img class="ci-scan-preview" id="ciScanPreview" alt="Anteprima documento"
           ${g._previewSrc ? 'src="'+g._previewSrc+'"' : ''}
           style="${g._previewSrc ? 'display:block' : 'display:none'}">
      <div id="ciScanBadge"></div>
      ` : ''}
      <!-- Nome / Cognome -->
      <div class="ci-field-row">
        <div class="ci-field">
          <label>Cognome *</label>
          <input type="text" value="${g.cognome||''}" placeholder="Rossi"
                 oninput="_ciEditGuests[${idx}].cognome=this.value">
        </div>
        <div class="ci-field">
          <label>Nome *</label>
          <input type="text" value="${g.nome||''}" placeholder="Mario"
                 oninput="_ciEditGuests[${idx}].nome=this.value">
        </div>
      </div>
      <!-- Sesso / Data nascita -->
      <div class="ci-field-row trio">
        <div class="ci-field">
          <label>Sesso *</label>
          <select onchange="_ciEditGuests[${idx}].sesso=this.value">
            ${SESSI.map(s=>`<option value="${s.cod}" ${g.sesso===s.cod?'selected':''}>${s.label}</option>`).join('')}
          </select>
        </div>
        <div class="ci-field" style="grid-column:span 2">
          <label>Data di nascita *</label>
          <input type="date" value="${g.dataNascita||''}"
                 oninput="_ciEditGuests[${idx}].dataNascita=this.value">
        </div>
      </div>
      <!-- Luogo nascita -->
      <div class="ci-field-row trio">
        <div class="ci-field" style="grid-column:span 2">
          <label>Comune di nascita *</label>
          <input type="text" value="${g.luogoNascita||''}" placeholder="Roma"
                 oninput="_ciEditGuests[${idx}].luogoNascita=this.value">
        </div>
        <div class="ci-field">
          <label>Prov.</label>
          <input type="text" value="${g.provNascita||''}" placeholder="RM" maxlength="2"
                 oninput="_ciEditGuests[${idx}].provNascita=this.value.toUpperCase();this.value=this.value.toUpperCase()">
        </div>
      </div>
      <!-- Stato estero (solo se non italiano) -->
      <div class="ci-field-row full" id="ciEstero${idx}" style="${(g.provNascita||'').trim()==='' && (g.statoEsteroNascita||'')!=='' || g.cittadinanza!=='IT' ? '' : 'display:none'}">
        <div class="ci-field">
          <label>Stato estero di nascita</label>
          <input type="text" value="${g.statoEsteroNascita||''}" placeholder="es. GERMANIA"
                 oninput="_ciEditGuests[${idx}].statoEsteroNascita=this.value.toUpperCase();this.value=this.value.toUpperCase()">
        </div>
      </div>
      <!-- Cittadinanza -->
      <div class="ci-field-row full">
        <div class="ci-field">
          <label>Cittadinanza (codice ISO) *</label>
          <input type="text" value="${g.cittadinanza||'IT'}" placeholder="IT" maxlength="3"
                 oninput="_ciEditGuests[${idx}].cittadinanza=this.value.toUpperCase();this.value=this.value.toUpperCase();
                          document.getElementById('ciEstero${idx}').style.display=this.value!=='IT'?'':'none'">
        </div>
      </div>
      ${isCapo ? `
      <!-- Documento (solo capogruppo) -->
      <div class="ci-section-title" style="margin-top:12px;">Documento d'identità</div>
      <div class="ci-field-row">
        <div class="ci-field">
          <label>Tipo documento *</label>
          <select onchange="_ciEditGuests[${idx}].tipoDoc=this.value">
            ${TIPI_DOC.map(t=>`<option value="${t.cod}" ${g.tipoDoc===t.cod?'selected':''}>${t.label}</option>`).join('')}
          </select>
        </div>
        <div class="ci-field">
          <label>Numero documento *</label>
          <input type="text" value="${g.numDoc||''}" placeholder="AB1234567"
                 oninput="_ciEditGuests[${idx}].numDoc=this.value.toUpperCase();this.value=this.value.toUpperCase()">
        </div>
      </div>
      <div class="ci-field-row full">
        <div class="ci-field">
          <label>Luogo di rilascio</label>
          <input type="text" value="${g.luogoRilascio||''}" placeholder="es. Roma"
                 oninput="_ciEditGuests[${idx}].luogoRilascio=this.value">
        </div>
      </div>` : ''}
    `;
    container.appendChild(block);
  });
}

function ciAddGuest() {
  _ciEditGuests.push(emptyGuest(false));
  renderCiGuests();
  // Scroll in fondo
  const panel = document.getElementById('ciPanel');
  setTimeout(() => panel.scrollTop = panel.scrollHeight, 50);
}

function ciRemoveGuest(idx) {
  _ciEditGuests.splice(idx, 1);
  renderCiGuests();
}

async function saveCiCheckin() {
  if (!_ciEditBookingId) return;

  // Validazione campi obbligatori capogruppo
  const capo = _ciEditGuests[0];
  const missing = [];
  if (!capo.cognome)     missing.push('Cognome');
  if (!capo.nome)        missing.push('Nome');
  if (!capo.dataNascita) missing.push('Data di nascita');
  if (!capo.luogoNascita)missing.push('Luogo di nascita');
  if (!capo.numDoc)      missing.push('Numero documento');
  if (missing.length > 0) {
    showToast('Campi obbligatori mancanti: ' + missing.join(', '), 'error');
    return;
  }

  const b = bookings.find(b => b.dbId === _ciEditBookingId);
  const room = ROOMS.find(r => r.id === b?.r);
  const existing = ciData[_ciEditBookingId] || ciData[b ? String(b.id) : ''] || null;

  const ciId    = existing?.ciId || 'CI-' + Date.now().toString(36).toUpperCase();
  const ciRow   = existing?.ciRow || null;
  const data    = new Date(b.s).toISOString().slice(0,10);
  const record  = {
    ciId, preId: _ciEditBookingId,
    camera: room?.name || b?.cameraName || '—',
    data,
    numOspiti: _ciEditGuests.length,
    guests: _ciEditGuests,
    ts: nowISO(),
    utente: document.getElementById('userAvatar')?.title || '',
    ciRow,
  };

  closeCiModal();
  showLoading('Salvataggio check-in…');
  try {
    await writeCiRow(_ciEditBookingId, record);
    ciData[_ciEditBookingId] = record;
    invalidateCiCache();
    localStorage.setItem(CI_CACHE_KEY, JSON.stringify({ts:Date.now(), data:ciData}));
    hideLoading();
    showToast('✓ Check-in salvato', 'success');
    renderCiTab();
    // Aggiorna anche il tab CI nel drawer se è aperto
    const drTabCI = document.getElementById('drTabCI');
    if (drTabCI && drTabCI.style.display !== 'none') {
      const bid = drTabCI.dataset.bookingId;
      if (bid && typeof renderCheckinDrawerTab === 'function') {
        renderCheckinDrawerTab(parseInt(bid));
      }
    }
  } catch(e) {
    hideLoading();
    showToast('Errore: ' + e.message, 'error');
  }
}

// ── Generazione file Alloggiati Web ──
// Formato fisso: record da 93 caratteri (versione 3)
// Ref: specifiche tecniche Polizia di Stato – tracciato record tipo 20
// --- TABELLE ALLOGGIATI WEB build 18.7.1 ---

function apriDialogExportAlloggiati(items, scope) {
  // De-duplica per ciId (evita doppi in Object.values se ciData ha chiavi multiple)
  const seenCiIds = new Set();
  items = items.filter(ci => {
    const k = ci.ciId || ci.preId;
    if (!k || seenCiIds.has(k)) return false;
    seenCiIds.add(k);
    return true;
  });
  if (items.length === 0) { showToast('Nessun check-in da esportare', 'error'); return; }
  // Salva in variabile globale — evita JSON.stringify gigantesco nell'onclick
  _exportDialogItems = items;

  const rows = items.map((ci, i) => {
    const cap = ci.guests && ci.guests[0] ? ci.guests[0] : {};
    const nome = cap.cognome ? cap.cognome + ' ' + cap.nome : '—';
    const doc  = cap.numDoc || '—';
    return '<div style="display:flex;align-items:center;gap:10px;padding:8px 0;border-bottom:1px solid var(--border)">' +
      '<input type="checkbox" id="exp_' + i + '" checked style="width:18px;height:18px;cursor:pointer;flex-shrink:0">' +
      '<label for="exp_' + i + '" style="flex:1;cursor:pointer">' +
        '<div style="font-weight:600;font-size:13px">' + escHtml(nome) + '</div>' +
        '<div style="font-size:11px;color:var(--text3)">Camera ' + escHtml(ci.camera) + ' · ' + ci.data + ' · ' + ci.numOspiti + ' ospite/i · Doc: ' + escHtml(doc) + '</div>' +
      '</label>' +
      '<button data-cikey="' + escHtml(ci.ciId||ci.preId||'') + '" class="_expmod" style="background:none;border:1px solid var(--border);border-radius:6px;padding:4px 8px;font-size:11px;cursor:pointer;color:var(--text2)">✎</button>' +
    '</div>';
  }).join('');

  const ov = document.createElement('div');
  ov.id = '_exportOv';
  ov.style.cssText = 'position:fixed;inset:0;background:rgba(0,0,0,.5);z-index:9999;display:flex;align-items:flex-end;justify-content:center;padding:0';

  const panel = document.createElement('div');
  panel.style.cssText = 'background:var(--surface);border-radius:16px 16px 0 0;padding:20px;width:100%;max-width:500px;max-height:80vh;display:flex;flex-direction:column';
  panel.innerHTML =
    '<div style="font-weight:700;font-size:15px;margin-bottom:4px">⬇ Esporta Alloggiati Web</div>' +
    '<div style="font-size:12px;color:var(--text3);margin-bottom:12px">' + items.length + ' check-in — deseleziona quelli da escludere</div>' +
    '<div id="_exportRows" style="overflow-y:auto;flex:1;margin-bottom:12px">' + rows + '</div>' +
    '<div style="display:flex;gap:8px">' +
      '<button id="_exportCancel" style="flex:1;background:none;border:1px solid var(--border);border-radius:8px;padding:10px;cursor:pointer;font-size:13px;color:var(--text2)">Annulla</button>' +
      '<button id="_exportConfirm" style="flex:1;background:var(--accent);color:#fff;border:none;border-radius:8px;padding:10px;font-weight:600;cursor:pointer;font-size:13px">⬇ Genera file</button>' +
    '</div>';

  ov.appendChild(panel);
  document.body.appendChild(ov);

  ov.addEventListener('click', e => { if(e.target===ov) ov.remove(); });
  document.getElementById('_exportCancel').addEventListener('click', () => ov.remove());
  panel.addEventListener('click', e => {
    const btn = e.target.closest('._expmod');
    if (btn) { apriModificaCiDaExport(btn.dataset.cikey); }
  });
  document.getElementById('_exportConfirm').addEventListener('click', () => {
    const selected = Array.from(document.querySelectorAll('[id^=exp_]:checked'))
      .map(cb => parseInt(cb.id.replace('exp_', '')));
    const toExport = _exportDialogItems.filter((_, i) => selected.includes(i));
    ov.remove();
    _esportaAlloggiatiDiretta(toExport, scope);
  });
}

function apriModificaCiDaExport(key) {
  document.querySelectorAll('div[style*="position:fixed"][style*="9999"]').forEach(el => el.remove());
  const ci = ciData[key];
  if (!ci) { showToast('Check-in non trovato', 'error'); return; }
  // Cerca la prenotazione con tutti i metodi disponibili
  let b = bookings.find(bk => bk.dbId === ci.preId);
  if (!b && ci.preId) {
    const numId = parseInt(ci.preId);
    if (!isNaN(numId)) b = bookings.find(x => x.id === numId);
  }
  if (!b) {
    // Fallback camera+data
    b = bookings.find(bk => {
      const camName = bk.cameraName || (typeof roomName==='function'?roomName(bk.r):'') || '';
      const arrivo  = bk.s instanceof Date ? bk.s.toISOString().slice(0,10) : '';
      return camName === ci.camera && arrivo === ci.data;
    });
  }
  if (!b) { showToast('Prenotazione non trovata', 'error'); return; }
  openCiModalWithBooking(b);
}

// Esporta direttamente una lista di ci già selezionata (chiamata dal dialog)
function _esportaAlloggiatiDiretta(items, scope) {
  if (!items || items.length === 0) { showToast('Nessun check-in selezionato', 'error'); return; }
  _exportAlloggiatiItems(items);
}

function exportAlloggiati(scope='today'){
  const today=new Date().toISOString().slice(0,10);
  let items;
  if(scope==='today'){items=Object.values(ciData).filter(ci=>ci.data===today);}
  else if(scope==='48h'){
    const cutoff=new Date(); cutoff.setHours(cutoff.getHours()-48);
    const cutoffStr=cutoff.toISOString().slice(0,10);
    items=Object.values(ciData).filter(ci=>ci.data>=cutoffStr);
  }
  else{items=Object.values(ciData).filter(ci=>{if(!_ciHistSearch)return true;const q=_ciHistSearch.toLowerCase();const cap=ci.guests[0]||{};return ci.camera.toLowerCase().includes(q)||((cap.cognome||'')+' '+(cap.nome||'')).toLowerCase().includes(q)||ci.data.includes(q);});}
  // Mostra dialog di selezione/preview prima di generare
  if(items.length>0){ apriDialogExportAlloggiati(items, scope); return; }
  if(items.length===0){showToast('Nessun check-in da esportare','error');return;}
  // Se arriviamo qui, il dialog è stato saltato — esegui direttamente
  _exportAlloggiatiItems(items);
}

// Nucleo dell'export — prende items già filtrati/selezionati
// Helper conversione ISO → nome stato (definiti FUORI dal loop)
const _ISO_TO_NOME = {'IT':'ITALIA','DE':'GERMANIA','FR':'FRANCIA','ES':'SPAGNA','GB':'REGNO UNITO','AT':'AUSTRIA','CH':'SVIZZERA','BE':'BELGIO','NL':'PAESI BASSI','PL':'POLONIA','RO':'ROMANIA','PT':'PORTOGALLO','GR':'GRECIA','CZ':'REPUBBLICA CECA','HU':'UNGHERIA','SE':'SVEZIA','DK':'DANIMARCA','FI':'FINLANDIA','SK':'REPUBBLICA SLOVACCA','SI':'SLOVENIA','HR':'CROAZIA','BG':'BULGARIA','LT':'LITUANIA','LV':'LETTONIA','EE':'ESTONIA','LU':'LUSSEMBURGO','MT':'MALTA','IE':'IRLANDA','CY':'CIPRO','US':'STATI UNITI D AMERICA','RU':'FEDERAZIONE RUSSA','CN':'CINA','JP':'GIAPPONE','IN':'INDIA','BR':'BRASILE','LY':'LIBIA','TN':'TUNISIA','MA':'MAROCCO','EG':'EGITTO','NG':'NIGERIA','GH':'GHANA','SN':'SENEGAL','CM':'CAMERUN','ET':'ETIOPIA','UA':'UCRAINA','TR':'TURCHIA','AL':'ALBANIA','MK':'MACEDONIA DEL NORD','RS':'SERBIA','BA':'BOSNIA ED ERZEGOVINA','ME':'MONTENEGRO','KO':'KOSOVO','XK':'KOSOVO'};
function _normCitNome(s) { const u=(s||'').toUpperCase().trim(); return _ISO_TO_NOME[u]||u; }
function _normCitIsIta(s) { return _alNorm(_normCitNome(s||'ITALIA')).includes('ITAL'); }

function _exportAlloggiatiItems(items){
  const today=new Date().toISOString().slice(0,10);
  const comuniMancanti=[],visti=new Set();
  items.forEach(ci=>{ci.guests.forEach((g,gIdx)=>{
    const _nc=((s)=>{const _i2n={'IT':'ITALIA','DE':'GERMANIA','FR':'FRANCIA','ES':'SPAGNA','GB':'REGNO UNITO','LV':'LETTONIA','EE':'ESTONIA','LT':'LITUANIA','AL':'ALBANIA','RS':'SERBIA','BA':'BOSNIA ED ERZEGOVINA','UA':'UCRAINA','TR':'TURCHIA','MA':'MAROCCO','TN':'TUNISIA','CN':'CINA','US':'STATI UNITI D AMERICA','RU':'FEDERAZIONE RUSSA'};const u=(s||'').toUpperCase().trim();return _i2n[u]||u;})(g.cittadinanza||'ITALIA');
    const isIta=_alNorm(_nc).includes('ITAL');
    if(isIta&&g.luogoNascita&&!_alCodiceComune(g.luogoNascita)){const k=_alNorm(g.luogoNascita);if(!visti.has('N:'+k)){visti.add('N:'+k);comuniMancanti.push({nomeComune:g.luogoNascita,label:'nascita - '+cleanAl(g.cognome)+' '+cleanAl(g.nome)});}}
    if(gIdx===0&&g.luogoRilascio&&!_alRisolviLuogo(g.luogoRilascio)){const k=_alNorm(g.luogoRilascio);if(!visti.has('R:'+k)){visti.add('R:'+k);comuniMancanti.push({nomeComune:g.luogoRilascio,label:'rilascio - '+cleanAl(g.cognome)+' '+cleanAl(g.nome)});}}
  });});
  if(comuniMancanti.length>0){_alChiediCodiciMancanti(comuniMancanti,()=>_exportAlloggiatiItems(items));return;}
  const toFmt=s=>{if(!s)return'          ';const c=s.replace(/-/g,'').replace(/\//g,'');if(c.length!==8)return'          ';let gg,mm,aaaa;if(/^(19|20)\d{6}$/.test(c)){aaaa=c.slice(0,4);mm=c.slice(4,6);gg=c.slice(6,8);}else{gg=c.slice(0,2);mm=c.slice(2,4);aaaa=c.slice(4,8);}return`${gg}/${mm}/${aaaa}`;};
  const nGiorni=(a,p)=>{try{const g=Math.round((new Date(p)-new Date(a))/86400000);return String(Math.min(Math.max(g,1),30)).padStart(2,'0');}catch{return'01';}};
  let lines=[];
  items.forEach(ci=>{
    const b=bookings.find(bk=>bk.dbId===ci.preId);
    const arrISO=ci.data,parISO=b?new Date(b.e).toISOString().slice(0,10):ci.data;
    ci.guests.forEach((g,gIdx)=>{
      const isCapo=gIdx===0,tipoAllog=isCapo?'16':'19';
      const cognome=padR(cleanAl(g.cognome),50),nome=padR(cleanAl(g.nome),30);
      const sesso=(_alNorm(g.sesso||'M').charAt(0)==='F')?'2':'1';
      const dataN=toFmt(g.dataNascita||'');
      const isIta=_normCitIsIta(g.cittadinanza||'ITALIA');
      let comN='         ',provN='  ';
      if(isIta&&g.luogoNascita){const found=_alCodiceComune(g.luogoNascita);if(found){comN=found.cod;provN=padR(found.prov,2);}}
      const statoN=_alCodiceStato(isIta?'ITALIA':_normCitNome(g.statoNascita||g.cittadinanza||'ITALIA')).padStart(9,'0');
      const cittad=_alCodiceStato(_normCitNome(g.cittadinanza||'ITALIA')).padStart(9,'0');
      const tipoDoc=isCapo?padR(_alCodiceDoc(g.tipoDoc),5):'     ';
      const numDoc=isCapo?padR(cleanAl(g.numDoc||''),20):'                    ';
      let luogoRil='         ';
      if(isCapo&&g.luogoRilascio){const rl=_alRisolviLuogo(g.luogoRilascio);if(rl)luogoRil=rl.cod;}
      const record=tipoAllog+toFmt(arrISO)+nGiorni(arrISO,parISO)+cognome+nome+sesso+dataN+comN+provN+statoN+cittad+tipoDoc+numDoc+luogoRil;
      if(record.length!==168)console.warn('[Alloggiati] len='+record.length,g.cognome);
      lines.push(record);
    });
  });
  const content=lines.join('\r\n');
  const blob=new Blob([content],{type:'text/plain;charset=utf-8'});
  const url=URL.createObjectURL(blob);const a=document.createElement('a');
  a.href=url;a.download=`alloggiati_${today}.txt`;a.click();URL.revokeObjectURL(url);
  showToast('File generato - '+lines.length+' record','success');
}

// Helper testo Alloggiati Web
function cleanAl(s) { return (s||'').toUpperCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/[^A-Z0-9 '.,\-]/g,''); }
function padR(s, n) { const t=(s||'').substring(0,n); return t+' '.repeat(Math.max(0,n-t.length)); }
function pad(s, n)  { const t=(s||'').substring(0,n); return '0'.repeat(Math.max(0,n-t.length))+t; }



// ═══════════════════════════════════════════════════════════════════
// OCR DOCUMENTO — Estrazione dati con Claude Vision
// ═══════════════════════════════════════════════════════════════════

let _ciOcrGuestIdx = 0; // indice ospite per cui stiamo scansionando


function ciTriggerScan(guestIdx) {
  const apiKey = localStorage.getItem('hotelGeminiApiKey') || '';
  if (!apiKey) {
    showToast('Inserisci prima la API Key Gemini nelle \u2699 Impostazioni', 'error');
    return;
  }
  _ciOcrGuestIdx = guestIdx;
  const inp = document.getElementById('ciDocInput');
  inp.value = '';
  inp.click();
}

async function testGeminiKey() {
  const apiKey = localStorage.getItem('hotelGeminiApiKey') || '';
  const result = document.getElementById('proxyTestResult');
  const detail = document.getElementById('proxyTestDetail');
  if (!apiKey) { result.textContent = '\u26a0 Chiave non inserita'; result.style.color='#e67e22'; return; }
  if (!apiKey.startsWith('AIza')) { result.textContent = '\u26a0 Formato non valido (deve iniziare con AIza)'; result.style.color='#e67e22'; return; }
  result.textContent = 'Ricerca modello disponibile\u2026'; result.style.color='var(--text3)';
  if (detail) detail.textContent = '';

  // Prova modelli in ordine finche uno funziona
  // gemini-1.5-flash prima — funziona con chiavi free tier standard
  const MODELS = [
    'gemini-1.5-flash',
    'gemini-1.5-flash-8b',
    'gemini-1.5-pro',
    'gemini-2.0-flash-lite',
    'gemini-2.0-flash',
    'gemini-2.5-flash-lite',
    'gemini-2.5-flash',
  ];
  let workingModel = null;
  let lastError = '';
  for (const model of MODELS) {
    try {
      result.textContent = 'Provo ' + model + '\u2026';
      const u = 'https://generativelanguage.googleapis.com/v1beta/models/' + model + ':generateContent?key=' + apiKey;
      const r = await fetch(u, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ contents: [{ parts: [{ text: 'ok' }] }] })
      });
      const d = await r.json();
      if (r.ok && d.candidates) { workingModel = model; break; }
      lastError = (d.error && d.error.message) || ('HTTP ' + r.status);
    } catch(ex) { lastError = ex.message; }
  }

  if (workingModel) {
    localStorage.setItem('hotelGeminiModel', workingModel);
    result.textContent = '\u2713 Chiave valida \u2014 ' + workingModel;
    result.style.color = 'var(--accent)';
    if (detail) detail.textContent = 'Modello salvato automaticamente.';
  } else {
    result.textContent = '\u2715 Nessun modello disponibile con questa chiave';
    result.style.color = 'var(--danger)';
    if (detail) detail.textContent = 'Ultimo errore: ' + lastError;
  }
}

async function ciHandleDocImage(input) {
  if (!input.files || !input.files[0]) return;
  const file = input.files[0];

  const preview = document.getElementById('ciScanPreview');
  const objUrl  = URL.createObjectURL(file);
  if (preview) { preview.src = objUrl; preview.style.display = 'block'; }
  if (_ciEditGuests[_ciOcrGuestIdx]) _ciEditGuests[_ciOcrGuestIdx]._previewSrc = objUrl;

  const base64 = await new Promise((res, rej) => {
    const r = new FileReader();
    r.onload  = () => res(r.result.split(',')[1]);
    r.onerror = () => rej(new Error('Lettura file fallita'));
    r.readAsDataURL(file);
  });

  const mediaType = file.type || 'image/jpeg';
  const overlay   = document.getElementById('ciOcrOverlay');
  const label     = document.getElementById('ciOcrLabel');
  overlay.classList.add('open');
  label.textContent = 'Analisi documento in corso\u2026';

  const btn = document.getElementById('ciScanBtn');
  if (btn) { btn.classList.add('loading'); btn.textContent = '\u23f3 Analisi in corso\u2026'; }

  try {
    const apiKey = localStorage.getItem('hotelGeminiApiKey') || '';
    if (!apiKey) throw new Error('API Key Gemini non configurata. Vai in \u2699 Impostazioni.');

    const prompt = 'Sei un assistente per la reception di un albergo italiano. '
      + 'Analizza questa immagine di un documento di identita e restituisci SOLO un JSON puro senza markdown:\n'
      + '{"cognome":"","nome":"","sesso":"M o F","dataNascita":"YYYY-MM-DD",'
      + '"luogoNascita":"","provNascita":"sigla 2 lettere se italiano vuoto se estero",'
      + '"statoEsteroNascita":"nome MAIUSCOLO solo se estero","cittadinanza":"IT o codice ISO 2 lettere",'
      + '"tipoDoc":"IDENT o PASOR o PATEN o ALTRO","numDoc":"per carta identita italiana usa il codice alfanumerico in alto a destra (es. CA22839IQ o AX1234567B) NON il numero a 6 cifre in basso. Per passaporto usa il codice 9 caratteri. Senza spazi.","luogoRilascio":""}';

    const model = localStorage.getItem('hotelGeminiModel') || 'gemini-1.5-flash';
    const url = 'https://generativelanguage.googleapis.com/v1beta/models/' + model + ':generateContent?key=' + apiKey;
    const response = await fetch(url, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        contents: [{ parts: [
          { text: prompt },
          { inline_data: { mime_type: mediaType, data: base64 } }
        ]}],
        generationConfig: { temperature: 0.1, maxOutputTokens: 500 }
      })
    });

    const data = await response.json();
    if (!response.ok) throw new Error('Gemini ' + response.status + ': ' + (data.error && data.error.message || 'errore sconosciuto'));

    const rawText = (((data.candidates || [])[0] || {}).content || {}).parts
                    ? data.candidates[0].content.parts[0].text : '';
    if (!rawText) throw new Error('Gemini ha restituito risposta vuota');

    let extracted;
    try {
      const clean = rawText.replace(/```json\s*/gi,'').replace(/```\s*/g,'').trim();
      extracted = JSON.parse(clean);
    } catch(pe) {
      const m = rawText.match(/\{[\s\S]*\}/);
      if (m) { try { extracted = JSON.parse(m[0]); } catch(e2) { throw new Error('JSON non valido: ' + rawText.slice(0,80)); } }
      else throw new Error('Nessun JSON: ' + rawText.slice(0,80));
    }
    console.log('OCR extracted:', JSON.stringify(extracted));

    const idx = _ciOcrGuestIdx;
    const g   = _ciEditGuests[idx];
    if (!g) throw new Error('Ospite non trovato');

    const previewSrc = document.getElementById('ciScanPreview') ? document.getElementById('ciScanPreview').src : '';
    const fields = ['cognome','nome','sesso','dataNascita','luogoNascita',
                    'provNascita','statoEsteroNascita','cittadinanza',
                    'tipoDoc','numDoc','luogoRilascio'];
    let filled = 0;
    fields.forEach(function(f) {
      const val = (extracted[f] || '').toString().trim();
      if (val !== '') { g[f] = val; filled++; }
    });

    overlay.classList.remove('open');
    renderCiGuests();

    const p2 = document.getElementById('ciScanPreview');
    if (p2 && previewSrc && previewSrc !== 'about:blank') { p2.src = previewSrc; p2.style.display = 'block'; }

    const badge = document.getElementById('ciScanBadge');
    if (badge) {
      badge.innerHTML = filled > 0
        ? '<div class="ci-scan-result-badge">\u2713 ' + filled + ' campi compilati \u2014 verifica e correggi se necessario</div>'
        : '<div class="ci-scan-result-badge" style="background:#fdecea;border-color:#f5c6cb;color:var(--danger);">\u26a0 Nessun campo estratto \u2014 riprova con foto pi\u00f9 nitida</div>';
    }

    const btn2 = document.getElementById('ciScanBtn');
    if (btn2) { btn2.classList.remove('loading'); btn2.textContent = '\ud83d\udcf7 Scansiona di nuovo'; }

    showToast(filled > 0
      ? '\u2713 Documento analizzato \u00b7 ' + filled + ' campi estratti'
      : 'Nessun dato estratto \u2014 riprova con foto pi\u00f9 nitida',
      filled > 0 ? 'success' : 'error');

  } catch(e) {
    overlay.classList.remove('open');
    console.error('OCR error:', e);
    showToast('Errore OCR: ' + e.message, 'error');
    const btn2 = document.getElementById('ciScanBtn');
    if (btn2) { btn2.classList.remove('loading'); btn2.textContent = '\ud83d\udcf7 Scansiona documento d\u2019identit\u00e0'; }
  }
}
// ═══════════════════════════════════════════════════════════════════
// GESTIONE FOGLI ANNUALI NELLE IMPOSTAZIONI
// ═══════════════════════════════════════════════════════════════════

