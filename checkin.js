// ═══════════════════════════════════════════════════════════════════
// checkin.js — Check-in operativo, Alloggiati Web, OCR Gemini
// Blip Hotel Management — build 18.7.x
// Dipende da: core.js, sync.js
// ═══════════════════════════════════════════════════════════════════


const BLIP_VER_CHECKIN = '21'; // ← incrementa ad ogni modifica

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
      const preId  = row[1] || '';
      const camera = row[2] || '';
      const data   = row[3] || '';
      try {
        const rec = {
          ciId:   row[0] || '',
          preId,
          camera,
          data,
          numOspiti: parseInt(row[4]) || 0,
          guests: JSON.parse(row[5] || '[]'),
          ts:     row[6] || '',
          utente: row[7] || '',
          ciRow:  i + 2,
        };
        // Chiave primaria per preId (se presente)
        if (preId) result[preId] = rec;
        // Chiave alternativa sempre: cam:CAMERA:DATA — permette lookup anche con preId vuoto
        if (camera && data) {
          const altKey = 'cam:' + camera + ':' + data;
          if (!result[altKey]) result[altKey] = rec; // non sovrascrivere se già presente
        }
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

// ── Render tab Oggi ──
function renderCiToday() {
  const body = document.getElementById('ciBody');
  const today = new Date(); today.setHours(0,0,0,0);
  const yesterday = new Date(today); yesterday.setDate(yesterday.getDate() - 1);

  // Arrivi di oggi e ieri
  const arriviOggi = bookings.filter(b => {
    const s = new Date(b.s); s.setHours(0,0,0,0);
    return s.getTime() === today.getTime();
  });
  const arriviIeri = bookings.filter(b => {
    const s = new Date(b.s); s.setHours(0,0,0,0);
    return s.getTime() === yesterday.getTime() && !ciData[b.dbId]; // ieri solo se non ancora registrati
  });
  const tuttiArrivi = [...arriviOggi, ...arriviIeri];

  // Contatori riepilogo (solo oggi per i contatori principali)
  const totArrivi   = arriviOggi.length;
  const completati  = arriviOggi.filter(b => ciData[b.dbId]).length;
  const daFare      = totArrivi - completati;
  const totOspiti   = arriviOggi.reduce((sum,b) => {
    const ci = ciData[b.dbId];
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
    const oggiDaFare  = arriviOggi.filter(b => !ciData[b.dbId]);
    const oggiDone    = arriviOggi.filter(b =>  ciData[b.dbId]);

    const renderCard = (b, isYesterday) => {
      const ci      = ciData[b.dbId];
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
        <button class="btn primary" style="flex-shrink:0" onclick="ciPreviewAlloggiati('72h')">⬇ Alloggiati</button>
      </div>`;
  }

  body.innerHTML = html;
}

// ── Render tab Storico ──
let _ciHistSearch = '';
function renderCiHistory() {
  const body = document.getElementById('ciBody');
  const allCi = Object.values(ciData).sort((a,b) => b.data.localeCompare(a.data));
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
          <div class="ci-history-row" onclick="openCiModalFromCi('${ci.ciId}')">
            <div>
              <div class="ci-history-name">${nomeCapo}</div>
              <div class="ci-history-sub">Camera ${ci.camera} · ${ci.numOspiti} ospite${ci.numOspiti!==1?'i':''} · ${ci.data ? ci.data.split('-').reverse().join('/') : ''}</div>
            </div>
            <div class="ci-history-cam">${ci.camera}</div>
          </div>`;
      });
    });
  }

  // Export tutto
  if (filtered.length > 0) {
    html += `
      <div class="ci-export-bar" style="margin-top:16px;">
        <div class="ci-export-info">Genera file Alloggiati Web per tutti i check-in${_ciHistSearch?' filtrati':''}</div>
        <button class="btn primary" style="flex-shrink:0" onclick="exportAlloggiati('filtered')">⬇ Genera .txt</button>
      </div>`;
  }

  body.innerHTML = html;
}

// ── Modale check-in ──
function openCiModal(bookingDbId) {
  // Trova il booking: prima per dbId esatto, poi per id numerico (fallback dal drawer)
  let b = bookings.find(x => x.dbId && x.dbId === bookingDbId);
  if (!b) b = bookings.find(x => String(x.id) === String(bookingDbId));
  if (!b) { showToast('Prenotazione non trovata', 'error'); return; }
  // Usa dbId se disponibile, altrimenti chiave locale stabile
  _ciEditBookingId = b.dbId || ('LOCAL-' + b.id);
  const room = ROOMS.find(r => r.id === b.r);
  document.getElementById('ciPanelTitle').textContent = `Check-in · Camera ${room?.name||b.cameraName}`;
  document.getElementById('ciPanelSub').textContent = `${b.n} · ${fmtDate(b.s)} → ${fmtDate(b.e)}`;

  // Carica ospiti esistenti o crea capogruppo vuoto
  const existing = ciData[bookingDbId];
  _ciEditGuests = existing ? JSON.parse(JSON.stringify(existing.guests)) : [emptyGuest(true)];
  renderCiGuests();
  document.getElementById('ciOverlay').classList.add('open');
  document.getElementById('ciPanel').scrollTop = 0;
}

function openCiModalFromCi(ciId) {
  // Trova il record CI per ciId
  const ci = Object.values(ciData).find(c => c.ciId === ciId);
  if (!ci) { showToast('Scheda check-in non trovata', 'error'); return; }

  // Prova a trovare il booking abbinato (per titolo panel)
  let b = bookings.find(x => x.dbId && x.dbId === ci.preId);
  if (!b) {
    // Fallback: cerca per camera + data arrivo
    b = bookings.find(x => {
      const room = ROOMS.find(r => r.id === x.r);
      const camName = room?.name || x.cameraName || '';
      const d = x.s instanceof Date ? x.s : new Date(x.s);
      const dataB = d.getFullYear() + '-' + String(d.getMonth()+1).padStart(2,'0') + '-' + String(d.getDate()).padStart(2,'0');
      return String(camName) === String(ci.camera) && dataB === ci.data;
    });
  }

  // Imposta l'ID editing — usa preId se disponibile, altrimenti ciId come chiave
  _ciEditBookingId = ci.preId || ('CI-DIRECT-' + ciId);
  _ciEditGuests = JSON.parse(JSON.stringify(ci.guests));

  // Titolo panel
  const room = b ? (ROOMS.find(r => r.id === b.r)) : null;
  const camLabel = room?.name || ci.camera;
  document.getElementById('ciPanelTitle').textContent = `Check-in · Camera ${camLabel}`;

  if (b) {
    document.getElementById('ciPanelSub').textContent = `${b.n} · ${fmtDate(b.s)} → ${fmtDate(b.e)}`;
  } else {
    const dataFmt = ci.data ? ci.data.split('-').reverse().join('/') : '—';
    document.getElementById('ciPanelSub').textContent = `Camera ${ci.camera} · ${dataFmt} · ${ci.numOspiti} ospite${ci.numOspiti!==1?'i':''}`;
  }

  renderCiGuests();
  document.getElementById('ciOverlay').classList.add('open');
  document.getElementById('ciPanel').scrollTop = 0;
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

  // Trova booking: per dbId, LOCAL- o CI-DIRECT- (da storico)
  let b = bookings.find(x => x.dbId === _ciEditBookingId)
       || bookings.find(x => ('LOCAL-' + x.id) === _ciEditBookingId);
  let existingCi = ciData[_ciEditBookingId];
  if (!existingCi && _ciEditBookingId.startsWith('CI-DIRECT-')) {
    const rawCiId = _ciEditBookingId.replace('CI-DIRECT-', '');
    existingCi = Object.values(ciData).find(c => c.ciId === rawCiId);
    if (existingCi && !b) {
      b = bookings.find(x => {
        const room2 = ROOMS.find(r => r.id === x.r);
        const cam2 = room2?.name || x.cameraName || '';
        const d2 = x.s instanceof Date ? x.s : new Date(x.s);
        const date2 = d2.getFullYear()+'-'+String(d2.getMonth()+1).padStart(2,'0')+'-'+String(d2.getDate()).padStart(2,'0');
        return String(cam2) === String(existingCi.camera) && date2 === existingCi.data;
      });
    }
  }
  const room = ROOMS.find(r => r.id === b?.r);
  const existing = existingCi || {};

  const ciId    = existing.ciId || 'CI-' + Date.now().toString(36).toUpperCase();
  const ciRow   = existing.ciRow || null;
  const localD  = d => { const dt = d instanceof Date ? d : new Date(d); return dt.getFullYear()+'-'+String(dt.getMonth()+1).padStart(2,'0')+'-'+String(dt.getDate()).padStart(2,'0'); };
  const data    = existing.data || (b ? localD(b.s) : localD(new Date()));
  const camera  = existing.camera || room?.name || b?.cameraName || '—';
  const preId   = b?.dbId || existing.preId || '';
  const record  = {
    ciId, preId,
    camera,
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
    // Aggiorna il tab CI nel drawer se è aperto
    const drPanel = document.getElementById('drTabCI');
    if (drPanel && drPanel.style.display !== 'none') {
      const booking = bookings.find(x => x.dbId === b?.dbId)
                   || bookings.find(x => ('LOCAL-' + x.id) === _ciEditBookingId);
      if (booking) {
        try { drPanel.innerHTML = renderDrawerCheckin(booking); } catch(e) {}
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

function exportAlloggiati(scope='72h'){
  // Date locali (no UTC shift)
  const localDate = n => { const d=new Date(Date.now()-n*864e5); return d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0')+'-'+String(d.getDate()).padStart(2,'0'); };
  const today=localDate(0), ieri=localDate(1), altroieri=localDate(2);
  let items;
  if(scope==='today'){items=Object.values(ciData).filter(ci=>ci.data===today);}
  else if(scope==='48h'||scope==='72h'){items=Object.values(ciData).filter(ci=>ci.data===today||ci.data===ieri||ci.data===altroieri);}
  else{items=Object.values(ciData).filter(ci=>{if(!_ciHistSearch)return true;const q=_ciHistSearch.toLowerCase();const cap=ci.guests[0]||{};return ci.camera.toLowerCase().includes(q)||((cap.cognome||'')+' '+(cap.nome||'')).toLowerCase().includes(q)||ci.data.includes(q);});}
  if(items.length===0){showToast('Nessun check-in da esportare','error');return;}
  const comuniMancanti=[],visti=new Set();
  items.forEach(ci=>{ci.guests.forEach((g,gIdx)=>{
    const isIta=_alNorm(g.cittadinanza||'ITALIA').includes('ITAL');
    if(isIta&&g.luogoNascita&&!_alCodiceComune(g.luogoNascita)){const k=_alNorm(g.luogoNascita);if(!visti.has('N:'+k)){visti.add('N:'+k);comuniMancanti.push({nomeComune:g.luogoNascita,label:'nascita - '+cleanAl(g.cognome)+' '+cleanAl(g.nome)});}}
    if(gIdx===0&&g.luogoRilascio&&!_alRisolviLuogo(g.luogoRilascio)&&!_alCodiceStato(g.luogoRilascio)){const k=_alNorm(g.luogoRilascio);if(!visti.has('R:'+k)){visti.add('R:'+k);comuniMancanti.push({nomeComune:g.luogoRilascio,label:'rilascio - '+cleanAl(g.cognome)+' '+cleanAl(g.nome)});}}
  });});
  if(comuniMancanti.length>0){_alChiediCodiciMancanti(comuniMancanti,()=>exportAlloggiati(scope));return;}
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
      const isIta=_alNorm(g.cittadinanza||'ITALIA').includes('ITAL');
      let comN='000000000',provN='  ';
      if(isIta&&g.luogoNascita){const found=_alCodiceComune(g.luogoNascita);if(found){comN=found.cod;provN=padR(found.prov,2);}}
      const statoN=_alCodiceStato(isIta?'ITALIA':(g.statoEsteroNascita||g.cittadinanza||'ITALIA')).padStart(9,'0');
      const cittad=_alCodiceStato(g.cittadinanza||'ITALIA').padStart(9,'0');
      const tipoDoc=isCapo?padR(_alCodiceDoc(g.tipoDoc),5):'     ';
      const numDoc=isCapo?padR(cleanAl(g.numDoc||''),20):'                    ';
      let luogoRil='000000000';
      if(isCapo&&g.luogoRilascio){
        const rl=_alRisolviLuogo(g.luogoRilascio);
        if(rl){luogoRil=rl.cod;}
        else{const sc=_alCodiceStato(g.luogoRilascio);if(sc&&sc!=='000000100')luogoRil=sc.padStart(9,'0');}}
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
      + '"tipoDoc":"IDENT o PASOR o PATEN o ALTRO","numDoc":"numero esatto senza spazi","luogoRilascio":""}';

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

// ── Toggle card ospite nel drawer ──
function ciDrToggle(cardId) {
  const body = document.getElementById(cardId);
  const arr  = document.getElementById(cardId + '_arr');
  if (!body) return;
  const open = body.style.display !== 'none';
  body.style.display = open ? 'none' : 'block';
  if (arr) arr.textContent = open ? '▸' : '▾';
}

// ── Anteprima TXT Alloggiati — modale con riga per ospite editabile ──
function ciPreviewAlloggiati(bookingIdOrScope) {
  // Determina gli items da mostrare
  const localDate = n => { const d=new Date(Date.now()-n*864e5); return d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0')+'-'+String(d.getDate()).padStart(2,'0'); };
  const today=localDate(0), ieri=localDate(1), altroieri=localDate(2);

  let items;
  if (bookingIdOrScope === 'filtered' || bookingIdOrScope === '72h' || bookingIdOrScope === '48h') {
    items = Object.values(ciData).filter(ci => ci.data===today||ci.data===ieri||ci.data===altroieri);
  } else {
    // bookingId numerico — trova il check-in abbinato
    const b = (typeof bookings !== 'undefined') ? bookings.find(x => String(x.id)===String(bookingIdOrScope)) : null;
    if (b) {
      const room = typeof ROOMS !== 'undefined' ? ROOMS.find(r => r.id === b.r) : null;
      const camName = room?.name || b.cameraName || '';
      const localD = d => { const dt=d instanceof Date?d:new Date(d); return dt.getFullYear()+'-'+String(dt.getMonth()+1).padStart(2,'0')+'-'+String(dt.getDate()).padStart(2,'0'); };
      const dataArr = localD(b.s);
      items = [ciData[b.dbId] || ciData['cam:'+camName+':'+dataArr]].filter(Boolean);
    }
    if (!items || !items.length) items = Object.values(ciData).filter(ci => ci.data===today||ci.data===ieri||ci.data===altroieri);
  }

  if (!items || items.length === 0) { showToast('Nessun check-in da esportare', 'error'); return; }

  // Costruisce la struttura dati editabile: array di righe ospite
  const rows = [];
  items.forEach(ci => {
    const b = typeof bookings !== 'undefined' ? bookings.find(bk => bk.dbId === ci.preId) : null;
    const nGiorni = b ? String(Math.min(Math.max(Math.round((new Date(b.e)-new Date(b.s))/86400000),1),30)).padStart(2,'0') : '01';
    ci.guests.forEach((g, gIdx) => {
      rows.push({
        ciId:      ci.ciId,
        isCapo:    gIdx === 0,
        gIdx,
        cognome:   g.cognome || '',
        nome:      g.nome || '',
        sesso:     g.sesso || 'M',
        dataNascita: g.dataNascita || '',
        luogoNascita: g.luogoNascita || '',
        provNascita:  g.provNascita || '',
        statoEsteroNascita: g.statoEsteroNascita || '',
        cittadinanza: g.cittadinanza || '',
        tipoDoc:   g.tipoDoc || '',
        numDoc:    g.numDoc || '',
        luogoRilascio: g.luogoRilascio || '',
        dataArrivo: ci.data || today,
        nGiorni,
        camera:    ci.camera || '',
      });
    });
  });

  // Costruisce HTML modale
  const rowsHtml = rows.map((r, idx) => {
    const isCapoLabel = r.isCapo ? '<span style="font-size:9px;background:var(--accent);color:#fff;padding:1px 5px;border-radius:3px;">CAPO</span>' : '<span style="font-size:9px;color:var(--text3);">fam.</span>';
    const docFields = r.isCapo ? `
      <input class="ci-prev-inp" placeholder="Tipo doc" value="${r.tipoDoc}" oninput="_ciPrevRows[${idx}].tipoDoc=this.value" style="width:60px">
      <input class="ci-prev-inp" placeholder="N° documento" value="${r.numDoc}" oninput="_ciPrevRows[${idx}].numDoc=this.value" style="flex:2">
      <input class="ci-prev-inp" placeholder="Luogo rilascio" value="${r.luogoRilascio}" oninput="_ciPrevRows[${idx}].luogoRilascio=this.value" style="flex:1.5">
    ` : `<span style="color:var(--text3);font-size:11px;padding:0 4px;">—</span>`;

    return `<div class="ci-prev-row">
      <div class="ci-prev-row-hdr">
        ${isCapoLabel}
        <strong style="font-size:12px;">${r.cognome} ${r.nome}</strong>
        <span style="font-size:11px;color:var(--text3);">Camera ${r.camera} · ${r.dataArrivo.split('-').reverse().join('/')}</span>
      </div>
      <div class="ci-prev-fields">
        <input class="ci-prev-inp" placeholder="Cognome *" value="${r.cognome}" oninput="_ciPrevRows[${idx}].cognome=this.value" style="flex:2">
        <input class="ci-prev-inp" placeholder="Nome *" value="${r.nome}" oninput="_ciPrevRows[${idx}].nome=this.value" style="flex:2">
        <select class="ci-prev-inp" onchange="_ciPrevRows[${idx}].sesso=this.value" style="width:70px">
          <option value="M" ${r.sesso==='M'?'selected':''}>M</option>
          <option value="F" ${r.sesso==='F'?'selected':''}>F</option>
        </select>
        <input class="ci-prev-inp" type="date" value="${r.dataNascita}" oninput="_ciPrevRows[${idx}].dataNascita=this.value" style="width:130px">
      </div>
      <div class="ci-prev-fields">
        <input class="ci-prev-inp" placeholder="Comune nascita" value="${r.luogoNascita}" oninput="_ciPrevRows[${idx}].luogoNascita=this.value" style="flex:2">
        <input class="ci-prev-inp" placeholder="Prov." value="${r.provNascita}" oninput="_ciPrevRows[${idx}].provNascita=this.value.toUpperCase();this.value=this.value.toUpperCase()" maxlength="2" style="width:50px">
        <input class="ci-prev-inp" placeholder="Cittadinanza" value="${r.cittadinanza}" oninput="_ciPrevRows[${idx}].cittadinanza=this.value.toUpperCase();this.value=this.value.toUpperCase()" style="flex:1.5">
        ${docFields}
      </div>
    </div>`;
  }).join('');

  // Crea/mostra overlay anteprima
  let overlay = document.getElementById('ciPrevOverlay');
  if (!overlay) {
    overlay = document.createElement('div');
    overlay.id = 'ciPrevOverlay';
    overlay.style.cssText = 'position:fixed;inset:0;background:rgba(0,0,0,.5);z-index:500;display:flex;flex-direction:column;';
    document.body.appendChild(overlay);
  }

  overlay.innerHTML = `
    <div style="background:var(--surface);flex:1;overflow-y:auto;padding:16px;max-width:600px;width:100%;margin:0 auto;">
      <div style="display:flex;align-items:center;gap:10px;margin-bottom:16px;">
        <button class="btn btn-icon" onclick="document.getElementById('ciPrevOverlay').style.display='none'">←</button>
        <div style="flex:1;font-family:'Playfair Display',serif;font-size:16px;font-weight:600;">Anteprima Alloggiati</div>
        <span style="font-size:11px;color:var(--text3);">${rows.length} ospiti</span>
      </div>
      <div id="ciPrevRows">${rowsHtml}</div>
    </div>
    <div style="background:var(--surface);border-top:1px solid var(--border);padding:12px 16px;display:flex;gap:10px;max-width:600px;width:100%;margin:0 auto;">
      <button class="btn" style="flex:1;justify-content:center;" onclick="document.getElementById('ciPrevOverlay').style.display='none'">Annulla</button>
      <button class="btn primary" style="flex:2;justify-content:center;" onclick="ciPrevGenerate()">⬇ Genera e scarica TXT</button>
    </div>`;

  window._ciPrevRows = rows;
  overlay.style.display = 'flex';
}

// Genera il TXT dall'anteprima editata
function ciPrevGenerate() {
  const rows = window._ciPrevRows;
  if (!rows || !rows.length) return;

  const toFmt = s => {
    if (!s) return '          ';
    const c = s.replace(/-/g,'').replace(/\//g,'');
    if (c.length !== 8) return '          ';
    if (/^(19|20)\d{6}$/.test(c)) return `${c.slice(6,8)}/${c.slice(4,6)}/${c.slice(0,4)}`;
    return `${c.slice(0,2)}/${c.slice(2,4)}/${c.slice(4,8)}`;
  };

  const lines = [];
  rows.forEach(r => {
    const isCapo    = r.isCapo;
    const tipoAllog = isCapo ? (rows.filter(x => x.ciId===r.ciId).length > 1 ? '17' : '16') : '19';
    const cognome   = padR(cleanAl(r.cognome), 50);
    const nome      = padR(cleanAl(r.nome), 30);
    const sesso     = (r.sesso||'M').toUpperCase() === 'F' ? '2' : '1';
    const dataN     = toFmt(r.dataNascita);
    const isIta     = _alNorm(r.cittadinanza||'ITALIA').includes('ITAL');
    let comN = '000000000', provN = '  ';
    if (isIta && r.luogoNascita) { const f=_alCodiceComune(r.luogoNascita); if(f){comN=f.cod;provN=padR(f.prov,2);} }
    const statoN = _alCodiceStato(isIta?'ITALIA':(r.statoEsteroNascita||r.cittadinanza||'ITALIA')).padStart(9,'0');
    const cittad = _alCodiceStato(r.cittadinanza||'ITALIA').padStart(9,'0');
    const tipoDoc = isCapo ? padR(_alCodiceDoc(r.tipoDoc), 5) : '     ';
    const numDoc  = isCapo ? padR(cleanAl(r.numDoc||''), 20) : ' '.repeat(20);
    let luogoRil = '000000000';
    if (isCapo && r.luogoRilascio) {
      const rl = _alRisolviLuogo(r.luogoRilascio);
      if (rl) { luogoRil = rl.cod; }
      else { const sc = _alCodiceStato(r.luogoRilascio); if (sc && sc !== '000000100') luogoRil = sc.padStart(9,'0'); }
    }
    const arrivo = r.dataArrivo.split('-').reverse().join('/').replace(/-/g,'/');
    const arrFmt = arrivo.length===10 ? arrivo : '          ';
    const record = tipoAllog + arrFmt + r.nGiorni + cognome + nome + sesso + dataN + comN + provN + statoN + cittad + tipoDoc + numDoc + luogoRil;
    if (record.length !== 168) console.warn('[Preview] len='+record.length, r.cognome);
    lines.push(record);
  });

  const content = lines.join('\r\n');
  const today = new Date(); const dd=String(today.getDate()).padStart(2,'0'), mm=String(today.getMonth()+1).padStart(2,'0'), yyyy=today.getFullYear();
  const blob = new Blob([content], {type:'text/plain;charset=utf-8'});
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement('a');
  a.href = url; a.download = `alloggiati_${yyyy}-${mm}-${dd}.txt`; a.click();
  URL.revokeObjectURL(url);
  document.getElementById('ciPrevOverlay').style.display = 'none';
  showToast('✓ File generato · ' + lines.length + ' record', 'success');
}

// ═══════════════════════════════════════════════════════════════════
// GESTIONE FOGLI ANNUALI NELLE IMPOSTAZIONI
// ═══════════════════════════════════════════════════════════════════


// ═══════════════════════════════════════════════════════════════════
// DRAWER CHECK-IN — Tab nella scheda prenotazione (gantt.js)
// ═══════════════════════════════════════════════════════════════════

/**
 * drTab(el, tabId) — switcha tra i tab del drawer prenotazione.
 * Definita qui (checkin.js caricato dopo gantt.js) per evitare duplicati.
 * Gestisce sia i tab esistenti (drTabInfo, drTabBill) che drTabCI.
 */
// drTabCheckin — gestisce SOLO il tab 🛎 Check-in nel drawer.
// Nome univoco per evitare conflitti con drTab di billing.js
function drTabCheckin(el, bookingId) {

  // Attiva il tab visivamente
  const tabsContainer = el.closest('.dr-bill-tabs');
  if (tabsContainer) {
    tabsContainer.querySelectorAll('.dr-bill-tab').forEach(t => t.classList.remove('active'));
    el.classList.add('active');
  }

  // Nascondi pannelli esistenti
  const parent = tabsContainer ? tabsContainer.parentElement : null;
  if (parent) {
    parent.querySelectorAll('[id^="drTab"]').forEach(p => p.style.display = 'none');
  }

  // Crea/trova il pannello CI
  let panel = parent ? parent.querySelector('#drTabCI') : null;
  if (!panel) {
    panel = document.createElement('div');
    panel.id = 'drTabCI';
    if (parent) parent.appendChild(panel);
  }
  panel.style.cssText = 'display:block;padding:0;';

  // Trova booking
  const allBooks = typeof bookings !== 'undefined' ? bookings : [];
  const b = allBooks.find(x => String(x.id) === String(bookingId));

  if (!b) {
    panel.innerHTML = `<div style="padding:12px;background:#fef2f2;border-radius:8px;font-size:12px;color:#991b1b;margin-top:8px;">
      ⚠ Booking non trovato (id=${bookingId}, tot=${allBooks.length})
    </div>`;
    return;
  }

  // Render immediato
  let html;
  try { html = renderDrawerCheckin(b); } catch(e) { html = ''; }

  if (!html || !html.trim()) {
    html = `<div style="padding:12px;background:#fff7ed;border-radius:8px;font-size:12px;color:#9a3412;margin-top:8px;">
      ⚠ Nessun contenuto (dbId=${b.dbId||'null'})
    </div>`;
  }
  panel.innerHTML = html;

  // Refresh silenzioso
  if (typeof loadCiData === 'function') {
    loadCiData().then(() => {
      if (!panel.isConnected) return;
      try {
        const u = renderDrawerCheckin(b);
        if (u && u.trim()) panel.innerHTML = u;
      } catch(e) {}
    }).catch(() => {});
  }
}

/**
 * renderDrawerCheckin(b) — HTML del tab 🛎 Check-in nel drawer.
 * Chiamata da gantt.js al render del drawer prenotazione.
 * Legge ciData (già in memoria) — nessuna fetch.
 */
function renderDrawerCheckin(b) {
  if (!b) return '';
  try {
    // Helper data locale (evita shift UTC: toISOString() usa UTC, non ora locale)
    const localDate = d => {
      const dt = d instanceof Date ? d : new Date(d);
      if (isNaN(dt.getTime())) return '';
      return dt.getFullYear() + '-'
        + String(dt.getMonth() + 1).padStart(2, '0') + '-'
        + String(dt.getDate()).padStart(2, '0');
    };

    // Lookup: per dbId, poi fallback cam:CAMERA:DATA_LOCALE
    const room    = typeof ROOMS !== 'undefined' ? ROOMS.find(r => r.id === b.r) : null;
    const camName = room?.name || b.cameraName || '';
    const dataArr = localDate(b.s);
    const altKey  = 'cam:' + camName + ':' + dataArr;
    const ci      = (ciData && b.dbId && ciData[b.dbId])
                 || (ciData && altKey && ciData[altKey])
                 || null;

    const now     = Date.now();
    const arrDate = b.s instanceof Date ? b.s : new Date(b.s);
    const arrival = isNaN(arrDate.getTime()) ? now : arrDate.getTime();
    const hoursToArrival = (arrival - now) / 36e5;

  // ── Check-in già fatto — card espandibili ──
  if (ci && ci.guests && ci.guests.length > 0) {
    const dataFmt = (ci.ts||'').slice(0,10).split('-').reverse().join('/');

    const cards = ci.guests.map((g, i) => {
      const isCapo = i === 0;
      const nomeDisplay = [g.cognome, g.nome].filter(Boolean).join(' ') || '—';
      const label = isCapo ? '👤 Capogruppo' : `👤 Accompagnatore ${i}`;
      const cardId = `ciDrCard_${b.id}_${i}`;

      const cittLabel = g.cittadinanza || '—';
      const nascitaLabel = g.luogoNascita
        ? (g.provNascita ? `${g.luogoNascita} (${g.provNascita})` : g.luogoNascita)
        : (g.statoEsteroNascita || '—');
      const dnFmt = g.dataNascita ? g.dataNascita.split('-').reverse().join('/') : '—';

      const docSection = isCapo ? `
        <div class="ci-dr-card-row"><span>Documento</span><span>${g.tipoDoc||'—'} ${g.numDoc||''}</span></div>
        <div class="ci-dr-card-row"><span>Rilasciato a</span><span>${g.luogoRilascio||'—'}</span></div>` : '';

      return `
        <div class="ci-dr-card" onclick="ciDrToggle('${cardId}')">
          <div class="ci-dr-card-hdr">
            <div>
              <div class="ci-dr-card-label">${label}</div>
              <div class="ci-dr-card-name">${nomeDisplay}</div>
            </div>
            <span class="ci-dr-card-arrow" id="${cardId}_arr">▸</span>
          </div>
          <div class="ci-dr-card-body" id="${cardId}" style="display:none;">
            <div class="ci-dr-card-row"><span>Data nascita</span><span>${dnFmt}</span></div>
            <div class="ci-dr-card-row"><span>Luogo nascita</span><span>${nascitaLabel}</span></div>
            <div class="ci-dr-card-row"><span>Cittadinanza</span><span>${cittLabel}</span></div>
            ${docSection}
          </div>
        </div>`;
    }).join('');

    return `
      <div class="ci-dr-done-badge">✓ Check-in · ${ci.numOspiti} ospite${ci.numOspiti!==1?'i':''} · ${dataFmt}</div>
      <div class="ci-dr-cards">${cards}</div>
      <div style="display:flex;gap:8px;margin-top:10px;">
        <button class="btn" style="flex:1;justify-content:center;" onclick="openCiModal('${b.id}')">✎ Modifica</button>
        <button class="btn primary" style="flex:1;justify-content:center;" onclick="ciPreviewAlloggiati('${b.id}')">⬇ Alloggiati</button>
      </div>`;
  }

  // ── Check-in NON fatto ──
  // In ritardo (arrivo passato)
  if (hoursToArrival < 0) {
    const giorni = Math.abs(Math.floor(hoursToArrival / 24));
    return `
      <div class="ci-dr-alert danger">
        ⚠ Check-in in ritardo<br>
        <span style="font-size:11px;opacity:.85;">Arrivo ${giorni > 0 ? giorni+' giorn'+(giorni===1?'o':'i')+' fa' : 'oggi'} — non ancora registrato</span>
      </div>
      <button class="btn primary" style="width:100%;justify-content:center;margin-top:4px;" onclick="openCiModal('${b.id}')">
        🛎 Registra check-in ora
      </button>`;
  }

  // Entro 48h dall'arrivo
  if (hoursToArrival <= 48) {
    const ore = Math.floor(hoursToArrival);
    const label = ore < 1 ? 'Meno di 1 ora' : ore < 24 ? `${ore} or${ore===1?'a':'e'}` : `${Math.floor(ore/24)} giorn${Math.floor(ore/24)===1?'o':'i'}`;
    return `
      <div class="ci-dr-alert warn">
        🕐 Arrivo tra ${label}<br>
        <span style="font-size:11px;opacity:.85;">Prepara il check-in in anticipo</span>
      </div>
      <button class="btn primary" style="width:100%;justify-content:center;margin-top:4px;" onclick="openCiModal('${b.id}')">
        🛎 Fai check-in
      </button>`;
  }

  // Arrivo lontano (>48h)
  const gg = Math.floor(hoursToArrival / 24);
  return `
    <div class="ci-dr-future">
      Arrivo tra <strong>${gg} giorni</strong><br>
      <span style="font-size:11px;color:var(--text3);">Il check-in può essere fatto fino a 48h prima o dopo l'arrivo.</span>
    </div>
    <button class="btn" style="width:100%;justify-content:center;margin-top:12px;" onclick="openCiModal('${b.id}')">
      🛎 Fai check-in in anticipo
    </button>`;
  } catch(e) {
    return `<div style="padding:12px;background:#fef2f2;border:1px solid #fecaca;border-radius:8px;font-size:12px;color:#991b1b;">
      ⚠ Errore rendering check-in: ${e.message}
    </div>`;
  }
}
