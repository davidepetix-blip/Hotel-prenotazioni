// ═══════════════════════════════════════════════════════════════════
// checkin.js — Check-in operativo, Alloggiati Web, OCR Gemini
// Blip Hotel Management — build 18.7.x
// Dipende da: core.js, sync.js
// ═══════════════════════════════════════════════════════════════════


const BLIP_VER_CHECKIN = '8'; // ← incrementa ad ogni modifica

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
      const preId = row[1] || '';
      if (!preId) return;
      try {
        result[preId] = {
          ciId:   row[0] || '',
          preId:  row[1] || '',
          camera: row[2] || '',
          data:   row[3] || '',
          numOspiti: parseInt(row[4]) || 0,
          guests: JSON.parse(row[5] || '[]'),
          ts:     row[6] || '',
          utente: row[7] || '',
          ciRow:  i + 2,
        };
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
        <button class="btn primary" style="flex-shrink:0" onclick="exportAlloggiati('today')">⬇ Genera .txt</button>
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
  const b = bookings.find(b => b.dbId === bookingDbId);
  if (!b) { showToast('Prenotazione non trovata', 'error'); return; }
  _ciEditBookingId = bookingDbId;
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

function openCiModalFromCi(preId) {
  openCiModal(preId);
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
  const existing = ciData[_ciEditBookingId];

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
  } catch(e) {
    hideLoading();
    showToast('Errore: ' + e.message, 'error');
  }
}

// ── Generazione file Alloggiati Web ──
// Formato fisso: record da 93 caratteri (versione 3)
// Ref: specifiche tecniche Polizia di Stato – tracciato record tipo 20
// --- TABELLE ALLOGGIATI WEB build 18.7.1 ---

function exportAlloggiati(scope='today'){
  const today=new Date().toISOString().slice(0,10);
  let items;
  if(scope==='today'){items=Object.values(ciData).filter(ci=>ci.data===today);}
  else{items=Object.values(ciData).filter(ci=>{if(!_ciHistSearch)return true;const q=_ciHistSearch.toLowerCase();const cap=ci.guests[0]||{};return ci.camera.toLowerCase().includes(q)||((cap.cognome||'')+' '+(cap.nome||'')).toLowerCase().includes(q)||ci.data.includes(q);});}
  if(items.length===0){showToast('Nessun check-in da esportare','error');return;}
  const comuniMancanti=[],visti=new Set();
  items.forEach(ci=>{ci.guests.forEach((g,gIdx)=>{
    const isIta=_alNorm(g.cittadinanza||'ITALIA').includes('ITAL');
    if(isIta&&g.luogoNascita&&!_alCodiceComune(g.luogoNascita)){const k=_alNorm(g.luogoNascita);if(!visti.has('N:'+k)){visti.add('N:'+k);comuniMancanti.push({nomeComune:g.luogoNascita,label:'nascita - '+cleanAl(g.cognome)+' '+cleanAl(g.nome)});}}
    if(gIdx===0&&g.luogoRilascio&&!_alRisolviLuogo(g.luogoRilascio)){const k=_alNorm(g.luogoRilascio);if(!visti.has('R:'+k)){visti.add('R:'+k);comuniMancanti.push({nomeComune:g.luogoRilascio,label:'rilascio - '+cleanAl(g.cognome)+' '+cleanAl(g.nome)});}}
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
function drTab(el, tabId, bookingId) {
  // Switch active tab
  const tabsContainer = el.closest('.dr-bill-tabs');
  if (!tabsContainer) return;
  tabsContainer.querySelectorAll('.dr-bill-tab').forEach(t => t.classList.remove('active'));
  el.classList.add('active');

  // Trova i pannelli nel parent (drbody) — querySelector locale, non getElementById
  const parent = tabsContainer.parentElement;
  if (!parent) return;

  // Nascondi tutti i pannelli
  parent.querySelectorAll('[id^="drTab"]').forEach(p => p.style.display = 'none');

  // Mostra il pannello target
  const panel = parent.querySelector('#' + tabId);
  if (panel) panel.style.display = '';

  // ── Lazy render Check-in ──
  if (tabId === 'drTabCI') {
    if (!panel) {
      // Pannello non trovato — crea al volo come ultimo figlio
      const fallback = document.createElement('div');
      fallback.id = 'drTabCI';
      fallback.style.padding = '12px';
      fallback.innerHTML = '<div style="color:red;font-size:12px;">Errore: pannello check-in non trovato nel DOM</div>';
      parent.appendChild(fallback);
      return;
    }

    const bid = bookingId !== undefined ? bookingId : panel.dataset.bookingId;
    panel.innerHTML = '<div style="padding:20px;text-align:center;color:var(--text3);font-size:12px;">⏳ Caricamento…</div>';

    const allBooks = typeof bookings !== 'undefined' ? bookings : [];
    const b = allBooks.find(x => String(x.id) === String(bid));

    if (!b) {
      panel.innerHTML = `<div style="padding:12px;background:#fef2f2;border:1px solid #fecaca;border-radius:8px;font-size:12px;color:#991b1b;">
        ⚠ Prenotazione non trovata (id=${bid})<br>
        <small>Prenotazioni in memoria: ${allBooks.length}</small><br>
        <button class="btn" style="margin-top:8px" onclick="openCheckin()">Apri modulo check-in</button>
      </div>`;
      return;
    }

    if (!b.dbId) {
      panel.innerHTML = `<div style="padding:12px;background:#fff7ed;border:1px solid #fed7aa;border-radius:8px;font-size:12px;color:#9a3412;">
        ⚠ Prenotazione non ancora nel database.<br>Premi <strong>↻</strong> per sincronizzare, poi riprova.
      </div>`;
      return;
    }

    loadCiData().then(() => {
      panel.innerHTML = renderDrawerCheckin(b);
    }).catch(e => {
      panel.innerHTML = `<div style="padding:12px;color:red;font-size:12px;">Errore: ${e.message}</div>`;
    });
  }
}

/**
 * renderDrawerCheckin(b) — HTML del tab 🛎 Check-in nel drawer.
 * Chiamata da gantt.js al render del drawer prenotazione.
 * Legge ciData (già in memoria) — nessuna fetch.
 */
function renderDrawerCheckin(b) {
  if (!b) return '';
  const ci      = ciData && ciData[b.dbId];
  const now     = Date.now();
  const arrival = b.s instanceof Date ? b.s.getTime() : new Date(b.s).getTime();
  const hoursToArrival = (arrival - now) / 36e5; // negativo = già arrivato

  // ── Check-in già fatto ──
  if (ci && ci.guests && ci.guests.length > 0) {
    const capo   = ci.guests[0];
    const nomeCapo = [capo.cognome, capo.nome].filter(Boolean).join(' ') || '—';
    const accomp = ci.guests.slice(1);

    const guestRows = ci.guests.map((g, i) => {
      const nome  = [g.cognome, g.nome].filter(Boolean).join(' ') || '—';
      const label = i === 0 ? 'Capogruppo' : `Accompagnatore ${i}`;
      return `<div class="ci-dr-guest-row">
        <span class="ci-dr-guest-label">${label}</span>
        <span class="ci-dr-guest-name">${nome}</span>
      </div>`;
    }).join('');

    return `
      <div class="ci-dr-done-badge">✓ Check-in registrato</div>
      <div class="ci-dr-guests-box">
        ${guestRows}
        <div class="ci-dr-meta">${ci.numOspiti} ospite${ci.numOspiti!==1?'i':''} · registrato il ${(ci.ts||'').slice(0,10).split('-').reverse().join('/')}</div>
      </div>
      <div style="display:flex;gap:8px;margin-top:12px;">
        <button class="btn" style="flex:1;justify-content:center;" onclick="openCiModal('${b.dbId}')">✎ Modifica</button>
        <button class="btn" style="flex:1;justify-content:center;" onclick="exportAlloggiati('filtered')">⬇ Alloggiati</button>
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
      <button class="btn primary" style="width:100%;justify-content:center;margin-top:4px;" onclick="openCiModal('${b.dbId}')">
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
      <button class="btn primary" style="width:100%;justify-content:center;margin-top:4px;" onclick="openCiModal('${b.dbId}')">
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
    <button class="btn" style="width:100%;justify-content:center;margin-top:12px;" onclick="openCiModal('${b.dbId}')">
      🛎 Fai check-in in anticipo
    </button>`;
}
