// ═══════════════════════════════════════════════════════════════════
// rooms.js — Dashboard camere, impostazioni, stato operativo
// Blip Hotel Management — build 18.10.5
//
// Responsabilità:
//   • Impostazioni camere (maxGuests, tipi letto)
//   • Dashboard stato giornaliero (pulizie, interventi)
//   • Modal modifica stato singola camera
//   • Drawer info camera (openRoomDrawer)
//   • Configurazione fogli annuali + test connessione DB
//
// Dipende da: core.js, api.js, auth.js, store.js
// Caricato PRIMA di: gantt.js
// ═══════════════════════════════════════════════════════════════════

const BLIP_VER_ROOMS = '1';

function openRoomDrawer(roomId) {
  const room = ROOMS.find(r=>r.id===roomId);
  if (!room) return;
  const rs    = roomStates[roomId] || {};
  const pLabel= {'pulita':'Pulita','da-pulire':'Da pulire','in-corso':'Controllare/Rassettare','fuori-servizio':'Fuori servizio'};
  // Apri drawer con info camera
  document.getElementById('drtitle').textContent = 'Camera ' + room.name;
  document.getElementById('drsub').textContent   = room.g;
  const stateHtml = `<div style="background:var(--surface2);border-radius:8px;padding:10px 12px;margin-bottom:12px;border:1px solid var(--border);">
    <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:6px;">
      <span style="font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.08em;color:var(--text3);">Stato camera</span>
      <button class="btn" style="font-size:10px;padding:3px 8px;" onclick="openRstateModal('${roomId}')">✎ Modifica</button>
    </div>
    <span class="rdash-badge badge-${rs.pulizia||'da-pulire'}">${pLabel[rs.pulizia||'da-pulire']||'—'}</span>
    ${rs.configurazione?'<div style="font-size:11px;color:var(--text2);margin-top:5px;">🛏 '+rs.configurazione+'</div>':''}
    ${rs.noteOps?'<div style="font-size:10px;color:var(--text3);margin-top:3px;font-style:italic;">'+rs.noteOps+'</div>':''}
  </div>`;
  // Prenotazioni della camera nel mese corrente
  const ms=new Date(curY,curM,1), me=new Date(curY,curM+1,0);
  const rbs=bookings.filter(b=>b.r===roomId&&b.s<=me&&b.e>=ms);
  let bhtml='';
  rbs.forEach(b=>{ bhtml+=`<div style="background:${b.c};padding:8px 10px;border-radius:6px;margin-bottom:6px;font-size:12px;"><strong>${b.n}</strong><br><span style="font-size:10px;">${fmt(b.s)} → ${fmt(b.e)}</span>${b.d?'<br><span style="font-size:10px;">'+b.d+'</span>':''}</div>`; });
  document.getElementById('drbody').innerHTML = stateHtml + (bhtml||'<div class="empty" style="padding:20px 0;"><div style="font-size:11px;color:var(--text3);">Nessuna prenotazione questo mese</div></div>');
  openDrawer();
}

// ═══════════════════════════════════════════════════════════════════
// INIT
// ═══════════════════════════════════════════════════════════════════
// ═══════════════════════════════════════════════════════════════════
// IMPOSTAZIONI CAMERE
// ═══════════════════════════════════════════════════════════════════
function openSettings() {
  renderSettingsBody();
  document.getElementById('settingsPage').classList.add('open');
}
function closeSettings() {
  document.getElementById('settingsPage').classList.remove('open');
}
function resetRoomSettings() {
  if (!confirm('Ripristinare le impostazioni predefinite per tutte le camere?')) return;
  roomSettings = JSON.parse(JSON.stringify(ROOM_DEFAULTS));
  saveRoomSettingsLS(roomSettings);
  renderSettingsBody();
  showToast('Impostazioni ripristinate', 'success');
}

function saveSettings() {
  // ── Salva fogli annuali
  const list = document.getElementById('annualSheetsList');
  if (list) {
    const newEntries = [];
    Array.from(list.children).forEach((row, i) => {
      const year = parseInt(document.getElementById('asr-year-' + i)?.value);
      const sid  = (document.getElementById('asr-id-'   + i)?.value || '').trim();
      // Salva la riga se ha almeno l'anno valido — anche con sheetId vuoto.
      // Prima il filtro "if (year && sid)" impediva di eliminare anni senza ID
      // perché le righe rimanenti non venivano mai scritte su localStorage.
      if (year) newEntries.push({ year, sheetId: sid, label: String(year) });
    });
    // Salva sempre, anche se la lista è vuota (l'utente ha rimosso tutte le righe)
    annualSheets = newEntries;
    saveAnnualSheetsLS(newEntries);
  }
  // ── Salva ID DATABASE
  const dbId = (document.getElementById('sDbSheetId')?.value || '').trim();
  if (dbId !== undefined) { DATABASE_SHEET_ID = dbId; saveDbSheetIdLS(dbId); }

  // Leggi tutti i valori dalla pagina impostazioni
  ROOMS.forEach(room => {
    const maxEl = document.getElementById('smax-' + room.id);
    if (!maxEl) return;
    const maxGuests = parseInt(maxEl.value) || 2;
    const allowedBeds = BED_TYPES
      .filter(bt => {
        const chip = document.getElementById('schip-' + room.id + '-' + bt.id);
        return chip && chip.classList.contains('on');
      })
      .map(bt => bt.id);
    roomSettings[room.id] = { maxGuests, allowedBeds: allowedBeds.length ? allowedBeds : ['m','s'] };
  });
  saveRoomSettingsLS(roomSettings);
  // Aggiorna roomStates e salva su DB
  ROOMS.forEach(room => {
    if (!roomStates[room.id]) roomStates[room.id] = {};
    roomStates[room.id].maxGuests   = roomSettings[room.id]?.maxGuests   || 2;
    roomStates[room.id].allowedBeds = roomSettings[room.id]?.allowedBeds || ['m','s'];
  });
  if (DATABASE_SHEET_ID) writeRoomsSheet(roomStates).catch(e => console.warn('writeRoomsSheet:', e.message));
  closeSettings();
  showToast('✓ Impostazioni salvate', 'success');
}


function renderSettingsBody() {
  const body = document.getElementById('settingsBody');
  const groups = [...new Set(ROOMS.map(r => r.g))];
  let html = '';

  // ── Sezione Alloggiati Web + OCR (Gemini diretto, no proxy)
  const _codStr    = localStorage.getItem('hotelCodiceStruttura') || '';
  const _geminiKey = localStorage.getItem('hotelGeminiApiKey') || '';
  html += `
    <div class="sgroup">
      <div class="sgroup-title">🛎 Alloggiati Web & OCR</div>
      <div class="sroom-card">
        <p style="font-size:11px;color:var(--text2);margin-bottom:12px;line-height:1.7;">
          Codice struttura ricettiva (8 cifre) fornito dalla Questura.
        </p>
        <div style="display:flex;align-items:center;gap:10px;margin-bottom:16px;">
          <input type="text" class="fi" id="codStrutturaInput" value="${_codStr}"
                 placeholder="es. 00123456" maxlength="9" style="max-width:180px;"
                 onchange="localStorage.setItem('hotelCodiceStruttura',this.value)">
          <span style="font-size:11px;color:var(--text3);">Codice struttura</span>
        </div>
        <div style="border-top:1px solid var(--border);padding-top:14px;">
          <p style="font-size:11px;color:var(--text2);margin-bottom:8px;line-height:1.7;">
            <strong>API Key Gemini</strong> — gratuita, per la scansione automatica dei documenti.<br>
            Ottienila su <a href="https://aistudio.google.com/apikey" target="_blank"
            style="color:var(--accent)">aistudio.google.com/apikey</a> → Create API Key.
          </p>
          <input type="text" class="fi" value="${_geminiKey}"
                 placeholder="AIza..."
                 style="font-size:12px;font-family:monospace;"
                 onchange="localStorage.setItem('hotelGeminiApiKey',this.value.trim())"
                 autocomplete="off" autocorrect="off" autocapitalize="off" spellcheck="false">
          <div style="margin-top:8px;display:flex;gap:8px;align-items:center;flex-wrap:wrap;">
            <button class="btn" style="font-size:11px;" onclick="testGeminiKey()">Testa chiave</button>
            <span id="proxyTestResult" style="font-size:11px;color:var(--text3);"></span>
          </div>
          <div id="proxyTestDetail" style="margin-top:6px;font-size:10px;color:var(--text2);line-height:1.6;"></div>
        </div>
      </div>
    </div>`;

  // ── Sezione Fogli annuali
  const annEntries = (loadAnnualSheets().length ? loadAnnualSheets() : DEFAULT_ANNUAL_SHEETS);
  const currentDbId = loadDbSheetId();
  html += `
    <div class="sgroup">
      <div class="sgroup-title">📅 Fogli Google Annuali</div>
      <div class="sroom-card">
        <p style="font-size:11px;color:var(--text2);margin-bottom:12px;line-height:1.7;">
          Aggiungi un foglio per ogni anno gestito. L'app leggerà le prenotazioni da tutti i fogli configurati.
        </p>
        <div id="annualSheetsList"></div>
        <button class="btn" style="font-size:11px;margin-top:8px;" onclick="addAnnualSheetRow()">+ Aggiungi anno</button>
      </div>
    </div>
    <div class="sgroup">
      <div class="sgroup-title">🗄️ Foglio DATABASE Centrale</div>
      <div class="sroom-card">
        <p style="font-size:11px;color:var(--text2);margin-bottom:10px;line-height:1.7;">
          Foglio Google separato con scheda <b>PRENOTAZIONI</b> — fonte di verità centrale per tutti gli anni.
          Crea il foglio, poi incolla il suo ID qui sotto.
        </p>
        <div class="sfield">
          <label>ID Foglio Database</label>
          <input type="text" id="sDbSheetId" value="${currentDbId}"
                 placeholder="es. 1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgVE2upms"
                 style="font-size:11px;font-family:monospace;">
        </div>
        <div style="margin-top:8px;">
          <button class="btn" style="font-size:11px;" onclick="testDbConnection()">🔗 Testa connessione DB</button>
          <span id="dbTestResult" style="font-size:11px;color:var(--text3);margin-left:8px;"></span>
        </div>
        <details style="margin-top:12px;">
          <summary style="font-size:11px;font-weight:600;color:var(--accent);cursor:pointer;user-select:none;">Come trovare l'ID del foglio ▾</summary>
          <p style="font-size:11px;color:var(--text2);margin-top:8px;line-height:1.8;">
            Apri il foglio Google nel browser. L'URL è:<br>
            <code style="font-size:10px;background:var(--surface2);padding:2px 6px;border-radius:4px;">docs.google.com/spreadsheets/d/<b>ID_QUI</b>/edit</code><br>
            Copia la parte in grassetto.
          </p>
        </details>
      </div>
    </div>`;

  // Popola lista fogli annuali con JS (dopo render)
  setTimeout(() => {
    const list = document.getElementById('annualSheetsList');
    if (!list) return;
    list.innerHTML = '';
    const entries = loadAnnualSheets();
    entries.forEach((e, i) => renderAnnualSheetRow(list, e, i));
  }, 50);

  groups.forEach(g => {
    html += `<div class="sgroup"><div class="sgroup-title">${g}</div>`;
    ROOMS.filter(r => r.g === g).forEach(room => {
      const cfg = roomSettings[room.id] || ROOM_DEFAULTS[room.id] || { maxGuests:2, allowedBeds:['m','s'] };
      html += `
        <div class="sroom-card">
          <div class="sroom-name">
            ${room.name}
            <span class="sroom-badge">${g}</span>
          </div>
          <div class="sroom-fields">
            <div class="sfield">
              <label>Capienza max ospiti</label>
              <input type="number" id="smax-${room.id}" value="${cfg.maxGuests}" min="1" max="20">
            </div>
          </div>
          <div class="sbed-types">
            <label>Tipi di letto consentiti</label>
            <div class="sbed-chips">
              ${BED_TYPES.map(bt => `
                <div class="sbed-chip ${cfg.allowedBeds.includes(bt.id) ? 'on' : ''}"
                     id="schip-${room.id}-${bt.id}"
                     onclick="this.classList.toggle('on')">
                  ${bt.label} (${bt.id})
                </div>
              `).join('')}
            </div>
          </div>
        </div>`;
    });
    html += '</div>';
  });

  // ── Sezione Manutenzione Database ──
  html += `
    <div class="sgroup">
      <div class="sgroup-title">🔧 Manutenzione Database</div>
      <div class="sroom-card" style="display:flex;flex-direction:column;gap:10px;">
        <p style="font-size:11px;color:var(--text2);margin:0;line-height:1.7;">
          Strumenti per riparare i legami univoci nel database (CONTI, PAGAMENTI, CHECK-IN)
          e per scrivere i BLIP_ID nella riga 46 del foglio Google.
          Usare dopo anomalie o migrazioni.
        </p>
        <div style="display:flex;gap:8px;flex-wrap:wrap;">
          <button class="btn" onclick="closeSettings();riparaDatabase()" style="flex:1;min-width:140px;">
            🔧 Ripara database
          </button>
          <button class="btn" onclick="closeSettings();backfillRow46()" style="flex:1;min-width:140px;">
            #46 Scrivi BLIP_ID riga 46
          </button>
        </div>
      </div>
    </div>`;

  body.innerHTML = html;
}

// ═══════════════════════════════════════════════════════════════════
// DASHBOARD CAMERE
// ═══════════════════════════════════════════════════════════════════

// ═══════════════════════════════════════════════════════════════════
// STATO OPERATIVO GIORNALIERO — calcola situazione di ogni camera oggi
// ═══════════════════════════════════════════════════════════════════

// Ritorna lo stato operativo di una camera per una data specifica
// { opId, label, uscente, entrante } 
// ── Calcola interventi automatici del giorno per una camera ──
// Ritorna array di { tipo, label, urgenza }
// tipo ∈ 'rassetto' | 'lenzuola'
// urgenza ∈ 'alta' | 'normale'

const GRUPPI_RASSETTO   = ['Scuola', 'Largo Roma'];     // solo questi gruppi
const SOGLIA_LENZUOLA   = 7;                             // ogni N notti dal check-in

function getRoomInterventions(roomId, date, room) {
  const d = new Date(date); d.setHours(12,0,0,0);
  const dMs = d.getTime();
  const interventions = [];

  // Cerca prenotazione attiva (ospite già dentro, non il giorno di check-in)
  const occ = bookings.find(b =>
    b.r === roomId &&
    b.s.getTime() < dMs &&   // check-in PRIMA di oggi
    b.e.getTime() > dMs      // check-out DOPO oggi
  );
  if (!occ) return interventions;

  // Camere/appartamenti con disposizione "aff" → gestione autonoma, nessun intervento
  const dispStr = (occ.d || '').toLowerCase();
  const isAff = dispStr.includes('aff');
  if (isAff || room.g === 'Appartamenti') {
    // Solo cambio lenzuola per appartamenti (non aff), mai rassetto
    if (room.g === 'Appartamenti' && !isAff) {
      const nottiPassate = Math.round((dMs - occ.s.getTime()) / 86400000);
      if (nottiPassate > 0 && nottiPassate % SOGLIA_LENZUOLA === 0) {
        interventions.push({ tipo:'lenzuola', label:'Cambio lenzuola', urgenza:'alta' });
      }
    }
    return interventions;
  }

  const nottiTotali  = Math.round((occ.e - occ.s) / 86400000);
  const nottiPassate = Math.round((dMs - occ.s.getTime()) / 86400000);
  const isCheckin    = occ.s.getTime() === dMs;

  // ── RASSETTO: Scuola + Largo Roma, soggiorni > 1 notte, non il giorno di check-in ──
  if (GRUPPI_RASSETTO.includes(room.g) && nottiTotali > 1 && !isCheckin) {
    interventions.push({ tipo:'rassetto', label:'Rassetto', urgenza:'normale' });
  }

  // ── CAMBIO LENZUOLA: ogni 7 notti esatte dal check-in ──
  if (nottiPassate > 0 && nottiPassate % SOGLIA_LENZUOLA === 0) {
    interventions.push({ tipo:'lenzuola', label:'Cambio lenzuola', urgenza:'alta' });
  }

  return interventions;
}

// opId ∈ occupata|cambio|uscita|arrivo|pronta|da-preparare|fuori-servizio
function getRoomDayStatus(roomId, date) {
  const d = new Date(date); d.setHours(12,0,0,0);
  const dStr = d.getTime();
  const rs = roomStates[roomId] || {};

  // Fuori servizio → priorità massima
  if (rs.pulizia === 'fuori-servizio') {
    return { opId:'fuori-servizio', label:'Fuori servizio', uscente:null, entrante:null };
  }

  // Trova prenotazioni attive / che iniziano / che finiscono in questa data
  const occupante = bookings.find(b =>
    b.r === roomId &&
    b.s.getTime() < dStr &&
    b.e.getTime() > dStr
  );
  const uscente = bookings.find(b =>
    b.r === roomId &&
    b.e.getTime() === dStr
  );
  const entrante = bookings.find(b =>
    b.r === roomId &&
    b.s.getTime() === dStr
  );

  if (uscente && entrante) {
    return { opId:'cambio', label:'Cambio cliente', uscente, entrante };
  }
  if (uscente) {
    return { opId:'uscita', label:'Check-out oggi', uscente, entrante:null };
  }
  if (entrante) {
    return { opId:'arrivo', label:'Arrivo oggi', uscente:null, entrante };
  }
  if (occupante) {
    return { opId:'occupata', label:'Occupata', uscente:null, entrante:null, occupante };
  }

  // Camera libera — cerca il prossimo arrivo futuro
  const prossimoArrivo = bookings
    .filter(b => b.r === roomId && b.s.getTime() > dStr)
    .sort((a,b) => a.s - b.s)[0] || null;

  // Distingui "controllare/rassettare" da semplice "da pulire"
  const pulizia = rs.pulizia || 'da-pulire';
  if (pulizia === 'pulita') {
    return { opId:'pronta', label:'Pronta', uscente:null, entrante:null, prossimoArrivo };
  }
  if (pulizia === 'in-corso') {
    return { opId:'controllare', label:'Controllare/Rassettare', uscente:null, entrante:null, prossimoArrivo };
  }
  // da-pulire
  return { opId:'da-preparare', label:'Da preparare', uscente:null, entrante:null, prossimoArrivo };
}

function puliziaBadge(p) {
  const labels = { 'pulita':'Pulita','da-pulire':'Da pulire','in-corso':'Controllare/Rassettare','fuori-servizio':'Fuori servizio' };
  return `<span class="rdash-badge badge-${p}">${labels[p]||p}</span>`;
}

function opBadge(opId, label) {
  return `<span class="rdash-badge badge-op-${opId}">${label}</span>`;
}

function openRoomDash() {
  _rdashFilter = 'tutti';
  renderRoomDash();
  document.getElementById('roomDashPage').classList.add('open');
}
function closeRoomDash() {
  document.getElementById('roomDashPage').classList.remove('open');
}

function renderRoomDash() {
  const today = new Date(); today.setHours(12,0,0,0);
  const todayStr = today.toLocaleDateString('it-IT',{weekday:'long',day:'numeric',month:'long'});

  // Filtri operativi
  const filtri = ['tutti','interventi','cambio','controllare','da-preparare','arrivo','uscita','occupata','pronta','fuori-servizio'];
  const fLabel  = {
    'tutti':'Tutti', 'interventi':'🧹 Interventi', 'cambio':'🔄 Cambio', 'controllare':'🟣 Controlla',
    'occupata':'🔴 Occupate', 'arrivo':'🔵 Arrivi', 'uscita':'🟡 Uscite',
    'pronta':'🟢 Pronte', 'da-preparare':'🟠 Da pulire', 'fuori-servizio':'⚫ Fuori serv.'
  };
  let filtersHtml = `<div style="padding:10px 14px 4px;font-size:11px;color:var(--text3);font-weight:600;">Oggi: ${todayStr}</div>
  <div class="rdash-filters">`;
  filtri.forEach(f => {
    filtersHtml += `<button class="rdash-filter-btn${_rdashFilter===f?' active':''}" onclick="_rdashFilter='${f}';renderRoomDash()">${fLabel[f]}</button>`;
  });
  filtersHtml += `</div>`;

  // ── Ordinamento per ordine di lavoro del personale ──
  // Bucket numerici: numero più basso = più urgente
  // 0  cambio cliente (doppio lavoro oggi)
  // 1  arrivo oggi + camera non pulita  (da fare SUBITO)
  // 2  uscita oggi (checkout → poi pulire)
  // 3  da preparare/controllare con arrivo imminente (oggi o domani)
  // 4  arrivo oggi + camera già pulita (pronta, nessuna azione)
  // 5  da preparare/controllare con arrivo tra 2-7gg
  // 6  da preparare/controllare con arrivo tra 8+gg
  // 7  da preparare/controllare senza arrivo previsto
  // 8  occupata (nessuna azione oggi), ordinate per checkout
  // 9  pronta senza arrivo imminente
  // 10 fuori servizio

  const todayMs = today.getTime();

  const allCards = [];
  ROOMS.forEach(room => {
    const st = getRoomDayStatus(room.id, today);
    const _ivs_filter = (st.opId === 'occupata') ? getRoomInterventions(room.id, today, room) : [];
    if (_rdashFilter === 'interventi') {
      if (_ivs_filter.length === 0) return; // nascondi camere senza interventi
    } else if (_rdashFilter !== 'tutti' && st.opId !== _rdashFilter) return;

    const rs    = roomStates[room.id] || {};
    const pul   = rs.pulizia || 'da-pulire';
    const gArr  = st.prossimoArrivo ? Math.round((st.prossimoArrivo.s - today) / 86400000) : null;

    let bucket, tiebreak;

    switch (st.opId) {
      case 'cambio':
        bucket = 0; tiebreak = 0; break;

      case 'arrivo':
        // Arrivo oggi: prima le non pulite (da fare subito), poi le già pronte
        bucket   = (pul === 'pulita') ? 4 : 1;
        tiebreak = 0; break;

      case 'uscita':
        bucket = 2; tiebreak = 0; break;

      case 'controllare':
      case 'da-preparare':
        if (gArr === null) {
          bucket = 7; tiebreak = 0;
        } else if (gArr <= 1) {
          bucket = 3; tiebreak = gArr;
        } else if (gArr <= 7) {
          bucket = 5; tiebreak = gArr;
        } else {
          bucket = 6; tiebreak = gArr;
        }
        break;

      case 'occupata': {
        const _ivs = getRoomInterventions(room.id, today, room);
        const _hasLenz = _ivs.some(i => i.tipo === 'lenzuola');
        const _hasRass = _ivs.some(i => i.tipo === 'rassetto');
        // Cambio lenzuola = bucket 2b (dopo uscite, prima di da-preparare)
        // Rassetto = bucket 2c
        // Occupata senza interventi = bucket 8
        if (_hasLenz)      { bucket = 2; tiebreak = 1; }
        else if (_hasRass) { bucket = 2; tiebreak = 2; }
        else               { bucket = 8; tiebreak = st.occupante ? st.occupante.e.getTime() : 0; }
        break;
      }

      case 'pronta':
        bucket = 9; tiebreak = gArr ?? Infinity; break;

      case 'fuori-servizio':
        bucket = 10; tiebreak = 0; break;

      default:
        bucket = 99; tiebreak = 0;
    }

    allCards.push({ room, st, bucket, tiebreak });
  });

  allCards.sort((a, b) => {
    if (a.bucket !== b.bucket) return a.bucket - b.bucket;
    if (a.tiebreak !== b.tiebreak) return a.tiebreak - b.tiebreak;
    return a.room.name.localeCompare(b.room.name, 'it', {numeric:true});
  });

  let gridHtml = `<div class="rdash-grid">`;
  allCards.forEach(({ room, st }) => {
    const opId = st.opId;
      const rs = roomStates[room.id] || {};
      const pulizia = rs.pulizia || 'da-pulire';
      const ts = rs.ts ? new Date(rs.ts).toLocaleString('it-IT',{day:'2-digit',month:'2-digit',hour:'2-digit',minute:'2-digit'}) : '';

      // Pallino pulizia: mostrato su tutte le card (prominente sugli arrivi)
      const dotSize   = (opId === 'arrivo') ? '11px' : '9px';
      const dotCfg    = { 'pulita':['#2d6a4f','Pulita'], 'da-pulire':['#e67e22','Da pulire'], 'in-corso':['#8e44ad','Controllare/Rassettare'], 'fuori-servizio':['#c0392b','Fuori servizio'] };
      const [dotColor, dotLabel] = dotCfg[pulizia] || dotCfg['da-pulire'];
      const puliziaDotHtml = `<span style="display:inline-flex;align-items:center;gap:4px;font-size:10px;color:${dotColor};font-weight:600;">
        <span style="width:${dotSize};height:${dotSize};border-radius:50%;background:${dotColor};display:inline-block;flex-shrink:0;"></span>
        ${opId === 'arrivo' ? dotLabel : ''}
      </span>`;

      let detailHtml = '';

      if (st.opId === 'cambio') {
        detailHtml += `<div class="rdash-section-title">In uscita</div>
          <div class="rdash-guest-row">
            <span class="rdash-guest-name">↑ ${st.uscente.n}</span>
            ${st.uscente.d ? `<span class="rdash-disp">${st.uscente.d}</span>` : ''}
          </div>
          <div style="text-align:center;margin:4px 0;font-size:18px;color:var(--accent);">↓</div>
          <div class="rdash-section-title">In arrivo</div>
          <div class="rdash-guest-row">
            <span class="rdash-guest-name">↓ ${st.entrante.n}</span>
            ${st.entrante.d ? `<span class="rdash-disp">${st.entrante.d}</span>` : ''}
          </div>`;
        if (st.uscente.d !== st.entrante.d) {
          detailHtml += `<div style="font-size:10px;color:#e67e22;margin-top:5px;font-weight:600;">⚙ ${st.uscente.d||'—'} → ${st.entrante.d||'—'}</div>`;
        }
      } else if (st.opId === 'occupata' && st.occupante) {
        const giorniRimasti = Math.round((st.occupante.e - today)/(86400000));
        const interventions = getRoomInterventions(room.id, today, room);
        // Badge interventi (rassetto e/o cambio lenzuola)
        const badgesHtml = interventions.map(iv =>
          `<span class="badge-${iv.tipo}">${iv.tipo === 'lenzuola' ? '🛏 ' : '🧹 '}${iv.label}</span>`
        ).join('');
        detailHtml += `<div class="rdash-guest-row">
          <span class="rdash-guest-name">${st.occupante.n}</span>
          ${st.occupante.d ? `<span class="rdash-disp">${st.occupante.d}</span>` : ''}
        </div>
        <div style="font-size:10px;color:var(--text3);margin-top:3px;">
          Uscita: ${fmt(st.occupante.e)} · ancora ${giorniRimasti} ${giorniRimasti===1?'notte':'notti'}
        </div>
        ${badgesHtml ? `<div style="margin-top:6px;">${badgesHtml}</div>` : ''}`;
      } else if (st.opId === 'uscita' && st.uscente) {
        detailHtml += `<div class="rdash-guest-row">
          <span class="rdash-guest-name">${st.uscente.n}</span>
          ${st.uscente.d ? `<span class="rdash-disp">${st.uscente.d}</span>` : ''}
        </div>
        <div style="font-size:10px;color:var(--text3);margin-top:3px;">Da pulire dopo uscita</div>`;
      } else if (st.opId === 'arrivo' && st.entrante) {
        detailHtml += `<div class="rdash-guest-row" style="margin-top:2px;">
          <span class="rdash-guest-name">${st.entrante.n}</span>
          ${st.entrante.d ? `<span class="rdash-disp">${st.entrante.d}</span>` : ''}
        </div>`;
        if (rs.configurazione) {
          const dispOk = rs.configurazione.replace(/\s/g,'').toLowerCase() === (st.entrante.d||'').replace(/\s/g,'').toLowerCase();
          detailHtml += `<div style="font-size:10px;margin-top:3px;color:${dispOk?'#2d6a4f':'#e67e22'};">
            ${dispOk ? '✓ Configurazione ok' : '⚙ ' + rs.configurazione + ' → ' + (st.entrante.d||'—')}
          </div>`;
        }
      } else if (st.opId === 'pronta') {
        detailHtml += `<div style="font-size:10px;color:#2d6a4f;margin-top:2px;">✓ Camera pronta</div>`;
        if (rs.configurazione) detailHtml += `<div style="font-size:10px;color:var(--text2);">🛏 ${rs.configurazione}</div>`;
      } else if (st.opId === 'controllare' || st.opId === 'da-preparare') {
        if (st.prossimoArrivo) {
          const giorniAl = Math.round((st.prossimoArrivo.s - today) / 86400000);
          const urgenza = giorniAl <= 0 ? 'color:#c0392b;font-weight:700' : giorniAl === 1 ? 'color:#e67e22;font-weight:600' : 'color:var(--text3)';
          const quando = giorniAl <= 0 ? '⚠ OGGI' : giorniAl === 1 ? '⚠ domani' : fmt(st.prossimoArrivo.s) + ' (tra ' + giorniAl + 'gg)';
          detailHtml += `<div class="rdash-guest-row" style="margin-top:4px;">
            <span class="rdash-guest-name" style="font-size:11px;">${st.prossimoArrivo.n}</span>
            ${st.prossimoArrivo.d ? '<span class="rdash-disp">' + st.prossimoArrivo.d + '</span>' : ''}
          </div>
          <div style="font-size:10px;margin-top:2px;${urgenza}">Arrivo: ${quando}</div>`;
          if (rs.configurazione && st.prossimoArrivo.d) {
            const dispOk = rs.configurazione.replace(/\s/g,'').toLowerCase() === st.prossimoArrivo.d.replace(/\s/g,'').toLowerCase();
            if (!dispOk) detailHtml += `<div style="font-size:10px;color:#e67e22;margin-top:2px;">⚙ ${rs.configurazione} → ${st.prossimoArrivo.d}</div>`;
          }
        } else {
          detailHtml += `<div style="font-size:10px;color:var(--text3);margin-top:2px;">Nessun arrivo in programma</div>`;
          if (rs.configurazione) detailHtml += `<div style="font-size:10px;color:var(--text2);">🛏 ${rs.configurazione}</div>`;
        }
      }

      if (rs.noteOps) detailHtml += `<div class="rdash-note" style="margin-top:5px;">📌 ${rs.noteOps}</div>`;

      gridHtml += `
        <div class="rdash-card op-${st.opId}" onclick="openRstateModal('${room.id}')">
          <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:4px;">
            <div class="rdash-cam-name">${room.name}</div>
            <div style="display:flex;align-items:center;gap:6px;">
              ${puliziaDotHtml}
              <div class="rdash-cam-group" style="margin-bottom:0;">${room.g}</div>
            </div>
          </div>
          <div style="margin-bottom:6px;">${opBadge(st.opId, st.label)}</div>
          ${detailHtml}
          <div class="rdash-ts">${ts ? 'Agg: '+ts : ''}</div>
        </div>`;
    });
  gridHtml += `</div>`;

  const body = document.getElementById('roomDashPage');
  const existingBar = body.querySelector('.spage-bar');
  body.innerHTML = '';
  if (existingBar) body.appendChild(existingBar);
  body.insertAdjacentHTML('beforeend', filtersHtml + gridHtml);
}

// ── Modal modifica stato singola camera ──
function openRstateModal(roomId) {
  _rstateEditRoom = roomId;
  const room = ROOMS.find(r => r.id === roomId);
  const s = roomStates[roomId] || {};
  document.getElementById('rstateCamTitle').textContent = `Camera ${room?.name || roomId}`;
  document.getElementById('rstateConfig').value = s.configurazione || '';
  document.getElementById('rstateNote').value   = s.noteOps || '';
  // Costruisci bottoni stati
  const btns = document.getElementById('rstateButtons');
  btns.dataset.selected = '';  // reset selezione precedente
  btns.innerHTML = PULIZIA_STATI.map(st =>
    `<button class="rstate-btn ${st.cls}${(s.pulizia||'da-pulire')===st.id?' sel':''}"
             data-state-id="${st.id}"
             onclick="selectRstate('${st.id}')">${st.label}</button>`
  ).join('');
  document.getElementById('rstateOverlay').classList.add('open');
}

function selectRstate(statoId) {
  document.querySelectorAll('.rstate-btn').forEach(b => b.classList.remove('sel'));
  const btns = document.getElementById('rstateButtons');
  const idx = PULIZIA_STATI.findIndex(s => s.id === statoId);
  if (idx >= 0) btns.children[idx]?.classList.add('sel');
  // Salva selezione temporanea nel dataset
  btns.dataset.selected = statoId;
}

function closeRstateModal() {
  document.getElementById('rstateOverlay').classList.remove('open');
  _rstateEditRoom = null;
}

function rstateOverlayClick(e) {
  if (e.target === document.getElementById('rstateOverlay')) closeRstateModal();
}

async function saveRoomState() {
  if (!_rstateEditRoom) return;
  const btns = document.getElementById('rstateButtons');

  // Leggi l'id dal dataset (impostato da selectRstate) oppure dal bottone .sel
  // MAI usare textContent — quello è il label visivo, non l'id
  const selBtn = btns.querySelector('.rstate-btn.sel');
  const pulizia = btns.dataset.selected          // cliccato durante questa sessione
    || selBtn?.dataset?.stateId                  // id nel dataset del bottone (v. sotto)
    || (roomStates[_rstateEditRoom]?.pulizia || 'da-pulire');

  const configurazione = document.getElementById('rstateConfig').value.trim();
  const noteOps        = document.getElementById('rstateNote').value.trim();
  const roomId         = _rstateEditRoom;        // salva prima di closeRstateModal

  closeRstateModal();
  showLoading('Salvataggio stato camera…');
  try {
    await updateSingleRoomState(roomId, { pulizia, configurazione, noteOps });
    hideLoading();
    showToast('✓ Stato camera aggiornato', 'success');
    render();
    if (document.getElementById('roomDashPage').classList.contains('open')) renderRoomDash();
  } catch(e) {
    hideLoading();
    showToast('Errore salvataggio: ' + e.message, 'error');
  }
}


// ═══════════════════════════════════════════════════════════════════
// MODULO CHECK-IN
// ═══════════════════════════════════════════════════════════════════


function renderAnnualSheetRow(container, entry, idx) {
  const row = document.createElement('div');
  row.id = 'asr-' + idx;
  row.style.cssText = 'display:grid;grid-template-columns:80px 1fr auto;gap:8px;align-items:center;margin-bottom:8px;';
  row.innerHTML = `
    <input type="number" value="${entry.year}" min="2020" max="2040"
           style="background:var(--surface2);border:1.5px solid var(--border);color:var(--text);
                  font-family:inherit;font-size:12px;padding:7px 9px;border-radius:var(--radius);outline:none;width:100%;"
           id="asr-year-${idx}" placeholder="Anno">
    <input type="text" value="${entry.sheetId}"
           style="background:var(--surface2);border:1.5px solid var(--border);color:var(--text);
                  font-family:monospace;font-size:11px;padding:7px 9px;border-radius:var(--radius);outline:none;width:100%;"
           id="asr-id-${idx}" placeholder="ID Foglio Google">
    <button class="btn btn-icon" style="color:var(--danger);border-color:var(--danger);" onclick="removeAnnualSheetRow(${idx})">✕</button>`;
  container.appendChild(row);
}

function addAnnualSheetRow() {
  const list = document.getElementById('annualSheetsList');
  if (!list) return;
  const currentCount = list.children.length;
  const nextYear = new Date().getFullYear() + 1;
  renderAnnualSheetRow(list, { year: nextYear, sheetId: '', label: String(nextYear) }, currentCount);
}

function removeAnnualSheetRow(idx) {
  const row = document.getElementById('asr-' + idx);
  if (row) row.remove();
  // Rinumera i rimanenti
  const list = document.getElementById('annualSheetsList');
  if (!list) return;
  Array.from(list.children).forEach((el, i) => { el.id = 'asr-' + i; });
}

async function testDbConnection() {
  const idEl = document.getElementById('sDbSheetId');
  const res  = document.getElementById('dbTestResult');
  const id   = idEl?.value.trim();
  if (!id) { res.textContent = "⚠ Inserisci l'ID del foglio"; res.style.color='var(--danger)'; return; }
  res.textContent = '⏳ Connessione…'; res.style.color='var(--text3)';
  try {
    const r = await fetch(
      `https://sheets.googleapis.com/v4/spreadsheets/${id}?fields=properties.title,sheets.properties.title`,
      { headers: { Authorization: 'Bearer ' + accessToken } }
    );
    if (!r.ok) throw new Error('HTTP ' + r.status);
    const d = await r.json();
    const sheetNames = d.sheets?.map(s => s.properties.title) || [];
    const hasPren = sheetNames.includes('PRENOTAZIONI');
    if (hasPren) {
      res.textContent = `✓ Connesso: "${d.properties.title}" — scheda PRENOTAZIONI trovata`;
      res.style.color = 'var(--accent)';
    } else {
      res.textContent = `⚠ Foglio trovato ma manca la scheda "PRENOTAZIONI". Creala nel foglio Google.`;
      res.style.color = 'var(--danger)';
    }
  } catch(e) {
    res.textContent = '✗ Errore: ' + e.message + ' — controlla l\'ID e i permessi';
    res.style.color = 'var(--danger)';
  }
}

