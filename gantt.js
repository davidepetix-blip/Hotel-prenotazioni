// ═══════════════════════════════════════════════════════════════════
// gantt.js — Render Gantt, booking drawer/modal, room dashboard, settings
// Blip Hotel Management — build 18.7.x
// Dipende da: core.js, sync.js
// ═══════════════════════════════════════════════════════════════════


const BLIP_VER_GANTT = '28'; // ← incrementa ad ogni modifica

let _billingPreloaded = false;
function render() {
  // Al primo render con DATABASE_SHEET_ID disponibile, precarica i dati di conto.
  // preloadContoDati() è asincrona: chiama render() di nuovo al completamento.
  if (!_billingPreloaded && typeof preloadContoDati === 'function' && typeof DATABASE_SHEET_ID !== 'undefined' && DATABASE_SHEET_ID) {
    _billingPreloaded = true;
    preloadContoDati(); // async — ri-renderizza da sola al completamento
  }
  const days = dim(curY, curM);
  const now  = new Date();
  const isNow = now.getFullYear()===curY && now.getMonth()===curM;
  const tod  = now.getDate();
  const CW   = 34;
  document.getElementById('mlabel').textContent = `${MONTHS_S[curM]} ${curY}`;

  let h = `<div class="legend">
    <span class="leg"><span style="display:inline-block;width:9px;height:9px;border-radius:2px;border-left:3px solid #34a853;background:var(--surface2);margin-right:4px;"></span>Pagato</span>
    <span class="leg"><span style="display:inline-block;width:9px;height:9px;border-radius:2px;border-left:3px solid #4285f4;background:var(--surface2);margin-right:4px;"></span>Fatturato</span>
    <span class="leg"><span style="display:inline-block;width:9px;height:9px;border-radius:2px;border-left:3px solid #fa7b17;background:var(--surface2);margin-right:4px;"></span>Emesso</span>
    <span class="leg"><span style="display:inline-block;width:9px;height:9px;border-radius:2px;border-left:3px solid #9e9e9e;background:var(--surface2);margin-right:4px;"></span>Bozza</span>
    <span class="lsep"></span>
    <span class="leg"><span style="display:inline-block;width:2px;height:9px;background:var(--accent);margin-right:4px;"></span>Oggi</span>
    <span class="leg"><span style="display:inline-block;width:9px;height:9px;border-radius:2px;outline:2px dashed #ff6b6b;margin-right:4px;"></span>Adiacente</span>
    <span class="leg"><span style="display:inline-block;width:14px;height:9px;border-radius:2px;background:var(--text3);opacity:.25;margin-right:1px;vertical-align:middle"></span><span style="display:inline-block;width:6px;height:9px;margin-right:4px;vertical-align:middle"></span>Checkout (parziale)</span>
  </div>`;

  h+=`<div class="dheader">`;
  for(let d=1;d<=days;d++){
    const dow=new Date(curY,curM,d).getDay();
    const cls=isNow&&d===tod?'today':dow===0?'sunday':'';
    h+=`<div class="dhcell ${cls}"><div class="dhnum">${d}</div><div class="dhdow">${DAYS_IT[dow]}</div></div>`;
  }
  h+=`</div>`;

  const ms=new Date(curY,curM,1), me=new Date(curY,curM+1,0);
  const groups=[...new Set(ROOMS.map(r=>r.g))];

  groups.forEach(g=>{
    const gr=ROOMS.filter(r=>r.g===g);
    let sc='';
    for(let d=1;d<=days;d++){ const dow=new Date(curY,curM,d).getDay(); sc+=`<div class="seccell${dow===0?' sunday':''}"></div>`; }
    h+=`<div class="secrow"><div class="seclabel">${g}</div><div class="seccells">${sc}</div></div>`;

    gr.forEach(room=>{
      const rb=bookings.filter(b=>b.r===room.id&&b.s<=me&&b.e>=ms);
      let cells='';
      for(let d=1;d<=days;d++){
        const dow=new Date(curY,curM,d).getDay();
        const tc=isNow&&d===tod?'todaycol':dow===0?'sunday':'';
        cells+=`<div class="dcell ${tc}" onclick="cellClick('${room.id}',${d})"></div>`;
      }
      let bars='';
      rb.forEach(b=>{
        const vs=b.s<ms?ms:b.s, ve=b.e>me?me:b.e;
        const sd=vs.getDate(), ed=ve.getDate();
        // 'continues' = prenotazione continua nel mese successivo → larghezza piena
        const continues = b.e > me;
        // Giorno di checkout: la barra occupa solo il 40% della cella —
        // così la camera risulta visivamente libera per il pomeriggio
        // e un check-in nella stessa giornata non si sovrappone.
        const checkoutFrac = 0.40;
        const w = continues
          ? (ed-sd+1)*CW - 2                          // continua: piena
          : (ed-sd)*CW + Math.round(CW*checkoutFrac); // ultimo giorno: 40%
        const lx=(sd-1)*CW+1;
        const tc = '#1a1916'; // testo sempre scuro — tutti i colori della palette sono pastello
        const adj=adjConflict(b).length>0;
        const _billBorder = (typeof billingBorderColor === 'function') ? billingBorderColor(b.dbId || b.id) : null;
        const _borderStyle = _billBorder ? `border-left:3px solid ${_billBorder};` : '';
        bars+=`<div class="bbar${adj?' adj':''}${b.pending?' pending':''}${continues?' continues':''}"
          style="left:${lx}px;width:${w}px;background:${b.c};color:${tc};${_borderStyle}"
          onclick="selBook(${b.id},event)"
          data-bid="${b.id}"
          onmouseenter="if(!('ontouchstart' in window))showTT(event,${b.id})"
          onmouseleave="hideTT()">
          ${b.n}<span class="bdisp">${b.d}</span></div>`;
      });
      let tv='';
      if(isNow){ const tx=(tod-1)*CW+CW/2-1; tv=`<div class="tvline" style="left:${tx}px"></div>`; }
      const _today=new Date(); _today.setHours(12,0,0,0);
      const _opst=getRoomDayStatus(room.id,_today);
      const _opLabels={'cambio':'Cambio','occupata':'Occupata','uscita':'Uscita oggi','arrivo':'Arrivo oggi','pronta':'Pronta','da-preparare':'Da preparare','controllare':'Controllare/Rassettare','fuori-servizio':'Fuori servizio'};
      const _dot=`<span class="room-status-dot rsd-op-${_opst.opId}" title="${_opLabels[_opst.opId]||''}"></span>`;
      h+=`<div class="rrow"><div class="rlabel" onclick="event.stopPropagation();openRoomDrawer('${room.id}')" style="cursor:pointer;">${room.name}${_dot}</div><div class="dcwrap">${cells}${bars}${tv}</div></div>`;
    });
  });

  document.getElementById('ginner').innerHTML = h;
  updateStats();
}

function updateStats(){
  const ms=new Date(curY,curM,1),me=new Date(curY,curM+1,0);
  const act=bookings.filter(b=>b.s<=me&&b.e>=ms);
  const occ=new Set(act.map(b=>b.r)).size;
  document.getElementById('sttot').textContent=act.length;
  document.getElementById('stocc').textContent=occ;
  document.getElementById('stfre').textContent=ROOMS.length-occ;
}

function changeMonth(d, animate=false){
  curM += d;
  if(curM > 11){ curM = 0; curY++; }
  if(curM <  0){ curM = 11; curY--; }
  render();
  // Breve fade-in sul nuovo contenuto
  if(animate){
    const gi = document.getElementById('ginner');
    if(gi){
      gi.style.opacity = '0';
      gi.style.transform = d > 0 ? 'translateX(32px)' : 'translateX(-32px)';
      requestAnimationFrame(() => {
        gi.style.transition = 'opacity .2s ease, transform .2s ease';
        gi.style.opacity = '1';
        gi.style.transform = 'translateX(0)';
        setTimeout(() => { gi.style.transition = ''; }, 220);
      });
    }
  }
}
function goToday(){ const t=new Date(); curM=t.getMonth(); curY=t.getFullYear(); render(); setTimeout(()=>{ document.getElementById('gscroll').scrollLeft=Math.max(0,(t.getDate()-4)*34+90); },60); }

// ── SWIPE orizzontale per cambiare mese ──
// Logica: se l'utente ha scrollato fino al bordo E continua a trascinare
// nella stessa direzione oltre una soglia → cambia mese.
(function initSwipeMonth() {
  let tx0=0, ty0=0, sx0=0, _locked=false, _dir=0;
  const THRESH_PX  = 55;   // px di overscroll per triggerare
  const MIN_HORIZ  = 28;   // movimento minimo orizzontale
  const MAX_VERT   = 40;   // massimo verticale (se troppo verticale = scroll normale)

  function gs() { return document.getElementById('gscroll'); }

  document.addEventListener('touchstart', e => {
    const el = gs(); if (!el || !el.contains(e.target)) return;
    tx0 = e.touches[0].clientX;
    ty0 = e.touches[0].clientY;
    sx0 = el.scrollLeft;
    _locked = false; _dir = 0;
  }, { passive: true });

  document.addEventListener('touchmove', e => {
    if (_locked) return;
    const el = gs(); if (!el) return;
    const dx = e.touches[0].clientX - tx0;
    const dy = Math.abs(e.touches[0].clientY - ty0);
    if (dy > MAX_VERT) { _locked = true; return; }
    if (Math.abs(dx) < MIN_HORIZ) return;

    const atLeft  = el.scrollLeft <= 0;
    const atRight = el.scrollLeft >= el.scrollWidth - el.clientWidth - 2;

    if (dx > THRESH_PX && atLeft && sx0 <= 0) {
      _locked = true; _dir = -1;
    } else if (dx < -THRESH_PX && atRight) {
      _locked = true; _dir = 1;
    }
  }, { passive: true });

  document.addEventListener('touchend', () => {
    if (_dir !== 0) {
      changeMonth(_dir, true);
      _dir = 0;
    }
    _locked = false;
  });
})();

// ═══════════════════════════════════════════════════════════════════
// INTERAZIONI
// ═══════════════════════════════════════════════════════════════════
function selBook(id,e){
  e&&e.stopPropagation();
  hideTT(); // forza chiusura tooltip prima di qualsiasi altra operazione
  const b=bookings.find(x=>x.id===id); if(!b) return;
  const adj=adjConflict(b);
  let adjHtml='';
  // Distingui tra adiacenza con stesso nome+dispo diversa (cambio legittimo)
  // e adiacenza con nome/dispo uguale (potenziale fusione indesiderata)
  const adjCambio = bookings.filter(o =>
    o.id !== b.id && o.r === b.r && o.c === b.c &&
    (o.e.getTime()===b.s.getTime()||o.s.getTime()===b.e.getTime()) &&
    o.n === b.n && o.d !== b.d
  );
  if(adj.length>0) adjHtml=`<div class="adjwarn">⚠ Adiacente a "<b>${adj[0].n}</b>" con stesso colore <b>${b.c}</b>. Lo script li unirebbe in un'unica prenotazione. Modifica il colore di una.</div>`;
  if(adjCambio.length>0) adjHtml=`<div style="font-size:11px;color:var(--accent);background:rgba(45,106,79,.08);border-radius:6px;padding:8px 10px;margin-bottom:8px">
    🔄 Cambio disposizione: <b>${adjCambio[0].d}</b> → <b>${b.d}</b> (${b.n})<br>
    <span style="color:var(--text3)">Le due prenotazioni sono collegate — usa il tab Gruppo per il conto completo.</span>
  </div>`;
  document.getElementById('drtitle').textContent=b.n;
  document.getElementById('drsub').textContent=`CAMERA ${roomName(b.r)} · ${roomGroup(b.r)}`;
  document.getElementById('drbody').innerHTML=`
    ${adjHtml}
    <div class="dr-bill-tabs">
      <div class="dr-bill-tab active" onclick="drTab(this,'drTabInfo')">📋 Dettagli</div>
      <div class="dr-bill-tab" onclick="drTab(this,'drTabBill')">💶 Conto</div>
      <div class="dr-bill-tab" onclick="drTabCheckin(this,${b.id})">🛎 Check-in</div>
    </div>
    <div id="drTabInfo">
    <div class="dcard">
      <div class="dcname"><span class="cpill" style="background:${b.c}"></span>${b.n}</div>
      <div class="drow"><span class="dkey">Camera</span><span class="dval">${roomName(b.r)}</span></div>
      <div class="drow"><span class="dkey">Gruppo</span><span class="dval">${roomGroup(b.r)}</span></div>
      <div class="drow"><span class="dkey">Check-in</span><span class="dval">${fmt(b.s)}</span></div>
      <div class="drow"><span class="dkey">Check-out</span><span class="dval">${fmt(b.e)}</span></div>
      <div class="drow"><span class="dkey">Notti</span><span class="dval">${nights(b.s,b.e)}</span></div>
      <div class="drow"><span class="dkey">Disposizione</span><span class="dval">${b.d}</span></div>
      ${b.note?`<div class="drow"><span class="dkey">Note</span><span class="dval">${b.note}</span></div>`:''}
      ${b.fromSheet?`<div class="drow"><span class="dkey">Fonte</span><span class="dval" style="color:var(--success)">📋 Google Sheets</span></div>`:''}
    </div>
    <div style="display:flex;gap:8px;">
      <button class="btn" onclick="editBook(${b.id})" style="flex:1;justify-content:center;">✎ Modifica</button>
      <button class="btn danger" onclick="delBook(${b.id})" style="flex:1;justify-content:center;">✕ Elimina</button>
    </div>
    <div class="synchint">
      ${(()=>{
        const frags = splitBookingByMonth(b);
        if(frags.length <= 1) {
          return `<b>→ Google Sheets</b><br>
          Colonna "${roomName(b.r)}" · Foglio "${b.sheetName || sheetName(b.s.getFullYear(),b.s.getMonth())}"<br>
          Celle ${fmt(b.s)} → ${fmt(b.e)} · Colore <span style="font-size:9px;padding:1px 5px;border-radius:2px;background:${b.c};color:${light(b.c)?'#222':'#eee'}">${b.c}</span>`;
        } else {
          return `<b>→ Google Sheets (prenotazione multi-mese)</b><br>
          Scritta su <b>${frags.length} fogli</b>: ${frags.map(f=>f.sName).join(', ')}<br>
          Camera "${roomName(b.r)}" · Colore <span style="font-size:9px;padding:1px 5px;border-radius:2px;background:${b.c};color:${light(b.c)?'#222':'#eee'}">${b.c}</span>`;
        }
      })()}
    </div>
    </div>
    <div id="drTabBill" style="display:none;">
      ${(()=>{ try { return renderDrawerBill(b); } catch(err) { syncLog('⚠ Conto: '+err.message,'wrn'); return '<div style="padding:12px;font-size:12px;color:var(--danger)">⚠ Errore caricamento conto: '+err.message+'</div>'; } })()}
    </div>
    <div id="drTabCI" style="display:none;" data-booking-id="${b.id}">
      <div style="padding:20px;text-align:center;color:var(--text3);font-size:12px;">Caricamento…</div>
    </div>`;
  openDrawer();
  setTimeout(() => { if (typeof renderCheckinDrawerTab === 'function') renderCheckinDrawerTab(b.id); }, 50);
}

function drTabCheckin(el, bookingNumId) {
  el.closest('#drbody').querySelectorAll('.dr-bill-tab').forEach(t=>t.classList.remove('active'));
  el.classList.add('active');
  ['drTabInfo','drTabBill','drTabCI'].forEach(id=>{
    const d=document.getElementById(id);
    if(d) d.style.display = id==='drTabCI'?'':'none';
  });
  if (typeof renderCheckinDrawerTab === 'function') renderCheckinDrawerTab(bookingNumId);
}

function renderCheckinDrawerTab(bookingNumId, _reloaded=false) {
  const tabEl = document.getElementById('drTabCI');
  if (!tabEl) return;
  const b = bookings.find(x => x.id === bookingNumId);
  if (!b) { tabEl.innerHTML = '<div style="padding:16px;color:var(--text3);font-size:12px">Prenotazione non trovata</div>'; return; }
  // Se ciData sembra vuota o non trova la prenotazione, forza ricaricamento
  const _ciKeys = (typeof ciData !== 'undefined') ? Object.keys(ciData).length : 0;
  if (typeof syncLog === 'function') syncLog('CI tab: id='+bookingNumId+' dbId='+(b.dbId||'null')+' ciData.keys='+_ciKeys, 'syn');
  let ci = (typeof getCiForBooking === 'function') ? getCiForBooking(b) : null;
  // Se ciData è vuoto e il DB è configurato → carica prima di renderizzare (una sola volta)
  if (!ci && _ciKeys === 0 && !_reloaded && typeof loadCiData === 'function' && typeof DATABASE_SHEET_ID !== 'undefined' && DATABASE_SHEET_ID) {
    tabEl.innerHTML = '<div style="padding:16px;color:var(--text3);font-size:12px">⏳ Caricamento check-in…</div>';
    loadCiData(false).then(() => renderCheckinDrawerTab(bookingNumId, true));
    return;
  }
  // Se ciData ha dati ma non trova il check-in → forza reload una sola volta
  // GUARD: senza _reloaded questo diventava un loop infinito per prenotazioni senza check-in
  if (!ci && _ciKeys > 0 && b.dbId && !_reloaded && typeof loadCiData === 'function') {
    loadCiData(true).then(() => renderCheckinDrawerTab(bookingNumId, true));
    return;
  }
  if (typeof syncLog === 'function') syncLog('CI trovato: '+(ci?'SI cam='+ci.camera+' data='+ci.data:'NO'), ci?'ok':'wrn');
  const oggi = new Date().toISOString().slice(0,10);
  const checkinDate = b.s.toISOString().slice(0,10);
  const isLate = checkinDate < oggi;
  if (ci) {
    const capo = ci.guests && ci.guests[0] ? ci.guests[0] : {};
    tabEl.innerHTML =
      '<div style="padding:14px">' +
      '<div style="background:#d1e7dd;border-radius:8px;padding:10px 12px;margin-bottom:12px;font-size:12px;color:#0f5132">' +
      '✓ Check-in ' + ci.data + ' · ' + ci.numOspiti + ' ospite/i</div>' +
      '<div style="font-size:13px;font-weight:600;margin-bottom:4px">' + escHtml(capo.cognome||'') + ' ' + escHtml(capo.nome||'') + '</div>' +
      '<div style="font-size:11px;color:var(--text3);margin-bottom:6px">' + escHtml(capo.tipoDoc||'') + ' ' + escHtml(capo.numDoc||'') + '</div>' +
      (ci.guests.length > 1 ? '<div style="font-size:11px;color:var(--text3);margin-bottom:8px">' + (ci.guests.length-1) + ' accompagnatore/i</div>' : '') +
      '<div style="display:flex;gap:8px">' +
      '<button class="btn primary" onclick="openCiModal(\'' + b.dbId + '\')" style="flex:1;justify-content:center">✎ Modifica</button>' +
      '<button class="btn" onclick="exportAlloggiati(\'all\')" style="flex:1;justify-content:center">⬇ .txt</button>' +
      '</div></div>';
  } else if (checkinDate <= oggi) {
    const giorni = Math.round((Date.now() - b.s.getTime()) / 86400000);
    tabEl.innerHTML =
      '<div style="padding:14px">' +
      (isLate ? '<div style="background:#f8d7da;border-radius:8px;padding:10px 12px;margin-bottom:12px;font-size:12px;color:#842029">' +
        '⚠ Check-in in ritardo<br><span style="font-size:11px">Arrivo ' + giorni + ' giorno/i fa</span></div>' : '') +
      '<button class="btn primary" onclick="openCiModal(\'' + b.dbId + '\')" style="width:100%;justify-content:center;margin-top:4px">' +
      '🛎 Registra check-in ora</button></div>';
  } else {
    const giorni = Math.ceil((b.s.getTime() - Date.now()) / 86400000);
    tabEl.innerHTML =
      '<div style="padding:14px;text-align:center;color:var(--text3);font-size:12px">' +
      '📅 Arrivo tra ' + giorni + ' giorno/i<br>' +
      '<button class="btn" onclick="openCiModal(\'' + b.dbId + '\')" style="margin-top:10px">' +
      '🛎 Pre-registra check-in</button></div>';
  }
}

function editBook(id){
  const b=bookings.find(x=>x.id===id); if(!b) return;
  editId=id;
  document.getElementById('fRoom').value=b.r;
  document.getElementById('fName').value=b.n;
  document.getElementById('fIn').value=b.s.toISOString().slice(0,10);
  document.getElementById('fOut').value=b.e.toISOString().slice(0,10);
  document.getElementById('fNotes').value=b.note;
  selColor=b.c;
  // Parsa la stringa disposizione nei contatori
  bedCounts = parseBedString(b.d);
  closeDrawer(); openModal(true);
}

async function delBook(id){
  const b=bookings.find(x=>x.id===id); if(!b) return;
  if(!confirm(`Eliminare la prenotazione di "${b.n}"?\nQuesto rimuoverà anche i colori dal foglio Google.`)) return;
  closeDrawer();
  showLoading('Rimozione…');
  try {
    await bridgeCancella(b);
    if(DATABASE_SHEET_ID && b.dbRow) await dbDelete(b, 'Eliminata dall\'app');
    bookings=bookings.filter(x=>x.id!==id);
    hideLoading(); render();
    showToast('Prenotazione eliminata', 'success');
  } catch(e) {
    hideLoading();
    showToast('Errore eliminazione: ' + e.message, 'error');
  }
}

function cellClick(rid,day){
  // Costruiamo la stringa YYYY-MM-DD in locale per evitare lo shift UTC
  // (toISOString() con fuso +1/+2 restituirebbe il giorno precedente)
  const pad = n => String(n).padStart(2,'0');
  document.getElementById('fRoom').value = rid;
  document.getElementById('fIn').value  = curY+'-'+pad(curM+1)+'-'+pad(day);
  document.getElementById('fOut').value = curY+'-'+pad(curM+1)+'-'+pad(day+1);
  openModal();
}

// ═══════════════════════════════════════════════════════════════════
// DRAWER
// ═══════════════════════════════════════════════════════════════════
function openDrawer(){
  // Nascondi tooltip con forza — su Android il tap può triggerare mouseenter dopo click
  const tt = document.getElementById('tt');
  if (tt) { tt.style.display = 'none'; tt.style.opacity = '0'; }
  document.getElementById('drawer').classList.add('open');
  document.getElementById('dov').classList.add('open');
  document.body.style.overflow = 'hidden';
  // Su mobile: dopo 100ms nascondi ancora il tooltip (race condition Android)
  if ('ontouchstart' in window) setTimeout(() => { if(tt) tt.style.display='none'; }, 100);
}

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
function closeDrawer(){
  document.getElementById('drawer').classList.remove('open');
  document.getElementById('dov').classList.remove('open');
  document.body.style.overflow = '';
}

// ═══════════════════════════════════════════════════════════════════
// MODAL
// ═══════════════════════════════════════════════════════════════════
function openModal(isEdit=false){
  if(!isEdit){
    editId=null;
    document.getElementById('fName').value='';
    document.getElementById('fNotes').value='';
    if(!document.getElementById('fIn').value){
      const t=new Date();
      document.getElementById('fIn').value=t.toISOString().slice(0,10);
      const t2=new Date(t); t2.setDate(t2.getDate()+1);
      document.getElementById('fOut').value=t2.toISOString().slice(0,10);
    }
    bedCounts={m:0,ms:0,s:0,c:0,aff:0};
    // Scegli automaticamente un colore alternato rispetto alle prenotazioni
    // già presenti nella camera selezionata
    const _rid = document.getElementById('fRoom')?.value;
    const _camBookings = _rid ? bookings.filter(b => b.r === _rid) : bookings;
    if (typeof nextBookingColor === 'function') {
      selColor = nextBookingColor(_camBookings);
    }
  }
  document.getElementById('mtitle').textContent=isEdit?'Modifica Prenotazione':'Nuova Prenotazione';

  // Gestione campo camera
  const _selRoom = document.getElementById('fRoom');
  const _cameraFg = _selRoom?.closest('.fg');
  const _cameraLbl = _cameraFg?.querySelector('label.fl');
  // Rimuovi eventuale div statico da sessione edit precedente
  _cameraFg?.querySelector('.camera-static-display')?.remove();
  if (isEdit) {
    // In modifica: mostra il nome camera come testo statico
    const _camName = window._editCameraName || _selRoom?.value || '—';
    if (_selRoom) _selRoom.style.display = 'none';
    const _staticDiv = document.createElement('div');
    _staticDiv.className = 'fi camera-static-display';
    _staticDiv.style.cssText = 'opacity:0.7;cursor:default;';
    _staticDiv.textContent = _camName;
    _selRoom?.after(_staticDiv);
    if (_cameraLbl) _cameraLbl.textContent = 'Camera';
  } else {
    // In nuova prenotazione: select visibile e abilitato
    if (_selRoom) { _selRoom.style.display = ''; _selRoom.disabled = false; _selRoom.style.opacity = ''; }
    if (_cameraLbl) _cameraLbl.textContent = 'Camera';
  }

  // Anagrafica: precarica clienti e preimposta se in edit mode
  if (typeof initAnagraficaModal === 'function') initAnagraficaModal();
  if (isEdit && editId) {
    const _eb = bookings.find(x=>x.id===editId);
    if (_eb?.clienteId && typeof preimpostaClienteModal === 'function') preimpostaClienteModal(_eb.clienteId);
    else if (typeof resetAnagraficaModal === 'function') resetAnagraficaModal();
  } else {
    if (typeof resetAnagraficaModal === 'function') resetAnagraficaModal();
  }
  rebuildBeds(isEdit ? bedCounts : null); rebuildColors();
  document.getElementById('errmsg').classList.remove('show');
  document.getElementById('mov').classList.add('open');
}
function closeModal(){
  document.getElementById('mov').classList.remove('open');
  editId=null;
  document.getElementById('fIn').value='';
  document.getElementById('fOut').value='';
  // Ripristina select camera
  const _s = document.getElementById('fRoom');
  if (_s) { _s.style.display = ''; _s.disabled = false; _s.style.opacity = ''; }
  document.getElementById('fRoom')?.closest('.fg')?.querySelector('.camera-static-display')?.remove();
  const _lbl = document.getElementById('fRoom')?.closest('.fg')?.querySelector('label.fl');
  if (_lbl) _lbl.textContent = 'Camera';
  window._editCameraName = null;
  if (typeof resetAnagraficaModal === 'function') resetAnagraficaModal();
}
function movClick(e){ if(e.target===document.getElementById('mov')) closeModal(); }

function rebuildBeds(initCounts) {
  const rid = document.getElementById('fRoom').value;
  const cfg = roomSettings[rid] || { maxGuests: 6, allowedBeds: ['m','s','c','aff'] };
  const el  = document.getElementById('bedCounters');
  el.innerHTML = '';

  // Se vengono passati contatori iniziali (modalità modifica), usali
  if (initCounts) {
    bedCounts = { ...initCounts };
  } else {
    // Reset contatori
    bedCounts = { m:0, ms:0, s:0, c:0, aff:0 };
  }

  BED_TYPES.forEach(({ id, label }) => {
    if (!cfg.allowedBeds.includes(id)) return; // non consentito per questa camera

    const row = document.createElement('div');
    row.className = 'bed-counter-row' + (bedCounts[id] > 0 ? ' active' : '');
    row.id = 'bcrow-' + id;

    const lbl = document.createElement('div');
    lbl.className = 'bed-counter-label';
    lbl.innerHTML = `${label} <small>(${id})</small>`;

    const ctrl = document.createElement('div');
    ctrl.className = 'bed-counter-ctrl';

    const btnMinus = document.createElement('button');
    btnMinus.className = 'bed-ctr-btn';
    btnMinus.textContent = '−';
    btnMinus.disabled = bedCounts[id] === 0;
    btnMinus.onclick = () => changeBedCount(id, -1, cfg.maxGuests);

    const valEl = document.createElement('div');
    valEl.className = 'bed-ctr-val';
    valEl.id = 'bcval-' + id;
    valEl.textContent = bedCounts[id];

    const btnPlus = document.createElement('button');
    btnPlus.className = 'bed-ctr-btn';
    btnPlus.textContent = '+';
    btnPlus.onclick = () => changeBedCount(id, +1, cfg.maxGuests);

    ctrl.append(btnMinus, valEl, btnPlus);
    row.append(lbl, ctrl);
    el.appendChild(row);
  });

  updateBedPreview();
}

function changeBedCount(id, delta, maxGuests) {
  const newVal = (bedCounts[id] || 0) + delta;
  if (newVal < 0) return;
  // Controlla capienza totale
  const total = Object.values(bedCounts).reduce((a,b) => a+b, 0) - (bedCounts[id]||0) + newVal;
  // Matrimoniale conta come 2 posti, singolo come 1, culla come 0 (non conta per capienza adulti)
  const guestCount = (bedCounts.m + (id==='m'?delta:0))*2 + (bedCounts.ms + (id==='ms'?delta:0))*1 + (bedCounts.s + (id==='s'?delta:0)) + (bedCounts.aff + (id==='aff'?delta:0));
  if (delta > 0 && guestCount > maxGuests * 2) return; // limite generoso

  bedCounts[id] = newVal;

  // Aggiorna UI
  const valEl = document.getElementById('bcval-' + id);
  if (valEl) valEl.textContent = newVal;

  const minusBtn = document.querySelector(`#bcrow-${id} .bed-ctr-btn`);
  if (minusBtn) minusBtn.disabled = newVal === 0;

  const row = document.getElementById('bcrow-' + id);
  if (row) row.classList.toggle('active', newVal > 0);

  updateBedPreview();
}

function updateBedPreview() {
  const str = buildBedString(bedCounts);
  const prev = document.getElementById('bedPreview');
  if (prev) prev.textContent = str !== 'ND' ? `→ ${str}` : '';
}
function rebuildColors(){
  const el=document.getElementById('colopts'); el.innerHTML='';
  const rid=document.getElementById('fRoom').value;
  const used=colorsUsed(rid,editId);
  PALETTE.forEach(({h,n})=>{
    const d=document.createElement('div');
    d.className='cdot'+(h===selColor?' sel':'')+(used.includes(h)?' inuse':'');
    d.style.background=h; d.title=n+(used.includes(h)?' (già usato — rischio adiacenza!)':'');
    d.onclick=()=>{ selColor=h; el.querySelectorAll('.cdot').forEach(x=>x.classList.remove('sel')); d.classList.add('sel'); };
    el.appendChild(d);
  });
}

// ═══════════════════════════════════════════════════════════════════
// SALVA PRENOTAZIONE → foglio Google
// ═══════════════════════════════════════════════════════════════════
async function saveBooking(){
  const rid  = document.getElementById('fRoom').value;
  const name = document.getElementById('fName').value.trim();
  const iv   = document.getElementById('fIn').value;
  const ov   = document.getElementById('fOut').value;
  const note = document.getElementById('fNotes').value.trim();
  const err  = document.getElementById('errmsg');
  const show = m => { err.textContent=m; err.classList.add('show'); };

  if(!rid)      { show('Seleziona una camera.'); return; }
  if(!name)     { show('Inserisci il nome del cliente.'); return; }
  if(!iv||!ov)  { show('Inserisci le date di check-in e check-out.'); return; }

  const sd=new Date(iv+'T12:00:00'), ed=new Date(ov+'T12:00:00');
  if(ed<=sd)    { show('Il check-out deve essere dopo il check-in.'); return; }

  // ── Info prenotazioni multi-mese ──
  // Nessun blocco: la scrittura è gestita automaticamente su più fogli mensili

  // ── OVERLAP BLOCK ──
  const conf=bookings.filter(b=>b.r===rid&&b.id!==editId&&b.s<ed&&b.e>sd);
  if(conf.length>0){
    // Caso speciale: stiamo modificando una prenotazione lunga (stessa camera, stesso nome)
    // e la "sovrapposizione" è con un frammento mensile della stessa prenotazione originale.
    const existingBCheck = editId ? bookings.find(b=>b.id===editId) : null;
    const isSelfOverlap  = existingBCheck &&
      conf.every(c => c.n === existingBCheck.n && c.r === existingBCheck.r && c.c === existingBCheck.c);
    if (!isSelfOverlap) {
      show(`⛔ Sovrapposizione con "${conf[0].n}" (${fmt(conf[0].s)} → ${fmt(conf[0].e)}). Impossibile salvare.`);
      return;
    }
    // isSelfOverlap → bridgeSalva gestisce il vecchio range automaticamente
  }

  // ── MODIFICA PRENOTAZIONE MULTI-MESE — conferma esplicita ──
  // Se stiamo modificando una prenotazione che copre più mesi mostriamo un riepilogo
  // di cosa verrà cancellato e cosa verrà scritto prima di procedere.
  if (editId) {
    const _eb4confirm = bookings.find(b=>b.id===editId);
    if (_eb4confirm) {
      const oldFrags   = splitBookingByMonth(_eb4confirm);
      const newFrags   = splitBookingByMonth({..._eb4confirm, s:sd, e:ed});
      const isMulti    = oldFrags.length > 1 || newFrags.length > 1;
      const dateChanged= Math.abs(_eb4confirm.s - sd) > 43200000 || Math.abs(_eb4confirm.e - ed) > 43200000;
      if (isMulti && dateChanged) {
        const oldFogli = [...new Set(oldFrags.map(f=>f.sName))].join(', ');
        const newFogli = [...new Set(newFrags.map(f=>f.sName))].join(', ');
        const msg =
          `📋 RIEPILOGO MODIFICA PRENOTAZIONE
` +
          `━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
` +
          `Cliente : ${_eb4confirm.n}
` +
          `Camera  : ${roomName(_eb4confirm.r)}

` +
          `🗑 VERRÀ CANCELLATO:
` +
          `  ${fmt(_eb4confirm.s)} → ${fmt(_eb4confirm.e)}
` +
          `  Fogli: ${oldFogli}

` +
          `✏ VERRÀ SCRITTO:
` +
          `  ${fmt(sd)} → ${fmt(ed)}
` +
          `  Fogli: ${newFogli}

` +
          `Confermi la modifica?`;
        if (!confirm(msg)) return;
      }
    }
  }

  // ── ADJACENT COLOR WARNING ──
  const adj=bookings.filter(b=>b.id!==editId&&b.r===rid&&b.c===selColor&&(b.e.getTime()===sd.getTime()||b.s.getTime()===ed.getTime()));
  if(adj.length>0){
    const ok=confirm(`⚠ "${adj[0].n}" è adiacente con stesso colore (${selColor}).\nLo script Google Sheets le unirebbe in un'unica prenotazione.\n\nCambia colore, oppure OK per salvare comunque.`);
    if(!ok) return;
  }

  const room = ROOMS.find(r=>r.id===rid);
  const bedStr = buildBedString(bedCounts);
  // Preserva dbId/dbRow se stiamo modificando una prenotazione esistente
  const existingB = editId ? bookings.find(b=>b.id===editId) : null;

  // Gestione anagrafica cliente (async, non blocca il salvataggio)
  let _clienteId = existingB?.clienteId || null;
  if (typeof gestisciClienteAlSalvataggio === 'function') {
    try { _clienteId = await gestisciClienteAlSalvataggio(name) || _clienteId; } catch(e) { console.warn('[anagrafica]:', e.message); }
  }

  const newB = {
    id: editId || nid++, r: rid, n: name, d: bedStr, c: selColor,
    s: sd, e: ed, note, fromSheet: false,
    sheetName: sheetName(sd.getFullYear(), sd.getMonth()),
    cameraName: room?.name || rid,
    pending: true,
    dbId:      existingB?.dbId      || null,
    dbRow:     existingB?.dbRow     || null,
    ts:        existingB?.ts        || null,
    fonte:     'app',
    clienteId: _clienteId,
  };

  // Aggiorna stato locale immediatamente (ottimistic UI)
  if(editId){ const i=bookings.findIndex(b=>b.id===editId); if(i>=0) bookings[i]=newB; editId=null; }
  else bookings.push(newB);

  closeModal(); render();
  showLoading('Scrittura sul foglio Google…');

  try {
    // bridgeSalva gestisce internamente cancel+write per le modifiche
    // e aggiorna bookings[] + render() dopo la risposta Apps Script
    await bridgeSalva(newB, existingB || null);
    const idx=bookings.findIndex(b=>b.id===newB.id);
    if(idx>=0){ bookings[idx].fromSheet=true; bookings[idx].pending=false; }
    hideLoading(); render();
    showToast(`✓ "${name}" salvato sul foglio Google`, 'success');
  } catch(e) {
    hideLoading();
    const idx=bookings.findIndex(b=>b.id===newB.id);
    if(idx>=0) bookings[idx].pending=false;
    render();
    showToast('Errore scrittura: ' + e.message, 'error');
    console.error(e);
  }
}

// ═══════════════════════════════════════════════════════════════════
// TOOLTIP
// ═══════════════════════════════════════════════════════════════════
function showTT(e,id){
  const b=bookings.find(x=>x.id===id); if(!b) return;
  const adj=adjConflict(b);
  const isMulti = b.s.getMonth() !== b.e.getMonth() || b.s.getFullYear() !== b.e.getFullYear();
  document.getElementById('tt-n').textContent=b.n+(adj.length?' ⚠':'')+(isMulti?' 📅':'');
  document.getElementById('tt-r').textContent=`Camera: ${roomName(b.r)}`;
  document.getElementById('tt-d').textContent=`${fmt(b.s)} → ${fmt(b.e)} (${nights(b.s,b.e)} notti${isMulti?' · multi-mese':''})`;
  document.getElementById('tt-b').textContent=`${b.d}`+(adj.length?' · ⚠ colore adiacente':'')+(b.pending?' · ⏳ in salvataggio':'');
  const tt=document.getElementById('tt'); tt.style.display='block'; movTT(e);
}
function movTT(e){
  const tt=document.getElementById('tt'); if(tt.style.display!=='block') return;
  let lx=e.clientX+14,ly=e.clientY-10;
  if(lx+160>window.innerWidth) lx=e.clientX-170;
  if(ly+100>window.innerHeight) ly=e.clientY-110;
  tt.style.left=lx+'px'; tt.style.top=ly+'px';
}
function hideTT(){ document.getElementById('tt').style.display='none'; }
document.addEventListener('mousemove', movTT);
document.addEventListener('touchstart', hideTT, { passive: true });


// ═══════════════════════════════════════════════════════════════════
// REPORT PRESENZE GIORNALIERE — stile "preconto mammana"
//
// Per ogni giorno del periodo mostra:
//   - Camere standard occupate (n. posti letto)
//   - Camere Duplex occupate (n. posti letto)
//   - Cene servite
//   - Supplemento cambio lenzuola
// Con totali mensili, prezzo unitario e importo per categoria.
// ═══════════════════════════════════════════════════════════════════

// Categorizzazione camere per report presenze
// Le categorie sono configurabili in cfg.reportCategorie:
//   { [roomId]: 'standard'|'duplex'|'appartamento'|'ignora' }
// Se non configurata usa i default in base al gruppo
function categorizzaCamera(roomId, clienteNome) {
  const cfg = loadBillSettings();
  // Override per cliente specifico: cfg.reportCategorieCliente[clienteNome][roomId]
  const clKey = (clienteNome||'').toLowerCase().trim();
  const ovClient = cfg.reportCategorieCliente?.[clKey]?.[roomId];
  if (ovClient) return ovClient;
  // Override globale per camera
  const ovGlobal = cfg.reportCategorie?.[roomId];
  if (ovGlobal) return ovGlobal;
  // Default per gruppo
  const r = ROOMS.find(x=>x.id===roomId);
  if (!r) return 'standard';
  if (r.g === 'Appartamenti') return 'appartamento';
  if (['r100','r101','r102','r103','r104','r21','r22','r23','r24','r25','r31'].includes(roomId)) return 'duplex';
  return 'standard';
}

// Conta posti letto dalla disposizione
function postiLetto(disp) {
  if (!disp) return 1;
  const d = disp.toLowerCase();
  let n = 0;
  const m  = d.match(/(\d+)m/); if (m) n += parseInt(m[1])*2;
  const ms = d.match(/(\d+)m\/s|(\d+)ms/); if (ms) n += parseInt(ms[1]||ms[2]||1);
  const s  = d.match(/(\d+)s/); if (s) n += parseInt(s[1]);
  return n || 1;
}

function apriReportPresenze() {
  const g = window._gruppoCorrente;
  if (!g || !g.bookings.length) { showToast('Cerca prima le prenotazioni', 'error'); return; }

  // Determina il periodo
  const dalD = g.dalD || g.bookings.reduce((m,b)=>b.s<m?b.s:m, g.bookings[0].s);
  const alD  = g.alD  || g.bookings.reduce((m,b)=>b.e>m?b.e:m, g.bookings[0].e);

  // Costruisci array giorni del periodo
  const giorni = [];
  let cur = new Date(dalD); cur.setHours(12,0,0,0);
  const fine = new Date(alD); fine.setHours(12,0,0,0);
  while (cur <= fine) { giorni.push(new Date(cur)); cur.setDate(cur.getDate()+1); }

  const cfg = loadBillSettings();

  // Per ogni giorno calcola presenze per categoria
  const righe = giorni.map(giorno => {
    const dg = giorno.getTime();
    let std=0, dup=0, cene=0, cambi=0;

    g.bookings.forEach(b => {
      const bs = new Date(b.s); bs.setHours(12,0,0,0);
      const be = new Date(b.e); be.setHours(12,0,0,0);
      // Presente = check-in <= giorno < check-out
      if (bs.getTime() <= dg && dg < be.getTime()) {
        const cat = categorizzaCamera(b.r, g.nome);
        const posti = postiLetto(b.d);
        if (cat === 'duplex') dup += posti;
        else if (cat === 'standard') std += posti;
        // Extra dal conto: cene e cambi lenzuola
        const extra = getExtraForBooking(b.id);
        extra.forEach(ex => {
          if (ex.label && ex.label.toLowerCase().includes('cena')) cene += ex.qty||1;
          if (ex.label && (ex.label.toLowerCase().includes('lenzuola')||ex.label.toLowerCase().includes('cambio'))) cambi += ex.qty||1;
        });
        // Note della prenotazione (es. "2 cambi")
        if (b.note) {
          const nc = b.note.match(/(\d+)\s*cambi?/i);
          if (nc) cambi += parseInt(nc[1]);
        }
      }
    });

    return { giorno, std, dup, cene, cambi };
  });

  // Recupera tariffe per categoria dalle impostazioni
  const tarStd = cfg.tariffeCamere?.['prezzoStd'] || cfg.tariffe?.s || 40;
  const tarDup = cfg.tariffeCamere?.['prezzoDup'] || cfg.tariffe?.m || 50;
  const tarCena= (cfg.extra||[]).find(e=>e.id==='cena')?.prezzo || 0;
  const tarCamb= (cfg.extra||[]).find(e=>e.id==='cambioLenzuola')?.prezzo || 0;

  // Totali colonna
  const totStd  = righe.reduce((s,r)=>s+r.std, 0);
  const totDup  = righe.reduce((s,r)=>s+r.dup, 0);
  const totCene = righe.reduce((s,r)=>s+r.cene, 0);
  const totCamb = righe.reduce((s,r)=>s+r.cambi, 0);

  const impStd  = parseFloat((totStd  * tarStd ).toFixed(2));
  const impDup  = parseFloat((totDup  * tarDup ).toFixed(2));
  const impCene = parseFloat((totCene * tarCena).toFixed(2));
  const impCamb = parseFloat((totCamb * tarCamb).toFixed(2));
  const totFatt = parseFloat((impStd+impDup+impCene+impCamb).toFixed(2));

  const periodoLabel = `${fmt(dalD)} — ${fmt(alD)}`;
  const meseLabel    = dalD.toLocaleDateString('it-IT',{month:'long',year:'numeric'});

  // ── Genera HTML del report ──
  const righeTabella = righe.map(r => {
    const hasData = r.std||r.dup||r.cene||r.cambi;
    return `<tr style="${hasData?'':'color:#bbb'}">
      <td style="padding:3px 6px;font-size:11px;white-space:nowrap">${fmt(r.giorno)}</td>
      <td style="padding:3px 6px;text-align:center;font-size:11px">${r.std||''}</td>
      <td style="padding:3px 6px;text-align:center;font-size:11px">${r.dup||''}</td>
      <td style="padding:3px 6px;text-align:center;font-size:11px">${r.cene||''}</td>
      <td style="padding:3px 6px;text-align:center;font-size:11px">${r.cambi||''}</td>
    </tr>`;
  }).join('');

  const html = `
    <div class="doc-header">
      <div class="doc-hotel-name">${cfg.hotelName}</div>
      ${cfg.hotelAddress?`<div class="doc-hotel-sub">${cfg.hotelAddress}</div>`:''}
      <div class="doc-type">Report Presenze — ${g.nome}</div>
    </div>
    <div class="doc-section">
      <div class="doc-section-title">Periodo: ${periodoLabel}</div>
      <table class="doc-table" style="width:100%;font-size:12px">
        <thead>
          <tr style="background:#f5f5f5">
            <th style="padding:5px 6px;text-align:left">Data</th>
            <th style="padding:5px 6px;text-align:center">Cam. Standard<br><span style="font-weight:400;font-size:10px">(posti)</span></th>
            <th style="padding:5px 6px;text-align:center">Cam. Duplex<br><span style="font-weight:400;font-size:10px">(posti)</span></th>
            <th style="padding:5px 6px;text-align:center">Cene<br><span style="font-weight:400;font-size:10px">(n.)</span></th>
            <th style="padding:5px 6px;text-align:center">Cambi<br><span style="font-weight:400;font-size:10px">(lenzuola)</span></th>
          </tr>
        </thead>
        <tbody>${righeTabella}</tbody>
        <tfoot>
          <tr style="border-top:2px solid #333">
            <td style="padding:5px 6px;font-weight:700;font-size:12px">Totale ${meseLabel}</td>
            <td style="padding:5px 6px;text-align:center;font-weight:700;color:#c0392b;font-size:13px">${totStd}</td>
            <td style="padding:5px 6px;text-align:center;font-weight:700;color:#c0392b;font-size:13px">${totDup}</td>
            <td style="padding:5px 6px;text-align:center;font-weight:700;color:#c0392b;font-size:13px">${totCene||'—'}</td>
            <td style="padding:5px 6px;text-align:center;font-weight:700;color:#c0392b;font-size:13px">${totCamb||'—'}</td>
          </tr>
          <tr>
            <td style="padding:4px 6px;font-size:11px;color:#666">Prezzo unitario</td>
            <td style="padding:4px 6px;text-align:center;font-size:11px;color:#666">€ ${tarStd.toFixed(2)}</td>
            <td style="padding:4px 6px;text-align:center;font-size:11px;color:#666">€ ${tarDup.toFixed(2)}</td>
            <td style="padding:4px 6px;text-align:center;font-size:11px;color:#666">${tarCena>0?'€ '+tarCena.toFixed(2):'—'}</td>
            <td style="padding:4px 6px;text-align:center;font-size:11px;color:#666">${tarCamb>0?'€ '+tarCamb.toFixed(2):'—'}</td>
          </tr>
          <tr>
            <td style="padding:4px 6px;font-size:11px;color:#666">Importo</td>
            <td style="padding:4px 6px;text-align:center;font-size:11px;color:#666">€ ${impStd.toFixed(2)}</td>
            <td style="padding:4px 6px;text-align:center;font-size:11px;color:#666">€ ${impDup.toFixed(2)}</td>
            <td style="padding:4px 6px;text-align:center;font-size:11px;color:#666">${impCene>0?'€ '+impCene.toFixed(2):'—'}</td>
            <td style="padding:4px 6px;text-align:center;font-size:11px;color:#666">${impCamb>0?'€ '+impCamb.toFixed(2):'—'}</td>
          </tr>
          <tr style="border-top:2px solid #333;background:#f9f9f9">
            <td colspan="3" style="padding:8px 6px;font-weight:700;font-size:14px">Totale fattura</td>
            <td colspan="2" style="padding:8px 6px;text-align:right;font-weight:700;font-size:14px;color:#2d6a4f">€ ${totFatt.toFixed(2)}</td>
          </tr>
        </tfoot>
      </table>
    </div>
    <div class="doc-section" style="font-size:11px;color:#666">
      <b>Dettaglio soggiorni (${g.bookings.length} prenotazioni)</b><br>
      ${g.bookings.map(b=>{
        const room=ROOMS.find(r=>r.id===b.r);
        return `${room?.name||roomName(b.r)} · ${fmt(b.s)}→${fmt(b.e)} · ${b.d||''}`;
      }).join('<br>')}
    </div>
    <div class="doc-footer">${cfg.hotelName} · Preconto ${meseLabel} · ${new Date().toLocaleDateString('it-IT')}</div>`;

  // Apri nel PDF overlay
  document.getElementById('pdfTitle').textContent = `Report presenze — ${g.nome} · ${meseLabel}`;
  document.getElementById('printDoc').innerHTML = html;
  document.getElementById('pdfOverlay').classList.add('open');
  _currentPdfBid = '__report__';
}

function esportaReportCSV() {
  const g = window._gruppoCorrente;
  if (!g || !g.bookings.length) { showToast('Cerca prima le prenotazioni','error'); return; }

  const dalD = g.dalD || g.bookings.reduce((m,b)=>b.s<m?b.s:m, g.bookings[0].s);
  const alD  = g.alD  || g.bookings.reduce((m,b)=>b.e>m?b.e:m, g.bookings[0].e);
  const cfg  = loadBillSettings();

  // Giorni del periodo
  const giorni = [];
  let cur = new Date(dalD); cur.setHours(12,0,0,0);
  const fine = new Date(alD); fine.setHours(12,0,0,0);
  while (cur <= fine) { giorni.push(new Date(cur)); cur.setDate(cur.getDate()+1); }

  // Calcola presenze per giorno
  const righe = giorni.map(giorno => {
    const dg = giorno.getTime();
    let std=0, dup=0, cene=0, cambi=0;
    const camOccupate = [];
    g.bookings.forEach(b => {
      const bs = new Date(b.s); bs.setHours(12,0,0,0);
      const be = new Date(b.e); be.setHours(12,0,0,0);
      if (bs.getTime() <= dg && dg < be.getTime()) {
        const cat   = categorizzaCamera(b.r, g.nome);
        const posti = postiLetto(b.d);
        const room  = ROOMS.find(r=>r.id===b.r);
        camOccupate.push(room?.name||roomName(b.r));
        if (cat==='duplex') dup+=posti;
        else if (cat==='standard') std+=posti;
        const extra = getExtraForBooking(b.id);
        extra.forEach(ex=>{
          if (ex.label?.toLowerCase().includes('cena')) cene+=ex.qty||1;
          if (ex.label?.toLowerCase().includes('lenzuola')||ex.label?.toLowerCase().includes('cambio')) cambi+=ex.qty||1;
        });
        if (b.note) { const nc=b.note.match(/(\d+)\s*cambi?/i); if(nc) cambi+=parseInt(nc[1]); }
      }
    });
    return { giorno, std, dup, cene, cambi, camere: camOccupate.join(';') };
  });

  const tarStd  = cfg.tariffe?.s || 40;
  const tarDup  = cfg.tariffe?.m || 50;
  const tarCena = (cfg.extra||[]).find(e=>e.id==='cena')?.prezzo || 0;
  const tarCamb = (cfg.extra||[]).find(e=>e.id==='cambioLenzuola')?.prezzo || 0;

  const totStd  = righe.reduce((s,r)=>s+r.std,0);
  const totDup  = righe.reduce((s,r)=>s+r.dup,0);
  const totCene = righe.reduce((s,r)=>s+r.cene,0);
  const totCamb = righe.reduce((s,r)=>s+r.cambi,0);

  // Costruisci CSV
  const sep = '\t'; // TSV — Google Sheets lo importa direttamente con copia-incolla
  const lines = [];

  // Intestazione
  lines.push(['Preconto', g.nome, fmt(dalD) + ' - ' + fmt(alD)].join(sep));
  lines.push('');
  lines.push(['Data','Camere standard (posti)','Camere Duplex (posti)','Cene servite','Cambio lenzuola','Camere occupate'].join(sep));

  // Righe dati
  righe.forEach(r => {
    lines.push([
      fmt(r.giorno),
      r.std || '',
      r.dup || '',
      r.cene || '',
      r.cambi || '',
      r.camere
    ].join(sep));
  });

  // Riga vuota + totali
  lines.push('');
  lines.push(['Totale mese', totStd, totDup, totCene||'', totCamb||'', ''].join(sep));
  lines.push(['Prezzo unitario', '€'+tarStd, '€'+tarDup, tarCena?'€'+tarCena:'', tarCamb?'€'+tarCamb:'', ''].join(sep));
  lines.push(['Importo',
    '€'+(totStd*tarStd).toFixed(2),
    '€'+(totDup*tarDup).toFixed(2),
    tarCena?'€'+(totCene*tarCena).toFixed(2):'',
    tarCamb?'€'+(totCamb*tarCamb).toFixed(2):'',
    ''
  ].join(sep));
  lines.push('');
  const tot = (totStd*tarStd)+(totDup*tarDup)+(totCene*tarCena)+(totCamb*tarCamb);
  lines.push(['Totale fattura','','','','€'+tot.toFixed(2),''].join(sep));
  lines.push('');
  lines.push('');
  lines.push(['Dettaglio soggiorni','','','','',''].join(sep));
  lines.push(['Camera','Nome','Check-in','Check-out','Notti','Disposizione'].join(sep));
  g.bookings.forEach(b => {
    const room = ROOMS.find(r=>r.id===b.r);
    lines.push([
      room?.name||roomName(b.r), b.n, fmt(b.s), fmt(b.e), nights(b.s,b.e), b.d||''
    ].join(sep));
  });

  const tsv   = lines.join('\n');
  const meseL = dalD.toLocaleDateString('it-IT',{month:'long',year:'numeric'});
  const nome  = 'preconto_'+g.nome.replace(/\s+/g,'_')+'_'+meseL.replace(/\s+/g,'_')+'.tsv';

  // Mostra overlay con download + istruzioni
  const blob = new Blob([tsv], {type:'text/tab-separated-values;charset=utf-8'});
  const url  = URL.createObjectURL(blob);
  const ov   = document.createElement('div');
  ov.style.cssText='position:fixed;inset:0;background:rgba(0,0,0,.6);z-index:9999;display:flex;align-items:center;justify-content:center;padding:20px;box-sizing:border-box';
  ov.innerHTML=`<div style="background:var(--surface);border-radius:16px;padding:24px;max-width:380px;width:100%;text-align:center">
    <div style="font-size:32px;margin-bottom:10px">📋</div>
    <div style="font-weight:700;font-size:15px;margin-bottom:6px">Report TSV pronto</div>
    <div style="font-size:11px;color:var(--text3);margin-bottom:6px;word-break:break-all">${nome}</div>
    <div style="font-size:11px;color:var(--text3);margin-bottom:18px;text-align:left;background:var(--surface2);padding:10px;border-radius:8px;line-height:1.6">
      💡 <b>Come aprire in Google Sheets:</b><br>
      1. Scarica il file<br>
      2. Vai su Google Sheets → File → Importa<br>
      3. Carica il file, separatore: <b>Tabulazione</b><br>
      oppure apri con Excel e copia-incolla
    </div>
    <a href="${url}" download="${nome}" id="_csvLink"
      style="display:block;background:#2d6a4f;color:#fff;padding:12px;border-radius:10px;text-decoration:none;font-weight:700;font-size:14px;margin-bottom:10px">
      ⬇ Scarica TSV
    </a>
    <button onclick="this.closest('div[style*=fixed]').remove()"
      style="background:none;border:1px solid var(--border);padding:10px;border-radius:8px;cursor:pointer;font-size:13px;color:var(--text2);width:100%">Chiudi</button>
  </div>`;
  document.body.appendChild(ov);
  ov.querySelector('#_csvLink').addEventListener('click',()=>setTimeout(()=>{URL.revokeObjectURL(url);ov.remove();},1000));
  ov.addEventListener('click',e=>{if(e.target===ov)ov.remove();});
}

// ═══════════════════════════════════════════════════════════════════
// RICERCA PRENOTAZIONI
// ═══════════════════════════════════════════════════════════════════

let _searchFilter = 'all';

function openSearch() {
  document.getElementById('searchOverlay').classList.add('open');
  hideTT();
  // Reset
  document.getElementById('searchInput').value = '';
  _searchFilter = 'all';
  document.querySelectorAll('.sf-chip').forEach(c => c.classList.toggle('active', c.dataset.f === 'all'));
  runSearch();
  setTimeout(() => document.getElementById('searchInput').focus(), 120);
}

function closeSearch() {
  document.getElementById('searchOverlay').classList.remove('open');
}

function setFilter(el, f) {
  _searchFilter = f;
  document.querySelectorAll('.sf-chip').forEach(c => c.classList.remove('active'));
  el.classList.add('active');
  runSearch();
}

function runSearch() {
  const q     = (document.getElementById('searchInput').value || '').toLowerCase().trim();
  const now   = new Date(); now.setHours(0,0,0,0);
  const mesS  = new Date(curY, curM, 1);
  const mesE  = new Date(curY, curM+1, 0);

  let risultati = [...bookings];

  // Filtro temporale
  if (_searchFilter === 'oggi') {
    risultati = risultati.filter(b => b.s <= now && b.e > now);
  } else if (_searchFilter === 'future') {
    risultati = risultati.filter(b => b.s > now);
  } else if (_searchFilter === 'passate') {
    risultati = risultati.filter(b => b.e <= now);
  } else if (_searchFilter === 'mese') {
    risultati = risultati.filter(b => b.s <= mesE && b.e >= mesS);
  }

  // Filtro testo
  if (q) {
    risultati = risultati.filter(b => {
      const cam = (roomName(b.r)||'').toLowerCase();
      const grp = (ROOMS.find(r=>r.id===b.r)?.g||'').toLowerCase();
      return b.n.toLowerCase().includes(q)
          || cam.includes(q)
          || grp.includes(q)
          || (b.d||'').toLowerCase().includes(q)
          || (b.note||'').toLowerCase().includes(q)
          || fmt(b.s).includes(q)
          || fmt(b.e).includes(q);
    });
  }

  // Ordina: prima le future/in corso, poi le passate; dentro ogni gruppo per data
  const ts = now.getTime();
  risultati.sort((a,b) => {
    const aPass = a.e.getTime() <= ts, bPass = b.e.getTime() <= ts;
    if (aPass !== bPass) return aPass ? 1 : -1;
    return a.s - b.s;
  });

  renderSearchResults(risultati, q);
}

function _highlight(text, q) {
  if (!q) return escHtml(text);
  const idx = text.toLowerCase().indexOf(q.toLowerCase());
  if (idx < 0) return escHtml(text);
  return escHtml(text.slice(0,idx))
    + '<mark>' + escHtml(text.slice(idx, idx+q.length)) + '</mark>'
    + escHtml(text.slice(idx+q.length));
}
function escHtml(s) {
  return String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
}

function renderSearchResults(lista, q) {
  const el = document.getElementById('searchResults');
  if (!lista.length) {
    el.innerHTML = `<div class="search-empty"><div class="emptyicon">🔍</div><div>${q?'Nessun risultato per "'+escHtml(q)+'"':'Nessuna prenotazione'}</div></div>`;
    return;
  }

  const now = new Date(); now.setHours(0,0,0,0);
  let lastGroup = null;

  el.innerHTML = lista.map(b => {
    const room   = ROOMS.find(r=>r.id===b.r);
    const n      = nights(b.s,b.e);
    const inCorso= b.s <= now && b.e > now;
    const futura = b.s > now;
    const group  = futura ? 'future' : inCorso ? 'in corso' : 'passate';

    let header = '';
    if (group !== lastGroup) {
      lastGroup = group;
      const label = group==='future'?'🗓 Future':group==='in corso'?'✅ In corso':'📁 Passate';
      header = `<div style="padding:6px 14px 4px;font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.08em;color:var(--text3);background:var(--surface2)">${label}</div>`;
    }

    return header + `<div class="sr-item" onclick="goToBooking(${b.id})">
      <div class="sr-dot" style="background:${b.c}"></div>
      <div class="sr-body">
        <div class="sr-name">${_highlight(b.n, q)}</div>
        <div class="sr-meta">
          <span class="sr-badge">${escHtml(room?.name||roomName(b.r))}</span>
          <span class="sr-badge">${fmt(b.s)} → ${fmt(b.e)}</span>
          <span class="sr-badge">${n} nott${n===1?'e':'i'}</span>
          ${b.d?`<span class="sr-badge">${escHtml(b.d)}</span>`:''}
          ${inCorso?'<span class="sr-badge" style="background:#d4edda;border-color:#c3e6cb;color:#155724">In corso</span>':''}
        </div>
        ${b.note?`<div style="font-size:10px;color:var(--text3);margin-top:3px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis">📝 ${escHtml(b.note)}</div>`:''}
      </div>
      <div style="font-size:11px;color:var(--text3);flex-shrink:0;text-align:right;padding-top:2px">${escHtml(room?.g||'')}</div>
    </div>`;
  }).join('');
}

function goToBooking(id) {
  closeSearch();
  const b = bookings.find(x=>x.id===id); if(!b) return;
  // Naviga al mese corretto
  const bm = b.s.getMonth(), by = b.s.getFullYear();
  if (curM !== bm || curY !== by) { curM=bm; curY=by; render(); }
  // Scrolla alla data nel gantt
  setTimeout(() => {
    const day = b.s.getDate();
    const gs  = document.getElementById('gscroll');
    if (gs) gs.scrollLeft = Math.max(0, (day-3)*34+90);
    // Apri drawer
    showBookingDetail(id);
  }, 80);
}

// Chiudi search con Escape
document.addEventListener('keydown', e => {
  if (e.key === 'Escape' && document.getElementById('searchOverlay').classList.contains('open')) {
    closeSearch();
  }
});

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

function buildRoomSelect(){
  const sel=document.getElementById('fRoom');
  if (!sel) return;
  sel.innerHTML = ''; // svuota prima — idempotente
  const groups=[...new Set(ROOMS.map(r=>r.g))];
  groups.forEach(g=>{
    const og=document.createElement('optgroup'); og.label=g;
    ROOMS.filter(r=>r.g===g).forEach(r=>{ const o=document.createElement('option'); o.value=r.id; o.textContent=r.name; og.appendChild(o); });
    sel.appendChild(og);
  });
}

function checkW(){
  const w=window.innerWidth>=640;
  document.getElementById('btnAdd').style.display=w?'inline-flex':'none';
  document.getElementById('fabGroup').style.display=w?'none':'flex';
}
window.addEventListener('resize',checkW);

// Attendi Google Identity Services
window.addEventListener('load', () => {
  // Controlla subito il token nel fragment — non aspettare GSI
  // (il redirect flow non ha bisogno di google.accounts)
  handleOAuthRedirect();
  // Inizializza UI — DOM è pronto a questo punto
  dbg('▶ buildRoomSelect'); buildRoomSelect();
  dbg('▶ checkW'); checkW();
  dbg('▶ render'); render();
  dbg('✓ avvio ok');
});

// ═══════════════════════════════════════════════════════════════════
// MODULO FATTURAZIONE — CONTI & TARIFFE
// ═══════════════════════════════════════════════════════════════════
