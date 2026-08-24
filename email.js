// ═══════════════════════════════════════════════════════════════════
// email.js — Pannello unificato "Richieste" (lato client)
// Blip Hotel Management — build 18.10.5
//
// Mostra in UN SOLO pannello le richieste in arrivo da DUE fonti/automazioni
// diverse, ciascuna gestita da un account e un'automazione indipendenti:
//
//   🌐 Sito web ("Mail da Sito Web")
//      - letta su ilborgosrl.montedoro@gmail.com da ScriptSito.gs
//        (progetto Apps Script standalone, fuori da questo repo)
//      - registrata da Codice.gs (bound a "Database prenotazioni",
//        davide.petix@gmail.com) nel tab RICHIESTE SITO, con prezzo e
//        disposizione già calcolati dal listino reale
//      - resta "in attesa" finché l'operatore non conferma/rifiuta dal
//        pannello "Richieste Sito Web" nel foglio (non da qui: quel
//        passaggio manda l'email di conferma finale via un webhook con
//        segreto condiviso, che non ha senso esporre in questa pagina
//        lato client)
//
//   🏝 Sicily Divide (partner/agenzia)
//      - elaborata interamente qui in blip-appscript.gs, ogni 10 min,
//        sull'account davide.petix@gmail.com (stesso login di Blip)
//      - se disponibile: crea già una pre-prenotazione tratteggiata sul
//        Gantt e notifica te via email — nessuna azione richiesta qui,
//        la trovi/confermi direttamente sul calendario
//      - se non disponibile: ha già risposto in automatico al cliente
//      - il tab EMAIL_LOG è quindi uno STORICO di email già gestite,
//        non una coda di cose da fare
//
// Le due automazioni di backend restano invariate: qui si UNIFICA SOLO
// la lettura/visualizzazione, leggendo entrambi i tab con lo stesso
// helper già usato per EMAIL_LOG (apiFetch + DATABASE_SHEET_ID — stessa
// identità Google dell'operatore, nessun permesso nuovo).
//
// Dipende da: core.js, api.js, store.js, billing.js (per il prezzo delle
// richieste dal sito: loadBillSettings/calcolaMoltiplicatoreStagionale,
// già usate per i conti — così il prezzo mostrato qui resta sempre
// coerente col listino vero, mai un valore ricalcolato a parte).
// ═══════════════════════════════════════════════════════════════════

const BLIP_VER_EMAIL = '2';
const EMAIL_LOG_SHEET_NAME = 'EMAIL_LOG';
const RICHIESTE_SITO_SHEET_NAME = 'RICHIESTE SITO';

// Ordine colonne del tab RICHIESTE SITO (scritto da Codice.gs).
const RS_COLS = ['ID_RICHIESTA','TS_RICEVUTA','GMAIL_MSG_ID','GMAIL_THREAD_ID','NOME','EMAIL',
  'NUM_PERSONE','CHECK_IN','CHECK_OUT','LETTI_MATRIMONIALI','LETTI_SINGOLI','COLAZIONE','BAMBINI_0_3',
  'MESSAGGIO','STANZA_RICHIESTA','STANZA_SUGGERITA','DISPOSIZIONE_SUGGERITA','STATO','DRAFT_RICHIESTA_ID',
  'DRAFT_RICHIESTA_URL','ID_PRENOTAZIONE','TS_ULTIMA_AZIONE','NOTE_OPERATORE','NUM_NOTTI','PREZZO_NOTTE',
  'PREZZO_TOTALE','SCONTO_PCT','ALT_DESCRIZIONE','ALT_PREZZO_NOTTE','ALT_PREZZO_TOTALE'];

let _richiesteSitoCache = [];

// ═══════════════════════════════════════════════════════════════════
// LETTURA — EMAIL_LOG (Sicily Divide, storico)
// ═══════════════════════════════════════════════════════════════════

async function loadEmailLog() {
  if (!DATABASE_SHEET_ID) return [];
  try {
    const range = encodeURIComponent(EMAIL_LOG_SHEET_NAME + '!A2:J200');
    const url   = `https://sheets.googleapis.com/v4/spreadsheets/${DATABASE_SHEET_ID}/values/${range}`;
    const r     = await apiFetch(url);
    if (!r.ok) return [];
    const data  = await r.json();
    return (data.values || []).map(row => ({
      data:        row[0] || '',
      mittente:    row[1] || '',
      nome:        row[2] || '',
      checkin:     row[3] || '',
      checkout:    row[4] || '',
      persone:     row[5] || '',
      disponibile: row[6] === '✅',
      camera:      row[7] || '',
      preBlipId:   row[8] || '',
      stato:       row[9] || '',
    })).reverse(); // più recenti prima
  } catch(e) {
    syncLog('⚠ Email: errore lettura EMAIL_LOG — ' + e.message, 'wrn');
    return [];
  }
}

// ═══════════════════════════════════════════════════════════════════
// LETTURA — RICHIESTE SITO (form sito web, in attesa)
// ═══════════════════════════════════════════════════════════════════

async function loadRichiesteSito() {
  if (!DATABASE_SHEET_ID) return [];
  try {
    const range = encodeURIComponent(`'${RICHIESTE_SITO_SHEET_NAME}'!A2:AD9999`);
    const url   = `https://sheets.googleapis.com/v4/spreadsheets/${DATABASE_SHEET_ID}/values/${range}`;
    const r     = await apiFetch(url);
    if (!r.ok) return [];
    const data  = await r.json();
    const rows  = data.values || [];
    return rows.map(row => {
      const o = {};
      RS_COLS.forEach((c, i) => { o[c] = row[i] !== undefined ? row[i] : ''; });
      return o;
    }).filter(o => o.ID_RICHIESTA && o.STATO !== 'Confermata' && o.STATO !== 'Rifiutata')
      .sort((a, b) => String(b.TS_RICEVUTA).localeCompare(String(a.TS_RICEVUTA)));
  } catch(e) {
    syncLog('⚠ Email: errore lettura RICHIESTE SITO — ' + e.message, 'wrn');
    return [];
  }
}

// ═══════════════════════════════════════════════════════════════════
// PANNELLO UNIFICATO
// ═══════════════════════════════════════════════════════════════════

async function openEmailPanel() {
  let ov = document.getElementById('emailPanelOv');
  if (!ov) {
    ov = document.createElement('div');
    ov.id = 'emailPanelOv';
    ov.onclick = e => { if (e.target === ov) closeEmailPanel(); };
    ov.style.cssText = 'position:fixed;inset:0;z-index:200;display:flex;align-items:flex-end;justify-content:flex-end;background:rgba(0,0,0,.2)';
    ov.innerHTML = `
      <div style="width:min(460px,100vw);height:min(680px,94vh);background:var(--surface);border-radius:16px 0 0 0;box-shadow:0 -4px 32px rgba(0,0,0,.18);display:flex;flex-direction:column;overflow:hidden">
        <div style="padding:14px 16px 10px;border-bottom:1px solid var(--border);display:flex;align-items:center;gap:10px">
          <span style="font-size:18px">📨</span>
          <div style="flex:1">
            <div style="font-weight:700;font-size:14px">Richieste</div>
            <div style="font-size:10px;color:var(--text3)">Sito web + Sicily Divide, elaborate automaticamente</div>
          </div>
          <button onclick="refreshEmailPanel()" style="background:var(--surface2);border:1px solid var(--border);border-radius:8px;padding:5px 10px;font-size:12px;cursor:pointer">↺ Aggiorna</button>
          <button onclick="closeEmailPanel()" style="background:transparent;border:none;font-size:20px;cursor:pointer;color:var(--text2);line-height:1">×</button>
        </div>
        <div id="emailPanelList" style="flex:1;overflow-y:auto;padding:10px 14px">
          <div style="color:var(--text3);font-size:12px;text-align:center;padding:30px">⏳ Caricamento...</div>
        </div>
        <div style="padding:10px 14px;border-top:1px solid var(--border);font-size:10px;color:var(--text3);line-height:1.5">
          🌐 Sito: conferma/rifiuta resta nel pannello "Richieste Sito Web" del foglio.<br>
          🏝 Sicily Divide: le pre-prenotazioni appaiono sul Gantt in grigio tratteggiato.
        </div>
      </div>`;
    document.body.appendChild(ov);
  }
  ov.style.display = 'flex';
  await refreshEmailPanel();
}

function closeEmailPanel() {
  const ov = document.getElementById('emailPanelOv');
  if (ov) ov.style.display = 'none';
}

async function refreshEmailPanel() {
  const list = document.getElementById('emailPanelList');
  if (!list) return;
  list.innerHTML = '<div style="color:var(--text3);font-size:12px;text-align:center;padding:30px">⏳ Lettura richieste...</div>';

  const [richiesteSito, emailLog] = await Promise.all([loadRichiesteSito(), loadEmailLog()]);
  _richiesteSitoCache = richiesteSito;

  if (!richiesteSito.length && !emailLog.length) {
    list.innerHTML = '<div style="color:var(--text3);font-size:12px;text-align:center;padding:30px">Nessuna richiesta al momento.</div>';
    return;
  }

  let html = '';

  html += `<div style="font-size:11px;font-weight:700;color:var(--text3);text-transform:uppercase;letter-spacing:.04em;margin:2px 0 8px">🌐 Dal sito web — in attesa (${richiesteSito.length})</div>`;
  html += richiesteSito.length
    ? richiesteSito.map((r, i) => `
      <div onclick="openRichiestaSitoReply(${i})" style="cursor:pointer;border:1px solid var(--border);border-radius:8px;padding:10px;margin-bottom:8px;background:var(--surface2)">
        <div style="display:flex;align-items:center;gap:6px;margin-bottom:3px">
          <span style="font-weight:600;font-size:13px">${_rsEsc(r.NOME)}</span>
          <span style="margin-left:auto;font-size:10px;color:var(--accent);background:var(--accent-light);padding:1px 7px;border-radius:8px;text-transform:uppercase">${_rsEsc(r.STATO || 'Nuova')}</span>
        </div>
        <div style="font-size:11px;color:var(--text3)">${_rsEsc(r.CHECK_IN)} → ${_rsEsc(r.CHECK_OUT)} · ${_rsEsc(r.NUM_PERSONE)} ospiti</div>
        <div style="font-size:11px;color:var(--text2);margin-top:2px">${_rsEsc(_rsDescrizionePlain(parseInt(r.LETTI_MATRIMONIALI)||0, parseInt(r.LETTI_SINGOLI)||0))}</div>
        ${r.PREZZO_TOTALE ? `<div style="font-size:11px;font-weight:600;color:var(--accent);margin-top:2px">€${_rsFmt(r.PREZZO_NOTTE)}/notte · totale €${_rsFmt(r.PREZZO_TOTALE)}${parseFloat(r.SCONTO_PCT)>0?' (sconto '+r.SCONTO_PCT+'%)':''}</div>` : ''}
      </div>`).join('')
    : `<div style="color:var(--text3);font-size:11px;padding:6px 2px 16px">Nessuna richiesta dal sito in attesa. 🎉</div>`;

  html += `<div style="font-size:11px;font-weight:700;color:var(--text3);text-transform:uppercase;letter-spacing:.04em;margin:14px 0 8px">🏝 Sicily Divide — ultime richieste (${emailLog.length})</div>`;
  html += emailLog.length
    ? emailLog.map(e => `
      <div style="border:1px solid var(--border);border-radius:8px;padding:10px;margin-bottom:8px;background:var(--surface2)">
        <div style="display:flex;align-items:center;gap:6px;margin-bottom:3px">
          <span>${e.disponibile ? '✅' : '❌'}</span>
          <span style="font-weight:600;font-size:13px">${_rsEsc(e.nome)}</span>
          <span style="margin-left:auto;font-size:10px;color:var(--text3)">${_rsEsc(e.data)}</span>
        </div>
        <div style="font-size:11px;color:var(--text3)">${_rsEsc(e.checkin)} → ${_rsEsc(e.checkout)} · ${_rsEsc(e.persone)} ospiti</div>
        ${e.camera ? `<div style="font-size:11px;margin-top:2px">📍 <b>${_rsEsc(e.camera)}</b> — cercala con 🔍 per aprirla sul calendario</div>` : ''}
        ${e.preBlipId ? `<div style="font-size:10px;color:var(--accent);margin-top:2px">Pre-pren: ${_rsEsc(e.preBlipId)}</div>` : ''}
        <div style="font-size:10px;margin-top:3px;color:${e.disponibile?'var(--success,#2d6a4f)':'var(--danger)'}">
          ${e.disponibile ? '✉ Notifica inviata a te + pre-prenotazione creata' : '✉ Risposta automatica già inviata al cliente (nessuna disponibilità)'}
        </div>
      </div>`).join('')
    : `<div style="color:var(--text3);font-size:11px;padding:6px 2px">Nessuna richiesta Sicily Divide recente.</div>`;

  list.innerHTML = html;
}

// ═══════════════════════════════════════════════════════════════════
// POPUP RISPOSTA — richieste dal sito web
// Mostra la stessa proposta (prezzo reale, mai il numero camera) già
// preparata nella bozza email da ScriptSito.gs, ricalcolata qui lato
// client con lo stesso listino (loadBillSettings) per restare coerente.
// ═══════════════════════════════════════════════════════════════════

function openRichiestaSitoReply(i) {
  const r = _richiesteSitoCache[i];
  if (!r) return;

  const m = parseInt(r.LETTI_MATRIMONIALI) || 0;
  const s = parseInt(r.LETTI_SINGOLI) || 0;
  const notti = parseInt(r.NUM_NOTTI) || 1;
  const cfg = (typeof loadBillSettings === 'function')
    ? loadBillSettings()
    : { tariffe: { s: 35, ms: 38, m: 45, ag: 15 }, stagioni: [] };

  let corpo;
  if (!r.PREZZO_TOTALE) {
    corpo = 'Le proponiamo una sistemazione adatta alle vostre esigenze; vi contatteremo a breve con il dettaglio del prezzo.';
  } else {
    const dIn = _rsParseDataIt(r.CHECK_IN), dOut = _rsParseDataIt(r.CHECK_OUT);
    const molt = (dIn && dOut && typeof calcolaMoltiplicatoreStagionale === 'function')
      ? calcolaMoltiplicatoreStagionale(dIn, dOut, cfg.stagioni) : 1;
    const parti = _rsParti(m, s, cfg.tariffe);
    const notteBase = parti.reduce((a, b) => a + b, 0);
    const importoTxt = parti.length <= 1
      ? `€${_rsFmt(notteBase)}`
      : `€${parti.map(_rsFmt).join('+')} = €${_rsFmt(notteBase)}`;
    const testoNotti = notti === 1 ? '1 notte' : `${notti} notti`;
    const scontoPct = parseFloat(r.SCONTO_PCT) || 0;

    corpo = `Le proponiamo ${_rsDescrizionePlain(m, s)}, al prezzo di ${importoTxt} a notte`;
    if (notti > 1) {
      corpo += ` (per ${testoNotti}: €${_rsFmt(r.PREZZO_TOTALE)}${scontoPct > 0 ? ', con lo sconto per soggiorni lunghi del ' + scontoPct + '%' : ''})`;
    }
    corpo += '.';

    if (r.ALT_DESCRIZIONE) {
      corpo += `\n\nIn alternativa possiamo proporre ${r.ALT_DESCRIZIONE} al prezzo di €${_rsFmt(r.ALT_PREZZO_NOTTE)} a notte`;
      if (notti > 1) corpo += ` (per ${testoNotti}: €${_rsFmt(r.ALT_PREZZO_TOTALE)})`;
      corpo += '.';
    }
    if (Math.abs(molt - 1) > 0.001) {
      corpo += `\n\n(Nota: periodo con moltiplicatore stagionale ×${molt.toFixed(2)} già incluso nel prezzo sopra.)`;
    }
  }

  const dbUrl = `https://docs.google.com/spreadsheets/d/${DATABASE_SHEET_ID}/edit`;
  const draftBtn = r.DRAFT_RICHIESTA_URL
    ? `<a class="btn primary" href="${_rsEsc(r.DRAFT_RICHIESTA_URL)}" target="_blank" rel="noopener" style="text-decoration:none;justify-content:center;">✉ Apri bozza email</a>`
    : `<div class="errmsg show">Bozza non ancora creata: verrà generata automaticamente al prossimo controllo posta (ogni 10 min), oppure lancia "Importa ora" dal menu Richieste Sito Web nel foglio.</div>`;

  let ov = document.getElementById('richiestaSitoReplyOv');
  if (!ov) {
    ov = document.createElement('div');
    ov.id = 'richiestaSitoReplyOv';
    ov.className = 'moverlay';
    ov.onclick = e => { if (e.target === ov) closeRichiestaSitoReply(); };
    ov.innerHTML = `<div class="modal" id="richiestaSitoReplyModal" style="max-width:480px;"></div>`;
    document.body.appendChild(ov);
  }

  document.getElementById('richiestaSitoReplyModal').innerHTML = `
    <div class="mdrag"></div>
    <div class="mtitle" style="font-size:17px;">${_rsEsc(r.NOME)}</div>
    <div class="msub">${_rsEsc(r.CHECK_IN)} → ${_rsEsc(r.CHECK_OUT)} · ${_rsEsc(r.NUM_PERSONE)} persone${r.EMAIL ? ' · ' + _rsEsc(r.EMAIL) : ''}</div>
    ${r.MESSAGGIO ? `<div class="fg"><label class="fl">Messaggio del cliente</label><div style="font-size:12.5px;color:var(--text2);white-space:pre-wrap;">${_rsEsc(r.MESSAGGIO)}</div></div>` : ''}
    <div class="fg">
      <label class="fl">Proposta precompilata (già nella bozza email)</label>
      <div style="font-size:12.5px;color:var(--text);white-space:pre-wrap;line-height:1.6;background:var(--surface2);border:1px solid var(--border);border-radius:var(--radius);padding:10px 12px;">${_rsEsc(corpo)}</div>
    </div>
    <div class="mfoot" style="flex-direction:column;align-items:stretch;gap:8px;">
      ${draftBtn}
      <a class="btn" href="${dbUrl}" target="_blank" rel="noopener" title="Per confermare o rifiutare, usa il pannello nel foglio">📋 Conferma/rifiuta dal foglio</a>
      <button class="btn" onclick="closeRichiestaSitoReply()">Chiudi</button>
    </div>
  `;
  ov.classList.add('open');
  ov.style.display = 'flex';
}

function closeRichiestaSitoReply() {
  const ov = document.getElementById('richiestaSitoReplyOv');
  if (ov) { ov.classList.remove('open'); ov.style.display = 'none'; }
}

// ═══════════════════════════════════════════════════════════════════
// HELPER — prezzo/disposizione (porting di Codice.gs, stessa logica)
// ═══════════════════════════════════════════════════════════════════

function _rsEsc(v) {
  return String(v == null ? '' : v).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

function _rsFmt(v) {
  const n = parseFloat(v);
  if (isNaN(n)) return v;
  const r = Math.round(n * 100) / 100;
  const str = r.toFixed(2);
  return str.slice(-3) === '.00' ? str.slice(0, -3) : str.replace('.', ',');
}

function _rsDescrizionePlain(m, s) {
  if (m > 0 && s > 0) return (m === 1 ? 'una camera matrimoniale' : m + ' camere matrimoniali') + ' con ' + s + (s === 1 ? ' letto singolo aggiunto' : ' letti singoli aggiunti');
  if (m > 0) return m === 1 ? 'una camera matrimoniale' : m + ' camere matrimoniali';
  if (s === 1) return 'una camera singola';
  if (s >= 2) return 'una camera con ' + s + ' letti singoli';
  return 'una sistemazione da definire insieme';
}

function _rsParti(m, s, tariffe) {
  const parti = [];
  if (m > 0 && s > 0) { for (let i = 0; i < m; i++) parti.push(tariffe.m); for (let j = 0; j < s; j++) parti.push(tariffe.ag); }
  else if (m > 0) { for (let i = 0; i < m; i++) parti.push(tariffe.m); }
  else if (s === 1) parti.push(tariffe.s);
  else if (s === 2) parti.push(tariffe.m);
  else if (s > 2) { parti.push(tariffe.m); for (let i = 0; i < s - 2; i++) parti.push(tariffe.ag); }
  return parti;
}

function _rsParseDataIt(v) {
  const parti = String(v || '').trim().split('/');
  if (parti.length !== 3) return null;
  const d = parseInt(parti[0], 10), mo = parseInt(parti[1], 10) - 1, y = parseInt(parti[2], 10);
  return new Date(Date.UTC(y, mo, d));
}
