// ═══════════════════════════════════════════════════════════════════
// email.js — Pannello monitoraggio email (lato client)
// Blip Hotel Management — build 18.10.5
//
// L'elaborazione automatica delle email avviene interamente
// in Apps Script (blip-appscript.gs) via trigger ogni 10 minuti.
// Questo modulo fornisce solo la UI di monitoraggio che legge
// il foglio EMAIL_LOG scritto da Apps Script.
//
// Dipende da: core.js, api.js, store.js
// ═══════════════════════════════════════════════════════════════════

const BLIP_VER_EMAIL = '1';
const EMAIL_LOG_SHEET_NAME = 'EMAIL_LOG';

// ═══════════════════════════════════════════════════════════════════
// LETTURA LOG DA FOGLIO EMAIL_LOG
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
    syncLog('⚠ Email: errore lettura log — ' + e.message, 'wrn');
    return [];
  }
}

// ═══════════════════════════════════════════════════════════════════
// PANNELLO EMAIL
// ═══════════════════════════════════════════════════════════════════

async function openEmailPanel() {
  let ov = document.getElementById('emailPanelOv');
  if (!ov) {
    ov = document.createElement('div');
    ov.id = 'emailPanelOv';
    ov.onclick = e => { if (e.target === ov) closeEmailPanel(); };
    ov.style.cssText = 'position:fixed;inset:0;z-index:200;display:flex;align-items:flex-end;justify-content:flex-end;background:rgba(0,0,0,.2)';
    ov.innerHTML = `
      <div style="width:min(440px,100vw);height:min(620px,92vh);background:var(--surface);border-radius:16px 0 0 0;box-shadow:0 -4px 32px rgba(0,0,0,.18);display:flex;flex-direction:column;overflow:hidden">
        <div style="padding:14px 16px 10px;border-bottom:1px solid var(--border);display:flex;align-items:center;gap:10px">
          <span style="font-size:18px">📧</span>
          <div style="flex:1">
            <div style="font-weight:700;font-size:14px">Richieste dal sito</div>
            <div style="font-size:10px;color:var(--text3)">Elaborate automaticamente da Apps Script ogni 10 min</div>
          </div>
          <button onclick="refreshEmailPanel()" style="background:var(--surface2);border:1px solid var(--border);border-radius:8px;padding:5px 10px;font-size:12px;cursor:pointer">↺ Aggiorna</button>
          <button onclick="closeEmailPanel()" style="background:transparent;border:none;font-size:20px;cursor:pointer;color:var(--text2);line-height:1">×</button>
        </div>
        <div id="emailPanelList" style="flex:1;overflow-y:auto;padding:10px 14px">
          <div style="color:var(--text3);font-size:12px;text-align:center;padding:30px">⏳ Caricamento...</div>
        </div>
        <div style="padding:10px 14px;border-top:1px solid var(--border);font-size:10px;color:var(--text3);line-height:1.5">
          📌 Le pre-prenotazioni appaiono sul Gantt in grigio.<br>
          ⚙ Configura IBAN e acconto in Tariffe → sezione Email.
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
  list.innerHTML = '<div style="color:var(--text3);font-size:12px;text-align:center;padding:30px">⏳ Lettura EMAIL_LOG...</div>';

  const entries = await loadEmailLog();

  if (!entries.length) {
    list.innerHTML = '<div style="color:var(--text3);font-size:12px;text-align:center;padding:30px">Nessuna richiesta elaborata.<br><br>Assicurati che il trigger Apps Script sia attivo:<br>menu 🏨 Script → <b>📧 Installa trigger email</b></div>';
    return;
  }

  list.innerHTML = entries.map(e => `
    <div style="border:1px solid var(--border);border-radius:8px;padding:10px;margin-bottom:8px;background:var(--surface2)">
      <div style="display:flex;align-items:center;gap:6px;margin-bottom:3px">
        <span>${e.disponibile ? '✅' : '❌'}</span>
        <span style="font-weight:600;font-size:13px">${e.nome}</span>
        <span style="margin-left:auto;font-size:10px;color:var(--text3)">${e.data}</span>
      </div>
      <div style="font-size:11px;color:var(--text3)">${e.checkin} → ${e.checkout} · ${e.persone} ospiti</div>
      ${e.camera ? `<div style="font-size:11px;margin-top:2px">📍 <b>${e.camera}</b></div>` : ''}
      ${e.preBlipId ? `<div style="font-size:10px;color:var(--accent);margin-top:2px">Pre-pren: ${e.preBlipId}</div>` : ''}
      <div style="font-size:10px;margin-top:3px;color:${e.disponibile?'var(--success,#2d6a4f)':'var(--danger)'}">
        ${e.disponibile ? '✉ Risposta inviata + pre-prenotazione creata' : '✉ Risposta inviata (nessuna disponibilità)'}
      </div>
    </div>`).join('');
}
