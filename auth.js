// ═══════════════════════════════════════════════════════════════════
// auth.js — OAuth 2.0, sessione, login/logout, rigenera
// Blip Hotel Management — build 18.10.4
//
// Responsabilità:
//   • OAuth 2.0 implicit flow (redirect + hash parsing)
//   • Generazione state CSRF
//   • onLoginSuccess: setup post-login (tariffe, loadFromSheets)
//   • logout, forceSync, rigenera
//
// Dipende da: core.js, api.js
// Chiamate runtime verso:
//   • loadFromSheets()      — definita in store.js
//   • loadBillSettingsDB()  — definita in billing.js
//   • invalidateDbCache()   — definita in store.js
//   • stopBgSync()          — definita in store.js
//   • render()              — definita in gantt.js
// ═══════════════════════════════════════════════════════════════════

const BLIP_VER_AUTH = '1'; // ← incrementa ad ogni modifica

// ═══════════════════════════════════════════════════════════════════
// HELPERS OAUTH
// ═══════════════════════════════════════════════════════════════════

function randomState() {
  return Array.from(crypto.getRandomValues(new Uint8Array(16)))
    .map(b => b.toString(16).padStart(2,'0')).join('');
}

function getRedirectUri() {
  const path = location.pathname.replace(/\/index\.html$/, '/').replace(/([^/])$/, '$1/');
  return location.origin + path;
}

// ═══════════════════════════════════════════════════════════════════
// LOGIN — redirect a Google OAuth
// ═══════════════════════════════════════════════════════════════════

function startLogin() {
  dbg('▶ startLogin');
  document.getElementById('loginErr').textContent = '';
  const state = randomState();
  sessionStorage.setItem('oauth_state', state);
  const params = new URLSearchParams({
    client_id:     CLIENT_ID,
    redirect_uri:  getRedirectUri(),
    response_type: 'token',
    scope:         SCOPES,
    state,
    include_granted_scopes: 'true',
  });
  location.href = 'https://accounts.google.com/o/oauth2/v2/auth?' + params.toString();
}

// ═══════════════════════════════════════════════════════════════════
// CALLBACK OAUTH — parsing hash dopo il redirect di Google
// ═══════════════════════════════════════════════════════════════════
// Chiamata da gantt.js nel listener window.load.
// Restituisce true se un hash OAuth era presente (redirect), false altrimenti.
// ═══════════════════════════════════════════════════════════════════

function handleOAuthRedirect() {
  dbg('▶ handleOAuthRedirect hash=' + (location.hash ? 'si' : 'no'));
  const hash = location.hash.slice(1);
  if (!hash) return false;
  const params = new URLSearchParams(hash);
  const token  = params.get('access_token');
  const error  = params.get('error');
  const state  = params.get('state');

  history.replaceState(null, '', location.pathname);

  if (error) {
    document.getElementById('loginErr').textContent = 'Accesso negato: ' + error;
    return true;
  }
  if (!token) return false;

  const saved = sessionStorage.getItem('oauth_state');
  sessionStorage.removeItem('oauth_state');
  if (saved && state !== saved) {
    document.getElementById('loginErr').textContent = 'Errore sicurezza. Riprova.';
    return true;
  }

  accessToken = token;
  onLoginSuccess();
  return true;
}

function initGoogleAuth() {
  // Wrapper mantenuto per compatibilità — handleOAuthRedirect viene chiamata
  // direttamente dal listener window.load in gantt.js.
  handleOAuthRedirect();
}

// ═══════════════════════════════════════════════════════════════════
// onLoginSuccess — setup post-autenticazione
// ═══════════════════════════════════════════════════════════════════
// Sequenza:
//   1. Nasconde login screen
//   2. Carica avatar utente (Google userinfo)
//   3. Carica tariffe dal DB (se DATABASE_SHEET_ID configurato)
//   4. Chiama loadFromSheets() → render iniziale
//
// loadBillSettingsDB e loadFromSheets sono chiamate lazily con ?. :
// billing.js e store.js potrebbero non essere ancora definiti al primo
// caricamento in scenari edge (non succede nell'ordine normale, ma
// il pattern difensivo evita eccezioni in caso di refactor futuro).
// ═══════════════════════════════════════════════════════════════════

async function onLoginSuccess() {
  document.getElementById('loginScreen').style.display = 'none';

  // Avatar utente
  try {
    const r = await fetch('https://www.googleapis.com/oauth2/v3/userinfo', {
      headers: { Authorization: 'Bearer ' + accessToken }
    });
    const u = await r.json();
    if (u.picture) {
      const av = document.getElementById('userAvatar');
      av.src = u.picture; av.style.display = 'block';
      av.title = `${u.name} — Clicca per uscire`;
    }
  } catch(e) {}

  // Tariffe dal DB (await: devono essere disponibili prima del primo render)
  if (DATABASE_SHEET_ID || loadDbSheetId()) {
    DATABASE_SHEET_ID = DATABASE_SHEET_ID || loadDbSheetId();
    try {
      // loadBillSettingsDB è in billing.js — chiamata lazy per sicurezza
      if (typeof loadBillSettingsDB === 'function') {
        const s = await loadBillSettingsDB();
        if (s) localStorage.setItem('hotelBillSettings', JSON.stringify(s));
      }
    } catch(e) {}
  }

  // Caricamento principale
  if (typeof loadFromSheets === 'function') await loadFromSheets();
}

// ═══════════════════════════════════════════════════════════════════
// RIGENERA — richiama la Web App Apps Script per rigenerare JSON_ANNUALE
// ═══════════════════════════════════════════════════════════════════

async function rigenera() {
  const url = (window._blipWebAppUrl || (typeof loadBillSettings === 'function' ? loadBillSettings().webAppUrl : '') || '').trim();
  if (!url) {
    showToast('URL Web App non configurato — vai in ⚙ Tariffe e salva', 'error');
    return;
  }
  showToast('📡 Chiamata Web App…', 'info');
  try {
    await fetch(`${url}?anno=${new Date().getFullYear()}&ts=${Date.now()}`, { method:'GET', mode:'no-cors' }).catch(() => {});
    showToast('⏳ Attendi 7 secondi…', 'info');
    await new Promise(r => setTimeout(r, 7000));

    // Verifica che il JSON_ANNUALE sia stato effettivamente aggiornato
    try {
      const currentYear = new Date().getFullYear();
      const allSheets   = loadAnnualSheets();
      const sheetEntry  = allSheets.find(e => e.sheetId && e.year === currentYear)
                       || allSheets.find(e => e.sheetId);
      if (sheetEntry && typeof readJSONAnnuale === 'function') {
        const fresh = await readJSONAnnuale(sheetEntry.sheetId);
        if (fresh && fresh.length > 0) {
          annualSheets = loadAnnualSheets();
          if (typeof loadFromSheets === 'function') await loadFromSheets();
          showToast('✓ Calendario rigenerato', 'success');
          return;
        }
      }
    } catch(e2) {}

    showToast('⚠ Web App non ha risposto — verifica il deploy in Apps Script (Distribuisci → Gestisci distribuzioni → Accesso: Chiunque)', 'warning');
  } catch(e) {
    console.warn('[rigenera] errore:', e.message);
    showToast('❌ Errore Web App: ' + e.message, 'error');
  }
}

// ═══════════════════════════════════════════════════════════════════
// FORCE SYNC / LOGOUT
// ═══════════════════════════════════════════════════════════════════

function forceSync() {
  if (typeof loadFromSheets === 'function') loadFromSheets._forceNext = true;
  if (typeof invalidateDbCache === 'function') invalidateDbCache();
  if (typeof stopBgSync === 'function') stopBgSync();
  if (typeof loadFromSheets === 'function') loadFromSheets();
}

function logout() {
  if (!confirm('Vuoi uscire?')) return;
  accessToken = null;
  bookings = [];
  // Svuota le cache interne di store.js
  // Nota: _sheetIdCaches non esiste (rimosso) — _metaCacheTitles è in store.js
  if (typeof _metaCacheTitles !== 'undefined') Object.keys(_metaCacheTitles).forEach(k => delete _metaCacheTitles[k]);
  if (typeof _dbSheetIdCache !== 'undefined') _dbSheetIdCache = null;
  if (typeof invalidateDbCache === 'function') invalidateDbCache();
  if (typeof stopBgSync === 'function') stopBgSync();
  document.getElementById('loginScreen').style.display = 'flex';
  document.getElementById('userAvatar').style.display = 'none';
}
