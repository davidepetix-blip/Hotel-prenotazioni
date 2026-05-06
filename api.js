// ═══════════════════════════════════════════════════════════════════
// api.js — HTTP layer: token bucket, apiFetch, re-auth silenzioso
// Blip Hotel Management — build 18.10.4
//
// Responsabilità:
//   • Unico modulo che chiama fetch() per le Google API
//   • Token bucket 45tok/900ms — rate limiting preventivo
//   • Retry automatico su 429 con backoff esponenziale
//   • Silent re-auth su 401 via iframe Google
//   • Banner "sessione scaduta" se il re-auth fallisce
//
// Dipende da: core.js (accessToken, CLIENT_ID, SCOPES, dbg)
// Chiamate runtime (non al caricamento) verso:
//   • syncLog()          — definita in store.js
//   • randomState()      — definita in auth.js
//   • getRedirectUri()   — definita in auth.js
// ═══════════════════════════════════════════════════════════════════

const BLIP_VER_API = '1'; // ← incrementa ad ogni modifica

// ═══════════════════════════════════════════════════════════════════
// TOKEN BUCKET — rate limiting preventivo (45 tok / 900ms)
// ═══════════════════════════════════════════════════════════════════
// Garantisce ≤ 67 chiamate/min per utente (quota Sheets: 100/min).
// Ogni apiFetch() deve prima acquisire un token — se il bucket è
// vuoto attende il ricaricamento senza bloccare altri task.
// ═══════════════════════════════════════════════════════════════════

let _tbTokens     = 45;        // token disponibili (bucket pieno all'avvio)
let _tbLastRefill = Date.now();
const _TB_CAPACITY  = 45;      // max token nel bucket
const _TB_REFILL_MS = 900;     // ms per token → ~67 token/min

async function _tbAcquire() {
  // Implementazione iterativa (non ricorsiva) — evita stack overflow
  // con molte chiamate in coda e lunghe attese
  while (true) {
    const now     = Date.now();
    const elapsed = now - _tbLastRefill;
    const gained  = Math.floor(elapsed / _TB_REFILL_MS);
    if (gained > 0) {
      _tbTokens    = Math.min(_TB_CAPACITY, _tbTokens + gained);
      _tbLastRefill = now - (elapsed % _TB_REFILL_MS);
    }
    if (_tbTokens > 0) {
      _tbTokens--;
      return; // token disponibile → procedi subito
    }
    // Bucket vuoto: aspetta il prossimo token
    const waitMs = _TB_REFILL_MS - (Date.now() - _tbLastRefill) + 20;
    // syncLog è definita in store.js — chiamata safe a runtime
    if (typeof syncLog === 'function') syncLog(`⏳ Rate limit preventivo — attesa ${Math.round(waitMs)}ms`, 'syn');
    await new Promise(r => setTimeout(r, waitMs));
    // loop → riprova senza ricorsione
  }
}

// ═══════════════════════════════════════════════════════════════════
// apiFetch — unico punto di uscita verso le Google API
// ═══════════════════════════════════════════════════════════════════
// Flusso:
//   1. Acquisisci token dal bucket (rate limiting)
//   2. Esegui fetch con Bearer token
//   3. 429 → svuota bucket + backoff esponenziale (max 3 retry)
//   4. 401 → prova silent re-auth via iframe
//            se riesce  → retry con nuovo token
//            se fallisce → mostra banner "sessione scaduta"
// ═══════════════════════════════════════════════════════════════════

async function apiFetch(url, options = {}, _attempt = 0) {
  if (!options.headers) options.headers = {};
  options.headers['Authorization'] = 'Bearer ' + accessToken;

  await _tbAcquire();

  let r;
  try {
    r = await fetch(url, options);
  } catch(e) {
    throw e; // AbortError, network error → passa al chiamante
  }

  // ── 429: safety net — retry con backoff esponenziale ──
  if (r.status === 429) {
    if (_attempt >= 3) {
      if (typeof syncLog === 'function') syncLog('❌ 429 quota esaurita dopo 3 retry — riprova tra 1 minuto', 'err');
      return r;
    }
    _tbTokens = 0; // svuota bucket
    const backoffMs = Math.min(32000, 5000 * Math.pow(2, _attempt)); // 5s, 10s, 20s
    if (typeof syncLog === 'function') syncLog(`⏱ 429 quota — retry ${_attempt + 1}/3 tra ${Math.round(backoffMs / 1000)}s`, 'wrn');
    await new Promise(res => setTimeout(res, backoffMs));
    const retryOpts = { ...options, headers: { ...options.headers } };
    delete retryOpts.signal; // AbortController potrebbe essere scaduto durante il backoff
    return apiFetch(url, retryOpts, _attempt + 1);
  }

  // ── 401 Sessione scaduta: silent re-auth ──
  if (r.status === 401) {
    const newToken = await trySilentReAuth();
    if (!newToken) {
      showSessionExpiredBanner();
      throw new Error('Sessione scaduta. Fai clic su "Riconnetti" per continuare.');
    }
    options.headers['Authorization'] = 'Bearer ' + newToken;
    return fetch(url, options);
  }

  return r;
}

// ═══════════════════════════════════════════════════════════════════
// SILENT RE-AUTH — refresh token via iframe nascosto
// ═══════════════════════════════════════════════════════════════════
// Google OAuth implicit flow non ha refresh token: all'expire (1h)
// proviamo un prompt=none flow in un iframe. Se l'utente è ancora
// loggato su Google, riceve un nuovo access_token silenziosamente.
// Se l'iframe fallisce (cookie scaduti, popup bloccato) → null.
// ═══════════════════════════════════════════════════════════════════

let _reAuthInProgress = false;

function trySilentReAuth() {
  // Se un re-auth è già in corso (da un'altra chiamata 401 concorrente)
  // aspettiamo il suo risultato invece di aprire un secondo iframe
  if (_reAuthInProgress) {
    return new Promise(res => {
      const poll = setInterval(() => {
        if (!_reAuthInProgress) { clearInterval(poll); res(accessToken); }
      }, 200);
      setTimeout(() => { clearInterval(poll); res(null); }, 10000);
    });
  }

  _reAuthInProgress = true;

  // randomState() e getRedirectUri() sono definite in auth.js (chiamata a runtime)
  return new Promise(resolve => {
    const state = randomState();
    const params = new URLSearchParams({
      client_id:     CLIENT_ID,
      redirect_uri:  getRedirectUri(),
      response_type: 'token',
      scope:         SCOPES,
      state,
      prompt:        'none',
      include_granted_scopes: 'true',
    });

    const iframe = document.createElement('iframe');
    iframe.style.cssText = 'display:none;width:1px;height:1px;position:fixed;top:-9999px';
    iframe.src = 'https://accounts.google.com/o/oauth2/v2/auth?' + params.toString();
    document.body.appendChild(iframe);

    const timeout = setTimeout(() => { cleanup(null); }, 8000);

    function cleanup(token) {
      clearTimeout(timeout);
      try { document.body.removeChild(iframe); } catch(e) {}
      _reAuthInProgress = false;
      if (token) { accessToken = token; hideSessionExpiredBanner(); }
      resolve(token);
    }

    iframe.onload = () => {
      try {
        const hash = iframe.contentWindow.location.hash;
        if (hash) {
          const p = new URLSearchParams(hash.slice(1));
          const token = p.get('access_token');
          if (token) { cleanup(token); return; }
        }
      } catch(e) {}
    };
  });
}

// ═══════════════════════════════════════════════════════════════════
// BANNER SESSIONE SCADUTA
// ═══════════════════════════════════════════════════════════════════

function showSessionExpiredBanner() {
  let b = document.getElementById('sessionExpiredBanner');
  if (!b) {
    b = document.createElement('div');
    b.id = 'sessionExpiredBanner';
    b.style.cssText = 'position:fixed;bottom:0;left:0;right:0;z-index:9999;background:#1a1a2e;color:#fff;padding:14px 20px;display:flex;align-items:center;justify-content:space-between;gap:12px;font-size:13px;box-shadow:0 -2px 12px rgba(0,0,0,.4)';
    b.innerHTML = '<span>⏰ <b>Sessione scaduta</b> — Il token Google è scaduto dopo 1 ora.</span>' +
      '<button onclick="startLogin()" style="background:#4285f4;color:#fff;border:none;padding:8px 18px;border-radius:6px;cursor:pointer;font-weight:600;white-space:nowrap">🔑 Riconnetti</button>';
    document.body.appendChild(b);
  }
  b.style.display = 'flex';
}

function hideSessionExpiredBanner() {
  const b = document.getElementById('sessionExpiredBanner');
  if (b) b.style.display = 'none';
}
