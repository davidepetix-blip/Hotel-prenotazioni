// ═══════════════════════════════════════════════════════════════════
// store-patch.js — Patch per store.js v5
// Caricato DOPO store.js — sovrascrive findMatch.
//
// NON usa `const` per variabili già dichiarate in store.js.
// `const` ridichiarata → SyntaxError → file ignorato silenziosamente.
// ═══════════════════════════════════════════════════════════════════

window._blipStorePatch = '5';
try { BLIP_VER_STORE = '5-patch'; } catch(e) {}

function findMatch(target, list) {
  // P1: BLIP_ID + camera (gruppi multi-camera)
  if (target.dbId) {
    const c1 = (target.cameraName || roomName(target.r) || '').toLowerCase().trim();
    const m1 = list.find(b =>
      b.dbId === target.dbId &&
      (b.cameraName || roomName(b.r) || '').toLowerCase().trim() === c1
    );
    if (m1) return m1;
    const m1b = list.find(b => b.dbId === target.dbId);
    if (m1b) return m1b;
  }

  const camT  = (target.cameraName || roomName(target.r) || '').toLowerCase().trim();
  const nomT  = _normName(target.n);
  const dayT  = Math.round((target.s?.getTime?.() || 0) / DAY_MS);
  const dispT = (target.d || '').trim().toLowerCase();

  // P2: nome + camera + data esatta
  let m = list.find(b => {
    if (_normName(b.n) !== nomT) return false;
    if ((b.cameraName || roomName(b.r) || '').toLowerCase().trim() !== camT) return false;
    return Math.round((b.s?.getTime?.() || 0) / DAY_MS) === dayT;
  });
  if (m) return m;

  // P3: fuzzy ±1 giorno
  m = list.find(b => {
    if (_normName(b.n) !== nomT) return false;
    if ((b.cameraName || roomName(b.r) || '').toLowerCase().trim() !== camT) return false;
    if (dispT && (b.d || '').trim().toLowerCase() !== dispT) return false;
    return Math.abs(Math.round((b.s?.getTime?.() || 0) / DAY_MS) - dayT) <= 1;
  });
  if (m) return m;

  // P4: overlap multi-mese ±1 giorno + normalizzazione camera
  const sT = target.s?.getTime?.() || 0;
  const eT = target.e?.getTime?.() || 0;
  const yT = target.s ? new Date(target.s).getFullYear() : 0;
  const dT = (target.d || '').trim().toLowerCase();
  const nc = c => String(c || '').toLowerCase().trim()
    .replace(/^camera\s+/, '').replace(/^cam\.\s*/, '').replace(/^0+(?=\d)/, '').trim();
  const camTN = nc(camT);

  m = list.find(b => {
    if (_normName(b.n) !== nomT) return false;
    if (nc(b.cameraName || roomName(b.r) || '') !== camTN) return false;
    const dDb = (b.d || '').trim().toLowerCase();
    if (dDb && dT && dDb !== dT) return false;
    const sDb = b.s?.getTime?.() || 0;
    const eDb = b.e?.getTime?.() || 0;
    if (b.s && new Date(b.s).getFullYear() !== yT) return false;
    return sDb < eT + DAY_MS && eDb + DAY_MS >= sT;
  });
  return m || null;
}

console.log('[Blip] store-patch v5 — findMatch P1+P4 + no const redecl');
