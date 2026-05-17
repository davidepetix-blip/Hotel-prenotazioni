// ═══════════════════════════════════════════════════════════════════
// store-patch.js — Patch per store.js
// Caricato DOPO store.js — sovrascrive solo le funzioni corrette.
// Dimensione minima per deploy mobile.
//
// Fix inclusi:
//   1. findMatch Priority 4 — overlap multi-mese + camera normalizzata
//   2. esportaLogSessione — mostra BLIP_VER_STORE/ROOMS/API/AUTH
//   3. BLIP_VER_STORE aggiornato a '2'
// ═══════════════════════════════════════════════════════════════════

// Aggiorna versione
if (typeof window !== 'undefined') window._STORE_PATCH = '4';
const BLIP_VER_STORE = '4'; // fix: findMatch P1 BLIP_ID+camera per gruppi multi-stanza

// Override findMatch con Priority 4 corretta
function findMatch(target, list) {
  // PRIORITÀ 1: BLIP_ID + camera (gruppi multi-camera: stesso ID su più stanze)
  if (target.dbId) {
    const camT1 = (target.cameraName || roomName(target.r) || '').toLowerCase().trim();
    const byIdCam = list.find(b =>
      b.dbId === target.dbId &&
      (b.cameraName||roomName(b.r)||'').toLowerCase().trim() === camT1
    );
    if (byIdCam) return byIdCam;
    const byId = list.find(b => b.dbId === target.dbId);
    if (byId) return byId;
  }

  const camT  = (target.cameraName || roomName(target.r) || '').toLowerCase().trim();
  const nomT  = _normName(target.n);
  const dayT  = Math.round((target.s?.getTime?.() || 0) / DAY_MS);
  const dispT = (target.d || '').trim().toLowerCase();

  // PRIORITÀ 2: nome + camera + data esatta
  let m = list.find(b => {
    if (_normName(b.n) !== nomT) return false;
    if ((b.cameraName||roomName(b.r)||'').toLowerCase().trim() !== camT) return false;
    return Math.round((b.s?.getTime?.() || 0) / DAY_MS) === dayT;
  });
  if (m) return m;

  // PRIORITÀ 3: fuzzy ±1 giorno
  m = list.find(b => {
    if (_normName(b.n) !== nomT) return false;
    if ((b.cameraName||roomName(b.r)||'').toLowerCase().trim() !== camT) return false;
    if (dispT && (b.d||'').trim().toLowerCase() !== dispT) return false;
    return Math.abs(Math.round((b.s?.getTime?.() || 0) / DAY_MS) - dayT) <= 1;
  });
  if (m) return m;

  // PRIORITÀ 4: overlap multi-mese — il frammento ha data inizio = 1° del mese
  // ma il record DB ha la data di inizio reale (es. 15 aprile).
  // Guard disposizione: se entrambi hanno disp. esplicita diversa → booking distinti.
  const sT = target.s?.getTime?.() || 0;
  const eT = target.e?.getTime?.() || 0;
  const yT = target.s ? new Date(target.s).getFullYear() : 0;
  const dT = (target.d || '').trim().toLowerCase();

  // Normalizza camera: "Camera 1"→"1", "cam.1"→"1", "01"→"1"
  const _nc = c => String(c||'').toLowerCase().trim()
    .replace(/^camera\s+/, '').replace(/^cam\.\s*/, '').replace(/^0+(?=\d)/, '').trim();
  const camTN = _nc(camT);

  m = list.find(b => {
    if (_normName(b.n) !== nomT) return false;
    if (_nc(b.cameraName || roomName(b.r) || '') !== camTN) return false;
    const dDb = (b.d || '').trim().toLowerCase();
    if (dDb && dT && dDb !== dT) return false;
    const sDb = b.s?.getTime?.() || 0;
    const eDb = b.e?.getTime?.() || 0;
    const yDb = b.s ? new Date(b.s).getFullYear() : 0;
    if (yDb !== yT) return false;
    // ±1 giorno per gestire frammenti adiacenti a cambio mese
    return sDb < eT + DAY_MS && eDb + DAY_MS >= sT;
  });
  return m || null;
}
