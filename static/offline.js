/* Aktions Tracker – Offline Queue
   IndexedDB-basierte Warteschlange für Besuche ohne Internetverbindung.
   Wird in base.html geladen und von neue_aktivitaet.html genutzt. */

const _DB_NAME  = 'at-offline-v1';
const _DB_STORE = 'pending';

function _dbOpen() {
  return new Promise((res, rej) => {
    const r = indexedDB.open(_DB_NAME, 1);
    r.onupgradeneeded = e => e.target.result.createObjectStore(_DB_STORE, { keyPath: 'id', autoIncrement: true });
    r.onsuccess = e => res(e.target.result);
    r.onerror   = e => rej(e.target.error);
  });
}

async function _dbSave(entry) {
  const db = await _dbOpen();
  return new Promise((res, rej) => {
    const r = db.transaction(_DB_STORE, 'readwrite').objectStore(_DB_STORE).add(entry);
    r.onsuccess = () => res(r.result);
    r.onerror   = () => rej(r.error);
  });
}

async function _dbGetAll() {
  const db = await _dbOpen();
  return new Promise((res, rej) => {
    const r = db.transaction(_DB_STORE, 'readonly').objectStore(_DB_STORE).getAll();
    r.onsuccess = () => res(r.result);
    r.onerror   = () => rej(r.error);
  });
}

async function _dbDelete(id) {
  const db = await _dbOpen();
  return new Promise((res, rej) => {
    const r = db.transaction(_DB_STORE, 'readwrite').objectStore(_DB_STORE).delete(id);
    r.onsuccess = () => res();
    r.onerror   = () => rej(r.error);
  });
}

async function _dbCount() {
  const db = await _dbOpen();
  return new Promise((res, rej) => {
    const r = db.transaction(_DB_STORE, 'readonly').objectStore(_DB_STORE).count();
    r.onsuccess = () => res(r.result);
    r.onerror   = () => rej(r.error);
  });
}

// Foto komprimieren: max 1600px, JPEG 80%
function _compressImage(file) {
  return new Promise(res => {
    const reader = new FileReader();
    reader.onload = e => {
      const img = new Image();
      img.onload = () => {
        const MAX = 1600;
        let w = img.width, h = img.height;
        if (w > MAX || h > MAX) {
          if (w > h) { h = Math.round(h * MAX / w); w = MAX; }
          else        { w = Math.round(w * MAX / h); h = MAX; }
        }
        const canvas = document.createElement('canvas');
        canvas.width = w; canvas.height = h;
        canvas.getContext('2d').drawImage(img, 0, 0, w, h);
        res(canvas.toDataURL('image/jpeg', 0.80));
      };
      img.src = e.target.result;
    };
    reader.readAsDataURL(file);
  });
}

// Banner mit Anzahl ausstehender Besuche aktualisieren
async function _updateBanner() {
  const count  = await _dbCount();
  const banner = document.getElementById('offlineSyncBanner');
  const badge  = document.getElementById('offlinePendingCount');
  if (!banner) return;
  if (count > 0) {
    if (badge) badge.textContent = count === 1 ? '1 Besuch' : count + ' Besuche';
    banner.classList.remove('d-none');
  } else {
    banner.classList.add('d-none');
  }
}

// Kleiner Alert am Seitenanfang
function _showAlert(html, type) {
  const el = document.createElement('div');
  el.className = 'alert alert-' + (type || 'success') + ' alert-dismissible fade show';
  el.innerHTML = html + '<button type="button" class="btn-close" data-bs-dismiss="alert"></button>';
  const c = document.querySelector('.container-fluid');
  if (c) c.insertBefore(el, c.firstChild.nextSibling || c.firstChild);
}

// Offline-Punkt in der Navbar aktualisieren
function _updateDot() {
  const dot = document.getElementById('offlineDot');
  if (!dot) return;
  dot.style.display = navigator.onLine ? 'none' : 'inline-flex';
}

// Alle ausstehenden Besuche zum Server hochladen
async function oqSync() {
  if (!navigator.onLine) return;
  const items = await _dbGetAll();
  if (!items.length) return;

  const btn = document.getElementById('btnOqSync');
  if (btn) { btn.disabled = true; btn.innerHTML = '<span class="spinner-border spinner-border-sm me-1"></span>Synchronisiere…'; }

  let ok = 0, fail = 0;
  for (const item of items) {
    try {
      const r = await fetch('/api/aktivitaet/offline-sync', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(item.data),
        credentials: 'same-origin',
      });
      if (!r.ok) throw new Error('HTTP ' + r.status);
      const j = await r.json();
      if (j.ok) { await _dbDelete(item.id); ok++; }
      else fail++;
    } catch { fail++; }
  }

  if (btn) { btn.disabled = false; btn.innerHTML = '<i class="bi bi-cloud-upload me-1"></i>Jetzt synchronisieren'; }
  await _updateBanner();

  if (ok > 0)   _showAlert('<i class="bi bi-check-circle me-2"></i><strong>' + ok + (ok === 1 ? ' Besuch' : ' Besuche') + ' synchronisiert!</strong> Die Daten sind jetzt auf dem Server gespeichert.');
  if (fail > 0) _showAlert('<i class="bi bi-exclamation-triangle me-2"></i>' + fail + (fail === 1 ? ' Besuch konnte' : ' Besuche konnten') + ' nicht übertragen werden. Bitte erneut versuchen.', 'warning');
}

// Initialisierung beim Seitenaufruf (inkl. bfcache-Restore)
async function _init() {
  await _updateBanner();
  _updateDot();
  if (navigator.onLine) {
    const c = await _dbCount();
    if (c > 0) oqSync();
  }
}
document.addEventListener('DOMContentLoaded', _init);
window.addEventListener('pageshow', e => { if (e.persisted) _init(); }); // bfcache

window.addEventListener('online',  () => { _updateDot(); setTimeout(oqSync, 1500); });
window.addEventListener('offline', _updateDot);

// Warteschlange komplett leeren (fehlgeschlagene / veraltete Einträge verwerfen)
async function _dbClear() {
  const db = await _dbOpen();
  return new Promise((res, rej) => {
    const r = db.transaction(_DB_STORE, 'readwrite').objectStore(_DB_STORE).clear();
    r.onsuccess = () => res();
    r.onerror   = () => rej(r.error);
  });
}

async function oqDiscard() {
  const count = await _dbCount();
  const label = count === 1 ? '1 offline gespeicherten Besuch' : count + ' offline gespeicherte Besuche';
  if (!confirm(`${label} endgültig verwerfen? Die Daten werden nicht auf den Server übertragen.`)) return;
  await _dbClear();
  await _updateBanner();
}

// Globale API für andere Skripte
window.OQ = {
  save:          _dbSave,
  compressImage: _compressImage,
  updateBanner:  _updateBanner,
  sync:          oqSync,
  discard:       oqDiscard,
};

// Direkte globale Referenzen für Inline-Onclick-Handler
window._dbClear    = _dbClear;
window._updateBanner = _updateBanner;
