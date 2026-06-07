const CACHE = 'at-v5';
const CDN_HOSTS = ['cdn.jsdelivr.net'];
const OFFLINE_PAGE = '/static/offline.html';

// App-Shell-Seiten: nach erstem Laden gecacht, dann offline verfügbar
const SHELL_PATHS = ['/aktivitaet/neu', '/aktivitaeten', '/dashboard'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll([
    OFFLINE_PAGE,  // Fallback-Seite immer vorhalten
    '/static/offline.js',
    'https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/css/bootstrap.min.css',
    'https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.css',
    'https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/js/bootstrap.bundle.min.js',
    'https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js',
  ]).catch(() => {})));
  self.skipWaiting();
});

self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(keys =>
    Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k)))
  ));
  self.clients.claim();
});

self.addEventListener('fetch', e => {
  if (e.request.method !== 'GET') return;
  const url = new URL(e.request.url);

  // CDN-Assets: Cache-First (blitzschnell, offline verfügbar)
  if (CDN_HOSTS.some(h => url.hostname.includes(h))) {
    e.respondWith(
      caches.match(e.request).then(r => r || fetch(e.request).then(res => {
        caches.open(CACHE).then(c => c.put(e.request, res.clone()));
        return res;
      }))
    );
    return;
  }

  // Static-Files (offline.js, manifest.json, icons)
  if (url.pathname.startsWith('/static/')) {
    e.respondWith(
      caches.match(e.request).then(r => r || fetch(e.request).then(res => {
        if (res.ok) caches.open(CACHE).then(c => c.put(e.request, res.clone()));
        return res;
      }).catch(() => caches.match(e.request)))
    );
    return;
  }

  // App-Shell-Seiten: Network-First, Cache als Fallback, Offline-Seite als letzter Ausweg
  if (SHELL_PATHS.some(p => url.pathname === p)) {
    e.respondWith(
      fetch(e.request)
        .then(res => {
          if (res.ok) caches.open(CACHE).then(c => c.put(e.request, res.clone()));
          return res;
        })
        .catch(() => caches.match(e.request).then(r => r || caches.match(OFFLINE_PAGE)))
    );
    return;
  }

  // Alle anderen Navigations-Requests (z.B. /login, /admin, unbekannte Seiten)
  // → Offline-Fallback wenn weder Netz noch Cache verfügbar
  if (e.request.mode === 'navigate') {
    e.respondWith(
      fetch(e.request)
        .catch(() => caches.match(e.request).then(r => r || caches.match(OFFLINE_PAGE)))
    );
    return;
  }
});
