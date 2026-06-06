/* ═══════════════════════════════════════════════════════════
   RAPSOSYN – Service Worker
   Cache offline + Background Sync
   ═══════════════════════════════════════════════════════════ */

const CACHE_NAME = 'rapsosyn-v1';
const SYNC_TAG   = 'rapsosyn-sync';

// Fichiers à mettre en cache pour le mode hors-ligne
const CACHE_FILES = [
  '/',
  '/index.html',
  '/canevas_temp.html',
  '/manifest.json',
  // Polices Google (si elles ont été chargées une première fois)
  'https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;500;600;700;800&family=JetBrains+Mono:wght@400;500&display=swap',
  // Firebase SDK
  'https://www.gstatic.com/firebasejs/9.22.0/firebase-app-compat.js',
  'https://www.gstatic.com/firebasejs/9.22.0/firebase-database-compat.js',
  'https://www.gstatic.com/firebasejs/9.22.0/firebase-auth-compat.js',
  // ExcelJS
  'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js'
];

/* ── Installation : mise en cache initiale ── */
self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(CACHE_NAME).then(cache => {
      // Cache les fichiers locaux en priorité, les CDN en best-effort
      return cache.addAll(['/', '/index.html', '/canevas_temp.html', '/manifest.json'])
        .then(() => {
          // CDN en best-effort (pas bloquant)
          const cdnFiles = CACHE_FILES.filter(f => f.startsWith('http'));
          return Promise.allSettled(cdnFiles.map(url =>
            cache.add(url).catch(() => {}) // silencieux si CDN indisponible
          ));
        });
    }).then(() => self.skipWaiting())
  );
});

/* ── Activation : nettoyage des anciens caches ── */
self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.filter(k => k !== CACHE_NAME).map(k => caches.delete(k)))
    ).then(() => self.clients.claim())
  );
});

/* ── Fetch : Cache-first pour les assets, Network-first pour les données ── */
self.addEventListener('fetch', event => {
  const url = new URL(event.request.url);

  // Ne pas intercepter les requêtes Firebase (temps réel)
  if (url.hostname.includes('firebasedatabase.app') ||
      url.hostname.includes('googleapis.com') && url.pathname.includes('/identitytoolkit')) {
    return;
  }

  // Network-first pour les pages HTML (toujours fraîches si possible)
  if (event.request.mode === 'navigate' ||
      event.request.headers.get('accept')?.includes('text/html')) {
    event.respondWith(
      fetch(event.request)
        .then(response => {
          const clone = response.clone();
          caches.open(CACHE_NAME).then(cache => cache.put(event.request, clone));
          return response;
        })
        .catch(() => caches.match(event.request).then(r => r || caches.match('/index.html')))
    );
    return;
  }

  // Cache-first pour les autres ressources (JS, CSS, fonts)
  event.respondWith(
    caches.match(event.request).then(cached => {
      if (cached) return cached;
      return fetch(event.request).then(response => {
        if (response && response.status === 200) {
          const clone = response.clone();
          caches.open(CACHE_NAME).then(cache => cache.put(event.request, clone));
        }
        return response;
      }).catch(() => cached);
    })
  );
});

/* ── Background Sync : déclenché au retour de connexion ── */
self.addEventListener('sync', event => {
  if (event.tag === SYNC_TAG) {
    event.waitUntil(notifyClientsToSync());
  }
});

async function notifyClientsToSync() {
  const clients = await self.clients.matchAll({ type: 'window' });
  clients.forEach(client => {
    client.postMessage({ type: 'sw_sync_trigger' });
  });
}

/* ── Message handler : communication avec les pages ── */
self.addEventListener('message', event => {
  if (event.data?.type === 'sw_skip_waiting') {
    self.skipWaiting();
  }
  if (event.data?.type === 'sw_cache_canevas') {
    // Mettre en cache une URL de canevas dynamique
    const url = event.data.url;
    if (url) {
      caches.open(CACHE_NAME).then(cache => cache.add(url).catch(() => {}));
    }
  }
});
