/* Simple offline-first service worker for Auto Report */
-2025-10-18d';
const PRECACHE = `autoreport-precache-${VERSION}`;
const RUNTIME = `autoreport-runtime-${VERSION}`;

// Files we know we need
const 
  './',
  './index.html',
  './manifest.webmanifest',
  './assets/ph/style.css',
  './assets/ph/ph.js',
  './assets/config.js',
  './assets/pwa.js',
  './assets/logo_nps.svg',
  'https://cdn.jsdelivr.net/npm/html2pdf.js@0.10.1/dist/html2pdf.bundle.min.js'
;

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(PRECACHE).then((cache) => cache.addAll(PRECACHE_URLS)).then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((keys) => Promise.all(keys.map((k) => { if (![PRECACHE, RUNTIME].includes(k)) return caches.delete(k); }))).then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', (event) => {
  const req = event.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);
  // Navigation requests: network first, fallback to cached index
  if (req.mode === 'navigate') {
    event.respondWith(
      fetch(req).catch(() => caches.match('./index.html'))
    );
    return;
  }
  // For others: cache-first, then network
  event.respondWith(
    caches.match(req).then((cached) => {
      if (cached) return cached;
      return fetch(req).then((res) => {
        const copy = res.clone();
        caches.open(RUNTIME).then((cache) => cache.put(req, copy)).catch(() => {});
        return res;
      }).catch(() => cached);
    })
  );
});


