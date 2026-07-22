// BNI Empower — Member Hub Service Worker
// Caches the app shell so the hub loads instantly. The presenter roster is
// always fetched fresh from the Apps Script endpoint (never cached).

const CACHE_NAME = 'bni-empower-hub-v3';

const SHELL = [
  './',
  './index.html',
  './manifest.json',
  './icon-192.png',
  './apple-touch-icon.png',
];

// ── Install — cache the app shell ────────────────────
self.addEventListener('install', event => {
  event.waitUntil(caches.open(CACHE_NAME).then(cache => cache.addAll(SHELL)));
  self.skipWaiting();
});

// ── Activate — remove old caches ─────────────────────
self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.filter(k => k !== CACHE_NAME).map(k => caches.delete(k)))
    )
  );
  self.clients.claim();
});

// ── Fetch — network for API, cache-first for shell ───
self.addEventListener('fetch', event => {
  const url = new URL(event.request.url);

  // Always go to the network for the Apps Script endpoint (live roster data)
  if (url.hostname.includes('script.google.com')) {
    event.respondWith(fetch(event.request));
    return;
  }

  // Network-first for the page itself, so design/content updates show up
  // immediately when online (fall back to cache only when offline).
  if (event.request.mode === 'navigate') {
    event.respondWith(
      fetch(event.request).then(response => {
        const clone = response.clone();
        caches.open(CACHE_NAME).then(cache => cache.put(event.request, clone));
        return response;
      }).catch(() => caches.match(event.request).then(c => c || caches.match('./index.html')))
    );
    return;
  }

  // Cache-first for static assets (icons, manifest, fonts)
  event.respondWith(
    caches.match(event.request).then(cached =>
      cached || fetch(event.request).then(response => {
        if (event.request.method === 'GET' && response.status === 200) {
          const clone = response.clone();
          caches.open(CACHE_NAME).then(cache => cache.put(event.request, clone));
        }
        return response;
      })
    )
  );
});
