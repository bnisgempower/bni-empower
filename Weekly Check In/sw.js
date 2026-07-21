// BNI Empower — Service Worker
// Caches the app shell so it loads instantly (and works offline to fill in,
// though submitting still needs a connection).

const CACHE_NAME = 'bni-empower-v1';

// Files to cache on install
const SHELL = [
  '/',
  '/index.html',
  '/attendance.html',
  '/icon.svg',
  '/manifest.json',
];

// ── Install — cache the app shell ────────────────────
self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(CACHE_NAME).then(cache => cache.addAll(SHELL))
  );
  self.skipWaiting(); // activate immediately
});

// ── Activate — remove old caches ─────────────────────
self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys().then(keys =>
      Promise.all(
        keys
          .filter(key => key !== CACHE_NAME)
          .map(key => caches.delete(key))
      )
    )
  );
  self.clients.claim(); // take control of open tabs immediately
});

// ── Fetch — cache-first for shell, network-first for API ─
self.addEventListener('fetch', event => {
  const url = new URL(event.request.url);

  // Always go network for the Apps Script endpoint
  if (url.hostname.includes('script.google.com')) {
    event.respondWith(fetch(event.request));
    return;
  }

  // Cache-first for everything else (fonts, qrcode CDN, app shell)
  event.respondWith(
    caches.match(event.request).then(cached => {
      return cached || fetch(event.request).then(response => {
        // Cache successful GET responses
        if (event.request.method === 'GET' && response.status === 200) {
          const clone = response.clone();
          caches.open(CACHE_NAME).then(cache => cache.put(event.request, clone));
        }
        return response;
      });
    }).catch(() => {
      // Offline fallback: serve the cached home page for navigation requests
      if (event.request.mode === 'navigate') {
        return caches.match('/') || caches.match('/index.html');
      }
    })
  );
});
