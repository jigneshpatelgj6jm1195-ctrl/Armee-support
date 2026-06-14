/* ═══════════════════════════════════════════════════════
   Service Worker — School Complaint Form PWA
   Caches: HTML form, school_data.json, fonts
   Offline queue: submissions stored in IndexedDB
═══════════════════════════════════════════════════════ */

const CACHE_NAME = 'school-form-v29';
const STATIC_ASSETS = [
  './index.html',
  './school_data.json',
  './manifest.json',
  './html5-qrcode.min.js',
  './armee_logo.png',
  './armee_logo_square.png',
];

// ── INSTALL: cache static assets ──
self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(CACHE_NAME).then(cache => {
      return cache.addAll(STATIC_ASSETS).catch(err => {
        console.warn('SW: Some assets failed to cache:', err);
      });
    })
  );
  self.skipWaiting();
});

// ── ACTIVATE: clean old caches ──
self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.filter(k => k !== CACHE_NAME).map(k => caches.delete(k)))
    )
  );
  self.clients.claim();
});

// ── FETCH: cache-first for assets, network-first for API ──
self.addEventListener('fetch', event => {
  const url = new URL(event.request.url);

  // Bypass service worker entirely for admin panel and Google Apps Script GET/POST calls
  if (url.pathname.includes('admin') || url.hostname.includes('script.google.com')) {
    return;
  }

  // school_data.json — cache first (it's big, rarely changes)
  if (event.request.url.includes('school_data.json')) {
    event.respondWith(
      caches.match(event.request).then(cached => {
        if (cached) return cached;
        return fetch(event.request).then(response => {
          const clone = response.clone();
          caches.open(CACHE_NAME).then(cache => cache.put(event.request, clone));
          return response;
        });
      })
    );
    return;
  }

  // Static assets — cache first
  event.respondWith(
    caches.match(event.request).then(cached => cached || fetch(event.request))
  );
});

// ── BACKGROUND SYNC: retry queued submissions ──
self.addEventListener('sync', event => {
  if (event.tag === 'submit-complaint') {
    event.waitUntil(retryQueuedSubmissions());
  }
});

async function retryQueuedSubmissions() {
  // Notify all clients to retry
  const clients = await self.clients.matchAll();
  clients.forEach(client => client.postMessage({ type: 'RETRY_QUEUE' }));
}
