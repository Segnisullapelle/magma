/* ================================================================
   MagMa — Service Worker
   Strategia: Cache-first per asset statici, Network-first per API
================================================================ */

const CACHE_NAME    = 'magma-v2';
const CACHE_CDN     = 'magma-cdn-v2';

const STATIC_ASSETS = [
  './',
  './index.html',
  './style.css',
  './app.js',
  './manifest.json',
  './icon.svg',
];

const CDN_URL_PATTERN = 'unpkg.com/@zxing';


/* ── Installazione: pre-cache degli asset statici ─────────── */
self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then(cache => cache.addAll(STATIC_ASSETS))
      .then(() => self.skipWaiting())
  );
});


/* ── Attivazione: rimozione cache vecchie ────────────────── */
self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys().then(keys =>
      Promise.all(
        keys
          .filter(k => k !== CACHE_NAME && k !== CACHE_CDN)
          .map(k => caches.delete(k))
      )
    ).then(() => self.clients.claim())
  );
});


/* ── Fetch: intercetta le richieste ─────────────────────── */
self.addEventListener('fetch', event => {
  const { request } = event;
  const url = new URL(request.url);

  // Ignora richieste non-GET e non-HTTP
  if (request.method !== 'GET') return;
  if (!url.protocol.startsWith('http')) return;

  // API Google Apps Script → Network-first, con fallback offline
  if (url.hostname.includes('script.google.com')) {
    event.respondWith(networkFirst(request, null));
    return;
  }

  // Libreria ZXing da CDN → Cache-first, aggiorna in background
  if (url.href.includes(CDN_URL_PATTERN)) {
    event.respondWith(cacheFirst(request, CACHE_CDN));
    return;
  }

  // Asset statici locali → Cache-first
  event.respondWith(cacheFirst(request, CACHE_NAME));
});


/* ── Strategie di caching ──────────────────────────────── */

async function cacheFirst(request, cacheName) {
  const cached = await caches.match(request);
  if (cached) return cached;

  try {
    const response = await fetch(request);
    if (response.ok && cacheName) {
      const cache = await caches.open(cacheName);
      cache.put(request, response.clone());
    }
    return response;
  } catch (_) {
    return offlineFallback(request);
  }
}

async function networkFirst(request) {
  try {
    return await fetch(request);
  } catch (_) {
    return new Response(
      JSON.stringify({
        success: false,
        message: 'Offline: impossibile raggiungere il server. Verifica la connessione.',
      }),
      { headers: { 'Content-Type': 'application/json' } }
    );
  }
}

function offlineFallback(request) {
  if (request.headers.get('Accept')?.includes('text/html')) {
    return caches.match('./index.html');
  }
  return new Response('', { status: 503, statusText: 'Service Unavailable' });
}
