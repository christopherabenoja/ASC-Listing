const CACHE_NAME = 'scannables-cache-v3';
const CORE_ASSETS = [
  './',
  './index.html',
  './app.js',
  './manifest.json',
  './icons/icon-192.png',
  './icons/icon-512.png',
  './vendor/xlsx.full.min.js',
  './vendor/JsBarcode.all.min.js',
  './vendor/jspdf.umd.min.js',
  './vendor/html2canvas.min.js',
  './vendor/jszip.min.js'
];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(CORE_ASSETS))
  );
  self.skipWaiting();
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(keys.filter((k) => k !== CACHE_NAME).map((k) => caches.delete(k)))
    )
  );
  self.clients.claim();
});

// Cache-first for the app shell, network-first for CDN libraries (so they stay fresh
// when online, but still work offline once cached once).
self.addEventListener('fetch', (event) => {
  const url = new URL(event.request.url);
  const isCore = CORE_ASSETS.some((a) => event.request.url.endsWith(a.replace('./', '')));

  if (isCore) {
    event.respondWith(
      caches.match(event.request).then((cached) => cached || fetch(event.request))
    );
    return;
  }

  event.respondWith(
    fetch(event.request)
      .then((response) => {
        const clone = response.clone();
        caches.open(CACHE_NAME).then((cache) => cache.put(event.request, clone));
        return response;
      })
      .catch(() => caches.match(event.request))
  );
});
