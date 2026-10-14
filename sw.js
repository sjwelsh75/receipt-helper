// Offline support: app files come from the network when online, cache when offline.
const CACHE = 'receipt-helper-v5';
const FILES = [
  './', 'index.html', 'styles.css', 'app.js', 'xlsx.js', 'manifest.webmanifest',
  'icon-180.png', 'icon-192.png', 'icon-512.png',
  'https://cdn.jsdelivr.net/npm/jszip@3.10.1/dist/jszip.min.js',
  'https://cdn.jsdelivr.net/npm/pdf-lib@1.17.1/dist/pdf-lib.min.js',
];
self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(FILES)).then(() => self.skipWaiting()));
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys()
    .then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k))))
    .then(() => self.clients.claim()));
});
self.addEventListener('fetch', e => {
  if (e.request.method !== 'GET') return;
  const sameOrigin = new URL(e.request.url).origin === self.location.origin;
  e.respondWith(caches.open(CACHE).then(async cache => {
    const cached = await cache.match(e.request, { ignoreSearch: true });
    // Libraries never change: use the cached copy. App files: get the latest when online.
    if (cached && !sameOrigin) return cached;
    try {
      const res = await fetch(e.request);
      if (res.ok) cache.put(e.request, res.clone());
      return res;
    } catch (err) {
      if (cached) return cached;
      throw err;
    }
  }));
});
