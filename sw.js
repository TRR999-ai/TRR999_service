// TRR999 Service Worker
// HTML = network-first (always get latest build when online, fall back to cache offline)
// Static assets = cache-first (fast load, refreshed when CACHE version bumps)
const CACHE = 'trr999-v3';

self.addEventListener('install', e => self.skipWaiting());

self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys()
      .then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET') return;

  const isHTML = req.mode === 'navigate' || req.destination === 'document';

  if (isHTML) {
    // Network-first — fresh HTML wins, cache is only a fallback when offline
    e.respondWith(
      fetch(req)
        .then(resp => {
          const clone = resp.clone();
          caches.open(CACHE).then(c => c.put(req, clone));
          return resp;
        })
        .catch(() => caches.match(req).then(r => r || caches.match('./')))
    );
  } else {
    // Cache-first for static assets (logo, fonts, SDK)
    e.respondWith(
      caches.match(req).then(cached =>
        cached || fetch(req).then(resp => {
          const clone = resp.clone();
          caches.open(CACHE).then(c => c.put(req, clone));
          return resp;
        })
      )
    );
  }
});
