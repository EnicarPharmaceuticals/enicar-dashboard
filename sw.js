/* Enicar Dashboard service worker.
   Strategy: NETWORK-FIRST for everything. The dashboard is a live report that
   rebuilds all day, so a cached copy must never shadow a fresh one; the cache
   exists only so the app still opens (with the last good data) in poor signal.
   Bump CACHE_VERSION to force old caches out. */
const CACHE_VERSION = 'enicar-v1';

self.addEventListener('install', (e) => self.skipWaiting());

self.addEventListener('activate', (e) => {
  e.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.filter(k => k !== CACHE_VERSION).map(k => caches.delete(k)))
    ).then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', (e) => {
  if (e.request.method !== 'GET') return;
  e.respondWith(
    fetch(e.request).then(resp => {
      if (resp && resp.ok) {
        const copy = resp.clone();
        caches.open(CACHE_VERSION).then(c => c.put(e.request, copy));
      }
      return resp;
    }).catch(() =>
      caches.match(e.request).then(hit => hit || caches.match('./index.html'))
    )
  );
});
