const CACHE_NAME = 'lookup-export-cache-v1';
const OFFLINE_URLS = [
  '/', // root
  'manifest.webmanifest',
  'Index', // HTML page
  'https://ssl.gstatic.com/docs/doclist/images/drive_icon_128.png'
];

self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(CACHE_NAME).then(cache => {
      return cache.addAll(OFFLINE_URLS);
    })
  );
  self.skipWaiting();
});

self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys().then(keys => {
      return Promise.all(keys.map(key => {
        if (key !== CACHE_NAME) return caches.delete(key);
      }));
    })
  );
  self.clients.claim();
});

self.addEventListener('fetch', event => {
  if (event.request.method !== 'GET') return;

  event.respondWith(
    caches.match(event.request).then(cached => {
      return cached || fetch(event.request).catch(() => {
        // Optional: return a fallback offline page
        return new Response('<h1>Offline</h1><p>Internet required for data operations.</p>', {
          headers: { 'Content-Type': 'text/html' }
        });
      });
    })
  );
});
