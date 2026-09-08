const CACHE_NAME = 'gpv-v4-layout-private-tasks';

const APP_FILES = [
  './',
  './index.html',
  './manifest.webmanifest',
  './gpv-icon.png',
  './gpv-icon-192.png',
  './gpv-icon-512.png'
];

self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(CACHE_NAME).then(cache => cache.addAll(APP_FILES))
  );
  self.skipWaiting();
});

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

  self.clients.claim();
});

self.addEventListener('fetch', event => {
  if (event.request.method !== 'GET') return;

  event.respondWith(
    fetch(
      event.request.mode === 'navigate'
        ? new Request(event.request, { cache: 'reload' })
        : event.request
    )
      .then(response => {
        const copy = response.clone();

        caches.open(CACHE_NAME).then(cache => {
          cache.put(event.request, copy);
        });

        return response;
      })
      .catch(() =>
        caches.match(event.request).then(cached => {
          return cached || caches.match('./index.html');
        })
      )
  );
});
