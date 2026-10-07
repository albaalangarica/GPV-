// Solo se guardan en el dispositivo los archivos de la propia app.
// Las peticiones a Google (hojas, Apps Script) no pasan por aquí: antes se copiaban
// todas y el almacenamiento del móvil crecía sin límite.
const CACHE_NAME = 'gpv-v9-app-shell';

const APP_FILES = [
  './',
  './index.html',
  './manifest.webmanifest',
  './gpv-icon.png',
  './gpv-icon-192.png',
  './gpv-icon-512.png',
  './favicon.png',
  './gpv-icon-maskable.png',
  './logo-ppvasco-blanco.png'
];

self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(CACHE_NAME).then(cache => cache.addAll(APP_FILES))
  );
  self.skipWaiting();
});

self.addEventListener('activate', event => {
  // Borra las cachés antiguas, incluidas las llenas de CSV de versiones anteriores.
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
  const request = event.request;
  if (request.method !== 'GET') return;

  const url = new URL(request.url);
  if (url.origin !== self.location.origin) return;

  // La página: siempre la versión más reciente; si no hay red, la guardada.
  if (request.mode === 'navigate') {
    event.respondWith(
      fetch(new Request(request, { cache: 'reload' }))
        .then(response => {
          if (response.ok) {
            const copy = response.clone();
            // Cada página se guarda con su propia dirección (la versión de prueba /nueva/
            // no debe sustituir a la app principal en el modo sin conexión).
            caches.open(CACHE_NAME).then(cache => cache.put(request.url.split('#')[0], copy));
          }
          return response;
        })
        .catch(() =>
          caches.match(request.url.split('#')[0]).then(cached => cached || caches.match('./index.html'))
        )
    );
    return;
  }

  // Iconos y manifest: se sirven al instante desde el dispositivo y se renuevan por detrás.
  event.respondWith(
    caches.match(request).then(cached => {
      const network = fetch(request)
        .then(response => {
          if (response.ok) {
            const copy = response.clone();
            caches.open(CACHE_NAME).then(cache => cache.put(request, copy));
          }
          return response;
        })
        .catch(() => cached);
      return cached || network;
    })
  );
});
