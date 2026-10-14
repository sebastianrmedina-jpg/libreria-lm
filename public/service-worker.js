// Antes había DOS service workers separados peleando por el mismo scope ("/"):
// este (caching + su propio manejo de "push") y el de OneSignal
// (OneSignalSDKWorker.js). Eso es lo que probablemente rompía las
// notificaciones push en el celular — dos registros para el mismo scope
// entran en conflicto y no hay garantía de cuál queda "activo". La forma
// correcta de integrar OneSignal cuando ya existe un service worker propio es
// importar el suyo DENTRO de este (ver su doc de "Existing Service Worker") —
// así queda un solo worker real, y OneSignal maneja sus propios eventos de
// push/notificationclick sin que el código de acá los duplique ni interfiera.
importScripts("https://cdn.onesignal.com/sdks/web/v16/OneSignalSDK.sw.js");

const CACHE_NAME = "libreria-madrid-v2";
const STATIC_ASSETS = ["/", "/index.html", "/manifest.json", "/icon-192.png", "/icon-512.png"];

self.addEventListener("install", (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(STATIC_ASSETS))
  );
  self.skipWaiting();
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(keys.filter((k) => k !== CACHE_NAME).map((k) => caches.delete(k)))
    )
  );
  self.clients.claim();
});

self.addEventListener("fetch", (event) => {
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
