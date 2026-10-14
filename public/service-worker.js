// Este archivo ya NO se usa — se volvió a OneSignal con su propio service
// worker separado (/OneSignalSDKWorker.js), sin combinarlo con uno propio,
// porque la combinación (importScripts) no terminó de mostrar las
// notificaciones de forma confiable en Android.
//
// Lo que queda acá es, a propósito, un service worker "autodestructivo": si
// algún dispositivo todavía tiene ESTE archivo registrado (de la versión
// anterior), Chrome lo va a revisar solo la próxima vez que abra la app,
// encontrar este contenido nuevo, instalarlo, y al activarse se da de baja
// solo y recarga la página — así ese dispositivo queda limpio sin que haga
// falta borrar datos del sitio a mano otra vez.
self.addEventListener("install", () => self.skipWaiting());
self.addEventListener("activate", (event) => {
  event.waitUntil(
    (async () => {
      await self.registration.unregister();
      const clientsList = await self.clients.matchAll({ type: "window" });
      clientsList.forEach((client) => client.navigate(client.url));
    })()
  );
});
