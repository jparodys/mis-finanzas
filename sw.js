// Guarda la app en el celular para que abra rápido y funcione sin señal.
// Las llamadas a Microsoft (login y OneDrive) nunca se guardan en caché.
const CACHE = "mis-finanzas-v5";
const ARCHIVOS = [
  "./", "index.html", "app.js", "xlsx-registro.js", "config.js", "manifest.webmanifest",
  "icons/icon-192.png", "icons/icon-512.png",
  "https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js",
];

self.addEventListener("install", (e) => {
  e.waitUntil(caches.open(CACHE).then((c) => c.addAll(ARCHIVOS)).then(() => self.skipWaiting()));
});
self.addEventListener("activate", (e) => {
  e.waitUntil(
    caches.keys().then((ks) => Promise.all(ks.filter((k) => k !== CACHE).map((k) => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});
self.addEventListener("fetch", (e) => {
  const url = new URL(e.request.url);
  const propio = url.origin === location.origin || url.hostname === "cdnjs.cloudflare.com";
  if (e.request.method !== "GET" || !propio) return;
  // Red primero (revalidando con el servidor para recibir actualizaciones), caché si no hay conexión
  const pedido = url.origin === location.origin ? new Request(e.request, { cache: "no-cache" }) : e.request;
  e.respondWith(
    fetch(pedido)
      .then((r) => {
        const copia = r.clone();
        caches.open(CACHE).then((c) => c.put(e.request, copia));
        return r;
      })
      .catch(() => caches.match(e.request, { ignoreSearch: true }))
  );
});
