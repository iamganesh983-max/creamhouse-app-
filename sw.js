// ================================================================
//  CREAM HOUSE — Service Worker (full app version)
//  Caches the entire app shell so it opens instantly, even offline.
//  All data entry is stored in the browser (localStorage equivalent
//  via in-memory state) and CSV/Sheets sync happens when online.
// ================================================================

const CACHE_NAME = "creamhouse-fullapp-v1";
const ASSETS = [
  "./",
  "./index.html",
  "./manifest.json",
  "./icon-192.png",
  "./icon-512.png"
];

self.addEventListener("install", (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(ASSETS))
  );
  self.skipWaiting();
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys().then((names) =>
      Promise.all(names.filter((n) => n !== CACHE_NAME).map((n) => caches.delete(n)))
    )
  );
  self.clients.claim();
});

self.addEventListener("fetch", (event) => {
  // Google Sheets / script.google.com calls always go to network directly
  if (event.request.url.indexOf("script.google.com") >= 0 ||
      event.request.url.indexOf("googleapis.com") >= 0) {
    return;
  }

  event.respondWith(
    caches.match(event.request).then((cached) => {
      if (cached) return cached;
      return fetch(event.request).catch(() => caches.match("./index.html"));
    })
  );
});
