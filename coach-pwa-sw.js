const CACHE_NAME = "coachflow-coach-20260502-0001";
const CORE_ASSETS = [
  "./leave-coach-sandbox.html?v=20260502-0001",
  "./config.js?v=20260502-0001",
  "./leave-sandbox.js?v=20260502-0001",
  "./coach-pwa.webmanifest?v=20260502-0001",
  "./coachflow-coach-icon-192.png",
  "./coachflow-coach-icon-512.png",
  "./coachflow-coach-icon.svg"
];

self.addEventListener("install", (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then((cache) => cache.addAll(CORE_ASSETS))
      .then(() => self.skipWaiting())
  );
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys()
      .then((keys) => Promise.all(keys
        .filter((key) => key.startsWith("coachflow-coach-") && key !== CACHE_NAME)
        .map((key) => caches.delete(key))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener("fetch", (event) => {
  if (event.request.method !== "GET") {
    return;
  }
  const url = new URL(event.request.url);
  if (url.origin !== self.location.origin) {
    return;
  }
  event.respondWith(
    fetch(event.request)
      .then((response) => {
        const copy = response.clone();
        caches.open(CACHE_NAME).then((cache) => cache.put(event.request, copy));
        return response;
      })
      .catch(() => caches.match(event.request).then((cached) => cached || caches.match("./leave-coach-sandbox.html?v=20260502-0001")))
  );
});
