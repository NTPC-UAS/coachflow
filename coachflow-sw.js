const COACHFLOW_PWA_VERSION = "20260503-0017";
const CACHE_NAME = `coachflow-system-${COACHFLOW_PWA_VERSION}`;
const CORE_ASSETS = [
  "./",
  `./index.html?v=${COACHFLOW_PWA_VERSION}`,
  `./admin.html?v=${COACHFLOW_PWA_VERSION}`,
  `./coach.html?v=${COACHFLOW_PWA_VERSION}`,
  `./student.html?v=${COACHFLOW_PWA_VERSION}`,
  `./install.html?v=${COACHFLOW_PWA_VERSION}`,
  `./styles.css?v=${COACHFLOW_PWA_VERSION}`,
  `./config.js?v=${COACHFLOW_PWA_VERSION}`,
  `./app.js?v=${COACHFLOW_PWA_VERSION}`,
  `./coachflow.webmanifest?v=${COACHFLOW_PWA_VERSION}`,
  "./coachflow-coach-icon-192.png",
  "./coachflow-coach-icon-512.png",
  "./coachflow-coach-icon.svg"
];
const NAVIGATION_FALLBACK = `./index.html?v=${COACHFLOW_PWA_VERSION}`;

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
        .filter((key) => (
          (key.startsWith("coachflow-system-") || key.startsWith("coachflow-coach-"))
          && key !== CACHE_NAME
        ))
        .map((key) => caches.delete(key))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener("fetch", (event) => {
  const request = event.request;
  if (request.method !== "GET") {
    return;
  }

  const url = new URL(request.url);
  if (url.origin !== self.location.origin) {
    return;
  }

  if (request.mode === "navigate") {
    event.respondWith(networkFirst(request, NAVIGATION_FALLBACK));
    return;
  }

  event.respondWith(networkFirst(request));
});

async function networkFirst(request, fallbackUrl = "") {
  const cache = await caches.open(CACHE_NAME);
  try {
    const response = await fetch(request);
    if (response && response.ok) {
      await cache.put(request, response.clone());
    }
    return response;
  } catch (error) {
    const cached = await caches.match(request);
    if (cached) {
      return cached;
    }
    if (fallbackUrl) {
      return caches.match(fallbackUrl);
    }
    throw error;
  }
}
