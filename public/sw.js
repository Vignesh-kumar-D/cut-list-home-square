/* Offline-first service worker (cache static assets for local/offline use).
 * Update strategy:
 * - network-first for HTML + model + JS/CSS so updates show up quickly
 * - cache-first for icons/other assets
 */

const CACHE_NAME = "cutlist-pwa-v1";
const ASSETS = [
  "./",
  "./index.html",
  "./styles.css",
  "./app.js",
  "./formula.js",
  "./model.json",
  "./manifest.webmanifest",
  "./icons/icon.svg"
];

self.addEventListener("install", (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(ASSETS)).then(() => self.skipWaiting())
  );
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    (async () => {
      const keys = await caches.keys();
      await Promise.all(keys.map((k) => (k === CACHE_NAME ? Promise.resolve() : caches.delete(k))));
      await self.clients.claim();
    })()
  );
});

self.addEventListener("fetch", (event) => {
  const req = event.request;
  const url = new URL(req.url);

  // Only handle same-origin GET requests
  if (req.method !== "GET" || url.origin !== self.location.origin) return;

  const path = url.pathname;
  const isNetworkFirst =
    req.mode === "navigate" ||
    path.endsWith("/") ||
    path.endsWith("/index.html") ||
    path.endsWith("/model.json") ||
    path.endsWith(".js") ||
    path.endsWith(".css") ||
    path.endsWith(".webmanifest");

  event.respondWith(
    (async () => {
      const cache = await caches.open(CACHE_NAME);
      if (isNetworkFirst) {
        try {
          const res = await fetch(req, { cache: "no-store" });
          if (res && res.ok) cache.put(req, res.clone());
          return res;
        } catch (e) {
          const cached = await cache.match(req);
          if (cached) return cached;
          if (req.mode === "navigate") {
            const fallback = await cache.match("./index.html");
            if (fallback) return fallback;
          }
          throw e;
        }
      }

      // cache-first for everything else
      const cached = await cache.match(req);
      if (cached) return cached;
      const res = await fetch(req);
      if (res && res.ok) cache.put(req, res.clone());
      return res;
    })()
  );
});


