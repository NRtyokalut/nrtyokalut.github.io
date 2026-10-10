// Vuorotaulun tarkastus — own service worker, scope ./ (= /tarkastus/). Separate from the root app.
const CACHE = "tarkastus-v2";
const ASSETS = [
  "./",
  "./index.html",
  "./styles.css",
  "./tarkastus.css",
  "./app.js",
  "./calc.js",
  "./manifest.webmanifest",
  "./lib/xlsx.full.min.js",
  "./icons/icon-192.png",
  "./icons/icon-512.png",
];

self.addEventListener("install", (e) => {
  e.waitUntil(
    caches
      .open(CACHE)
      .then((c) => c.addAll(ASSETS.map((u) => new Request(u, { cache: "reload" }))))
      .then(() => self.skipWaiting())
  );
});

// Only clean up this app's own old caches ("tarkastus-*"); the root app's caches are left alone.
self.addEventListener("activate", (e) => {
  e.waitUntil(
    caches
      .keys()
      .then((keys) => Promise.all(keys.filter((k) => k !== CACHE && k.startsWith("tarkastus-")).map((k) => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

// Network first, cache as offline fallback. Only requests inside this scope.
self.addEventListener("fetch", (e) => {
  const req = e.request;
  if (req.method !== "GET" || !req.url.startsWith(self.registration.scope)) return;
  e.respondWith(
    fetch(req, { cache: "no-store" })
      .then((res) => {
        if (res && res.ok) {
          const copy = res.clone();
          caches.open(CACHE).then((c) => c.put(req, copy)).catch(() => {});
        }
        return res;
      })
      .catch(() => caches.match(req, { ignoreSearch: true }))
  );
});

/*! © 2026 Lämpöpumppu Mafia. Kaikki oikeudet pidätetään. All rights reserved. */
