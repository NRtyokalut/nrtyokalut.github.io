const CACHE = "pekantyokalut-v76";
// The checker app lives at ./tarkastus/ with its own service worker and caches ("tarkastus-*").
const SUBAPP = new URL("./tarkastus/", self.registration.scope).href;
const ASSETS = [
  "./",
  "./index.html",
  "./styles.css",
  "./app.js",
  "./calc.js",
  "./manifest.webmanifest",
  "./lib/xlsx.full.min.js",
  "./icons/icon-192.png",
  "./icons/icon-512.png",
];

// Always fetch fresh copies at install so old and new files never mix.
self.addEventListener("install", (e) => {
  e.waitUntil(
    caches
      .open(CACHE)
      .then((c) => c.addAll(ASSETS.map((u) => new Request(u, { cache: "reload" }))))
      .then(() => self.skipWaiting())
  );
});

self.addEventListener("activate", (e) => {
  e.waitUntil(
    caches
      .keys()
      .then((keys) =>
        Promise.all(keys.filter((k) => k !== CACHE && !k.startsWith("tarkastus-")).map((k) => caches.delete(k)))
      )
      .then(() => self.clients.claim())
  );
});

// Network first (fresh files when online), cache as offline fallback.
self.addEventListener("fetch", (e) => {
  const req = e.request;
  if (req.method !== "GET") return;
  // Never touch the checker app: its own SW (scope ./tarkastus/) handles it; before that, plain network.
  if (req.url.startsWith(SUBAPP)) return;
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
