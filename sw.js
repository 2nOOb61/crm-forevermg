// FOREVER MG CRM — Service Worker v3.0
const CACHE = "forevermg-crm-v3";
const ASSETS = [
  "./index.html",
  "./manifest.json"
];

self.addEventListener("install", e => {
  e.waitUntil(
    caches.open(CACHE).then(c => c.addAll(ASSETS)).then(() => self.skipWaiting())
  );
});

self.addEventListener("activate", e => {
  e.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k)))
    ).then(() => self.clients.claim())
  );
});

self.addEventListener("fetch", e => {
  const req = e.request;

  // Requêtes GAS : toujours réseau (jamais de cache)
  if (req.url.includes("script.google.com")) return;

  // Page HTML / navigation : RÉSEAU D'ABORD (toujours la dernière version),
  // cache uniquement en secours hors-ligne.
  const isHTML = req.mode === "navigate"
    || req.destination === "document"
    || req.url.endsWith("/")
    || req.url.endsWith(".html");

  if (isHTML) {
    e.respondWith(
      fetch(req).then(res => {
        if (res && res.status === 200) {
          const clone = res.clone();
          caches.open(CACHE).then(c => c.put("./index.html", clone));
        }
        return res;
      }).catch(() => caches.match("./index.html"))
    );
    return;
  }

  // Autres ressources statiques : cache d'abord
  e.respondWith(
    caches.match(req).then(cached => {
      if (cached) return cached;
      return fetch(req).then(res => {
        if (res && res.status === 200 && req.method === "GET") {
          const clone = res.clone();
          caches.open(CACHE).then(c => c.put(req, clone));
        }
        return res;
      }).catch(() => caches.match("./index.html"));
    })
  );
});
