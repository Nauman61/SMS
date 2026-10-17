// Minimal service worker: caches the app SHELL (the HTML/JS/CSS files
// themselves) so the site can be installed as a PWA and reopens even with
// a flaky connection.
//
// IMPORTANT: this must NEVER intercept requests to Supabase (or any other
// cross-origin API) — doing so previously caused different devices to see
// different, out-of-sync data, since a cached response or a stale cached
// app bundle could silently keep a device running on old code/old data
// instead of talking to the live database. The fetch handler below only
// ever touches same-origin requests; everything else (Supabase calls
// included) is left completely untouched and goes straight to the network
// exactly as it would with no service worker at all.

const CACHE_NAME = "hss-school-shell-v2";
const APP_SHELL = ["/", "/index.html", "/manifest.webmanifest"];

self.addEventListener("install", (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(APP_SHELL)).catch(() => {})
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

// Network-first for same-origin navigation/JS/CSS, so a new deployment is
// picked up immediately when online; falls back to the cached shell only
// when genuinely offline. Cross-origin requests (Supabase, fonts, etc.)
// are never intercepted — they're left to the browser's normal handling.
self.addEventListener("fetch", (event) => {
  if (event.request.method !== "GET") return;
  if (new URL(event.request.url).origin !== self.location.origin) return;

  event.respondWith(
    fetch(event.request)
      .then((response) => {
        const copy = response.clone();
        caches.open(CACHE_NAME).then((cache) => cache.put(event.request, copy)).catch(() => {});
        return response;
      })
      .catch(() => caches.match(event.request).then((cached) => cached || caches.match("/index.html")))
  );
});
