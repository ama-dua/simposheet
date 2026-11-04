/* SimpoSheet service worker.
   Deliberately conservative: it caches only the app shell and static assets
   so the installed app opens fast. Pages are network-first, and document
   conversions (/convert, file uploads/downloads) are NEVER cached. */
const CACHE = "simposheet-v1";
const SHELL = [
  "/static/styles.css",
  "/static/app.js",
  "/static/icons/icon-192.png",
  "/static/icons/icon-512.png",
  "/static/manifest.webmanifest"
];

self.addEventListener("install", (event) => {
  event.waitUntil(
    caches.open(CACHE).then((cache) => cache.addAll(SHELL)).then(() => self.skipWaiting())
  );
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys()
      .then((keys) => Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener("fetch", (event) => {
  const url = new URL(event.request.url);
  if (event.request.method !== "GET") return; // uploads/conversions: network only
  if (url.pathname.startsWith("/static/")) {
    event.respondWith(
      caches.match(event.request).then((hit) => hit || fetch(event.request))
    );
    return;
  }
  // Everything else (pages, /health, downloads): network first, no caching.
  event.respondWith(fetch(event.request).catch(() => caches.match("/")));
});
