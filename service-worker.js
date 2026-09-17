const CACHE_PREFIX = "work-payment-log-";
const CACHE = CACHE_PREFIX + "1.11.4-005a7d170ea5";
const SHELL = [
  "./",
  "./index.html",
  "./styles.css",
  "./seed.js",
  "./sync-core.js",
  "./cloud.js",
  "./app.js",
  "./manifest.webmanifest",
  "./quick-log.html",
  "./quick-log.webmanifest",
  "./widget.html",
  "./widget.css",
  "./widget.js",
  "./quicklog/",
  "./icons/icon-180.png",
  "./icons/icon-192.png",
  "./icons/icon-512.png"
];

self.addEventListener("install", event => {
  event.waitUntil(caches.open(CACHE).then(cache => cache.addAll(SHELL)).then(() => self.skipWaiting()));
});

self.addEventListener("activate", event => {
  event.waitUntil(
    caches.keys()
      .then(keys => Promise.all(keys.filter(key => key.startsWith(CACHE_PREFIX) && key !== CACHE).map(key => caches.delete(key))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener("fetch", event => {
  const request = event.request;
  if (request.method !== "GET") return;
  const url = new URL(request.url);
  if (url.origin !== self.location.origin) return;

  const scope = self.registration.scope;
  const path = url.pathname;
  // Cache the complete release at install, then serve it immediately even on
  // a slow connection. The worker's content hash installs the next release.
  // HTML uses ?v=... URLs; these must resolve to the precached asset paths.
  const shellPath = SHELL.find(file => new URL(file, scope).pathname === path);
  const quickLog = new URL("./quicklog/", scope).pathname;
  const key = shellPath || (path === quickLog + "index.html" || path === quickLog.slice(0, -1) ? "./quicklog/" : null);
  if (key) {
    event.respondWith(caches.open(CACHE).then(async cache => {
      const cached = await cache.match(new URL(key, scope).href);
      if (cached) return cached;
      // Never substitute HTML for missing JavaScript, CSS or images.
      return fetch(request);
    }));
  }
  // Other requests (including update manifests) stay on the network. Cloud
  // data is managed by the app's local store and is not put in this cache.
});
