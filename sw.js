// Caches the app shell so it opens instantly and installs as an app.
// Calendar API calls are never cached: your calendar is always the source of truth.
const CACHE = "taskboard-v13";
const SHELL = ["./", "index.html", "styles.css", "app.js", "config.js", "i18n.js", "core.js", "backends.js", "manifest.webmanifest", "icons/icon.svg", "vendor/msal-browser.min.js"];

self.addEventListener("install", (e) => {
  e.waitUntil(caches.open(CACHE).then((c) => c.addAll(SHELL.map((u) => new Request(u, { cache: "reload" })))).then(() => self.skipWaiting()));
});
self.addEventListener("activate", (e) => {
  e.waitUntil(
    caches.keys().then((keys) => Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});
self.addEventListener("fetch", (e) => {
  const url = new URL(e.request.url);
  if (e.request.method !== "GET" || url.origin !== location.origin) return;
  // Network first so updates show up, cache as offline fallback. "no-cache" makes the browser
  // check with the server, so a fresh page never runs with an older copy of a script.
  e.respondWith(
    fetch(e.request, { cache: "no-cache" })
      .then((res) => {
        const copy = res.clone();
        caches.open(CACHE).then((c) => c.put(e.request, copy));
        return res;
      })
      .catch(() => caches.match(e.request, { ignoreSearch: true }))
  );
});
