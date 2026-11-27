// Offline cache for the app shell.
// The page itself is fetched network-first, so a new version shows on the next open;
// the cached copy is used only when there is no connection. Icons and the manifest are
// served cache-first. Bump VERSION when shipping, so old caches are cleared.
const VERSION = 'paynote-v13';
const SHELL = ['./', './index.html', './manifest.webmanifest', './icon-192.png', './icon-512.png'];

self.addEventListener('install', e => {
  // cache:'reload' skips the browser's HTTP cache, so a new worker never stores a stale page
  e.waitUntil(caches.open(VERSION)
    .then(c => c.addAll(SHELL.map(u => new Request(u, {cache: 'reload'}))))
    .then(() => self.skipWaiting()));
});

self.addEventListener('activate', e => {
  e.waitUntil(caches.keys()
    .then(ks => Promise.all(ks.filter(k => k !== VERSION).map(k => caches.delete(k))))
    .then(() => self.clients.claim()));
});

self.addEventListener('fetch', e => {
  const req = e.request, url = new URL(req.url);
  if (req.method !== 'GET' || url.origin !== location.origin) return;   // the sheet API and fonts go straight to the network
  const isPage = req.mode === 'navigate' || url.pathname.endsWith('/') || url.pathname.endsWith('/index.html');
  if (isPage) {
    e.respondWith(fetch(req, {cache: 'no-store'})
      .then(res => { if (res.ok) { const copy = res.clone(); caches.open(VERSION).then(c => c.put('./index.html', copy)); } return res; })
      .catch(() => caches.match('./index.html')));
    return;
  }
  e.respondWith(caches.match(req).then(hit => hit || fetch(req).then(res => {
    if (res.ok) { const copy = res.clone(); caches.open(VERSION).then(c => c.put(req, copy)); }
    return res;
  })));
});
