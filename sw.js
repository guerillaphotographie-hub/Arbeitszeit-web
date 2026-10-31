/* Service Worker: macht die App offline nutzbar (Dateien werden beim ersten Besuch gespeichert). */
var CACHE = 'arbeitszeit-v2';
var FILES = ['./', 'index.html', 'app.js', 'core.js', 'xlsx-assets.js', 'manifest.webmanifest',
  'icons/icon-180.png', 'icons/icon-192.png', 'icons/icon-512.png'];

self.addEventListener('install', function (ev) {
  ev.waitUntil(caches.open(CACHE).then(function (c) { return c.addAll(FILES); }).then(function () { return self.skipWaiting(); }));
});

self.addEventListener('activate', function (ev) {
  ev.waitUntil(caches.keys().then(function (keys) {
    return Promise.all(keys.filter(function (k) { return k !== CACHE; }).map(function (k) { return caches.delete(k); }));
  }).then(function () { return self.clients.claim(); }));
});

// Erst das Netz versuchen (damit Updates ankommen), sonst den gespeicherten Stand nehmen.
self.addEventListener('fetch', function (ev) {
  if (ev.request.method !== 'GET') return;
  ev.respondWith(
    fetch(ev.request).then(function (res) {
      var copy = res.clone();
      caches.open(CACHE).then(function (c) { c.put(ev.request, copy); });
      return res;
    }).catch(function () { return caches.match(ev.request).then(function (r) { return r || caches.match('index.html'); }); })
  );
});
