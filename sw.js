// SGS School Management System — Service Worker (V19.36)
// धोरण: नेटवर्क आधी, नेट नसल्यास Cache. त्यामुळे नवीन आवृत्ती Deploy केल्यावर जुना कोड अडकून राहत नाही.
// Apps Script (script.google.com), Google Drive फोटो व googleusercontent कधीही Cache केले जात नाहीत.
const VERSION = 'sgs-v19.36-1';
const SHELL = ['./', './index.html', './style.css', './script.js', './manifest.json', './icon-192.png', './icon-512.png'];
const NEVER_CACHE = /(^|\.)(script|drive|docs)\.google\.com$|googleusercontent\.com$/;
const CACHEABLE_CDN = /cdnjs\.cloudflare\.com$|cdn\.jsdelivr\.net$|fonts\.googleapis\.com$|fonts\.gstatic\.com$/;

self.addEventListener('install', function (e) {
  e.waitUntil(caches.open(VERSION).then(function (c) {
    return Promise.all(SHELL.map(function (u) { return c.add(u).catch(function () {}); }));
  }).then(function () { return self.skipWaiting(); }));
});

self.addEventListener('activate', function (e) {
  e.waitUntil(caches.keys().then(function (keys) {
    return Promise.all(keys.filter(function (k) { return k !== VERSION; }).map(function (k) { return caches.delete(k); }));
  }).then(function () { return self.clients.claim(); }));
});

self.addEventListener('fetch', function (e) {
  var req = e.request;
  if (req.method !== 'GET') return;
  var url = new URL(req.url);
  if (NEVER_CACHE.test(url.hostname)) return;
  var sameOrigin = url.origin === self.location.origin;
  if (!sameOrigin && !CACHEABLE_CDN.test(url.hostname)) return;
  e.respondWith(
    fetch(req).then(function (res) {
      if (res && (res.ok || res.type === 'opaque')) {
        var copy = res.clone();
        caches.open(VERSION).then(function (c) { c.put(req, copy); }).catch(function () {});
      }
      return res;
    }).catch(function () {
      return caches.match(req).then(function (hit) {
        if (hit) return hit;
        if (req.mode === 'navigate') return caches.match('./index.html');
        return Response.error();
      });
    })
  );
});
