/* Lika's Project service worker — offline support.
   Pages: network first (so updates arrive), cache as fallback when offline.
   Fonts/icons: cache first. Google Apps Script (sync) is never cached. */
const CACHE = 'lika-v4';
const CORE = ['./', './index.html', './manifest.webmanifest', './icons/icon-192.png', './icons/icon-512.png'];
self.addEventListener('install', e => { e.waitUntil(caches.open(CACHE).then(c => c.addAll(CORE)).then(() => self.skipWaiting())); });
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});
self.addEventListener('fetch', e => {
  const u = new URL(e.request.url);
  if (e.request.method !== 'GET' || u.hostname.includes('script.google') || u.hostname.includes('googleusercontent') || u.hostname.includes('anthropic')) return;
  if (e.request.mode === 'navigate' || u.pathname.endsWith('.html')) {
    e.respondWith(fetch(e.request).then(r => { const cp = r.clone(); caches.open(CACHE).then(c => c.put(e.request, cp)); return r; })
      .catch(() => caches.match(e.request).then(r => r || caches.match('./index.html'))));
    return;
  }
  e.respondWith(caches.match(e.request).then(r => r || fetch(e.request).then(res => {
    if (res.ok && (u.origin === location.origin || u.hostname.includes('gstatic') || u.hostname.includes('googleapis'))) {
      const cp = res.clone(); caches.open(CACHE).then(c => c.put(e.request, cp));
    }
    return res;
  })));
});
