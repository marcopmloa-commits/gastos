// Abre o app mesmo sem sinal. Os dados ficam no aparelho (localStorage) e sobem quando a internet volta.
const V = 'chile-v2';
const SHELL = ['./', 'index.html', 'icon.png', 'manifest.webmanifest'];
self.addEventListener('install', e => { e.waitUntil(caches.open(V).then(c => c.addAll(SHELL)).then(() => self.skipWaiting())); });
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== V).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});
self.addEventListener('fetch', e => {
  const req = e.request, u = new URL(req.url);
  if (req.method !== 'GET' || u.origin !== location.origin) return;
  e.respondWith((async () => {
    const c = await caches.open(V);
    const hit = await c.match(req, {ignoreSearch: true});
    const net = fetch(req).then(r => { if (r.ok) c.put(req.mode === 'navigate' ? 'index.html' : req, r.clone()); return r; }).catch(() => null);
    if (hit) { e.waitUntil(net); return hit; }
    return (await net) || (await c.match('index.html')) || Response.error();
  })());
});
