/* Albarán Tienda: funciona sin cobertura. Solo toca sus propias cachés (albaran-*), nunca las de TraceQueso o Recogida. */
const CACHE = 'albaran-v48';
const ARCHIVOS = ['./', './index.html', './manifest.json', './icon-192.png', './icon-512.png'];
self.addEventListener('install', e => e.waitUntil(caches.open(CACHE).then(c => c.addAll(ARCHIVOS)).then(() => self.skipWaiting())));
self.addEventListener('activate', e => e.waitUntil(
  caches.keys().then(k => Promise.all(k.filter(n => n.startsWith('albaran-') && n !== CACHE).map(n => caches.delete(n)))).then(() => self.clients.claim())));
self.addEventListener('fetch', e => {
  const r = e.request;
  if (r.method !== 'GET') return;
  if (r.mode === 'navigate' || r.url.endsWith('/index.html')) {   // primero red (mejoras), si no hay, la guardada
    e.respondWith(fetch(r).then(res => { const c2 = res.clone(); caches.open(CACHE).then(c => c.put('./index.html', c2)); return res; })
      .catch(() => caches.match('./index.html')));
    return;
  }
  e.respondWith(caches.match(r).then(hit => hit || fetch(r)));
});
