/* Gestor de Condomínios — Service Worker v7
   Estratégia: NETWORK-FIRST. A rede é sempre consultada primeiro;
   o cache serve apenas como fallback offline. Isso evita o problema
   clássico de PWA servir versões antigas do app. */
const CACHE = 'gestor-cond-v7';
const ESTATICOS = ['./', './index.html', './style.css?v=7', './app.js?v=7',
                   './manifest.json', './icon-192.png'];

self.addEventListener('install', (e) => {
  e.waitUntil(caches.open(CACHE).then((c) => c.addAll(ESTATICOS)));
  self.skipWaiting();
});

self.addEventListener('activate', (e) => {
  e.waitUntil(
    caches.keys().then((ks) =>
      Promise.all(ks.filter((k) => k !== CACHE).map((k) => caches.delete(k)))
    ).then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', (e) => {
  const url = new URL(e.request.url);
  // Nunca interceptar API (Render), Supabase ou métodos não-GET
  if (e.request.method !== 'GET' || url.origin !== self.location.origin) return;
  e.respondWith(
    fetch(e.request)
      .then((resp) => {
        const copia = resp.clone();
        caches.open(CACHE).then((c) => c.put(e.request, copia));
        return resp;
      })
      .catch(() => caches.match(e.request))
  );
});
