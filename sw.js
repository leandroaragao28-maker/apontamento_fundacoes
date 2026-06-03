// Service Worker — Apontamento de Fundações DC Pecém
const CACHE_NAME = "apontamento-v7";

// Páginas HTML: sempre busca versão mais recente na rede
const NETWORK_FIRST = ["index.html", "apontamento_datas.html", "terraplanagem.html", "/apontamento_fundacoes/"];

self.addEventListener("install", event => {
  self.skipWaiting();
});

self.addEventListener("activate", event => {
  event.waitUntil(
    caches.keys()
      .then(keys => Promise.all(keys.filter(k => k !== CACHE_NAME).map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener("fetch", event => {
  const url = event.request.url;

  // Ignora URLs que não sejam http/https (ex: chrome-extension://)
  if (!url.startsWith("http")) return;

  // Apps Script: sempre rede, sem cache
  if (url.includes("script.google.com")) return;

  // POSTs nunca passam por cache
  if (event.request.method !== "GET") return;

  const isNetworkFirst = NETWORK_FIRST.some(p => url.endsWith(p));

  if (isNetworkFirst) {
    // REDE PRIMEIRO: garante sempre a versão mais recente
    event.respondWith(
      fetch(event.request)
        .then(res => {
          const clone = res.clone();
          caches.open(CACHE_NAME).then(c => c.put(event.request, clone));
          return res;
        })
        .catch(() => caches.match(event.request))
    );
  } else {
    // CACHE PRIMEIRO: ícones e recursos estáticos (performance)
    event.respondWith(
      caches.match(event.request).then(cached => {
        if (cached) return cached;
        return fetch(event.request).then(res => {
          if (res && res.ok && new URL(url).origin === self.location.origin) {
            const clone = res.clone();
            caches.open(CACHE_NAME).then(c => c.put(event.request, clone));
          }
          return res;
        }).catch(() => cached);
      })
    );
  }
});
