// Service Worker — Apontamento de Fundações DC Pecém
// Cache simples para permitir instalação como PWA e funcionamento offline básico.

const CACHE_NAME = "apontamento-v5";
const FILES_TO_CACHE = [
  "./",
  "./index.html",
  "./apontamento_datas.html",
  "./terraplanagem.html",
  "./manifest.webmanifest",
  "./icon-192.png",
  "./icon-512.png",
  "./icon-512-maskable.png",
];

self.addEventListener("install", event => {
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then(cache => cache.addAll(FILES_TO_CACHE))
      .then(() => self.skipWaiting())
  );
});

self.addEventListener("activate", event => {
  event.waitUntil(
    caches.keys()
      .then(keys => Promise.all(keys.filter(k => k !== CACHE_NAME).map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener("fetch", event => {
  const url = new URL(event.request.url);

  // NUNCA cachear o Apps Script (precisamos sempre da resposta fresca)
  if (url.hostname === "script.google.com") return;

  // POSTs nunca passam por cache
  if (event.request.method !== "GET") return;

  // Cache-first para os arquivos da app
  event.respondWith(
    caches.match(event.request).then(cached => {
      if (cached) return cached;
      return fetch(event.request).then(response => {
        // Só cachear respostas próprias (mesma origem)
        if (response.ok && url.origin === self.location.origin) {
          const clone = response.clone();
          caches.open(CACHE_NAME).then(c => c.put(event.request, clone));
        }
        return response;
      }).catch(() => {
        // Offline e sem cache: retorna o hub para nav requests
        if (event.request.mode === "navigate") {
          return caches.match("./index.html");
        }
      });
    })
  );
});
