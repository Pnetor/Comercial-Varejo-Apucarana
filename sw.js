// Service worker do painel Estoque Apucarana.
//
// Objetivo: deixar o site instalável (PWA) e ainda abrir (com o último
// dado visto) se a conexão cair no meio de um pedido/consulta - mas SEM
// nunca preferir uma versão antiga do estoque/preços/pedidos quando tem
// internet. Por isso as páginas HTML são "network-first" (tenta buscar a
// versão mais nova sempre; só usa o cache se a rede falhar de verdade).
// Só os ícones/manifest (que não mudam) ficam em cache-first.

const CACHE_VERSION = 'estoque-apucarana-v1';

const ARQUIVOS_ESTATICOS = [
  './manifest.json',
  './icon-192.png',
  './icon-512.png',
];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_VERSION).then((cache) => cache.addAll(ARQUIVOS_ESTATICOS))
  );
  self.skipWaiting();
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((nomes) =>
      Promise.all(
        nomes
          .filter((nome) => nome !== CACHE_VERSION)
          .map((nome) => caches.delete(nome))
      )
    )
  );
  self.clients.claim();
});

self.addEventListener('fetch', (event) => {
  const req = event.request;
  if (req.method !== 'GET') return;

  const url = new URL(req.url);
  if (url.origin !== self.location.origin) return;

  const ehPagina = req.mode === 'navigate' || req.headers.get('accept')?.includes('text/html');

  if (ehPagina) {
    // Network-first: sempre busca a versão mais nova do painel; só cai
    // pro cache (última versão vista) se estiver sem internet.
    event.respondWith(
      fetch(req)
        .then((resp) => {
          const copia = resp.clone();
          caches.open(CACHE_VERSION).then((cache) => cache.put(req, copia));
          return resp;
        })
        .catch(() => caches.match(req))
    );
    return;
  }

  // Estático (ícone, manifest): cache-first, com atualização em segundo plano.
  event.respondWith(
    caches.match(req).then((cacheado) => {
      const buscaRede = fetch(req)
        .then((resp) => {
          const copia = resp.clone();
          caches.open(CACHE_VERSION).then((cache) => cache.put(req, copia));
          return resp;
        })
        .catch(() => cacheado);
      return cacheado || buscaRede;
    })
  );
});
