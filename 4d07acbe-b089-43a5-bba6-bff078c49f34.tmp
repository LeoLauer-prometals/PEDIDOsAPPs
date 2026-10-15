// Service worker do app "Pedidos" (Prometals).
//
// Objetivo: só deixar o app instalável no celular (ícone na tela inicial,
// abre em tela cheia) e funcionar um pouco melhor com internet fraca.
// NÃO guarda os pedidos/dados em cache — isso sempre vem direto do Supabase
// pela internet, exatamente como no navegador normal.
//
// IMPORTANTE pra quem for atualizar o app no futuro: sempre que subir uma
// versão nova do pedidosapp.html no Supabase, aumente o número abaixo
// (CACHE_VERSION). Isso força os celulares que já têm o app instalado a
// descartar a versão antiga guardada e buscar a nova. Sem isso, o celular
// pode continuar mostrando uma versão desatualizada do app por um tempo.
const CACHE_VERSION = 'v1';
const CACHE_NAME = 'pedidos-prometals-' + CACHE_VERSION;

const ARQUIVOS_DO_APP = [
  './pedidosapp.html',
  './manifest.json',
  './icon-192.png',
  './icon-512.png',
  './icon-512-maskable.png',
  './apple-touch-icon.png'
];

self.addEventListener('install', function (event) {
  self.skipWaiting();
  event.waitUntil(
    caches.open(CACHE_NAME).then(function (cache) {
      return cache.addAll(ARQUIVOS_DO_APP);
    })
  );
});

self.addEventListener('activate', function (event) {
  event.waitUntil(
    caches.keys().then(function (nomes) {
      return Promise.all(
        nomes
          .filter(function (nome) { return nome !== CACHE_NAME; })
          .map(function (nome) { return caches.delete(nome); })
      );
    }).then(function () { return self.clients.claim(); })
  );
});

self.addEventListener('fetch', function (event) {
  const req = event.request;

  // Só cuida de navegação/arquivos do próprio app (GET). Tudo que for pro
  // Supabase (dados, login, storage de anexos etc.) passa direto pela rede,
  // sem passar pelo cache — o app tem que sempre mostrar dados atuais.
  if (req.method !== 'GET') return;

  const url = new URL(req.url);
  const ehArquivoDoApp = ARQUIVOS_DO_APP.some(function (arq) {
    return url.pathname.endsWith(arq.replace('./', '/'));
  });
  if (!ehArquivoDoApp) return;

  // Rede primeiro (pra sempre pegar a versão mais nova quando tem internet);
  // só usa o que tá guardado se a rede falhar (sem internet no momento).
  event.respondWith(
    fetch(req)
      .then(function (resposta) {
        const copia = resposta.clone();
        caches.open(CACHE_NAME).then(function (cache) { cache.put(req, copia); });
        return resposta;
      })
      .catch(function () {
        return caches.match(req);
      })
  );
});
