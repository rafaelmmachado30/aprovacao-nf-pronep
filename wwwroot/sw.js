/**
 * Service Worker — Sistema de Aprovacao de NF Pronep
 *
 * Estrategia:
 *  - NUNCA intercepta /api/* (precisa sempre bater no backend - dados dinamicos)
 *  - NUNCA intercepta /.auth/* (fluxo de login Easy Auth quebra se cachear)
 *  - NUNCA intercepta hosts externos (graph.microsoft.com, login.microsoftonline.com, etc)
 *  - NUNCA intercepta POST/PATCH/DELETE/PUT (so navegacao e GET de assets)
 *  - Cacheia o shell (index.html + vendor/* + icones) pra abrir offline
 *  - Network-first pro index.html (sempre tenta versao fresca; fallback pro cache)
 *  - Cache-first pra assets estaticos (vendor scripts e icones)
 *
 * IMPORTANTE: pra atualizar o SW a cada deploy, bumpa CACHE_VERSION.
 */

const CACHE_VERSION = 'pronep-nf-v5-html-integro-20260824';
const SHELL_CACHE = `${CACHE_VERSION}-shell`;

/* ============================================================================
   HTML TRUNCADO NO CACHE — o defeito que este bloco existe para impedir.

   Em 24/08/2026 o Rafael abriu o sistema e viu uma pagina TOTALMENTE BRANCA. O
   DevTools mostrou o documento terminando depois do </head>: sem <body>, altura
   zero, e o console VAZIO — porque sem body nao ha script para rodar nem erro para
   reportar. O arquivo no servidor estava intacto (589.582 bytes, identico ao
   repositorio); o que estava truncado era a copia no cache do navegador.

   A CAUSA: `resp.ok` e verdadeiro num HTTP 200 mesmo que o CORPO chegue cortado no
   meio (queda de conexao, proxy, rede instavel). O codigo antigo era:

       if (resp && resp.ok) c.put('/index.html', resp.clone());

   ou seja, guardava o pedaco que chegou. Depois, no primeiro `fetch` que falhasse, o
   `.catch` servia esse pedaco — e continuava servindo PARA SEMPRE, porque nada
   invalida um cache que "existe". O usuario nao tem como diagnosticar: pagina branca,
   sem erro, e recarregar nao resolve.

   A REGRA AGORA: HTML so entra no cache se estiver COMPLETO, e so sai do cache se
   estiver completo. Um documento sem `</html>` no fim, ou sem `<body`, e descartado
   nas duas direcoes. Guardar so o que esta inteiro e mais importante que guardar.
   ========================================================================== */
const HTML_TERMINA = /<\/html>\s*$/i;

async function htmlCompleto(resp) {
  try {
    const txt = await resp.clone().text();
    return (HTML_TERMINA.test(txt) && txt.indexOf('<body') >= 0) ? txt : null;
  } catch (e) {
    /* Corpo ilegivel (stream abortado) tambem e HTML incompleto. */
    return null;
  }
}

/* Guarda o TEXTO VALIDADO, e nao a resposta original: assim o que fica no cache e
   exatamente o que foi conferido, sem chance de a stream ser consumida pela metade
   entre a checagem e a gravacao. */
async function guardarHtmlSeCompleto(chave, resp) {
  const txt = await htmlCompleto(resp);
  if (!txt) {
    console.warn('[SW] HTML incompleto recusado pelo cache:', chave);
    return false;
  }
  const cache = await caches.open(SHELL_CACHE);
  await cache.put(chave, new Response(txt, {
    status: 200,
    headers: { 'Content-Type': resp.headers.get('content-type') || 'text/html; charset=utf-8' }
  }));
  return true;
}

/* Le do cache SO se estiver completo. Se achar truncado, APAGA — senao um cache
   envenenado por uma queda de rede antiga sobreviveria a todas as visitas futuras. */
async function lerHtmlDoCacheSeCompleto(chave) {
  const achado = await caches.match(chave);
  if (!achado) return null;
  const txt = await htmlCompleto(achado);
  if (txt) return achado;
  console.warn('[SW] cache tinha HTML truncado — descartando:', chave);
  try { const c = await caches.open(SHELL_CACHE); await c.delete(chave); } catch (e) {}
  return null;
}
const SHELL_ASSETS = [
  '/',
  '/index.html',
  '/offline.html',
  '/manifest.webmanifest',
  '/favicon.svg',
  '/favicon-256.png',
  '/icon-192.png',
  '/icon-512.png',
  '/apple-touch-icon.png',
  '/pronep-logo.png',
  '/vendor/chart.umd.min.js',
  '/vendor/teams-js.min.js',
  '/vendor/xlsx.full.min.js'
];

// Instalacao: faz pre-cache do shell
self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(SHELL_CACHE).then((cache) => {
      // Best-effort: se algum asset falhar, continua mesmo assim
      /* index.html sai do cache.add e passa pela validacao: cache.add usa fetch+put por
         dentro e guardaria um corpo truncado do mesmo jeito — era a segunda porta para o
         mesmo defeito. */
      return Promise.allSettled(SHELL_ASSETS.map(function (url) {
        if (url === '/index.html') {
          return fetch(url, { cache: 'reload' })
            .then(r => (r && r.ok) ? guardarHtmlSeCompleto('/index.html', r) : null)
            .catch(() => null);
        }
        return cache.add(url).catch(() => null);
      }));
    }).then(() => self.skipWaiting())
  );
});

// Ativacao: limpa caches de versoes antigas
self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(
        keys.filter(k => k.startsWith('pronep-nf-') && !k.startsWith(CACHE_VERSION))
            .map(k => caches.delete(k))
      )
    ).then(() => self.clients.claim())
  );
});

// Decide se deve interceptar um request
function deveBypassar(url, method) {
  // Bypass total: POST/PATCH/PUT/DELETE
  if (method !== 'GET') return true;
  // Bypass: API
  if (url.pathname.startsWith('/api/')) return true;
  // Bypass: auth Easy Auth
  if (url.pathname.startsWith('/.auth/')) return true;
  // Bypass: hosts externos (Microsoft, Graph, Brasilapi, CDN externos)
  if (url.origin !== self.location.origin) return true;
  return false;
}

self.addEventListener('fetch', (event) => {
  const req = event.request;
  let url;
  try { url = new URL(req.url); } catch (e) { return; }

  if (deveBypassar(url, req.method)) {
    // Deixa o request seguir normal pelo browser (sem nos meter)
    return;
  }

  // Navegacao (HTML): network-first com fallback pro cache, depois offline.html
  const accept = req.headers.get('accept') || '';
  const isNavigation = req.mode === 'navigate' || accept.includes('text/html');

  if (isNavigation) {
    event.respondWith(
      fetch(req)
        .then((resp) => {
          /* Clone EXPLICITO aqui, antes de a resposta seguir para o browser: a validacao
             precisa de um corpo proprio para ler, e a stream original nao pode ser tocada. */
          if (resp && resp.ok) {
            const copia = resp.clone();
            guardarHtmlSeCompleto('/index.html', copia);
          }
          return resp;
        })
        .catch(() =>
          lerHtmlDoCacheSeCompleto('/index.html')
            .then(c => c || caches.match('/offline.html')))
    );
    return;
  }

  /* CACHE-FIRST SO PARA O QUE E IMUTAVEL NA PRATICA.
     Antes TODO asset do mesmo origem era cache-first, styles.css incluso. Efeito:
     enquanto CACHE_VERSION nao mudasse, o navegador servia o CSS antigo PARA SEMPRE,
     ignorando qualquer header do servidor. O CACHE_VERSION ficou parado em 03/06/2026
     e, com ele, o CSS de junho — o cabecalho azul do modal, o modal horizontal, o
     total em vermelho e o botao desabilitado nunca chegaram a quem ja tinha aberto o
     app. O usuario relatou como "algo se perdeu"; nada se perdeu, o SW nao entregava.

     Um ritual manual ("bumpa a versao a cada deploy") que ninguem executa por dois
     meses nao e mecanismo. Agora a regra depende do TIPO de arquivo, e nao da memoria
     de quem faz deploy:
       /vendor/ e icones -> cache-first (bibliotecas e imagens, trocam junto com versao)
       o resto (css, js) -> network-first com fallback no cache (offline continua vivo) */
  const imutavel = url.pathname.startsWith('/vendor/') ||
                   /\.(png|jpg|jpeg|svg|webp|woff2?|ttf)$/i.test(url.pathname);

  if (imutavel) {
    event.respondWith(
      caches.match(req).then((cached) => {
        if (cached) return cached;
        return fetch(req).then((resp) => {
          if (resp && resp.ok && resp.type === 'basic') {
            const copy = resp.clone();
            caches.open(SHELL_CACHE).then(c => c.put(req, copy));
          }
          return resp;
        }).catch(() => undefined);
      })
    );
    return;
  }

  /* Network-first: o servidor manda max-age=30 com must-revalidate, entao a rede
     responde 304 quase sempre — custa poucos bytes e mantem a tela sincronizada com
     o deploy. Sem rede, cai no cache e o app continua abrindo. */
  event.respondWith(
    fetch(req).then((resp) => {
      if (resp && resp.ok && resp.type === 'basic') {
        const copy = resp.clone();
        caches.open(SHELL_CACHE).then(c => c.put(req, copy));
      }
      return resp;
    }).catch(() => caches.match(req))
  );
});

// Permite o front pedir ativacao imediata em deploy novo
self.addEventListener('message', (event) => {
  if (event.data && event.data.type === 'SKIP_WAITING') {
    self.skipWaiting();
  }
});

// ============================================================================
// PUSH NOTIFICATIONS
// ============================================================================

// Recebe push do servidor — mostra notificacao nativa no SO
self.addEventListener('push', (event) => {
  let data = {};
  try {
    if (event.data) data = event.data.json();
  } catch (e) {
    // payload pode ser texto puro
    try { data = { title: 'Aprovacao NF Pronep', body: event.data.text() }; }
    catch (e2) { data = { title: 'Aprovacao NF Pronep', body: 'Nova atualizacao' }; }
  }

  const title = data.title || 'Aprovacao NF Pronep';
  const options = {
    body: data.body || '',
    icon: '/icon-192.png',
    badge: '/icon-192.png',
    tag: data.tag || 'nf-default',          // substitui notif anterior da mesma NF
    renotify: true,                          // toca/vibra de novo mesmo se substituindo
    requireInteraction: false,
    data: {
      url: data.url || '/',
      evento: data.evento,
      nfId: data.nfId,
      timestamp: data.timestamp || Date.now()
    }
  };
  // Acoes contextuais — aparece no formato "expanded" no Android/Desktop
  if (data.evento === 'lancada') {
    options.actions = [
      { action: 'open', title: 'Abrir' },
      { action: 'dismiss', title: 'Depois' }
    ];
  }

  event.waitUntil(self.registration.showNotification(title, options));
});

// Quando o user clica na notificacao — abre o app na URL certa
self.addEventListener('notificationclick', (event) => {
  event.notification.close();
  if (event.action === 'dismiss') return;
  const url = (event.notification.data && event.notification.data.url) || '/';
  const fullUrl = new URL(url, self.location.origin).href;

  event.waitUntil(
    self.clients.matchAll({ type: 'window', includeUncontrolled: true }).then((clientList) => {
      // Se ja tem janela do app aberta, foca nela
      for (const client of clientList) {
        try {
          const sameOrigin = new URL(client.url).origin === self.location.origin;
          if (sameOrigin && 'focus' in client) {
            client.focus();
            // Manda mensagem pro front saber qual URL/NF abrir
            if (client.postMessage) {
              client.postMessage({ type: 'push-click', url: url, nfId: event.notification.data && event.notification.data.nfId });
            }
            return;
          }
        } catch (e) {}
      }
      // Senao, abre nova janela
      if (self.clients.openWindow) return self.clients.openWindow(fullUrl);
    })
  );
});
