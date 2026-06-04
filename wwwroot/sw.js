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

const CACHE_VERSION = 'pronep-nf-v3-fixdesconhecido-20260603';
const SHELL_CACHE = `${CACHE_VERSION}-shell`;
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
      return Promise.allSettled(SHELL_ASSETS.map(url => cache.add(url).catch(() => null)));
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
          // Cacheia copia fresca do index pra futuro fallback offline
          if (resp && resp.ok) {
            const copy = resp.clone();
            caches.open(SHELL_CACHE).then(c => c.put('/index.html', copy));
          }
          return resp;
        })
        .catch(() => caches.match('/index.html').then(c => c || caches.match('/offline.html')))
    );
    return;
  }

  // Assets estaticos (vendor, icones, css): cache-first com fallback pra rede
  event.respondWith(
    caches.match(req).then((cached) => {
      if (cached) return cached;
      return fetch(req).then((resp) => {
        // Cacheia se for resposta ok do mesmo origem
        if (resp && resp.ok && resp.type === 'basic') {
          const copy = resp.clone();
          caches.open(SHELL_CACHE).then(c => c.put(req, copy));
        }
        return resp;
      }).catch(() => undefined);
    })
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
