/**
 * /api/Warmup (GET) — mantem o Function App "quente" (anti cold-start).
 *
 * As Functions do SWA (Consumption) hibernam apos ~20min ociosas; a 1a chamada
 * depois disso paga o cold start (carregar Node + node_modules + token Graph),
 * que e justamente o que faz o boot demorar (pior ainda dentro do Teams).
 *
 * Um cron (GitHub Actions / uptime monitor) chama este endpoint a cada poucos
 * minutos em horario comercial. Como as Functions do SWA compartilham o mesmo
 * processo, manter ESTE endpoint quente mantem TODAS quentes. Ele ainda toca o
 * Graph (resolve o site) pra aquecer credencial/token e o caminho de rede.
 *
 * Seguranca: se a App Setting WARMUP_SECRET existir, exige o header
 * X-Warmup-Secret igual. Se nao existir, funciona aberto (nao expoe dado algum).
 */

require('isomorphic-fetch');

module.exports = async function (context, req) {
  const required = process.env.WARMUP_SECRET;
  if (required) {
    const got = (req.headers && (req.headers['x-warmup-secret'] || req.headers['X-Warmup-Secret'])) || '';
    if (got !== required) {
      context.res = { status: 401, headers: { 'Content-Type': 'application/json' }, body: { error: 'unauthorized' } };
      return;
    }
  }

  const diag = { warmedModules: false, warmedSite: false };
  try {
    // Carrega os modulos pesados no processo (JIT/cache) — o maior custo do cold start.
    const { ClientSecretCredential } = require('@azure/identity');
    const { Client } = require('@microsoft/microsoft-graph-client');
    const { TokenCredentialAuthenticationProvider } =
      require('@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials');
    diag.warmedModules = true;

    const t = process.env.AAD_TENANT_ID, c = process.env.AAD_CLIENT_ID, s = process.env.AAD_CLIENT_SECRET;
    const host = process.env.SHAREPOINT_SITE_HOSTNAME, path = process.env.SHAREPOINT_SITE_PATH;
    if (t && c && s && host && path) {
      const cred = new ClientSecretCredential(t, c, s);
      const authProvider = new TokenCredentialAuthenticationProvider(cred, {
        scopes: ['https://graph.microsoft.com/.default']
      });
      const client = Client.initWithMiddleware({ authProvider });
      // Resolve o site: aquece token Graph + DNS + caminho de rede.
      await client.api('/sites/' + host + ':' + path).get();
      diag.warmedSite = true;
    }
  } catch (e) {
    diag.err = (e && e.message) || String(e);
  }

  context.res = {
    status: 200,
    headers: { 'Content-Type': 'application/json', 'Cache-Control': 'no-store' },
    body: { ok: true, warm: true, ts: new Date().toISOString(), diag: diag }
  };
};
