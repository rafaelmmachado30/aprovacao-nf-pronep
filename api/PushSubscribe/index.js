/**
 * /api/PushSubscribe
 *
 * Recebe a subscription do navegador e salva no SharePoint (PRONEP-NF-PushSubscriptions),
 * vinculada ao email do usuario autenticado.
 *
 * Body: { subscription: PushSubscription, userAgent?: string }
 *
 * Retorna:
 *   200 { ok, id, criado|atualizado }
 *   401 { error: "Nao autenticado" }
 *   400 { error: "subscription invalida" }
 *   500 { error, diag }
 */

require('isomorphic-fetch');
const { getUser } = require('../shared/auth');
const { salvarSubscription } = require('../shared/pushNotif');
const { ClientSecretCredential } = require('@azure/identity');
const { Client } = require('@microsoft/microsoft-graph-client');
const { TokenCredentialAuthenticationProvider } =
  require('@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials');

function getGraphClient() {
  const tenantId = process.env.AAD_TENANT_ID;
  const clientId = process.env.AAD_CLIENT_ID;
  const clientSecret = process.env.AAD_CLIENT_SECRET;
  if (!tenantId || !clientId || !clientSecret) throw new Error('AAD_* incompletas');
  const credential = new ClientSecretCredential(tenantId, clientId, clientSecret);
  const authProvider = new TokenCredentialAuthenticationProvider(credential, {
    scopes: ['https://graph.microsoft.com/.default']
  });
  return Client.initWithMiddleware({ authProvider });
}

async function resolveSiteId(client) {
  const host = process.env.SHAREPOINT_SITE_HOSTNAME;
  const path = process.env.SHAREPOINT_SITE_PATH;
  if (!host || !path) throw new Error('SHAREPOINT_* incompletas');
  const siteResp = await client.api(`/sites/${host}:${path}`).get();
  return siteResp.id;
}

module.exports = async function (context, req) {
  const diag = { step: 'start' };
  try {
    diag.step = 'auth';
    const user = await getUser(req);
    if (!user || !user.email) {
      context.res = { status: 401, headers: { 'Content-Type': 'application/json' },
        body: { error: 'Nao autenticado' } };
      return;
    }

    diag.step = 'parse_body';
    const body = req.body || {};
    const sub = body.subscription;
    if (!sub || !sub.endpoint || !sub.keys || !sub.keys.p256dh || !sub.keys.auth) {
      context.res = { status: 400, headers: { 'Content-Type': 'application/json' },
        body: { error: 'subscription invalida — precisa endpoint + keys.p256dh + keys.auth' } };
      return;
    }
    const userAgent = body.userAgent || (req.headers && req.headers['user-agent']) || '';

    diag.step = 'graph';
    const client = getGraphClient();
    const siteId = await resolveSiteId(client);

    diag.step = 'salvar';
    const result = await salvarSubscription(client, siteId, user.email, sub, userAgent);

    context.res = {
      status: 200,
      headers: { 'Content-Type': 'application/json' },
      body: { ok: true, ...result, email: user.email }
    };
  } catch (err) {
    if (context.log && context.log.error) context.log.error('PushSubscribe erro:', err);
    context.res = {
      status: 500,
      headers: { 'Content-Type': 'application/json' },
      body: { error: (err && err.message) || String(err), diag }
    };
  }
};
