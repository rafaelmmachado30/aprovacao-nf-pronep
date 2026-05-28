/**
 * /api/PushUnsubscribe
 *
 * Remove a subscription do SharePoint.
 *
 * Body: { endpoint: string }
 *
 * Retorna:
 *   200 { ok, removidos }
 *   401 { error: "Nao autenticado" }
 *   400 { error: "endpoint obrigatorio" }
 */

require('isomorphic-fetch');
const { getUser } = require('../shared/auth');
const { removerSubscription } = require('../shared/pushNotif');
const { ClientSecretCredential } = require('@azure/identity');
const { Client } = require('@microsoft/microsoft-graph-client');
const { TokenCredentialAuthenticationProvider } =
  require('@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials');

function getGraphClient() {
  const tenantId = process.env.AAD_TENANT_ID;
  const clientId = process.env.AAD_CLIENT_ID;
  const clientSecret = process.env.AAD_CLIENT_SECRET;
  const credential = new ClientSecretCredential(tenantId, clientId, clientSecret);
  const authProvider = new TokenCredentialAuthenticationProvider(credential, {
    scopes: ['https://graph.microsoft.com/.default']
  });
  return Client.initWithMiddleware({ authProvider });
}

module.exports = async function (context, req) {
  try {
    const user = await getUser(req);
    if (!user) {
      context.res = { status: 401, headers: { 'Content-Type': 'application/json' },
        body: { error: 'Nao autenticado' } };
      return;
    }
    const endpoint = (req.body || {}).endpoint;
    if (!endpoint) {
      context.res = { status: 400, headers: { 'Content-Type': 'application/json' },
        body: { error: 'endpoint obrigatorio' } };
      return;
    }
    const client = getGraphClient();
    const host = process.env.SHAREPOINT_SITE_HOSTNAME;
    const path = process.env.SHAREPOINT_SITE_PATH;
    const siteResp = await client.api(`/sites/${host}:${path}`).get();
    const result = await removerSubscription(client, siteResp.id, endpoint);
    context.res = { status: 200, headers: { 'Content-Type': 'application/json' },
      body: { ok: true, ...result } };
  } catch (err) {
    context.res = { status: 500, headers: { 'Content-Type': 'application/json' },
      body: { error: (err && err.message) || String(err) } };
  }
};
