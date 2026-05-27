/**
 * /api/MarcarProcessado (POST)
 *
 * Marca/desmarca uma NF como Processado (financeiro liberou pra integracao).
 * Body: { id: '123', processado: true|false }
 */

require('isomorphic-fetch');
const { getUser } = require('../shared/auth');
const { ClientSecretCredential } = require('@azure/identity');
const { Client } = require('@microsoft/microsoft-graph-client');
const { TokenCredentialAuthenticationProvider } =
  require('@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials');

const LIST_NOTAS = 'PRONEP-NF-NotasFiscais';
const cache = { siteId: null, listId: null };

async function getGraphClient() {
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

async function resolveSiteAndList(client) {
  if (cache.siteId && cache.listId) return cache;
  const host = process.env.SHAREPOINT_SITE_HOSTNAME;
  const path = process.env.SHAREPOINT_SITE_PATH;
  const siteResp = await client.api('/sites/' + host + ':' + path).get();
  cache.siteId = siteResp.id;
  const lists = await client.api('/sites/' + cache.siteId + '/lists')
    .filter("displayName eq '" + LIST_NOTAS + "'").get();
  if (!lists.value || !lists.value.length) throw new Error("Lista nao encontrada");
  cache.listId = lists.value[0].id;
  return cache;
}

module.exports = async function (context, req) {
  try {
    const user = await getUser(req);
    if (!user) {
      context.res = { status: 401, body: { error: 'Nao autenticado' } };
      return;
    }
    const body = req.body || {};
    const itemId = body.id;
    const processado = !!body.processado;
    if (!itemId) {
      context.res = { status: 400, body: { error: 'id obrigatorio' } };
      return;
    }
    const client = await getGraphClient();
    const { siteId, listId } = await resolveSiteAndList(client);
    await client.api('/sites/' + siteId + '/lists/' + listId + '/items/' + itemId + '/fields')
      .patch({ Processado: processado });
    context.res = {
      status: 200,
      headers: { 'Content-Type': 'application/json' },
      body: { ok: true, id: itemId, processado: processado, por: user.email }
    };
  } catch (err) {
    context.log && context.log.error && context.log.error('MarcarProcessado:', err);
    context.res = {
      status: 500,
      headers: { 'Content-Type': 'application/json' },
      body: {
        error: (err && err.message) || String(err),
        statusCode: err && err.statusCode,
        graphBody: err && err.body
      }
    };
  }
};
