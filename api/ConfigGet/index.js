/**
 * /api/ConfigGet
 *
 * Retorna config global do sistema. Schema:
 *   {
 *     multiNivel: {
 *       habilitado: false,
 *       valorLimite: 0,
 *       modoAprovador: 'global' | 'porDiretoria',
 *       gestorMasterGlobal: 'email@pronep.com.br',
 *       gestoresPorDiretoria: { 'Tecnologia': 'email', ... }
 *     }
 *   }
 *
 * Se nao existe config salva, retorna defaults (multi-nivel desabilitado).
 */

require('isomorphic-fetch');
const { getUser } = require('../shared/auth');
const { ClientSecretCredential } = require('@azure/identity');
const { Client } = require('@microsoft/microsoft-graph-client');
const { TokenCredentialAuthenticationProvider } =
  require('@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials');

const LIST_NAME = 'PRONEP-NF-Config';
const CONFIG_TITLE = 'global';
const cache = { siteId: null, listId: null };

const DEFAULT_CONFIG = {
  multiNivel: {
    habilitado: false,
    valorLimite: 0,
    modoAprovador: 'global',
    gestorMasterGlobal: '',
    gestoresPorDiretoria: {}
  }
};

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
    .filter("displayName eq '" + LIST_NAME + "'")
    .get();
  if (!lists.value || !lists.value.length) {
    throw new Error("Lista '" + LIST_NAME + "' nao encontrada");
  }
  cache.listId = lists.value[0].id;
  return cache;
}

async function findConfigItem(client, siteId, listId) {
  const resp = await client.api('/sites/' + siteId + '/lists/' + listId + '/items')
    .expand('fields')
    .filter("fields/Title eq '" + CONFIG_TITLE + "'")
    .top(1)
    .get();
  return (resp.value && resp.value[0]) || null;
}

module.exports = async function (context, req) {
  try {
    const user = await getUser(req);
    if (!user) {
      context.res = { status: 401, body: { error: 'Nao autenticado' } };
      return;
    }

    const client = await getGraphClient();
    const { siteId, listId } = await resolveSiteAndList(client);
    const item = await findConfigItem(client, siteId, listId);

    let config = DEFAULT_CONFIG;
    if (item && item.fields && item.fields.ConfigJson) {
      try { config = JSON.parse(item.fields.ConfigJson); }
      catch (e) { context.log && context.log.warn && context.log.warn('ConfigJson invalido, usando defaults'); }
    }

    context.res = {
      status: 200,
      headers: { 'Content-Type': 'application/json' },
      body: { ok: true, config: config, itemId: item ? item.id : null }
    };
  } catch (err) {
    context.log && context.log.error && context.log.error('ConfigGet error:', err);
    context.res = {
      status: 500,
      headers: { 'Content-Type': 'application/json' },
      body: { error: (err && err.message) || String(err), config: DEFAULT_CONFIG }
    };
  }
};
