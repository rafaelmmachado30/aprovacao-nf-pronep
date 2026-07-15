/**
 * shared/graph.js — cliente Microsoft Graph + resolucao de site/lista, num lugar so.
 *
 * Consolida o boilerplate que estava duplicado em ~52 endpoints (getGraphClient)
 * e ~38 (resolveSite*). Comportamento IDENTICO ao das copias locais; a unica
 * diferenca e que o nome da lista agora e PARAMETRO em vez de constante do modulo.
 *
 * API:
 *   getGraphClient()                       -> Client Graph (autenticado por app)
 *   resolveSiteId(client)                  -> siteId (string), cacheado no processo
 *   resolveList(client, siteId, nomeLista) -> listId (string), cacheado por nome
 *   resolveSiteAndList(client, nomeLista)  -> { siteId, listId }
 *
 * Env vars exigidas (App Settings), as mesmas de sempre:
 *   AAD_TENANT_ID / AAD_CLIENT_ID / AAD_CLIENT_SECRET
 *   SHAREPOINT_SITE_HOSTNAME / SHAREPOINT_SITE_PATH
 *
 * Nota de compatibilidade: getGraphClient e SINCRONO (Client.initWithMiddleware
 * nao e assincrono). Chamadas antigas com `await getGraphClient()` continuam
 * funcionando — await sobre valor nao-Promise resolve pro proprio valor.
 */

require('isomorphic-fetch');
const { ClientSecretCredential } = require('@azure/identity');
const { Client } = require('@microsoft/microsoft-graph-client');
const { TokenCredentialAuthenticationProvider } =
  require('@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials');

// Cache em memoria do processo. O Azure Functions reaproveita a instancia entre
// invocacoes, entao siteId/listId sao resolvidos uma vez e reutilizados.
let _siteId = null;
const _listCache = Object.create(null); // displayName -> listId

function getGraphClient() {
  const tenantId = process.env.AAD_TENANT_ID;
  const clientId = process.env.AAD_CLIENT_ID;
  const clientSecret = process.env.AAD_CLIENT_SECRET;
  if (!tenantId || !clientId || !clientSecret) {
    throw new Error('App Settings AAD_* incompletas (AAD_TENANT_ID / AAD_CLIENT_ID / AAD_CLIENT_SECRET)');
  }
  const credential = new ClientSecretCredential(tenantId, clientId, clientSecret);
  const authProvider = new TokenCredentialAuthenticationProvider(credential, {
    scopes: ['https://graph.microsoft.com/.default']
  });
  return Client.initWithMiddleware({ authProvider });
}

async function resolveSiteId(client) {
  if (_siteId) return _siteId;
  const host = process.env.SHAREPOINT_SITE_HOSTNAME;
  const path = process.env.SHAREPOINT_SITE_PATH;
  if (!host || !path) {
    throw new Error('App Settings SHAREPOINT_* incompletas (SHAREPOINT_SITE_HOSTNAME / SHAREPOINT_SITE_PATH)');
  }
  const siteResp = await client.api(`/sites/${host}:${path}`).get();
  _siteId = siteResp.id;
  return _siteId;
}

async function resolveList(client, siteId, displayName) {
  if (_listCache[displayName]) return _listCache[displayName];
  const listsResp = await client
    .api(`/sites/${siteId}/lists`)
    .filter(`displayName eq '${displayName}'`)
    .get();
  if (!listsResp.value || !listsResp.value.length) {
    throw new Error(`Lista '${displayName}' nao encontrada`);
  }
  _listCache[displayName] = listsResp.value[0].id;
  return _listCache[displayName];
}

async function resolveSiteAndList(client, displayName) {
  const siteId = await resolveSiteId(client);
  const listId = await resolveList(client, siteId, displayName);
  return { siteId, listId };
}

module.exports = { getGraphClient, resolveSiteId, resolveList, resolveSiteAndList };
