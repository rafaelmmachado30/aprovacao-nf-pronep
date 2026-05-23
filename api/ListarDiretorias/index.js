/**
 * Sistema de Aprovacao de NF - ListarDiretorias
 *
 * Le a lista PRONEP-NF-Diretorias no SharePoint via Graph API.
 * Retorna a matriz Unidade x Diretoria com aprovador responsavel.
 *
 * Uso pelo front:
 *   - Tela Lancamento: descobre quem aprova (Unidade + Diretoria do fornecedor)
 *   - Tela Mapa de Aprovadores: mostra a matriz completa
 */

require('isomorphic-fetch');
const { ClientSecretCredential } = require('@azure/identity');
const { Client } = require('@microsoft/microsoft-graph-client');
const { TokenCredentialAuthenticationProvider } =
  require('@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials');

const LIST_NAME = 'PRONEP-NF-Diretorias';
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
  if (!host || !path) throw new Error('SHAREPOINT_* obrigatorias');
  const siteResp = await client.api(`/sites/${host}:${path}`).get();
  cache.siteId = siteResp.id;
  const listsResp = await client
    .api(`/sites/${cache.siteId}/lists`)
    .filter(`displayName eq '${LIST_NAME}'`)
    .get();
  if (!listsResp.value || !listsResp.value.length) {
    throw new Error(`Lista '${LIST_NAME}' nao encontrada`);
  }
  cache.listId = listsResp.value[0].id;
  return cache;
}

module.exports = async function (context, req) {
  const diag = { step: 'start' };
  try {
    diag.step = 'graph_client';
    const client = await getGraphClient();

    diag.step = 'resolve_site';
    const { siteId, listId } = await resolveSiteAndList(client);
    diag.siteId = siteId; diag.listId = listId;

    diag.step = 'fetch_items';
    const all = [];
    let url = `/sites/${siteId}/lists/${listId}/items?expand=fields&$top=500`;
    while (url) {
      const resp = await client.api(url).get();
      all.push(...(resp.value || []));
      url = resp['@odata.nextLink']
        ? resp['@odata.nextLink'].replace('https://graph.microsoft.com/v1.0','')
        : null;
    }
    diag.totalItems = all.length;

    diag.step = 'transform';
    diag.rawFieldsDebug = all.slice(0,3).map(item => item.fields);
    const diretorias = all.map(item => {
      const f = item.fields || {};
      return {
        id: item.id,
        chave: f.Title || '',
        unidade: f.Unidade || '',
        diretoria: f.Diretoria || '',
        aprovadorEmail: f.AprovadorEmail || '',
        aprovadorNome: f.AprovadorNome || '',
        grupoEntraId: f.GrupoEntraId || ''
      };
    });

    // Tambem retorna em formato dicionario (chave -> aprovador) pra consumo direto
    const mapa = {};
    diretorias.forEach(d => {
      if (d.chave) mapa[d.chave] = { email: d.aprovadorEmail, nome: d.aprovadorNome };
    });

    context.res = {
      status: 200,
      headers: { 'Content-Type': 'application/json', 'Cache-Control': 'no-store' },
      body: { total: diretorias.length, diretorias, mapa, diag }
    };
  } catch (err) {
    context.log && context.log.error && context.log.error('ListarDiretorias error:', err);
    context.res = {
      status: 500,
      headers: { 'Content-Type': 'application/json' },
      body: {
        error: (err && err.message) || String(err),
        statusCode: err && err.statusCode,
        diag
      }
    };
  }
};
