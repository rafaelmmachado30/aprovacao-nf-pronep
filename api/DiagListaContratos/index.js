/**
 * /api/DiagListaContratos
 *
 * Diagnostico DETALHADO da lista PRONEP-NF-Contratos. Le propriedades raw
 * (name interno, hidden, webUrl), lista os primeiros 10 itens, e busca
 * tambem na LIXEIRA do SP (caso a lista tenha sido excluida e ainda esteja
 * recuperavel).
 *
 * NAO chama Claude. NAO crawl. Pura inspecao Graph.
 *
 * RBAC: admin only.
 */

require('isomorphic-fetch');
const { ClientSecretCredential } = require('@azure/identity');
const { Client } = require('@microsoft/microsoft-graph-client');
const { TokenCredentialAuthenticationProvider } =
  require('@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials');

const LIST_NAME = 'PRONEP-NF-Contratos';

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

function readClientPrincipal(req) {
  const header = req.headers && req.headers['x-ms-client-principal'];
  if (!header) return null;
  try { return JSON.parse(Buffer.from(header, 'base64').toString('utf-8')); } catch (e) { return null; }
}

async function isAdmin(req) {
  const p = readClientPrincipal(req);
  const roles = (p && p.userRoles) || [];
  if (roles.includes('administrador') || roles.includes('admin')) return true;
  try {
    const { getUser } = require('../shared/auth');
    const user = await getUser(req);
    if (!user || !user.oid) return false;
    const { getUserRoles } = require('../shared/userRoles');
    const userRoles = await getUserRoles(user);
    return (userRoles || []).includes('administrador');
  } catch (e) { return false; }
}

module.exports = async function (context, req) {
  try {
    if (!(await isAdmin(req))) {
      context.res = { status: 403, body: { error: 'Apenas admin' } };
      return;
    }
    const client = await getGraphClient();
    const host = process.env.SHAREPOINT_SITE_HOSTNAME;
    const path = process.env.SHAREPOINT_SITE_PATH;
    const siteResp = await client.api('/sites/' + host + ':' + path).get();
    const siteId = siteResp.id;

    const diag = { site: { siteId, webUrl: siteResp.webUrl }, etapas: {} };

    // 1. Listar TODAS as listas do site, mostrando name + displayName + hidden
    diag.etapas.todasListas = [];
    try {
      const all = await client.api('/sites/' + siteId + '/lists').select('id,name,displayName,webUrl,list,createdDateTime').get();
      for (const l of (all.value || [])) {
        diag.etapas.todasListas.push({
          id: l.id,
          name: l.name,
          displayName: l.displayName,
          webUrl: l.webUrl,
          template: l.list && l.list.template,
          hidden: l.list && l.list.hidden,
          contentTypesEnabled: l.list && l.list.contentTypesEnabled,
          createdDateTime: l.createdDateTime
        });
      }
    } catch (e) {
      diag.etapas.todasListas_erro = e.message;
    }

    // 2. Procurar PRONEP-NF-Contratos por displayName E por name
    diag.etapas.listasComPrefixoPronep = diag.etapas.todasListas.filter(function(l){
      return (l.displayName || '').toLowerCase().indexOf('contrato') >= 0 ||
             (l.name || '').toLowerCase().indexOf('contrato') >= 0 ||
             (l.displayName || '').toLowerCase().indexOf('pronep-nf') >= 0;
    });

    // 3. GET direto pelo ID conhecido
    const listIdConhecido = 'd5742aea-cc86-4a19-b354-aee6ce2f8b0f';
    try {
      const detalhe = await client.api('/sites/' + siteId + '/lists/' + listIdConhecido).expand('columns,items($top=10;$expand=fields)').get();
      diag.etapas.listaPorId = {
        id: detalhe.id,
        name: detalhe.name,
        displayName: detalhe.displayName,
        webUrl: detalhe.webUrl,
        template: detalhe.list && detalhe.list.template,
        hidden: detalhe.list && detalhe.list.hidden,
        createdDateTime: detalhe.createdDateTime,
        itemCount: (detalhe.items || []).length,
        primeirosItens: (detalhe.items || []).map(function(it){
          return {
            id: it.id,
            createdDateTime: it.createdDateTime,
            fields: it.fields ? {
              Title: it.fields.Title,
              Diretoria: it.fields.Diretoria,
              Unidade: it.fields.Unidade,
              Fornecedor: it.fields.Fornecedor,
              Status: it.fields.Status,
              DataInicio: it.fields.DataInicio,
              DataFim: it.fields.DataFim,
              DriveItemId: it.fields.DriveItemId
            } : null
          };
        })
      };
    } catch (e) {
      diag.etapas.listaPorId_erro = { error: e.message, statusCode: e.statusCode, code: e.code };
    }

    // 4. Contar itens com GET separado (caso o $expand items nao puxe todos)
    try {
      const itemsResp = await client.api('/sites/' + siteId + '/lists/' + listIdConhecido + '/items?$top=999&$expand=fields').get();
      diag.etapas.itemsTotal = (itemsResp.value || []).length;
      diag.etapas.itemsSample = (itemsResp.value || []).slice(0, 5).map(function(it){
        return {
          id: it.id,
          fields: it.fields ? {
            Title: it.fields.Title,
            Diretoria: it.fields.Diretoria,
            Unidade: it.fields.Unidade,
            Fornecedor: it.fields.Fornecedor,
            Status: it.fields.Status,
            DataInicio: it.fields.DataInicio,
            DataFim: it.fields.DataFim,
            DriveItemId: it.fields.DriveItemId
          } : null
        };
      });
    } catch (e) {
      diag.etapas.itemsTotal_erro = { error: e.message, statusCode: e.statusCode, code: e.code };
    }

    // 5. Olhar a Lixeira do site (recycleBin)
    try {
      const rb = await client.api('/sites/' + siteId + '/drive/items/root/children').get();
      diag.etapas.driveRoot = (rb.value || []).map(function(it){ return { name: it.name, isFolder: !!it.folder }; });
    } catch (e) {
      diag.etapas.driveRoot_erro = e.message;
    }

    context.res = {
      status: 200,
      headers: { 'Content-Type': 'application/json' },
      body: diag
    };
  } catch (err) {
    context.res = {
      status: 500,
      headers: { 'Content-Type': 'application/json' },
      body: { error: err.message, stack: (err.stack || '').split('\n').slice(0, 8) }
    };
  }
};
