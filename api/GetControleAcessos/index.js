/**
 * /api/GetControleAcessos (GET) — ADMIN ONLY
 *
 * Dados da tela Controle de Acessos (modelo POR DIRETORIA):
 *   {
 *     folders:    ['Comercial','Tecnologia','Suprimentos',...]  // pastas (acervo de contratos)
 *     diretorias: ['Tecnologia','RH',...]                       // pool de diretorias atribuiveis
 *     acessos:    { 'Tecnologia': ['Tecnologia','Comercial'], ... } // mapa atual (pasta -> diretorias)
 *   }
 *
 * As PESSOAS vem automaticamente da matriz de gestores (PRONEP-NF-Diretorias): quem for
 * gestor de uma diretoria liberada ve a pasta. Por isso a tela trabalha com diretorias, nao e-mails.
 */

require('isomorphic-fetch');
const { requireAdmin } = require('../shared/authz');
const { lerMapaAcessos } = require('../shared/acessoContratos');
const { ClientSecretCredential } = require('@azure/identity');
const { Client } = require('@microsoft/microsoft-graph-client');
const { TokenCredentialAuthenticationProvider } =
  require('@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials');

const LIST_DIR = 'PRONEP-NF-Diretorias';
const LIST_CONTRATOS = 'PRONEP-NF-Contratos';

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

async function resolveSite(client) {
  const host = process.env.SHAREPOINT_SITE_HOSTNAME;
  const path = process.env.SHAREPOINT_SITE_PATH;
  if (!host || !path) throw new Error('SHAREPOINT_* incompletas');
  const siteResp = await client.api('/sites/' + host + ':' + path).get();
  return siteResp.id;
}

async function resolveListId(client, siteId, displayName) {
  const lists = await client.api('/sites/' + siteId + '/lists')
    .filter("displayName eq '" + displayName + "'").get();
  return (lists.value && lists.value.length) ? lists.value[0].id : null;
}

module.exports = async function (context, req) {
  try {
    const authz = await requireAdmin(context, req);
    if (!authz) return;

    const client = getGraphClient();
    const siteId = await resolveSite(client);

    // 1) Diretorias da matriz (pool atribuivel)
    const dirSet = new Set();
    const listDirId = await resolveListId(client, siteId, LIST_DIR);
    if (listDirId) {
      const resp = await client.api('/sites/' + siteId + '/lists/' + listDirId + '/items?expand=fields&$top=300').get();
      for (const it of (resp.value || [])) {
        const dir = (String((it.fields || {}).Title || '').split('|')[1] || '').trim();
        if (dir) dirSet.add(dir);
      }
    }

    // 2) Pastas de contrato (dropdown) = diretorias distintas no acervo de contratos
    const folderSet = new Set();
    try {
      const listContrId = await resolveListId(client, siteId, LIST_CONTRATOS);
      if (listContrId) {
        const cols = await client.api('/sites/' + siteId + '/lists/' + listContrId + '/columns').get();
        let cDir = 'Diretoria';
        for (const c of (cols.value || [])) { if (c.displayName === 'Diretoria' && c.name) cDir = c.name; }
        let url = '/sites/' + siteId + '/lists/' + listContrId + '/items?expand=fields&$top=999';
        let pages = 0;
        while (url && pages < 50) {
          const r = await client.api(url).get();
          for (const it of (r.value || [])) {
            const d = String((it.fields || {})[cDir] || '').trim();
            if (d) folderSet.add(d);
          }
          pages++;
          const nx = r['@odata.nextLink'];
          url = nx ? nx.replace('https://graph.microsoft.com/v1.0', '') : null;
        }
      }
    } catch (e) { /* sem acervo acessivel -> usa a matriz como pastas */ }
    if (!folderSet.size) { for (const d of dirSet) folderSet.add(d); }

    // 3) Pool de diretorias atribuiveis = uniao (matriz + pastas existentes)
    const poolSet = new Set();
    for (const d of dirSet) poolSet.add(d);
    for (const d of folderSet) poolSet.add(d);

    const acessos = await lerMapaAcessos(client, siteId, null);

    const folders = Array.from(folderSet).sort(function (a, b) { return a.localeCompare(b); });
    const diretorias = Array.from(poolSet).sort(function (a, b) { return a.localeCompare(b); });

    context.res = {
      status: 200,
      headers: { 'Content-Type': 'application/json', 'Cache-Control': 'no-store' },
      body: { ok: true, folders: folders, diretorias: diretorias, acessos: acessos }
    };
  } catch (err) {
    context.log && context.log.error && context.log.error('GetControleAcessos:', err);
    context.res = { status: 500, headers: { 'Content-Type': 'application/json' },
      body: { error: (err && err.message) || String(err) } };
  }
};
