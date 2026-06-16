/**
 * /api/GetControleAcessos (GET) — ADMIN ONLY
 *
 * Dados da tela Controle de Acessos (modelo POR GRUPO/diretoria):
 *   {
 *     folders: ['Comercial','Tecnologia','Particulares',...]      // pastas reais do acervo
 *     grupos:  [{ role:'gestor_tecnologia', label:'Tecnologia' }, ...] // grupos atribuiveis (Entra)
 *     acessos: { 'Comercial': ['gestor_juridica'], ... }          // mapa atual (pasta -> roles)
 *   }
 *
 * As pessoas vem da pertinencia ao grupo do Entra. Por isso a tela trabalha com GRUPOS
 * (lista canonica e limpa), nao com os nomes soltos da matriz/acervo.
 */

require('isomorphic-fetch');
const { requireAdmin } = require('../shared/authz');
const { lerMapaAcessos } = require('../shared/acessoContratos');
const { gruposContrato } = require('../shared/userRoles');
const { ClientSecretCredential } = require('@azure/identity');
const { Client } = require('@microsoft/microsoft-graph-client');
const { TokenCredentialAuthenticationProvider } =
  require('@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials');

const LIST_CONTRATOS = 'PRONEP-NF-Contratos';
const LIST_DIR = 'PRONEP-NF-Diretorias';

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

    // Pastas de contrato (dropdown) = diretorias distintas no acervo de contratos.
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
    } catch (e) { /* sem acervo acessivel */ }

    // Fallback: se nao achou pastas no acervo, usa as diretorias da matriz como pastas.
    if (!folderSet.size) {
      try {
        const listDirId = await resolveListId(client, siteId, LIST_DIR);
        if (listDirId) {
          const resp = await client.api('/sites/' + siteId + '/lists/' + listDirId + '/items?expand=fields&$top=300').get();
          for (const it of (resp.value || [])) {
            const dir = (String((it.fields || {}).Title || '').split('|')[1] || '').trim();
            if (dir) folderSet.add(dir);
          }
        }
      } catch (e) { /* ignora */ }
    }

    const acessos = await lerMapaAcessos(client, siteId, null);
    const folders = Array.from(folderSet).sort(function (a, b) { return a.localeCompare(b); });
    const grupos = gruposContrato(); // lista canonica [{role,label}] dos grupos do Entra

    context.res = {
      status: 200,
      headers: { 'Content-Type': 'application/json', 'Cache-Control': 'no-store' },
      body: { ok: true, folders: folders, grupos: grupos, acessos: acessos }
    };
  } catch (err) {
    context.log && context.log.error && context.log.error('GetControleAcessos:', err);
    context.res = { status: 500, headers: { 'Content-Type': 'application/json' },
      body: { error: (err && err.message) || String(err) } };
  }
};
