/**
 * /api/GetControleAcessos (GET) — ADMIN ONLY
 *
 * Retorna os dados da tela Controle de Acessos a contratos:
 *   {
 *     diretorias: ['Tecnologia','RH',...],          // pastas de contrato (diretorias)
 *     gestores:   [{ nome, email }, ...],            // pool atribuivel
 *     acessos:    { 'Tecnologia': ['a@x','b@x'], ... } // mapa atual (config explicita)
 *   }
 *
 * O pool de gestores = aprovadores cadastrados na PRONEP-NF-Diretorias + membros do
 * grupo Financeiro-Gestao (pra permitir liberar o Financeiro, conforme decisao do Rafa).
 */

require('isomorphic-fetch');
const { requireAdmin } = require('../shared/authz');
const { lerMapaAcessos } = require('../shared/acessoContratos');
const { ClientSecretCredential } = require('@azure/identity');
const { Client } = require('@microsoft/microsoft-graph-client');
const { TokenCredentialAuthenticationProvider } =
  require('@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials');

const LIST_DIR = 'PRONEP-NF-Diretorias';
const FINANCEIRO_GROUP_ID = process.env.GESTOR_FINANCEIRO_GROUP_ID || 'c2a73d16-4659-4b3c-93a1-0c0fbfaaaa96';

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
    if (!authz) return; // requireAdmin ja setou 401/403

    const client = getGraphClient();
    const siteId = await resolveSite(client);
    const listDirId = await resolveListId(client, siteId, LIST_DIR);

    // 1) Diretorias (pastas) + pool de gestores a partir da matriz de NF
    const diretoriasSet = new Set();
    const poolMap = {}; // email -> nome
    if (listDirId) {
      const resp = await client.api('/sites/' + siteId + '/lists/' + listDirId + '/items?expand=fields&$top=300').get();
      for (const it of (resp.value || [])) {
        const f = it.fields || {};
        const dir = (String(f.Title || '').split('|')[1] || '').trim();
        if (dir) diretoriasSet.add(dir);
        const email = String(f.field_3 || '').toLowerCase().trim();
        const nome = String(f.field_4 || '').trim();
        if (email) poolMap[email] = poolMap[email] || nome || email;
      }
    }

    // 2) Adiciona membros do grupo Financeiro-Gestao ao pool (pra poder liberar Financeiro)
    try {
      const grp = await client.api('/groups/' + FINANCEIRO_GROUP_ID + '/members')
        .select('id,displayName,mail,userPrincipalName').top(100).get();
      for (const u of (grp.value || [])) {
        const email = String(u.mail || u.userPrincipalName || '').toLowerCase().trim();
        if (email && !poolMap[email]) poolMap[email] = u.displayName || email;
      }
    } catch (e) { /* grupo opcional — segue sem ele */ }

    // 3) Mapa atual de acessos
    const acessos = await lerMapaAcessos(client, siteId, null);

    const gestores = Object.keys(poolMap)
      .map(function (em) { return { email: em, nome: poolMap[em] }; })
      .sort(function (a, b) { return String(a.nome).localeCompare(String(b.nome)); });
    const diretorias = Array.from(diretoriasSet).sort(function (a, b) { return a.localeCompare(b); });

    context.res = {
      status: 200,
      headers: { 'Content-Type': 'application/json', 'Cache-Control': 'no-store' },
      body: { ok: true, diretorias: diretorias, gestores: gestores, acessos: acessos }
    };
  } catch (err) {
    context.log && context.log.error && context.log.error('GetControleAcessos:', err);
    context.res = { status: 500, headers: { 'Content-Type': 'application/json' },
      body: { error: (err && err.message) || String(err) } };
  }
};
