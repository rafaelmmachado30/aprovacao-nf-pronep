/**
 * /api/RelatorioAcessosContratos (GET) — ADMIN ONLY
 *
 * Relatorio "quem ve o que": para cada PASTA de contrato, resolve os grupos/pessoas
 * liberados no Controle de Acessos e EXPANDE cada grupo nos membros reais do Entra.
 *
 * Resposta:
 * {
 *   ok: true,
 *   adminJuridicaVeemTudo: true,
 *   totalPessoasUnicas: N,
 *   pastas: [
 *     { pasta, configurada, fallback, grupos:[{role,label,membros:[{nome,email}]}], pessoas:[email] }
 *   ]
 * }
 *
 * Reusa a mesma logica de descoberta de pastas do GetControleAcessos e de membros do
 * ListarMembrosGrupo. Admin e Juridico veem tudo (nota global — nao repetido por pasta).
 */

require('isomorphic-fetch');
const { requireAdmin } = require('../shared/authz');
const { lerMapaAcessos, folderParaRole } = require('../shared/acessoContratos');
const { ROLE_LABELS, roleParaGrupoId } = require('../shared/userRoles');
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

function _norm(s) {
  return String(s || '').toLowerCase().normalize('NFD').replace(/\p{Diacritic}/gu, '').replace(/[^a-z0-9]/g, '');
}

async function descobrirPastas(client, siteId) {
  const set = new Set();
  try {
    const listId = await resolveListId(client, siteId, LIST_CONTRATOS);
    if (listId) {
      const cols = await client.api('/sites/' + siteId + '/lists/' + listId + '/columns').get();
      let cDir = 'Diretoria';
      for (const c of (cols.value || [])) { if (c.displayName === 'Diretoria' && c.name) cDir = c.name; }
      let url = '/sites/' + siteId + '/lists/' + listId + '/items?expand=fields&$top=999';
      let pages = 0;
      while (url && pages < 50) {
        const r = await client.api(url).get();
        for (const it of (r.value || [])) {
          const d = String((it.fields || {})[cDir] || '').trim();
          if (d) set.add(d);
        }
        pages++;
        const nx = r['@odata.nextLink'];
        url = nx ? nx.replace('https://graph.microsoft.com/v1.0', '') : null;
      }
    }
  } catch (e) { /* ignora */ }
  if (!set.size) {
    try {
      const listDirId = await resolveListId(client, siteId, LIST_DIR);
      if (listDirId) {
        const resp = await client.api('/sites/' + siteId + '/lists/' + listDirId + '/items?expand=fields&$top=300').get();
        for (const it of (resp.value || [])) {
          const dir = (String((it.fields || {}).Title || '').split('|')[1] || '').trim();
          if (dir) set.add(dir);
        }
      }
    } catch (e) { /* ignora */ }
  }
  return Array.from(set).sort(function (a, b) { return a.localeCompare(b); });
}

async function membrosDoRole(client, role, cache) {
  if (cache[role]) return cache[role];
  const gid = roleParaGrupoId(role);
  const membros = [];
  if (gid) {
    let url = '/groups/' + gid + '/members?$select=id,displayName,mail,userPrincipalName&$top=100';
    let pages = 0;
    while (url && pages < 20) {
      const resp = await client.api(url).get();
      for (const u of (resp.value || [])) {
        const email = String(u.mail || u.userPrincipalName || '').toLowerCase().trim();
        if (email) membros.push({ nome: u.displayName || email, email: email });
      }
      pages++;
      url = resp['@odata.nextLink'] ? resp['@odata.nextLink'].replace('https://graph.microsoft.com/v1.0', '') : null;
    }
    membros.sort(function (a, b) { return String(a.nome).localeCompare(String(b.nome)); });
  }
  cache[role] = membros;
  return membros;
}

module.exports = async function (context, req) {
  try {
    const authz = await requireAdmin(context, req);
    if (!authz) return;

    const client = getGraphClient();
    const siteId = await resolveSite(client);
    const pastas = await descobrirPastas(client, siteId);
    const mapa = await lerMapaAcessos(client, siteId, null);
    const roleKeys = Object.keys(ROLE_LABELS);
    const cache = {};
    const pessoasUnicas = new Set();

    const out = [];
    for (const pasta of pastas) {
      // acha config da pasta (case/acento-insensitive)
      let tokens = null;
      for (const k of Object.keys(mapa || {})) {
        if (_norm(k) === _norm(pasta) && Array.isArray(mapa[k])) { tokens = mapa[k]; break; }
      }
      const configurada = Array.isArray(tokens);
      // fallback: pasta sem config -> grupo de mesmo nome
      let fallback = false;
      if (!configurada) {
        const ownRole = folderParaRole(pasta);
        tokens = ownRole ? [ownRole] : [];
        fallback = !!ownRole;
      }

      const grupos = [];
      const pessoas = [];
      for (const tk of (tokens || [])) {
        const t = String(tk || '').trim();
        if (!t) continue;
        if (t.indexOf('@') >= 0) {
          pessoas.push(t.toLowerCase());
          pessoasUnicas.add(t.toLowerCase());
        } else {
          // role (grupo inteiro)
          const roleKey = roleKeys.indexOf(t) >= 0 ? t : (roleKeys.find(function (r) { return r.toLowerCase() === t.toLowerCase(); }) || t);
          const membros = await membrosDoRole(client, roleKey, cache);
          membros.forEach(function (m) { pessoasUnicas.add(m.email); });
          grupos.push({ role: roleKey, label: ROLE_LABELS[roleKey] || roleKey, membros: membros });
        }
      }

      out.push({
        pasta: pasta,
        configurada: configurada,
        fallback: fallback,
        grupos: grupos,
        pessoas: pessoas,
        semAcessoExtra: (grupos.length === 0 && pessoas.length === 0)
      });
    }

    context.res = {
      status: 200,
      headers: { 'Content-Type': 'application/json', 'Cache-Control': 'no-store' },
      body: {
        ok: true,
        adminJuridicaVeemTudo: true,
        totalPastas: out.length,
        totalPessoasUnicas: pessoasUnicas.size,
        pastas: out
      }
    };
  } catch (err) {
    context.log && context.log.error && context.log.error('RelatorioAcessosContratos:', err);
    context.res = { status: 500, headers: { 'Content-Type': 'application/json' },
      body: { error: (err && err.message) || String(err) } };
  }
};
