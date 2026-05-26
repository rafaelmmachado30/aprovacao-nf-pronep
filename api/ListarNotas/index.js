/**
 * Sistema de Aprovacao de NF - ListarNotas
 *
 * Le PRONEP-NF-NotasFiscais via Graph e aplica filtro RBAC server-side:
 *  - admin / financeiro      -> ve tudo
 *  - gestor                  -> ve onde AprovadorAtual = email do user
 *  - submitter               -> ve onde LancadoPor = email do user
 *
 * Query: ?status=Lancada|Aprovada|Rejeitada ?unidade=RJ|SP|ES ?diretoria=...
 */

require('isomorphic-fetch');
const { getUser } = require('../shared/auth');
const { ClientSecretCredential } = require('@azure/identity');
const { Client } = require('@microsoft/microsoft-graph-client');
const { TokenCredentialAuthenticationProvider } =
  require('@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials');

const LIST_NOTAS = 'PRONEP-NF-NotasFiscais';
const cache = { siteId: null, listId: null, colMap: null, invColMap: null };

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
  if (!host || !path) throw new Error('SHAREPOINT_* incompletas');
  const siteResp = await client.api(`/sites/${host}:${path}`).get();
  cache.siteId = siteResp.id;
  const listsResp = await client.api(`/sites/${cache.siteId}/lists`).filter(`displayName eq '${LIST_NOTAS}'`).get();
  if (!listsResp.value || !listsResp.value.length) throw new Error(`Lista ${LIST_NOTAS} nao encontrada`);
  cache.listId = listsResp.value[0].id;
  return cache;
}

async function getColumnMap(client, siteId, listId) {
  if (cache.colMap) return cache.colMap;
  const resp = await client.api(`/sites/${siteId}/lists/${listId}/columns`).get();
  const map = {};
  for (const col of (resp.value || [])) {
    if (!col.displayName || !col.name) continue;
    if (col.readOnly === true) continue;
    if (col.hidden === true) continue;
    if (col.name.startsWith('_')) continue;
    if (['LinkTitle','LinkTitleNoMenu','Edit','DocIcon','ItemChildCount',
         'FolderChildCount','AppAuthor','AppEditor','Attachments'].includes(col.name)) continue;
    map[col.displayName] = col.name;
  }
  cache.colMap = map;
  cache.invColMap = {};
  for (const [k, v] of Object.entries(map)) cache.invColMap[v] = k;
  return map;
}

function readClientPrincipal(req) {
  const header = req.headers && req.headers['x-ms-client-principal'];
  if (!header) return null;
  try { return JSON.parse(Buffer.from(header, 'base64').toString('utf-8')); }
  catch (e) { return null; }
}

function readClientPrincipalRoles(req) {
  const principal = readClientPrincipal(req);
  return (principal && principal.userRoles) || [];
}

function normalizeItem(item, invColMap) {
  const f = item.fields || {};
  const out = { id: item.id };
  for (const [internal, val] of Object.entries(f)) {
    const display = invColMap[internal];
    if (display) out[display] = val;
  }
  return out;
}

module.exports = async function (context, req) {
  const diag = { step: 'start' };
  try {
    diag.step = 'principal';
    const user = await getUser(req);
    if (!user) {
      context.res = { status: 401, body: { error: 'Nao autenticado' } };
      return;
    }
    const userEmail = (user.email || '').toLowerCase();
    const userRoles = (user.claims && user.claims.roles) || (user.source === 'easy-auth' ? readClientPrincipalRoles(req) : []);
    diag.userEmail = userEmail;
    diag.userRoles = userRoles;

    diag.step = 'graph';
    const client = await getGraphClient();
    const { siteId, listId } = await resolveSiteAndList(client);

    diag.step = 'columns';
    await getColumnMap(client, siteId, listId);
    const invColMap = cache.invColMap;

    diag.step = 'fetch';
    const all = [];
    let url = `/sites/${siteId}/lists/${listId}/items?expand=fields&$top=500`;
    let pages = 0;
    while (url) {
      const resp = await client.api(url).get();
      all.push(...(resp.value || []));
      pages++;
      url = resp['@odata.nextLink']
        ? resp['@odata.nextLink'].replace('https://graph.microsoft.com/v1.0','')
        : null;
      if (pages >= 30) break;
    }
    diag.totalLidos = all.length;

    diag.step = 'normalize';
    let notas = all.map(item => normalizeItem(item, invColMap));

    diag.step = 'rbac';
    const qTodos = (req.query && req.query.todos) === '1';
    const isAdmin = userRoles.includes('administrador') || userEmail === 'rafael.machado@pronep.com.br';
    const isFinanceiro = userRoles.includes('financeiro_nf');
    const isGestor = userRoles.some(r => r.startsWith('gestor'));
    diag.userScope = { isAdmin, isFinanceiro, isGestor };

    let notasFiltradas;
    if (isAdmin || isFinanceiro || qTodos) {
      notasFiltradas = notas;
    } else if (isGestor) {
      notasFiltradas = notas.filter(n => (n.AprovadorAtual || '').toLowerCase() === userEmail);
    } else {
      notasFiltradas = notas.filter(n => (n.LancadoPor || '').toLowerCase() === userEmail);
    }

    // Filtros query string adicionais
    const qStatus    = (req.query && req.query.status) || '';
    const qUnidade   = (req.query && req.query.unidade) || '';
    const qDiretoria = (req.query && req.query.diretoria) || '';
    if (qStatus)    notasFiltradas = notasFiltradas.filter(n => (n.Status || '').toLowerCase() === qStatus.toLowerCase());
    if (qUnidade)   notasFiltradas = notasFiltradas.filter(n => (n.Unidade || '') === qUnidade);
    if (qDiretoria) notasFiltradas = notasFiltradas.filter(n => (n.Diretoria || '') === qDiretoria);

    // Ordena: pendentes primeiro, depois mais recentes
    notasFiltradas.sort((a, b) => {
      const sa = (a.Status || ''); const sb = (b.Status || '');
      if (sa === 'Lancada' && sb !== 'Lancada') return -1;
      if (sb === 'Lancada' && sa !== 'Lancada') return 1;
      return (b.LancadoEm || '').localeCompare(a.LancadoEm || '');
    });

    context.res = {
      status: 200,
      headers: { 'Content-Type': 'application/json', 'Cache-Control': 'no-store' },
      body: {
        total: notasFiltradas.length,
        totalAntesFiltros: notas.length,
        diag,
        notas: notasFiltradas
      }
    };
  } catch (err) {
    context.log && context.log.error && context.log.error('ListarNotas error:', err);
    context.res = {
      status: 500,
      headers: { 'Content-Type': 'application/json' },
      body: {
        error: (err && err.message) || String(err),
        statusCode: err && err.statusCode,
        body: err && err.body,
        diag
      }
    };
  }
};
