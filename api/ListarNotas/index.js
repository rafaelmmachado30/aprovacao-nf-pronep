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
const { getUserRoles } = require('../shared/userRoles');
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
    // userRoles vem de 3 fontes mescladas:
    //  1. claims.roles do JWT (Teams SSO — Entra popula com grupos AAD nas claims)
    //  2. principal.userRoles do Easy Auth (SWA — so retorna ["authenticated","anonymous"], inutil)
    //  3. Graph API consultando /users/{user}/transitiveMemberOf (fonte de verdade)
    // SEMPRE chamamos o Graph (com cache 5min) porque o (2) NUNCA tem grupos AAD —
    // se confiar so nele, gestor/financeiro/ti caem no fluxo do submetedor e veem nada.
    const claimsRoles = (user.claims && user.claims.roles) || [];
    const principalRoles = (user.source === 'easy-auth' ? readClientPrincipalRoles(req) : []) || [];
    // Filtra as roles default do SWA que nao indicam grupo real
    const usefulPrincipalRoles = principalRoles.filter(r => r !== 'authenticated' && r !== 'anonymous');
    const graphRoles = await getUserRoles(user);
    // Mescla, dedupe
    const userRoles = Array.from(new Set([...claimsRoles, ...usefulPrincipalRoles, ...graphRoles]));
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
    // Pendentes (Lancada ou AguardandoN2) primeiro, depois ordenado por data
    function isPendente(s) { return s === 'Lancada' || s === 'AguardandoN2'; }
    notasFiltradas.sort((a, b) => {
      const pa = isPendente(a.Status); const pb = isPendente(b.Status);
      if (pa && !pb) return -1;
      if (pb && !pa) return 1;
      return (b.LancadoEm || '').localeCompare(a.LancadoEm || '');
    });

    // === ENRICHMENT: adiciona FornecedorRazao/FornecedorFantasia em cada nota ===
    // Lookup pela lista PRONEP-NF-Fornecedores (cache em memoria) — evita o front
    // depender de fornecedorIdx que pode bater errado se a lista local estiver defasada.
    diag.step = 'enrich_fornecedores';
    try {
      const fornListResp = await client.api(`/sites/${siteId}/lists`).filter(`displayName eq 'PRONEP-NF-Fornecedores'`).get();
      if (fornListResp.value && fornListResp.value.length) {
        const fornListId = fornListResp.value[0].id;
        // Carrega colunas + items uma vez
        const colResp = await client.api(`/sites/${siteId}/lists/${fornListId}/columns`).get();
        const fornInvMap = {};
        for (const c of (colResp.value || [])) { if (c.displayName && c.name) fornInvMap[c.name] = c.displayName; }
        const allForn = [];
        let urlF = `/sites/${siteId}/lists/${fornListId}/items?expand=fields&$top=500`;
        let fpages = 0;
        while (urlF && fpages < 30) {
          const r = await client.api(urlF).get();
          allForn.push(...(r.value || []));
          fpages++;
          urlF = r['@odata.nextLink'] ? r['@odata.nextLink'].replace('https://graph.microsoft.com/v1.0','') : null;
        }
        // Indexa por CNPJ normalizado (so digitos)
        const byDoc = {};
        for (const it of allForn) {
          const f = {}; for (const [k, v] of Object.entries(it.fields || {})) { const d = fornInvMap[k]; if (d) f[d] = v; }
          const doc = String(f.Documento || f.CNPJ || f.field_2 || '').replace(/\D/g, '');
          if (doc && !byDoc[doc]) {
            byDoc[doc] = {
              razao: f.Title || f.Razao || f.RazaoSocial || '',
              fantasia: f.NomeFantasia || f.field_3 || ''
            };
          }
        }
        diag.fornecedoresIndexados = Object.keys(byDoc).length;
        // Popular cada nota
        for (const n of notasFiltradas) {
          const doc = String(n.CNPJFornecedor || '').replace(/\D/g, '');
          const hit = doc && byDoc[doc];
          if (hit) {
            n.FornecedorRazao = hit.razao;
            n.FornecedorFantasia = hit.fantasia;
          }
        }
      }
    } catch (e) {
      diag.enrichErro = e.message;
    }

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
