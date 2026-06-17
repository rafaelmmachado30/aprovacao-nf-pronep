/**
 * /api/SalvarRecorrente  (POST)
 *
 * Grava a decisao do gestor sobre uma conta recorrente (idempotente por
 * CNPJ+Diretoria+Unidade) na lista PRONEP-NF-Recorrentes.
 *
 * Body JSON: {
 *   cnpj, fornecedor, diretoria, unidade,
 *   ehRecorrente: true|false,
 *   dataFim: 'AAAA-MM-DD' | '' (opcional, fim do lembrete),
 *   diaVencimento, valorEstimado, ativo
 * }
 *
 * RBAC: admin OU gestor da diretoria informada. Demais: 403.
 */

require('isomorphic-fetch');
const { getUser } = require('../shared/auth');
const { getUserRoles, ROLE_LABELS } = require('../shared/userRoles');
const { isAdminEmail } = require('../shared/authz');
const { salvarDecisao, _norm } = require('../shared/recorrentes');
const { registrar: auditRegistrar } = require('../shared/auditLog');
const { ClientSecretCredential } = require('@azure/identity');
const { Client } = require('@microsoft/microsoft-graph-client');
const { TokenCredentialAuthenticationProvider } =
  require('@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials');

const cache = { siteId: null };

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

async function resolveSiteId(client) {
  if (cache.siteId) return cache.siteId;
  const host = process.env.SHAREPOINT_SITE_HOSTNAME;
  const path = process.env.SHAREPOINT_SITE_PATH;
  if (!host || !path) throw new Error('SHAREPOINT_* incompletas');
  const siteResp = await client.api('/sites/' + host + ':' + path).get();
  cache.siteId = siteResp.id;
  return cache.siteId;
}

function readClientPrincipalRoles(req) {
  const header = req.headers && req.headers['x-ms-client-principal'];
  if (!header) return [];
  try { const p = JSON.parse(Buffer.from(header, 'base64').toString('utf-8')); return (p && p.userRoles) || []; }
  catch (e) { return []; }
}

module.exports = async function (context, req) {
  try {
    if (!req.body || typeof req.body !== 'object') {
      context.res = { status: 400, body: { error: 'Body invalido (JSON esperado)' } };
      return;
    }
    const user = await getUser(req);
    if (!user) { context.res = { status: 401, body: { error: 'Nao autenticado' } }; return; }
    const userEmail = (user.email || '').toLowerCase();

    const claimsRoles = (user.claims && user.claims.roles) || [];
    const principalRoles = (user.source === 'easy-auth' ? readClientPrincipalRoles(req) : []) || [];
    const usefulPrincipalRoles = principalRoles.filter(function (r) { return r !== 'authenticated' && r !== 'anonymous'; });
    const graphRoles = await getUserRoles(user);
    const userRoles = Array.from(new Set([].concat(claimsRoles, usefulPrincipalRoles, graphRoles)));
    const isAdmin = userRoles.includes('administrador') || isAdminEmail(userEmail);
    const gestorLabels = userRoles.filter(function (r) { return r.indexOf('gestor') === 0; })
      .map(function (r) { return ROLE_LABELS[r]; }).filter(Boolean);

    const b = req.body;
    const diretoria = String(b.diretoria || '').trim();
    const cnpj = String(b.cnpj || '').replace(/\D/g, '');
    if (!cnpj || !diretoria) {
      context.res = { status: 400, body: { error: 'cnpj e diretoria sao obrigatorios' } };
      return;
    }
    // RBAC: admin ou gestor DA diretoria informada
    const podeNaDiretoria = isAdmin || gestorLabels.some(function (l) { return _norm(l) === _norm(diretoria); });
    if (!podeNaDiretoria) {
      context.res = { status: 403, body: { error: 'Voce nao gerencia a diretoria ' + diretoria } };
      return;
    }

    const client = getGraphClient();
    const siteId = await resolveSiteId(client);

    const resultado = await salvarDecisao(client, siteId, null, {
      cnpj: cnpj,
      fornecedor: b.fornecedor,
      diretoria: diretoria,
      unidade: b.unidade,
      ehRecorrente: (b.ehRecorrente === true || b.ehRecorrente === 'true' || b.ehRecorrente === 'Sim'),
      diaVencimento: b.diaVencimento,
      valorEstimado: b.valorEstimado,
      dataFim: b.dataFim,
      ativo: !(b.ativo === false || b.ativo === 'Nao'),
      confirmadoPor: userEmail
    });

    auditRegistrar(user, 'recorrente_decisao',
      { tipo: 'recorrente', id: resultado.chave, numero: b.fornecedor || cnpj },
      'sucesso',
      { diretoria: diretoria, unidade: b.unidade, ehRecorrente: b.ehRecorrente, dataFim: b.dataFim || null, action: resultado.action }
    ).catch(function () {});

    context.res = { status: 200, headers: { 'Content-Type': 'application/json' }, body: resultado };
  } catch (err) {
    context.log && context.log.error && context.log.error('SalvarRecorrente error:', err);
    context.res = { status: 500, headers: { 'Content-Type': 'application/json' }, body: { error: (err && err.message) || String(err) } };
  }
};
