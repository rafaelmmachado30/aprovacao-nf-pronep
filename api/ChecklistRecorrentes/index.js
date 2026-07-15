/**
 * /api/ChecklistRecorrentes  (GET)  — Estagio 2.
 *
 * Para cada conta RECORRENTE confirmada (PRONEP-NF-Recorrentes, EhRecorrente=Sim,
 * Ativo, dentro da DataFim), cruza com as NFs do mes-alvo e calcula o status:
 *   - 'aprovada'   : NF do mes ja aprovada (ou integrada/processada)
 *   - 'lancada'    : NF do mes lancada, aguardando aprovacao (fila)
 *   - 'atrasada'   : sem NF e o vencimento esperado JA passou
 *   - 'risco'      : sem NF e o vencimento entra no prazo D+5 (<=5 dias uteis)
 *   - 'aguardando' : sem NF e ainda nao chegou a hora
 *   - 'conciliada' : gestor vinculou manualmente uma NF aprovada (Merge)
 *
 * O CALCULO vive em shared/checklistRecorrentes.js (computar), reutilizado tambem
 * pela SAN (tool fechamento_pendencias). Este endpoint so faz RBAC + escopo +
 * envelope da resposta.
 *
 * RBAC: admin (todas) | gestor (sua(s) diretoria(s)) | demais 403.
 * Query: ?diretoria=Tecnologia ?mes=AAAA-MM (default: mes corrente, BRT)
 */

require('isomorphic-fetch');
const { getUser } = require('../shared/auth');
const { getUserRoles, ROLE_LABELS } = require('../shared/userRoles');
const { isAdminEmail } = require('../shared/authz');
const { _norm } = require('../shared/recorrentes');
const { getGraphClient } = require('../shared/graph');
const { computar } = require('../shared/checklistRecorrentes');

const LIST_NOTAS = 'PRONEP-NF-NotasFiscais';
const cache = { siteId: null, listNotasId: null, invColMap: null };

async function resolveSite(client) {
  if (cache.siteId && cache.listNotasId) return cache;
  const host = process.env.SHAREPOINT_SITE_HOSTNAME;
  const path = process.env.SHAREPOINT_SITE_PATH;
  if (!host || !path) throw new Error('SHAREPOINT_* incompletas');
  const siteResp = await client.api('/sites/' + host + ':' + path).get();
  cache.siteId = siteResp.id;
  const lists = await client.api('/sites/' + cache.siteId + '/lists').filter("displayName eq '" + LIST_NOTAS + "'").get();
  if (!lists.value || !lists.value.length) throw new Error('Lista ' + LIST_NOTAS + ' nao encontrada');
  cache.listNotasId = lists.value[0].id;
  return cache;
}

async function getInvColMap(client, siteId, listId) {
  if (cache.invColMap) return cache.invColMap;
  const resp = await client.api('/sites/' + siteId + '/lists/' + listId + '/columns').get();
  const inv = {};
  for (const c of (resp.value || [])) { if (c.displayName && c.name) inv[c.name] = c.displayName; }
  cache.invColMap = inv;
  return inv;
}

function readClientPrincipalRoles(req) {
  const header = req.headers && req.headers['x-ms-client-principal'];
  if (!header) return [];
  try { const p = JSON.parse(Buffer.from(header, 'base64').toString('utf-8')); return (p && p.userRoles) || []; }
  catch (e) { return []; }
}

module.exports = async function (context, req) {
  const diag = { step: 'start' };
  try {
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
    if (!isAdmin && gestorLabels.length === 0) {
      context.res = { status: 403, body: { error: 'Apenas gestores ou admin' } }; return;
    }

    const qDiretoria = (req.query && req.query.diretoria) || '';
    let scopeDiretorias = null;
    if (qDiretoria && qDiretoria !== '__todas__') scopeDiretorias = [qDiretoria];
    else if (qDiretoria === '__todas__') scopeDiretorias = null;
    else if (isAdmin) scopeDiretorias = null;
    else if (gestorLabels.length) scopeDiretorias = gestorLabels;
    const scopeNorm = scopeDiretorias ? scopeDiretorias.map(_norm) : null;

    // Mes-alvo (BRT). Default: mes corrente.
    const agoraBRT = new Date(Date.now() - 3 * 60 * 60 * 1000);
    let ano = agoraBRT.getUTCFullYear();
    let mes = agoraBRT.getUTCMonth(); // 0-based
    const qMes = (req.query && req.query.mes) || '';
    const mMes = qMes.match(/^(\d{4})-(\d{2})$/);
    if (mMes) { ano = Number(mMes[1]); mes = Number(mMes[2]) - 1; }
    diag.mesAlvo = ano + '-' + String(mes + 1).padStart(2, '0');

    diag.step = 'compute';
    const client = getGraphClient();
    const { siteId, listNotasId } = await resolveSite(client);
    const invColMap = await getInvColMap(client, siteId, listNotasId);
    const r = await computar(client, siteId, listNotasId, invColMap, { scopeNorm: scopeNorm, ano: ano, mes: mes });

    if (r.listaDecisoesExiste === false) {
      context.res = { status: 200, headers: { 'Content-Type': 'application/json' },
        body: { ok: true, listaDecisoesExiste: false, mes: r.mes, contas: [], resumo: {}, diretoriasDisponiveis: [] } };
      return;
    }

    diag.step = 'done';
    context.res = {
      status: 200, headers: { 'Content-Type': 'application/json', 'Cache-Control': 'no-store' },
      body: {
        ok: true, listaDecisoesExiste: true, mes: r.mes,
        scope: { isAdmin: isAdmin, diretorias: scopeDiretorias },
        diretoriasDisponiveis: r.diretoriasDisponiveis || [],
        total: r.total, resumo: r.resumo, contas: r.contas, diag: diag
      }
    };
  } catch (err) {
    context.log && context.log.error && context.log.error('ChecklistRecorrentes error:', err);
    context.res = { status: 500, headers: { 'Content-Type': 'application/json' },
      body: { error: (err && err.message) || String(err), diag: diag } };
  }
};
