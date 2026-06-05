/**
 * POST /api/AtualizarStatusContrato
 *
 * Body:
 *   { id: "<spListItemId>", status: "Ativo"|"Cancelado"|"Vencido"|... , observacoes?: "..." }
 *
 * RBAC: admin OU gestor da diretoria do contrato.
 *
 * Casos de uso:
 *  - Marcar contrato como Cancelado (pastas vermelhas no SP nao tem metadata Graph)
 *  - Reativar contrato cancelado
 *  - Editar observacoes
 */

const { ClientSecretCredential } = require('@azure/identity');
const { Client } = require('@microsoft/microsoft-graph-client');
const { TokenCredentialAuthenticationProvider } = require('@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials');
const { getUser } = require('../shared/auth');
const { getUserRoles } = require('../shared/userRoles');
const auditLog = require('../shared/auditLog');

const STATUS_VALIDOS = new Set(['Ativo', 'Cancelado', 'Vencido', 'Vencendo30', 'Vencendo60', 'Vencendo90', 'SemVigencia', 'Indeterminado']);

const cache = { siteId: null, listId: null, listDirId: null };

function getGraphClient() {
  const credential = new ClientSecretCredential(
    process.env.AAD_TENANT_ID, process.env.AAD_CLIENT_ID, process.env.AAD_CLIENT_SECRET
  );
  const authProvider = new TokenCredentialAuthenticationProvider(credential, {
    scopes: ['https://graph.microsoft.com/.default']
  });
  return Client.initWithMiddleware({ authProvider });
}

async function resolveSiteELists(client) {
  if (cache.siteId && cache.listId) return cache;
  const siteResp = await client.api('/sites/' + process.env.SHAREPOINT_SITE_HOSTNAME + ':' + process.env.SHAREPOINT_SITE_PATH).get();
  cache.siteId = siteResp.id;
  const lists = await client.api('/sites/' + cache.siteId + '/lists').get();
  for (const l of (lists.value || [])) {
    if (l.displayName === 'PRONEP-NF-Contratos') cache.listId = l.id;
    if (l.displayName === 'PRONEP-NF-Diretorias') cache.listDirId = l.id;
  }
  return cache;
}

function readClientPrincipal(req) {
  const header = req.headers && req.headers['x-ms-client-principal'];
  if (!header) return null;
  try { return JSON.parse(Buffer.from(header, 'base64').toString('utf-8')); } catch (e) { return null; }
}
function readClientPrincipalRoles(req) {
  const p = readClientPrincipal(req);
  return (p && p.userRoles) || [];
}

async function diretoriasDoGestor(client, siteId, listDirId, userEmail) {
  if (!listDirId || !userEmail) return [];
  const resp = await client.api('/sites/' + siteId + '/lists/' + listDirId + '/items?expand=fields&$top=200').get();
  const set = new Set();
  for (const it of (resp.value || [])) {
    const f = it.fields || {};
    const emailDir = String(f.field_3 || '').toLowerCase().trim();
    if (emailDir === userEmail) {
      const dir = String(f.Title || '').split('|')[1] || '';
      if (dir) set.add(dir.trim());
    }
  }
  return Array.from(set);
}

module.exports = async function (context, req) {
  try {
    const body = req.body || {};
    const id = body.id;
    const observacoes = body.observacoes;
    const dataInicio = body.dataInicio;  // 'YYYY-MM-DD' ou null
    const dataFim = body.dataFim;        // 'YYYY-MM-DD' ou null
    const valor = body.valor;             // number ou null
    let status = body.status;             // se vier 'auto' ou nada, recalcula com base na dataFim

    if (!id) {
      context.res = { status: 400, body: { error: 'id obrigatorio' } };
      return;
    }
    // status manual valida; mas se nao vier OU vier 'auto', recalcula com base em dataFim
    if (status && status !== 'auto' && !STATUS_VALIDOS.has(status)) {
      context.res = { status: 400, body: { error: 'status invalido. Validos: ' + Array.from(STATUS_VALIDOS).join(', ') + ' ou \'auto\'' } };
      return;
    }
    // Validacao de datas (formato YYYY-MM-DD)
    const dateRe = /^\d{4}-\d{2}-\d{2}$/;
    if (dataInicio && dataInicio !== '' && !dateRe.test(dataInicio)) {
      context.res = { status: 400, body: { error: 'dataInicio deve ser YYYY-MM-DD' } }; return;
    }
    if (dataFim && dataFim !== '' && !dateRe.test(dataFim)) {
      context.res = { status: 400, body: { error: 'dataFim deve ser YYYY-MM-DD' } }; return;
    }

    // Auth
    const user = await getUser(req);
    if (!user || !user.email) {
      context.res = { status: 401, body: { error: 'Nao autenticado' } };
      return;
    }

    // RBAC
    const claimsRoles = readClientPrincipalRoles(req) || [];
    const userRoles = await getUserRoles(user);
    const allRoles = Array.from(new Set([].concat(claimsRoles, userRoles || [])));
    const isAdmin = allRoles.includes('administrador') || allRoles.includes('admin');
    const isGestor = allRoles.includes('gestor');
    if (!isAdmin && !isGestor) {
      context.res = { status: 403, body: { error: 'Acesso negado' } };
      return;
    }

    const client = getGraphClient();
    const { siteId, listId, listDirId } = await resolveSiteELists(client);
    if (!listId) {
      context.res = { status: 404, body: { error: 'Lista PRONEP-NF-Contratos nao existe' } };
      return;
    }

    // Carrega item e diretoria pra checar escopo do gestor
    const item = await client.api('/sites/' + siteId + '/lists/' + listId + '/items/' + id + '?expand=fields').get();
    const diretoria = (item.fields && item.fields.Diretoria) || '';

    if (!isAdmin && isGestor) {
      const dirs = await diretoriasDoGestor(client, siteId, listDirId, user.email);
      if (!dirs.includes(diretoria)) {
        context.res = { status: 403, body: { error: 'Contrato fora do seu escopo de diretoria' } };
        return;
      }
    }

    // Recalcula status automaticamente se nao foi forcado E temos dataFim
    if ((!status || status === 'auto') && dataFim) {
      const contratos = require('../shared/contratos');
      status = contratos.calcularStatus(dataFim, false);
    } else if (!status) {
      // Se nao passou status nem dataFim, mantem o existente
      status = (item.fields && item.fields.Status) || 'Indeterminado';
    }

    // Constroi patch (so atualiza campos passados)
    const patch = { Status: status };
    if (observacoes !== undefined) patch.Observacoes = String(observacoes).slice(0, 30000);
    if (dataInicio !== undefined) patch.DataInicio = dataInicio ? (dataInicio + 'T00:00:00Z') : null;
    if (dataFim !== undefined) patch.DataFim = dataFim ? (dataFim + 'T00:00:00Z') : null;
    if (valor !== undefined) {
      if (valor === null || valor === '') {
        patch.ValorContrato = null;
      } else {
        const n = typeof valor === 'number' ? valor : parseFloat(valor);
        if (!isNaN(n) && n >= 0) patch.ValorContrato = n;
      }
    }
    // Marca como leitura manual quando user edita vigencia
    if (dataInicio !== undefined || dataFim !== undefined) {
      patch.LeituraIAStatus = 'manual';
    }

    await client.api('/sites/' + siteId + '/lists/' + listId + '/items/' + id + '/fields').patch(patch);

    // Audit
    auditLog.registrar(
      { oid: user.oid, email: user.email, name: user.name },
      'contrato_status',
      { tipo: 'contrato', id: id, numero: (item.fields && item.fields.Title) || '' },
      'sucesso',
      { statusNovo: status, statusAnterior: (item.fields && item.fields.Status) || '', diretoria, observacoes }
    ).catch(function(){});

    context.res = {
      status: 200,
      headers: { 'Content-Type': 'application/json' },
      body: { ok: true, id, status, observacoes: observacoes !== undefined ? observacoes : null }
    };
  } catch (err) {
    context.log && context.log.error && context.log.error('AtualizarStatusContrato error:', err);
    context.res = {
      status: 500,
      headers: { 'Content-Type': 'application/json' },
      body: { error: err.message, stack: (err.stack || '').split('\n').slice(0, 6) }
    };
  }
};
