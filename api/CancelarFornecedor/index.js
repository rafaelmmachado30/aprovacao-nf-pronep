/**
 * POST /api/CancelarFornecedor
 *
 * Marca TODOS os contratos de um fornecedor como Cancelado (ou Reativa) em batch.
 *
 * Body:
 *   {
 *     diretoria: "Tecnologia",
 *     unidade: "CORPORATIVO",
 *     fornecedor: "ABS FATURAMENTO HOSPITALAR",
 *     acao: "cancelar" | "reativar"
 *   }
 *
 * Comportamento:
 *   - cancelar:  para cada contrato (Diretoria+Unidade+Fornecedor) — set Status=Cancelado,
 *                e grava o Status anterior em Observacoes (linha "_statusAntes=<X>") pra
 *                permitir reativar com o status original
 *   - reativar:  para cada contrato CANCELADO do fornecedor — restaura Status:
 *                  1) se tem "_statusAntes=<X>" em Observacoes, usa X
 *                  2) senao, recalcula pela DataFim (calcularStatus)
 *
 * RBAC: admin OU gestor da diretoria.
 * Custo Claude: zero.
 */

const { ClientSecretCredential } = require('@azure/identity');
const { Client } = require('@microsoft/microsoft-graph-client');
const { TokenCredentialAuthenticationProvider } = require('@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials');
const { getUser } = require('../shared/auth');
const { getUserRoles } = require('../shared/userRoles');
const contratosShared = require('../shared/contratos');
const auditLog = require('../shared/auditLog');

const cache = { siteId: null, listId: null, listDirId: null, colMap: null };

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

async function getColMap(client, siteId, listId) {
  if (cache.colMap) return cache.colMap;
  const r = await client.api('/sites/' + siteId + '/lists/' + listId + '/columns').get();
  const m = {};
  for (const c of (r.value || [])) if (c.displayName && c.name) m[c.displayName] = c.name;
  cache.colMap = m;
  return m;
}

function readClientPrincipal(req) {
  const h = req.headers && req.headers['x-ms-client-principal'];
  if (!h) return null;
  try { return JSON.parse(Buffer.from(h, 'base64').toString('utf-8')); } catch (e) { return null; }
}

async function diretoriasDoGestor(client, siteId, listDirId, userEmail) {
  if (!listDirId || !userEmail) return [];
  const r = await client.api('/sites/' + siteId + '/lists/' + listDirId + '/items?expand=fields&$top=200').get();
  const set = new Set();
  for (const it of (r.value || [])) {
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
    // 1. Auth
    const user = await getUser(req);
    if (!user || !user.email) {
      context.res = { status: 401, body: { error: 'Nao autenticado' } };
      return;
    }
    const claimsRoles = (readClientPrincipal(req) && readClientPrincipal(req).userRoles) || [];
    const userRoles = await getUserRoles(user);
    const allRoles = Array.from(new Set([].concat(claimsRoles, userRoles || [])));
    const isAdmin = allRoles.includes('administrador') || allRoles.includes('admin');
    const isGestor = allRoles.includes('gestor');
    if (!isAdmin && !isGestor) {
      context.res = { status: 403, body: { error: 'Acesso restrito a Gestores e Admins.' } };
      return;
    }

    // 2. Parse body
    const body = req.body || {};
    const diretoria = String(body.diretoria || '').trim();
    const unidade = String(body.unidade || '').trim();
    const fornecedor = String(body.fornecedor || '').trim();
    const acao = String(body.acao || '').toLowerCase().trim();
    if (!diretoria || !unidade || !fornecedor) {
      context.res = { status: 400, body: { error: 'diretoria, unidade e fornecedor obrigatorios' } };
      return;
    }
    if (acao !== 'cancelar' && acao !== 'reativar') {
      context.res = { status: 400, body: { error: 'acao deve ser "cancelar" ou "reativar"' } };
      return;
    }

    // 3. Resolve site/lists
    const client = getGraphClient();
    const { siteId, listId, listDirId } = await resolveSiteELists(client);
    if (!listId) {
      context.res = { status: 500, body: { error: 'Lista PRONEP-NF-Contratos nao encontrada' } };
      return;
    }

    // 4. RBAC: gestor so pode mexer nas SUAS diretorias
    if (!isAdmin && isGestor) {
      const minhasDir = await diretoriasDoGestor(client, siteId, listDirId, user.email);
      if (!minhasDir.includes(diretoria)) {
        context.res = { status: 403, body: { error: 'Voce nao eh gestor da diretoria ' + diretoria } };
        return;
      }
    }

    // 5. Carrega TODOS os contratos do fornecedor (paginando)
    const colMap = await getColMap(client, siteId, listId);
    const cD = colMap['Diretoria'] || 'Diretoria';
    const cU = colMap['Unidade'] || 'Unidade';
    const cF = colMap['Fornecedor'] || 'Fornecedor';
    const cSt = colMap['Status'] || 'Status';
    const cDF = colMap['DataFim'] || 'DataFim';
    const cObs = colMap['Observacoes'] || 'Observacoes';

    const todos = [];
    let nextUrl = '/sites/' + siteId + '/lists/' + listId + '/items?expand=fields&$top=999';
    let pages = 0;
    while (nextUrl && pages < 50) {
      pages++;
      const r = await client.api(nextUrl).get();
      for (const it of (r.value || [])) {
        const f = it.fields || {};
        if ((f[cD] || '') === diretoria && (f[cU] || '') === unidade && (f[cF] || '') === fornecedor) {
          todos.push({ id: it.id, fields: f });
        }
      }
      const next = r['@odata.nextLink'];
      if (next) {
        const idx = next.indexOf('/v1.0/');
        nextUrl = idx >= 0 ? next.substring(idx + 5) : null;
      } else nextUrl = null;
    }

    if (!todos.length) {
      context.res = { status: 404, body: { error: 'Nenhum contrato encontrado para ' + diretoria + '/' + unidade + '/' + fornecedor } };
      return;
    }

    // 6. Aplica a acao
    const resultados = { ok: 0, erros: [], total: todos.length };
    for (const c of todos) {
      try {
        const statusAtual = c.fields[cSt] || '';
        const obsAtuais = c.fields[cObs] || '';
        let novoStatus, novasObs;

        if (acao === 'cancelar') {
          if (statusAtual === 'Cancelado') { resultados.ok++; continue; } // ja estava
          novoStatus = 'Cancelado';
          // Preserva o status anterior em Observacoes pra permitir reativar
          const obsLinha = '_statusAntes=' + statusAtual;
          if (obsAtuais.indexOf('_statusAntes=') < 0) {
            novasObs = (obsAtuais ? obsAtuais + '\n' : '') + obsLinha;
          } else {
            novasObs = obsAtuais;  // ja tinha registro anterior, preserva
          }
        } else { // reativar
          if (statusAtual !== 'Cancelado') { resultados.ok++; continue; } // nao estava cancelado
          // Tenta extrair status anterior das observacoes
          const m = obsAtuais.match(/_statusAntes=([A-Za-z0-9]+)/);
          if (m && m[1] && m[1] !== 'Cancelado') {
            novoStatus = m[1];
          } else {
            // Recalcula via DataFim
            const dataFim = c.fields[cDF];
            const dfStr = dataFim ? String(dataFim).substring(0, 10) : null;
            novoStatus = contratosShared.calcularStatus(dfStr, false);
          }
          // Remove a linha _statusAntes pra evitar acumulo
          novasObs = obsAtuais.replace(/\n?_statusAntes=[A-Za-z0-9]+/g, '');
        }

        await client.api('/sites/' + siteId + '/lists/' + listId + '/items/' + c.id + '/fields')
          .patch({ [cSt]: novoStatus, [cObs]: novasObs });
        resultados.ok++;
      } catch (eItem) {
        resultados.erros.push({ id: c.id, error: eItem.message });
      }
    }

    // 7. Audit log
    auditLog.auditRegistrar(user, acao + '_fornecedor',
      { tipo: 'fornecedor', diretoria, unidade, fornecedor },
      resultados.erros.length ? 'parcial' : 'sucesso',
      { total: resultados.total, ok: resultados.ok, erros: resultados.erros.length }
    ).catch(function(){});

    context.res = {
      status: 200,
      headers: { 'Content-Type': 'application/json' },
      body: {
        ok: true,
        acao,
        diretoria, unidade, fornecedor,
        total: resultados.total,
        atualizados: resultados.ok,
        erros: resultados.erros
      }
    };
  } catch (err) {
    context.log && context.log.error && context.log.error('CancelarFornecedor:', err);
    context.res = { status: 500, body: { error: err.message, stack: (err.stack || '').split('\n').slice(0, 6) } };
  }
};
