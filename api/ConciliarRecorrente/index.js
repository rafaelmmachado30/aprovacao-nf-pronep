/**
 * /api/ConciliarRecorrente (POST) — admin ou gestor da diretoria da conta.
 *
 * "Merge" do Fechamento do Mes: vincula manualmente uma NF Aprovada a uma conta
 * recorrente NAQUELE mes, para resolver o falso "atrasada" quando o casamento
 * automatico nao acha a NF (filial de CNPJ / diretoria diferente).
 *
 * Body:
 *   { action:'set', chave, mes:'AAAA-MM', diretoria, nota:{ id, numero, status, valor } }
 *   { action:'remover', chave, mes:'AAAA-MM', diretoria }
 *
 * O vinculo vale SO para o mes informado (chave@@mes). Nao "aprende" meses futuros.
 */

require('isomorphic-fetch');
const { resolveAuthz } = require('../shared/authz');
const { ROLE_LABELS } = require('../shared/userRoles');
const { _norm } = require('../shared/recorrentes');
const { lerConciliacoes, salvarConciliacoes, resolveConfigListId, chaveConc } = require('../shared/conciliacaoRecorrentes');
const { registrar: auditRegistrar } = require('../shared/auditLog');
const { getGraphClient, resolveSiteId } = require('../shared/graph');

module.exports = async function (context, req) {
  try {
    const authz = await resolveAuthz(req);
    if (!authz) { context.res = { status: 401, body: { error: 'Nao autenticado' } }; return; }

    const body = req.body || {};
    const action = body.action || 'set';
    const chave = String(body.chave || '').trim();
    const mes = String(body.mes || '').trim();
    const diretoria = String(body.diretoria || '').trim();
    if (!chave || !/^\d{4}-\d{2}$/.test(mes)) {
      context.res = { status: 400, body: { error: 'Envie chave e mes (AAAA-MM) validos' } };
      return;
    }

    // RBAC: admin OU gestor da diretoria da conta.
    const gestorLabels = (authz.roles || []).filter(function (r) { return String(r).indexOf('gestor') === 0; })
      .map(function (r) { return ROLE_LABELS[r]; }).filter(Boolean);
    const podeNaDiretoria = authz.isAdmin || gestorLabels.some(function (l) { return _norm(l) === _norm(diretoria); });
    if (!podeNaDiretoria) {
      context.res = { status: 403, body: { error: 'Voce nao gerencia a diretoria ' + (diretoria || '(vazia)') } };
      return;
    }

    const client = getGraphClient();
    const siteId = await resolveSiteId(client);
    const cfgId = await resolveConfigListId(client, siteId);
    if (!cfgId) { context.res = { status: 500, body: { error: "Lista 'PRONEP-NF-Config' nao encontrada" } }; return; }

    const mapa = await lerConciliacoes(client, siteId, cfgId);
    const k = chaveConc(chave, mes);

    if (action === 'remover') {
      delete mapa[k];
    } else {
      const nota = body.nota || {};
      if (!nota.id) { context.res = { status: 400, body: { error: 'Para conciliar, envie nota.id' } }; return; }
      mapa[k] = {
        notaId: String(nota.id),
        numero: String(nota.numero || ''),
        status: String(nota.status || 'Aprovada'),
        valor: Number(nota.valor || 0) || 0,
        por: authz.email,
        em: new Date().toISOString()
      };
    }

    const r = await salvarConciliacoes(client, siteId, cfgId, mapa);

    auditRegistrar(authz.user, 'config_update',
      { tipo: 'conciliacao_recorrente', id: r.itemId },
      'sucesso',
      { action: action, chave: chave, mes: mes, nota: (body.nota && body.nota.id) || null }
    ).catch(function () {});

    context.res = {
      status: 200, headers: { 'Content-Type': 'application/json' },
      body: { ok: true, action: action, chave: chave, mes: mes }
    };
  } catch (err) {
    context.log && context.log.error && context.log.error('ConciliarRecorrente:', err);
    context.res = { status: 500, headers: { 'Content-Type': 'application/json' },
      body: { error: (err && err.message) || String(err) } };
  }
};
