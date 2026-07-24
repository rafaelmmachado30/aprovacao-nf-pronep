/**
 * /api/ConfirmarAlinhamento  (POST) — RBAC: Financeiro-Gestao (financeiro_nf) ou admin.
 *
 * Fecha o gate de alinhamento das NFs vencendo em < D+5. Uma NF em
 * Status=AguardandoAlinhamento (o aprovador declarou "alinhei com o financeiro")
 * so segue quando o Financeiro-Gestao age aqui:
 *   - acao='confirmar' -> executa a aprovacao REAL (reusa AprovarNota com bypass;
 *     carimba, move o PDF pra Aprovadas, Status=Aprovada, notifica o submitter).
 *   - acao='rejeitar'  -> rejeita a NF (reusa RejeitarNota; Status=Rejeitada, motivo).
 *
 * Ambos rodam em-processo com um principal do APROVADOR original (AprovadorAtual da NF),
 * porque o RBAC do AprovarNota/RejeitarNota valida pelo aprovador atribuido.
 *
 * Body: { id, acao: 'confirmar'|'rejeitar', motivo? }
 */

require('isomorphic-fetch');
const { resolveAuthz } = require('../shared/authz');
const { getGraphClient } = require('../shared/graph');
const { registrar: auditRegistrar } = require('../shared/auditLog');

const LIST_NAME = 'PRONEP-NF-NotasFiscais';

async function resolveSiteAndList(client) {
  const host = process.env.SHAREPOINT_SITE_HOSTNAME;
  const path = process.env.SHAREPOINT_SITE_PATH;
  const siteResp = await client.api('/sites/' + host + ':' + path).get();
  const siteId = siteResp.id;
  const lists = await client.api('/sites/' + siteId + '/lists').filter("displayName eq '" + LIST_NAME + "'").get();
  if (!lists.value || !lists.value.length) throw new Error('Lista ' + LIST_NAME + ' nao encontrada');
  return { siteId: siteId, listId: lists.value[0].id };
}

module.exports = async function (context, req) {
  const diag = { step: 'init' };
  try {
    const authz = await resolveAuthz(req);
    if (!authz) { context.res = { status: 401, headers: { 'Content-Type': 'application/json' }, body: { error: 'Nao autenticado' } }; return; }
    if (!authz.isAdmin && !authz.isFinanceiro) {
      context.res = { status: 403, headers: { 'Content-Type': 'application/json' }, body: { error: 'Acesso restrito ao Financeiro-Gestao (ou admin)' } };
      return;
    }

    const body = req.body || {};
    const itemId = String(body.id || '').trim();
    const acao = String(body.acao || '').trim().toLowerCase();
    const motivo = String(body.motivo || '').trim();
    if (!itemId) { context.res = { status: 400, headers: { 'Content-Type': 'application/json' }, body: { error: 'id obrigatorio' } }; return; }
    if (acao !== 'confirmar' && acao !== 'rejeitar') {
      context.res = { status: 400, headers: { 'Content-Type': 'application/json' }, body: { error: "acao deve ser 'confirmar' ou 'rejeitar'" } }; return;
    }

    const client = await getGraphClient();
    const { siteId, listId } = await resolveSiteAndList(client);

    diag.step = 'fetch_item';
    const item = await client.api('/sites/' + siteId + '/lists/' + listId + '/items/' + itemId + '?expand=fields').get();
    const fields = item.fields || {};
    const statusAtual = fields.Status || fields.field_9 || '';
    const aprovador = String(fields.AprovadorAtual || fields.field_8 || '').toLowerCase();
    const gestorFinanceiroAlinhado = fields.GestorFinanceiroAlinhado || '';

    if (statusAtual !== 'AguardandoAlinhamento') {
      context.res = { status: 409, headers: { 'Content-Type': 'application/json' },
        body: { error: 'NF nao esta aguardando alinhamento (status atual: ' + statusAtual + ').', status: statusAtual } };
      return;
    }
    if (!aprovador) {
      context.res = { status: 409, headers: { 'Content-Type': 'application/json' }, body: { error: 'NF sem aprovador atribuido.' } };
      return;
    }

    // Principal do APROVADOR original (o RBAC do AprovarNota/RejeitarNota valida por ele).
    const principalAprovador = Buffer.from(JSON.stringify({
      userDetails: aprovador, userId: aprovador, userRoles: ['authenticated']
    })).toString('base64');

    if (acao === 'confirmar') {
      diag.step = 'confirmar';
      const aprovarHandler = require('../AprovarNota/index.js');
      const ctxA = { res: null, log: { error: function () {}, info: function () {} } };
      await aprovarHandler(ctxA, {
        body: { id: itemId, bypassAlinhamento: true, alinhouFinanceiro: true, gestorFinanceiroAlinhado: gestorFinanceiroAlinhado },
        headers: { 'x-ms-client-principal': principalAprovador }, query: {}
      });
      const okA = ctxA.res && ctxA.res.status === 200 && ctxA.res.body && ctxA.res.body.ok;
      if (!okA) {
        context.res = { status: 502, headers: { 'Content-Type': 'application/json' },
          body: { error: 'Falha ao aprovar apos confirmacao.', detalhe: ctxA.res && ctxA.res.body, diag: diag } };
        return;
      }
      auditRegistrar(authz.user, 'alinhamento_confirmado', { tipo: 'nf', id: itemId }, 'sucesso',
        { confirmadoPor: authz.email, aprovadorOriginal: aprovador, gestorFinanceiroAlinhado: gestorFinanceiroAlinhado }).catch(function () {});
      context.res = { status: 200, headers: { 'Content-Type': 'application/json' },
        body: { ok: true, acao: 'confirmar', itemId: itemId, mensagem: 'Alinhamento confirmado — NF aprovada e seguiu para pagamento.' } };
      return;
    }

    // acao === 'rejeitar'
    diag.step = 'rejeitar';
    const rejeitarHandler = require('../RejeitarNota/index.js');
    const motivoFinal = 'Alinhamento com o financeiro NAO confirmado' + (motivo ? (' — ' + motivo) : '') + ' (por ' + authz.email + ')';
    const ctxR = { res: null, log: { error: function () {}, info: function () {} } };
    await rejeitarHandler(ctxR, {
      body: { id: itemId, motivo: motivoFinal, observacao: '' },
      headers: { 'x-ms-client-principal': principalAprovador }, query: {}
    });
    const okR = ctxR.res && ctxR.res.status === 200;
    if (!okR) {
      context.res = { status: 502, headers: { 'Content-Type': 'application/json' },
        body: { error: 'Falha ao rejeitar.', detalhe: ctxR.res && ctxR.res.body, diag: diag } };
      return;
    }
    // Atribuicao: quem rejeitou foi o FINANCEIRO (nao o gestor/aprovador que o RejeitarNota
    // registrou pelo principal falso). Corrige RejeitadoPor pro financeiro que agiu.
    diag.step = 'corrige_rejeitado_por';
    try {
      await client.api('/sites/' + siteId + '/lists/' + listId + '/items/' + itemId + '/fields')
        .patch({ RejeitadoPor: authz.email, RejeitadoEm: new Date().toISOString() });
    } catch (ePatch) { diag.rejeitadoPorPatchError = (ePatch && ePatch.message) || String(ePatch); }
    auditRegistrar(authz.user, 'alinhamento_rejeitado', { tipo: 'nf', id: itemId }, 'sucesso',
      { rejeitadoPor: authz.email, aprovadorOriginal: aprovador, motivo: motivoFinal }).catch(function () {});
    context.res = { status: 200, headers: { 'Content-Type': 'application/json' },
      body: { ok: true, acao: 'rejeitar', itemId: itemId, mensagem: 'Alinhamento rejeitado — NF rejeitada.' } };
  } catch (err) {
    context.log && context.log.error && context.log.error('ConfirmarAlinhamento:', err);
    context.res = { status: 500, headers: { 'Content-Type': 'application/json' },
      body: { error: (err && err.message) || String(err), diag: diag } };
  }
};
