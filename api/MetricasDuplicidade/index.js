/**
 * /api/MetricasDuplicidade (GET) — admin | gestor | financeiro (quem ve o Dashboard).
 *
 * Conta quantas NFs foram BARRADAS por DUPLICIDADE no lancamento, num mes. Esses
 * bloqueios nao viram registro em PRONEP-NF-NotasFiscais (a NF e barrada, nao salva),
 * entao o unico lugar com o dado e o AuditLog (evento lancamento/bloqueado/duplicidade).
 *
 * Query: ?mes=AAAA-MM (default: mes corrente BRT).
 * Resposta: { ok, mes, duplicidadesEvitadas, amostraTruncada }
 */

require('isomorphic-fetch');
const { resolveAuthz } = require('../shared/authz');
const { listar } = require('../shared/auditLog');

module.exports = async function (context, req) {
  try {
    const authz = await resolveAuthz(req);
    if (!authz) { context.res = { status: 401, body: { error: 'Nao autenticado' } }; return; }
    if (!(authz.isAdmin || authz.isGestor || authz.isFinanceiro)) {
      context.res = { status: 403, body: { error: 'Acesso restrito a gestores, financeiro e admin' } };
      return;
    }

    // Mes-alvo (BRT). Default: mes corrente.
    const agoraBRT = new Date(Date.now() - 3 * 60 * 60 * 1000);
    let ano = agoraBRT.getUTCFullYear();
    let mes0 = agoraBRT.getUTCMonth();
    const qMes = (req.query && req.query.mes) || '';
    const m = /^(\d{4})-(\d{2})$/.exec(qMes);
    if (m) { ano = Number(m[1]); mes0 = Number(m[2]) - 1; }
    const mesKey = ano + '-' + String(mes0 + 1).padStart(2, '0');
    const dataDe = new Date(Date.UTC(ano, mes0, 1, 0, 0, 0)).toISOString();
    const dataAte = new Date(Date.UTC(ano, mes0 + 1, 0, 23, 59, 59)).toISOString();

    // So eventos de lancamento BLOQUEADO no periodo (conjunto pequeno).
    const limit = 200;
    const r = await listar({ acao: 'lancamento', resultado: 'bloqueado', dataDe: dataDe, dataAte: dataAte, limit: limit });
    const eventos = (r && r.events) || [];

    // Conta os que foram barrados por duplicidade (o motivo fica em detalhes.motivo).
    let count = 0;
    for (const ev of eventos) {
      const det = ev.detalhes;
      let motivo = '';
      if (det && typeof det === 'object') motivo = String(det.motivo || '');
      else if (typeof det === 'string') motivo = det;
      if (/duplic/i.test(motivo)) count++;
    }

    context.res = {
      status: 200,
      headers: { 'Content-Type': 'application/json', 'Cache-Control': 'no-store' },
      body: {
        ok: true,
        mes: mesKey,
        duplicidadesEvitadas: count,
        // Sinaliza se a amostra bateu no teto (raro): nesse caso pode haver subcontagem.
        amostraTruncada: eventos.length >= limit
      }
    };
  } catch (err) {
    context.log && context.log.error && context.log.error('MetricasDuplicidade:', err);
    context.res = { status: 500, headers: { 'Content-Type': 'application/json' },
      body: { error: (err && err.message) || String(err) } };
  }
};
