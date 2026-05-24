/**
 * /api/EnviarNotificacao - delega pra shared/email
 *
 * Body: { evento, destinatarios: [emails], cc?, dados? }
 *
 * Retorna o resultado de cada canal (email + teams)
 */

const { notificar } = require('../shared/email');

module.exports = async function (context, req) {
  try {
    const body = req.body || {};
    const { evento, destinatarios, cc, dados } = body;
    if (!evento) {
      context.res = { status: 400, body: { error: 'evento obrigatorio' } };
      return;
    }
    if (!destinatarios || !destinatarios.length) {
      context.res = { status: 400, body: { error: 'destinatarios obrigatorio (array)' } };
      return;
    }
    const result = await notificar(evento, destinatarios, dados || {}, cc || []);
    context.res = {
      status: 200,
      headers: { 'Content-Type': 'application/json' },
      body: { ok: true, evento, result }
    };
  } catch (err) {
    context.log && context.log.error && context.log.error('EnviarNotificacao:', err);
    context.res = {
      status: 500,
      headers: { 'Content-Type': 'application/json' },
      body: { error: (err && err.message) || String(err), statusCode: err && err.statusCode, body: err && err.body }
    };
  }
};
