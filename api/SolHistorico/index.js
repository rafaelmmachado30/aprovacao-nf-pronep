/**
 * /api/SolHistorico
 *
 * GET    — retorna ultimas N msgs do user logado (default 20, max 100)
 *          Body: { history: [{role, content, timestamp, tokensUsed?}], total }
 *
 * DELETE — apaga TODO o historico do user (botao "Limpar Conversa")
 *          Body: { ok: true, removidos: N }
 */

require('isomorphic-fetch');
const { getUser } = require('../shared/auth');
const { lerUltimas, limparHistorico } = require('../shared/solHistorico');

module.exports = async function (context, req) {
  try {
    const user = await getUser(req);
    if (!user) {
      context.res = { status: 401, body: { error: 'Nao autenticado' } };
      return;
    }

    if (req.method === 'GET') {
      const limit = parseInt(req.query && req.query.limit || '20');
      const history = await lerUltimas(user, limit);
      context.res = {
        status: 200,
        headers: { 'Content-Type': 'application/json' },
        body: { history: history, total: history.length }
      };
      return;
    }

    if (req.method === 'DELETE') {
      const result = await limparHistorico(user);
      context.res = {
        status: 200,
        headers: { 'Content-Type': 'application/json' },
        body: { ok: true, removidos: result.removidos }
      };
      return;
    }

    context.res = { status: 405, body: { error: 'Method not allowed' } };
  } catch (err) {
    context.log && context.log.error && context.log.error('SolHistorico error:', err);
    context.res = {
      status: 500,
      headers: { 'Content-Type': 'application/json' },
      body: {
        error: 'Erro interno',
        detail: (err && err.message) || String(err),
        stack: (err && err.stack || '').split('\n').slice(0, 6)
      }
    };
  }
};
