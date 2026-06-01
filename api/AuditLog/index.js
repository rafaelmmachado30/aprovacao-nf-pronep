/**
 * /api/AuditLog (GET) — Trilha de Auditoria. So admin pode acessar.
 *
 * Query params (todos opcionais):
 *   acao=<lancamento|aprovacao|rejeicao|processado|fornecedor_criar|...>
 *   userEmail=<email>
 *   dataDe=<ISO>
 *   dataAte=<ISO>
 *   limit=<N>  (default 50, max 200)
 *
 * Resposta:
 *   200 { events: [...], total: N }
 *   401 nao autenticado
 *   403 sem permissao (so admin)
 */

require('isomorphic-fetch');
const { getUser } = require('../shared/auth');
const { listar } = require('../shared/auditLog');

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

// Mapping grupo Entra ID -> role (mesma fonte de verdade do MeusGrupos)
const GROUP_ADMIN = '480a1595-bdc3-492a-9ef2-317f148a237e';

function isAdminFromHeader(req) {
  const principal = readClientPrincipal(req);
  if (!principal) return false;
  const roles = principal.userRoles || [];
  return roles.includes('administrador') || roles.includes('admin');
}

module.exports = async function (context, req) {
  try {
    const user = await getUser(req);
    if (!user) {
      context.res = { status: 401, body: { error: 'Nao autenticado' } };
      return;
    }
    const userEmail = (user.email || '').toLowerCase();
    const userRoles = (user.claims && user.claims.roles) || readClientPrincipalRoles(req);

    // CHECK ADMIN: aceita se claims.roles inclui admin OU email eh do admin master.
    // (Sistema atual mapeia grupos via MeusGrupos, mas pra simplicidade, fallback no email.)
    const isAdmin = userRoles.includes('administrador')
      || isAdminFromHeader(req)
      || userEmail === 'rafael.machado@pronep.com.br';

    if (!isAdmin) {
      context.res = {
        status: 403,
        headers: { 'Content-Type': 'application/json' },
        body: { error: 'Acesso negado. Trilha de auditoria eh restrita ao perfil Admin.' }
      };
      return;
    }

    const q = req.query || {};
    const filtros = {
      acao: q.acao,
      userOid: q.userOid,
      userEmail: q.userEmail,
      dataDe: q.dataDe,
      dataAte: q.dataAte,
      limit: q.limit
    };

    const result = await listar(filtros);
    context.res = {
      status: 200,
      headers: { 'Content-Type': 'application/json' },
      body: { events: result.events, total: result.total }
    };
  } catch (err) {
    context.log && context.log.error && context.log.error('AuditLog error:', err);
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
