/**
 * shared/authz.js — Autorizacao centralizada.
 *
 * Resolve o usuario autenticado (getUser) + suas roles mescladas de 3 fontes
 * (claims do Teams SSO, principal do Easy Auth, Graph transitiveMemberOf) e
 * expoe flags prontas (isAdmin / isFinanceiro / isGestor).
 *
 * Substitui a logica de merge de roles que estava duplicada em ListarNotas,
 * RejeitarNota, IntegrarOmie, AuditLog, etc. e o e-mail de admin hardcoded
 * espalhado pelo codigo (agora via env ADMIN_EMAILS, com fallback ao atual).
 */

const { getUser } = require('./auth');
const { getUserRoles } = require('./userRoles');

// E-mails com privilegio de admin (alem do grupo AAD 'administrador').
// Centraliza o que antes estava hardcoded em varios arquivos (A4).
const ADMIN_EMAILS = (process.env.ADMIN_EMAILS || 'rafael.machado@pronep.com.br')
  .split(',')
  .map(s => s.trim().toLowerCase())
  .filter(Boolean);

// Le as roles do x-ms-client-principal (Easy Auth). So uteis pra fonte 2 do merge.
function readClientPrincipalRoles(req) {
  try {
    const h = req && req.headers &&
      (req.headers['x-ms-client-principal'] || req.headers['X-MS-CLIENT-PRINCIPAL']);
    if (!h) return [];
    const principal = JSON.parse(Buffer.from(h, 'base64').toString('utf-8'));
    return (principal && principal.userRoles) || [];
  } catch (e) {
    return [];
  }
}

// Mescla as roles das 3 fontes e remove as roles default inuteis do SWA.
async function getMergedRoles(req, user) {
  const claimsRoles = (user && user.claims && user.claims.roles) || [];
  const principalRoles = (user && user.source === 'easy-auth' ? readClientPrincipalRoles(req) : []) || [];
  const usefulPrincipalRoles = principalRoles.filter(r => r !== 'authenticated' && r !== 'anonymous');
  const graphRoles = await getUserRoles(user);
  return Array.from(new Set([...claimsRoles, ...usefulPrincipalRoles, ...graphRoles]));
}

function isAdminEmail(email) {
  return ADMIN_EMAILS.includes((email || '').toLowerCase());
}

/**
 * Resolve usuario + roles + flags de uma vez.
 * @returns {Promise<null | { user, email, roles, isAdmin, isFinanceiro, isGestor }>}
 *          null = nao autenticado.
 */
async function resolveAuthz(req) {
  const user = await getUser(req);
  if (!user) return null;
  const roles = await getMergedRoles(req, user);
  const email = (user.email || '').toLowerCase();
  return {
    user,
    email,
    roles,
    isAdmin: roles.includes('administrador') || isAdminEmail(email),
    isFinanceiro: roles.includes('financeiro_nf'),
    isGestor: roles.some(r => typeof r === 'string' && r.startsWith('gestor'))
  };
}

/**
 * Guarda de admin pronta pra usar no topo de um endpoint.
 * Seta context.res (401/403) e retorna null se nao for admin.
 * Retorna o objeto authz se for admin.
 *
 *   const authz = await requireAdmin(context, req);
 *   if (!authz) return;
 */
async function requireAdmin(context, req) {
  const authz = await resolveAuthz(req);
  if (!authz) {
    context.res = { status: 401, headers: { 'Content-Type': 'application/json' }, body: { error: 'Nao autenticado' } };
    return null;
  }
  if (!authz.isAdmin) {
    context.res = { status: 403, headers: { 'Content-Type': 'application/json' }, body: { error: 'Acesso restrito a administradores' } };
    return null;
  }
  return authz;
}

module.exports = {
  ADMIN_EMAILS,
  readClientPrincipalRoles,
  getMergedRoles,
  isAdminEmail,
  resolveAuthz,
  requireAdmin
};
