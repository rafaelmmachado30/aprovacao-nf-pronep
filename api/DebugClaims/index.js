/**
 * DEBUG — retorna o conteúdo completo do x-ms-client-principal header
 * pra diagnosticar problema de mapeamento de roles.
 *
 * REMOVER após resolver o problema!
 */

module.exports = async function (context, req) {
  const headerB64 = req.headers['x-ms-client-principal'];
  if (!headerB64) {
    context.res = {
      status: 401,
      headers: { 'Content-Type': 'application/json' },
      body: { error: 'No x-ms-client-principal header. Você está autenticado?' }
    };
    return;
  }
  try {
    const principal = JSON.parse(Buffer.from(headerB64, 'base64').toString('utf-8'));
    context.res = {
      status: 200,
      headers: { 'Content-Type': 'application/json' },
      body: {
        message: 'Principal decoded successfully',
        identityProvider: principal.identityProvider,
        userId: principal.userId,
        userDetails: principal.userDetails,
        userRoles: principal.userRoles,
        claimsCount: (principal.claims || []).length,
        claimsTypes: [...new Set((principal.claims || []).map(c => c.typ))],
        claims: principal.claims || []
      }
    };
  } catch (err) {
    context.res = {
      status: 500,
      body: { error: err.message }
    };
  }
};
