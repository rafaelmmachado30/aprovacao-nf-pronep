/**
 * /api/PushPublicKey
 *
 * Retorna a VAPID public key (necessaria no frontend pra criar a subscription).
 * Endpoint publico — VAPID public key nao eh secreta.
 *
 * Retorna:
 *   200 { publicKey: "BJP9M..." }
 *   500 { error: "VAPID_PUBLIC_KEY nao configurado" }
 */

module.exports = async function (context, req) {
  const pub = process.env.VAPID_PUBLIC_KEY;
  if (!pub) {
    context.res = { status: 500, headers: { 'Content-Type': 'application/json' },
      body: { error: 'VAPID_PUBLIC_KEY nao configurado no servidor' } };
    return;
  }
  context.res = {
    status: 200,
    headers: { 'Content-Type': 'application/json', 'Cache-Control': 'public, max-age=3600' },
    body: { publicKey: pub }
  };
};
