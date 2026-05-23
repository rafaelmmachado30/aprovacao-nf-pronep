/**
 * DEBUG v2 — decodifica o JWT completo do header x-ms-token-aad-id-token
 * (esse header tem TODOS os claims, ao contrário do x-ms-client-principal resumido).
 *
 * REMOVER após resolver o problema!
 */

function decodeJwtPayload(jwt) {
  if (!jwt) return null;
  const parts = jwt.split('.');
  if (parts.length !== 3) return null;
  // base64url -> base64
  const payload = parts[1].replace(/-/g, '+').replace(/_/g, '/');
  // pad
  const padded = payload + '='.repeat((4 - payload.length % 4) % 4);
  try {
    return JSON.parse(Buffer.from(padded, 'base64').toString('utf-8'));
  } catch (e) {
    return { _decode_error: e.message };
  }
}

module.exports = async function (context, req) {
  try {
    const allHeaders = req.headers || {};
    // Lista headers úteis para debug (sem expor tokens completos)
    const interestingHeaders = {};
    for (const [k, v] of Object.entries(allHeaders)) {
      if (k.startsWith('x-ms-') || k === 'authorization') {
        interestingHeaders[k] = (typeof v === 'string' && v.length > 60)
          ? v.substring(0, 30) + '...(' + v.length + ' chars total)'
          : v;
      }
    }

    // Decodifica o ID Token
    const idToken = allHeaders['x-ms-token-aad-id-token'];
    const decodedId = decodeJwtPayload(idToken);

    // Decodifica o Access Token
    const accessToken = allHeaders['x-ms-token-aad-access-token'];
    const decodedAccess = decodeJwtPayload(accessToken);

    // Principal resumido
    const principalB64 = allHeaders['x-ms-client-principal'];
    let principal = null;
    if (principalB64) {
      try {
        principal = JSON.parse(Buffer.from(principalB64, 'base64').toString('utf-8'));
      } catch (e) {
        principal = { _decode_error: e.message };
      }
    }

    context.res = {
      status: 200,
      headers: { 'Content-Type': 'application/json' },
      body: {
        message: 'JWT debug',
        '_HEADERS_PRESENTES': Object.keys(allHeaders).filter(h => h.startsWith('x-ms-')),
        '_HEADER_x-ms-token-aad-id-token_existe': !!idToken,
        '_HEADER_x-ms-token-aad-access-token_existe': !!accessToken,
        'principal_resumido': principal,
        'id_token_payload': decodedId,
        'access_token_payload': decodedAccess,
        '_INTERESTING_HEADERS': interestingHeaders
      }
    };
  } catch (err) {
    context.res = {
      status: 500,
      body: { error: err.message, stack: err.stack }
    };
  }
};
