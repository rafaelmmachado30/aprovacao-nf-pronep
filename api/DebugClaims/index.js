/**
 * DEBUG v3 — decodifica TODOS os JWTs presentes nos headers
 * Foco em x-ms-auth-token e authorization, que devem conter os grupos.
 *
 * REMOVER após resolver o problema!
 */

function decodeJwtPayload(jwt) {
  if (!jwt) return null;
  // Remove "Bearer " prefix se tiver
  jwt = jwt.replace(/^Bearer\s+/i, '');
  const parts = jwt.split('.');
  if (parts.length !== 3) return { _error: 'JWT mal formado', _parts_count: parts.length };
  const payload = parts[1].replace(/-/g, '+').replace(/_/g, '/');
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

    // Decodifica todos os possíveis tokens
    const decoded = {};
    for (const headerName of [
      'authorization',
      'x-ms-auth-token',
      'x-ms-token-aad-id-token',
      'x-ms-token-aad-access-token',
      'x-ms-token-aad-refresh-token'
    ]) {
      if (allHeaders[headerName]) {
        decoded[headerName] = {
          length: allHeaders[headerName].length,
          payload: decodeJwtPayload(allHeaders[headerName])
        };
      } else {
        decoded[headerName] = '(não enviado pelo SWA)';
      }
    }

    // Lista todos os headers x-ms-* (sem valores grandes)
    const xmsHeaders = {};
    for (const [k, v] of Object.entries(allHeaders)) {
      if (k.startsWith('x-ms-')) {
        xmsHeaders[k] = (typeof v === 'string' && v.length > 100)
          ? `(${v.length} chars)`
          : v;
      }
    }

    context.res = {
      status: 200,
      headers: { 'Content-Type': 'application/json' },
      body: {
        message: 'JWT debug v3',
        '_HEADERS_X-MS': xmsHeaders,
        'TOKENS_DECODIFICADOS': decoded
      }
    };
  } catch (err) {
    context.res = {
      status: 500,
      body: { error: err.message, stack: err.stack }
    };
  }
};
