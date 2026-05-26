/**
 * Helper de autenticacao com 2 caminhos:
 *  1. Easy Auth (cookie SWA) via header x-ms-client-principal
 *  2. Teams SSO via Authorization: Bearer <jwt> (validado contra JWKS Microsoft)
 *
 * getUser(req) retorna { email, name, oid, source } ou null
 */

require('isomorphic-fetch');
const jwt = require('jsonwebtoken');
const jwksClient = require('jwks-rsa');

const TENANT_ID = process.env.AAD_TENANT_ID;
const CLIENT_ID = process.env.AAD_CLIENT_ID;
// Aceita tokens cujo audience seja o App ID URI OU o clientId puro
const APP_ID_URI = process.env.APP_ID_URI
  || 'api://purple-forest-09588fe10.7.azurestaticapps.net/' + CLIENT_ID;
const ACCEPTED_AUDIENCES = [APP_ID_URI, CLIENT_ID];

let _jwks = null;
function getJwksClient() {
  if (_jwks) return _jwks;
  _jwks = jwksClient({
    jwksUri: 'https://login.microsoftonline.com/' + TENANT_ID + '/discovery/v2.0/keys',
    cache: true,
    cacheMaxAge: 60 * 60 * 1000, // 1h
    rateLimit: true
  });
  return _jwks;
}

function getKey(header, cb) {
  getJwksClient().getSigningKey(header.kid, function (err, key) {
    if (err) return cb(err);
    cb(null, key.getPublicKey());
  });
}

function validateTeamsToken(token) {
  return new Promise(function (resolve, reject) {
    const expectedIssuers = [
      'https://login.microsoftonline.com/' + TENANT_ID + '/v2.0',
      'https://sts.windows.net/' + TENANT_ID + '/'
    ];
    jwt.verify(token, getKey, {
      audience: ACCEPTED_AUDIENCES,
      issuer: expectedIssuers,
      algorithms: ['RS256']
    }, function (err, decoded) {
      if (err) return reject(err);
      resolve(decoded);
    });
  });
}

// Formata GUID sem hifens pro formato padrao 8-4-4-4-12
function formatGuid(s) {
  if (!s) return s;
  const v = String(s).replace(/-/g, '').toLowerCase();
  if (v.length !== 32) return String(s);
  return v.slice(0,8) + '-' + v.slice(8,12) + '-' + v.slice(12,16) + '-' + v.slice(16,20) + '-' + v.slice(20);
}

function userFromEasyAuth(req) {
  const header = req.headers && (req.headers['x-ms-client-principal'] || req.headers['X-MS-CLIENT-PRINCIPAL']);
  if (!header) return null;
  try {
    const decoded = Buffer.from(header, 'base64').toString('utf-8');
    const principal = JSON.parse(decoded);
    if (!principal || !principal.userId) return null;
    // IMPORTANTE: principal.userId eh o ID interno do SWA, NAO o oid do Entra ID.
    // O oid real vem dos claims (claim 'oid' ou objectidentifier).
    const claims = principal.claims || [];
    let realOid = null;
    for (const c of claims) {
      if (!c || !c.val) continue;
      if (c.typ === 'http://schemas.microsoft.com/identity/claims/objectidentifier' || c.typ === 'oid') {
        if (/^[0-9a-f-]{32,36}$/i.test(c.val)) {
          realOid = formatGuid(c.val);
          break;
        }
      }
    }
    return {
      email: (principal.userDetails || '').toLowerCase(),
      name: principal.userDetails || '',
      oid: realOid || principal.userId, // fallback so se nao achar oid claim
      source: 'easy-auth'
    };
  } catch (e) {
    return null;
  }
}

async function userFromTeamsToken(req) {
  // IMPORTANTE: SWA Easy Auth SUBSTITUI o header Authorization por token interno do Azure
  // (aud=azurewebsites.net/azurefunctions). Por isso usamos header custom X-Teams-Token
  // pra o token MSAL do Teams passar intacto.
  let token = null;
  if (req.headers) {
    token = req.headers['x-teams-token'] || req.headers['X-Teams-Token'] || null;
  }
  // Fallback: tentar Authorization Bearer (caso o SWA nao toque no header)
  if (!token && req.headers) {
    const auth = req.headers.authorization || req.headers.Authorization;
    if (auth && auth.startsWith('Bearer ')) token = auth.substring(7).trim();
  }
  if (!token) return null;
  // Decodifica payload sem validar (pra log/debug)
  let unverifiedPayload = null;
  try {
    const parts = token.split('.');
    if (parts.length === 3) {
      unverifiedPayload = JSON.parse(Buffer.from(parts[1], 'base64').toString('utf-8'));
    }
  } catch (e) {}
  try {
    const claims = await validateTeamsToken(token);
    return {
      email: (claims.preferred_username || claims.upn || claims.email || '').toLowerCase(),
      name: claims.name || claims.preferred_username || '',
      oid: claims.oid || claims.sub,
      source: 'teams-sso',
      claims: claims
    };
  } catch (e) {
    // Anexa info de debug no erro pra Functions logarem
    const debug = unverifiedPayload ? {
      aud: unverifiedPayload.aud,
      iss: unverifiedPayload.iss,
      scp: unverifiedPayload.scp,
      tid: unverifiedPayload.tid,
      appid: unverifiedPayload.appid,
      preferred_username: unverifiedPayload.preferred_username,
      ver: unverifiedPayload.ver
    } : null;
    const err = new Error('Teams token rejeitado: ' + (e.message || e));
    err.tokenDebug = debug;
    err.expectedAudiences = ACCEPTED_AUDIENCES;
    err.expectedIssuers = [
      'https://login.microsoftonline.com/' + TENANT_ID + '/v2.0',
      'https://sts.windows.net/' + TENANT_ID + '/'
    ];
    err.original = e.message;
    throw err;
  }
}

async function getUser(req) {
  // 1. Easy Auth primeiro (mais comum, no browser)
  const easy = userFromEasyAuth(req);
  if (easy) return easy;
  // 2. Teams SSO Bearer (dentro do iframe Teams)
  try {
    const teams = await userFromTeamsToken(req);
    if (teams) return teams;
  } catch (e) {
    // Salva no req pra Function poder retornar info de debug ao front
    if (req) req._authError = {
      message: e.message,
      tokenDebug: e.tokenDebug,
      expectedAudiences: e.expectedAudiences,
      expectedIssuers: e.expectedIssuers,
      original: e.original
    };
    return null;
  }
  return null;
}

module.exports = { getUser, userFromEasyAuth, userFromTeamsToken, validateTeamsToken };
