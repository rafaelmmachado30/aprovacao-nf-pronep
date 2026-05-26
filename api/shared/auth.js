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

function userFromEasyAuth(req) {
  const header = req.headers && (req.headers['x-ms-client-principal'] || req.headers['X-MS-CLIENT-PRINCIPAL']);
  if (!header) return null;
  try {
    const decoded = Buffer.from(header, 'base64').toString('utf-8');
    const principal = JSON.parse(decoded);
    if (!principal || !principal.userId) return null;
    return {
      email: (principal.userDetails || '').toLowerCase(),
      name: principal.userDetails || '',
      oid: principal.userId,
      source: 'easy-auth'
    };
  } catch (e) {
    return null;
  }
}

async function userFromTeamsToken(req) {
  const auth = req.headers && (req.headers.authorization || req.headers.Authorization);
  if (!auth || !auth.startsWith('Bearer ')) return null;
  const token = auth.substring(7).trim();
  if (!token) return null;
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
    return null;
  }
}

async function getUser(req) {
  // 1. Easy Auth primeiro (mais comum, no browser)
  const easy = userFromEasyAuth(req);
  if (easy) return easy;
  // 2. Teams SSO Bearer (dentro do iframe Teams)
  const teams = await userFromTeamsToken(req);
  if (teams) return teams;
  return null;
}

module.exports = { getUser, userFromEasyAuth, userFromTeamsToken, validateTeamsToken };
