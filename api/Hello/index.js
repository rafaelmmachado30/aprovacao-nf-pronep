// Hello v4 — testa fetch + identity + microsoft-graph-client
let fetchErr = null, identityErr = null, graphErr = null, authProvErr = null;

try { require('isomorphic-fetch'); } catch (e) { fetchErr = e.message; }

let ClientSecretCredentialType = null;
try {
  const id = require('@azure/identity');
  ClientSecretCredentialType = typeof id.ClientSecretCredential;
} catch (e) { identityErr = e.message; }

let GraphClientType = null, AuthProvType = null;
try {
  const g = require('@microsoft/microsoft-graph-client');
  GraphClientType = typeof g.Client;
} catch (e) { graphErr = e.message; }

try {
  const a = require('@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials');
  AuthProvType = typeof a.TokenCredentialAuthenticationProvider;
} catch (e) { authProvErr = e.message; }

module.exports = async function (context, req) {
  context.res = {
    status: 200,
    headers: { 'Content-Type': 'application/json' },
    body: {
      ok: true,
      time: new Date().toISOString(),
      node: process.version,
      step: 'checkpoint-4: testing microsoft-graph-client',
      requires: {
        isomorphicFetch: { error: fetchErr, fetchType: typeof fetch },
        azureIdentity:   { error: identityErr, ClientSecretCredential: ClientSecretCredentialType },
        graphClient:     { error: graphErr, ClientType: GraphClientType },
        authProvider:    { error: authProvErr, ProviderType: AuthProvType }
      }
    }
  };
};
