// Hello v3 — testa isomorphic-fetch + @azure/identity
let fetchErr = null, identityErr = null;
try { require('isomorphic-fetch'); } catch (e) { fetchErr = e.message; }
let ClientSecretCredential = null;
try {
  const id = require('@azure/identity');
  ClientSecretCredential = typeof id.ClientSecretCredential;
} catch (e) {
  identityErr = e.message;
}

module.exports = async function (context, req) {
  context.res = {
    status: 200,
    headers: { 'Content-Type': 'application/json' },
    body: {
      ok: true,
      time: new Date().toISOString(),
      node: process.version,
      step: 'checkpoint-3: testing @azure/identity',
      fetchRequireError: fetchErr,
      fetchAvailable: typeof fetch,
      identityRequireError: identityErr,
      ClientSecretCredentialType: ClientSecretCredential,
      env: {
        AAD_TENANT_ID: !!process.env.AAD_TENANT_ID,
        AAD_CLIENT_ID: !!process.env.AAD_CLIENT_ID
      }
    }
  };
};
