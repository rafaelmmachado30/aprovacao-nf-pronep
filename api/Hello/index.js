// Hello v2 — testa se isomorphic-fetch consegue ser carregado
let fetchRequireError = null;
try {
  require('isomorphic-fetch');
} catch (e) {
  fetchRequireError = e.message;
}

module.exports = async function (context, req) {
  context.res = {
    status: 200,
    headers: { 'Content-Type': 'application/json' },
    body: {
      ok: true,
      time: new Date().toISOString(),
      node: process.version,
      step: 'checkpoint-2: testing isomorphic-fetch',
      fetchRequireError: fetchRequireError,
      fetchAvailable: typeof fetch,
      env: {
        AAD_TENANT_ID:           !!process.env.AAD_TENANT_ID,
        SHAREPOINT_SITE_HOSTNAME: process.env.SHAREPOINT_SITE_HOSTNAME || null
      }
    }
  };
};
