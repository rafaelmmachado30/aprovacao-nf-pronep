// Function minima de diagnostico — sem require, sem deps externas
module.exports = async function (context, req) {
  context.res = {
    status: 200,
    headers: { 'Content-Type': 'application/json' },
    body: {
      ok: true,
      time: new Date().toISOString(),
      node: process.version,
      env: {
        AAD_TENANT_ID:           !!process.env.AAD_TENANT_ID,
        AAD_CLIENT_ID:           !!process.env.AAD_CLIENT_ID,
        AAD_CLIENT_SECRET:       !!process.env.AAD_CLIENT_SECRET,
        SHAREPOINT_SITE_HOSTNAME: process.env.SHAREPOINT_SITE_HOSTNAME || null,
        SHAREPOINT_SITE_PATH:     process.env.SHAREPOINT_SITE_PATH || null
      },
      message: 'Se voce esta vendo isso, o runtime Functions esta OK. Problema deve ser em deps especificas.'
    }
  };
};
