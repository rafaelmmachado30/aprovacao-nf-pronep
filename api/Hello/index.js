// Hello — function viva pra teste rapido
module.exports = async function (context, req) {
  context.res = {
    status: 200,
    headers: { 'Content-Type': 'application/json' },
    body: {
      ok: true,
      time: new Date().toISOString(),
      node: process.version,
      step: 'checkpoint-5: Hello + MeusGrupos',
      env: {
        AAD_TENANT_ID: !!process.env.AAD_TENANT_ID,
        SHAREPOINT_SITE_HOSTNAME: process.env.SHAREPOINT_SITE_HOSTNAME || null
      }
    }
  };
};
