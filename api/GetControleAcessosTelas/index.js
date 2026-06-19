/**
 * /api/GetControleAcessosTelas (GET) — ADMIN ONLY
 *
 * Dados da aba "Telas do sistema" da Central de Controle de Acessos:
 *   {
 *     telas:   [{ id:'aprovadas', label:'Notas Aprovadas' }, ...]  // telas configuraveis
 *     grupos:  [{ role:'gestor_fiscal_contabil', label:'Fiscal-Contábil' }, ...] // grupos Entra
 *     acessos: { 'aprovadas': ['gestor_fiscal_contabil'], ... }    // mapa atual (tela -> tokens)
 *   }
 *
 * Espelha GetControleAcessos (contratos), mas o alvo e uma TELA em vez de uma pasta.
 */

require('isomorphic-fetch');
const { requireAdmin } = require('../shared/authz');
const { lerMapaTelas, TELAS } = require('../shared/acessoTelas');
const { gruposContrato } = require('../shared/userRoles');
const { ClientSecretCredential } = require('@azure/identity');
const { Client } = require('@microsoft/microsoft-graph-client');
const { TokenCredentialAuthenticationProvider } =
  require('@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials');

function getGraphClient() {
  const tenantId = process.env.AAD_TENANT_ID;
  const clientId = process.env.AAD_CLIENT_ID;
  const clientSecret = process.env.AAD_CLIENT_SECRET;
  if (!tenantId || !clientId || !clientSecret) throw new Error('AAD_* incompletas');
  const credential = new ClientSecretCredential(tenantId, clientId, clientSecret);
  const authProvider = new TokenCredentialAuthenticationProvider(credential, {
    scopes: ['https://graph.microsoft.com/.default']
  });
  return Client.initWithMiddleware({ authProvider });
}

async function resolveSite(client) {
  const host = process.env.SHAREPOINT_SITE_HOSTNAME;
  const path = process.env.SHAREPOINT_SITE_PATH;
  if (!host || !path) throw new Error('SHAREPOINT_* incompletas');
  const siteResp = await client.api('/sites/' + host + ':' + path).get();
  return siteResp.id;
}

module.exports = async function (context, req) {
  try {
    const authz = await requireAdmin(context, req);
    if (!authz) return;

    const client = getGraphClient();
    const siteId = await resolveSite(client);
    const acessos = await lerMapaTelas(client, siteId, null);
    const grupos = gruposContrato(); // lista canonica [{role,label}] dos grupos do Entra

    context.res = {
      status: 200,
      headers: { 'Content-Type': 'application/json', 'Cache-Control': 'no-store' },
      body: { ok: true, telas: TELAS, grupos: grupos, acessos: acessos }
    };
  } catch (err) {
    context.log && context.log.error && context.log.error('GetControleAcessosTelas:', err);
    context.res = { status: 500, headers: { 'Content-Type': 'application/json' },
      body: { error: (err && err.message) || String(err) } };
  }
};
