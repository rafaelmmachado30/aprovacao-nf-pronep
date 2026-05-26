/**
 * /api/ListarGestoresFinanceiro
 *
 * Retorna membros do grupo Entra ID PRONEP-NF-Gestor_Financeira (autorizados a aprovar
 * negociacao de vencimento fora do prazo D+5).
 *
 * App Setting opcional:
 *   GESTOR_FINANCEIRO_GROUP_ID — OID do grupo. Default: 6b77405b-ba89-47ee-af21-58ec19bb3ff7
 */

require('isomorphic-fetch');
const { getUser } = require('../shared/auth');
const { ClientSecretCredential } = require('@azure/identity');
const { Client } = require('@microsoft/microsoft-graph-client');
const { TokenCredentialAuthenticationProvider } =
  require('@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials');

const DEFAULT_GROUP_ID = 'c2a73d16-4659-4b3c-93a1-0c0fbfaaaa96'; // PRONEP-NF-Financeiro-Gestao (autorizados a aprovar vencimento fora do prazo)

async function getGraphClient() {
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

module.exports = async function (context, req) {
  try {
    const user = await getUser(req);
    if (!user) {
      context.res = { status: 401, headers: { 'Content-Type': 'application/json' }, body: { error: 'Nao autenticado' } };
      return;
    }
    const groupId = process.env.GESTOR_FINANCEIRO_GROUP_ID || DEFAULT_GROUP_ID;
    const client = await getGraphClient();
    const resp = await client
      .api(`/groups/${groupId}/members`)
      .select('id,displayName,mail,userPrincipalName,jobTitle')
      .top(50)
      .get();

    const membros = (resp.value || []).map(u => ({
      id: u.id,
      nome: u.displayName || '',
      email: (u.mail || u.userPrincipalName || '').toLowerCase(),
      cargo: u.jobTitle || ''
    })).filter(m => m.email);

    context.res = {
      status: 200,
      headers: { 'Content-Type': 'application/json', 'Cache-Control': 'no-store' },
      body: { groupId, total: membros.length, membros }
    };
  } catch (err) {
    context.log && context.log.error && context.log.error('ListarGestoresFinanceiro:', err);
    context.res = {
      status: 500,
      headers: { 'Content-Type': 'application/json' },
      body: { error: (err && err.message) || String(err), statusCode: err && err.statusCode, body: err && err.body }
    };
  }
};
