/**
 * shared/userRoles.js — Resolve roles do usuario logado consultando grupos no Entra ID.
 *
 * O SWA Easy Auth NAO popula claims.roles a partir dos grupos AAD por default.
 * Pra saber se um user eh gestor/admin/financeiro/etc, precisamos consultar
 * /users/{oid|email}/transitiveMemberOf via Graph e mapear pelos OIDs conhecidos.
 *
 * Mantem o GROUP_TO_ROLE SINCRONIZADO com /api/MeusGrupos/index.js — quando
 * adicionar um grupo novo (ex: PRONEP-NF-TI), atualizar nos DOIS lugares.
 */

require('isomorphic-fetch');
const { ClientSecretCredential } = require('@azure/identity');
const { Client } = require('@microsoft/microsoft-graph-client');
const { TokenCredentialAuthenticationProvider } =
  require('@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials');

const GROUP_TO_ROLE = {
  '01d540d1-8596-42d0-9a20-de5c361c7c96': 'submitter',
  '480a1595-bdc3-492a-9ef2-317f148a237e': 'administrador',
  'a8425e5e-6497-4bf5-a8e3-e206b0294340': 'ti',
  'c2a73d16-4659-4b3c-93a1-0c0fbfaaaa96': 'financeiro_nf',
  '2d9f5bcf-2ae0-494e-957b-a1c69016664d': 'gestor_suprimentos',
  '6b77405b-ba89-47ee-af21-58ec19bb3ff7': 'gestor_financeira',
  'a7826b5c-7c29-4a24-836b-a7432aa941ec': 'gestor_tecnologia',
  'a6711877-8746-4ca5-a955-c15980c7e90d': 'gestor_qualidade',
  '13a544d8-3dde-4820-9695-c492e58a2782': 'gestor_rh_dp',
  'b9272d98-3e26-4e2d-aae6-ff9057f57e5c': 'gestor_fiscal_contabil',
  '5aa9fc6b-900d-40eb-861d-8bbf72499da1': 'gestor_juridica',
  '4f28f31b-b704-4615-961b-a9ca0898cea8': 'gestor_administrativa',
  'fc3a375b-329c-4d9c-81be-06180f0598af': 'gestor_tecnica_sp',
  '334eb19b-c138-4551-8e45-a36ca4e32e48': 'gestor_tecnica_rjes'
};

function formatGuid(guid) {
  if (!guid) return guid;
  const d = String(guid).replace(/-/g, '').toLowerCase();
  if (d.length !== 32) return guid;
  return d.slice(0,8) + '-' + d.slice(8,12) + '-' + d.slice(12,16) + '-' + d.slice(16,20) + '-' + d.slice(20,32);
}

// Cache simples por user (TTL 5min) pra evitar bater no Graph a cada chamada
const _cache = new Map();
const _TTL_MS = 5 * 60 * 1000;

/**
 * Retorna as roles do user (array de strings tipo 'administrador', 'gestor_suprimentos', etc).
 * Lista vazia se user nao for membro de nenhum grupo conhecido.
 *
 * @param {Object} user - { oid?, email? } do auth
 * @returns {Promise<string[]>}
 */
async function getUserRoles(user) {
  if (!user) return [];

  // Define a chave de busca
  let userKey = null;
  if (user.oid && /^[0-9a-f-]{32,36}$/i.test(user.oid)) {
    userKey = formatGuid(user.oid);
  } else if (user.email && /@/.test(user.email)) {
    userKey = user.email;
  } else {
    return [];
  }

  // Cache
  const cached = _cache.get(userKey);
  if (cached && (Date.now() - cached.ts) < _TTL_MS) return cached.roles;

  try {
    const tenantId = process.env.AAD_TENANT_ID;
    const clientId = process.env.AAD_CLIENT_ID;
    const clientSecret = process.env.AAD_CLIENT_SECRET;
    if (!tenantId || !clientId || !clientSecret) return [];

    const credential = new ClientSecretCredential(tenantId, clientId, clientSecret);
    const authProvider = new TokenCredentialAuthenticationProvider(credential, {
      scopes: ['https://graph.microsoft.com/.default']
    });
    const client = Client.initWithMiddleware({ authProvider });

    const result = await client
      .api(`/users/${encodeURIComponent(userKey)}/transitiveMemberOf`)
      .select('id,displayName')
      .top(200)
      .get();

    const groupIds = ((result && result.value) || [])
      .filter(g => g && g.id)
      .map(g => String(g.id).toLowerCase());

    const roles = [];
    for (const gid of groupIds) {
      const role = GROUP_TO_ROLE[gid];
      if (role && !roles.includes(role)) roles.push(role);
    }

    _cache.set(userKey, { ts: Date.now(), roles });
    return roles;
  } catch (e) {
    console.error('[getUserRoles] erro:', e && e.message);
    return [];
  }
}

module.exports = { getUserRoles, GROUP_TO_ROLE };
