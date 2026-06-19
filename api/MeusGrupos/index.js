/**
 * Sistema de Aprovação de NF — MeusGrupos
 *
 * Lê x-ms-client-principal (Easy Auth), descobre o OID do usuário,
 * obtém token Application via Client Credentials e consulta o Graph
 * pra retornar os grupos do usuário + roles mapeadas.
 *
 * App Settings exigidas no SWA:
 *   - AAD_CLIENT_ID
 *   - AAD_CLIENT_SECRET
 *   - AAD_TENANT_ID
 *
 * Permissões Application (Microsoft Graph) no AppReg:
 *   - GroupMember.Read.All (com admin consent)
 */

require('isomorphic-fetch');
const { getUser } = require('../shared/auth');
const { lerMapaTelas, telasLiberadasPara } = require('../shared/acessoTelas');
const { ClientSecretCredential } = require('@azure/identity');
const { Client } = require('@microsoft/microsoft-graph-client');
const { TokenCredentialAuthenticationProvider } =
  require('@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials');

const GROUP_TO_ROLE = {
  '01d540d1-8596-42d0-9a20-de5c361c7c96': 'submitter',
  '480a1595-bdc3-492a-9ef2-317f148a237e': 'administrador',
  'a8425e5e-6497-4bf5-a8e3-e206b0294340': 'ti',  // PRONEP-NF-TI
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

function readClientPrincipal(req) {
  const header = req.headers && req.headers['x-ms-client-principal'];
  if (!header) return null;
  try {
    const decoded = Buffer.from(header, 'base64').toString('utf-8');
    return JSON.parse(decoded);
  } catch (e) {
    return null;
  }
}

function extractUserOid(principal) {
  if (!principal) return null;
  const claims = principal.claims || [];
  const oidClaim = claims.find(c =>
    c.typ === 'http://schemas.microsoft.com/identity/claims/objectidentifier' ||
    c.typ === 'oid'
  );
  if (oidClaim && oidClaim.val) return oidClaim.val;
  if (principal.userId) return principal.userId;
  return null;
}

// Easy Auth as vezes manda OID sem hifens (32 chars). Graph exige UUID com hifens.
function formatGuid(guid) {
  if (!guid) return guid;
  const d = String(guid).replace(/-/g, '').toLowerCase();
  if (d.length !== 32) return guid;
  return d.slice(0,8) + '-' + d.slice(8,12) + '-' + d.slice(12,16) + '-' + d.slice(16,20) + '-' + d.slice(20,32);
}

module.exports = async function (context, req) {
  // Diagnóstico: sempre retorna JSON, mesmo em caso de erro
  const diag = { step: 'start' };
  try {
    diag.step = 'env';
    const tenantId = process.env.AAD_TENANT_ID;
    const clientId = process.env.AAD_CLIENT_ID;
    const clientSecret = process.env.AAD_CLIENT_SECRET;

    if (!tenantId || !clientId || !clientSecret) {
      context.res = {
        status: 500,
        headers: { 'Content-Type': 'application/json' },
        body: {
          error: 'App Settings incompletas',
          missing: {
            AAD_TENANT_ID: !tenantId,
            AAD_CLIENT_ID: !clientId,
            AAD_CLIENT_SECRET: !clientSecret
          }
        }
      };
      return;
    }

    diag.step = 'principal';
    const user = await getUser(req);
    if (!user) {
      context.res = {
        status: 401,
        headers: { 'Content-Type': 'application/json' },
        body: {
          error: 'Nao autenticado (nem Easy Auth nem Bearer Teams).',
          authError: req._authError || null
        }
      };
      return;
    }
    diag.authSource = user.source;

    diag.step = 'identify_user';
    // Prioridade: oid do JWT (autoritativo do Entra ID) > email
    let userKey = null;
    let userKeyKind = null;
    if (user.oid && /^[0-9a-f-]{32,36}$/i.test(user.oid)) {
      userKey = formatGuid(user.oid);
      userKeyKind = user.source === 'teams-sso' ? 'oid_teams_token' : 'oid_claim';
    } else if (user.email && /@/.test(user.email)) {
      userKey = user.email;
      userKeyKind = 'email';
    } else {
      context.res = {
        status: 400,
        headers: { 'Content-Type': 'application/json' },
        body: {
          error: 'Nao foi possivel identificar o usuario (sem oid nem email).',
          user: user
        }
      };
      return;
    }
    diag.userKey = userKey;
    // Cria principal sintetico pra compatibilidade com codigo existente
    const principal = { userDetails: user.email, userId: user.oid, claims: (user.claims ? Object.entries(user.claims).map(function(kv){ return { typ: kv[0], val: String(kv[1]) }; }) : []) };
    diag.userKeyKind = userKeyKind;

    diag.step = 'credential';
    const credential = new ClientSecretCredential(tenantId, clientId, clientSecret);
    const authProvider = new TokenCredentialAuthenticationProvider(credential, {
      scopes: ['https://graph.microsoft.com/.default']
    });

    diag.step = 'graph_client';
    const client = Client.initWithMiddleware({ authProvider });

    diag.step = 'graph_call';
    const result = await client
      .api(`/users/${encodeURIComponent(userKey)}/transitiveMemberOf`)
      .select('id,displayName')
      .top(200)
      .get();

    const items = (result && result.value) || [];
    const groups = items
      .filter(it => it && it.id)
      .map(it => ({ id: String(it.id).toLowerCase(), displayName: it.displayName || '' }));

    const roleMap = {};
    for (const [k, v] of Object.entries(GROUP_TO_ROLE)) roleMap[k.toLowerCase()] = v;
    const roles = groups.map(g => roleMap[g.id]).filter(Boolean);

    const aggregated = new Set(roles);
    if (roles.some(r => r.startsWith('gestor_'))) aggregated.add('gestor');

    // Telas extras liberadas via Central de Controle de Acessos (ADITIVO ao canSee
    // do papel). Best-effort: se a config falhar, segue sem telas extras.
    diag.step = 'telas_extras';
    let telasExtras = [];
    try {
      const host = process.env.SHAREPOINT_SITE_HOSTNAME;
      const sitePath = process.env.SHAREPOINT_SITE_PATH;
      if (host && sitePath) {
        const siteResp = await client.api('/sites/' + host + ':' + sitePath).get();
        const mapaTelas = await lerMapaTelas(client, siteResp.id, null);
        telasExtras = telasLiberadasPara(principal.userDetails || '', Array.from(aggregated), mapaTelas);
      }
    } catch (e) { /* sem telas extras */ }

    context.res = {
      status: 200,
      headers: { 'Content-Type': 'application/json', 'Cache-Control': 'no-store' },
      body: {
        userId: userKey,
        userKeyKind: userKeyKind,
        userDetails: principal.userDetails || null,
        identityProvider: principal.identityProvider || null,
        groups,
        roles,
        rolesAggregated: Array.from(aggregated),
        telasExtras: telasExtras
      }
    };
  } catch (err) {
    context.log && context.log.error && context.log.error('MeusGrupos error:', err);
    context.res = {
      status: 500,
      headers: { 'Content-Type': 'application/json' },
      body: {
        error: (err && err.message) || String(err),
        code: err && err.code,
        statusCode: err && err.statusCode,
        body: err && err.body,
        diag
      }
    };
  }
};
