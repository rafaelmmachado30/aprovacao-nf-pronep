/**
 * Sistema de Aprovação de NF — MeusGrupos
 *
 * Substitui o `rolesSource` (que não estava sendo invocado pelo SWA Free).
 *
 * Fluxo:
 *   1) Lê o header `x-ms-client-principal` injetado pelo Easy Auth (SWA),
 *      decodifica o base64 e descobre o `oid` (Object ID) do usuário logado.
 *   2) Faz Client Credentials flow contra o Entra ID (mesmo AppReg, com
 *      Application permission Group.Read.All / Directory.Read.All) e obtém
 *      um token para o Microsoft Graph.
 *   3) Chama GET https://graph.microsoft.com/v1.0/users/{oid}/transitiveMemberOf
 *      e devolve a lista de grupos do usuário + as roles mapeadas.
 *
 * App Settings exigidas no SWA (Configuration):
 *   - AAD_CLIENT_ID         (já existe, reusada do Easy Auth)
 *   - AAD_CLIENT_SECRET     (já existe, reusada do Easy Auth)
 *   - AAD_TENANT_ID         (novo — colocar 4b30645b-0888-45c0-9481-712bde435ffd)
 *
 * Permissões Application no AppReg (Microsoft Graph):
 *   - GroupMember.Read.All  (com admin consent)
 *      ou alternativamente Directory.Read.All
 */

const GROUP_TO_ROLE = {
  '01d540d1-8596-42d0-9a20-de5c361c7c96': 'submitter',
  '480a1595-bdc3-492a-9ef2-317f148a237e': 'administrador',
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
  const header = req.headers['x-ms-client-principal'];
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
  // userId do SWA já costuma ser o oid do Entra ID, mas vamos validar
  // procurando explicitamente um claim "oid" se existir.
  const claims = principal.claims || [];
  const oidClaim = claims.find(c =>
    c.typ === 'http://schemas.microsoft.com/identity/claims/objectidentifier' ||
    c.typ === 'oid'
  );
  if (oidClaim && oidClaim.val) return oidClaim.val;
  if (principal.userId) return principal.userId;
  return null;
}

async function getGraphAppToken(tenantId, clientId, clientSecret) {
  const url = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;
  const params = new URLSearchParams();
  params.append('client_id', clientId);
  params.append('client_secret', clientSecret);
  params.append('scope', 'https://graph.microsoft.com/.default');
  params.append('grant_type', 'client_credentials');

  const resp = await fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: params.toString()
  });

  if (!resp.ok) {
    const text = await resp.text();
    throw new Error(`Token endpoint retornou ${resp.status}: ${text}`);
  }

  const data = await resp.json();
  return data.access_token;
}

async function fetchUserGroups(graphToken, userOid) {
  // transitiveMemberOf devolve TODOS os grupos (incluindo grupos aninhados)
  const url = `https://graph.microsoft.com/v1.0/users/${userOid}/transitiveMemberOf?$select=id,displayName&$top=200`;

  const resp = await fetch(url, {
    headers: { 'Authorization': `Bearer ${graphToken}` }
  });

  if (!resp.ok) {
    const text = await resp.text();
    throw new Error(`Graph ${url} retornou ${resp.status}: ${text}`);
  }

  const data = await resp.json();
  const items = data.value || [];
  // Filtra só grupos (microsoft.graph.group), ignora roles de diretório
  const groups = items
    .filter(it => it['@odata.type'] === '#microsoft.graph.group' || it.id)
    .map(it => ({ id: (it.id || '').toLowerCase(), displayName: it.displayName || '' }));
  return groups;
}

module.exports = async function (context, req) {
  try {
    const principal = readClientPrincipal(req);

    if (!principal) {
      context.res = {
        status: 401,
        headers: { 'Content-Type': 'application/json' },
        body: { error: 'Usuário não autenticado (x-ms-client-principal ausente).' }
      };
      return;
    }

    const userOid = extractUserOid(principal);
    if (!userOid) {
      context.res = {
        status: 400,
        headers: { 'Content-Type': 'application/json' },
        body: { error: 'Não foi possível identificar o OID do usuário.', principal }
      };
      return;
    }

    const tenantId = process.env.AAD_TENANT_ID || '4b30645b-0888-45c0-9481-712bde435ffd';
    const clientId = process.env.AAD_CLIENT_ID;
    const clientSecret = process.env.AAD_CLIENT_SECRET;

    if (!clientId || !clientSecret) {
      context.res = {
        status: 500,
        headers: { 'Content-Type': 'application/json' },
        body: { error: 'AAD_CLIENT_ID ou AAD_CLIENT_SECRET não configurados nas App Settings.' }
      };
      return;
    }

    const graphToken = await getGraphAppToken(tenantId, clientId, clientSecret);
    const groups = await fetchUserGroups(graphToken, userOid);

    // Mapeia GUIDs → roles internas
    const roleMap = {};
    for (const [k, v] of Object.entries(GROUP_TO_ROLE)) roleMap[k.toLowerCase()] = v;
    const roles = groups
      .map(g => roleMap[g.id])
      .filter(Boolean);

    // Roles "agregadas" pra simplificar o front-end
    const aggregated = new Set(roles);
    const isGestor = roles.some(r => r.startsWith('gestor_'));
    if (isGestor) aggregated.add('gestor');

    context.res = {
      status: 200,
      headers: {
        'Content-Type': 'application/json',
        'Cache-Control': 'no-store'
      },
      body: {
        userId: userOid,
        userDetails: principal.userDetails || null,
        identityProvider: principal.identityProvider || null,
        groups,          // [{ id, displayName }]
        roles,           // ex.: ['submitter','gestor_financeira']
        rolesAggregated: Array.from(aggregated) // inclui 'gestor' genérico
      }
    };
  } catch (err) {
    context.log.error('MeusGrupos error:', err);
    context.res = {
      status: 500,
      headers: { 'Content-Type': 'application/json' },
      body: { error: err.message }
    };
  }
};
