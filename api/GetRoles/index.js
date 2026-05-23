/**
 * Sistema de Aprovação de NF — GetRoles
 *
 * VERSÃO DE TESTE: retorna roles FIXAS pra confirmar que o SWA está chamando este endpoint.
 * Se as roles fixas aparecerem no `.auth/me` após relogar, o SWA está chamando.
 * Se não aparecerem, o `rolesSource` no staticwebapp.config.json não está sendo executado.
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

const GROUP_CLAIM_TYPES = [
  'groups',
  'http://schemas.microsoft.com/ws/2008/06/identity/claims/groupsid',
  'http://schemas.microsoft.com/claims/groups',
  'http://schemas.xmlsoap.org/claims/Group'
];

module.exports = async function (context, req) {
  try {
    context.log('=== GetRoles invoked ===');
    context.log('Method:', req.method);
    context.log('Body:', JSON.stringify(req.body || {}));
    context.log('Headers:', JSON.stringify(Object.keys(req.headers || {})));

    // TESTE: ignora totalmente o body e SEMPRE retorna roles fixas.
    // Se depois disso "administrador" aparecer no .auth/me, o SWA está chamando o GetRoles.
    // Se "administrador" continuar não aparecendo, o rolesSource não está sendo invocado.
    const TEST_ROLES = ['submitter', 'administrador', 'TEST_HARDCODED'];

    context.log('Returning TEST_ROLES:', TEST_ROLES.join(', '));

    // Lógica real (comentada durante teste)
    /*
    const claims = (req.body && req.body.claims) || [];
    const groupClaims = claims.filter(c =>
      GROUP_CLAIM_TYPES.includes(c.typ) ||
      (c.typ && c.typ.toLowerCase().includes('group'))
    );
    const groups = groupClaims.map(c => (c.val || '').toLowerCase());
    const groupToRoleLower = {};
    for (const [k, v] of Object.entries(GROUP_TO_ROLE)) groupToRoleLower[k.toLowerCase()] = v;
    let roles = groups.map(g => groupToRoleLower[g]).filter(Boolean);
    */

    context.res = {
      status: 200,
      headers: { 'Content-Type': 'application/json' },
      body: { roles: TEST_ROLES }
    };
  } catch (err) {
    context.log.error('GetRoles error:', err);
    context.res = {
      status: 500,
      body: { error: err.message, roles: [] }
    };
  }
};
