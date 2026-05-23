/**
 * Sistema de Aprovação de NF — GetRoles
 *
 * Azure Function chamada pelo Azure Static Web Apps após o login.
 * Recebe os claims do usuário (incluindo grupos do Entra ID) e devolve
 * as roles que ele deve ter na plataforma.
 *
 * Configurada no staticwebapp.config.json via auth.rolesSource = /api/GetRoles
 *
 * MAPEAMENTO DE GRUPOS:
 *   Os GUIDs abaixo precisam ser substituídos pelos IDs reais dos grupos
 *   criados no Entra ID. Veja GUIA_ENTRA_ID.md na raiz do projeto para o
 *   passo-a-passo de criação dos 12 grupos.
 *
 *   Grupos a criar:
 *     PRONEP-NF-Submitter
 *     PRONEP-NF-Admin
 *     PRONEP-Financeiro-Gestao        (Izabel Rocha — para negociação D+5)
 *     PRONEP-NF-Gestor-Suprimentos    (Bruno Hioka)
 *     PRONEP-NF-Gestor-Financeira     (Henrico Molina)
 *     PRONEP-NF-Gestor-Tecnologia     (Rafael Machado)
 *     PRONEP-NF-Gestor-Qualidade      (Sabrina Fernandes)
 *     PRONEP-NF-Gestor-RH-DP          (Janilene Santos)
 *     PRONEP-NF-Gestor-Fiscal-Contabil(Janilene Santos)
 *     PRONEP-NF-Gestor-Juridica       (Rafaella Santos)
 *     PRONEP-NF-Gestor-Administrativa (Rafaella Santos)
 *     PRONEP-NF-Gestor-Tecnica-SP     (Nicolle Boas)
 *     PRONEP-NF-Gestor-Tecnica-RJES   (Vitor Amaral)
 */

// ------------------------------------------------------------------
// CONFIGURAÇÃO — TROCAR OS GUIDS ABAIXO PELOS REAIS DEPOIS DE CRIAR
// OS GRUPOS NO ENTRA ID (Azure Portal → Microsoft Entra ID → Grupos)
// ------------------------------------------------------------------
const GROUP_TO_ROLE = {
  // Roles globais
  '01d540d1-8596-42d0-9a20-de5c361c7c96': 'submitter',          // PRONEP-NF-Submitter
  '480a1595-bdc3-492a-9ef2-317f148a237e': 'administrador',      // PRONEP-NF-Admin
  'c2a73d16-4659-4b3c-93a1-0c0fbfaaaa96': 'financeiro_nf',      // PRONEP-NF-Financeiro-Gestao

  // Gestores por diretoria (10 grupos)
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

// Lista de todas as roles de gestor (usada na herança)
const TODAS_ROLES_GESTOR = [
  'gestor_suprimentos', 'gestor_financeira', 'gestor_tecnologia',
  'gestor_qualidade', 'gestor_rh_dp', 'gestor_fiscal_contabil',
  'gestor_juridica', 'gestor_administrativa',
  'gestor_tecnica_sp', 'gestor_tecnica_rjes'
];

// Role "gestor" agregada — verdadeira se o usuário tem QUALQUER role de gestor
// (usada nas rotas do staticwebapp.config.json para regras genéricas)
const ROLE_INHERITANCE = {
  administrador:  ['submitter', 'financeiro_nf', 'gestor', ...TODAS_ROLES_GESTOR],
  // Gestor de qualquer diretoria também ganha "submitter" (pode lançar NFs) e "gestor" (agregada)
  gestor_suprimentos:      ['submitter', 'gestor'],
  gestor_financeira:       ['submitter', 'gestor'],
  gestor_tecnologia:       ['submitter', 'gestor'],
  gestor_qualidade:        ['submitter', 'gestor'],
  gestor_rh_dp:            ['submitter', 'gestor'],
  gestor_fiscal_contabil:  ['submitter', 'gestor'],
  gestor_juridica:         ['submitter', 'gestor'],
  gestor_administrativa:   ['submitter', 'gestor'],
  gestor_tecnica_sp:       ['submitter', 'gestor'],
  gestor_tecnica_rjes:     ['submitter', 'gestor'],
  // Quem está no time financeiro também pode lançar NFs
  financeiro_nf:           ['submitter']
};

// Tipos de claim possíveis pra grupos (varia entre v1/v2 do token AAD)
const GROUP_CLAIM_TYPES = [
  'groups',
  'http://schemas.microsoft.com/ws/2008/06/identity/claims/groupsid',
  'http://schemas.microsoft.com/claims/groups',
  'http://schemas.xmlsoap.org/claims/Group'
];

module.exports = async function (context, req) {
  try {
    context.log('=== GetRoles (Aprovação NF) invoked ===');
    const claims = (req.body && req.body.claims) || [];
    context.log(`Total claims received: ${claims.length}`);

    // Log dos tipos de claim recebidos pra debug
    const claimTypes = [...new Set(claims.map(c => c.typ))];
    context.log('Claim types present:', claimTypes.join(' | '));

    // Extrai grupos tolerando diferentes formatos de claim type
    const groupClaims = claims.filter(c =>
      GROUP_CLAIM_TYPES.includes(c.typ) ||
      (c.typ && c.typ.toLowerCase().includes('group'))
    );
    const groups = groupClaims.map(c => (c.val || '').toLowerCase());
    context.log(`Group claims found: ${groups.length}`);

    // Mapeia grupos -> roles (case-insensitive)
    const groupToRoleLower = {};
    for (const [k, v] of Object.entries(GROUP_TO_ROLE)) {
      groupToRoleLower[k.toLowerCase()] = v;
    }
    let roles = groups
      .map(g => groupToRoleLower[g])
      .filter(Boolean);

    // Aplica herança (até 2 níveis — gestor → submitter+gestor)
    for (let i = 0; i < 2; i++) {
      const novasRoles = [];
      for (const r of roles) {
        if (ROLE_INHERITANCE[r]) novasRoles.push(...ROLE_INHERITANCE[r]);
      }
      roles = [...new Set([...roles, ...novasRoles])];
    }

    context.log(`Final roles assigned: ${roles.join(', ') || '(none)'}`);

    context.res = {
      status: 200,
      body: { roles }
    };
  } catch (err) {
    context.log.error('GetRoles error:', err);
    context.res = {
      status: 500,
      body: { error: err.message, roles: [] }
    };
  }
};
