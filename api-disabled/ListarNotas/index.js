/**
 * Listar Notas Fiscais
 *
 * GET /api/ListarNotas?status=pendente&unidade=RJ&diretoria=Suprimentos
 *
 * Headers do SWA já contêm o usuário autenticado em X-MS-CLIENT-PRINCIPAL.
 * Filtra automaticamente pelo escopo do usuário (Submitter vê só as próprias;
 * Gestor vê só sua diretoria; Financeiro/Admin veem todas).
 *
 * Status: ESQUELETO - implementar na Sprint 2/3.
 *
 * TODO:
 *  1. Decodificar X-MS-CLIENT-PRINCIPAL para obter email e roles
 *  2. Conectar no SharePoint via Microsoft Graph (Managed Identity)
 *  3. Consultar a lista "NotasFiscais" com filtros do query string
 *  4. Aplicar filtro RBAC: Submitter -> próprias; Gestor -> matriz aprovadoresReais
 *  5. Devolver JSON paginado
 */

module.exports = async function (context, req) {
  context.log('=== ListarNotas invoked ===');
  // Decodifica o principal do SWA
  const headerB64 = req.headers['x-ms-client-principal'];
  const principal = headerB64
    ? JSON.parse(Buffer.from(headerB64, 'base64').toString('utf-8'))
    : null;

  if (!principal) {
    context.res = { status: 401, body: { error: 'Não autenticado' } };
    return;
  }

  // Placeholder: devolve lista vazia.
  // Implementação real lê do SharePoint List 'NotasFiscais' via Graph.
  context.res = {
    status: 200,
    body: {
      principal: { userId: principal.userId, userDetails: principal.userDetails, roles: principal.userRoles },
      filtros: req.query,
      itens: [],
      _todo: 'Conectar com SharePoint List NotasFiscais via Microsoft Graph SDK'
    }
  };
};
