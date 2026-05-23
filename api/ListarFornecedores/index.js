/**
 * Listar / criar / editar fornecedores
 *
 * GET  /api/ListarFornecedores            -> lista todos (com filtros via querystring)
 * POST /api/ListarFornecedores            -> cria/edita (corpo JSON)
 *
 * Status: ESQUELETO - implementar na Sprint 1.
 *
 * TODO:
 *  - GET: consultar SharePoint List 'Fornecedores' via Graph
 *  - POST: validar CNPJ (DV), checar duplicidade, gravar
 *  - Apenas administradores podem POST/PUT/DELETE (validar role)
 */

module.exports = async function (context, req) {
  context.log('=== ListarFornecedores invoked ===');
  context.res = {
    status: 501,
    body: { error: 'Not implemented yet — Sprint 1', _ver_arquivo: 'api/ListarFornecedores/index.js' }
  };
};
