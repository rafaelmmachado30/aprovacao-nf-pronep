/**
 * GET /api/ListarSubpastasContratos?path=/...
 *
 * Lista APENAS os filhos imediatos de uma pasta no SP de contratos.
 * Resposta rapida (1 chamada Graph), pra BFS no orquestrador front.
 */
module.exports = async function (context, req) {
  try {
    const path = (req.query && req.query.path) || '';
    if (!path) {
      context.res = { status: 400, body: { error: 'path obrigatorio' } };
      return;
    }
    const contratos = require('../shared/contratos');
    const client = contratos.getGraphClient();
    const { driveId } = await contratos.resolveContratosSite(client);
    const listing = await contratos.listarPasta(client, driveId, path);
    context.res = {
      status: 200,
      headers: { 'Content-Type': 'application/json' },
      body: {
        ok: true,
        pasta: path,
        pdfs: listing.files.map(function(f){ return { nome: f.name, id: f.id, path: f.path, size: f.size }; }),
        subpastas: listing.folders.map(function(f){ return { nome: f.name, path: f.path, childCount: f.childCount }; })
      }
    };
  } catch (err) {
    context.res = {
      status: 500,
      headers: { 'Content-Type': 'application/json' },
      body: { error: err.message, statusCode: err.statusCode }
    };
  }
};
