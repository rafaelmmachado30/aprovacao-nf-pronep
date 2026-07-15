/**
 * /api/MarcarProcessado (POST)
 *
 * Marca/desmarca uma NF como Processado (financeiro liberou pra integracao).
 * Body: { id: '123', processado: true|false }
 */

require('isomorphic-fetch');
const { getUser } = require('../shared/auth');
const { registrar: auditRegistrar } = require('../shared/auditLog');
const { getGraphClient, resolveSiteAndList } = require('../shared/graph');

const LIST_NOTAS = 'PRONEP-NF-NotasFiscais';

async function getColMap(client, siteId, listId) {
  const resp = await client.api('/sites/' + siteId + '/lists/' + listId + '/columns').get();
  const map = {};
  for (const c of (resp.value || [])) { if (c.displayName && c.name) map[c.displayName] = c.name; }
  return map;
}

module.exports = async function (context, req) {
  try {
    const user = await getUser(req);
    if (!user) {
      context.res = { status: 401, body: { error: 'Nao autenticado' } };
      return;
    }
    const body = req.body || {};
    const itemId = body.id;
    const processado = !!body.processado;
    if (!itemId) {
      context.res = { status: 400, body: { error: 'id obrigatorio' } };
      return;
    }
    const client = await getGraphClient();
    const { siteId, listId } = await resolveSiteAndList(client, LIST_NOTAS);
    // Resolve internal name de Processado (SP pode renomear)
    let processadoInternal = 'Processado';
    try {
      const cm = await getColMap(client, siteId, listId);
      processadoInternal = cm['Processado'] || 'Processado';
    } catch (e) {}
    const patch = {};
    patch[processadoInternal] = processado;
    await client.api('/sites/' + siteId + '/lists/' + listId + '/items/' + itemId + '/fields')
      .patch(patch);

    auditRegistrar(user, 'processado',
      { tipo: 'nf', id: itemId },
      'sucesso',
      { processado: processado }
    ).catch(function(){});

    context.res = {
      status: 200,
      headers: { 'Content-Type': 'application/json' },
      body: { ok: true, id: itemId, processado: processado, por: user.email }
    };
  } catch (err) {
    context.log && context.log.error && context.log.error('MarcarProcessado:', err);
    context.res = {
      status: 500,
      headers: { 'Content-Type': 'application/json' },
      body: {
        error: (err && err.message) || String(err),
        statusCode: err && err.statusCode,
        graphBody: err && err.body
      }
    };
  }
};
