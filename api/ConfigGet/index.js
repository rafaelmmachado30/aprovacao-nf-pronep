/**
 * /api/ConfigGet
 *
 * Retorna config global do sistema. Schema:
 *   {
 *     multiNivel: {
 *       habilitado: false,
 *       valorLimite: 0,
 *       modoAprovador: 'global' | 'porDiretoria',
 *       gestorMasterGlobal: 'email@pronep.com.br',
 *       gestoresPorDiretoria: { 'Tecnologia': 'email', ... }
 *     }
 *   }
 *
 * Se nao existe config salva, retorna defaults (multi-nivel desabilitado).
 */

require('isomorphic-fetch');
const { getUser } = require('../shared/auth');
const { getGraphClient, resolveSiteAndList } = require('../shared/graph');

const LIST_NAME = 'PRONEP-NF-Config';
const CONFIG_TITLE = 'global';

const DEFAULT_CONFIG = {
  multiNivel: {
    habilitado: false,
    valorLimite: 0,
    modoAprovador: 'global',
    gestorMasterGlobal: '',
    gestoresPorDiretoria: {}
  }
};

async function findConfigItem(client, siteId, listId) {
  // SP nao deixa filtrar Title sem indice. Lista todos (lista tem 1 ou 2 items) e filtra local.
  const resp = await client.api('/sites/' + siteId + '/lists/' + listId + '/items')
    .expand('fields')
    .top(20)
    .get();
  if (!resp.value) return null;
  return resp.value.find(function (x) {
    return x.fields && (x.fields.Title === CONFIG_TITLE);
  }) || null;
}

module.exports = async function (context, req) {
  try {
    const user = await getUser(req);
    if (!user) {
      context.res = { status: 401, body: { error: 'Nao autenticado' } };
      return;
    }

    const client = await getGraphClient();
    const { siteId, listId } = await resolveSiteAndList(client, LIST_NAME);
    const item = await findConfigItem(client, siteId, listId);

    let config = DEFAULT_CONFIG;
    if (item && item.fields && item.fields.ConfigJson) {
      try { config = JSON.parse(item.fields.ConfigJson); }
      catch (e) { context.log && context.log.warn && context.log.warn('ConfigJson invalido, usando defaults'); }
    }

    context.res = {
      status: 200,
      headers: { 'Content-Type': 'application/json' },
      body: { ok: true, config: config, itemId: item ? item.id : null }
    };
  } catch (err) {
    context.log && context.log.error && context.log.error('ConfigGet error:', err);
    context.res = {
      status: 500,
      headers: { 'Content-Type': 'application/json' },
      body: { error: (err && err.message) || String(err), config: DEFAULT_CONFIG }
    };
  }
};
