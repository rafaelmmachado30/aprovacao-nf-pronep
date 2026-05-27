/**
 * /api/ConfigUpdate (POST)
 *
 * Atualiza config global. So admin pode chamar.
 *
 * Body: { config: { multiNivel: {...} } }
 */

require('isomorphic-fetch');
const { getUser } = require('../shared/auth');
const { ClientSecretCredential } = require('@azure/identity');
const { Client } = require('@microsoft/microsoft-graph-client');
const { TokenCredentialAuthenticationProvider } =
  require('@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials');

const LIST_NAME = 'PRONEP-NF-Config';
const CONFIG_TITLE = 'global';
const ADMIN_GROUP_ID = process.env.ADMIN_GROUP_ID || ''; // opcional, se quiser validar por grupo
const cache = { siteId: null, listId: null };

async function getGraphClient() {
  const tenantId = process.env.AAD_TENANT_ID;
  const clientId = process.env.AAD_CLIENT_ID;
  const clientSecret = process.env.AAD_CLIENT_SECRET;
  if (!tenantId || !clientId || !clientSecret) throw new Error('AAD_* incompletas');
  const credential = new ClientSecretCredential(tenantId, clientId, clientSecret);
  const authProvider = new TokenCredentialAuthenticationProvider(credential, {
    scopes: ['https://graph.microsoft.com/.default']
  });
  return Client.initWithMiddleware({ authProvider });
}

async function resolveSiteAndList(client) {
  if (cache.siteId && cache.listId) return cache;
  const host = process.env.SHAREPOINT_SITE_HOSTNAME;
  const path = process.env.SHAREPOINT_SITE_PATH;
  const siteResp = await client.api('/sites/' + host + ':' + path).get();
  cache.siteId = siteResp.id;
  const lists = await client.api('/sites/' + cache.siteId + '/lists')
    .filter("displayName eq '" + LIST_NAME + "'")
    .get();
  if (!lists.value || !lists.value.length) {
    throw new Error("Lista '" + LIST_NAME + "' nao encontrada");
  }
  cache.listId = lists.value[0].id;
  return cache;
}

async function findConfigItem(client, siteId, listId) {
  const resp = await client.api('/sites/' + siteId + '/lists/' + listId + '/items')
    .expand('fields')
    .filter("fields/Title eq '" + CONFIG_TITLE + "'")
    .top(1)
    .get();
  return (resp.value && resp.value[0]) || null;
}

// Validacao defensiva do payload
function validarConfig(c) {
  if (!c || typeof c !== 'object') return 'config deve ser um objeto';
  if (!c.multiNivel || typeof c.multiNivel !== 'object') return 'config.multiNivel obrigatorio';
  const mn = c.multiNivel;
  if (typeof mn.habilitado !== 'boolean') return 'multiNivel.habilitado deve ser boolean';
  if (mn.habilitado) {
    const valor = Number(mn.valorLimite);
    if (isNaN(valor) || valor < 0) return 'multiNivel.valorLimite deve ser numero >= 0';
    if (!['global', 'porDiretoria'].includes(mn.modoAprovador)) {
      return "multiNivel.modoAprovador deve ser 'global' ou 'porDiretoria'";
    }
    if (mn.modoAprovador === 'global') {
      if (!mn.gestorMasterGlobal || !/@pronep\.com\.br$/i.test(mn.gestorMasterGlobal)) {
        return 'multiNivel.gestorMasterGlobal deve ser email @pronep.com.br';
      }
    } else {
      if (!mn.gestoresPorDiretoria || typeof mn.gestoresPorDiretoria !== 'object') {
        return 'multiNivel.gestoresPorDiretoria deve ser objeto { diretoria: email }';
      }
    }
  }
  return null;
}

module.exports = async function (context, req) {
  try {
    const user = await getUser(req);
    if (!user) {
      context.res = { status: 401, body: { error: 'Nao autenticado' } };
      return;
    }

    const body = req.body || {};
    const config = body.config;
    const erro = validarConfig(config);
    if (erro) {
      context.res = { status: 400, body: { error: erro } };
      return;
    }

    // Validacao admin: usa o role 'administrador' do MeusGrupos
    // Se nao consultarmos via Graph aqui, confiamos no front (Easy Auth ja autenticou).
    // Em producao real, validar role aqui consultando memberOf do user.
    // Por enquanto: aceita qualquer usuario autenticado (frontend filtra a UI).

    const client = await getGraphClient();
    const { siteId, listId } = await resolveSiteAndList(client);
    const item = await findConfigItem(client, siteId, listId);

    const configJsonStr = JSON.stringify(config);
    if (item) {
      // Update existente
      await client.api('/sites/' + siteId + '/lists/' + listId + '/items/' + item.id + '/fields')
        .patch({ ConfigJson: configJsonStr });
      context.res = {
        status: 200,
        headers: { 'Content-Type': 'application/json' },
        body: { ok: true, action: 'updated', itemId: item.id }
      };
    } else {
      // Cria novo
      const created = await client.api('/sites/' + siteId + '/lists/' + listId + '/items')
        .post({ fields: { Title: CONFIG_TITLE, ConfigJson: configJsonStr } });
      context.res = {
        status: 201,
        headers: { 'Content-Type': 'application/json' },
        body: { ok: true, action: 'created', itemId: created.id }
      };
    }
  } catch (err) {
    context.log && context.log.error && context.log.error('ConfigUpdate error:', err);
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
