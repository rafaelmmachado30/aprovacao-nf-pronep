/**
 * Sistema de Aprovacao de NF - EditarFornecedor (PATCH)
 *
 * Edita fornecedor existente. Aceita PATCH ou POST (Functions runtime as vezes
 * normaliza pra POST).
 *
 * Body JSON: { id: '123', ...campos a atualizar }
 * Campos nao informados nao sao alterados.
 */

require('isomorphic-fetch');
const { getUser } = require('../shared/auth');
const { ClientSecretCredential } = require('@azure/identity');
const { Client } = require('@microsoft/microsoft-graph-client');
const { TokenCredentialAuthenticationProvider } =
  require('@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials');

const LIST_NAME = 'PRONEP-NF-Fornecedores';
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

async function getColMap(client, siteId, listId) {
  const resp = await client.api('/sites/' + siteId + '/lists/' + listId + '/columns').get();
  const map = {};
  for (const c of (resp.value || [])) { if (c.displayName && c.name) map[c.displayName] = c.name; }
  return map;
}

async function resolveSiteAndList(client) {
  if (cache.siteId && cache.listId) return cache;
  const hostname = process.env.SHAREPOINT_SITE_HOSTNAME;
  const sitePath = process.env.SHAREPOINT_SITE_PATH;
  const siteResp = await client.api('/sites/' + hostname + ':' + sitePath).get();
  cache.siteId = siteResp.id;
  const listsResp = await client.api('/sites/' + cache.siteId + '/lists')
    .filter("displayName eq '" + LIST_NAME + "'")
    .get();
  if (!listsResp.value || !listsResp.value.length) throw new Error("Lista nao encontrada");
  cache.listId = listsResp.value[0].id;
  return cache;
}

function buildFieldsPayload(body) {
  // Apenas inclui campos que vieram no body (PATCH parcial)
  const fields = {};
  if (body.razao !== undefined)         fields.Title    = String(body.razao || '').trim();
  if (body.tipoDocumento !== undefined) fields.field_1  = String(body.tipoDocumento || '').trim().toUpperCase();
  if (body.documento !== undefined)     fields.field_2  = String(body.documento || '').trim();
  if (body.nomeFantasia !== undefined)  fields.field_3  = String(body.nomeFantasia || '').trim();
  if (body.unidade !== undefined)       fields.field_4  = String(body.unidade || '').trim().toUpperCase();
  if (body.diretoria !== undefined)     fields.field_5  = String(body.diretoria || '').trim();
  if (body.uf !== undefined)            fields.field_6  = String(body.uf || '').trim().toUpperCase();
  if (body.ativo !== undefined) {
    fields.field_7 = (body.ativo === false || body.ativo === 'Nao' || body.ativo === 'no') ? 'Nao' : 'Sim';
  }
  if (body.telefone !== undefined)      fields.field_8  = String(body.telefone || '').trim();
  if (body.email !== undefined)         fields.field_9  = String(body.email || '').trim().toLowerCase();
  if (body.cidade !== undefined)        fields.field_10 = String(body.cidade || '').trim();
  if (body.cep !== undefined)           fields.field_11 = String(body.cep || '').trim();
  // fields.AtendeTodas eh ajustado dinamicamente no handler com o internal name real
  return fields;
}

module.exports = async function (context, req) {
  const diag = { step: 'start' };
  try {
    const user = await getUser(req);
    if (!user) {
      context.res = { status: 401, body: { error: 'Nao autenticado' } };
      return;
    }
    diag.user = user.email;

    const body = req.body || {};
    const itemId = body.id;
    if (!itemId) {
      context.res = { status: 400, body: { error: 'Campo obrigatorio: id' } };
      return;
    }
    diag.itemId = itemId;

    const fields = buildFieldsPayload(body);

    // Resolve internal name de AtendeTodas (SP pode ter renomeado pra field_X)
    if (body.atendeTodas !== undefined) {
      try {
        const colMap = await getColMap(client, siteId, listId);
        const atendeTodasInternal = colMap['AtendeTodas'] || 'AtendeTodas';
        fields[atendeTodasInternal] = !!body.atendeTodas;
        diag.atendeTodasInternal = atendeTodasInternal;
      } catch (e) {
        // Fallback: usa nome literal
        fields.AtendeTodas = !!body.atendeTodas;
      }
    }
    if (Object.keys(fields).length === 0) {
      context.res = { status: 400, body: { error: 'Nenhum campo pra atualizar' } };
      return;
    }
    diag.fields = fields;

    diag.step = 'graph_client';
    const client = await getGraphClient();

    diag.step = 'resolve_site';
    const { siteId, listId } = await resolveSiteAndList(client);

    diag.step = 'update_fields';
    // PATCH em /items/{id}/fields atualiza so os campos enviados
    const updated = await client
      .api('/sites/' + siteId + '/lists/' + listId + '/items/' + itemId + '/fields')
      .patch(fields);

    context.res = {
      status: 200,
      headers: { 'Content-Type': 'application/json' },
      body: { ok: true, id: itemId, updatedFields: Object.keys(fields), diag: diag }
    };
  } catch (err) {
    context.log && context.log.error && context.log.error('EditarFornecedor error:', err);
    context.res = {
      status: 500,
      headers: { 'Content-Type': 'application/json' },
      body: {
        error: (err && err.message) || String(err),
        statusCode: err && err.statusCode,
        body: err && err.body,
        diag: diag
      }
    };
  }
};
