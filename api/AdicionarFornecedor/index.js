/**
 * Sistema de Aprovacao de NF - AdicionarFornecedor (POST)
 *
 * Adiciona novo fornecedor na lista PRONEP-NF-Fornecedores do SharePoint via Graph API.
 *
 * Body JSON: {
 *   razao, tipoDocumento, documento, nomeFantasia, unidade, diretoria, uf, ativo,
 *   telefone, email, cidade, cep
 * }
 *
 * Retorna { ok: true, id, fornecedor } ou erro.
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

async function resolveSiteAndList(client) {
  if (cache.siteId && cache.listId) return cache;
  const hostname = process.env.SHAREPOINT_SITE_HOSTNAME;
  const sitePath = process.env.SHAREPOINT_SITE_PATH;
  const siteResp = await client.api('/sites/' + hostname + ':' + sitePath).get();
  cache.siteId = siteResp.id;
  const listsResp = await client.api('/sites/' + cache.siteId + '/lists')
    .filter("displayName eq '" + LIST_NAME + "'")
    .get();
  if (!listsResp.value || !listsResp.value.length) {
    throw new Error("Lista '" + LIST_NAME + "' nao encontrada");
  }
  cache.listId = listsResp.value[0].id;
  return cache;
}

function onlyDigits(s) { return String(s || '').replace(/\D/g, ''); }

function buildFieldsPayload(body) {
  // Mapeamento: Title=razao, field_1=tipoDoc, field_2=doc, field_3=fantasia,
  // field_4=unidade, field_5=diretoria, field_6=uf, field_7=ativo, field_8=tel,
  // field_9=email, field_10=cidade, field_11=cep
  const fields = {};
  if (body.razao)          fields.Title    = String(body.razao).trim();
  if (body.tipoDocumento)  fields.field_1  = String(body.tipoDocumento).trim().toUpperCase(); // CNPJ ou CPF
  if (body.documento)      fields.field_2  = String(body.documento).trim();
  if (body.nomeFantasia)   fields.field_3  = String(body.nomeFantasia).trim();
  if (body.unidade)        fields.field_4  = String(body.unidade).trim().toUpperCase();
  if (body.diretoria)      fields.field_5  = String(body.diretoria).trim();
  if (body.uf)             fields.field_6  = String(body.uf).trim().toUpperCase();
  fields.field_7 = (body.ativo === false || body.ativo === 'Nao' || body.ativo === 'no') ? 'Nao' : 'Sim';
  if (body.telefone)       fields.field_8  = String(body.telefone).trim();
  if (body.email)          fields.field_9  = String(body.email).trim().toLowerCase();
  if (body.cidade)         fields.field_10 = String(body.cidade).trim();
  if (body.cep)            fields.field_11 = String(body.cep).trim();
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
    if (!body.razao || !String(body.razao).trim()) {
      context.res = { status: 400, body: { error: 'Campo obrigatorio: razao' } };
      return;
    }
    if (!body.documento || !String(body.documento).trim()) {
      context.res = { status: 400, body: { error: 'Campo obrigatorio: documento (CNPJ ou CPF)' } };
      return;
    }
    if (!body.unidade || !body.diretoria) {
      context.res = { status: 400, body: { error: 'Campos obrigatorios: unidade e diretoria' } };
      return;
    }

    diag.step = 'graph_client';
    const client = await getGraphClient();

    diag.step = 'resolve_site';
    const { siteId, listId } = await resolveSiteAndList(client);
    diag.siteId = siteId; diag.listId = listId;

    // Verifica se ja existe (mesmo documento) pra evitar duplicata
    const docDigits = onlyDigits(body.documento);
    diag.step = 'check_duplicate';
    if (docDigits) {
      const checkResp = await client.api('/sites/' + siteId + '/lists/' + listId + '/items')
        .expand('fields(select=Title,field_2)')
        .filter("fields/field_2 eq '" + String(body.documento).trim().replace(/'/g, "''") + "'")
        .header('Prefer', 'HonorNonIndexedQueriesWarningMayFailRandomly')
        .top(1)
        .get();
      if (checkResp.value && checkResp.value.length > 0) {
        context.res = {
          status: 409,
          body: {
            error: 'Fornecedor ja existe com este documento',
            existing: { id: checkResp.value[0].id, razao: checkResp.value[0].fields.Title }
          }
        };
        return;
      }
    }

    diag.step = 'create_item';
    const fields = buildFieldsPayload(body);
    diag.fields = fields;
    const created = await client.api('/sites/' + siteId + '/lists/' + listId + '/items')
      .post({ fields: fields });

    context.res = {
      status: 201,
      headers: { 'Content-Type': 'application/json' },
      body: {
        ok: true,
        id: created.id,
        razao: fields.Title,
        documento: fields.field_2,
        diag: diag
      }
    };
  } catch (err) {
    context.log && context.log.error && context.log.error('AdicionarFornecedor error:', err);
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
