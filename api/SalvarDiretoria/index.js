/**
 * /api/SalvarDiretoria (POST) — ADMIN ONLY.
 *
 * Atualiza o APROVADOR de uma linha da lista PRONEP-NF-Diretorias (Unidade x Diretoria).
 * Grava field_3 (email) e field_4 (nome). Fonte de verdade do roteamento de NFs.
 *
 * OBS: afeta apenas NFs FUTURAS. As NFs ja pendentes guardam o AprovadorAtual do momento
 * da criacao (nao sao reescritas aqui) — comportamento definido com o negocio.
 *
 * Body: { id: "<listItemId>", email: "nome@pronep.com.br", nome: "Nome do aprovador" }
 */

require('isomorphic-fetch');
const { resolveAuthz } = require('../shared/authz');
const { ClientSecretCredential } = require('@azure/identity');
const { Client } = require('@microsoft/microsoft-graph-client');
const { TokenCredentialAuthenticationProvider } =
  require('@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials');

const LIST_NAME = 'PRONEP-NF-Diretorias';
const cache = { siteId: null, listId: null };

function getGraphClient() {
  const credential = new ClientSecretCredential(
    process.env.AAD_TENANT_ID, process.env.AAD_CLIENT_ID, process.env.AAD_CLIENT_SECRET
  );
  const authProvider = new TokenCredentialAuthenticationProvider(credential, {
    scopes: ['https://graph.microsoft.com/.default']
  });
  return Client.initWithMiddleware({ authProvider });
}

async function resolveSiteAndList(client) {
  if (cache.siteId && cache.listId) return cache;
  const host = process.env.SHAREPOINT_SITE_HOSTNAME;
  const path = process.env.SHAREPOINT_SITE_PATH;
  if (!host || !path) throw new Error('SHAREPOINT_* obrigatorias');
  const siteResp = await client.api('/sites/' + host + ':' + path).get();
  cache.siteId = siteResp.id;
  const listsResp = await client.api('/sites/' + cache.siteId + '/lists')
    .filter("displayName eq '" + LIST_NAME + "'").get();
  if (!listsResp.value || !listsResp.value.length) throw new Error("Lista '" + LIST_NAME + "' nao encontrada");
  cache.listId = listsResp.value[0].id;
  return cache;
}

module.exports = async function (context, req) {
  try {
    // Admin OU TI podem alterar aprovadores.
    const authz = await resolveAuthz(req);
    if (!authz) { context.res = { status: 401, headers: { 'Content-Type': 'application/json' }, body: { error: 'Nao autenticado' } }; return; }
    const podeEditar = authz.isAdmin || (authz.roles || []).indexOf('ti') >= 0;
    if (!podeEditar) { context.res = { status: 403, headers: { 'Content-Type': 'application/json' }, body: { error: 'Acesso restrito a Admin ou TI' } }; return; }

    const body = req.body || {};
    const id = String(body.id || '').trim();
    const email = String(body.email || '').trim().toLowerCase();
    const nome = String(body.nome || '').trim();

    if (!id) { context.res = { status: 400, headers: { 'Content-Type': 'application/json' }, body: { error: 'id obrigatorio' } }; return; }
    if (!email || !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) {
      context.res = { status: 400, headers: { 'Content-Type': 'application/json' }, body: { error: 'email invalido' } }; return;
    }
    if (!/@pronep\.com\.br$/i.test(email)) {
      context.res = { status: 400, headers: { 'Content-Type': 'application/json' }, body: { error: 'O e-mail do aprovador deve ser @pronep.com.br' } }; return;
    }

    const client = getGraphClient();
    const { siteId, listId } = await resolveSiteAndList(client);

    // field_3 = Email do aprovador · field_4 = Nome do aprovador
    await client.api('/sites/' + siteId + '/lists/' + listId + '/items/' + id + '/fields')
      .patch({ field_3: email, field_4: nome });

    context.res = { status: 200, headers: { 'Content-Type': 'application/json', 'Cache-Control': 'no-store' },
      body: { ok: true, id: id, email: email, nome: nome, atualizadoPor: authz.email } };
  } catch (err) {
    context.log && context.log.error && context.log.error('SalvarDiretoria:', err);
    context.res = { status: 500, headers: { 'Content-Type': 'application/json' },
      body: { error: (err && err.message) || String(err) } };
  }
};
