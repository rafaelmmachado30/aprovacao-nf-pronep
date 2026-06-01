/**
 * shared/solHistorico.js
 *
 * Persistencia do historico de conversas da SAN na lista SharePoint
 * PRONEP-NF-SOL-Conversas (memoria persistente entre sessoes/dispositivos).
 *
 * Colunas da lista (resolve internalNames dinamicamente):
 *   Title      — chave curta (timestamp ms)
 *   UserOid    — oid do user (indexado)
 *   UserEmail  — email
 *   Role       — choice: user | assistant | system
 *   Content    — multi-line text
 *   Timestamp  — datetime
 *   ConversaId — opcional (futuro: agrupar sessoes)
 *   TokensUsed — number (so pra assistant msgs)
 */

require('isomorphic-fetch');
const { ClientSecretCredential } = require('@azure/identity');
const { Client } = require('@microsoft/microsoft-graph-client');
const { TokenCredentialAuthenticationProvider } =
  require('@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials');

const LIST_NAME = 'PRONEP-NF-SOL-Conversas';
const cache = { siteId: null, listId: null, colMap: null };

function getGraphClient() {
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
  if (!host || !path) throw new Error('SHAREPOINT_SITE_HOSTNAME/PATH nao configurados');
  const siteResp = await client.api('/sites/' + host + ':' + path).get();
  cache.siteId = siteResp.id;
  const lists = await client.api('/sites/' + cache.siteId + '/lists')
    .filter("displayName eq '" + LIST_NAME + "'").get();
  if (!lists.value || !lists.value.length) {
    throw new Error("Lista '" + LIST_NAME + "' nao encontrada no site");
  }
  cache.listId = lists.value[0].id;
  return cache;
}

async function getColMap(client) {
  if (cache.colMap) return cache.colMap;
  const { siteId, listId } = await resolveSiteAndList(client);
  const resp = await client.api('/sites/' + siteId + '/lists/' + listId + '/columns').get();
  const map = {};
  for (const c of (resp.value || [])) {
    if (c.displayName && c.name) map[c.displayName] = c.name;
  }
  cache.colMap = map;
  return map;
}

/**
 * Le as ultimas N mensagens do user, em ordem cronologica (mais antiga primeiro).
 * @returns Array<{role, content, timestamp, tokensUsed?}>
 */
async function lerUltimas(user, limit) {
  const lim = Math.max(1, Math.min(parseInt(limit || 20), 100));
  const client = getGraphClient();
  const { siteId, listId } = await resolveSiteAndList(client);
  const cm = await getColMap(client);
  const oidCol = cm['UserOid'] || 'UserOid';
  const roleCol = cm['Role'] || 'Role';
  const contentCol = cm['Content'] || 'Content';
  const tsCol = cm['Timestamp'] || 'Timestamp';
  const tokCol = cm['TokensUsed'] || 'TokensUsed';

  // Query: filter por UserOid, order by Timestamp desc, top N
  const items = await client.api('/sites/' + siteId + '/lists/' + listId + '/items')
    .expand('fields')
    .filter("fields/" + oidCol + " eq '" + (user.oid || '') + "'")
    .orderby("fields/" + tsCol + " desc")
    .top(lim)
    .header('Prefer', 'HonorNonIndexedQueriesWarningMayFailRandomly')
    .get();

  const list = (items.value || []).map(it => {
    const f = it.fields || {};
    return {
      id: it.id,
      role: f[roleCol] || 'user',
      content: f[contentCol] || '',
      timestamp: f[tsCol] || it.lastModifiedDateTime,
      tokensUsed: f[tokCol] || 0
    };
  });
  // Reverte pra ordem cronologica (mais antigo primeiro) — front renderiza topo->fundo
  list.reverse();
  return list;
}

/**
 * Salva uma mensagem (user ou assistant) no historico.
 * Best-effort: nao throw — apenas loga em caso de falha pra nao impactar a UX.
 */
async function salvar(user, role, content, opts) {
  opts = opts || {};
  try {
    if (!user || !user.oid) return { ok: false, reason: 'sem oid' };
    if (!role || !content) return { ok: false, reason: 'role/content vazios' };
    const client = getGraphClient();
    const { siteId, listId } = await resolveSiteAndList(client);
    const cm = await getColMap(client);
    const oidCol = cm['UserOid'] || 'UserOid';
    const emailCol = cm['UserEmail'] || 'UserEmail';
    const roleCol = cm['Role'] || 'Role';
    const contentCol = cm['Content'] || 'Content';
    const tsCol = cm['Timestamp'] || 'Timestamp';
    const conversaCol = cm['ConversaId'] || 'ConversaId';
    const tokCol = cm['TokensUsed'] || 'TokensUsed';

    const fields = {};
    fields['Title'] = String(Date.now());
    fields[oidCol] = user.oid;
    fields[emailCol] = user.email || '';
    fields[roleCol] = role;
    fields[contentCol] = String(content).slice(0, 32000); // trunca pra evitar erro de tamanho
    fields[tsCol] = new Date().toISOString();
    if (opts.conversaId) fields[conversaCol] = opts.conversaId;
    if (opts.tokensUsed) fields[tokCol] = opts.tokensUsed;

    const created = await client.api('/sites/' + siteId + '/lists/' + listId + '/items')
      .post({ fields });
    return { ok: true, id: created.id };
  } catch (e) {
    console.error('[solHistorico.salvar] erro:', e && e.message);
    return { ok: false, error: e && e.message };
  }
}

/**
 * Deleta todas as mensagens do usuario logado.
 */
async function limparHistorico(user) {
  if (!user || !user.oid) throw new Error('user.oid obrigatorio');
  const client = getGraphClient();
  const { siteId, listId } = await resolveSiteAndList(client);
  const cm = await getColMap(client);
  const oidCol = cm['UserOid'] || 'UserOid';

  // Lista todos os items do user (paginando pra pegar tudo)
  let removidos = 0;
  let next = "/sites/" + siteId + "/lists/" + listId + "/items?$expand=fields&$filter=fields/" + oidCol + " eq '" + user.oid + "'&$top=100";
  while (next) {
    const page = await client.api(next).header('Prefer', 'HonorNonIndexedQueriesWarningMayFailRandomly').get();
    const items = page.value || [];
    for (const it of items) {
      try {
        await client.api('/sites/' + siteId + '/lists/' + listId + '/items/' + it.id).delete();
        removidos++;
      } catch (e) {
        console.error('[limparHistorico] falha ao deletar item', it.id, e && e.message);
      }
    }
    next = page['@odata.nextLink'] || null;
  }
  return { removidos };
}

module.exports = { lerUltimas, salvar, limparHistorico };
