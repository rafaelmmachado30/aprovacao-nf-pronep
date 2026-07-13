/**
 * shared/auditLog.js — Trilha de auditoria do sistema.
 *
 * Registra eventos de ESCRITA no SharePoint (lista PRONEP-NF-AuditLog) pra
 * compliance e rastreabilidade. Retencao: indefinida (admin pode purgar quando
 * passar dos limites do plano SP).
 *
 * Acoes registradas (tudo que muda dados):
 *   - lancamento, aprovacao, rejeicao, processado
 *   - fornecedor_criar, fornecedor_editar, fornecedor_inativar
 *   - limpar_base, config_update
 *
 * NAO registra leituras (listar, abrir PDF, ver detalhes).
 *
 * Padrao: SEMPRE best-effort (fire-and-forget). Falha no log nao pode bloquear
 * a operacao principal.
 */

require('isomorphic-fetch');
const { ClientSecretCredential } = require('@azure/identity');
const { Client } = require('@microsoft/microsoft-graph-client');
const { TokenCredentialAuthenticationProvider } =
  require('@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials');

const LIST_NAME = 'PRONEP-NF-AuditLog';
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
 * Registra um evento de auditoria. BEST-EFFORT — falha nunca propaga.
 *
 * @param {Object} user      - { oid, email, name } do user autenticado
 * @param {string} acao      - Ex: 'lancamento', 'aprovacao', 'rejeicao', etc
 * @param {Object} objeto    - { tipo: 'nf'|'fornecedor'|'config', id, numero }
 * @param {string} resultado - 'sucesso' | 'falha' | 'bloqueado'
 * @param {Object} detalhes  - Info extra (motivo, valor, fornecedor, etc) — serializado em JSON
 * @returns {Promise<{ok, id?, error?}>}
 */
async function registrar(user, acao, objeto, resultado, detalhes) {
  try {
    if (!user || !user.oid) return { ok: false, reason: 'sem oid' };
    if (!acao) return { ok: false, reason: 'sem acao' };
    objeto = objeto || {};
    resultado = resultado || 'sucesso';

    const client = getGraphClient();
    const { siteId, listId } = await resolveSiteAndList(client);
    const cm = await getColMap(client);

    const fields = {};
    fields['Title'] = String(Date.now()) + '-' + (acao || '').slice(0, 30);
    fields[cm['Timestamp'] || 'Timestamp'] = new Date().toISOString();
    fields[cm['UserOid'] || 'UserOid'] = user.oid;
    fields[cm['UserEmail'] || 'UserEmail'] = user.email || '';
    fields[cm['Acao'] || 'Acao'] = acao;
    fields[cm['ObjetoTipo'] || 'ObjetoTipo'] = objeto.tipo || 'sistema';
    if (objeto.id) fields[cm['ObjetoId'] || 'ObjetoId'] = String(objeto.id);
    if (objeto.numero) fields[cm['ObjetoNumero'] || 'ObjetoNumero'] = String(objeto.numero);
    fields[cm['Resultado'] || 'Resultado'] = resultado;
    if (detalhes) {
      // Trunca detalhes pra evitar erro de tamanho (Graph aceita ~ 4M, mas SP single-line tem limite)
      const json = typeof detalhes === 'string' ? detalhes : JSON.stringify(detalhes);
      fields[cm['Detalhes'] || 'Detalhes'] = String(json).slice(0, 32000);
    }

    const created = await client.api('/sites/' + siteId + '/lists/' + listId + '/items')
      .post({ fields });
    return { ok: true, id: created.id };
  } catch (e) {
    console.error('[auditLog.registrar] erro (best-effort):', e && e.message);
    return { ok: false, error: e && e.message };
  }
}

/**
 * Lista eventos de auditoria com filtros opcionais.
 * @param {Object} filtros  - { acao?, userOid?, userEmail?, dataDe?, dataAte?, limit? }
 * @returns {Promise<{events: Array, total: number}>}
 */
async function listar(filtros) {
  filtros = filtros || {};
  const limit = Math.max(1, Math.min(parseInt(filtros.limit || 50), 200));
  const client = getGraphClient();
  const { siteId, listId } = await resolveSiteAndList(client);
  const cm = await getColMap(client);

  // Constroi filtro OData
  const conds = [];
  if (filtros.acao) {
    conds.push("fields/" + (cm['Acao'] || 'Acao') + " eq '" + String(filtros.acao).replace(/'/g, "''") + "'");
  }
  if (filtros.resultado) {
    conds.push("fields/" + (cm['Resultado'] || 'Resultado') + " eq '" + String(filtros.resultado).replace(/'/g, "''") + "'");
  }
  if (filtros.userOid) {
    conds.push("fields/" + (cm['UserOid'] || 'UserOid') + " eq '" + String(filtros.userOid).replace(/'/g, "''") + "'");
  }
  if (filtros.userEmail) {
    conds.push("fields/" + (cm['UserEmail'] || 'UserEmail') + " eq '" + String(filtros.userEmail).toLowerCase().replace(/'/g, "''") + "'");
  }
  if (filtros.dataDe) {
    conds.push("fields/" + (cm['Timestamp'] || 'Timestamp') + " ge '" + filtros.dataDe + "'");
  }
  if (filtros.dataAte) {
    conds.push("fields/" + (cm['Timestamp'] || 'Timestamp') + " le '" + filtros.dataAte + "'");
  }
  const filterStr = conds.length ? conds.join(' and ') : null;

  let request = client.api('/sites/' + siteId + '/lists/' + listId + '/items')
    .expand('fields')
    .orderby("fields/" + (cm['Timestamp'] || 'Timestamp') + " desc")
    .top(limit)
    .header('Prefer', 'HonorNonIndexedQueriesWarningMayFailRandomly');
  if (filterStr) request = request.filter(filterStr);
  const resp = await request.get();

  const events = (resp.value || []).map(it => {
    const f = it.fields || {};
    let det = f[cm['Detalhes'] || 'Detalhes'];
    if (det && typeof det === 'string') {
      try { det = JSON.parse(det); } catch (e) { /* deixa como string */ }
    }
    return {
      id: it.id,
      timestamp: f[cm['Timestamp'] || 'Timestamp'] || it.lastModifiedDateTime,
      userOid: f[cm['UserOid'] || 'UserOid'] || '',
      userEmail: f[cm['UserEmail'] || 'UserEmail'] || '',
      acao: f[cm['Acao'] || 'Acao'] || '',
      objetoTipo: f[cm['ObjetoTipo'] || 'ObjetoTipo'] || '',
      objetoId: f[cm['ObjetoId'] || 'ObjetoId'] || '',
      objetoNumero: f[cm['ObjetoNumero'] || 'ObjetoNumero'] || '',
      resultado: f[cm['Resultado'] || 'Resultado'] || 'sucesso',
      detalhes: det || null
    };
  });
  return { events, total: events.length };
}

module.exports = { registrar, listar };
