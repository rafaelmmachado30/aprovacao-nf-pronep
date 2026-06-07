/**
 * POST /api/MarcarNFsRejeitadasVistas
 *
 * Marca NFs rejeitadas como "vistas pelo submetedor" - adiciona
 * `_visto_rejeicao_<email>=<date>` na coluna Observacao da NF.
 *
 * Apos marcada, essa NF nao volta a aparecer na saudacao proativa da SAN
 * pra esse user. Se a NF voltar pra status Lancada (re-submetida) e for
 * rejeitada DE NOVO depois, o status muda e o ciclo recomeca.
 *
 * Body:
 *   { ids: [ "<spListItemId>", ... ] }
 *
 * Idempotente: se a tag ja existe, nao duplica.
 * Custo Claude: ZERO.
 */

require('isomorphic-fetch');
const { ClientSecretCredential } = require('@azure/identity');
const { Client } = require('@microsoft/microsoft-graph-client');
const { TokenCredentialAuthenticationProvider } = require('@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials');
const { getUser } = require('../shared/auth');

const LIST_NOTAS = 'PRONEP-NF-NotasFiscais';

const _cache = { siteId: null, listId: null, colMap: null };

function getGraphClient() {
  const credential = new ClientSecretCredential(
    process.env.AAD_TENANT_ID, process.env.AAD_CLIENT_ID, process.env.AAD_CLIENT_SECRET
  );
  const authProvider = new TokenCredentialAuthenticationProvider(credential, { scopes: ['https://graph.microsoft.com/.default'] });
  return Client.initWithMiddleware({ authProvider });
}

async function resolveSite(client) {
  if (_cache.siteId && _cache.listId) return _cache;
  const host = process.env.SHAREPOINT_SITE_HOSTNAME;
  const path = process.env.SHAREPOINT_SITE_PATH;
  const siteResp = await client.api('/sites/' + host + ':' + path).get();
  _cache.siteId = siteResp.id;
  const lists = await client.api('/sites/' + _cache.siteId + '/lists').get();
  for (const l of (lists.value || [])) {
    if (l.displayName === LIST_NOTAS) _cache.listId = l.id;
  }
  if (!_cache.listId) throw new Error('Lista ' + LIST_NOTAS + ' nao encontrada');
  const cols = await client.api('/sites/' + _cache.siteId + '/lists/' + _cache.listId + '/columns').get();
  _cache.colMap = {};
  for (const c of (cols.value || [])) {
    if (c.displayName && c.name) { _cache.colMap[c.displayName] = c.name; }
  }
  return _cache;
}

module.exports = async function (context, req) {
  try {
    const user = await getUser(req);
    if (!user || !user.email) {
      context.res = { status: 401, body: { error: 'Nao autenticado' } };
      return;
    }
    const emailLow = String(user.email).toLowerCase();
    const body = req.body || {};
    const ids = Array.isArray(body.ids) ? body.ids : [];
    if (!ids.length) {
      context.res = { status: 200, body: { ok: true, processados: 0, msg: 'nada a marcar' } };
      return;
    }
    const client = getGraphClient();
    const { siteId, listId, colMap } = await resolveSite(client);
    // Procura coluna Observacao (pode ter nome variado)
    const colObs = colMap['Observacao'] || colMap['Observacoes'] || colMap['Obs'] || 'Observacao';
    const hoje = new Date(new Date().getTime() - 3*60*60*1000).toISOString().substring(0,10);
    const tag = '_visto_rejeicao_' + emailLow + '=';
    const stats = { ok: 0, jaTinha: 0, erros: 0 };
    const detalhes = [];

    for (const id of ids) {
      try {
        // Le NF atual (precisa ler observacao + lancadoPor pra checar permissao)
        const item = await client.api('/sites/' + siteId + '/lists/' + listId + '/items/' + id + '?expand=fields').get();
        const lancadoPor = String((item.fields && (item.fields.LancadoPor || item.fields[colMap['LancadoPor'] || 'LancadoPor'])) || '').toLowerCase();
        if (lancadoPor !== emailLow) {
          // Seguranca: nao deixa user marcar NFs de outros
          stats.erros++;
          detalhes.push({ id, status: 'permissao_negada', motivo: 'NF nao pertence ao user' });
          continue;
        }
        const obsAtuais = (item.fields && (item.fields[colObs] || item.fields.Observacao || item.fields.Observacoes)) || '';
        if (obsAtuais.indexOf(tag) >= 0) {
          stats.jaTinha++;
          detalhes.push({ id, status: 'ja_visto' });
          continue;
        }
        const novaObs = (obsAtuais ? obsAtuais + '\n' : '') + tag + hoje;
        await client.api('/sites/' + siteId + '/lists/' + listId + '/items/' + id + '/fields')
          .patch({ [colObs]: novaObs });
        stats.ok++;
        detalhes.push({ id, status: 'marcado' });
      } catch (e) {
        stats.erros++;
        detalhes.push({ id, status: 'erro', error: e.message });
      }
    }

    context.res = {
      status: 200,
      headers: { 'Content-Type': 'application/json' },
      body: { ok: true, processados: ids.length, stats, detalhes }
    };
  } catch (err) {
    context.log && context.log.error && context.log.error('MarcarNFsRejeitadasVistas:', err);
    context.res = { status: 500, body: { error: err.message, stack: (err.stack || '').split('\n').slice(0, 6) } };
  }
};
