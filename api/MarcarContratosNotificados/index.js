/**
 * POST /api/MarcarContratosNotificados
 *
 * Marca um ou mais contratos como "notificados" para uma janela especifica,
 * adicionando linha _alerta_<janela>=<data> em Observacoes do contrato.
 *
 * Body:
 *   { marcadores: [ { id: "<spListItemId>", janela: 30|60|90|"vencido" }, ... ] }
 *
 * Compartilha o MESMO campo de dedup do AlertaContratosDiario (email). Apos
 * marcar via SAN proativa, o contrato nao volta a aparecer pra essa janela
 * em emails diarios nem em saudacoes proativas.
 *
 * Idempotente: se a linha ja existe, nao duplica.
 * Custo Claude: ZERO. So patch SP.
 */

require('isomorphic-fetch');
const { getGraphClient } = require('../shared/graph');
const { getUser } = require('../shared/auth');

const LIST_CONTRATOS = 'PRONEP-NF-Contratos';

const _cache = { siteId: null, listId: null, colMap: null, invColMap: null };

async function resolveSite(client) {
  if (_cache.siteId && _cache.listId) return _cache;
  const host = process.env.SHAREPOINT_SITE_HOSTNAME;
  const path = process.env.SHAREPOINT_SITE_PATH;
  const siteResp = await client.api('/sites/' + host + ':' + path).get();
  _cache.siteId = siteResp.id;
  const lists = await client.api('/sites/' + _cache.siteId + '/lists').get();
  for (const l of (lists.value || [])) {
    if (l.displayName === LIST_CONTRATOS) _cache.listId = l.id;
  }
  if (!_cache.listId) throw new Error('Lista ' + LIST_CONTRATOS + ' nao encontrada');
  const cols = await client.api('/sites/' + _cache.siteId + '/lists/' + _cache.listId + '/columns').get();
  _cache.colMap = {}; _cache.invColMap = {};
  for (const c of (cols.value || [])) {
    if (c.displayName && c.name) { _cache.colMap[c.displayName] = c.name; _cache.invColMap[c.name] = c.displayName; }
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
    const body = req.body || {};
    const marcadores = Array.isArray(body.marcadores) ? body.marcadores : [];
    if (!marcadores.length) {
      context.res = { status: 200, body: { ok: true, processados: 0, msg: 'nada a marcar' } };
      return;
    }
    const client = getGraphClient();
    const { siteId, listId, colMap } = await resolveSite(client);
    const colObs = colMap['Observacoes'] || 'Observacoes';

    const hoje = new Date(new Date().getTime() - 3*60*60*1000).toISOString().substring(0,10);
    const stats = { ok: 0, jaTinha: 0, erros: 0 };
    const detalhes = [];

    for (const m of marcadores) {
      if (!m || !m.id || !m.janela) { stats.erros++; continue; }
      const janela = String(m.janela);
      // janela aceita: '30', '60', '90', 'vencido'
      if (!['30','60','90','vencido'].includes(janela)) { stats.erros++; continue; }
      const tag = '_alerta_' + janela + '=';
      try {
        // Le observacoes atuais
        const item = await client.api('/sites/' + siteId + '/lists/' + listId + '/items/' + m.id + '?expand=fields').get();
        const obsAtuais = (item.fields && (item.fields[colObs] || item.fields.Observacoes)) || '';
        // Pra janela 'vencido', SUBSTITUI a linha anterior (pra atualizar a data)
        // Pra outras janelas, so insere se nao existe
        let novaObs;
        if (janela === 'vencido') {
          // Remove qualquer _alerta_vencido= anterior e adiciona novo
          const sem = String(obsAtuais).replace(/(^|\n)_alerta_vencido=\d{4}-\d{2}-\d{2}/g, '').replace(/^\n/, '');
          novaObs = (sem ? sem + '\n' : '') + tag + hoje;
        } else {
          if (obsAtuais.indexOf(tag) >= 0) { stats.jaTinha++; detalhes.push({ id: m.id, janela, status: 'ja_tinha' }); continue; }
          novaObs = (obsAtuais ? obsAtuais + '\n' : '') + tag + hoje;
        }
        await client.api('/sites/' + siteId + '/lists/' + listId + '/items/' + m.id + '/fields')
          .patch({ [colObs]: novaObs });
        stats.ok++;
        detalhes.push({ id: m.id, janela, status: 'marcado' });
      } catch (e) {
        stats.erros++;
        detalhes.push({ id: m.id, janela, status: 'erro', error: e.message });
      }
    }

    context.res = {
      status: 200,
      headers: { 'Content-Type': 'application/json' },
      body: { ok: true, processados: marcadores.length, stats, detalhes }
    };
  } catch (err) {
    context.log && context.log.error && context.log.error('MarcarContratosNotificados:', err);
    context.res = { status: 500, body: { error: err.message, stack: (err.stack || '').split('\n').slice(0, 6) } };
  }
};
