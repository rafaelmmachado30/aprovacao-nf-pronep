/**
 * shared/conciliacaoRecorrentes.js — Conciliacao MANUAL do Fechamento do Mes ("Merge").
 *
 * Quando o casamento automatico nao acha a NF de uma conta recorrente (porque foi
 * relancada sob CNPJ de filial / diretoria diferente), o gestor aponta manualmente
 * uma NF Aprovada que quita aquela conta NAQUELE mes. Isso evita o falso "atrasada".
 *
 * Guardado como JSON num item da lista PRONEP-NF-Config (Title='recorrentesConciliacao',
 * ConfigJson): { "<chave>@@<AAAA-MM>": { notaId, numero, status, valor, por, em } }
 *   chave = chaveRecorrente (cnpj|diretoria|unidade);  o vinculo vale SO para o mes.
 */

require('isomorphic-fetch');

const LIST_CONFIG = 'PRONEP-NF-Config';
const CONC_TITLE = 'recorrentesConciliacao';
const SEP = '@@';

const _cfgCache = {};
async function resolveConfigListId(client, siteId) {
  if (_cfgCache[siteId]) return _cfgCache[siteId];
  const lists = await client.api('/sites/' + siteId + '/lists')
    .filter("displayName eq '" + LIST_CONFIG + "'").get();
  const id = (lists.value && lists.value.length) ? lists.value[0].id : null;
  if (id) _cfgCache[siteId] = id;
  return id;
}

async function findConcItem(client, siteId, listConfigId) {
  const resp = await client.api('/sites/' + siteId + '/lists/' + listConfigId + '/items')
    .expand('fields').top(50).get();
  return (resp.value || []).find(function (x) { return x.fields && x.fields.Title === CONC_TITLE; }) || null;
}

function chaveConc(chave, mes) {
  return String(chave || '') + SEP + String(mes || '');
}

// Le o mapa { "chave@@mes": {...} }. {} se nao existe / invalido.
async function lerConciliacoes(client, siteId, listConfigId) {
  try {
    const cfgId = listConfigId || await resolveConfigListId(client, siteId);
    if (!cfgId) return {};
    const item = await findConcItem(client, siteId, cfgId);
    if (item && item.fields && item.fields.ConfigJson) {
      const m = JSON.parse(item.fields.ConfigJson);
      if (m && typeof m === 'object' && !Array.isArray(m)) return m;
    }
  } catch (e) { /* sem config -> {} */ }
  return {};
}

async function salvarConciliacoes(client, siteId, listConfigId, mapa) {
  const cfgId = listConfigId || await resolveConfigListId(client, siteId);
  if (!cfgId) throw new Error("Lista '" + LIST_CONFIG + "' nao encontrada");
  const json = JSON.stringify(mapa || {});
  const item = await findConcItem(client, siteId, cfgId);
  if (item) {
    await client.api('/sites/' + siteId + '/lists/' + cfgId + '/items/' + item.id + '/fields')
      .patch({ ConfigJson: json });
    return { ok: true, itemId: item.id, action: 'updated' };
  }
  const created = await client.api('/sites/' + siteId + '/lists/' + cfgId + '/items')
    .post({ fields: { Title: CONC_TITLE, ConfigJson: json } });
  return { ok: true, itemId: created.id, action: 'created' };
}

module.exports = {
  LIST_CONFIG,
  CONC_TITLE,
  SEP,
  resolveConfigListId,
  chaveConc,
  lerConciliacoes,
  salvarConciliacoes
};
