/**
 * shared/acessoContratos.js — Controle de Acessos a contratos.
 *
 * Desacopla "quem ve quais contratos" da matriz de aprovacao de NF.
 * O mapa { "Diretoria": ["email1","email2"], ... } fica salvo num item da lista
 * PRONEP-NF-Config (Title = 'acessoContratos', campo ConfigJson).
 *
 * Modelo COMPLEMENTAR (decisao do Rafa):
 *   - Diretoria COM config explicita no Controle de Acessos -> vale a config.
 *   - Diretoria SEM config -> fallback: o aprovador de NF daquela diretoria ve os contratos
 *     (PRONEP-NF-Diretorias). Assim ninguem perde acesso antes de configurar.
 *
 * Admin e Juridico (gestor_juridica) veem tudo — tratado fora deste helper.
 */

require('isomorphic-fetch');

const LIST_CONFIG = 'PRONEP-NF-Config';
const ACESSO_TITLE = 'acessoContratos';

// cache do id da lista Config por siteId (instancia da Function reaproveita)
const _cfgCache = {};
async function resolveConfigListId(client, siteId) {
  if (_cfgCache[siteId]) return _cfgCache[siteId];
  const lists = await client.api('/sites/' + siteId + '/lists')
    .filter("displayName eq '" + LIST_CONFIG + "'").get();
  const id = (lists.value && lists.value.length) ? lists.value[0].id : null;
  if (id) _cfgCache[siteId] = id;
  return id;
}

// Acha o item Config 'acessoContratos' (ou null).
async function findAcessoItem(client, siteId, listConfigId) {
  const resp = await client.api('/sites/' + siteId + '/lists/' + listConfigId + '/items')
    .expand('fields').top(50).get();
  return (resp.value || []).find(function (x) { return x.fields && x.fields.Title === ACESSO_TITLE; }) || null;
}

// Le o mapa { diretoria: [emails] }. {} se nao existe / invalido.
async function lerMapaAcessos(client, siteId, listConfigId) {
  try {
    const cfgId = listConfigId || await resolveConfigListId(client, siteId);
    if (!cfgId) return {};
    const item = await findAcessoItem(client, siteId, cfgId);
    if (item && item.fields && item.fields.ConfigJson) {
      const m = JSON.parse(item.fields.ConfigJson);
      if (m && typeof m === 'object' && !Array.isArray(m)) return m;
    }
  } catch (e) { /* sem config -> {} */ }
  return {};
}

// Grava o mapa completo (cria o item se nao existir). Retorna { ok, itemId }.
async function salvarMapaAcessos(client, siteId, listConfigId, mapa) {
  const cfgId = listConfigId || await resolveConfigListId(client, siteId);
  if (!cfgId) throw new Error("Lista '" + LIST_CONFIG + "' nao encontrada");
  const json = JSON.stringify(mapa || {});
  const item = await findAcessoItem(client, siteId, cfgId);
  if (item) {
    await client.api('/sites/' + siteId + '/lists/' + cfgId + '/items/' + item.id + '/fields')
      .patch({ ConfigJson: json });
    return { ok: true, itemId: item.id, action: 'updated' };
  }
  const created = await client.api('/sites/' + siteId + '/lists/' + cfgId + '/items')
    .post({ fields: { Title: ACESSO_TITLE, ConfigJson: json } });
  return { ok: true, itemId: created.id, action: 'created' };
}

// Le a matriz de aprovadores de NF -> { diretoria: Set(emails) } (usada no fallback).
async function lerAprovadoresPorDiretoria(client, siteId, listDirId) {
  const out = {};
  if (!listDirId) return out;
  try {
    const resp = await client.api('/sites/' + siteId + '/lists/' + listDirId + '/items?expand=fields&$top=300').get();
    for (const it of (resp.value || [])) {
      const f = it.fields || {};
      const email = String(f.field_3 || '').toLowerCase().trim();
      const dir = (String(f.Title || '').split('|')[1] || '').trim();
      if (dir && email) {
        if (!out[dir]) out[dir] = new Set();
        out[dir].add(email);
      }
    }
  } catch (e) { /* lista pode nao existir */ }
  return out;
}

// Resolve o conjunto de diretorias que o email acessa, no modelo complementar.
// Retorna array de nomes de diretoria.
async function diretoriasAcessiveis(client, siteId, listDirId, email) {
  const em = String(email || '').toLowerCase().trim();
  if (!em) return [];
  const cfgId = await resolveConfigListId(client, siteId);
  const mapa = await lerMapaAcessos(client, siteId, cfgId);
  const aprov = await lerAprovadoresPorDiretoria(client, siteId, listDirId);
  const dirs = new Set();
  // 1) Diretorias COM config explicita: vale a lista de e-mails configurada.
  for (const d of Object.keys(mapa)) {
    const lista = Array.isArray(mapa[d]) ? mapa[d].map(function (x) { return String(x).toLowerCase().trim(); }) : [];
    if (lista.indexOf(em) >= 0) dirs.add(d);
  }
  // 2) Diretorias SEM config: fallback no aprovador de NF.
  for (const d of Object.keys(aprov)) {
    if (!Array.isArray(mapa[d]) && aprov[d].has(em)) dirs.add(d);
  }
  return Array.from(dirs);
}

module.exports = {
  LIST_CONFIG,
  ACESSO_TITLE,
  resolveConfigListId,
  lerMapaAcessos,
  salvarMapaAcessos,
  lerAprovadoresPorDiretoria,
  diretoriasAcessiveis
};
