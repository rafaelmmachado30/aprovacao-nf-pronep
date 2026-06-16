/**
 * shared/acessoContratos.js — Controle de Acessos a contratos (por GRUPO/diretoria).
 *
 * Modelo (decisao do Rafa): cada PASTA de contrato (ex: "Comercial", "Tecnologia") libera
 * acesso a um ou mais GRUPOS de acesso do Entra (= diretorias: Suprimentos, Tecnologia,
 * Juridica, Financeira, ... e Financeiro-Gestao). Quem for MEMBRO de um grupo liberado ve a
 * pasta. As pessoas vem automaticamente da pertinencia ao grupo (Graph), nao de e-mails fixos.
 *
 * Mapa salvo num item da lista PRONEP-NF-Config (Title='acessoContratos', ConfigJson):
 *   { "Comercial": ["gestor_juridica","gestor_financeira"], "Tecnologia": [...], ... }
 *   chave = pasta (diretoria do contrato);  valor = ROLES (grupos) com acesso.
 *
 * Modelo COMPLEMENTAR:
 *   - Pasta COM config -> so os grupos listados (via seus membros) tem acesso.
 *   - Pasta SEM config -> fallback: o grupo cujo nome = nome da pasta (ex: pasta "Tecnologia"
 *     -> grupo gestor_tecnologia). Pasta sem grupo correspondente: so admin/juridico.
 *   - Admin e Juridico veem tudo (tratado fora deste helper).
 */

require('isomorphic-fetch');
const { ROLE_LABELS } = require('./userRoles');

const LIST_CONFIG = 'PRONEP-NF-Config';
const ACESSO_TITLE = 'acessoContratos';

const _cfgCache = {};
async function resolveConfigListId(client, siteId) {
  if (_cfgCache[siteId]) return _cfgCache[siteId];
  const lists = await client.api('/sites/' + siteId + '/lists')
    .filter("displayName eq '" + LIST_CONFIG + "'").get();
  const id = (lists.value && lists.value.length) ? lists.value[0].id : null;
  if (id) _cfgCache[siteId] = id;
  return id;
}

async function findAcessoItem(client, siteId, listConfigId) {
  const resp = await client.api('/sites/' + siteId + '/lists/' + listConfigId + '/items')
    .expand('fields').top(50).get();
  return (resp.value || []).find(function (x) { return x.fields && x.fields.Title === ACESSO_TITLE; }) || null;
}

// Le o mapa { pasta: [roles] }. {} se nao existe / invalido.
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

// Normaliza pra comparacao: minusculas, sem acento, so alfanumerico.
function _norm(s) {
  return String(s || '').toLowerCase().normalize('NFD').replace(/\p{Diacritic}/gu, '').replace(/[^a-z0-9]/g, '');
}

// Pasta -> role do grupo de mesmo nome (ex: "Tecnologia" -> gestor_tecnologia). null se nenhum.
function folderParaRole(folder) {
  const f = _norm(folder);
  for (const role of Object.keys(ROLE_LABELS)) {
    if (_norm(ROLE_LABELS[role]) === f) return role;
  }
  return null;
}

// Decide se um usuario (com seus roles/grupos) ve os contratos de uma pasta. Complementar.
function podeVerContrato(folderDiretoria, userRoles, mapa) {
  const folder = String(folderDiretoria || '').trim();
  if (!folder) return false;
  const roles = (userRoles || []).map(function (r) { return String(r).toLowerCase(); });
  // Config explicita da pasta (chave case/acento-insensitive)?
  let liberadas = null;
  for (const k of Object.keys(mapa || {})) {
    if (_norm(k) === _norm(folder) && Array.isArray(mapa[k])) { liberadas = mapa[k]; break; }
  }
  if (liberadas) {
    const set = liberadas.map(function (r) { return String(r).toLowerCase(); });
    return roles.some(function (r) { return set.indexOf(r) >= 0; });
  }
  // Fallback: grupo cujo nome = nome da pasta.
  const ownRole = folderParaRole(folder);
  return ownRole ? roles.indexOf(ownRole.toLowerCase()) >= 0 : false;
}

module.exports = {
  LIST_CONFIG,
  ACESSO_TITLE,
  resolveConfigListId,
  lerMapaAcessos,
  salvarMapaAcessos,
  folderParaRole,
  podeVerContrato
};
