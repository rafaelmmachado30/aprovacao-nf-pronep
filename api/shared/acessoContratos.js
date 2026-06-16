/**
 * shared/acessoContratos.js — Controle de Acessos a contratos (por DIRETORIA).
 *
 * Modelo (decisao do Rafa): cada PASTA de contrato (uma diretoria, ex: "Tecnologia")
 * pode ser acessada por uma ou mais DIRETORIAS. Quem for gestor/aprovador de qualquer
 * diretoria liberada ve a pasta — as PESSOAS vem automaticamente da matriz de gestores
 * (PRONEP-NF-Diretorias). Assim, trocar o gestor de uma diretoria ajusta o acesso sozinho.
 *
 * Mapa salvo num item da lista PRONEP-NF-Config (Title='acessoContratos', ConfigJson):
 *   { "Tecnologia": ["Tecnologia","Comercial"], "Suprimentos": [...], ... }
 *   chave  = pasta (diretoria do contrato);  valor = diretorias com acesso.
 *
 * Modelo COMPLEMENTAR:
 *   - Pasta COM config -> so as diretorias listadas tem acesso (via seus gestores).
 *   - Pasta SEM config -> fallback: a propria diretoria da pasta (aprovador de NF dela).
 *   - Admin e Juridico veem tudo (tratado fora deste helper).
 */

require('isomorphic-fetch');

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

// Le o mapa { pasta: [diretorias] }. {} se nao existe / invalido.
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

// Grava o mapa completo (cria o item se nao existir).
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

// Diretorias em que o e-mail e gestor/aprovador (matriz PRONEP-NF-Diretorias).
// Title formato "Unidade|Diretoria"; field_3 = e-mail do aprovador.
async function dirsDoUsuario(client, siteId, listDirId, email) {
  const out = new Set();
  const em = String(email || '').toLowerCase().trim();
  if (!listDirId || !em) return [];
  try {
    const resp = await client.api('/sites/' + siteId + '/lists/' + listDirId + '/items?expand=fields&$top=300').get();
    for (const it of (resp.value || [])) {
      const f = it.fields || {};
      if (String(f.field_3 || '').toLowerCase().trim() === em) {
        const dir = (String(f.Title || '').split('|')[1] || '').trim();
        if (dir) out.add(dir);
      }
    }
  } catch (e) { /* lista pode nao existir */ }
  return Array.from(out);
}

// Normaliza nome de diretoria pra comparacao (lower + trim).
function _norm(s) { return String(s || '').toLowerCase().trim(); }

// Decide se um usuario (com suas diretorias dirsUsuario) pode ver os contratos de uma
// pasta (folderDiretoria), dado o mapa de acessos. Modelo complementar.
function podeVerContrato(folderDiretoria, dirsUsuario, mapa) {
  const folder = String(folderDiretoria || '').trim();
  if (!folder) return false;
  const dirsU = (dirsUsuario || []).map(_norm);
  // Acha a config da pasta (case-insensitive na chave)
  let liberadas = null;
  for (const k of Object.keys(mapa || {})) {
    if (_norm(k) === _norm(folder) && Array.isArray(mapa[k])) { liberadas = mapa[k]; break; }
  }
  if (liberadas) {
    // Pasta configurada: usuario ve se e gestor de alguma diretoria liberada.
    const set = liberadas.map(_norm);
    return dirsU.some(function (d) { return set.indexOf(d) >= 0; });
  }
  // Pasta sem config: fallback — so a propria diretoria da pasta (aprovador de NF dela).
  return dirsU.indexOf(_norm(folder)) >= 0;
}

module.exports = {
  LIST_CONFIG,
  ACESSO_TITLE,
  resolveConfigListId,
  lerMapaAcessos,
  salvarMapaAcessos,
  dirsDoUsuario,
  podeVerContrato
};
