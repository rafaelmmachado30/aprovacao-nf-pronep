/**
 * shared/acessoTelas.js — Central de Controle de Acessos a TELAS do sistema.
 *
 * Mesmo padrao do shared/acessoContratos.js, porem o "alvo" e uma TELA (view) do app
 * em vez de uma pasta de contrato. Cada tela pode liberar acesso a um ou mais GRUPOS
 * de acesso do Entra (roles gestor_* e financeiro_nf) e/ou a PESSOAS especificas (email).
 *
 * Mapa salvo num item da lista PRONEP-NF-Config (Title='acessoTelas', ConfigJson):
 *   { "aprovadas": ["gestor_fiscal_contabil","fulano@pronep.com.br"], "fila-aprovacao": [...] }
 *   chave = id da tela (mesmo id usado no canSee do front);  valor = tokens (role OU email).
 *
 * Modelo ADITIVO (decisao do Rafa): este controle SO ADICIONA acesso. Nunca remove o
 * que o papel ja enxerga hoje. Admin/Juridico e as regras atuais continuam valendo.
 *   - Tela COM tokens -> quem for membro de um grupo liberado, ou estiver na lista de
 *     pessoas, passa a ver a tela (alem de quem ja via pelo papel).
 *   - Tela SEM tokens -> nada muda (segue so o canSee do papel).
 *
 * IMPORTANTE: TELAS é a lista canonica de telas configuraveis. Sao apenas telas de
 * OPERACAO. Telas administrativas criticas (controle-acessos, configuracoes, auditoria)
 * ficam de fora de proposito — nao sao liberaveis por aqui.
 */

require('isomorphic-fetch');

const LIST_CONFIG = 'PRONEP-NF-Config';
const ACESSO_TITLE = 'acessoTelas';

// Telas configuraveis pela Central. id = mesmo id de view usado no front (canSee/menuItems).
const TELAS = [
  { id: 'dashboard',          label: 'Dashboard' },
  { id: 'fila-aprovacao',     label: 'Fila de Aprovação' },
  { id: 'aprovadas',          label: 'Notas Aprovadas' },
  { id: 'rejeitadas',         label: 'Notas Rejeitadas' },
  { id: 'nova-nf',            label: 'Lançamento de Notas Fiscais' },
  { id: 'minhas-nfs',         label: 'Minhas NFs' },
  { id: 'caixa-entrada',      label: 'Caixa de Entrada (PDFs)' },
  { id: 'contas-recorrentes', label: 'Contas Recorrentes' },
  { id: 'fechamento-mes',     label: 'Fechamento do Mês' },
  { id: 'fornecedores',       label: 'Fornecedores' },
  { id: 'contratos',          label: 'Contratos' },
  { id: 'mapa-aprovadores',   label: 'Mapa de Aprovadores' }
];
const TELAS_IDS = TELAS.map(function (t) { return t.id; });

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

// Le o mapa { tela: [tokens] }. {} se nao existe / invalido.
async function lerMapaTelas(client, siteId, listConfigId) {
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

async function salvarMapaTelas(client, siteId, listConfigId, mapa) {
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

// Decide se um usuario ganhou acesso a uma tela via Central (ADITIVO).
// Tokens em mapa[tela] podem ser ROLE (grupo inteiro) OU EMAIL (pessoa especifica).
function podeVerTela(viewId, userEmail, userRoles, mapa) {
  const v = String(viewId || '').trim();
  if (!v) return false;
  const liberadas = (mapa && Array.isArray(mapa[v])) ? mapa[v] : null;
  if (!liberadas || !liberadas.length) return false;
  const email = String(userEmail || '').toLowerCase().trim();
  const roles = (userRoles || []).map(function (r) { return String(r).toLowerCase(); });
  const set = liberadas.map(function (t) { return String(t).toLowerCase().trim(); });
  if (email && set.indexOf(email) >= 0) return true;             // pessoa especifica liberada
  return roles.some(function (r) { return set.indexOf(r) >= 0; }); // grupo inteiro liberado
}

// Lista de viewIds que o usuario ganhou via Central (so telas configuraveis validas).
function telasLiberadasPara(userEmail, userRoles, mapa) {
  const out = [];
  for (const v of TELAS_IDS) {
    if (podeVerTela(v, userEmail, userRoles, mapa)) out.push(v);
  }
  return out;
}

module.exports = {
  LIST_CONFIG,
  ACESSO_TITLE,
  TELAS,
  TELAS_IDS,
  resolveConfigListId,
  lerMapaTelas,
  salvarMapaTelas,
  podeVerTela,
  telasLiberadasPara
};
