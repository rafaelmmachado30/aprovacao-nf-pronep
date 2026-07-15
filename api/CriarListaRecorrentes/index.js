/**
 * /api/CriarListaRecorrentes
 *
 * Cria (ou valida) a lista PRONEP-NF-Recorrentes e suas colunas.
 * NAO chama Claude. Operacao 100% sobre SharePoint. RBAC: admin only.
 *
 * Mesmo padrao de CriarListaContratos: idempotente, cria colunas que faltam
 * uma a uma, retorna diagnostico por etapa.
 */

require('isomorphic-fetch');
const { getGraphClient } = require('../shared/graph');
const { LIST_RECORRENTES, COLUNAS_RECORRENTES } = require('../shared/recorrentes');

function readClientPrincipal(req) {
  const header = req.headers && req.headers['x-ms-client-principal'];
  if (!header) return null;
  try { return JSON.parse(Buffer.from(header, 'base64').toString('utf-8')); } catch (e) { return null; }
}

async function isAdmin(req) {
  const p = readClientPrincipal(req);
  const roles = (p && p.userRoles) || [];
  if (roles.includes('administrador') || roles.includes('admin')) return true;
  try {
    const { getUser } = require('../shared/auth');
    const user = await getUser(req);
    if (!user) return false;
    const { isAdminEmail } = require('../shared/authz');
    if (isAdminEmail((user.email || '').toLowerCase())) return true;
    const { getUserRoles } = require('../shared/userRoles');
    const userRoles = await getUserRoles(user);
    return (userRoles || []).includes('administrador');
  } catch (e) {
    return false;
  }
}

module.exports = async function (context, req) {
  const diag = {
    step: 'init', listName: LIST_RECORRENTES, listaJaExistia: null, listaCriada: false,
    listIdAposGarantir: null, colunasCriadas: [], colunasComFalha: [], erros: [], timeMs: 0
  };
  const t0 = Date.now();
  try {
    if (!(await isAdmin(req))) {
      context.res = { status: 403, headers: { 'Content-Type': 'application/json' }, body: { error: 'Apenas admin' } };
      return;
    }
    const client = await getGraphClient();

    diag.step = 'resolve_site';
    const host = process.env.SHAREPOINT_SITE_HOSTNAME;
    const path = process.env.SHAREPOINT_SITE_PATH;
    if (!host || !path) throw new Error('SHAREPOINT_SITE_HOSTNAME/PATH nao configurados');
    const siteResp = await client.api('/sites/' + host + ':' + path).get();
    const siteId = siteResp.id;
    diag.site = { host, path, siteId, webUrl: siteResp.webUrl };

    diag.step = 'search_list';
    let listId = null;
    try {
      const lists = await client.api('/sites/' + siteId + '/lists')
        .filter("displayName eq '" + LIST_RECORRENTES + "'").get();
      if (lists.value && lists.value.length) { listId = lists.value[0].id; diag.listaJaExistia = true; }
      else diag.listaJaExistia = false;
    } catch (eSearch) {
      diag.erros.push({ step: 'search_list_filter', error: eSearch.message });
      const allLists = await client.api('/sites/' + siteId + '/lists').get();
      const found = (allLists.value || []).find(function (l) { return l.displayName === LIST_RECORRENTES; });
      if (found) { listId = found.id; diag.listaJaExistia = true; } else diag.listaJaExistia = false;
    }

    if (!listId) {
      diag.step = 'create_list';
      try {
        const newList = await client.api('/sites/' + siteId + '/lists').post({
          displayName: LIST_RECORRENTES, list: { template: 'genericList' }
        });
        listId = newList.id; diag.listaCriada = true;
      } catch (eCreate) {
        diag.erros.push({ step: 'create_list', error: eCreate.message, graphBody: eCreate.body });
        diag.timeMs = Date.now() - t0;
        context.res = { status: 500, headers: { 'Content-Type': 'application/json' },
          body: Object.assign({ error: 'Falha ao criar lista: ' + eCreate.message }, diag) };
        return;
      }
    }
    diag.listIdAposGarantir = listId;

    diag.step = 'list_columns';
    const colsResp = await client.api('/sites/' + siteId + '/lists/' + listId + '/columns').get();
    const colsExistentes = new Set();
    for (const c of (colsResp.value || [])) {
      if (c.displayName) colsExistentes.add(c.displayName);
      if (c.name) colsExistentes.add(c.name);
    }

    diag.step = 'create_missing_columns';
    for (const col of COLUNAS_RECORRENTES) {
      if (colsExistentes.has(col.name)) continue;
      try {
        const payload = Object.assign({ name: col.name }, col.def);
        await client.api('/sites/' + siteId + '/lists/' + listId + '/columns').post(payload);
        diag.colunasCriadas.push(col.name);
      } catch (eCol) {
        diag.colunasComFalha.push({ coluna: col.name, error: eCol.message, graphBody: eCol.body });
      }
    }

    diag.step = 'done';
    diag.timeMs = Date.now() - t0;
    context.res = {
      status: 200, headers: { 'Content-Type': 'application/json' },
      body: Object.assign({
        ok: true,
        mensagem: diag.listaCriada
          ? 'Lista ' + LIST_RECORRENTES + ' CRIADA. Colunas: ' + diag.colunasCriadas.length
          : 'Lista ' + LIST_RECORRENTES + ' ja existia. Colunas novas: ' + diag.colunasCriadas.length + ' (' + diag.colunasComFalha.length + ' falharam)'
      }, diag)
    };
  } catch (err) {
    diag.timeMs = Date.now() - t0;
    context.log && context.log.error && context.log.error('CriarListaRecorrentes:', err);
    context.res = { status: 500, headers: { 'Content-Type': 'application/json' },
      body: Object.assign({ error: err.message }, diag) };
  }
};
