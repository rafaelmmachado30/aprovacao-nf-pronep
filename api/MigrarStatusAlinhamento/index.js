/**
 * /api/MigrarStatusAlinhamento
 *
 * Adiciona a opcao "AguardandoAlinhamento" na coluna Status (choice) da lista
 * PRONEP-NF-NotasFiscais. Necessario pro gate de alinhamento (NF <D+5 que o
 * aprovador declarou "alinhei com o financeiro" fica nesse estado ate o
 * Financeiro-Gestao confirmar). Idempotente. RBAC: admin only.
 */

require('isomorphic-fetch');
const { getGraphClient } = require('../shared/graph');

const LIST_NAME = 'PRONEP-NF-NotasFiscais';
const NOVO_STATUS = 'AguardandoAlinhamento';

function readClientPrincipal(req) {
  const h = req.headers && req.headers['x-ms-client-principal'];
  if (!h) return null;
  try { return JSON.parse(Buffer.from(h, 'base64').toString('utf-8')); } catch (e) { return null; }
}
async function isAdmin(req) {
  const p = readClientPrincipal(req);
  const roles = (p && p.userRoles) || [];
  if (roles.includes('administrador') || roles.includes('admin')) return true;
  try {
    const { getUser } = require('../shared/auth');
    const user = await getUser(req);
    if (!user || !user.oid) return false;
    const { getUserRoles } = require('../shared/userRoles');
    const userRoles = await getUserRoles(user);
    return (userRoles || []).includes('administrador');
  } catch (e) { return false; }
}

module.exports = async function (context, req) {
  const diag = { step: 'init' };
  try {
    if (!(await isAdmin(req))) { context.res = { status: 403, body: { error: 'Apenas admin' } }; return; }
    const client = await getGraphClient();
    const host = process.env.SHAREPOINT_SITE_HOSTNAME;
    const path = process.env.SHAREPOINT_SITE_PATH;
    const siteResp = await client.api('/sites/' + host + ':' + path).get();
    const siteId = siteResp.id;
    const lists = await client.api('/sites/' + siteId + '/lists').get();
    const lista = (lists.value || []).find(function (l) { return l.displayName === LIST_NAME; });
    if (!lista) throw new Error('Lista ' + LIST_NAME + ' nao encontrada');
    const listId = lista.id;

    diag.step = 'find_status_column';
    const colsResp = await client.api('/sites/' + siteId + '/lists/' + listId + '/columns').get();
    const statusCol = (colsResp.value || []).find(function (c) {
      return (c.displayName === 'Status' || c.name === 'Status') && c.choice;
    });
    if (!statusCol) throw new Error('Coluna Status (choice) nao encontrada');
    const choices = (statusCol.choice && statusCol.choice.choices) || [];
    diag.choicesAntes = choices.slice();

    if (choices.indexOf(NOVO_STATUS) >= 0) {
      diag.step = 'done';
      context.res = { status: 200, headers: { 'Content-Type': 'application/json' },
        body: { ok: true, jaExistia: true, choices: choices } };
      return;
    }

    diag.step = 'patch_choices';
    const novas = choices.concat([NOVO_STATUS]);
    await client.api('/sites/' + siteId + '/lists/' + listId + '/columns/' + statusCol.id)
      .patch({ choice: { choices: novas } });

    diag.step = 'done';
    context.res = { status: 200, headers: { 'Content-Type': 'application/json' },
      body: { ok: true, criado: NOVO_STATUS, choices: novas } };
  } catch (err) {
    context.res = { status: 500, headers: { 'Content-Type': 'application/json' },
      body: { error: (err && err.message) || String(err), diag: diag } };
  }
};
