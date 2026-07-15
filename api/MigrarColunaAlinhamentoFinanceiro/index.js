/**
 * /api/MigrarColunaAlinhamentoFinanceiro
 *
 * Adiciona 2 colunas em PRONEP-NF-NotasFiscais para auditar aprovacoes
 * de NFs vencendo em < D+5 dias:
 *   - AlinhouFinanceiro     (boolean Yes/No) - se o aprovador alinhou com Financeiro antes
 *   - GestorFinanceiroAlinhado (text)        - email do gestor financeiro alinhado (se sim)
 *
 * Idempotente: se ja existe, retorna sem fazer nada.
 *
 * RBAC: admin only. Custo Claude: zero.
 */

require('isomorphic-fetch');
const { getGraphClient } = require('../shared/graph');

const LIST_NAME = 'PRONEP-NF-NotasFiscais';
const COLS = [
  { name: 'AlinhouFinanceiro',        def: { boolean: {} } },
  { name: 'GestorFinanceiroAlinhado', def: { text: {} } }
];

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
  const diag = { step: 'init', criadas: [], jaExistia: [], falhas: [] };
  try {
    if (!(await isAdmin(req))) {
      context.res = { status: 403, body: { error: 'Apenas admin' } };
      return;
    }
    const client = await getGraphClient();
    const host = process.env.SHAREPOINT_SITE_HOSTNAME;
    const path = process.env.SHAREPOINT_SITE_PATH;
    const siteResp = await client.api('/sites/' + host + ':' + path).get();
    const siteId = siteResp.id;
    const lists = await client.api('/sites/' + siteId + '/lists').get();
    const lista = (lists.value || []).find(l => l.displayName === LIST_NAME);
    if (!lista) throw new Error('Lista ' + LIST_NAME + ' nao encontrada');
    const listId = lista.id;

    const colsResp = await client.api('/sites/' + siteId + '/lists/' + listId + '/columns').get();
    const existentes = new Set();
    for (const c of (colsResp.value || [])) {
      if (c.displayName) existentes.add(c.displayName);
      if (c.name) existentes.add(c.name);
    }

    for (const col of COLS) {
      if (existentes.has(col.name)) { diag.jaExistia.push(col.name); continue; }
      try {
        await client.api('/sites/' + siteId + '/lists/' + listId + '/columns')
          .post(Object.assign({ name: col.name }, col.def));
        diag.criadas.push(col.name);
      } catch (eCol) {
        // Tenta fallback text se boolean nao aceito
        if (col.def.boolean) {
          try {
            await client.api('/sites/' + siteId + '/lists/' + listId + '/columns')
              .post({ name: col.name, text: {} });
            diag.criadas.push(col.name + ' (fallback text)');
            continue;
          } catch (eText) { /* falha geral abaixo */ }
        }
        diag.falhas.push({ coluna: col.name, error: eCol.message, code: eCol.code, body: eCol.body });
      }
    }

    diag.step = 'done';
    context.res = {
      status: 200,
      headers: { 'Content-Type': 'application/json' },
      body: Object.assign({ ok: diag.falhas.length === 0 }, diag)
    };
  } catch (err) {
    context.res = { status: 500, body: Object.assign({ error: err.message }, diag) };
  }
};
