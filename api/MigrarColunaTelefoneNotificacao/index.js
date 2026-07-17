/**
 * /api/MigrarColunaTelefoneNotificacao
 *
 * Adiciona coluna `TelefoneNotificacao` (texto) na lista PRONEP-NF-Diretorias.
 * Guarda o telefone E.164 (ex.: +5511999998888) que RECEBE a notificacao de
 * WhatsApp quando o VarrerEmailsNF encontra NFs novas (Fase 2d / n8n).
 * Match por Unidade+Diretoria (a lista e chaveada por Title = "Unidade|Diretoria").
 *
 * Idempotente — se a coluna ja existe, retorna sem fazer nada.
 *
 * RBAC: admin only.
 * Custo Claude: zero.
 */

require('isomorphic-fetch');
const { ClientSecretCredential } = require('@azure/identity');
const { Client } = require('@microsoft/microsoft-graph-client');
const { TokenCredentialAuthenticationProvider } =
  require('@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials');

const LIST_NAME = 'PRONEP-NF-Diretorias';
const COL_NAME = 'TelefoneNotificacao';

async function getGraphClient() {
  const tenantId = process.env.AAD_TENANT_ID;
  const clientId = process.env.AAD_CLIENT_ID;
  const clientSecret = process.env.AAD_CLIENT_SECRET;
  const credential = new ClientSecretCredential(tenantId, clientId, clientSecret);
  const authProvider = new TokenCredentialAuthenticationProvider(credential, {
    scopes: ['https://graph.microsoft.com/.default']
  });
  return Client.initWithMiddleware({ authProvider });
}

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
    if (!user || !user.oid) return false;
    const { getUserRoles } = require('../shared/userRoles');
    const userRoles = await getUserRoles(user);
    return (userRoles || []).includes('administrador');
  } catch (e) { return false; }
}

module.exports = async function (context, req) {
  const diag = { step: 'init', colunaJaExistia: null, colunaCriada: false, erros: [] };
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

    // 1. Achar a lista
    diag.step = 'find_list';
    const lists = await client.api('/sites/' + siteId + '/lists').get();
    const lista = (lists.value || []).find(function(l){ return l.displayName === LIST_NAME; });
    if (!lista) throw new Error('Lista ' + LIST_NAME + ' nao encontrada');
    const listId = lista.id;
    diag.listId = listId;

    // 2. Ver colunas existentes
    diag.step = 'list_columns';
    const cols = await client.api('/sites/' + siteId + '/lists/' + listId + '/columns').get();
    const existentes = new Set();
    for (const c of (cols.value || [])) {
      if (c.displayName) existentes.add(c.displayName);
      if (c.name) existentes.add(c.name);
    }
    diag.colunaJaExistia = existentes.has(COL_NAME);

    // 3. Criar se nao existe (Texto — telefone E.164)
    if (!diag.colunaJaExistia) {
      diag.step = 'create_column';
      try {
        await client.api('/sites/' + siteId + '/lists/' + listId + '/columns').post({
          name: COL_NAME,
          text: {}
        });
        diag.colunaCriada = true;
      } catch (eCol) {
        diag.erros.push({
          step: 'create_column',
          error: eCol.message,
          graphCode: eCol.code,
          graphStatusCode: eCol.statusCode,
          graphBody: eCol.body
        });
      }
    }

    diag.step = 'done';
    context.res = {
      status: 200,
      headers: { 'Content-Type': 'application/json' },
      body: Object.assign({
        ok: !diag.erros.length,
        mensagem: diag.colunaJaExistia
          ? 'Coluna ' + COL_NAME + ' ja existia. Nada a fazer.'
          : (diag.colunaCriada ? 'Coluna ' + COL_NAME + ' CRIADA com sucesso.' : 'Falha ao criar coluna.')
      }, diag)
    };
  } catch (err) {
    context.res = {
      status: 500,
      headers: { 'Content-Type': 'application/json' },
      body: Object.assign({ error: err.message }, diag)
    };
  }
};
