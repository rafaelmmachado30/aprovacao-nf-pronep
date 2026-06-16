/**
 * /api/AbrirContrato?id={spContratoId}
 *
 * Recebe o ID do item da lista PRONEP-NF-Contratos, busca o DriveItemId
 * gravado, e:
 *   - Se o cliente eh PDF: retorna o conteudo (stream) inline
 *   - Se for DOCX: redireciona pro webUrl do SharePoint (abre no Office Online)
 *
 * RBAC: identico ao ListarContratos
 *   - admin: tudo
 *   - gestor: so contratos da diretoria dele
 *   - outros: 403
 */

const { ClientSecretCredential } = require('@azure/identity');
const { Client } = require('@microsoft/microsoft-graph-client');
const { TokenCredentialAuthenticationProvider } = require('@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials');
const { getUser } = require('../shared/auth');
const { getUserRoles } = require('../shared/userRoles');
const { podeVerContrato, lerMapaAcessos } = require('../shared/acessoContratos');

const cache = { siteId: null, listId: null, listDirId: null, colMap: null };

function getGraphClient() {
  const credential = new ClientSecretCredential(
    process.env.AAD_TENANT_ID,
    process.env.AAD_CLIENT_ID,
    process.env.AAD_CLIENT_SECRET
  );
  const authProvider = new TokenCredentialAuthenticationProvider(credential, {
    scopes: ['https://graph.microsoft.com/.default']
  });
  return Client.initWithMiddleware({ authProvider });
}

async function resolveSiteELists(client) {
  if (cache.siteId && cache.listId) return cache;
  const siteResp = await client.api('/sites/' + process.env.SHAREPOINT_SITE_HOSTNAME + ':' + process.env.SHAREPOINT_SITE_PATH).get();
  cache.siteId = siteResp.id;
  const lists = await client.api('/sites/' + cache.siteId + '/lists').get();
  for (const l of (lists.value || [])) {
    if (l.displayName === 'PRONEP-NF-Contratos') cache.listId = l.id;
    if (l.displayName === 'PRONEP-NF-Diretorias') cache.listDirId = l.id;
  }
  return cache;
}

async function resolveContratosSite(client) {
  const host = process.env.SHAREPOINT_CONTRATOS_HOSTNAME || 'pronepadmin.sharepoint.com';
  const path = process.env.SHAREPOINT_CONTRATOS_PATH || '/sites/CONTRATOS-SERVICOS-CONTRATOS';
  const siteResp = await client.api('/sites/' + host + ':' + path).get();
  const driveResp = await client.api('/sites/' + siteResp.id + '/drive').get();
  return { siteId: siteResp.id, driveId: driveResp.id };
}

function readClientPrincipal(req) {
  const header = req.headers && req.headers['x-ms-client-principal'];
  if (!header) return null;
  try { return JSON.parse(Buffer.from(header, 'base64').toString('utf-8')); } catch (e) { return null; }
}
function readClientPrincipalRoles(req) {
  const p = readClientPrincipal(req);
  return (p && p.userRoles) || [];
}

async function diretoriasDoGestor(client, siteId, listDirId, userEmail) {
  if (!listDirId || !userEmail) return [];
  const resp = await client.api('/sites/' + siteId + '/lists/' + listDirId + '/items?expand=fields&$top=200').get();
  const set = new Set();
  for (const it of (resp.value || [])) {
    const f = it.fields || {};
    const emailDir = String(f.field_3 || '').toLowerCase().trim();
    if (emailDir === userEmail) {
      const dir = String(f.Title || '').split('|')[1] || '';
      if (dir) set.add(dir.trim());
    }
  }
  return Array.from(set);
}

module.exports = async function (context, req) {
  try {
    const contratoId = (req.query && req.query.id) || '';
    if (!contratoId) {
      context.res = { status: 400, body: { error: 'id obrigatorio' } };
      return;
    }

    // 1. Auth
    const user = await getUser(req);
    if (!user || !user.email) {
      context.res = { status: 401, body: { error: 'Nao autenticado' } };
      return;
    }

    // 2. RBAC
    const claimsRoles = readClientPrincipalRoles(req) || [];
    const userRoles = await getUserRoles(user);
    const allRoles = Array.from(new Set([].concat(claimsRoles, userRoles || [])));
    const isAdmin = allRoles.includes('administrador') || allRoles.includes('admin');
    const veTodos = isAdmin || allRoles.includes('gestor_juridica'); // admin/juridico veem tudo
    // Sem 403 fixo aqui — acesso e por GRUPO (Controle de Acessos), conferido por-contrato abaixo.

    // 3. Busca contrato no SP
    const client = getGraphClient();
    const { siteId, listId, listDirId } = await resolveSiteELists(client);
    if (!listId) {
      context.res = { status: 404, body: { error: 'Lista de contratos nao existe' } };
      return;
    }
    const item = await client.api('/sites/' + siteId + '/lists/' + listId + '/items/' + contratoId + '?expand=fields').get();
    const f = item.fields || {};
    const diretoria = f.Diretoria || '';
    const driveItemId = f.DriveItemId || '';
    const caminhoSP = f.CaminhoSharepoint || '';

    // 4. Confere acesso por GRUPO (mesmo criterio do ListarContratos: Controle de Acessos
    // + fallback no grupo de mesmo nome da pasta).
    if (!veTodos) {
      const mapa = await lerMapaAcessos(client, siteId, null);
      if (!podeVerContrato(diretoria, allRoles, mapa)) {
        context.res = { status: 403, body: { error: 'Voce nao tem acesso a esta pasta de contratos. Solicite liberacao ao Admin (Controle de Acessos).' } };
        return;
      }
    }

    // 5. Serve o arquivo PELO SISTEMA (identidade do app, que tem Sites.Read.All), pra
    //    NAO depender da permissao do usuario no SharePoint — o acesso e 100% controlado
    //    pelo Controle de Acessos. Sem driveItemId (sync antigo), cai no redirect (legado).
    if (!driveItemId) {
      if (caminhoSP) { context.res = { status: 302, headers: { Location: caminhoSP } }; return; }
      context.res = { status: 404, body: { error: 'Registro sem DriveItemId nem CaminhoSharepoint' } };
      return;
    }

    const { driveId } = await resolveContratosSite(client);
    const meta = await client.api('/drives/' + driveId + '/items/' + driveItemId).get();
    const dlUrl = meta && meta['@microsoft.graph.downloadUrl'];
    const nomeArq = (meta && meta.name) || f.NomeArquivo || 'contrato';
    const mime = (meta && meta.file && meta.file.mimeType) || 'application/octet-stream';

    if (!dlUrl) {
      // Sem downloadUrl: ultimo fallback no webUrl (requer permissao SP).
      if (meta && meta.webUrl) { context.res = { status: 302, headers: { Location: meta.webUrl } }; return; }
      context.res = { status: 404, body: { error: 'Nao foi possivel obter o conteudo do arquivo' } };
      return;
    }

    require('isomorphic-fetch');
    const dlResp = await fetch(dlUrl);
    if (!dlResp.ok) {
      context.res = { status: 502, body: { error: 'Falha ao baixar o arquivo do SharePoint (HTTP ' + dlResp.status + ')' } };
      return;
    }
    const buf = Buffer.from(await dlResp.arrayBuffer());
    context.res = {
      status: 200,
      headers: {
        'Content-Type': mime,
        'Content-Disposition': 'inline; filename="' + encodeURIComponent(nomeArq) + '"',
        'Cache-Control': 'private, no-store'
      },
      isRaw: true,
      body: buf
    };
  } catch (err) {
    context.log && context.log.error && context.log.error('AbrirContrato error:', err);
    context.res = {
      status: 500,
      headers: { 'Content-Type': 'application/json' },
      body: { error: err.message }
    };
  }
};
