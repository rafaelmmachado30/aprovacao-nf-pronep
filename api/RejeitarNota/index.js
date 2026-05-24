/**
 * Sistema de Aprovacao de NF - RejeitarNota
 *
 * POST /api/RejeitarNota
 * Body JSON: { id, motivo, observacao }
 *
 * Fluxo:
 *  1. Valida role do usuario (aprovador atual ou admin)
 *  2. Le item da lista NotasFiscais
 *  3. Move o PDF de /Pendentes/{Unidade}/Diretoria {Diretoria}/ pra
 *     /Notas Fiscais/Rejeitadas/{Unidade}/Diretoria {Diretoria}/
 *  4. Atualiza Status=Rejeitada, MotivoRejeicao=<motivo + observacao>, AprovadoEm=<now>
 */

require('isomorphic-fetch');
const { ClientSecretCredential } = require('@azure/identity');
const { Client } = require('@microsoft/microsoft-graph-client');
const { TokenCredentialAuthenticationProvider } =
  require('@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials');

const LIST_NOTAS = 'PRONEP-NF-NotasFiscais';
const cache = { siteId: null, driveId: null, listNotasId: null, colMap: null, invColMap: null, colTypes: null, colMapCachedAt: 0 };
const CACHE_TTL_MS = 5 * 60 * 1000;

async function getGraphClient() {
  const tenantId = process.env.AAD_TENANT_ID;
  const clientId = process.env.AAD_CLIENT_ID;
  const clientSecret = process.env.AAD_CLIENT_SECRET;
  if (!tenantId || !clientId || !clientSecret) throw new Error('AAD_* incompletas');
  const credential = new ClientSecretCredential(tenantId, clientId, clientSecret);
  const authProvider = new TokenCredentialAuthenticationProvider(credential, {
    scopes: ['https://graph.microsoft.com/.default']
  });
  return Client.initWithMiddleware({ authProvider });
}

async function resolveSiteAndList(client) {
  if (cache.siteId && cache.driveId && cache.listNotasId) return cache;
  const host = process.env.SHAREPOINT_SITE_HOSTNAME;
  const path = process.env.SHAREPOINT_SITE_PATH;
  if (!host || !path) throw new Error('SHAREPOINT_* incompletas');
  const siteResp = await client.api(`/sites/${host}:${path}`).get();
  cache.siteId = siteResp.id;
  const driveResp = await client.api(`/sites/${cache.siteId}/drive`).get();
  cache.driveId = driveResp.id;
  const lists = await client.api(`/sites/${cache.siteId}/lists`).filter(`displayName eq '${LIST_NOTAS}'`).get();
  if (!lists.value || !lists.value.length) throw new Error(`Lista ${LIST_NOTAS} nao encontrada`);
  cache.listNotasId = lists.value[0].id;
  return cache;
}

async function getColumnMap(client, siteId, listId) {
  const age = Date.now() - (cache.colMapCachedAt || 0);
  if (cache.colMap && age < CACHE_TTL_MS) return cache.colMap;
  const resp = await client.api(`/sites/${siteId}/lists/${listId}/columns`).get();
  const map = {}; const types = {};
  for (const col of (resp.value || [])) {
    if (!col.displayName || !col.name) continue;
    if (col.readOnly === true) continue;
    if (col.hidden === true) continue;
    if (col.name.startsWith('_')) continue;
    if (['LinkTitle','LinkTitleNoMenu','Edit','DocIcon','ItemChildCount',
         'FolderChildCount','AppAuthor','AppEditor','Attachments'].includes(col.name)) continue;
    map[col.displayName] = col.name;
    let t = 'text';
    if (col.text) t='text'; else if (col.number) t='number';
    else if (col.dateTime) t='dateTime'; else if (col.boolean) t='boolean';
    else if (col.choice) t='choice'; else if (col.hyperlinkOrPicture) t='hyperlink';
    else if (col.personOrGroup) t='person'; else if (col.currency) t='currency';
    types[col.name] = t;
  }
  cache.colMap = map; cache.colTypes = types; cache.colMapCachedAt = Date.now();
  cache.invColMap = {};
  for (const [k, v] of Object.entries(map)) cache.invColMap[v] = k;
  return map;
}

function readClientPrincipal(req) {
  const header = req.headers && req.headers['x-ms-client-principal'];
  if (!header) return null;
  try { return JSON.parse(Buffer.from(header, 'base64').toString('utf-8')); }
  catch (e) { return null; }
}

function normalizeFields(fields, invColMap) {
  const out = {};
  for (const [k, v] of Object.entries(fields)) {
    if (invColMap[k]) out[invColMap[k]] = v;
  }
  return out;
}

function buildPatchPayload(displayPayload, colMap, colTypes) {
  const fields = {};
  for (const [displayName, value] of Object.entries(displayPayload)) {
    const internal = colMap[displayName];
    if (!internal) continue;
    if (value === null || value === undefined) continue;
    const t = colTypes[internal] || 'text';
    let formatted = value;
    if (t === 'hyperlink') formatted = { Url: String(value), Description: '' };
    else if (t === 'number' || t === 'currency') formatted = Number(value);
    else if (t === 'boolean') formatted = (value === true || value === 'Sim');
    else if (t === 'dateTime') formatted = value;
    else formatted = String(value);
    fields[internal] = formatted;
  }
  return fields;
}

module.exports = async function (context, req) {
  const diag = { step: 'start' };
  try {
    diag.step = 'parse_body';
    const body = req.body || {};
    const itemId = body.id;
    const motivo = body.motivo || '';
    const observacao = body.observacao || '';
    if (!itemId)  return errResp(context, 400, 'id obrigatorio');
    if (!motivo)  return errResp(context, 400, 'motivo obrigatorio');
    diag.itemId = itemId;

    diag.step = 'principal';
    const principal = readClientPrincipal(req);
    if (!principal) return errResp(context, 401, 'Nao autenticado');
    const aprovadorEmail = (principal.userDetails || '').toLowerCase();

    diag.step = 'graph';
    const client = await getGraphClient();
    const { siteId, driveId, listNotasId } = await resolveSiteAndList(client);

    diag.step = 'columns';
    await getColumnMap(client, siteId, listNotasId);
    const { colMap, invColMap, colTypes } = cache;

    diag.step = 'fetch_item';
    const item = await client.api(`/sites/${siteId}/lists/${listNotasId}/items/${itemId}?expand=fields`).get();
    const f = normalizeFields(item.fields || {}, invColMap);
    diag.item = { Title: f.Title, Status: f.Status, Unidade: f.Unidade, Diretoria: f.Diretoria, AprovadorAtual: f.AprovadorAtual };

    // RBAC
    const isAdmin = aprovadorEmail === 'rafael.machado@pronep.com.br';
    const isAprovadorAtribuido = (f.AprovadorAtual || '').toLowerCase() === aprovadorEmail;
    if (!isAdmin && !isAprovadorAtribuido) {
      return errResp(context, 403, 'Voce nao eh o aprovador desta NF', { aprovadorAtual: f.AprovadorAtual, voce: aprovadorEmail });
    }
    if (f.Status === 'Aprovada')  return errResp(context, 409, 'NF ja foi aprovada, nao pode ser rejeitada');
    if (f.Status === 'Rejeitada') return errResp(context, 409, 'NF ja foi rejeitada');

    diag.step = 'find_pdf';
    const folderPendente = `Notas Fiscais/Pendentes/${f.Unidade}/Diretoria ${f.Diretoria}`;
    const folderRejeitada = `Notas Fiscais/Rejeitadas/${f.Unidade}/Diretoria ${f.Diretoria}`;
    let pdfTarget = null;
    try {
      const folderListResp = await client.api(`/sites/${siteId}/drive/root:/${folderPendente}:/children`).get();
      const files = (folderListResp.value || []).filter(x => x.file);
      const numero = String(f.NumeroNF || '');
      pdfTarget = files.find(x => x.name && x.name.startsWith(numero + '_'));
      if (!pdfTarget && files.length > 0) {
        pdfTarget = files.sort((a,b) => (b.lastModifiedDateTime||'').localeCompare(a.lastModifiedDateTime||''))[0];
      }
    } catch (e) {
      diag.findPdfWarning = e.message;
    }

    diag.step = 'move_pdf';
    let urlPDFRejeitado = '';
    if (pdfTarget) {
      // Move o PDF: copia pra Rejeitadas + deleta original.
      // Graph nao tem "move" entre pastas em 1 chamada, entao fazemos copy + delete
      const downloadUrl = pdfTarget['@microsoft.graph.downloadUrl'];
      const dlResp = await fetch(downloadUrl);
      const pdfBuffer = Buffer.from(await dlResp.arrayBuffer());

      const uploadPath = `/sites/${siteId}/drive/root:/${encodeURIComponent(folderRejeitada)}/${encodeURIComponent(pdfTarget.name)}:/content`;
      const uploadResp = await client.api(uploadPath).header('Content-Type', 'application/pdf').put(pdfBuffer);
      urlPDFRejeitado = uploadResp.webUrl;
      diag.movedTo = urlPDFRejeitado;

      // Deleta original
      await client.api(`/sites/${siteId}/drive/items/${pdfTarget.id}`).delete();
      diag.deletedFromPendentes = true;
    } else {
      diag.pdfNotFound = true;
    }

    diag.step = 'update_list';
    const motivoCompleto = observacao ? `${motivo} — ${observacao}` : motivo;
    const patchPayload = buildPatchPayload({
      Status: 'Rejeitada',
      MotivoRejeicao: motivoCompleto,
      AprovadoEm: new Date().toISOString()
    }, colMap, colTypes);
    await client.api(`/sites/${siteId}/lists/${listNotasId}/items/${itemId}/fields`).patch(patchPayload);

    context.res = {
      status: 200,
      headers: { 'Content-Type': 'application/json' },
      body: { ok: true, itemId, motivo: motivoCompleto, urlPDFRejeitado, diag }
    };
  } catch (err) {
    context.log && context.log.error && context.log.error('RejeitarNota error:', err);
    context.res = {
      status: 500,
      headers: { 'Content-Type': 'application/json' },
      body: {
        error: (err && err.message) || String(err),
        statusCode: err && err.statusCode,
        graphBody: err && err.body,
        diag
      }
    };
  }
};

function errResp(context, status, msg, extra) {
  context.res = {
    status,
    headers: { 'Content-Type': 'application/json' },
    body: Object.assign({ error: msg }, extra || {})
  };
}
