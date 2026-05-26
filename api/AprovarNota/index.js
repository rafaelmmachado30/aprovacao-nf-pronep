/**
 * Sistema de Aprovacao de NF - AprovarNota
 *
 * POST /api/AprovarNota
 * Body JSON: { id: "<spListItemId>" }
 *
 * Fluxo (replica o PA antigo):
 *  1. Valida role do usuario (precisa ser gestor da diretoria correspondente)
 *  2. Le item da lista NotasFiscais pra obter UrlPDF, Unidade, Diretoria
 *  3. Baixa o PDF de /Pendentes/{Unidade}/Diretoria {Diretoria}/
 *  4. Aplica watermark APROVADO + timestamp + email aprovador (3 linhas azul)
 *  5. Sobe PDF pra /Notas Aprovadas/{Unidade}/{AAAA-MM-DD}/ (cria pasta do dia)
 *  6. Deleta o original de Pendentes
 *  7. Atualiza Status=Aprovada na lista
 *
 * IMPORTANTE: pdf-lib eh carregado LAZY (require dentro da funcao) pra nao
 * quebrar o startup das outras Functions caso pdf-lib tenha problema.
 */

require('isomorphic-fetch');
const { getUser } = require('../shared/auth');
const { notificar } = require('../shared/notificar');
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

// Le um item normalizado (internal -> display)
function normalizeFields(fields, invColMap) {
  const out = {};
  for (const [k, v] of Object.entries(fields)) {
    if (invColMap[k]) out[invColMap[k]] = v;
  }
  return out;
}

// Constroi payload pra PATCH usando colMap + colTypes
function buildPatchPayload(displayNamePayload, colMap, colTypes) {
  const fields = {};
  for (const [displayName, value] of Object.entries(displayNamePayload)) {
    const internal = colMap[displayName];
    if (!internal) continue;
    if (value === null || value === undefined) continue;
    const t = colTypes[internal] || 'text';
    let formatted = value;
    if (t === 'hyperlink') {
      const url = String(value);
      let descr = 'Ver PDF';
      try {
        const decoded = decodeURIComponent(url);
        const m = decoded.match(/\/([^\/?#]+\.pdf)(?:[?#]|$)/i);
        if (m && m[1]) descr = m[1].substring(0, 100);
      } catch (e) {}
      formatted = { Url: url, Description: descr };
    }
    else if (t === 'number' || t === 'currency') formatted = Number(value);
    else if (t === 'boolean') formatted = (value === true || value === 'Sim');
    else if (t === 'dateTime') formatted = value;
    else formatted = String(value);
    fields[internal] = formatted;
  }
  return fields;
}

// Aplica watermark APROVADO no PDF (3 linhas azul, igual PA antigo)
async function aplicarWatermark(pdfBuffer, aprovadorEmail) {
  // LAZY require — so carrega pdf-lib quando essa funcao for chamada
  const { PDFDocument, StandardFonts, rgb } = require('pdf-lib');

  const pdfDoc = await PDFDocument.load(pdfBuffer);
  const helveticaBold = await pdfDoc.embedFont(StandardFonts.HelveticaBold);
  const helvetica = await pdfDoc.embedFont(StandardFonts.Helvetica);

  // Data/hora BR (timezone America/Sao_Paulo aproximada via UTC-3)
  const now = new Date();
  const brOffset = -3 * 60; // minutos UTC-3
  const local = new Date(now.getTime() + (brOffset + now.getTimezoneOffset()) * 60 * 1000);
  const pad = n => String(n).padStart(2,'0');
  const dataStr = `${pad(local.getDate())}/${pad(local.getMonth()+1)}/${local.getFullYear()} ${pad(local.getHours())}:${pad(local.getMinutes())}:${pad(local.getSeconds())}`;
  const prefixoEmail = aprovadorEmail.split('@')[0];

  // Cor azul com transparencia 0.3
  const azul = rgb(0, 0, 1); // #0000FF
  const alpha = 0.3;

  // Aplica nas paginas (primeira pagina principal)
  const pages = pdfDoc.getPages();
  for (const page of pages) {
    const { width, height } = page.getSize();

    // Linha 1: APROVADO (grande)
    const aprovadoFontSize = Math.min(width / 8, 80);
    const aprovadoWidth = helveticaBold.widthOfTextAtSize('APROVADO', aprovadoFontSize);
    page.drawText('APROVADO', {
      x: (width - aprovadoWidth) / 2,
      y: height / 2 + 60,
      size: aprovadoFontSize,
      font: helveticaBold,
      color: azul,
      opacity: alpha,
      rotate: { type: 'degrees', angle: 0 }
    });

    // Linha 2: data/hora
    const dataFontSize = Math.min(width / 22, 24);
    const dataWidth = helvetica.widthOfTextAtSize(dataStr, dataFontSize);
    page.drawText(dataStr, {
      x: (width - dataWidth) / 2,
      y: height / 2 + 10,
      size: dataFontSize,
      font: helvetica,
      color: azul,
      opacity: alpha
    });

    // Linha 3: Por: prefixo-email
    const porTexto = `Por: ${prefixoEmail}`;
    const porFontSize = Math.min(width / 24, 20);
    const porWidth = helvetica.widthOfTextAtSize(porTexto, porFontSize);
    page.drawText(porTexto, {
      x: (width - porWidth) / 2,
      y: height / 2 - 20,
      size: porFontSize,
      font: helvetica,
      color: azul,
      opacity: alpha
    });
  }

  const pdfBytes = await pdfDoc.save();
  return Buffer.from(pdfBytes);
}

module.exports = async function (context, req) {
  const diag = { step: 'start' };
  try {
    diag.step = 'parse_body';
    const body = req.body || {};
    const itemId = body.id;
    if (!itemId) {
      context.res = { status: 400, body: { error: 'id obrigatorio' } };
      return;
    }
    diag.itemId = itemId;

    diag.step = 'principal';
    const user = await getUser(req);
    if (!user) {
      context.res = { status: 401, body: { error: 'Nao autenticado' } };
      return;
    }
    const aprovadorEmail = (user.email || '').toLowerCase();
    diag.aprovador = aprovadorEmail;
    diag.authSource = user.source;

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

    // RBAC: so o aprovador atribuido (ou admin) pode aprovar
    const isAdmin = aprovadorEmail === 'rafael.machado@pronep.com.br';
    const isAprovadorAtribuido = (f.AprovadorAtual || '').toLowerCase() === aprovadorEmail;
    if (!isAdmin && !isAprovadorAtribuido) {
      context.res = { status: 403, body: {
        error: 'Voce nao eh o aprovador desta NF',
        aprovadorAtual: f.AprovadorAtual,
        voce: aprovadorEmail
      }};
      return;
    }

    if (f.Status === 'Aprovada') {
      context.res = { status: 409, body: { error: 'NF ja esta aprovada' } };
      return;
    }
    if (f.Status === 'Rejeitada') {
      context.res = { status: 409, body: { error: 'NF foi rejeitada, nao pode ser aprovada' } };
      return;
    }

    diag.step = 'find_pdf';
    // Encontra o PDF em /Pendentes/{Unidade}/Diretoria {Diretoria}/
    const folder = `Notas Fiscais/Pendentes/${f.Unidade}/Diretoria ${f.Diretoria}`;
    // Lista arquivos da pasta e acha o que bate com o NumeroNF (ou usa UrlPDF do item)
    const folderListResp = await client.api(`/sites/${siteId}/drive/root:/${folder}:/children`).get();
    const files = (folderListResp.value || []).filter(x => x.file);
    // Tenta achar pelo numero (primeiro chunk do filename)
    const numero = String(f.NumeroNF || '');
    let target = files.find(x => x.name && x.name.startsWith(numero + '_'));
    if (!target && files.length > 0) {
      // Fallback: pega o mais recente
      target = files.sort((a,b) => (b.lastModifiedDateTime||'').localeCompare(a.lastModifiedDateTime||''))[0];
    }
    if (!target) {
      context.res = { status: 404, body: { error: 'PDF nao encontrado em Pendentes', folder } };
      return;
    }
    diag.pdfFound = { name: target.name, id: target.id };

    diag.step = 'download_pdf';
    const downloadUrl = target['@microsoft.graph.downloadUrl'];
    const dlResp = await fetch(downloadUrl);
    const pdfBuffer = Buffer.from(await dlResp.arrayBuffer());
    diag.pdfSize = pdfBuffer.length;

    diag.step = 'watermark';
    const stampedPdf = await aplicarWatermark(pdfBuffer, aprovadorEmail);
    diag.stampedSize = stampedPdf.length;

    diag.step = 'upload_aprovado';
    const hoje = new Date();
    const dataPasta = `${hoje.getFullYear()}-${String(hoje.getMonth()+1).padStart(2,'0')}-${String(hoje.getDate()).padStart(2,'0')}`;
    const folderAprov = `Notas Fiscais/Notas Aprovadas/${f.Unidade}/${dataPasta}`;
    const uploadPath = `/sites/${siteId}/drive/root:/${encodeURIComponent(folderAprov)}/${encodeURIComponent(target.name)}:/content`;
    const uploadResp = await client.api(uploadPath).header('Content-Type', 'application/pdf').put(stampedPdf);
    diag.uploadedTo = uploadResp.webUrl;

    diag.step = 'delete_pendente';
    await client.api(`/sites/${siteId}/drive/items/${target.id}`).delete();

    diag.step = 'update_list';
    // QUIRK Graph API: hyperlinks misturados com outros campos no mesmo PATCH as vezes
    // dao 400. Separa em 2 PATCHes: primeiro campos base, depois hyperlinks.
    const basePayload = buildPatchPayload({
      Status: 'Aprovada',
      AprovadoEm: new Date().toISOString()
    }, colMap, colTypes);
    await client.api(`/sites/${siteId}/lists/${listNotasId}/items/${itemId}/fields`).patch(basePayload);

    diag.step = 'update_list_hyperlink';
    const hyperlinkPayload = buildPatchPayload({
      UrlPDFAprovado: uploadResp.webUrl
    }, colMap, colTypes);
    try {
      await client.api(`/sites/${siteId}/lists/${listNotasId}/items/${itemId}/fields`).patch(hyperlinkPayload);
      diag.urlPDFAprovadoPatchOk = true;
    } catch (urlErr) {
      // Nao falha a aprovacao se o PATCH do hyperlink der erro
      diag.urlPDFAprovadoPatchError = {
        message: urlErr.message,
        statusCode: urlErr.statusCode,
        body: urlErr.body
      };
    }

    // Dispara notificacao pro submitter (quem lancou)
    diag.step = 'notify';
    await notificar('aprovada', [f.LancadoPor || aprovadorEmail], {
      numero: f.NumeroNF, fornecedor: f.CNPJFornecedor, valor: f.Valor,
      vencimento: f.DataVencimento, unidade: f.Unidade, diretoria: f.Diretoria,
      aprovador: aprovadorEmail, submitter: f.LancadoPor,
      urlPDF: uploadResp.webUrl
    });

    context.res = {
      status: 200,
      headers: { 'Content-Type': 'application/json' },
      body: {
        ok: true,
        itemId,
        urlPDFAprovado: uploadResp.webUrl,
        diag
      }
    };
  } catch (err) {
    context.log && context.log.error && context.log.error('AprovarNota error:', err);
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
