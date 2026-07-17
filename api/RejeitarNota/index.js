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
const { getUser } = require('../shared/auth');
const { getUserRoles } = require('../shared/userRoles');
const { isAdminEmail } = require('../shared/authz');
const { notificar } = require('../shared/notificar');
const { registrar: auditRegistrar } = require('../shared/auditLog');
const { getGraphClient } = require('../shared/graph');

const LIST_NOTAS = 'PRONEP-NF-NotasFiscais';
// Resolver local: alem de siteId/listId, resolve driveId + colMap desta lista.
const cache = { siteId: null, driveId: null, listNotasId: null, colMap: null, invColMap: null, colTypes: null, colMapCachedAt: 0 };
const CACHE_TTL_MS = 5 * 60 * 1000;

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

// Extrai a URL de um campo hyperlink (pode vir string ou { Url, Description }).
function urlDeCampo(v) {
  if (!v) return '';
  if (typeof v === 'string' && v.indexOf('http') === 0) return v;
  if (typeof v === 'object' && v.Url && String(v.Url).indexOf('http') === 0) return v.Url;
  return '';
}
// Nome exato do arquivo a partir da URL armazenada na nota (unico: inclui o valor).
function nomeArquivoDeUrl(url) {
  if (!url) return '';
  try { return decodeURIComponent(String(url).split('?')[0].split('/').pop() || ''); }
  catch (e) { return String(url).split('?')[0].split('/').pop() || ''; }
}

// Aplica watermark REJEITADA (vermelho, 4 linhas: REJEITADA + data + por + motivo)
async function aplicarWatermarkRejeitado(pdfBuffer, aprovadorEmail, motivo) {
  // LAZY require — so carrega pdf-lib quando essa funcao for chamada
  const { PDFDocument, StandardFonts, rgb } = require('pdf-lib');

  // ignoreEncryption: PDFs de NF as vezes vem com criptografia/permissoes — sem isso
  // o pdf-lib falha no watermark com "document is encrypted".
  const pdfDoc = await PDFDocument.load(pdfBuffer, { ignoreEncryption: true });
  // Se criptografado, salvar com watermark geraria arquivo corrompido — arquiva o
  // original intacto (sem o carimbo) pra nao perder o documento.
  if (pdfDoc.isEncrypted) return pdfBuffer;
  const helveticaBold = await pdfDoc.embedFont(StandardFonts.HelveticaBold);
  const helvetica = await pdfDoc.embedFont(StandardFonts.Helvetica);

  // Data/hora BR
  const now = new Date();
  const brOffset = -3 * 60;
  const local = new Date(now.getTime() + (brOffset + now.getTimezoneOffset()) * 60 * 1000);
  const pad = n => String(n).padStart(2,'0');
  const dataStr = `${pad(local.getDate())}/${pad(local.getMonth()+1)}/${local.getFullYear()} ${pad(local.getHours())}:${pad(local.getMinutes())}:${pad(local.getSeconds())}`;
  const prefixoEmail = aprovadorEmail.split('@')[0];

  // Cor vermelha (#C62828) com transparencia 0.35
  const vermelho = rgb(0.776, 0.157, 0.157);
  const alpha = 0.35;

  // Trunca motivo pra caber na pagina (max ~80 chars)
  const motivoTexto = String(motivo || '').slice(0, 90);

  const pages = pdfDoc.getPages();
  for (const page of pages) {
    const { width, height } = page.getSize();

    // Linha 1: REJEITADA (grande, vermelho)
    const txt1 = 'REJEITADA';
    const fs1 = Math.min(width / 8, 78);
    const w1 = helveticaBold.widthOfTextAtSize(txt1, fs1);
    page.drawText(txt1, {
      x: (width - w1) / 2, y: height / 2 + 80,
      size: fs1, font: helveticaBold, color: vermelho, opacity: alpha
    });

    // Linha 2: data/hora
    const fs2 = Math.min(width / 22, 24);
    const w2 = helvetica.widthOfTextAtSize(dataStr, fs2);
    page.drawText(dataStr, {
      x: (width - w2) / 2, y: height / 2 + 30,
      size: fs2, font: helvetica, color: vermelho, opacity: alpha
    });

    // Linha 3: Por: email
    const txt3 = `Por: ${prefixoEmail}`;
    const fs3 = Math.min(width / 24, 20);
    const w3 = helvetica.widthOfTextAtSize(txt3, fs3);
    page.drawText(txt3, {
      x: (width - w3) / 2, y: height / 2,
      size: fs3, font: helvetica, color: vermelho, opacity: alpha
    });

    // Linha 4: Motivo
    const txt4 = `Motivo: ${motivoTexto}`;
    const fs4 = Math.min(width / 30, 16);
    const w4 = helvetica.widthOfTextAtSize(txt4, fs4);
    page.drawText(txt4, {
      x: (width - w4) / 2, y: height / 2 - 30,
      size: fs4, font: helvetica, color: vermelho, opacity: alpha
    });
  }

  const pdfBytes = await pdfDoc.save();
  return Buffer.from(pdfBytes);
}

function buildPatchPayload(displayPayload, colMap, colTypes) {
  const fields = {};
  for (const [displayName, value] of Object.entries(displayPayload)) {
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
    const user = await getUser(req);
    if (!user) return errResp(context, 401, 'Nao autenticado');
    const aprovadorEmail = (user.email || '').toLowerCase();
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

    // RBAC: 3 fontes de roles (claims JWT + principal Easy Auth + Graph AAD groups).
    // SEMPRE mescla com Graph porque SWA principal nao popula grupos AAD.
    function readPrincipalRolesLocal(req) {
      const h = req.headers && req.headers['x-ms-client-principal'];
      if (!h) return [];
      try {
        const p = JSON.parse(Buffer.from(h, 'base64').toString('utf-8'));
        return p.userRoles || [];
      } catch (e) { return []; }
    }
    const claimsRoles = (user.claims && user.claims.roles) || [];
    const principalRolesRaw = readPrincipalRolesLocal(req);
    const usefulPrincipalRoles = principalRolesRaw.filter(r => r !== 'authenticated' && r !== 'anonymous');
    const graphRoles = await getUserRoles(user);
    const userRoles = Array.from(new Set([...claimsRoles, ...usefulPrincipalRoles, ...graphRoles]));
    diag.userRoles = userRoles;
    const isAdmin = userRoles.includes('administrador') || isAdminEmail(aprovadorEmail); // A4: centralizado
    const isFinanceiro = userRoles.includes('financeiro_nf');
    const isAprovadorAtribuido = (f.AprovadorAtual || '').toLowerCase() === aprovadorEmail;

    // ESTORNO: se NF ja esta Aprovada, SO admin ou financeiro pode rejeitar (estornar)
    const ehEstorno = f.Status === 'Aprovada';
    if (ehEstorno) {
      if (!isAdmin && !isFinanceiro) {
        return errResp(context, 403, 'NF ja aprovada — apenas Admin ou Financeiro pode estornar', { status: f.Status });
      }
    } else {
      // Fluxo normal: rejeicao de NF pendente
      if (!isAdmin && !isAprovadorAtribuido) {
        return errResp(context, 403, 'Voce nao eh o aprovador desta NF', { aprovadorAtual: f.AprovadorAtual, voce: aprovadorEmail });
      }
    }
    if (f.Status === 'Rejeitada') return errResp(context, 409, 'NF ja foi rejeitada');
    diag.ehEstorno = ehEstorno;

    diag.step = 'find_pdf';
    // Se estorno (status era Aprovada): busca em "Notas Aprovadas" (sem subpastas, pasta plana legacy).
    // Se rejeicao normal: busca em "Notas Fiscais/Pendentes/{Unidade}/Diretoria {Diretoria}".
    const folderPendente = `Notas Fiscais/Pendentes/${f.Unidade}/Diretoria ${f.Diretoria}`;
    const folderAprovada = `Notas Aprovadas`;
    // CORRECAO: estrutura de subpastas por Unidade+Diretoria igual Pendentes (antes salvava plano).
    // O AbrirPdfDaNota tem fallback pra pasta raiz pra NFs rejeitadas antes desta correcao.
    const folderRejeitada = `Notas Fiscais/Rejeitadas/${f.Unidade}/Diretoria ${f.Diretoria}`;
    const folderOrigem = ehEstorno ? folderAprovada : folderPendente;
    diag.folderOrigem = folderOrigem;
    let pdfTarget = null;
    try {
      const folderListResp = await client.api(`/sites/${siteId}/drive/root:/${folderOrigem}:/children`).get();
      const files = (folderListResp.value || []).filter(x => x.file);

      // (1) FONTE DA VERDADE: nome EXATO do arquivo da propria nota, extraido da URL
      // armazenada (UrlPDFAprovado no estorno; UrlPDF no fluxo normal). O nome inclui o
      // valor -> unico. Imune a NumeroNF duplicado/colidindo com o sequencial do arquivo.
      const urlNota = ehEstorno
        ? (urlDeCampo(f.UrlPDFAprovadoStr) || urlDeCampo(f.UrlPDFAprovado) || urlDeCampo(f.UrlPDFStr) || urlDeCampo(f.UrlPDF))
        : (urlDeCampo(f.UrlPDFStr) || urlDeCampo(f.UrlPDF));
      const nomeExato = nomeArquivoDeUrl(urlNota);
      if (nomeExato) pdfTarget = files.find(x => x.name === nomeExato);
      diag.matchPor = pdfTarget ? 'nome_exato' : null;

      // (2) FALLBACK ESTRITO (so se a nota nao tem URL/nome): exige numero E valor no
      // nome, e so aceita se for UNICO. NUNCA aceita match frouxo unico por numero
      // (era o bug: NumeroNF=3 casava com o sequencial _3_ de outro arquivo).
      if (!pdfTarget) {
        const numero = String(f.NumeroNF || '').trim();
        const valorNum = (typeof f.Valor === 'number' ? f.Valor : Number(f.Valor)) || 0;
        const valorStr = valorNum > 0 ? valorNum.toFixed(2).replace('.', ',') : '';
        if (numero && valorStr) {
          const cand = files.filter(x => x.name
            && (x.name.startsWith(numero + '_') || x.name.includes('_' + numero + '_'))
            && (x.name.includes('_' + valorStr + '_') || x.name.includes('_' + valorStr + '.')));
          if (cand.length === 1) { pdfTarget = cand[0]; diag.matchPor = 'numero+valor'; }
          else diag.matchAmbiguo = { numero, valorStr, encontrados: cand.length };
        }
      }
      // Sem identificacao CONFIAVEL, nao move/carimba/deleta nada (evita mexer no arquivo
      // errado). A NF e rejeitada; o PDF fica pra reconciliacao manual.
    } catch (e) {
      diag.findPdfWarning = e.message;
    }

    diag.step = 'watermark_and_move';
    const motivoBase = observacao ? `${motivo} — ${observacao}` : motivo;
    const motivoCompletoStr = ehEstorno ? `[ESTORNO PELO FINANCEIRO] ${motivoBase}` : motivoBase;
    let urlPDFRejeitado = '';
    if (pdfTarget) {
      // Baixa, aplica watermark REJEITADA, sobe pra Rejeitadas, deleta original
      const downloadUrl = pdfTarget['@microsoft.graph.downloadUrl'];
      const dlResp = await fetch(downloadUrl);
      const pdfBuffer = Buffer.from(await dlResp.arrayBuffer());
      diag.pdfSize = pdfBuffer.length;

      const stampedPdf = await aplicarWatermarkRejeitado(pdfBuffer, aprovadorEmail, motivoCompletoStr);
      diag.stampedSize = stampedPdf.length;

      const uploadPath = `/sites/${siteId}/drive/root:/${encodeURIComponent(folderRejeitada)}/${encodeURIComponent(pdfTarget.name)}:/content`;
      const uploadResp = await client.api(uploadPath).header('Content-Type', 'application/pdf').put(stampedPdf);
      urlPDFRejeitado = uploadResp.webUrl;
      diag.movedTo = urlPDFRejeitado;

      // Deleta original
      await client.api(`/sites/${siteId}/drive/items/${pdfTarget.id}`).delete();
      diag.deletedFromPendentes = true;
    } else {
      diag.pdfNotFound = true;
    }

    diag.step = 'update_list';
    const patchPayload = buildPatchPayload({
      Status: 'Rejeitada',
      MotivoRejeicao: motivoCompletoStr,
      AprovadoEm: new Date().toISOString()
    }, colMap, colTypes);
    await client.api(`/sites/${siteId}/lists/${listNotasId}/items/${itemId}/fields`).patch(patchPayload);

    // Grava a URL do PDF rejeitado na nota -> o "Ver" (AbrirPdfDaNota) redireciona
    // DIRETO pra ela, sem varrer a pasta por numero. Patch SEPARADO e best-effort:
    // colunas hyperlink antigas as vezes falham no Graph, e isso nao pode quebrar a rejeicao.
    if (urlPDFRejeitado) {
      try {
        const urlBuilt = buildPatchPayload({ UrlPDFStr: urlPDFRejeitado, UrlPDF: urlPDFRejeitado }, colMap, colTypes);
        if (Object.keys(urlBuilt).length) {
          await client.api(`/sites/${siteId}/lists/${listNotasId}/items/${itemId}/fields`).patch(urlBuilt);
          diag.urlPDFGravada = true;
        }
      } catch (e) { diag.urlPDFGravadaErro = e.message; }
    }

    // Dispara notificacao pro submitter (quem lancou).
    // Se for ESTORNO, notifica tambem o gestor que aprovou originalmente (AprovadorAtual da NF).
    diag.step = 'notify';
    const destinatarios = [f.LancadoPor || aprovadorEmail];
    if (ehEstorno && f.AprovadorAtual && (f.AprovadorAtual || '').toLowerCase() !== (f.LancadoPor || '').toLowerCase()) {
      destinatarios.push(f.AprovadorAtual);
    }
    await notificar('rejeitada', destinatarios, {
      numero: f.NumeroNF, fornecedor: f.CNPJFornecedor, valor: f.Valor,
      vencimento: f.DataVencimento, unidade: f.Unidade, diretoria: f.Diretoria,
      aprovador: aprovadorEmail, submitter: f.LancadoPor,
      motivo: motivoCompletoStr,
      urlPDF: urlPDFRejeitado
    });

    auditRegistrar(user, 'rejeicao',
      { tipo: 'nf', id: itemId, numero: f.NumeroNF },
      'sucesso',
      { fornecedor: f.CNPJFornecedor, valor: f.Valor, motivo: motivoCompletoStr, unidade: f.Unidade, diretoria: f.Diretoria, posAprovacao: ehEstorno, gestorOriginal: ehEstorno ? f.AprovadorAtual : undefined }
    ).catch(function(){});

    context.res = {
      status: 200,
      headers: { 'Content-Type': 'application/json' },
      body: { ok: true, itemId, motivo: motivoCompletoStr, urlPDFRejeitado, diag }
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
