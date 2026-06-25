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
const { getMergedRoles, isAdminEmail } = require('../shared/authz');
const { registrar: auditRegistrar } = require('../shared/auditLog');
const { notificar } = require('../shared/notificar');
const { ClientSecretCredential } = require('@azure/identity');
const { Client } = require('@microsoft/microsoft-graph-client');
const { TokenCredentialAuthenticationProvider } =
  require('@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials');

const LIST_NOTAS = 'PRONEP-NF-NotasFiscais';
const LIST_CONFIG = 'PRONEP-NF-Config';
const CONFIG_TITLE = 'global';
const configCache = { config: null, ts: 0 };
const CONFIG_TTL = 5 * 60 * 1000;

async function getConfigSistema(client, siteId) {
  if (configCache.config && (Date.now() - configCache.ts) < CONFIG_TTL) return configCache.config;
  try {
    const lists = await client.api('/sites/' + siteId + '/lists')
      .filter("displayName eq '" + LIST_CONFIG + "'")
      .get();
    if (!lists.value || !lists.value.length) {
      configCache.config = { multiNivel: { habilitado: false } };
      configCache.ts = Date.now();
      return configCache.config;
    }
    const listId = lists.value[0].id;
    const items = await client.api('/sites/' + siteId + '/lists/' + listId + '/items')
      .expand('fields').top(20).get();
    const item = (items.value || []).find(function (x) { return x.fields && x.fields.Title === CONFIG_TITLE; });
    if (item && item.fields && item.fields.ConfigJson) {
      try { configCache.config = JSON.parse(item.fields.ConfigJson); }
      catch (e) { configCache.config = { multiNivel: { habilitado: false } }; }
    } else {
      configCache.config = { multiNivel: { habilitado: false } };
    }
    configCache.ts = Date.now();
    return configCache.config;
  } catch (e) {
    return { multiNivel: { habilitado: false } };
  }
}

// Resolve quem eh o gestor master do 2o nivel
function resolverGestorMaster(config, diretoria) {
  if (!config || !config.multiNivel || !config.multiNivel.habilitado) return null;
  const mn = config.multiNivel;
  if (mn.modoAprovador === 'global') return mn.gestorMasterGlobal || null;
  if (mn.modoAprovador === 'porDiretoria') {
    return (mn.gestoresPorDiretoria || {})[diretoria] || null;
  }
  return null;
}
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

  // ignoreEncryption: alguns PDFs de NF vem com criptografia/permissoes (mesmo sem
  // senha de abertura). Sem isso o pdf-lib lanca "document is encrypted" no watermark.
  const pdfDoc = await PDFDocument.load(pdfBuffer, { ignoreEncryption: true });
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
    // Campos opcionais de compliance (so vem populados quando NF vencendo em < D+5)
    const alinhouFinanceiro = body.alinhouFinanceiro === true || body.alinhouFinanceiro === 'true';
    const naoAlinhou = body.alinhouFinanceiro === false || body.alinhouFinanceiro === 'false';
    const compliancePresente = alinhouFinanceiro || naoAlinhou;
    const gestorFinanceiroAlinhado = String(body.gestorFinanceiroAlinhado || '').trim();

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

    // RBAC: so o aprovador atribuido (ou admin) pode aprovar.
    // A4: admin vem do grupo Entra 'administrador' OU da whitelist central ADMIN_EMAILS
    // (antes era o e-mail hardcoded — e membro do grupo admin nem conseguia aprovar).
    const userRolesAA = await getMergedRoles(req, user);
    const isAdmin = userRolesAA.includes('administrador') || isAdminEmail(aprovadorEmail);
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

    // === LOGICA MULTI-NIVEL: verifica se precisa encaminhar pro 2o aprovador ===
    diag.step = 'check_multinivel';
    const config = await getConfigSistema(client, siteId);
    diag.configMultiNivel = config && config.multiNivel ? {
      habilitado: config.multiNivel.habilitado,
      valorLimite: config.multiNivel.valorLimite,
      modoAprovador: config.multiNivel.modoAprovador
    } : null;
    const valorNF = Number(f.Valor || 0);
    const jaEstaNoSegundoNivel = (f.Status === 'AguardandoN2');
    const precisaSegundoNivel = !jaEstaNoSegundoNivel
      && config && config.multiNivel && config.multiNivel.habilitado
      && valorNF > Number(config.multiNivel.valorLimite || 0);

    if (precisaSegundoNivel) {
      const gestorMaster = resolverGestorMaster(config, f.Diretoria);
      diag.gestorMaster = gestorMaster;
      if (!gestorMaster) {
        context.res = { status: 400, body: {
          error: 'Aprovacao multi-nivel habilitada mas nao tem gestor master definido pra diretoria ' + f.Diretoria,
          diag: diag
        }};
        return;
      }
      {
        // Sempre encaminha pro 2o nivel quando valor > limite, mesmo que aprovador == gestor master.
        // Isso forca registro auditavel de 2 aprovacoes distintas (boa pratica de governanca).
        diag.step = 'encaminhar_n2';
        const patchFields = buildPatchPayload({
          Status: 'AguardandoN2',
          AprovadorAtual: gestorMaster
        }, colMap, colTypes);
        await client.api(`/sites/${siteId}/lists/${listNotasId}/items/${itemId}/fields`).patch(patchFields);

        diag.step = 'notify_n2';
        try {
          await notificar('lancada', [gestorMaster], {
            itemId: itemId,
            numero: f.NumeroNF, fornecedor: f.CNPJFornecedor, valor: f.Valor,
            vencimento: f.DataVencimento, unidade: f.Unidade, diretoria: f.Diretoria,
            aprovador: 'Aprovacao 2o nivel - acima do limite',
            submitter: f.LancadoPor,
            urlPDF: f.UrlPDF
          });
        } catch (notifErr) {
          diag.notifyN2Error = notifErr.message;
        }

        context.res = {
          status: 200,
          headers: { 'Content-Type': 'application/json' },
          body: {
            ok: true,
            encaminhado: true,
            mensagem: 'NF encaminhada para aprovacao do 2o nivel',
            proximoAprovador: gestorMaster,
            itemId: itemId,
            diag: diag
          }
        };
        return;
      }
    }
    // === FIM lógica multi-nivel — segue fluxo de aprovacao final ===

    diag.step = 'find_pdf';
    // Encontra o PDF em /Pendentes/{Unidade}/Diretoria {Diretoria}/
    const folder = `Notas Fiscais/Pendentes/${f.Unidade}/Diretoria ${f.Diretoria}`;
    // Lista arquivos da pasta e acha o que bate com o NumeroNF (ou usa UrlPDF do item)
    const folderListResp = await client.api(`/sites/${siteId}/drive/root:/${folder}:/children`).get();
    const files = (folderListResp.value || []).filter(x => x.file);
    // Tenta achar pelo numero. Suporta:
    //   - Padrao antigo: {numero}_{razao}_{ts}_*.pdf (numero no inicio, sem padding)
    //   - Padrao intermediario: {data}_{012345}_..._.pdf (numero zero-padded - 1a versao)
    //   - Padrao atual: {data}_{12345}_..._.pdf (numero sem zeros a esquerda)
    const numero = String(f.NumeroNF || '');
    const numClean = numero.replace(/[^A-Za-z0-9]/g, '');
    const numUnpadded = /^\d+$/.test(numClean) ? (numClean.replace(/^0+/, '') || '0') : numClean;
    const numPadded = /^\d+$/.test(numClean) ? numClean.padStart(6, '0') : numClean;
    const candidates = Array.from(new Set([numero, numClean, numUnpadded, numPadded].filter(Boolean)));
    let target = null;
    for (const n of candidates) {
      target = files.find(x => x.name && x.name.startsWith(n + '_'));
      if (target) break;
    }
    if (!target) {
      for (const n of candidates) {
        target = files.find(x => x.name && x.name.indexOf('_' + n + '_') >= 0);
        if (target) break;
      }
    }
    // A2: SEM fallback "mais recente". Antes, se nenhum arquivo batesse pelo numero,
    // pegava o PDF modificado mais recentemente da pasta COMPARTILHADA e o carimbava/
    // movia/deletava — podendo agir sobre a NF de OUTRO fornecedor. Agora falha com 404.
    if (!target) {
      const nomes = files.map(x => x.name).slice(0, 15);
      context.res = { status: 404, body: { error: 'PDF da NF ' + numero + ' nao encontrado em Pendentes pelo numero. Verifique o arquivo no SharePoint (pode ter sido renomeado).', folder, arquivosDisponiveis: nomes } };
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
    // Data BRT (UTC-3): server roda em UTC, ajusta pra fuso de Brasilia
    // Senao apos 21h BRT vira dia seguinte no UTC
    const agora = new Date();
    const brt = new Date(agora.getTime() - 3 * 60 * 60 * 1000);
    const dataPasta = `${brt.getUTCFullYear()}-${String(brt.getUTCMonth()+1).padStart(2,'0')}-${String(brt.getUTCDate()).padStart(2,'0')}`;
    diag.dataPastaBRT = dataPasta;
    // Padrao Pronep: adiciona _APROVADA_{dataAprovacao} antes da extensao
    // Funciona com nomes antigos e novos — apenas insere o sufixo
    let aprovadoName = target.name;
    if (/\.pdf$/i.test(aprovadoName)) {
      aprovadoName = aprovadoName.replace(/\.pdf$/i, `_APROVADA_${dataPasta}.pdf`);
    } else {
      aprovadoName = `${aprovadoName}_APROVADA_${dataPasta}.pdf`;
    }
    diag.aprovadoName = aprovadoName;
    const folderAprov = `Notas Fiscais/Notas Aprovadas/${f.Unidade}/${dataPasta}`;
    const uploadPath = `/sites/${siteId}/drive/root:/${encodeURIComponent(folderAprov)}/${encodeURIComponent(aprovadoName)}:/content`;
    const uploadResp = await client.api(uploadPath).header('Content-Type', 'application/pdf').put(stampedPdf);
    diag.uploadedTo = uploadResp.webUrl;

    diag.step = 'delete_pendente';
    await client.api(`/sites/${siteId}/drive/items/${target.id}`).delete();

    diag.step = 'update_list';
    // QUIRK Graph API: hyperlinks misturados com outros campos no mesmo PATCH as vezes
    // dao 400. Separa em 2 PATCHes: primeiro campos base, depois hyperlinks.
    const camposBase = {
      Status: 'Aprovada',
      AprovadoEm: new Date().toISOString()
    };
    // Compliance: aprovacao de NF vencendo em < D+5 - registra se alinhou com Financeiro
    if (compliancePresente) {
      camposBase.AlinhouFinanceiro = alinhouFinanceiro;
      camposBase.GestorFinanceiroAlinhado = alinhouFinanceiro ? gestorFinanceiroAlinhado : 'NAO_ALINHADO_SEGUIU_MESMO_ASSIM';
      diag.compliance = { alinhou: alinhouFinanceiro, gestor: camposBase.GestorFinanceiroAlinhado };
    }
    const basePayload = buildPatchPayload(camposBase, colMap, colTypes);
    await client.api(`/sites/${siteId}/lists/${listNotasId}/items/${itemId}/fields`).patch(basePayload);

    // FIX: grava URL como STRING na coluna UrlPDFAprovadoStr (text-multiline criada via Graph).
    // A coluna UrlPDFAprovado original tem tipo nao reconhecido pelo Graph (sem 'text' nem
    // 'hyperlinkOrPicture' no schema) - sempre falha. Migrar pra Str resolve.
    diag.step = 'update_list_url_string';
    try {
      await client.api(`/sites/${siteId}/lists/${listNotasId}/items/${itemId}/fields`)
        .patch({ UrlPDFAprovadoStr: uploadResp.webUrl });
      diag.urlPDFAprovadoStrOk = true;
    } catch (urlStrErr) {
      diag.urlPDFAprovadoStrError = {
        message: urlStrErr.message,
        statusCode: urlStrErr.statusCode,
        body: urlStrErr.body
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

    // Audit log — best effort
    auditRegistrar(user, 'aprovacao',
      { tipo: 'nf', id: itemId, numero: f.NumeroNF },
      'sucesso',
      { fornecedor: f.CNPJFornecedor, valor: f.Valor, vencimento: f.DataVencimento, unidade: f.Unidade, diretoria: f.Diretoria, urlPDF: uploadResp.webUrl }
    ).catch(function(){});

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
