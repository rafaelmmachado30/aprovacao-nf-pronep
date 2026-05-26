/**
 * Sistema de Aprovacao de NF - PostNota
 *
 * Recebe nova NF via POST JSON com PDF em base64.
 * Sobe PDF pra SharePoint Pendentes + cria item na lista NotasFiscais.
 *
 * Body JSON esperado:
 *   {
 *     fornecedorCNPJ: "00.000.000/0001-00",  // ou CPF
 *     fornecedorRazao: "Razao Social",
 *     numero: "12345",
 *     serie: "1",
 *     valor: 1500.00,
 *     vencimento: "2026-06-30",
 *     unidade: "RJ",       // SP|RJ|ES
 *     diretoria: "Suprimentos",  // do fornecedor
 *     negociadoCom: null,  // email do gestor financeiro quando vencimento < D+5
 *     descricao: "Servicos de manutencao",
 *     fileBase64: "data:application/pdf;base64,JVBERi0xLjcK...",
 *     fileName: "NF-12345.pdf"
 *   }
 *
 * Retorna:
 *   200 { ok: true, id, urlPDF, aprovador: {email, nome} }
 *   400 { error: "..." }
 *   409 { error: "Duplicidade", duplicada: {...} }
 *   500 { error: "...", diag: {...} }
 */

require('isomorphic-fetch');
const { notificar } = require('../shared/notificar');
const crypto = require('crypto');
const { ClientSecretCredential } = require('@azure/identity');
const { Client } = require('@microsoft/microsoft-graph-client');
const { TokenCredentialAuthenticationProvider } =
  require('@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials');

const LIST_NOTAS = 'PRONEP-NF-NotasFiscais';
const LIST_DIRETORIAS = 'PRONEP-NF-Diretorias';
const cache = { siteId: null, driveId: null, listNotasId: null, listDirId: null, colMap: null, colTypes: null, colMapCachedAt: 0 };
const CACHE_TTL_MS = 5 * 60 * 1000; // 5 minutos

// Descobre mapping displayName -> internalName das colunas da lista NotasFiscais
// IMPORTANTE: filtra colunas read-only (LinkTitle, Edit, ItemChildCount, etc) e colunas
// de sistema (_UIVersionString, _ComplianceTag, etc) pra nao bagunçar o mapeamento
async function getColumnMap(client, siteId, listId) {
  const age = Date.now() - (cache.colMapCachedAt || 0);
  if (cache.colMap && age < CACHE_TTL_MS) return cache.colMap;
  const resp = await client
    .api(`/sites/${siteId}/lists/${listId}/columns`)
    .get();
  const map = {};
  const types = {};
  for (const col of (resp.value || [])) {
    if (!col.displayName || !col.name) continue;
    if (col.readOnly === true) continue;
    if (col.hidden === true) continue;
    if (col.name.startsWith('_')) continue;
    if (['LinkTitle','LinkTitleNoMenu','Edit','DocIcon','ItemChildCount',
         'FolderChildCount','AppAuthor','AppEditor','Attachments'].includes(col.name)) continue;
    map[col.displayName] = col.name;
    // Captura o TIPO da coluna pra logica especifica
    // Possiveis: text, number, dateTime, boolean, choice, hyperlinkOrPicture,
    //            personOrGroup, currency, calculated, lookup, ...
    let detectedType = 'text';
    if (col.text)                detectedType = 'text';
    else if (col.number)         detectedType = 'number';
    else if (col.dateTime)       detectedType = 'dateTime';
    else if (col.boolean)        detectedType = 'boolean';
    else if (col.choice)         detectedType = 'choice';
    else if (col.hyperlinkOrPicture) detectedType = 'hyperlink';
    else if (col.personOrGroup)  detectedType = 'person';
    else if (col.currency)       detectedType = 'currency';
    else if (col.calculated)     detectedType = 'calculated';
    else if (col.lookup)         detectedType = 'lookup';
    types[col.name] = detectedType;
  }
  cache.colMap = map;
  cache.colTypes = types;
  cache.colMapCachedAt = Date.now();
  return map;
}

// Constroi o objeto fields com internalNames corretos
function buildFieldsObject(colMap, data) {
  const fields = {};
  for (const [displayName, value] of Object.entries(data)) {
    const internal = colMap[displayName] || displayName;
    fields[internal] = value;
  }
  return fields;
}

// Formata valor de acordo com o tipo da coluna no SharePoint
function formatByType(internalName, value) {
  const t = (cache.colTypes && cache.colTypes[internalName]) || 'text';
  if (value === null || value === undefined || value === '') return undefined;
  switch (t) {
    case 'hyperlink':
      return { Url: String(value), Description: '' };
    case 'currency':
    case 'number': {
      const num = Number(value);
      return isNaN(num) ? undefined : num;
    }
    case 'boolean':
      return value === true || value === 'true' || value === 'Sim';
    case 'dateTime':
      return value;
    default:
      return String(value);
  }
}

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

async function resolveSiteAndDrive(client) {
  if (cache.siteId && cache.driveId && cache.listNotasId && cache.listDirId) return cache;
  const host = process.env.SHAREPOINT_SITE_HOSTNAME;
  const path = process.env.SHAREPOINT_SITE_PATH;
  if (!host || !path) throw new Error('SHAREPOINT_* incompletas');
  const siteResp = await client.api(`/sites/${host}:${path}`).get();
  cache.siteId = siteResp.id;
  const driveResp = await client.api(`/sites/${cache.siteId}/drive`).get();
  cache.driveId = driveResp.id;
  // List IDs
  const lists = await client.api(`/sites/${cache.siteId}/lists`).get();
  for (const l of lists.value || []) {
    if (l.displayName === LIST_NOTAS) cache.listNotasId = l.id;
    if (l.displayName === LIST_DIRETORIAS) cache.listDirId = l.id;
  }
  if (!cache.listNotasId) throw new Error(`Lista '${LIST_NOTAS}' nao encontrada`);
  if (!cache.listDirId)   throw new Error(`Lista '${LIST_DIRETORIAS}' nao encontrada`);
  return cache;
}

function readClientPrincipal(req) {
  const header = req.headers && req.headers['x-ms-client-principal'];
  if (!header) return null;
  try {
    const decoded = Buffer.from(header, 'base64').toString('utf-8');
    return JSON.parse(decoded);
  } catch (e) { return null; }
}

// Resolve quem eh o aprovador a partir da matriz Diretorias
async function resolveAprovador(client, siteId, listDirId, unidade, diretoria) {
  const chave = `${unidade}|${diretoria}`;
  const resp = await client
    .api(`/sites/${siteId}/lists/${listDirId}/items?expand=fields&$top=200`)
    .get();
  const found = (resp.value || []).find(item => (item.fields || {}).Title === chave);
  if (!found) return null;
  const f = found.fields || {};
  return { email: f.field_3 || '', nome: f.field_4 || '' };
}

// Checa duplicata: hash igual OU (CNPJ + numero + serie iguais)
async function verificaDuplicata(client, siteId, listNotasId, hash, cnpj, numero, serie) {
  // Le todas as notas atuais (TODO paginacao se muito grande)
  const resp = await client
    .api(`/sites/${siteId}/lists/${listNotasId}/items?expand=fields&$top=999`)
    .get();
  for (const item of (resp.value || [])) {
    const f = item.fields || {};
    // field_X reais sao mapeados conforme ListarNotas — provisoriamente comparamos fields conhecidos
    const itemHash   = f.HashSHA256 || f.field_14 || '';  // ajustaremos quando soubermos os field_N
    const itemDoc    = f.CNPJFornecedor || f.field_2 || '';
    const itemNum    = f.NumeroNF || f.field_1 || '';
    if (hash && itemHash && hash === itemHash) {
      return { motivo: 'hash', notaId: item.id, status: f.Status || f.field_8 };
    }
    if (cnpj && numero && itemDoc === cnpj && itemNum === numero) {
      return { motivo: 'cnpj_numero', notaId: item.id, status: f.Status || f.field_8 };
    }
  }
  return null;
}

module.exports = async function (context, req) {
  const diag = { step: 'start' };
  try {
    diag.step = 'parse_body';
    const body = req.body || {};
    const {
      fornecedorCNPJ, fornecedorRazao, numero, serie, valor, vencimento,
      unidade, diretoria, negociadoCom, descricao,
      fileBase64, fileName
    } = body;

    // Validacao
    if (!fornecedorCNPJ) return ctxErr(context, 400, 'fornecedorCNPJ obrigatorio');
    if (!numero)         return ctxErr(context, 400, 'numero obrigatorio');
    if (typeof valor !== 'number' || valor <= 0) return ctxErr(context, 400, 'valor invalido');
    if (!vencimento)     return ctxErr(context, 400, 'vencimento obrigatorio');
    if (!unidade)        return ctxErr(context, 400, 'unidade obrigatoria');
    if (!diretoria)      return ctxErr(context, 400, 'diretoria obrigatoria');
    if (!fileBase64)     return ctxErr(context, 400, 'fileBase64 obrigatorio');
    if (!fileName)       return ctxErr(context, 400, 'fileName obrigatorio');

    diag.step = 'decode_base64';
    // Remove o prefixo "data:application/pdf;base64,"
    const base64 = fileBase64.includes(',') ? fileBase64.split(',')[1] : fileBase64;
    const pdfBuffer = Buffer.from(base64, 'base64');
    if (pdfBuffer.length > 6 * 1024 * 1024) {
      return ctxErr(context, 400, 'Arquivo excede 6 MB (' + (pdfBuffer.length/1024/1024).toFixed(2) + ' MB)');
    }
    diag.fileSize = pdfBuffer.length;

    diag.step = 'hash';
    const hash = crypto.createHash('sha256').update(pdfBuffer).digest('hex');
    diag.hash = hash;

    diag.step = 'principal';
    const principal = readClientPrincipal(req);
    const submitterEmail = (principal && principal.userDetails) || 'desconhecido@pronep.com.br';
    diag.submitter = submitterEmail;

    diag.step = 'graph';
    const client = await getGraphClient();
    const { siteId, driveId, listNotasId, listDirId } = await resolveSiteAndDrive(client);
    diag.siteId = siteId;

    diag.step = 'aprovador';
    const aprovador = await resolveAprovador(client, siteId, listDirId, unidade, diretoria);
    if (!aprovador || !aprovador.email) {
      return ctxErr(context, 400, `Nao encontrado aprovador para ${unidade}|${diretoria}`);
    }
    diag.aprovador = aprovador.email;

    diag.step = 'check_duplicate';
    const dup = await verificaDuplicata(client, siteId, listNotasId, hash, fornecedorCNPJ, numero, serie);
    if (dup) {
      context.res = { status: 409, headers: { 'Content-Type': 'application/json' },
        body: { error: 'Duplicidade detectada', motivo: dup.motivo, notaId: dup.notaId, status: dup.status, diag } };
      return;
    }

    diag.step = 'upload_pdf';
    // Caminho: /Notas Fiscais/Pendentes/{Unidade}/Diretoria {Diretoria}/{fileName-saneado}
    const safeName = fileName.replace(/[^A-Za-z0-9._\\-\\s\\(\\)]/g, '').trim() || 'nota.pdf';
    const finalName = `${numero}_${fornecedorRazao || 'NF'}_${Date.now()}_${safeName}`
      .replace(/[^A-Za-z0-9._\\-\\s\\(\\)]/g, '_').slice(0, 200);
    const folder = `Notas Fiscais/Pendentes/${unidade}/Diretoria ${diretoria}`;
    const uploadPath = `/sites/${siteId}/drive/root:/${encodeURIComponent(folder)}/${encodeURIComponent(finalName)}:/content`;
    diag.uploadPath = uploadPath;
    const uploadResp = await client.api(uploadPath)
      .header('Content-Type', 'application/pdf')
      .put(pdfBuffer);
    diag.driveItemId = uploadResp.id;
    diag.webUrl = uploadResp.webUrl;

    diag.step = 'discover_columns';
    const colMap = await getColumnMap(client, siteId, listNotasId);
    diag.colMap = colMap;
    diag.colTypes = cache.colTypes;

    diag.step = 'create_list_item';
    // Title = "NF {numero}" ou descricao se nao tem numero
    const title = numero ? `NF ${numero}` : (descricao || 'NF sem numero').slice(0, 80);
    // Constroi o objeto com displayNames; buildFieldsObject converte pra internalNames
    // Formata datas como ISO completo (DataVencimento precisa de hora pra SharePoint Date+Time)
    // Se a coluna for soh Date, o SharePoint aceita ISO completo tambem (truncar)
    const isoVenc = (vencimento && vencimento.length === 10) ? (vencimento + 'T00:00:00Z') : vencimento;
    const rawFields = {
      Title:           title,
      NumeroNF:        numero || '',
      CNPJFornecedor:  fornecedorCNPJ,
      Descricao:       descricao || '',
      Valor:           valor,                       // <-- NUMERO (nao string)
      DataVencimento:  isoVenc,
      Unidade:         unidade,
      Diretoria:       diretoria,
      AprovadorAtual:  aprovador.email,
      Status:          'Lancada',
      LancadoPor:      submitterEmail,
      LancadoEm:       new Date().toISOString(),
      AprovadoEm:      null,                        // <-- null em vez de '' (campo data)
      MotivoRejeicao:  '',
      HashSHA256:      hash,
      UrlPDF:          uploadResp.webUrl || '',  // formatByType detecta Hyperlink e formata { Url, Description }
      UrlPDFAprovado:  null                       // preenchido na aprovacao
    };
    const itemFieldsRaw = buildFieldsObject(colMap, rawFields);
    // Aplica formatacao por tipo (com log detalhado pra debug)
    const itemFields = {};
    diag.formatLog = {};
    for (const [k, v] of Object.entries(itemFieldsRaw)) {
      const detectedType = (cache.colTypes && cache.colTypes[k]) || 'text';
      const formatted = formatByType(k, v);
      diag.formatLog[k] = {
        colType: detectedType,
        inputValue: v,
        inputType: typeof v,
        outputValue: formatted,
        outputType: typeof formatted
      };
      if (formatted !== undefined) {
        itemFields[k] = formatted;
      }
    }
    diag.itemFieldsUsados = Object.keys(itemFields);
    diag.itemFieldsPayload = itemFields;
    let itemResp;
    try {
      itemResp = await client
        .api(`/sites/${siteId}/lists/${listNotasId}/items`)
        .post({ fields: itemFields });
      diag.itemId = itemResp.id;
    } catch (createErr) {
      // Re-throw com info mais util no diag
      diag.createListError = {
        message: createErr.message,
        statusCode: createErr.statusCode,
        body: createErr.body,
        code: createErr.code
      };
      throw createErr;
    }

    // Dispara notificacao (nao-bloqueante) pro aprovador
    diag.step = 'notify';
    const notifResult = await notificar('lancada', [aprovador.email], {
      itemId: itemResp.id,    // <-- pra gerar links assinados nos botoes
      numero, fornecedor: fornecedorRazao, valor, vencimento,
      unidade, diretoria, aprovador: aprovador.nome,
      submitter: submitterEmail,
      urlPDF: uploadResp.webUrl
    });
    diag.notificado = true;
    diag.notifResult = notifResult;  // <-- expoe detalhe (email, teamsAtividade, teamsWebhook) pra debug

    context.res = {
      status: 200,
      headers: { 'Content-Type': 'application/json' },
      body: {
        ok: true,
        id: itemResp.id,
        title,
        urlPDF: uploadResp.webUrl,
        aprovador,
        diag
      }
    };
  } catch (err) {
    context.log && context.log.error && context.log.error('PostNota error:', err);
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

function ctxErr(context, status, msg) {
  context.res = {
    status,
    headers: { 'Content-Type': 'application/json' },
    body: { error: msg }
  };
}
