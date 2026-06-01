/**
 * /api/IntegrarOmie (POST)
 *
 * Integra uma NF aprovada com o Omie ERP:
 *  1. Le NF do SP (numero, CNPJ, valor, unidade, urlPDFAprovado)
 *  2. Seleciona credenciais Omie baseado em Unidade (SP/RJ/ES)
 *  3. Busca conta a pagar correspondente no Omie (match: CNPJ + numero NF)
 *  4. Baixa PDF aprovado do SP
 *  5. Anexa PDF na conta encontrada (call IncluirAnexo)
 *  6. Atualiza colunas IntegradoOmie* no SP
 *  7. Audit log
 *
 * Body JSON: { id: '<spListItemId>' }
 *
 * Apenas Financeiro/Admin pode chamar. Front filtra UI, mas backend tambem checa.
 *
 * Respostas:
 *   200 { ok: true, codigoLancamentoOmie, anexoId, empresa, diag }
 *   400 NF nao aprovada / sem dados pra integrar
 *   403 sem permissao
 *   404 NF nao encontrada (ou conta a pagar nao encontrada no Omie)
 *   500 erro generico
 */

require('isomorphic-fetch');
const { getUser } = require('../shared/auth');
const { getUserRoles } = require('../shared/userRoles');
const { registrar: auditRegistrar } = require('../shared/auditLog');
const { getCredentials, buscarContaPagar, anexarPDF } = require('../shared/omie');
const { ClientSecretCredential } = require('@azure/identity');
const { Client } = require('@microsoft/microsoft-graph-client');
const { TokenCredentialAuthenticationProvider } =
  require('@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials');

const LIST_NOTAS = 'PRONEP-NF-NotasFiscais';
const cache = { siteId: null, driveId: null, listNotasId: null, colMap: null, invColMap: null };

function readClientPrincipal(req) {
  const header = req.headers && req.headers['x-ms-client-principal'];
  if (!header) return null;
  try { return JSON.parse(Buffer.from(header, 'base64').toString('utf-8')); }
  catch (e) { return null; }
}
function readClientPrincipalRoles(req) {
  const p = readClientPrincipal(req);
  return (p && p.userRoles) || [];
}

function getGraphClient() {
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

async function resolveSite(client) {
  if (cache.siteId && cache.driveId && cache.listNotasId) return cache;
  const host = process.env.SHAREPOINT_SITE_HOSTNAME;
  const path = process.env.SHAREPOINT_SITE_PATH;
  const siteResp = await client.api('/sites/' + host + ':' + path).get();
  cache.siteId = siteResp.id;
  const driveResp = await client.api('/sites/' + cache.siteId + '/drive').get();
  cache.driveId = driveResp.id;
  const lists = await client.api('/sites/' + cache.siteId + '/lists').filter("displayName eq '" + LIST_NOTAS + "'").get();
  cache.listNotasId = lists.value[0].id;
  return cache;
}

async function getColMap(client) {
  if (cache.colMap) return cache.colMap;
  const { siteId, listNotasId } = cache;
  const resp = await client.api('/sites/' + siteId + '/lists/' + listNotasId + '/columns').get();
  const map = {}; const inv = {};
  for (const c of (resp.value || [])) {
    if (c.displayName && c.name) { map[c.displayName] = c.name; inv[c.name] = c.displayName; }
  }
  cache.colMap = map; cache.invColMap = inv;
  return map;
}

function normalizeFields(fields, invMap) {
  const out = {};
  for (const [k, v] of Object.entries(fields || {})) {
    out[invMap[k] || k] = v;
  }
  return out;
}

module.exports = async function (context, req) {
  const diag = { step: 'start' };
  try {
    diag.step = 'principal';
    const user = await getUser(req);
    if (!user) {
      context.res = { status: 401, body: { error: 'Nao autenticado' } };
      return;
    }
    const userEmail = (user.email || '').toLowerCase();

    diag.step = 'rbac';
    const claimsRoles = (user.claims && user.claims.roles) || [];
    const principalRoles = readClientPrincipalRoles(req) || [];
    const usefulPrincipalRoles = principalRoles.filter(r => r !== 'authenticated' && r !== 'anonymous');
    const graphRoles = await getUserRoles(user);
    const userRoles = Array.from(new Set([...claimsRoles, ...usefulPrincipalRoles, ...graphRoles]));
    const isAdmin = userRoles.includes('administrador') || userEmail === 'rafael.machado@pronep.com.br';
    const isFinanceiro = userRoles.includes('financeiro_nf');
    if (!isAdmin && !isFinanceiro) {
      context.res = { status: 403, body: { error: 'Apenas Financeiro ou Admin pode integrar com Omie', userRoles } };
      return;
    }
    diag.userScope = { isAdmin, isFinanceiro };

    diag.step = 'parse_body';
    const body = req.body || {};
    const itemId = body.id;
    if (!itemId) {
      context.res = { status: 400, body: { error: 'id obrigatorio no body' } };
      return;
    }
    diag.itemId = itemId;

    diag.step = 'graph';
    const client = getGraphClient();
    const { siteId, driveId, listNotasId } = await resolveSite(client);
    const colMap = await getColMap(client);

    diag.step = 'fetch_item';
    const item = await client.api('/sites/' + siteId + '/lists/' + listNotasId + '/items/' + itemId + '?expand=fields').get();
    const f = normalizeFields(item.fields || {}, cache.invColMap);
    diag.nf = { numero: f.NumeroNF, fornecedor: f.CNPJFornecedor, valor: f.Valor, unidade: f.Unidade, status: f.Status };

    if (f.Status !== 'Aprovada') {
      context.res = { status: 400, body: { error: 'NF nao esta aprovada (status atual: ' + f.Status + ')', diag } };
      return;
    }
    if (f.IntegradoOmie === true || f.IntegradoOmie === 'Sim') {
      context.res = { status: 409, body: { error: 'NF ja foi integrada com Omie anteriormente', diag, ref: f.IntegradoOmieRef || null } };
      return;
    }

    diag.step = 'credenciais';
    const creds = getCredentials(f.Unidade);
    diag.empresa = creds.empresa;

    diag.step = 'buscar_conta';
    const busca = await buscarContaPagar({
      cnpj: f.CNPJFornecedor,
      numero: f.NumeroNF,
      valor: f.Valor
    }, creds);
    diag.buscaOmie = busca.diag;

    if (!busca.found) {
      // Audit log: tentou mas conta nao existe no Omie
      auditRegistrar(user, 'integrar_omie',
        { tipo: 'nf', id: itemId, numero: f.NumeroNF },
        'falha',
        { motivo: 'conta_nao_encontrada_omie', empresa: creds.empresa, paginasLidas: busca.diag.paginas, totalLidos: busca.diag.totalLidos }
      ).catch(function(){});

      context.res = {
        status: 404,
        body: {
          error: 'Conta a pagar nao encontrada no Omie ' + creds.empresa,
          detail: 'Verifique se a conta foi lancada no Omie com CNPJ ' + f.CNPJFornecedor + ' e numero NF ' + f.NumeroNF,
          diag
        }
      };
      return;
    }

    diag.step = 'baixar_pdf';
    // Pega URL do PDF aprovado. Pode estar em UrlPDFAprovado (hyperlink) — extrai o URL
    let urlPDF = '';
    if (f.UrlPDFAprovado) {
      urlPDF = typeof f.UrlPDFAprovado === 'object' ? (f.UrlPDFAprovado.Url || '') : String(f.UrlPDFAprovado);
    }
    if (!urlPDF) {
      // Fallback: tenta encontrar o PDF na pasta Notas Aprovadas pelo numero
      diag.step = 'find_pdf_fallback';
      const folder = 'Notas Aprovadas';
      try {
        const list = await client.api('/sites/' + siteId + '/drive/root:/' + folder + ':/children').get();
        const numero = String(f.NumeroNF || '');
        const pdf = (list.value || []).find(x => x.file && x.name && (x.name.includes('_' + numero + '_') || x.name.startsWith(numero + '_')));
        if (pdf) urlPDF = pdf['@microsoft.graph.downloadUrl'] || pdf.webUrl || '';
      } catch (e) { diag.fallbackErr = e.message; }
    }
    if (!urlPDF) {
      context.res = { status: 400, body: { error: 'PDF aprovado nao encontrado (UrlPDFAprovado vazio e fallback falhou)', diag } };
      return;
    }

    diag.step = 'download_pdf';
    // Se a URL eh um link do SharePoint webUrl, precisa de download direto.
    // Tentamos primeiro chamar Graph pra pegar o downloadUrl do arquivo:
    let pdfBuffer = null;
    try {
      // Se a URL ja eh um @microsoft.graph.downloadUrl, baixa direto (sem auth)
      const dlResp = await fetch(urlPDF);
      if (dlResp.ok) {
        pdfBuffer = Buffer.from(await dlResp.arrayBuffer());
      }
    } catch (e) { diag.downloadErr = e.message; }
    if (!pdfBuffer || pdfBuffer.length === 0) {
      // Tenta resolver o arquivo via Graph (procura por nome na pasta Aprovadas)
      try {
        const list = await client.api('/sites/' + siteId + '/drive/root:/Notas Aprovadas:/children').get();
        const numero = String(f.NumeroNF || '');
        const pdf = (list.value || []).find(x => x.file && x.name && (x.name.includes('_' + numero + '_') || x.name.startsWith(numero + '_')));
        if (pdf) {
          const downloadUrl = pdf['@microsoft.graph.downloadUrl'];
          if (downloadUrl) {
            const dlResp = await fetch(downloadUrl);
            if (dlResp.ok) pdfBuffer = Buffer.from(await dlResp.arrayBuffer());
          }
        }
      } catch (e) { diag.fallbackDownloadErr = e.message; }
    }
    if (!pdfBuffer || pdfBuffer.length === 0) {
      context.res = { status: 400, body: { error: 'Nao foi possivel baixar o PDF aprovado pra anexar no Omie', diag } };
      return;
    }
    diag.pdfSize = pdfBuffer.length;

    // Limite 10MB pro Omie (estimativa segura)
    if (pdfBuffer.length > 10 * 1024 * 1024) {
      context.res = { status: 400, body: { error: 'PDF excede 10MB — Omie pode rejeitar', size: pdfBuffer.length } };
      return;
    }

    diag.step = 'anexar_omie';
    const nomeArq = 'NF-' + f.NumeroNF + '_PRONEP.pdf';
    const codigoLancamento = busca.conta.codigo_lancamento_omie;
    const anexoResp = await anexarPDF({
      codigoLancamento: codigoLancamento,
      nomeArquivo: nomeArq,
      pdfBuffer: pdfBuffer,
      codIntegracao: 'PRONEP-NF-' + itemId
    }, creds);
    diag.anexoResp = anexoResp;

    diag.step = 'update_sp';
    const patch = {};
    if (colMap['IntegradoOmie']) patch[colMap['IntegradoOmie']] = true;
    if (colMap['IntegradoOmieEm']) patch[colMap['IntegradoOmieEm']] = new Date().toISOString();
    if (colMap['IntegradoOmieRef']) patch[colMap['IntegradoOmieRef']] = String(codigoLancamento);
    if (colMap['IntegradoOmieEmpresa']) patch[colMap['IntegradoOmieEmpresa']] = creds.empresa;
    await client.api('/sites/' + siteId + '/lists/' + listNotasId + '/items/' + itemId + '/fields').patch(patch);

    diag.step = 'audit';
    auditRegistrar(user, 'integrar_omie',
      { tipo: 'nf', id: itemId, numero: f.NumeroNF },
      'sucesso',
      {
        empresa: creds.empresa,
        codigoLancamento: codigoLancamento,
        nomeArquivo: nomeArq,
        pdfSize: pdfBuffer.length,
        fornecedor: f.CNPJFornecedor,
        valor: f.Valor
      }
    ).catch(function(){});

    context.res = {
      status: 200,
      headers: { 'Content-Type': 'application/json' },
      body: {
        ok: true,
        codigoLancamentoOmie: codigoLancamento,
        empresa: creds.empresa,
        anexoResp: anexoResp,
        diag
      }
    };
  } catch (err) {
    context.log && context.log.error && context.log.error('IntegrarOmie error:', err);
    // Audit do erro
    try {
      const user = await getUser(req);
      if (user) {
        auditRegistrar(user, 'integrar_omie',
          { tipo: 'nf', id: req.body && req.body.id },
          'falha',
          { motivo: (err && err.message) || String(err), step: diag.step, omieFault: err && err.omieFault }
        ).catch(function(){});
      }
    } catch (e) {}

    context.res = {
      status: 500,
      headers: { 'Content-Type': 'application/json' },
      body: {
        error: (err && err.message) || String(err),
        omieFault: err && err.omieFault,
        diag,
        stack: (err && err.stack || '').split('\n').slice(0, 8)
      }
    };
  }
};
