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
const { getCredentials, buscarContaPagar, buscarContaPagarPF, anexarPDF } = require('../shared/omie');
const { getGraphClient } = require('../shared/graph');

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
  // DEBUG: garante que SEMPRE retorna algo, mesmo se algo der erro logo no inicio
  try { context.log && context.log('IntegrarOmie called, method=' + req.method); } catch(e){}
  const diag = { step: 'start', method: req.method, hasBody: !!req.body, bodyKeys: req.body ? Object.keys(req.body) : [] };
  try {
    // Validacao precoce do body — se nao for JSON valido com {id}, retorna 400 com diag
    if (!req.body || typeof req.body !== 'object') {
      context.res = { status: 400, headers: {'Content-Type':'application/json'}, body: { error: 'Body invalido (esperado JSON com {id})', diag, bodyType: typeof req.body, bodyRaw: typeof req.body === 'string' ? req.body.slice(0,200) : null } };
      return;
    }
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
    diag.nf = { numero: f.NumeroNF, fornecedor: f.CNPJFornecedor, valor: f.Valor, unidade: f.Unidade, status: f.Status, dataVencimento: f.DataVencimento, aprovadoEm: f.AprovadoEm };

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
    // Detecta se eh PF (CPF 11 digitos) ou PJ (CNPJ 14 digitos)
    const docDigitos = String(f.CNPJFornecedor || '').replace(/\D/g, '');
    const ehPF = docDigitos.length === 11;
    diag.modoBusca = ehPF ? 'PF (CPF) — match por valor+data' : 'PJ (CNPJ) — match por numero NF';

    const busca = ehPF
      ? await buscarContaPagarPF({
          cnpj: f.CNPJFornecedor,
          valor: f.Valor,
          dataVencimento: f.DataVencimento
        }, creds)
      : await buscarContaPagar({
          cnpj: f.CNPJFornecedor,
          numero: f.NumeroNF,
          valor: f.Valor,
          dataVencimento: f.DataVencimento
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

    diag.step = 'achar_pdf';
    // Procura o PDF aprovado em "Notas Fiscais/Notas Aprovadas/{Unidade}/{AAAA-MM-DD}/".
    // BUG CORRIGIDO: a pasta e nomeada pela data BRT (UTC-3) da aprovacao (ver AprovarNota,
    // que ajusta o fuso), mas AprovadoEm e gravado em UTC. Apos 21h BRT a data UTC "vira" o
    // dia seguinte e a busca caia na pasta errada -> "PDF nao encontrado". Agora convertemos
    // pra BRT, tentamos tambem a data UTC, e como rede de seguranca varremos TODAS as
    // subpastas (datas) da unidade.
    const aprovadoEm = f.AprovadoEm || '';
    const numero = String(f.NumeroNF || '');
    const unidade = String(f.Unidade || '');

    function _fmtData(d) {
      return d.getUTCFullYear() + '-' + String(d.getUTCMonth() + 1).padStart(2, '0') + '-' + String(d.getUTCDate()).padStart(2, '0');
    }
    const datasCandidatas = [];
    if (aprovadoEm) {
      const dUtc = new Date(aprovadoEm);
      if (!isNaN(dUtc.getTime())) {
        const dBrt = new Date(dUtc.getTime() - 3 * 60 * 60 * 1000);
        datasCandidatas.push(_fmtData(dBrt)); // pasta BRT (a correta)
        const utcStr = String(aprovadoEm).substring(0, 10);
        if (datasCandidatas.indexOf(utcStr) < 0) datasCandidatas.push(utcStr); // fallback: data UTC (nomes antigos)
      } else {
        datasCandidatas.push(String(aprovadoEm).substring(0, 10));
      }
    }
    diag.pdfBusca = { aprovadoEm, datasCandidatas, numero, unidade, pastasTentadas: [] };

    const { normalizaNumeroNF } = require('../shared/omie');
    const numAlvoFmt = normalizaNumeroNF(numero);
    diag.pdfBusca.numAlvoFmt = numAlvoFmt;
    function achaPdfNaLista(arquivos) {
      return arquivos.find(function (x) {
        if (!x.name) return false;
        const base = x.name.replace(/\.pdf$/i, '');
        // 1) algum token (separado por _ - espaco) igual ao numero normalizado
        const tokens = base.split(/[_\-\s]+/);
        if (tokens.some(function (t) { return numAlvoFmt && normalizaNumeroNF(t) === numAlvoFmt; })) return true;
        // 2) numero normalizado aparece no nome inteiro normalizado (cobre nomes "colados")
        return numAlvoFmt && normalizaNumeroNF(base).indexOf(numAlvoFmt) >= 0;
      });
    }

    const folderUnidade = 'Notas Fiscais/Notas Aprovadas/' + unidade;
    let pdfMatch = null;
    let arquivosCandidatos = [];

    // 1) Caminho rapido: tenta cada data candidata (BRT primeiro, depois UTC)
    for (const data of datasCandidatas) {
      const folder = folderUnidade + '/' + data;
      try {
        const resp = await client.api('/sites/' + siteId + '/drive/root:/' + folder + ':/children').get();
        const arqs = (resp.value || []).filter(function (x) { return x.file; });
        diag.pdfBusca.pastasTentadas.push({ folder: folder, arquivos: arqs.length });
        const m = achaPdfNaLista(arqs);
        if (m) { pdfMatch = m; arquivosCandidatos = arqs; diag.pdfBusca.folderUsado = folder; break; }
        arquivosCandidatos = arquivosCandidatos.concat(arqs);
      } catch (e) {
        diag.pdfBusca.pastasTentadas.push({ folder: folder, erro: e.message });
      }
    }

    // 2) Rede de seguranca: varre TODAS as subpastas (datas) da unidade
    if (!pdfMatch) {
      try {
        const sub = await client.api('/sites/' + siteId + '/drive/root:/' + folderUnidade + ':/children').get();
        for (const sf of (sub.value || []).filter(function (x) { return x.folder; })) {
          try {
            const filesResp = await client.api('/sites/' + siteId + '/drive/items/' + sf.id + '/children').get();
            const arqs = (filesResp.value || []).filter(function (x) { return x.file; });
            const m = achaPdfNaLista(arqs);
            if (m) { pdfMatch = m; diag.pdfBusca.folderUsado = folderUnidade + '/' + sf.name + ' (varredura)'; break; }
          } catch (e2) {}
        }
        diag.pdfBusca.varreuSubpastas = true;
      } catch (e) {
        diag.pdfBusca.erroVarredura = e.message;
      }
    }

    if (!pdfMatch) {
      diag.pdfBusca.nomesEncontrados = arquivosCandidatos.slice(0, 8).map(function (x) { return x.name; });
      context.res = {
        status: 404,
        body: {
          error: 'PDF aprovado nao encontrado na pasta',
          detail: 'NF ' + numero + ' (aprovada em ' + (aprovadoEm || '?') + '). Procurei nas pastas ' + (datasCandidatas.join(', ') || '?') + ' e varri as subpastas de ' + unidade + '.',
          diag
        }
      };
      return;
    }
    diag.pdfBusca.pdfMatch = { id: pdfMatch.id, name: pdfMatch.name };

    diag.step = 'download_pdf';
    let pdfBuffer = null;
    const downloadUrl = pdfMatch['@microsoft.graph.downloadUrl'];
    if (downloadUrl) {
      try {
        const dlResp = await fetch(downloadUrl);
        if (dlResp.ok) {
          pdfBuffer = Buffer.from(await dlResp.arrayBuffer());
        } else {
          diag.downloadStatus = dlResp.status;
        }
      } catch (e) {
        diag.downloadErr = e.message;
      }
    }
    if (!pdfBuffer || pdfBuffer.length === 0) {
      // Fallback: baixa via Graph API direto
      try {
        const ab = await client.api('/sites/' + siteId + '/drive/items/' + pdfMatch.id + '/content').getStream
          ? await client.api('/sites/' + siteId + '/drive/items/' + pdfMatch.id + '/content').get()
          : null;
        if (ab) pdfBuffer = Buffer.from(ab);
      } catch (e) { diag.fallbackDownloadErr = e.message; }
    }
    if (!pdfBuffer || pdfBuffer.length === 0) {
      context.res = { status: 500, body: { error: 'PDF localizado mas nao foi possivel baixar', diag } };
      return;
    }
    diag.pdfSize = pdfBuffer.length;

    // Limite 10MB pro Omie (estimativa segura)
    if (pdfBuffer.length > 10 * 1024 * 1024) {
      context.res = { status: 400, body: { error: 'PDF excede 10MB — Omie pode rejeitar', size: pdfBuffer.length } };
      return;
    }

    diag.step = 'anexar_omie';
    // Nome do anexo: pra PF usa observacao (mais descritivo), pra PJ usa numero NF
    let nomeArq;
    if (ehPF && f.Observacao) {
      const slug = String(f.Observacao).slice(0, 50).replace(/[^A-Za-z0-9]+/g, '-').replace(/^-|-$/g, '');
      nomeArq = 'PF-' + (docDigitos.slice(-6) || 'XXX') + '_' + (slug || 'reembolso') + '.pdf';
    } else {
      nomeArq = 'NF-' + f.NumeroNF + '_PRONEP.pdf';
    }
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
    // Tambem marca como Processado (integracao com Omie = fim do fluxo financeiro)
    if (colMap['Processado']) patch[colMap['Processado']] = true;
    if (colMap['ProcessadoPor']) patch[colMap['ProcessadoPor']] = user.email || '';
    if (colMap['ProcessadoEm']) patch[colMap['ProcessadoEm']] = new Date().toISOString();
    diag.spPatch = patch;
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
        valor: f.Valor,
        anexoResp: anexoResp,
        codIntegracao: 'PRONEP-NF-' + itemId
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
