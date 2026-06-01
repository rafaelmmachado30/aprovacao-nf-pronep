/**
 * shared/omie.js — Integracao com Omie ERP (REST API v1).
 *
 * Doc: https://developer.omie.com.br/
 * Endpoints usados:
 *   POST https://app.omie.com.br/api/v1/financas/contapagar/   (call=ListarContasPagar)
 *   POST https://app.omie.com.br/api/v1/geral/anexo/           (call=IncluirAnexo)
 *
 * Auth: cada empresa Omie (SP/RJ/ES da Pronep) tem seu proprio par app_key+app_secret.
 * App Settings:
 *   OMIE_APP_KEY_SP / OMIE_APP_SECRET_SP
 *   OMIE_APP_KEY_RJ / OMIE_APP_SECRET_RJ
 *   OMIE_APP_KEY_ES / OMIE_APP_SECRET_ES
 *
 * Rate limit Omie: ~60 req/min por app_key (varia). Helpers nao fazem throttle —
 * caller eh responsavel se chamar em batch.
 */

require('isomorphic-fetch');

const OMIE_BASE = 'https://app.omie.com.br/api/v1';

/**
 * Retorna { appKey, appSecret, empresa } pra unidade dada.
 * Lanca erro se a unidade nao tiver credenciais configuradas.
 */
function getCredentials(unidade) {
  const u = String(unidade || '').toUpperCase();
  let appKey, appSecret, empresa;
  if (u === 'SP') {
    appKey = process.env.OMIE_APP_KEY_SP;
    appSecret = process.env.OMIE_APP_SECRET_SP;
    empresa = 'PRONEP SP';
  } else if (u === 'RJ') {
    appKey = process.env.OMIE_APP_KEY_RJ;
    appSecret = process.env.OMIE_APP_SECRET_RJ;
    empresa = 'PRONEP RJ';
  } else if (u === 'ES') {
    appKey = process.env.OMIE_APP_KEY_ES;
    appSecret = process.env.OMIE_APP_SECRET_ES;
    empresa = 'PRONEP ES';
  } else {
    throw new Error('Unidade nao suportada pra Omie: ' + unidade);
  }
  if (!appKey || !appSecret) {
    throw new Error('Credenciais Omie nao configuradas pra unidade ' + u + ' (OMIE_APP_KEY_' + u + ' / OMIE_APP_SECRET_' + u + ')');
  }
  return { appKey, appSecret, empresa };
}

/**
 * Helper REST: monta payload no formato Omie e chama endpoint.
 * Retorna o body parsed (ja JSON).
 */
async function callOmie(endpoint, call, paramObj, creds) {
  const url = OMIE_BASE + endpoint;
  const body = {
    call: call,
    app_key: creds.appKey,
    app_secret: creds.appSecret,
    param: [paramObj]
  };
  const resp = await fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json; charset=utf-8' },
    body: JSON.stringify(body)
  });
  const text = await resp.text();
  let data;
  try { data = JSON.parse(text); }
  catch (e) {
    throw new Error('Omie retornou resposta nao-JSON (status ' + resp.status + '): ' + text.slice(0, 200));
  }
  // Omie retorna erros com campo "faultstring" mesmo com HTTP 200 as vezes
  if (data && data.faultstring) {
    const err = new Error('Omie erro: ' + data.faultstring + ' (' + (data.faultcode || 'sem codigo') + ')');
    err.omieFault = data;
    throw err;
  }
  if (!resp.ok) {
    const err = new Error('Omie HTTP ' + resp.status + ': ' + text.slice(0, 200));
    err.status = resp.status;
    err.body = text;
    throw err;
  }
  return data;
}

/**
 * Normaliza CNPJ/CPF removendo tudo que nao for digito.
 */
function normalizaDoc(doc) {
  return String(doc || '').replace(/\D/g, '');
}

/**
 * Normaliza numero da NF removendo zeros a esquerda e nao-alfanumericos.
 */
function normalizaNumeroNF(num) {
  const limpo = String(num || '').replace(/[^A-Za-z0-9]/g, '');
  if (/^\d+$/.test(limpo)) return limpo.replace(/^0+/, '') || '0';
  return limpo;
}

/**
 * Busca uma conta a pagar no Omie por CNPJ do fornecedor + numero da NF.
 *
 * Retorna { found: true, conta: {...} } ou { found: false, candidatos: N, tentativas: [...] }.
 *
 * Estrategia: pagina pelos contas a pagar do fornecedor e filtra por numero NF.
 * Omie tem o campo "numero_documento" (com parcela tipo "12345/1") e a
 * "nota_fiscal" (so o numero da NF). Tentamos match em ambos.
 */
async function buscarContaPagar(opts, creds) {
  const cnpjAlvo = normalizaDoc(opts.cnpj);
  const numAlvo = normalizaNumeroNF(opts.numero);
  const valorAlvo = Number(opts.valor || 0);
  const diag = { cnpjAlvo, numAlvo, valorAlvo, paginas: 0, totalLidos: 0, candidatos: [] };

  // ListarContasPagar paginado. Filtra por CNPJ no servidor se possivel.
  // Tentamos primeiro com filtro especifico de CNPJ.
  for (let pagina = 1; pagina <= 20; pagina++) {
    diag.paginas++;
    const param = {
      pagina: pagina,
      registros_por_pagina: 50,
      apenas_importado_api: 'N',
      filtrar_por_cnpj: cnpjAlvo  // se Omie nao reconhece, ignora silenciosamente
    };
    let resp;
    try {
      resp = await callOmie('/financas/contapagar/', 'ListarContasPagar', param, creds);
    } catch (e) {
      diag.erroNoListar = e.message;
      break;
    }
    const items = (resp && (resp.conta_pagar_cadastro || resp.contas_pagar_cadastro)) || [];
    diag.totalLidos += items.length;
    if (items.length === 0) break;

    for (const it of items) {
      const itDoc = normalizaDoc(it.cnpj_cpf_fornecedor || it.cnpj_cpf || '');
      const itNum = normalizaNumeroNF(it.numero_documento || '');
      const itNotaFiscal = normalizaNumeroNF(it.numero_documento_fiscal || it.nota_fiscal || '');
      const itValor = Number(it.valor_documento || 0);

      // Match: doc bate E (numero bate em numero_documento OU nota_fiscal)
      const docOk = itDoc && itDoc === cnpjAlvo;
      const numOk = (itNum && itNum === numAlvo) || (itNotaFiscal && itNotaFiscal === numAlvo);
      if (docOk && numOk) {
        diag.candidatos.push({
          codigo_lancamento_omie: it.codigo_lancamento_omie,
          codigo_lancamento_integracao: it.codigo_lancamento_integracao,
          numero_documento: it.numero_documento,
          nota_fiscal: it.numero_documento_fiscal || it.nota_fiscal,
          valor: itValor,
          status: it.status_titulo
        });
      }
    }

    // Se ja achou algum candidato exato, retorna
    if (diag.candidatos.length > 0) {
      const escolhido = diag.candidatos[0];
      return { found: true, conta: escolhido, diag: diag };
    }

    // Paginacao
    const totalPags = (resp && resp.total_de_paginas) || pagina;
    if (pagina >= totalPags) break;
  }

  return { found: false, diag: diag };
}

/**
 * Anexa PDF a uma conta a pagar do Omie.
 * @param opts.codigoLancamento — codigo_lancamento_omie da conta
 * @param opts.nomeArquivo — ex: "NF-12345_FORNECEDOR.pdf"
 * @param opts.pdfBuffer — Buffer com bytes do PDF
 * @param opts.codIntegracao — chave de idempotencia (ex: "PRONEP-NF-{itemId}")
 */
async function anexarPDF(opts, creds) {
  const pdfBase64 = Buffer.isBuffer(opts.pdfBuffer)
    ? opts.pdfBuffer.toString('base64')
    : String(opts.pdfBuffer || '');
  if (!pdfBase64) throw new Error('pdfBuffer vazio');
  if (!opts.codigoLancamento) throw new Error('codigoLancamento obrigatorio');

  const param = {
    cCodIntAnexo: String(opts.codIntegracao || ('PRONEP-' + Date.now())).slice(0, 100),
    nIdAnexo: 0,  // 0 = novo (Omie cria id)
    cTabela: 'conta-pagar',
    nId: Number(opts.codigoLancamento),
    cNomeArquivo: String(opts.nomeArquivo || 'NF.pdf').slice(0, 100),
    cTipoArquivo: 'pdf',
    cArquivo: pdfBase64
  };

  const resp = await callOmie('/geral/anexo/', 'IncluirAnexo', param, creds);
  return resp;
}

module.exports = {
  getCredentials,
  buscarContaPagar,
  anexarPDF,
  normalizaDoc,
  normalizaNumeroNF
};
