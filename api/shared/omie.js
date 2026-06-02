/**
 * shared/omie.js — Integracao com Omie ERP (REST API v1).
 *
 * Doc: https://developer.omie.com.br/
 * Endpoints usados:
 *   POST https://app.omie.com.br/api/v1/financas/contapagar/   (ListarContasPagar)
 *   POST https://app.omie.com.br/api/v1/geral/clientes/        (ListarClientes)
 *   POST https://app.omie.com.br/api/v1/geral/anexo/           (IncluirAnexo)
 *
 * Auth: cada empresa Omie (SP/RJ/ES da Pronep) tem seu proprio par app_key+app_secret.
 *
 * App Settings:
 *   OMIE_APP_KEY_SP / OMIE_APP_SECRET_SP
 *   OMIE_APP_KEY_RJ / OMIE_APP_SECRET_RJ
 *   OMIE_APP_KEY_ES / OMIE_APP_SECRET_ES
 *
 * Rate limit Omie: ~60 req/min por app_key.
 *
 * IMPORTANTE: o lcpListarRequest do ListarContasPagar NAO aceita filtro de cliente
 * (clientesFiltro, codigo_cliente_fornecedor, filtrar_por_cnpj, etc — todos rejeitados).
 * Solucao: filtramos por janela de data de vencimento (filtrar_por_data_de e
 * filtrar_por_data_ate) ao redor da data de vencimento da NF.
 */

require('isomorphic-fetch');
const crypto = require('crypto');
const zlib = require('zlib');

const OMIE_BASE = 'https://app.omie.com.br/api/v1';

// Janela em dias antes/depois do vencimento da NF
const JANELA_DIAS_ANTES = 30;
const JANELA_DIAS_DEPOIS = 60;
// Limite de paginas pra evitar timeout SWA (30s)
const MAX_PAGINAS = 50;

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
    throw new Error('Credenciais Omie nao configuradas pra unidade ' + u);
  }
  return { appKey, appSecret, empresa };
}

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
    headers: {
      'Content-Type': 'application/json',
      'Accept': 'application/json',
      'User-Agent': 'PronepNF/1.0 (Azure SWA Functions)'
    },
    body: JSON.stringify(body)
  });
  const text = await resp.text();
  let data;
  try { data = JSON.parse(text); }
  catch (e) {
    throw new Error('Omie retornou resposta nao-JSON (status ' + resp.status + '): ' + text.slice(0, 200));
  }
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

function normalizaDoc(doc) {
  return String(doc || '').replace(/\D/g, '');
}

function normalizaNumeroNF(num) {
  const limpo = String(num || '').replace(/[^A-Za-z0-9]/g, '');
  if (/^\d+$/.test(limpo)) return limpo.replace(/^0+/, '') || '0';
  return limpo;
}

/**
 * Formata Date em DD/MM/AAAA (formato que o Omie usa).
 */
function fmtDataOmie(d) {
  const dia = String(d.getDate()).padStart(2, '0');
  const mes = String(d.getMonth() + 1).padStart(2, '0');
  const ano = d.getFullYear();
  return dia + '/' + mes + '/' + ano;
}

/**
 * Busca cliente por CNPJ usando ListarClientes (ConsultarCliente NAO aceita CNPJ).
 */
async function buscarCliente(cnpj, creds) {
  const cnpjLimpo = normalizaDoc(cnpj);
  try {
    const resp = await callOmie(
      '/geral/clientes/',
      'ListarClientes',
      {
        pagina: 1,
        registros_por_pagina: 5,
        apenas_importado_api: 'N',
        clientesFiltro: { cnpj_cpf: cnpjLimpo }
      },
      creds
    );
    const lista = (resp && (resp.clientes_cadastro || resp.clientes_cadastro_resumido)) || [];
    if (lista.length === 0) {
      return { found: false, error: 'Nenhum cliente com CNPJ ' + cnpjLimpo };
    }
    const exato = lista.find(function (c) {
      return normalizaDoc(c.cnpj_cpf) === cnpjLimpo;
    }) || lista[0];
    return {
      found: true,
      codigo_cliente_omie: exato.codigo_cliente_omie,
      razao: exato.razao_social || exato.nome_fantasia || '',
      totalEncontrados: lista.length
    };
  } catch (e) {
    return { found: false, error: e.message };
  }
}

/**
 * Busca uma conta a pagar no Omie.
 *
 * Estrategia em 2 passos:
 *  1. buscarCliente(cnpj) -> codigo_cliente_omie
 *  2. ListarContasPagar filtrado por janela de data, match por
 *     codigo_cliente_fornecedor + numero NF
 *
 * @param opts.cnpj — CNPJ do fornecedor
 * @param opts.numero — Numero da NF
 * @param opts.valor — Valor da NF (informativo, nao usado pra match)
 * @param opts.dataVencimento — Date | string ISO | DD/MM/AAAA
 */
async function buscarContaPagar(opts, creds) {
  const cnpjAlvo = normalizaDoc(opts.cnpj);
  const numAlvo = normalizaNumeroNF(opts.numero);
  const valorAlvo = Number(opts.valor || 0);
  const diag = {
    cnpjAlvo, numAlvo, valorAlvo,
    paginas: 0, totalLidos: 0,
    candidatos: [], primeirosDocs: [], primeiroDocCompleto: null
  };

  // PASSO 1: resolver codigo_cliente_omie via CNPJ
  const cli = await buscarCliente(cnpjAlvo, creds);
  diag.clienteOmie = cli;
  if (!cli.found) {
    diag.erroNoListar = 'Fornecedor com CNPJ ' + cnpjAlvo + ' nao cadastrado no Omie';
    return { found: false, diag: diag };
  }
  const codClienteAlvo = Number(cli.codigo_cliente_omie);

  // Calcula janela de data ao redor do vencimento
  let dtRef = null;
  if (opts.dataVencimento) {
    const v = opts.dataVencimento;
    if (v instanceof Date) dtRef = v;
    else if (typeof v === 'string') {
      // Tenta ISO primeiro, depois DD/MM/AAAA
      const d = new Date(v);
      if (!isNaN(d.getTime())) dtRef = d;
      else {
        const m = v.match(/^(\d{2})\/(\d{2})\/(\d{4})/);
        if (m) dtRef = new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
      }
    }
  }
  if (!dtRef || isNaN(dtRef.getTime())) {
    // Fallback: usa hoje como referencia
    dtRef = new Date();
    diag.dataReferenciaSource = 'fallback_hoje';
  } else {
    diag.dataReferenciaSource = 'dataVencimento_do_SP';
  }
  const dtDe = new Date(dtRef.getTime() - JANELA_DIAS_ANTES * 86400 * 1000);
  const dtAte = new Date(dtRef.getTime() + JANELA_DIAS_DEPOIS * 86400 * 1000);
  diag.janela = { de: fmtDataOmie(dtDe), ate: fmtDataOmie(dtAte), dias: JANELA_DIAS_ANTES + JANELA_DIAS_DEPOIS };

  for (let pagina = 1; pagina <= MAX_PAGINAS; pagina++) {
    diag.paginas++;
    const param = {
      pagina: pagina,
      registros_por_pagina: 50,
      apenas_importado_api: 'N',
      filtrar_por_data_de: fmtDataOmie(dtDe),
      filtrar_por_data_ate: fmtDataOmie(dtAte)
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

    // Guarda alguns documentos pra debug
    if (pagina === 1) {
      diag.primeirosDocs = items.slice(0, 3).map(function (it) {
        return {
          codigo_cliente_fornecedor: it.codigo_cliente_fornecedor,
          codigo_lancamento_omie: it.codigo_lancamento_omie,
          numero_documento: it.numero_documento,
          nota_fiscal: it.numero_documento_fiscal || it.nota_fiscal,
          valor: it.valor_documento
        };
      });
      // 1a conta completa pra debug
      if (items[0]) {
        diag.primeiroDocCompleto = Object.keys(items[0]);
      }
    }

    for (const it of items) {
      const itCodCli = Number(it.codigo_cliente_fornecedor || 0);
      const itNum = normalizaNumeroNF(it.numero_documento || '');
      const itNotaFiscal = normalizaNumeroNF(it.numero_documento_fiscal || it.nota_fiscal || '');
      const itValor = Number(it.valor_documento || 0);

      const docOk = itCodCli && itCodCli === codClienteAlvo;
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

    if (diag.candidatos.length > 0) {
      return { found: true, conta: diag.candidatos[0], diag: diag };
    }

    const totalPags = (resp && resp.total_de_paginas) || pagina;
    if (pagina >= totalPags) break;
  }

  return { found: false, diag: diag };
}


/**
 * Calcula CRC-32 (IEEE 802.3 / standard ZIP).
 */
function crc32(buf) {
  let crc = 0xFFFFFFFF;
  for (let i = 0; i < buf.length; i++) {
    crc = crc ^ buf[i];
    for (let j = 0; j < 8; j++) {
      crc = (crc >>> 1) ^ (0xEDB88320 & -(crc & 1));
    }
  }
  return (crc ^ 0xFFFFFFFF) >>> 0;
}

/**
 * Cria um arquivo ZIP minimo (formato PKZip standard) contendo um unico arquivo.
 * Retorna Buffer pronto pra usar.
 */
function criarZipSimples(filename, fileBuffer) {
  const filenameBuf = Buffer.from(filename, 'utf-8');
  const fnLen = filenameBuf.length;
  const compressed = zlib.deflateRawSync(fileBuffer);
  const crc = crc32(fileBuffer);
  const uncompSize = fileBuffer.length;
  const compSize = compressed.length;

  // Local File Header (30 bytes + filename)
  const lfh = Buffer.alloc(30);
  lfh.writeUInt32LE(0x04034b50, 0);     // signature
  lfh.writeUInt16LE(20, 4);             // version needed (2.0)
  lfh.writeUInt16LE(0, 6);              // flags
  lfh.writeUInt16LE(8, 8);              // compression method: DEFLATE
  lfh.writeUInt16LE(0, 10);             // mod time
  lfh.writeUInt16LE(0x21, 12);          // mod date (qualquer)
  lfh.writeUInt32LE(crc, 14);           // crc32
  lfh.writeUInt32LE(compSize, 18);      // compressed size
  lfh.writeUInt32LE(uncompSize, 22);    // uncompressed size
  lfh.writeUInt16LE(fnLen, 26);         // filename length
  lfh.writeUInt16LE(0, 28);             // extra length

  // Central Directory File Header (46 bytes + filename)
  const cdh = Buffer.alloc(46);
  cdh.writeUInt32LE(0x02014b50, 0);     // signature
  cdh.writeUInt16LE(20, 4);             // version made by
  cdh.writeUInt16LE(20, 6);             // version needed
  cdh.writeUInt16LE(0, 8);              // flags
  cdh.writeUInt16LE(8, 10);             // method
  cdh.writeUInt16LE(0, 12);             // mod time
  cdh.writeUInt16LE(0x21, 14);          // mod date
  cdh.writeUInt32LE(crc, 16);
  cdh.writeUInt32LE(compSize, 20);
  cdh.writeUInt32LE(uncompSize, 24);
  cdh.writeUInt16LE(fnLen, 28);         // filename len
  cdh.writeUInt16LE(0, 30);             // extra
  cdh.writeUInt16LE(0, 32);             // comment
  cdh.writeUInt16LE(0, 34);             // disk num
  cdh.writeUInt16LE(0, 36);             // internal attrs
  cdh.writeUInt32LE(0, 38);             // external attrs
  cdh.writeUInt32LE(0, 42);             // offset of LFH

  // End of Central Directory (22 bytes)
  const centralDirSize = cdh.length + fnLen;
  const centralDirOffset = lfh.length + fnLen + compSize;
  const eocd = Buffer.alloc(22);
  eocd.writeUInt32LE(0x06054b50, 0);    // signature
  eocd.writeUInt16LE(0, 4);             // disk
  eocd.writeUInt16LE(0, 6);             // disk with cd
  eocd.writeUInt16LE(1, 8);             // entries on this disk
  eocd.writeUInt16LE(1, 10);            // total entries
  eocd.writeUInt32LE(centralDirSize, 12);
  eocd.writeUInt32LE(centralDirOffset, 16);
  eocd.writeUInt16LE(0, 20);            // comment len

  return Buffer.concat([
    lfh, filenameBuf, compressed,
    cdh, filenameBuf,
    eocd
  ]);
}

/**
 * Anexa PDF a uma conta a pagar do Omie via IncluirAnexo.
 */
async function anexarPDF(opts, creds) {
  if (!opts.codigoLancamento) throw new Error('codigoLancamento obrigatorio');
  if (!opts.pdfBuffer) throw new Error('pdfBuffer vazio');
  const pdfBin = Buffer.isBuffer(opts.pdfBuffer) ? opts.pdfBuffer : Buffer.from(String(opts.pdfBuffer), 'base64');
  if (pdfBin.length === 0) throw new Error('pdfBuffer vazio');

  // Omie EXIGE que cArquivo seja um ZIP base64 contendo o arquivo, nao o PDF direto.
  // Status code 6 da resposta: 'arquivo X nao foi encontrado no arquivo zip encaminhado'.
  const fileName = String(opts.nomeArquivo || 'NF.pdf').slice(0, 100);
  const zipBuf = criarZipSimples(fileName, pdfBin);
  const pdfBase64 = zipBuf.toString('base64');

  // cMd5: hash MD5 sobre a STRING base64 do ZIP (Omie calcula sobre o que recebe via JSON).
  const cMd5 = crypto.createHash('md5').update(pdfBase64).digest('hex');

  const param = {
    cCodIntAnexo: String(opts.codIntegracao || ('PRONEP-' + Date.now())).slice(0, 100),
    cTabela: 'conta-pagar',
    nId: Number(opts.codigoLancamento),
    cNomeArquivo: fileName,
    cTipoArquivo: 'pdf',
    cMd5: cMd5,
    cArquivo: pdfBase64
  };

  const resp = await callOmie('/geral/anexo/', 'IncluirAnexo', param, creds);
  // Validacao do status retornado:
  //  '0' = sucesso explicito
  //  '6' = sucesso com warning ('arquivo X nao encontrado no zip' — msg
  //         confusa do Omie, mas anexo eh CRIADO mesmo assim, owner=Integracao)
  // Qualquer outro codigo eh falha real e deve lancar erro
  const status = resp && resp.cCodStatus;
  if (status && status !== '0' && status !== '6') {
    const err = new Error('Omie rejeitou anexo: ' + (resp.cDesStatus || 'cCodStatus=' + status));
    err.omieFault = resp;
    throw err;
  }
  return resp;
}

module.exports = {
  getCredentials,
  buscarCliente,
  buscarContaPagar,
  anexarPDF,
  normalizaDoc,
  normalizaNumeroNF,
  fmtDataOmie
};
