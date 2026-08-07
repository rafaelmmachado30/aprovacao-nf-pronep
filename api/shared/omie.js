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
 * O QUE O lcpListarRequest NAO TEM (tudo sondado, ver DiagOmieFiltros):
 *   filtro de cliente/fornecedor  clientesFiltro, codigo_cliente_fornecedor,
 *                                 filtrar_por_cliente, filtrar_por_fornecedor,
 *                                 filtrar_cliente_fornecedor — todos recusados
 *   filtro por vencimento         filtrar_por_vencimento nao existe
 *   filtro por pagamento          filtrar_por_pagamento nao existe
 *
 * ARMADILHA QUE JA CUSTOU UM BUG: filtrar_por_data_de/ate parece um filtro de
 * vencimento e NAO E — filtra por data de ALTERACAO. Este arquivo dizia o
 * contrario, e as buscas montavam a janela em torno do vencimento. Em conta de
 * vencimento longo a alteracao fica fora da janela, a conta some do resultado, e
 * o erro sai como "conta a pagar nao encontrada no Omie": acusando o Omie de nao
 * ter uma conta que ele tem. So use esses dois campos quando quiser de fato
 * recortar por alteracao, e diga isso no nome da variavel.
 *
 * Sem filtro util, achar UMA conta exige varrer paginas. Por isso o caminho
 * normal do lancamento nao passa por aqui: acharCodigoOmieDaNota le o
 * codigo_lancamento_omie que a sincronizacao ja gravou no SharePoint.
 */

require('isomorphic-fetch');
const crypto = require('crypto');
const zlib = require('zlib');

const OMIE_BASE = 'https://app.omie.com.br/api/v1';

/* JANELA_DIAS_ANTES/DEPOIS foram removidas junto com o filtro de data das buscas.
   Elas descreviam uma janela de vencimento que o Omie nunca aplicou. */
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
 * Busca uma conta a pagar no Omie varrendo o ListarContasPagar.
 *
 * PLANO B. O caminho normal e acharCodigoOmieDaNota, que le o codigo direto do
 * nosso SharePoint, onde a sincronizacao ja gravou. Aqui so chega nota que nao
 * esta sincronizada — e entao nao ha alternativa a varredura, porque o Omie nao
 * tem filtro por vencimento nem por fornecedor (ambos sondados e recusados pelo
 * nome; ver DiagOmieFiltros).
 *
 * Passos:
 *  1. buscarCliente(cnpj) -> codigo_cliente_omie
 *  2. varre as paginas em duas passadas (em aberto, depois todas), casando por
 *     codigo_cliente_fornecedor + numero da NF
 *
 * @param opts.cnpj   — CNPJ do fornecedor
 * @param opts.numero — Numero da NF
 * @param opts.valor  — Valor da NF (informativo, nao usado no match)
 *
 * NAO recebe mais dataVencimento: a janela de data que existia aqui filtrava por
 * data de ALTERACAO e escondia justamente as contas de vencimento longo.
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

  /* SEM JANELA DE DATA — e essa a correcao.
     O codigo antigo montava uma janela em torno do VENCIMENTO e a mandava em
     filtrar_por_data_de/ate, que no Omie filtra por data de ALTERACAO. Numa conta
     de vencimento longo (IPTU em 10 cotas, parcela 010/013) a alteracao foi hoje e
     a janela esta meses a frente: a conta nunca voltava, e o lancamento falhava
     com "conta a pagar nao encontrada" — um erro que acusava o Omie de nao ter a
     conta quando o filtro e que a escondia.
     Filtro por vencimento e por fornecedor foram sondados e NAO existem
     (DiagOmieFiltros). Entao aqui nao ha filtro esperto possivel: varre-se.

     Este caminho e o PLANO B. O plano A e o IntegrarOmie achar o codigo direto no
     nosso proprio SharePoint, onde a sincronizacao ja o gravou. Aqui so chega nota
     que nao esta sincronizada. */
  diag.estrategia = 'varredura sem filtro de data (filtro de vencimento nao existe no Omie)';

  /* Duas passadas. A primeira restringe a EMABERTO — sao ~550 contas no RJ contra
     6.300 no total, entao cabe em 11 paginas e resolve o caso normal, que e anexar
     o PDF numa conta ainda nao paga. A segunda so roda se a primeira nao achou, e
     cobre a conta que ja foi paga antes da integracao. Comecar pela ampla gastaria
     127 paginas sempre. */
  const passadas = [
    ['em aberto', { filtrar_por_status: 'EMABERTO' }],
    ['todas',     {}]
  ];

  for (const [rotuloPassada, filtroExtra] of passadas) {
    let truncou = false;

    for (let pagina = 1; pagina <= MAX_PAGINAS; pagina++) {
      diag.paginas++;
      const param = Object.assign({
        pagina: pagina,
        registros_por_pagina: 50,
        apenas_importado_api: 'N'
      }, filtroExtra);

      let resp;
      try {
        resp = await callOmie('/financas/contapagar/', 'ListarContasPagar', param, creds);
        if (pagina === 1) {
          diag['respostaPag1_' + rotuloPassada] = {
            total_de_paginas: resp && resp.total_de_paginas,
            total_de_registros: resp && resp.total_de_registros,
            registros_retornados: ((resp && (resp.conta_pagar_cadastro || resp.contas_pagar_cadastro)) || []).length
          };
        }
      } catch (e) {
        diag.erroNoListar = '[' + rotuloPassada + '] ' + e.message;
        break;
      }

      const items = (resp && (resp.conta_pagar_cadastro || resp.contas_pagar_cadastro)) || [];
      diag.totalLidos += items.length;
      if (items.length === 0) break;

      if (pagina === 1 && !diag.primeirosDocs.length) {
        diag.primeirosDocs = items.slice(0, 3).map(function (it) {
          return {
            codigo_cliente_fornecedor: it.codigo_cliente_fornecedor,
            codigo_lancamento_omie: it.codigo_lancamento_omie,
            numero_documento: it.numero_documento,
            nota_fiscal: it.numero_documento_fiscal || it.nota_fiscal,
            valor: it.valor_documento
          };
        });
        if (items[0]) diag.primeiroDocCompleto = Object.keys(items[0]);
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
        diag.achadoNaPassada = rotuloPassada;
        return { found: true, conta: diag.candidatos[0], diag: diag };
      }

      const totalPags = (resp && resp.total_de_paginas) || pagina;
      if (pagina >= totalPags) break;
      if (pagina === MAX_PAGINAS && totalPags > MAX_PAGINAS) truncou = true;
    }

    /* Truncou = paramos no teto de paginas com base ainda por ler. Sem esse aviso,
       "nao encontrada" viraria "nao existe", e alguem iria procurar no Omie uma
       conta que estava na pagina 51. */
    if (truncou) {
      diag.truncado = true;
      diag.avisoTruncado = 'A passada "' + rotuloPassada + '" parou em ' + MAX_PAGINAS +
        ' paginas com registros ainda por ler. NAO ENCONTRADA aqui nao significa ' +
        'que a conta nao exista no Omie.';
    }
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
 * Busca conta a pagar pra PF (CPF) — match por valor + janela de data.
 *
 * Diferente da buscarContaPagar (PJ), nao tenta match por numero NF
 * (reembolso/ferias/13o/etc raramente tem NF formal).
 *
 * Estrategia:
 *  1. buscarCliente(cpf) -> codigo_cliente_omie
 *  2. ListarContasPagar filtrado por data
 *  3. Match: codigo_cliente_fornecedor === codCliente
 *           + valor exato (toleranci de 1 centavo)
 *           + dentro da janela
 *  4. Se varios candidatos: escolhe o mais proximo do vencimento
 */
async function buscarContaPagarPF(opts, creds) {
  const cpfAlvo = normalizaDoc(opts.cnpj);  // pode ser CPF ou CNPJ, normalizaDoc trata
  const valorAlvo = Number(opts.valor || 0);
  const diag = {
    cpfAlvo, valorAlvo,
    paginas: 0, totalLidos: 0,
    candidatos: [], todasContasDoCliente: []
  };

  // PASSO 1: resolver codigo_cliente_omie via CPF
  const cli = await buscarCliente(cpfAlvo, creds);
  diag.clienteOmie = cli;
  if (!cli.found) {
    diag.erroNoListar = 'Colaborador com CPF ' + cpfAlvo + ' nao cadastrado no Omie';
    return { found: false, diag: diag };
  }
  const codClienteAlvo = Number(cli.codigo_cliente_omie);

  /* A data de vencimento continua sendo usada — mas para DESEMPATAR, nunca para
     filtrar a consulta. Colaborador com duas contas do mesmo valor no mes e caso
     real; a mais proxima do vencimento e a aposta certa. */
  let dtRef = null;
  if (opts.dataVencimento) {
    const v = opts.dataVencimento;
    if (v instanceof Date) dtRef = v;
    else if (typeof v === 'string') {
      const d = new Date(v);
      if (!isNaN(d.getTime())) dtRef = d;
      else {
        const m = v.match(/^(\d{2})\/(\d{2})\/(\d{4})/);
        if (m) dtRef = new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
      }
    }
  }
  if (!dtRef || isNaN(dtRef.getTime())) {
    dtRef = null;
    diag.dataReferenciaSource = 'ausente — desempate por vencimento fica indisponivel';
  } else {
    diag.dataReferenciaSource = 'dataVencimento_do_SP (so desempate)';
  }

  /* SEM JANELA DE DATA na consulta — mesma correcao do buscarContaPagar.
     filtrar_por_data_de/ate filtra por ALTERACAO, nao por vencimento. A janela
     montada em torno do vencimento escondia a conta de vencimento longo, e o erro
     saia como "colaborador nao tem contas no periodo" — acusando o cadastro de um
     problema que era do filtro. */
  diag.estrategia = 'varredura sem filtro de data; vencimento so desempata';

  const passadas = [
    ['em aberto', { filtrar_por_status: 'EMABERTO' }],
    ['todas',     {}]
  ];

  for (const [rotuloPassada, filtroExtra] of passadas) {
    let truncou = false;

    for (let pagina = 1; pagina <= MAX_PAGINAS; pagina++) {
      diag.paginas++;
      const param = Object.assign({
        pagina: pagina,
        registros_por_pagina: 50,
        apenas_importado_api: 'N'
      }, filtroExtra);

      let resp;
      try {
        resp = await callOmie('/financas/contapagar/', 'ListarContasPagar', param, creds);
        if (pagina === 1) {
          diag['respostaPag1_' + rotuloPassada] = {
            total_de_paginas: resp && resp.total_de_paginas,
            total_de_registros: resp && resp.total_de_registros,
            registros_retornados: ((resp && (resp.conta_pagar_cadastro || resp.contas_pagar_cadastro)) || []).length
          };
        }
      } catch (e) {
        diag.erroNoListar = '[' + rotuloPassada + '] ' + e.message;
        break;
      }
      const items = (resp && (resp.conta_pagar_cadastro || resp.contas_pagar_cadastro)) || [];
      diag.totalLidos += items.length;
      if (items.length === 0) break;

      for (const it of items) {
        const itCodCli = Number(it.codigo_cliente_fornecedor || 0);
        const itValor = Number(it.valor_documento || 0);
        const itVenc = it.data_vencimento || '';
        if (itCodCli && itCodCli === codClienteAlvo) {
          diag.todasContasDoCliente.push({
            codigo_lancamento_omie: it.codigo_lancamento_omie,
            numero_documento: it.numero_documento,
            nota_fiscal: it.numero_documento_fiscal || it.nota_fiscal,
            valor: itValor,
            data_vencimento: itVenc,
            status: it.status_titulo
          });
          // Match por valor (tolerancia 1 centavo pra arredondamento)
          if (Math.abs(itValor - valorAlvo) < 0.01) {
            diag.candidatos.push({
              codigo_lancamento_omie: it.codigo_lancamento_omie,
              numero_documento: it.numero_documento,
              nota_fiscal: it.numero_documento_fiscal || it.nota_fiscal,
              valor: itValor,
              data_vencimento: itVenc,
              status: it.status_titulo
            });
          }
        }
      }

      const totalPags = (resp && resp.total_de_paginas) || pagina;
      if (pagina >= totalPags) break;
      if (pagina === MAX_PAGINAS && totalPags > MAX_PAGINAS) truncou = true;
    }

    if (truncou) {
      diag.truncado = true;
      diag.avisoTruncado = 'A passada "' + rotuloPassada + '" parou em ' + MAX_PAGINAS +
        ' paginas com registros ainda por ler. NAO ENCONTRADA aqui nao significa ' +
        'que a conta nao exista no Omie.';
    }

    /* Achou na passada estreita? Nao varre a base inteira atras de mais nada. */
    if (diag.candidatos.length > 0) { diag.achadoNaPassada = rotuloPassada; break; }
  }

  if (diag.candidatos.length === 0) {
    diag.erroNoListar = diag.todasContasDoCliente.length > 0
      ? ('Colaborador tem ' + diag.todasContasDoCliente.length + ' contas no periodo mas nenhuma com valor R$ ' + valorAlvo.toFixed(2))
      : 'Colaborador nao tem contas a pagar no periodo';
    return { found: false, diag: diag };
  }

  // Se varios candidatos com mesmo valor: escolhe o mais proximo do vencimento
  let escolhido = diag.candidatos[0];
  if (diag.candidatos.length > 1 && dtRef) {
    const dtRefMs = dtRef.getTime();
    let menorDiff = Infinity;
    for (const c of diag.candidatos) {
      const m = String(c.data_vencimento || '').match(/^(\d{2})\/(\d{2})\/(\d{4})/);
      if (!m) continue;
      const ms = new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1])).getTime();
      const diff = Math.abs(ms - dtRefMs);
      if (diff < menorDiff) { menorDiff = diff; escolhido = c; }
    }
    diag.escolhidoCriterio = 'mais_proximo_do_vencimento';
  } else {
    diag.escolhidoCriterio = 'unico_candidato_com_valor';
  }
  diag.totalCandidatos = diag.candidatos.length;

  return { found: true, conta: escolhido, diag: diag };
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

/**
 * Lista contas a pagar por janela de VENCIMENTO, paginando ate o fim.
 *
 * Este e o caminho que alimenta o quadro "NFs a Pagar". Nasceu porque consumir a
 * SEFAZ direto colidia com o proprio Omie: a SEFAZ conta as consultas por CNPJ, e
 * com os dois puxando o mesmo canal veio cStat 656 (consumo indevido) nas tres
 * filiais. Lendo do Omie, cada sistema fica com a sua fonte e ninguem atrapalha
 * ninguem.
 *
 * O ListarContasPagar NAO aceita filtro de fornecedor — so janela de data (ver o
 * comentario no topo deste arquivo). Para um quadro de contas a pagar isso serve:
 * a janela de vencimento e exatamente o recorte que interessa.
 *
 * @param {{de:Date, ate:Date, maxPaginas?:number}} opts
 * @returns {{contas:Array, paginas:number, totalRegistros:number, truncado:boolean}}
 */
async function listarContasPagarPorVencimento(opts, creds) {
  const maxPag = Math.min(opts.maxPaginas || 20, MAX_PAGINAS);
  const contas = [];
  let paginas = 0, totalRegistros = 0, totalPags = 1;

  /* A janela de data e OPCIONAL: filtrar_por_status sozinho ja e o recorte certo
     para "o que esta em aberto", e nao existe filtro por vencimento no Omie
     (medido — ver DiagOmieFiltros). Mandar filtrar_por_data_de sem querer filtrar
     por data restringiria por data de ALTERACAO, que nao e o que se pediu. */
  const base = { pagina: 0, registros_por_pagina: 50, apenas_importado_api: 'N' };
  if (opts.de && opts.ate) {
    base.filtrar_por_data_de = fmtDataOmie(opts.de);
    base.filtrar_por_data_ate = fmtDataOmie(opts.ate);
  }
  Object.assign(base, opts.filtroExtra || {});

  /* RETOMADA POR PAGINA. Sem isto, uma unidade maior que maxPag NUNCA sincroniza
     por inteiro: cada execucao le da pagina 1 e traz sempre os mesmos registros.
     SP tem 2.053 contas em aberto contra o teto de 1.000 de uma execucao — metade
     ficaria fora para sempre, e o quadro pareceria completo.
     Reler faixas sobrepostas e seguro: a gravacao casa por codigo_lancamento_omie
     e atualiza em vez de duplicar. */
  const inicio = Math.max(1, Number(opts.paginaInicial) || 1);
  let ultimaLida = inicio - 1;

  for (let pagina = inicio; pagina < inicio + maxPag; pagina++) {
    const resp = await callOmie('/financas/contapagar/', 'ListarContasPagar',
      Object.assign({}, base, { pagina: pagina }), creds);

    paginas++;
    ultimaLida = pagina;
    totalRegistros = (resp && resp.total_de_registros) || totalRegistros;
    totalPags = (resp && resp.total_de_paginas) || pagina;
    const items = (resp && (resp.conta_pagar_cadastro || resp.contas_pagar_cadastro)) || [];
    if (!items.length) break;
    contas.push.apply(contas, items);
    if (pagina >= totalPags) break;
  }

  /* truncado avisa que sobrou base por ler. Sem esse sinal, um quadro incompleto
     passa por completo — o pior tipo de erro num painel de contas a pagar.
     proximaPagina diz por onde retomar, em vez de deixar a conta para o usuario. */
  const truncado = ultimaLida < totalPags;
  return { contas, paginas, totalRegistros, totalPaginas: totalPags,
           ultimaPaginaLida: ultimaLida,
           proximaPagina: truncado ? ultimaLida + 1 : null,
           truncado: truncado };
}

/* Cache em memoria do codigo interno do Omie -> CNPJ. A instancia da Function e
   reaproveitada entre chamadas, entao na pratica cada fornecedor e consultado uma
   vez por instancia. Chave inclui a empresa: o mesmo codigo significa fornecedores
   diferentes em empresas Omie diferentes. */
const _cacheFornecedor = {};

/**
 * Resolve codigo_cliente_fornecedor -> { cnpj, razao }.
 *
 * A conta a pagar do Omie traz so o codigo interno, e o que decide unidade e
 * diretoria no nosso quadro e o CNPJ. Nao existe consulta em lote por codigo, so
 * uma por vez — por isso o LIMITE por execucao.
 *
 * O limite nao perde nada: o CNPJ resolvido e gravado na propria linha e nao
 * precisa ser consultado de novo. Com 3 sincronizacoes por dia, um cadastro de
 * algumas centenas de fornecedores termina de resolver no primeiro dia. A
 * alternativa — resolver tudo de uma vez — estouraria os 45s da Function e nao
 * gravaria nada, que e estritamente pior.
 */
async function resolverFornecedoresPorCodigo(codigos, creds, limite) {
  const out = {};
  let gastos = 0;
  for (const cod of codigos) {
    const ck = creds.empresa + '|' + cod;
    if (_cacheFornecedor[ck]) { out[cod] = _cacheFornecedor[ck]; continue; }
    if (gastos >= (limite || 40)) continue;   /* fica para a proxima execucao */
    gastos++;
    try {
      const resp = await callOmie('/geral/clientes/', 'ConsultarCliente',
        { codigo_cliente_omie: Number(cod) }, creds);
      const r = {
        cnpj: normalizaDoc(resp && resp.cnpj_cpf),
        razao: (resp && (resp.razao_social || resp.nome_fantasia)) || ''
      };
      _cacheFornecedor[ck] = r;
      out[cod] = r;
    } catch (e) {
      /* Fornecedor que nao resolve nao derruba a sincronizacao: a conta entra sem
         CNPJ e aparece no quadro como "fornecedor sem cadastro", que ja tem
         tratamento proprio na tela. */
      out[cod] = { cnpj: '', razao: '', erro: e.message };
    }
  }
  return { mapa: out, consultados: gastos, pendentes: codigos.length - Object.keys(out).length };
}

module.exports = {
  getCredentials,
  buscarCliente,
  listarContasPagarPorVencimento,
  resolverFornecedoresPorCodigo,
  buscarContaPagar,
  buscarContaPagarPF,
  anexarPDF,
  normalizaDoc,
  normalizaNumeroNF,
  fmtDataOmie
};
