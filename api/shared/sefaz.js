/**
 * shared/sefaz.js — cliente do webservice NFeDistribuicaoDFe.
 *
 * Baixa as NF-e emitidas CONTRA os nossos CNPJs. E o unico jeito oficial de saber
 * de uma nota antes de o fornecedor mandar o PDF por e-mail.
 *
 * POR QUE ISTO RODA AQUI E NAO NO SISTEMA NOVO: a SEFAZ exige mTLS com certificado
 * A1, e o Node faz isso nativamente com https.Agent({pfx}). O Deno do Supabase nao
 * faz — la o conector precisa de um processo externo. Aqui roda dentro da propria
 * aplicacao.
 *
 * A SENHA DO CERTIFICADO nunca aparece no codigo nem em log: vem de App Setting e
 * e passada direto ao agente TLS.
 *
 * Cuidados que a NT 2014.002 impoe e que estao implementados:
 *   - a resposta vem em lotes de ATE 50 documentos, cada um gzip+base64
 *   - `ultNSU` e o ponto de retomada; `maxNSU` diz se ainda ha lote
 *   - a janela e de 90 dias: documento nao buscado nesse prazo some para sempre
 *   - consulta repetida sem avancar o NSU e "consumo indevido" e pode bloquear o
 *     CNPJ temporariamente — por isso o ponteiro so avanca depois de gravar
 */

const https = require('https');
const zlib = require('zlib');

const HOST = 'www1.nfe.fazenda.gov.br';
const CAMINHO = '/NFeDistribuicaoDFe/NFeDistribuicaoDFe.asmx';

/* tpAmb: 1 producao, 2 homologacao. cUFAutor: codigo IBGE da UF do autor. */
const UF_IBGE = {
  AC: 12, AL: 27, AP: 16, AM: 13, BA: 29, CE: 23, DF: 53, ES: 32, GO: 52,
  MA: 21, MT: 51, MS: 50, MG: 31, PA: 15, PB: 25, PR: 41, PE: 26, PI: 22,
  RJ: 33, RN: 24, RS: 43, RO: 11, RR: 14, SC: 42, SP: 35, SE: 28, TO: 17
};

function envelope(cnpj, ultNSU, tpAmb, cUFAutor) {
  /* ultNSU com 15 digitos, zero a esquerda — a SEFAZ recusa formato diferente. */
  const nsu = String(ultNSU || 0).padStart(15, '0');
  return '<?xml version="1.0" encoding="utf-8"?>' +
    '<soap12:Envelope xmlns:soap12="http://www.w3.org/2003/05/soap-envelope">' +
    '<soap12:Body>' +
    '<nfeDistDFeInteresse xmlns="http://www.portalfiscal.inf.br/nfe/wsdl/NFeDistribuicaoDFe">' +
    '<nfeDadosMsg>' +
    '<distDFeInt xmlns="http://www.portalfiscal.inf.br/nfe" versao="1.01">' +
    '<tpAmb>' + tpAmb + '</tpAmb>' +
    '<cUFAutor>' + cUFAutor + '</cUFAutor>' +
    '<CNPJ>' + cnpj + '</CNPJ>' +
    '<distNSU><ultNSU>' + nsu + '</ultNSU></distNSU>' +
    '</distDFeInt>' +
    '</nfeDadosMsg>' +
    '</nfeDistDFeInteresse>' +
    '</soap12:Body></soap12:Envelope>';
}

function postSoap(corpo, cert, timeoutMs) {
  return new Promise(function (resolve, reject) {
    const dados = Buffer.from(corpo, 'utf8');
    const req = https.request({
      host: HOST,
      path: CAMINHO,
      method: 'POST',
      /* O certificado do cliente entra aqui — e isto que o Deno nao permite. */
      pfx: cert.pfx,
      passphrase: cert.passphrase,
      headers: {
        'Content-Type': 'application/soap+xml; charset=utf-8',
        'Content-Length': dados.length
      },
      timeout: timeoutMs || 25000
    }, function (res) {
      const partes = [];
      res.on('data', function (d) { partes.push(d); });
      res.on('end', function () {
        resolve({ status: res.statusCode, corpo: Buffer.concat(partes).toString('utf8') });
      });
    });
    req.on('timeout', function () {
      req.destroy(new Error('SEFAZ nao respondeu em ' + (timeoutMs || 25000) + 'ms'));
    });
    req.on('error', reject);
    req.write(dados);
    req.end();
  });
}

/* Extrator simples por tag — evita dependencia de parser XML no runtime das
   Functions. Os campos que interessam nao tem aninhamento ambiguo. */
function tag(xml, nome) {
  const m = new RegExp('<(?:\\w+:)?' + nome + '(?:\\s[^>]*)?>([\\s\\S]*?)</(?:\\w+:)?' + nome + '>')
    .exec(xml || '');
  return m ? m[1].trim() : '';
}

function todos(xml, nome) {
  const re = new RegExp('<(?:\\w+:)?' + nome + '(\\s[^>]*)?>([\\s\\S]*?)</(?:\\w+:)?' + nome + '>', 'g');
  const out = [];
  let m;
  while ((m = re.exec(xml || ''))) out.push({ attrs: m[1] || '', conteudo: m[2] });
  return out;
}

function attr(s, nome) {
  const m = new RegExp(nome + '="([^"]*)"').exec(s || '');
  return m ? m[1] : '';
}

/**
 * Consulta um lote a partir de ultNSU.
 * @returns {{cStat,xMotivo,ultNSU,maxNSU,documentos:[{nsu,schema,xml}]}}
 */
async function consultarLote(opcoes) {
  const { cnpj, ultNSU, cert, uf, ambiente } = opcoes;
  const tpAmb = ambiente === 'homologacao' ? 2 : 1;
  const cUF = UF_IBGE[String(uf || 'RJ').toUpperCase()] || 33;

  const r = await postSoap(envelope(cnpj, ultNSU, tpAmb, cUF), cert, opcoes.timeoutMs);
  if (r.status !== 200) {
    throw new Error('SEFAZ devolveu HTTP ' + r.status + ': ' + (r.corpo || '').slice(0, 300));
  }

  const ret = r.corpo;
  const cStat = tag(ret, 'cStat');
  const xMotivo = tag(ret, 'xMotivo');

  const documentos = [];
  for (const z of todos(ret, 'docZip')) {
    let xml = '';
    try {
      /* Cada docZip vem base64 e comprimido em gzip. */
      xml = zlib.gunzipSync(Buffer.from(z.conteudo, 'base64')).toString('utf8');
    } catch (e) {
      /* Um documento ilegivel nao pode derrubar o lote inteiro — os outros 49
         seriam perdidos junto, e a SEFAZ so guarda 90 dias. */
      documentos.push({ nsu: attr(z.attrs, 'NSU'), schema: attr(z.attrs, 'schema'),
                        erro: 'falha ao descomprimir: ' + e.message });
      continue;
    }
    documentos.push({ nsu: attr(z.attrs, 'NSU'), schema: attr(z.attrs, 'schema'), xml: xml });
  }

  return {
    cStat: cStat,
    xMotivo: xMotivo,
    ultNSU: Number(tag(ret, 'ultNSU') || 0),
    maxNSU: Number(tag(ret, 'maxNSU') || 0),
    documentos: documentos
  };
}

/**
 * Le o que interessa de um documento do lote.
 * O `resNFe` e o RESUMO (o que vem antes da manifestacao do destinatario); o
 * `nfeProc`/`NFe` e o XML completo. Os dois trazem chave, emitente e valor, entao
 * o quadro ja funciona com o resumo.
 * Devolve null para o que nao for NF-e (evento, CT-e, resumo de evento).
 */
function extrairNFe(doc) {
  const xml = doc.xml || '';
  const schema = doc.schema || '';

  if (/^procEventoNFe|^resEvento/.test(schema)) return null;   /* evento, nao e nota */

  const ehResumo = /^resNFe/.test(schema) || /<resNFe/.test(xml);
  const ehCompleto = /^procNFe/.test(schema) || /<nfeProc/.test(xml) || /<infNFe/.test(xml);
  if (!ehResumo && !ehCompleto) return null;

  let chave = '';
  let emitCNPJ = '';
  let emitNome = '';
  let valor = null;
  let numero = '';
  let serie = '';
  let emissao = '';
  let vencimento = '';
  let destCNPJ = '';

  if (ehResumo) {
    chave = tag(xml, 'chNFe');
    emitCNPJ = tag(xml, 'CNPJ');
    emitNome = tag(xml, 'xNome');
    valor = parseFloat(tag(xml, 'vNF') || '0') || null;
    emissao = (tag(xml, 'dhEmi') || '').slice(0, 10);
    destCNPJ = tag(xml, 'CNPJDest') || '';
    /* O resumo nao traz numero e serie separados: derivam da propria chave —
       posicoes 25..34 (numero) e 22..24 (serie). */
    const d = String(chave).replace(/\D/g, '');
    if (d.length === 44) {
      serie = String(parseInt(d.slice(22, 25), 10));
      numero = String(parseInt(d.slice(25, 34), 10));
    }
  } else {
    const inf = /<infNFe[\s\S]*?<\/infNFe>/.exec(xml);
    const bloco = inf ? inf[0] : xml;
    chave = (/(?:Id=")NFe([0-9]{44})/.exec(bloco) || [])[1] || tag(xml, 'chNFe');
    const emit = (/<emit>[\s\S]*?<\/emit>/.exec(bloco) || [])[0] || '';
    const dest = (/<dest>[\s\S]*?<\/dest>/.exec(bloco) || [])[0] || '';
    const ide = (/<ide>[\s\S]*?<\/ide>/.exec(bloco) || [])[0] || '';
    const tot = (/<ICMSTot>[\s\S]*?<\/ICMSTot>/.exec(bloco) || [])[0] || '';
    const dup = (/<dup>[\s\S]*?<\/dup>/.exec(bloco) || [])[0] || '';
    emitCNPJ = tag(emit, 'CNPJ');
    emitNome = tag(emit, 'xNome');
    destCNPJ = tag(dest, 'CNPJ');
    numero = tag(ide, 'nNF');
    serie = tag(ide, 'serie');
    emissao = (tag(ide, 'dhEmi') || tag(ide, 'dEmi') || '').slice(0, 10);
    valor = parseFloat(tag(tot, 'vNF') || '0') || null;
    vencimento = tag(dup, 'dVenc') || '';
  }

  const chaveLimpa = String(chave).replace(/\D/g, '');
  if (chaveLimpa.length !== 44) return null;

  return {
    chaveAcesso: chaveLimpa,
    numeroNF: numero,
    serie: serie,
    emitenteCNPJ: String(emitCNPJ).replace(/\D/g, ''),
    emitenteNome: emitNome,
    valor: valor,
    dataEmissao: emissao || null,
    dataVencimento: vencimento || null,
    cnpjDestino: String(destCNPJ).replace(/\D/g, ''),
    nsu: doc.nsu ? Number(doc.nsu) : null,
    completo: !!ehCompleto,
    xml: xml
  };
}

module.exports = { consultarLote, extrairNFe, UF_IBGE, envelope, tag, todos, attr };
