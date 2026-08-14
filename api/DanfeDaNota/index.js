/**
 * /api/DanfeDaNota?docId=<id do documento no quadro>
 *
 * Abre a DANFE DE VERDADE — o PDF gerado pelo Omie a partir do XML autorizado,
 * com protocolo de autorizacao de uso. E o mesmo arquivo que o botao "Exibir
 * DANFE do Fornecedor" abre dentro do Omie.
 *
 * &xml=1  devolve o XML autorizado (nfeProc) em vez do PDF.
 *
 * POR QUE ISTO SUBSTITUI O ESPELHO
 * Eu tinha afirmado que o Omie nao guardava o XML e que, sem o protocolo, so daria
 * para montar um espelho sem valor fiscal. Estava errado. A sonda antiga bateu em
 * /produtos/dfedocsfiscais/ e levou 404, e eu li "esse endereco nao existe" como
 * "o Omie nao tem o XML" — afirmacoes completamente diferentes. O servico real e
 * /contador/xml/. Quem achou foi o Rafael, clicando na tela do Omie.
 *
 * O CAMINHO, EM DOIS PASSOS
 *   1. /contador/xml/ ListarDocumentos  cModelo 55, cOperacao 0 (ENTRADA), nChave
 *      -> cXml (nfeProc completo) + nIdNF
 *   2. /produtos/dfedocs/ ObterNfe  nIdNfe
 *      -> cPdf, a URL da DANFE renderizada
 *
 * POR QUE PROXY E NAO REDIRECT — a decisao menos obvia daqui
 * O cPdf e um link curto (click.omie.com/pdfnfe-...) que redireciona para um CDN
 * com URL assinada. Medido: ele abre SEM AUTENTICACAO NENHUMA, e a assinatura
 * vale 24h. Um redirect deixaria essa URL no historico do navegador, pronta para
 * ser copiada e aberta por qualquer um — documento fiscal da empresa atras de um
 * link publico que circula sozinho. Baixando aqui e devolvendo o conteudo, o link
 * nunca sai do servidor. Sao ~13 KB por nota; o custo nao se compara.
 *
 * E POR QUE NADA DISSO E GRAVADO
 * A URL assinada expira em 24h. Guardada, viraria um botao que funciona hoje e
 * quebra amanha — calado, que e o pior jeito de quebrar.
 *
 * NOTA CANCELADA
 * cStatus 10 e cancelada, 20 e denegada. O XML de uma nota cancelada continua
 * valido e bonito, e e exatamente por isso que o aviso aparece na tela: quem
 * abre a DANFE para conferir antes de pagar precisa ver isso primeiro.
 */

require('isomorphic-fetch');
const { getGraphClient, resolveSiteId } = require('../shared/graph');
const { getCredentials } = require('../shared/omie');
const { montarEscopo, podeVer } = require('../shared/escopoNF');
const {
  resolverDocumento, buscarXmlAutorizado, obterPdfDaNfe
} = require('../shared/recebimentoOmie');

function esc(s) {
  return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) {
    return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
  });
}

/* Pagina de aviso, nao de erro tecnico: sempre que possivel aponta o espelho, que
   continua servindo quando o documento original nao esta no Omie. */
function aviso(titulo, msg, docId, tom) {
  const cor = tom === 'alerta' ? '#C62828' : '#1F4E79';
  return '<!DOCTYPE html><html lang="pt-BR"><head><meta charset="UTF-8">' +
    '<meta name="viewport" content="width=device-width,initial-scale=1">' +
    '<title>' + esc(titulo) + '</title><style>' +
    "body{font-family:'Segoe UI',Arial,sans-serif;background:#F4F8FB;margin:0;padding:60px 16px;color:#2C3E50}" +
    '.c{max-width:560px;margin:0 auto;background:#fff;padding:30px;border-radius:10px;' +
    'box-shadow:0 4px 16px rgba(31,78,121,.1)}' +
    'h1{color:' + cor + ';font-size:19px;margin:0 0 12px}p{line-height:1.6}' +
    '.b{display:inline-block;margin-top:16px;background:#1F4E79;color:#fff;padding:9px 18px;' +
    'border-radius:5px;text-decoration:none;font-size:14px}' +
    '</style></head><body><div class="c"><h1>' + esc(titulo) + '</h1><p>' + msg + '</p>' +
    (docId ? '<a class="b" href="/api/EspelhoDaNota?docId=' + encodeURIComponent(docId) +
             '">Ver espelho da nota</a>' : '') +
    '</div></body></html>';
}

module.exports = async function (context, req) {
  const q = req.query || {};
  const docId = String(q.docId || '').trim();
  const querXml = String(q.xml || '') === '1';

  const html = function (status, body) {
    context.res = { status: status,
      headers: { 'Content-Type': 'text/html; charset=utf-8', 'Cache-Control': 'no-store' },
      body: body };
  };

  try {
    if (!docId) {
      return html(400, aviso('Parametro faltando', 'Uso: ?docId=&lt;id do documento&gt;.', null));
    }

    const client = await getGraphClient();
    const siteId = await resolveSiteId(client);

    /* Escopo antes de qualquer chamada ao Omie: nao gasta cota para descobrir
       depois que a pessoa nao podia ver a nota. */
    const escopo = await montarEscopo(client, siteId, req, String(q.verComo || '').trim());
    if (!escopo.autenticado) {
      return html(401, aviso('Nao autenticado', 'Faca login para abrir a nota.', null));
    }

    const res = await resolverDocumento(client, siteId, docId);
    if (!podeVer(res.card, escopo)) {
      /* 404, nao 403 — mesmo criterio do resto: "existe mas voce nao pode ver" ja
         entrega que aquela NF existe naquele id. */
      return html(404, aviso('Documento nao encontrado',
        'Nao ha documento com este id no seu escopo.', null));
    }

    if (res.chave.length !== 44) {
      return html(200, aviso('Nota sem chave de acesso',
        'Este documento nao tem chave de acesso — NFS-e de servico ou lancamento ' +
        'manual. NF-e e o unico tipo que tem DANFE.', docId));
    }

    const creds = getCredentials(res.unidade);

    const doc = await buscarXmlAutorizado(res.chave, creds);
    if (!doc) {
      return html(404, aviso('DANFE nao disponivel',
        'O Omie nao tem o documento fiscal desta nota. Isso acontece quando a NF-e ' +
        'nao passou pelo modulo de recebimento. O espelho, montado com os dados que ' +
        'o Omie guarda, continua disponivel.', docId));
    }

    /* --- XML cru, quando pedido --- */
    if (querXml) {
      context.res = { status: 200,
        headers: { 'Content-Type': 'application/xml; charset=utf-8',
                   'Content-Disposition': 'inline; filename="' + res.chave + '.xml"',
                   'Cache-Control': 'no-store' },
        body: doc.xml };
      return;
    }

    /* --- nota cancelada/denegada: avisa ANTES de mostrar o documento ---
       `forcar` existe porque bloquear de vez seria pior: as vezes a pessoa precisa
       justamente ver a DANFE da nota cancelada para tratar com o fornecedor. O que
       nao pode e ela chegar no documento SEM ter lido o aviso. */
    if (doc.status && doc.status !== '00' && String(q.forcar || '') !== '1') {
      const rotulo = doc.status === '10' ? 'CANCELADA'
                   : doc.status === '20' ? 'DENEGADA' : 'com status ' + esc(doc.status);
      return html(200, aviso('Atencao: nota ' + rotulo,
        'A NF-e ' + esc(String(doc.numero || '')) + ' esta <b>' + rotulo + '</b> na SEFAZ. ' +
        'O documento continua existindo e a DANFE ainda abre, mas <b>esta nota nao ' +
        'deve ser paga</b> sem conferir com o fornecedor.<br><br>' +
        '<a href="/api/DanfeDaNota?docId=' + encodeURIComponent(docId) +
        '&amp;forcar=1">Abrir a DANFE mesmo assim</a>', docId, 'alerta'));
    }

    if (!doc.nIdNF) {
      return html(200, aviso('Sem PDF, mas com XML',
        'O Omie tem o XML autorizado desta nota, mas nao devolveu o id que gera a ' +
        'DANFE em PDF. <a href="/api/DanfeDaNota?docId=' + encodeURIComponent(docId) +
        '&amp;xml=1">Baixar o XML</a>.', docId));
    }

    const pdf = await obterPdfDaNfe(doc.nIdNF, creds);
    if (!pdf || !pdf.url) {
      return html(200, aviso('DANFE nao gerada',
        'O Omie tem o XML desta nota mas nao gerou o PDF agora. ' +
        '<a href="/api/DanfeDaNota?docId=' + encodeURIComponent(docId) +
        '&amp;xml=1">Baixar o XML</a> ou ver o espelho.', docId));
    }
    /* Identidade de novo, no fim: o PDF vem por um id intermediario, e um id
       trocado no meio do caminho entregaria a DANFE de outra nota. */
    if (pdf.chave && pdf.chave !== res.chave) {
      return html(502, aviso('Resposta inconsistente',
        'O Omie devolveu um PDF de chave diferente da pedida. Nada foi exibido de ' +
        'proposito: mostrar isso seria arriscar aprovar pagamento pelo documento ' +
        'errado.', docId, 'alerta'));
    }

    /* --- PROXY, nao redirect. Ver o cabecalho do arquivo. --- */
    const r = await fetch(pdf.url, { redirect: 'follow' });
    if (!r.ok) {
      return html(502, aviso('O PDF nao veio',
        'O Omie gerou o link da DANFE mas ele respondeu ' + r.status + '. ' +
        'Tente de novo em alguns segundos.', docId));
    }
    const buf = Buffer.from(await r.arrayBuffer());
    const tipo = String(r.headers.get('content-type') || '');
    if (!/pdf/i.test(tipo)) {
      /* Sem isto, uma pagina de erro do CDN seria salva como "NF-4469.pdf" e so
         apareceria como arquivo corrompido na mao de quem tentasse abrir. */
      return html(502, aviso('O link nao devolveu um PDF',
        'O endereco da DANFE respondeu ' + esc(tipo || 'sem tipo') + ' em vez de PDF. ' +
        'Nada foi baixado.', docId));
    }

    context.res = { status: 200,
      headers: {
        'Content-Type': 'application/pdf',
        'Content-Disposition': 'inline; filename="NF-' +
          String(doc.numero || res.chave).replace(/[^\w.-]/g, '') + '.pdf"',
        'Content-Length': String(buf.length),
        'Cache-Control': 'no-store'
      },
      body: buf, isRaw: true };
  } catch (e) {
    context.log('DanfeDaNota erro: ' + (e && e.stack || e));
    html(500, aviso('Erro inesperado', esc((e && e.message) || String(e)), docId));
  }
};
