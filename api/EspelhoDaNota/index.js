/**
 * /api/EspelhoDaNota?docId=<id do documento no quadro>
 *
 * Abre a nota fiscal na tela, no layout da DANFE, a partir do que o Omie tem.
 * Sai HTML pronto para imprimir — o navegador salva em PDF sem biblioteca nenhuma
 * do lado do servidor.
 *
 * POR QUE ISTO EXISTE E DE ONDE VEM O DADO
 * O Omie NAO guarda o XML da NF-e (medido: seis sondas em /produtos/, todas sem
 * campo de xml/danfe/link). Mas o call ConsultarRecebimento devolve o CONTEUDO da
 * nota — itens com NCM e CFOP, ICMS por item, ICMS-ST, totais, transporte. Em 49
 * notas das tres unidades a taxa de resposta foi 100%, zero divergente. Como o
 * Omie nao tem janela de retencao, isto alcanca o backlog inteiro, que e o que a
 * SEFAZ nao faz: la o XML se perde em 90 dias.
 *
 * POR QUE NAO SE CHAMA DANFE
 * A DANFE exige o PROTOCOLO DE AUTORIZACAO DE USO, que so existe no XML
 * autorizado. O Omie nao tem, e nao ha de onde tirar. Chamar isto de DANFE seria
 * prometer valor fiscal que o documento nao tem — serve para conferir e aprovar
 * pagamento, nao para acompanhar mercadoria nem para prova perante o fisco. Por
 * isso a tarja, e por isso o rodape lista o que falta em vez de omitir.
 * (Decisao minha, nao pedida: se preferir outro rotulo, e trocar a tarja.)
 *
 * O QUE NUNCA VIRA ZERO
 * Campo que o Omie nao mandou nunca vira 0,00 — sai como traco (ausencia
 * esperada) ou como "nao informado" (ausencia anomala). Duas das 49 notas medidas
 * vieram sem cSitTrib nenhum, e imprimir 0,00 ali diria "imposto zerado" para
 * quem esta aprovando pagamento: afirmar um numero que ninguem mediu.
 *
 * SEGURANCA E REUSO
 * Entra por docId e passa pelo shared/recebimentoOmie — o MESMO caminho do
 * DetalheNFOmie, que ja resolvia o documento, aplicava escopo, consultava o Omie
 * e conferia a identidade da chave. A primeira versao deste arquivo refez tudo
 * isso do zero, com indice proprio e entrada por chave; duas copias da regra de
 * acesso a documento fiscal divergem, e a que diverge e a que ninguem olha.
 *
 * O que NAO e compartilhado e a formatacao, de proposito: o modal converte
 * ausente em 0 para somar, e aqui ausente PRECISA continuar ausente. Um mapeador
 * comum obrigaria uma das duas telas a mentir.
 */

require('isomorphic-fetch');
const { getGraphClient, resolveSiteId } = require('../shared/graph');
const { getCredentials } = require('../shared/omie');
const { soDigitos } = require('../shared/documentosFiscais');
const { montarEscopo, podeVer } = require('../shared/escopoNF');
const {
  resolverDocumento, consultarRecebimento, fichaFornecedor, conferirIdentidade
} = require('../shared/recebimentoOmie');

function esc(s) {
  return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) {
    return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
  });
}

/* ------------------------------------------------------------------ formato */
/* DOIS marcadores, e a diferenca entre eles e o ponto.
   Na primeira versao tudo que faltava virava um aviso amarelo. Renderizado com os
   payloads reais, uma nota CST 60 saiu com NOVE avisos — IPI, desconto, PIS,
   COFINS, todos ausentes por serem zero — e a unica ausencia que realmente
   importava (o CST que o Omie nao mandou, em duas notas do ES) ficou identica as
   outras oito. Marcador que aparece em tudo nao avisa nada: treina a pessoa a
   ignorar exatamente o caso que ele existe para pegar.
   Entao: traco discreto para ausencia ESPERADA (o Omie so manda o campo quando ha
   valor), aviso amarelo so para ausencia ANOMALA — o campo deveria estar la e nao
   esta. Nenhum dos dois vira 0,00, que e o unico erro grave: afirmar um numero
   que ninguem mediu para quem esta aprovando pagamento. */
const TRACO = '<span class="traco" title="Nao veio do Omie — normalmente porque nao ha valor">&mdash;</span>';
const VAZIO = '<span class="vazio" title="Este campo deveria vir preenchido e veio vazio">nao informado</span>';

/* CSTs em que o ICMS proprio NAO existe: ja foi recolhido por substituicao (10,
   30, 60, 70) ou a operacao e isenta/nao tributada/suspensa (40, 41, 50, 51).
   Em ambos, campo de ICMS vazio e a nota ser assim — nao falta de dado. */
const CST_SEM_ICMS_PROPRIO = ['10', '30', '40', '41', '50', '51', '60', '70'];
const CST_ST = ['10', '30', '60', '70'];

function normCst(v) {
  if (v == null || v === '') return null;
  return String(v).trim().padStart(2, '0');
}

function num(v) {
  if (v == null || v === '') return null;
  const n = Number(v);
  return isFinite(n) ? n : null;
}
/* `ausente` decide QUAL marcador usar quando nao ha valor. O default e o traco:
   o aviso e excecao, e quem chama precisa dizer por que aquela ausencia e anomala. */
function moeda(v, ausente) {
  const n = num(v);
  return n == null ? (ausente || TRACO)
                   : n.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}
function qtd(v, ausente) {
  const n = num(v);
  return n == null ? (ausente || TRACO) : n.toLocaleString('pt-BR', { maximumFractionDigits: 4 });
}
function txt(v, ausente) {
  const s = String(v == null ? '' : v).trim();
  return s ? esc(s) : (ausente || TRACO);
}
/* O Omie manda dd/mm/aaaa. Nao passa por new Date() de proposito: uma nota
   emitida as 22:40 -03:00 vira o dia seguinte se converter, e a data da nota
   e o campo que o financeiro confere contra o boleto. */
function data(v) {
  const s = String(v == null ? '' : v).trim();
  return /^\d{2}\/\d{2}\/\d{4}$/.test(s) ? s : (s ? esc(s) : VAZIO);
}
function chaveFormatada(ch) {
  return String(ch).replace(/(\d{4})(?=\d)/g, '$1 ').trim();
}

/* --------------------------------------------------------------------- HTML */
function pagina(corpo, titulo) {
  return '<!DOCTYPE html><html lang="pt-BR"><head><meta charset="UTF-8">' +
    '<meta name="viewport" content="width=device-width,initial-scale=1">' +
    '<title>' + esc(titulo) + '</title><style>' + CSS + '</style></head><body>' +
    corpo + '</body></html>';
}

const CSS = `
*{box-sizing:border-box}
body{font-family:'Segoe UI',Arial,sans-serif;background:#F4F8FB;margin:0;padding:20px;color:#2C3E50;font-size:12px}
.folha{max-width:1000px;margin:0 auto;background:#fff;padding:18px;box-shadow:0 4px 16px rgba(31,78,121,.12);border-radius:6px}
.tarja{background:#8A4B00;color:#fff;padding:9px 14px;border-radius:4px;margin-bottom:14px;line-height:1.5}
.tarja b{letter-spacing:.5px}
.acoes{max-width:1000px;margin:0 auto 12px;display:flex;gap:8px}
.btn{background:#1F4E79;color:#fff;border:0;padding:8px 16px;border-radius:5px;cursor:pointer;font-size:13px;font-family:inherit}
.btn.sec{background:#647883}
h2{font-size:13px;margin:16px 0 6px;text-transform:uppercase;letter-spacing:.5px;color:#1F4E79;border-bottom:1px solid #D6E4F0;padding-bottom:3px}
table{width:100%;border-collapse:collapse;margin-bottom:4px}
th,td{border:1px solid #C8D6E0;padding:4px 6px;text-align:left;vertical-align:top}
th{background:#EDF3F8;font-weight:600;font-size:11px}
td.n,th.n{text-align:right;white-space:nowrap}
.rot{font-size:10px;color:#647883;text-transform:uppercase;display:block;line-height:1.3}
.val{font-size:12px;font-weight:600}
.topo{display:flex;gap:10px;align-items:stretch;margin-bottom:8px}
.emit{flex:1;border:1px solid #C8D6E0;padding:8px}
.emit .nome{font-size:15px;font-weight:700;margin-bottom:3px}
.ident{width:230px;border:1px solid #C8D6E0;padding:8px;text-align:center}
.ident .tit{font-size:17px;font-weight:700;letter-spacing:1px}
.chave{border:1px solid #C8D6E0;padding:6px 8px;margin-bottom:8px;font-family:'Courier New',monospace;
  font-size:13px;letter-spacing:.5px;word-break:break-all}
.vazio{color:#8A6D00;background:#FFF6DB;padding:0 4px;border-radius:3px;font-style:italic;font-size:11px;font-weight:400}
.traco{color:#A8B6C0}
.nota-st{background:#EDF3F8;border-left:3px solid #1F4E79;padding:7px 10px;margin:4px 0 8px;font-size:11px;line-height:1.5}
.rodape{margin-top:16px;background:#FFF6DB;border-left:3px solid #C8A200;padding:10px 12px;line-height:1.6;font-size:11px}
.rodape b{display:block;margin-bottom:4px}
.rodape li{margin:2px 0}
.wrap{overflow-x:auto}
.erro{max-width:560px;margin:60px auto;background:#fff;padding:28px;border-radius:10px;text-align:center;
  box-shadow:0 4px 16px rgba(31,78,121,.1)}
.erro h1{color:#C62828;font-size:19px;margin:0 0 10px}
@media print{
  body{background:#fff;padding:0;font-size:10px}
  .folha{box-shadow:none;max-width:none;padding:0;border-radius:0}
  .acoes{display:none}
  .tarja{border:2px solid #000;color:#000;background:#fff}
}`;

function erro(titulo, msg) {
  return pagina('<div class="erro"><h1>' + esc(titulo) + '</h1><p>' + esc(msg) + '</p></div>', titulo);
}

function linhaCampo(rotulo, valor) {
  return '<div><span class="rot">' + esc(rotulo) + '</span><span class="val">' + valor + '</span></div>';
}

function montar(d, emit, ctx) {
  const cab = d.cabec || {};
  const tot = d.totais || {};
  const tr = d.transporte || {};
  const brutos = d.itensRecebimento;
  const itens = Array.isArray(brutos) ? brutos : (brutos ? [brutos] : []);
  const chave = soDigitos(cab.cChaveNFe);

  let h = '<div class="acoes">' +
    '<button class="btn" onclick="window.print()">Imprimir / salvar em PDF</button>' +
    '<button class="btn sec" onclick="window.close()">Fechar</button></div>';

  h += '<div class="folha">';

  h += '<div class="tarja"><b>ESPELHO DA NOTA — SEM VALOR FISCAL.</b> ' +
       'Reproduz os dados que o Omie guarda desta NF-e para conferencia e aprovacao de ' +
       'pagamento. Nao substitui a DANFE: nao tem protocolo de autorizacao de uso e nao ' +
       'serve para acompanhar mercadoria nem como prova perante o fisco.</div>';

  /* --- emitente + identificacao ---
     `emit` vem da fichaFornecedor compartilhada, que ja tenta cadastro da Pronep,
     depois Omie, depois o proprio documento. A ficha do cadastro nao guarda IE nem
     logradouro completo; a do Omie guarda. Por isso o endereco pode faltar mesmo
     com ficha presente — e ai o marcador diz de onde ela veio, em vez de deixar a
     pessoa achando que o fornecedor nao tem endereco. */
  const e = emit || {};
  const nomeEmit = e.razao || e.fantasia || cab.cRazaoSocial || cab.cNome || '';
  const endereco = [e.logradouro, e.bairro].filter(Boolean).join(' - ');
  const municipio = [[e.cidade, e.uf].filter(Boolean).join(' / '), e.cep].filter(Boolean).join(' - ');

  h += '<div class="topo"><div class="emit">' +
       '<div class="nome">' + txt(nomeEmit, VAZIO) + '</div>';
  h += linhaCampo('Endereco', endereco ? esc(endereco) : VAZIO);
  h += linhaCampo('Municipio / UF / CEP', municipio ? esc(municipio) : VAZIO);
  h += linhaCampo('Inscricao estadual', txt(e.inscricaoEstadual, VAZIO));
  h += linhaCampo('CNPJ', txt(e.documento || cab.cCNPJ_CPF, VAZIO));
  if (!endereco || !e.inscricaoEstadual) {
    /* Nomear a origem transforma "faltou dado" em "olhe aqui para corrigir". */
    h += '<div class="rot" style="margin-top:5px">Ficha vinda de ' +
         esc(e.origem === 'cadastro' ? 'cadastro de fornecedores da Pronep'
             : e.origem === 'omie' ? 'cadastro do Omie'
             : e.origem === 'documento' ? 'apenas do proprio documento'
             : 'nenhuma fonte') + '</div>';
  }
  h += '</div>';

  h += '<div class="ident"><div class="tit">ESPELHO</div>' +
       '<div class="rot" style="margin:2px 0 6px">Documento auxiliar (sem valor fiscal)</div>' +
       linhaCampo('Numero', txt(cab.cNumeroNFe)) +
       linhaCampo('Serie', txt(cab.cSerieNFe)) +
       linhaCampo('Modelo', txt(cab.cModeloNFe)) +
       linhaCampo('Emissao', data(cab.dEmissaoNFe)) +
       '</div></div>';

  h += '<div class="chave"><span class="rot">Chave de acesso</span>' +
       esc(chaveFormatada(chave)) + '</div>';

  /* --- destinatario --- */
  h += '<h2>Destinatario</h2><table><tr>' +
       '<td>' + linhaCampo('Filial', txt(ctx.empresa)) + '</td>' +
       '<td>' + linhaCampo('CNPJ', txt(ctx.cnpjDestino)) + '</td>' +
       '<td>' + linhaCampo('Unidade / diretoria',
                 txt([ctx.unidade, ctx.diretoria].filter(Boolean).join(' · '))) + '</td>' +
       '</tr></table>';

  /* O CST do primeiro item decide como LER o quadro de imposto da nota inteira.
     Em CST 60 o ICMS proprio vazio nao e buraco: o imposto foi recolhido antes e
     esta nas colunas de ST. Sem dizer isso, a pessoa ve o quadro em branco e
     conclui que a nota veio sem imposto. */
  const cstNota = normCst(((itens[0] || {}).itensICMS || {}).cSitTrib);
  const semIcmsProprio = CST_SEM_ICMS_PROPRIO.indexOf(cstNota) >= 0;
  /* CST ausente e o unico caso em que o ICMS vazio e mesmo anomalo: nao ha como
     saber se a nota tem imposto ou nao. Foi o que apareceu em 2 das 49 medidas. */
  const marcaIcms = (cstNota == null || !semIcmsProprio) ? VAZIO : TRACO;

  h += '<h2>Calculo do imposto</h2>';
  if (semIcmsProprio) {
    h += '<div class="nota-st">' + (CST_ST.indexOf(cstNota) >= 0
      ? 'CST <b>' + esc(cstNota) + '</b> — ICMS recolhido por <b>substituicao tributaria</b>. ' +
        'A nota nao tem ICMS proprio; o imposto aparece nas colunas <b>BC ST</b> e ' +
        '<b>Vlr ST</b> da tabela de produtos.'
      : 'CST <b>' + esc(cstNota) + '</b> — operacao <b>isenta, nao tributada ou com ' +
        'suspensao</b>. A nota nao tem ICMS a destacar.') + '</div>';
  } else if (cstNota == null) {
    h += '<div class="nota-st">O Omie nao informou o CST desta nota, entao nao da para ' +
         'dizer se ela tem ICMS. Os campos abaixo ficam como <span class="vazio">nao ' +
         'informado</span> de proposito — <b>nao leia como zero</b>.</div>';
  }

  h += '<div class="wrap"><table><tr>' +
       '<th class="n">BC do ICMS</th><th class="n">Valor do ICMS</th>' +
       '<th class="n">Valor do IPI</th><th class="n">Descontos</th>' +
       '<th class="n">Total dos produtos</th><th class="n">Total da nota</th></tr><tr>' +
       '<td class="n">' + moeda(tot.bcICMS, marcaIcms) + '</td>' +
       '<td class="n">' + moeda(tot.vICMS, marcaIcms) + '</td>' +
       /* IPI, desconto, PIS, COFINS e desonerado: o Omie so manda quando ha valor.
          Ausencia aqui e o caso comum, e merece traco, nao alarme. */
       '<td class="n">' + moeda(tot.vTotalIPI) + '</td>' +
       '<td class="n">' + moeda(tot.vTotalDescontos) + '</td>' +
       '<td class="n">' + moeda(tot.vTotalProdutos, VAZIO) + '</td>' +
       '<td class="n"><b>' + moeda(tot.vTotalNFe, VAZIO) + '</b></td>' +
       '</tr></table>';
  h += '<table><tr>' +
       '<th class="n">PIS</th><th class="n">COFINS</th><th class="n">ICMS desonerado</th>' +
       '<th class="n">Tributos aprox. (Lei 12.741)</th></tr><tr>' +
       '<td class="n">' + moeda(tot.vTotalPIS) + '</td>' +
       '<td class="n">' + moeda(tot.vTotalCOFINS) + '</td>' +
       '<td class="n">' + moeda(tot.vICMSDesonerado) + '</td>' +
       '<td class="n">' + moeda(tot.vAproxTributos) + '</td>' +
       '</tr></table></div>';

  /* --- transporte --- */
  const temTransp = !!(tr.cNomeTransp || tr.cRazaoTransp);
  h += '<h2>Transportador / volumes</h2>';
  if (temTransp) {
    h += '<div class="wrap"><table><tr>' +
      '<th>Transportador</th><th>CNPJ/CPF</th><th>Frete por conta</th>' +
      '<th class="n">Volumes</th><th>Especie</th><th>Marca</th><th class="n">Peso bruto</th></tr><tr>' +
      '<td>' + txt(tr.cRazaoTransp || tr.cNomeTransp) + '</td>' +
      '<td>' + txt(tr.cCnpjCpfTransp) + '</td>' +
      '<td>' + (tr.cTipoFrete === '0' ? 'Emitente' : tr.cTipoFrete === '1' ? 'Destinatario'
                : tr.cTipoFrete === '9' ? 'Sem frete' : txt(tr.cTipoFrete)) + '</td>' +
      '<td class="n">' + qtd(tr.nQtdeVolume) + '</td>' +
      '<td>' + txt(tr.cEspecieVolume) + '</td>' +
      '<td>' + txt(tr.cMarcaVolume) + '</td>' +
      '<td class="n">' + qtd(tr.nPesoBruto) + '</td>' +
      '</tr></table></div>';
  } else {
    /* Nota sem transportadora e comum (retirada, entrega propria, servico) —
       quase metade da amostra medida. Dizer isso evita procurar defeito. */
    h += '<p class="rot" style="padding:6px 0">Sem transportador informado — comum em ' +
         'retirada, entrega propria ou servico.</p>';
  }

  /* --- itens --- */
  h += '<h2>Produtos e servicos</h2><div class="wrap"><table><tr>' +
       '<th>Codigo</th><th>Descricao</th><th>NCM</th><th>CST</th><th>CFOP</th><th>Un</th>' +
       '<th class="n">Qtd</th><th class="n">Vlr unit.</th><th class="n">Vlr total</th>' +
       '<th class="n">BC ICMS</th><th class="n">Aliq</th><th class="n">Vlr ICMS</th>' +
       '<th class="n">BC ST</th><th class="n">Vlr ST</th></tr>';

  for (const item of itens) {
    const ic = item.itensCabec || {};
    const icms = item.itensICMS || {};
    const st = ((item.itensAjustes || {}).itensSitTribEnt || {}).itensSitTribICMSSTEnt || {};
    const cstItem = normCst(icms.cSitTrib);
    const ehST = CST_ST.indexOf(cstItem) >= 0;
    /* Para CST 60 o ICMS foi retido antes, e o valor que a nota carrega esta em
       nValorStRecAnt — nao em nValorST, que veio zerado em toda a amostra. Cair no
       campo errado mostraria zero onde ha R$ 198,07 de imposto. */
    const bcSt = num(st.nBCSTE) ? st.nBCSTE : icms.nBCStRecAnt;
    const vlSt = num(st.nValorST) ? st.nValorST : icms.nValorStRecAnt;
    /* ICMS vazio so alarma quando o CST diz que deveria haver ICMS. Em CST de ST
       ou isencao o vazio e esperado; sem CST nenhum, nada e esperado. */
    const mIcms = (cstItem == null || CST_SEM_ICMS_PROPRIO.indexOf(cstItem) < 0) ? VAZIO : TRACO;
    /* ST so e cobrado de quem tem ST. Numa nota CST 00 a coluna vazia e correta. */
    const mSt = ehST ? VAZIO : TRACO;
    h += '<tr>' +
      '<td>' + txt(ic.cCodigoProduto) + '</td>' +
      '<td>' + txt(ic.cDescricaoProduto, VAZIO) + '</td>' +
      '<td>' + txt(ic.cNCM, VAZIO) + '</td>' +
      '<td>' + txt(cstItem, VAZIO) + '</td>' +
      '<td>' + txt(ic.cCFOP, VAZIO) + '</td>' +
      '<td>' + txt(ic.cUnidadeNfe) + '</td>' +
      '<td class="n">' + qtd(ic.nQtdeNFe, VAZIO) + '</td>' +
      '<td class="n">' + moeda(ic.nPrecoUnit, VAZIO) + '</td>' +
      '<td class="n">' + moeda(ic.vTotalItem, VAZIO) + '</td>' +
      '<td class="n">' + moeda(icms.nBC, mIcms) + '</td>' +
      '<td class="n">' + (num(icms.nAliq) == null ? mIcms : esc(String(icms.nAliq)) + '%') + '</td>' +
      '<td class="n">' + moeda(icms.nValor, mIcms) + '</td>' +
      '<td class="n">' + moeda(bcSt, mSt) + '</td>' +
      '<td class="n">' + moeda(vlSt, mSt) + '</td>' +
      '</tr>';
  }
  if (!itens.length) {
    h += '<tr><td colspan="14">' + VAZIO + ' — o Omie nao devolveu itens para esta nota.</td></tr>';
  }
  h += '</table></div>';

  /* --- o que falta, dito na cara --- */
  h += '<div class="rodape"><b>O que este espelho nao tem, e por que</b><ul>' +
    '<li><b>Protocolo de autorizacao de uso</b> — so existe no XML autorizado da SEFAZ. ' +
    'E a ausencia dele que impede este documento de ser uma DANFE.</li>' +
    '<li><b>Natureza da operacao</b> — nao vem no recebimento do Omie; o CFOP de cada item ' +
    'e o que mais se aproxima.</li>' +
    '<li><b>Frete, seguro e outras despesas acessorias</b> — nao aparecem nos totais do Omie ' +
    '(conferido nas tres unidades), entao nao entram no calculo acima.</li>' +
    '<li><span class="traco">&mdash;</span> significa que o Omie nao mandou o campo, ' +
    'normalmente porque nao ha valor (IPI, desconto e ST em nota sem eles).</li>' +
    '<li><span class="vazio">nao informado</span> e diferente: o campo <b>deveria</b> vir ' +
    'preenchido e veio vazio. <b>Nem um nem outro e zero</b> — nao leia como imposto zerado.</li>' +
    '</ul></div>';

  h += '</div>';
  return pagina(h, 'Nota ' + (cab.cNumeroNFe || '') + ' — ' + (nomeEmit || 'espelho'));
}

/* ------------------------------------------------------------------ handler */
module.exports = async function (context, req) {
  const html = function (status, body) {
    context.res = { status: status,
      headers: { 'Content-Type': 'text/html; charset=utf-8', 'Cache-Control': 'no-store' },
      body: body };
  };

  try {
    const q = req.query || {};
    const docId = String(q.docId || '').trim();
    if (!docId) {
      return html(400, erro('Parametro faltando',
        'Uso: ?docId=<id do documento no quadro>.'));
    }

    const client = await getGraphClient();
    const siteId = await resolveSiteId(client);

    /* Escopo ANTES de qualquer chamada ao Omie: nao gasta cota para descobrir
       depois que a pessoa nao podia ver a nota. */
    const escopo = await montarEscopo(client, siteId, req, String(q.verComo || '').trim());
    if (!escopo.autenticado) {
      return html(401, erro('Nao autenticado', 'Faca login para abrir a nota.'));
    }

    const res = await resolverDocumento(client, siteId, docId);
    if (!podeVer(res.card, escopo)) {
      /* 404, nao 403 — mesmo criterio do DetalheNFOmie: dizer "existe mas voce
         nao pode ver" ja entrega que aquela NF existe naquele id. */
      return html(404, erro('Documento nao encontrado',
        'Nao ha documento com este id no seu escopo.'));
    }

    if (res.chave.length !== 44) {
      return html(200, erro('Nota sem chave de acesso',
        'Este documento nao tem chave de acesso — NFS-e de servico ou lancamento ' +
        'manual. O Omie so guarda o conteudo de NF-e recebida, entao nao ha espelho ' +
        'a montar. O anexo do fornecedor continua sendo o caminho para este documento.'));
    }

    let creds;
    try { creds = getCredentials(res.unidade); }
    catch (e) { return html(500, erro('Credencial do Omie', e.message)); }

    let receb;
    try {
      receb = await consultarRecebimento(res.chave, creds);
    } catch (e) {
      if (e.naoEncontrado) {
        return html(404, erro('Sem detalhe no Omie',
          'Esta nota nao tem recebimento registrado no Omie, entao nao ha itens nem ' +
          'impostos para mostrar. O anexo enviado pelo fornecedor continua sendo o ' +
          'caminho para este documento.'));
      }
      return html(502, erro('O Omie nao respondeu', e.message || String(e)));
    }

    const id = conferirIdentidade(receb, res.chave);
    if (!id.ok) {
      return html(502, erro('Resposta inconsistente',
        'O Omie devolveu uma nota com chave diferente da pedida. Nada foi exibido de ' +
        'proposito: mostrar isso seria arriscar aprovar pagamento pelo documento errado.'));
    }

    const emit = await fichaFornecedor(res, res.unidade);

    html(200, montar(receb, emit, {
      empresa: creds.empresa,
      cnpjDestino: soDigitos(res.f.CNPJDestino),
      unidade: res.card.unidade,
      diretoria: res.card.diretoria
    }));
  } catch (e) {
    context.log('EspelhoDaNota erro: ' + (e && e.stack || e));
    html(500, erro('Erro inesperado', (e && e.message) || String(e)));
  }
};
