/**
 * /api/DetalheNFOmie  (GET ?docId=<id>) — itens da NF-e para o modal do card.
 *
 * Traz do Omie o conteudo da nota: itens com codigo, descricao, NCM, CFOP,
 * quantidade, unitario e total, mais os totais e o transporte. E a tela de
 * Recebimento de NF-e do Omie, dentro do card.
 *
 * DE ONDE VEM: /produtos/recebimentonfe/ call ConsultarRecebimento, por
 * cChaveNFe. Descoberto por sondagem (ver DiagOmieRecebimento) — os candidatos
 * /produtos/xml/, /produtos/nfe/ e /produtos/dfedocsfiscais/ nao existem, e
 * conta a pagar nao guarda anexo.
 *
 * O QUE ESTA API **NAO** DA: o arquivo. Nao ha XML nem link de DANFE na resposta
 * — so o conteudo estruturado. O PDF continua vindo do e-mail ou do anexo manual.
 * Isso nao e limitacao do nosso codigo; e o que o modulo expoe.
 *
 * SO VALE PARA NF-e. NFS-e de servico nao tem chave de acesso e nao passa pelo
 * recebimento, entao nao ha o que consultar — o endpoint devolve
 * `semChave: true` e a tela explica, em vez de mostrar erro.
 *
 * DUAS TRAVAS QUE NAO PODEM CAIR:
 *
 * 1. ESCOPO. Aplica o MESMO podeVer do quadro. Sem isto, um gestor que nao ve o
 *    card na lista poderia ler o conteudo da nota chamando este endpoint direto
 *    com o id — esconder na listagem e nao barrar aqui nao e controle de acesso,
 *    e decoracao.
 *
 * 2. IDENTIDADE. Confere a chave que o Omie devolveu contra a que foi pedida e
 *    RECUSA se divergir. HTTP 200 nao prova que veio a nota certa; se um dia a
 *    API passar a devolver o ultimo recebimento ou um registro vizinho, o card
 *    mostraria os itens de OUTRA nota para quem autoriza o pagamento. Preferimos
 *    nao mostrar nada a mostrar algo plausivel e errado.
 */

require('isomorphic-fetch');
const { getGraphClient, resolveSiteId } = require('../shared/graph');
const { montarEscopo, podeVer } = require('../shared/escopoNF');
const { getCredentials } = require('../shared/omie');
/* Resolucao, escopo, consulta e trava de identidade sairam daqui para o shared
   quando o EspelhoDaNota precisou do mesmo caminho. Ficaram identicas — o que
   mudou e que agora ha uma copia so, e as duas telas erram ou acertam juntas. */
const {
  resolverDocumento, consultarRecebimento, fichaFornecedor, conferirIdentidade
} = require('../shared/recebimentoOmie');

function num(v) { const n = Number(v); return isFinite(n) ? n : 0; }

/* Achata o item do Omie para o que a tela mostra. A resposta traz doze blocos
   por item (CBS, IBS, ICMSST, custo de estoque...) — jogar tudo na tela seria
   despejar a nota fiscal crua no rosto de quem so quer conferir o que comprou. */
function mapearItem(it) {
  const c = (it && it.itensCabec) || {};
  const icms = (it && it.itensICMS) || {};
  const ipi = (it && it.itensIPI) || {};
  const aj = (it && it.itensAjustes) || {};
  return {
    sequencia: num(c.nSequencia),
    codigo: c.cCodigoProduto || '',
    descricao: c.cDescricaoProduto || '',
    ncm: c.cNCM || '',
    ean: c.cEAN || '',
    cfop: c.cCFOP || '',
    unidade: c.cUnidadeNfe || aj.cUnidade || '',
    quantidade: num(c.nQtdeNFe),
    /* Quanto CHEGOU pode diferir do que foi faturado. Quando difere, e
       justamente o que o aprovador precisa ver antes de autorizar. */
    quantidadeRecebida: aj.nQtdeRecebida != null ? num(aj.nQtdeRecebida) : null,
    precoUnitario: num(c.nPrecoUnit),
    desconto: num(c.vDesconto),
    total: num(c.vTotalItem),
    icms: { aliquota: num(icms.nAliq), base: num(icms.nBC), valor: num(icms.nValor),
            cst: icms.cSitTrib || '' },
    ipi: { aliquota: num(ipi.nAliqIPI), valor: num(ipi.nValIPI) }
  };
}

module.exports = async function (context, req) {
  const t0 = Date.now();
  const q = req.query || {};
  const docId = String(q.docId || '');
  const verComo = String(q.verComo || '').trim();

  try {
    if (!docId) throw new Error('docId obrigatorio');

    const client = await getGraphClient();
    const siteId = await resolveSiteId(client);

    const escopo = await montarEscopo(client, siteId, req, verComo);
    if (!escopo.autenticado) {
      context.res = { status: 401, headers: { 'Content-Type': 'application/json' },
        body: { error: 'Nao autenticado' } };
      return;
    }

    const res = await resolverDocumento(client, siteId, docId);
    if (!podeVer(res.card, escopo)) {
      /* 404, nao 403: dizer "existe mas voce nao pode ver" ja entrega que aquela
         NF existe naquele id. */
      context.res = { status: 404, headers: { 'Content-Type': 'application/json' },
        body: { error: 'Documento nao encontrado' } };
      return;
    }

    const unidade = res.unidade;

    /* Cabecalho do fornecedor. Vem ANTES do desvio de "sem chave" porque nao
       depende de NF-e nenhuma: e justamente na NFS-e, que nao tem itens para
       mostrar, que a ficha do fornecedor sustenta a tela sozinha. */
    const fornecedor = await fichaFornecedor(res, unidade);

    const chave = res.chave;
    if (chave.length !== 44) {
      context.res = { status: 200, headers: { 'Content-Type': 'application/json' },
        body: { ok: true, semChave: true, fornecedor: fornecedor,
          timeMs: Date.now() - t0,
          motivo: 'Documento sem chave de acesso — NFS-e de servico ou lancamento ' +
                  'manual. O Omie so guarda itens de NF-e recebida.' } };
      return;
    }

    const d = await consultarRecebimento(chave, getCredentials(unidade));

    const cab = (d && d.cabec) || {};
    const ident = conferirIdentidade(d, chave);
    if (!ident.ok) {
      /* Trava 2. Nao renderiza nada quando a identidade nao fecha. */
      context.res = { status: 502, headers: { 'Content-Type': 'application/json' },
        body: { error: 'O Omie devolveu uma nota diferente da solicitada. ' +
                       'Nada foi exibido por seguranca.',
                chavePedida: chave, chaveDevolvida: ident.chaveDevolvida } };
      return;
    }

    const brutos = d.itensRecebimento;
    const arr = Array.isArray(brutos) ? brutos : (brutos ? [brutos] : []);
    const itens = arr.map(mapearItem).sort(function (a, b) { return a.sequencia - b.sequencia; });
    const tot = d.totais || {};
    const tr = d.transporte || {};

    context.res = { status: 200, headers: { 'Content-Type': 'application/json' },
      body: {
        ok: true,
        chave: chave,
        fornecedor: fornecedor,
        cabec: {
          numeroNF: cab.cNumeroNFe || '', serie: cab.cSerieNFe || '',
          modelo: cab.cModeloNFe || '', emissao: cab.dEmissaoNFe || '',
          emitente: cab.cRazaoSocial || cab.cNome || '', cnpj: cab.cCNPJ_CPF || '',
          /* O total da NF costuma diferir do valor do card, que e UMA parcela.
             A tela precisa nomear os dois, senao parece divergencia. */
          valorTotalNF: num(cab.nValorNFe), etapa: cab.cEtapa || ''
        },
        itens: itens,
        totais: { produtos: num(tot.vTotalProdutos), nota: num(tot.vTotalNFe),
                  baseICMS: num(tot.bcICMS), valorICMS: num(tot.vICMS) },
        transporte: tr.cNomeTransp || tr.cRazaoTransp
          ? { transportadora: tr.cNomeTransp || tr.cRazaoTransp,
              cnpj: tr.cCnpjCpfTransp || '', volumes: tr.nQtdeVolume || '',
              especie: tr.cEspecieVolume || '', pesoBruto: num(tr.nPesoBruto) }
          : null,
        timeMs: Date.now() - t0
      } };
  } catch (err) {
    context.res = { status: 500, headers: { 'Content-Type': 'application/json' },
      body: { error: (err && err.message) || String(err), timeMs: Date.now() - t0 } };
  }
};
