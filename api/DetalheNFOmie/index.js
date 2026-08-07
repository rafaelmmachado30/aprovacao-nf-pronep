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
const { resolveListId, LIST_DOCFIS, soDigitos } = require('../shared/documentosFiscais');
const { montarEscopo, podeVer, todosItens } = require('../shared/escopoNF');
const { getCredentials } = require('../shared/omie');

const OMIE_BASE = 'https://app.omie.com.br/api/v1';

/* O cadastro de fornecedores muda raramente e a lista inteira custa segundos.
   Sem cache, abrir cinco cards seguidos releria tudo cinco vezes. TTL curto para
   um cadastro corrigido agora aparecer no proximo card, nao so amanha. */
const TTL_FORN_MS = 5 * 60 * 1000;
let _fornCache = { em: 0, mapa: null };

async function fornecedorPorCnpj(client, siteId) {
  if (_fornCache.mapa && (Date.now() - _fornCache.em) < TTL_FORN_MS) return _fornCache.mapa;
  const mapa = {};
  const idForn = await resolveListId(client, siteId, 'PRONEP-NF-Fornecedores');
  if (idForn) {
    for (const it of await todosItens(client, siteId, idForn)) {
      const f = it.fields || {};
      /* field_2=documento, field_4=unidade, field_5=diretoria, field_7=ativo
         (mesmo mapa de ListarFornecedores — se um mudar, os dois mudam). */
      const doc = soDigitos(f.field_2);
      if (doc.length !== 14 || mapa[doc]) continue;
      mapa[doc] = { unidade: f.field_4 || '', diretoria: f.field_5 || '',
                    ativo: String(f.field_7 || '').toLowerCase() === 'sim' };
    }
  }
  _fornCache = { em: Date.now(), mapa: mapa };
  return mapa;
}

async function consultarRecebimento(chave, creds) {
  const resp = await fetch(OMIE_BASE + '/produtos/recebimentonfe/', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', 'Accept': 'application/json',
               'User-Agent': 'PronepNF/1.0 (Azure SWA Functions)' },
    body: JSON.stringify({ call: 'ConsultarRecebimento', app_key: creds.appKey,
                           app_secret: creds.appSecret, param: [{ cChaveNFe: chave }] })
  });
  const texto = await resp.text();
  let data;
  try { data = JSON.parse(texto); }
  catch (e) { throw new Error('Omie respondeu em formato inesperado (HTTP ' + resp.status + ')'); }
  if (data && data.faultstring) throw new Error(String(data.faultstring).slice(0, 200));
  if (!resp.ok) throw new Error('Omie HTTP ' + resp.status);
  return data;
}

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
    const listId = await resolveListId(client, siteId, LIST_DOCFIS);
    if (!listId) throw new Error('Lista de documentos nao existe');

    const escopo = await montarEscopo(client, siteId, req, verComo);
    if (!escopo.autenticado) {
      context.res = { status: 401, headers: { 'Content-Type': 'application/json' },
        body: { error: 'Nao autenticado' } };
      return;
    }

    const item = await client.api('/sites/' + siteId + '/lists/' + listId + '/items/' + docId)
      .expand('fields').get();
    const f = (item && item.fields) || {};
    const cnpj = soDigitos(f.EmitenteCNPJ);
    const forn = cnpj ? (await fornecedorPorCnpj(client, siteId))[cnpj] : null;

    /* Mesmo formato que podeVer espera no quadro — nao um parecido. */
    const card = {
      unidade: (forn && forn.unidade) || f.UnidadeOmie || '',
      diretoria: (forn && forn.diretoria) || '',
      fornecedorCadastrado: !!forn,
      cadastroIncompleto: !!forn && !(forn.unidade && forn.diretoria)
    };
    if (!podeVer(card, escopo)) {
      /* 404, nao 403: dizer "existe mas voce nao pode ver" ja entrega que aquela
         NF existe naquele id. */
      context.res = { status: 404, headers: { 'Content-Type': 'application/json' },
        body: { error: 'Documento nao encontrado' } };
      return;
    }

    const chave = soDigitos(f.ChaveAcesso);
    if (chave.length !== 44) {
      context.res = { status: 200, headers: { 'Content-Type': 'application/json' },
        body: { ok: true, semChave: true, timeMs: Date.now() - t0,
          motivo: 'Documento sem chave de acesso — NFS-e de servico ou lancamento ' +
                  'manual. O Omie so guarda itens de NF-e recebida.' } };
      return;
    }

    const unidade = String(f.UnidadeOmie || 'RJ').toUpperCase();
    const d = await consultarRecebimento(chave, getCredentials(unidade));

    const cab = (d && d.cabec) || {};
    const voltou = soDigitos(cab.cChaveNFe);
    if (voltou !== chave) {
      /* Trava 2. Nao renderiza nada quando a identidade nao fecha. */
      context.res = { status: 502, headers: { 'Content-Type': 'application/json' },
        body: { error: 'O Omie devolveu uma nota diferente da solicitada. ' +
                       'Nada foi exibido por seguranca.',
                chavePedida: chave, chaveDevolvida: voltou || '(vazia)' } };
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
