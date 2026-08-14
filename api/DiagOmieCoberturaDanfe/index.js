/**
 * /api/DiagOmieCoberturaDanfe  (GET) — ADMIN. Medicao, nao producao.
 *
 * RESPONDE UMA PERGUNTA SO: dos cards que tem chave de acesso, para quantos o
 * Omie devolve o detalhe da nota (itens + impostos) pelo ConsultarRecebimento?
 *
 * POR QUE ISTO PRECISA SER MEDIDO ANTES DE CONSTRUIR
 * O DiagOmieRecebimento provou que UMA chave responde, com itens completos. Isso
 * prova que o caminho existe — nao prova que ele cobre o quadro. O registro que
 * voltou trazia `cEtapa: "60"`, ou seja, e um recebimento REGISTRADO no modulo de
 * recebimento do Omie. Nem toda conta a pagar passa por esse modulo. Se a taxa de
 * resposta for alta, montar a DANFE a partir do Omie resolve o backlog inteiro
 * (nao ha janela de 90 dias como na SEFAZ). Se for baixa, a SEFAZ sobe de
 * prioridade. Uma sonda de 20 minutos decide semanas de trabalho.
 *
 * O QUE ELA MEDE ALEM DO "RESPONDEU"
 * Responder 200 nao basta: a DANFE precisa de campos especificos. Para cada nota
 * que responde, confere bloco a bloco (itens, impostos, totais, transporte,
 * emitente) e diz em quantas cada bloco veio COMPLETO. Um endpoint que responde
 * sempre mas devolve item sem NCM nao serve, e "taxa de resposta 100%" esconderia
 * exatamente isso.
 *
 * SO LE. Um unico call, ConsultarRecebimento, fixo no codigo — nao ha lista de
 * verbos para sondar aqui, porque a pergunta ja esta respondida. Nada de
 * Incluir/Alterar/Excluir chega perto deste arquivo: ele roda contra a base de
 * producao da Pronep.
 *
 * Query:
 *   ?amostra=8     quantas chaves por unidade (default 8, teto 40)
 *   ?unidade=RJ    so uma unidade — use para amostra maior sem estourar o tempo
 *   ?detalhe=1     lista chave a chave o que respondeu e o que faltou
 */

require('isomorphic-fetch');
const { getGraphClient } = require('../shared/graph');
const { getCredentials } = require('../shared/omie');
const { LIST_DOCFIS, resolveListId, soDigitos } = require('../shared/documentosFiscais');

const OMIE_BASE = 'https://app.omie.com.br/api/v1';
const EP = '/produtos/recebimentonfe/';
const CALL = 'ConsultarRecebimento';

/* O Omie limita ~60 req/min por app_key e derruba rajada com "Consumo redundante".
   400ms e o mesmo espacamento que o DiagOmieRecebimento usa e que se mostrou
   suficiente. O orcamento existe porque o SWA corta a execucao: melhor devolver
   uma amostra menor DIZENDO que foi menor do que morrer no timeout sem resposta.
   38s deixa folga sob o teto do SWA e ainda cabe amostra util. */
const PAUSA_MS = 400;
const ORCAMENTO_MS = 38000;

/* Na primeira medicao o SharePoint comeu mais da metade do orcamento e sobrou
   tempo para 7 de 20 consultas. O culpado era o `expand=fields` sem recorte:
   1895 linhas com TODAS as colunas para usar tres delas. Pedindo so as tres, a
   leitura deixa de disputar tempo com a pergunta que a sonda existe para fazer. */
const CAMPOS_NECESSARIOS = 'ChaveAcesso,UnidadeOmie,Descartado';

function dorme(ms) { return new Promise(function (r) { setTimeout(r, ms); }); }

async function lerDocumentosEnxuto(client, siteId, listId) {
  const out = [];
  let url = '/sites/' + siteId + '/lists/' + listId + '/items' +
            '?$expand=' + encodeURIComponent('fields($select=' + CAMPOS_NECESSARIOS + ')') +
            '&$top=999';
  let p = 0;
  while (url && p < 20) {
    const r = await client.api(url).get();
    out.push.apply(out, r.value || []);
    p++;
    const nl = r['@odata.nextLink'];
    url = nl ? nl.replace('https://graph.microsoft.com/v1.0', '') : null;
  }
  return out;
}

function readClientPrincipal(req) {
  const h = req.headers && req.headers['x-ms-client-principal'];
  if (!h) return null;
  try { return JSON.parse(Buffer.from(h, 'base64').toString('utf-8')); } catch (e) { return null; }
}

async function isAdmin(req) {
  const p = readClientPrincipal(req);
  const roles = (p && p.userRoles) || [];
  if (roles.includes('administrador') || roles.includes('admin')) return true;
  try {
    const { getUser } = require('../shared/auth');
    const user = await getUser(req);
    if (!user) return false;
    const { isAdminEmail } = require('../shared/authz');
    if (isAdminEmail((user.email || '').toLowerCase())) return true;
    const { getUserRoles } = require('../shared/userRoles');
    return ((await getUserRoles(user)) || []).includes('administrador');
  } catch (e) { return false; }
}

/* ---------------------------------------------------------------- amostragem */
/* Espacada, nao aleatoria. Duas razoes: a mesma chamada devolve o mesmo resultado
   (da para repetir a medicao e comparar), e pegar o comeco da lista traria so as
   notas mais antigas ou mais novas — que e justamente onde a cobertura tende a
   ser diferente. Ordenar por chave antes espalha por emitente e por data, porque
   os 6 primeiros digitos da chave sao UF + AAMM. */
/* `pagina` desloca a amostra dentro de cada balde, sem reordenar nada: pagina 1
   pega o primeiro de cada faixa, pagina 2 o segundo, e assim por diante. Serve
   para acumular amostra em rodadas — a latencia do Omie e de ~2,4s por nota, e
   uma execucao so nunca vai fechar 40. A alternativa seria disparar consultas em
   paralelo, mas isto roda contra a producao e o teto e 60/min por app_key:
   arriscar o rate limit para medir mais rapido estragaria a propria medicao. */
function amostrar(chaves, quantas, pagina) {
  const ordenadas = chaves.slice().sort();
  const desloc = Math.max(0, (pagina || 1) - 1);
  if (ordenadas.length <= quantas) return desloc ? [] : ordenadas;
  const passo = ordenadas.length / quantas;
  const out = [];
  for (let i = 0; i < quantas; i++) {
    const idx = Math.floor(i * passo) + desloc;
    /* Nao deixa a pagina vazar para o balde seguinte: repetir chave ja medida
       inflaria a amostra com a mesma nota contada duas vezes. */
    if (idx >= Math.floor((i + 1) * passo) || idx >= ordenadas.length) continue;
    out.push(ordenadas[idx]);
  }
  return out;
}

/* ------------------------------------------------------------ um ConsultarRec */
async function consultar(chave, creds) {
  let resp, texto;
  try {
    resp = await fetch(OMIE_BASE + EP, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'Accept': 'application/json',
                 'User-Agent': 'PronepNF/1.0 (Azure SWA Functions)' },
      body: JSON.stringify({ call: CALL, app_key: creds.appKey,
                             app_secret: creds.appSecret, param: [{ cChaveNFe: chave }] })
    });
    texto = await resp.text();
  } catch (e) {
    return { estado: 'erro', motivo: 'rede: ' + ((e && e.message) || String(e)) };
  }
  let data;
  try { data = JSON.parse(texto); }
  catch (e) { return { estado: 'erro', motivo: 'HTTP ' + resp.status + ' nao-JSON' }; }

  if (data && data.faultstring) {
    const fs = String(data.faultstring);
    /* "Nao encontrado" e uma RESPOSTA, nao uma falha: significa que aquela nota
       nao passou pelo modulo de recebimento. E exatamente o numero que a sonda
       existe para levantar. Misturar com erro de rede/credencial inflaria a
       cobertura ou a esconderia, dependendo do lado. */
    const naoAchou = /n[aã]o (foi )?(encontrad|localizad)|inexistente|nenhum registro/i.test(fs);
    return { estado: naoAchou ? 'sem_recebimento' : 'erro', motivo: fs.slice(0, 180) };
  }
  if (!resp.ok) return { estado: 'erro', motivo: 'HTTP ' + resp.status };
  return { estado: 'ok', data: data };
}

/* --------------------------------------------------- completude para a DANFE */
/* Cada bloco e um pedaco obrigatorio da DANFE. Conferir "veio o objeto" nao serve:
   o Omie devolve a estrutura mesmo vazia. O teste e pelo CAMPO que a impressao usa. */
function conferirBlocos(d, chavePedida) {
  const cab = d.cabec || {};
  const tot = d.totais || {};
  const tr = d.transporte || {};
  const brutos = d.itensRecebimento;
  const itens = Array.isArray(brutos) ? brutos : (brutos ? [brutos] : []);

  const chaveVolta = soDigitos(cab.cChaveNFe);
  /* Identidade antes de tudo. Um 200 com a nota ERRADA e o pior resultado
     possivel: montaria a DANFE de outra nota para quem vai aprovar pagamento. */
  if (chaveVolta !== chavePedida) {
    return { divergente: true, chaveDevolvida: chaveVolta || '(vazia)' };
  }

  const temItem = itens.length > 0;
  const it = temItem ? (itens[0].itensCabec || {}) : {};
  const icms = temItem ? (itens[0].itensICMS || {}) : {};
  /* O ST mora fundo: itensAjustes > itensSitTribEnt > itensSitTribICMSSTEnt.
     Ele importa porque na rodada anterior TODAS as notas com imposto "incompleto"
     eram CST 60 — ICMS ja recolhido por substituicao. Concluir "nao falta nada,
     e a natureza da operacao" sem olhar o ST seria deducao, nao medicao: se o
     Omie tambem nao tiver BC e valor do ST, a DANFE dessas notas sai SEM imposto
     nenhum, e imposto errado na tela de quem aprova pagamento e o pior defeito
     possivel aqui. */
  const ajustes = temItem ? (itens[0].itensAjustes || {}) : {};
  const sitTrib = ajustes.itensSitTribEnt || {};
  const st = sitTrib.itensSitTribICMSSTEnt || {};

  /* Registra QUAL campo faltou, nao so que o bloco falhou. Na primeira medicao
     dois registros vieram sem `impostosItem` e sem `totais` e eu nao tinha como
     saber se era campo ausente ou nota sem ICMS — sao coisas opostas: a primeira
     e buraco de dado, a segunda e a nota ser assim mesmo, e a DANFE dela tambem
     nao mostra imposto. Sem este detalhe eu estaria adivinhando. */
  const nulos = [];
  function exige(bloco, obj, campos, teste) {
    let inteiro = true;
    for (const c of campos) {
      const v = obj[c];
      const bom = teste === 'texto' ? !!v : (v != null);
      if (!bom) { nulos.push(bloco + '.' + c); inteiro = false; }
    }
    return inteiro;
  }

  /* `&` e nao `&&`, de proposito. Com `&&` o curto-circuito pula o segundo exige
     assim que o primeiro falha — e camposNulos, que existe justamente para dizer
     TUDO que faltou, sairia contando so o primeiro. Eu leria "faltou o CST"
     quando faltaram quatro campos, e consertaria um buraco achando que era o
     unico. Os dois lados precisam rodar; so o veredito e conjuncao. */
  /* nao basta ter item: a linha da DANFE precisa de descricao, NCM, CFOP,
     unidade, quantidade, unitario e total. */
  const bItens = temItem &&
    (exige('item', it, ['cDescricaoProduto', 'cNCM', 'cCFOP', 'cUnidadeNfe'], 'texto') &
     exige('item', it, ['nQtdeNFe', 'nPrecoUnit', 'vTotalItem'], 'valor')) === 1;
  if (!temItem) nulos.push('itensRecebimento(vazio)');

  /* coluna de imposto por item: CST, aliquota, base e valor do ICMS. */
  const bImpostos = temItem &&
    (exige('icms', icms, ['cSitTrib'], 'texto') &
     exige('icms', icms, ['nAliq', 'nBC', 'nValor'], 'valor')) === 1;

  /* quadro "calculo do imposto" no rodape. */
  const bTotais = exige('totais', tot, ['vTotalNFe', 'vTotalProdutos', 'bcICMS', 'vICMS'], 'valor');

  /* quadro do transportador. Ausencia aqui nao invalida a nota — venda sem
     frete existe — mas muda o que da para imprimir, entao e medido a parte. */
  const bTransporte = !!tr.cNomeTransp || !!tr.cRazaoTransp;
  if (!bTransporte) nulos.push('transporte.cNomeTransp');

  /* identificacao do emitente NO TOPO da DANFE. O Omie da razao social e CNPJ;
     endereco e IE nao vem aqui (ficam no cadastro de fornecedor). */
  const bEmitente = !!(cab.cRazaoSocial || cab.cNome) && !!cab.cCNPJ_CPF;

  /* identificacao do documento: numero, serie, modelo e emissao. */
  const bIdent = exige('cabec', cab, ['cNumeroNFe', 'cSerieNFe', 'dEmissaoNFe'], 'texto');

  /* ---- ICMS ST ----
     CST 10/30/60/70 sao as situacoes com substituicao tributaria. So nelas o
     quadro de ST da DANFE tem o que mostrar; cobrar BC/valor de ST numa nota CST
     00 acusaria falta onde nao ha operacao. Por isso o bloco so e JULGADO quando
     e aplicavel — nas demais fica null, que e diferente de false. */
  const cstTxt = icms.cSitTrib != null ? String(icms.cSitTrib).padStart(2, '0') : null;
  const stAplicavel = ['10', '30', '60', '70'].indexOf(cstTxt) >= 0;
  let bST = null;
  if (stAplicavel) {
    bST = (exige('st', st, ['nBCSTE', 'nValorST'], 'valor')) === true;
  }

  /* Descoberta, nao julgamento: eu nao conheco o shape completo do Omie e nao vou
     inventar. Listar as chaves que REALMENTE vieram deixa o proximo passo apoiado
     no que existe, em vez de num campo que eu supus que existisse. */
  const chavesVistas = {
    totais: Object.keys(tot).sort(),
    st: Object.keys(st).sort(),
    icmsItem: Object.keys(icms).sort()
  };

  return {
    divergente: false,
    qtdItens: itens.length,
    camposNulos: nulos,
    stAplicavel: stAplicavel,
    chavesVistas: chavesVistas,
    /* O CST decide a leitura de um ICMS zerado ou ausente: 40/41/50/51/60 sao
       isencao/nao-tributado/suspensao/ST-ja-recolhida, e nessas a nota REALMENTE
       nao tem ICMS proprio. Sem carregar o CST junto, "impostosItem incompleto"
       seria lido como falha do Omie quando pode ser a natureza da operacao. */
    cstPrimeiroItem: icms.cSitTrib != null ? String(icms.cSitTrib) : null,
    cnpjEmitente: soDigitos(cab.cCNPJ_CPF) || null,
    blocos: {
      itens: bItens,
      impostosItem: bImpostos,
      totais: bTotais,
      transporte: bTransporte,
      emitenteBasico: bEmitente,
      identificacao: bIdent,
      /* null = nao se aplica a esta nota. Contar null como falha rebaixaria toda
         nota tributada normalmente; contar como acerto inflaria a cobertura do
         ST com notas que nunca o tiveram. As duas leituras enganam. */
      impostoST: bST
    },
    /* Existe para a resposta nao parecer que a DANFE sai completa daqui: estes
       campos sao obrigatorios na DANFE e o Omie nao tem nenhum deles. */
    faltaEstruturalNoOmie: ['protocolo de autorizacao de uso',
                            'endereco e IE do emitente',
                            'natureza da operacao'],
    /* nIdFornecedor abre o caminho para fechar endereco/IE por ConsultarFornecedor.
       Registrado aqui porque decide se aquele buraco e contornavel. */
    temIdFornecedor: cab.nIdFornecedor != null
  };
}

/* ------------------------------------------------------------------- handler */
module.exports = async function (context, req) {
  const t0 = Date.now();
  const q = req.query || {};
  const filtroUnidade = String(q.unidade || '').toUpperCase();
  const detalhe = String(q.detalhe || '') === '1';
  let amostra = parseInt(q.amostra, 10);
  if (!isFinite(amostra) || amostra < 1) amostra = 8;
  if (amostra > 40) amostra = 40;
  let pagina = parseInt(q.pagina, 10);
  if (!isFinite(pagina) || pagina < 1) pagina = 1;

  const out = { ok: true, call: CALL, endpoint: EP, amostraPedidaPorUnidade: amostra,
                pagina: pagina, unidades: {}, timeMs: 0 };

  try {
    if (!(await isAdmin(req))) {
      context.res = { status: 403, headers: { 'Content-Type': 'application/json' },
        body: { error: 'Apenas admin' } };
      return;
    }

    /* ---------- de onde saem as chaves: o proprio quadro ---------- */
    const client = await getGraphClient();
    const site = await client.api('/sites/' + process.env.SHAREPOINT_SITE_HOSTNAME +
                                  ':' + process.env.SHAREPOINT_SITE_PATH).get();
    const idDoc = await resolveListId(client, site.id, LIST_DOCFIS);
    if (!idDoc) {
      context.res = { status: 200, headers: { 'Content-Type': 'application/json' },
        body: { ok: true, aviso: 'Lista ' + LIST_DOCFIS + ' nao existe.' } };
      return;
    }

    const docs = await lerDocumentosEnxuto(client, site.id, idDoc);
    /* Medir a leitura a parte foi o que revelou o gargalo da primeira rodada.
       Fica na resposta: se a amostra truncar de novo, este numero diz se a culpa
       e do SharePoint ou da latencia do Omie. */
    out.tempos = { sharePointMs: Date.now() - t0 };

    /* Uma NF parcelada e N linhas com a MESMA chave. Sem deduplicar, a amostra
       gastaria chamadas na mesma nota e a taxa sairia enviesada para os
       fornecedores que parcelam. */
    const porUnidade = {};
    const vistas = {};
    let semChave = 0, descartados = 0;
    for (const d of docs) {
      const f = d.fields || {};
      if (String(f.Descartado || '') === 'Sim') { descartados++; continue; }
      const ch = soDigitos(f.ChaveAcesso);
      if (ch.length !== 44) { semChave++; continue; }
      if (vistas[ch]) continue;
      vistas[ch] = true;
      const u = String(f.UnidadeOmie || '').toUpperCase();
      /* Sem UnidadeOmie nao da para escolher o app_key. Contabiliza e segue —
         chutar a unidade consultaria a empresa errada e voltaria "nao encontrado",
         que seria lido como falta de cobertura. */
      if (u !== 'RJ' && u !== 'SP' && u !== 'ES') continue;
      if (filtroUnidade && u !== filtroUnidade) continue;
      (porUnidade[u] = porUnidade[u] || []).push(ch);
    }
    out.universo = { linhas: docs.length, descartados: descartados, semChave: semChave,
                     notasUnicasComChave: Object.keys(vistas).length };

    /* ---------- consulta ---------- */
    for (const u of Object.keys(porUnidade).sort()) {
      const chaves = porUnidade[u];
      const alvo = amostrar(chaves, amostra, pagina);
      const res = { comChaveNaUnidade: chaves.length, sorteadas: alvo.length,
                    consultadas: 0, ok: 0, semRecebimento: 0, erro: 0, divergente: 0,
                    blocosCompletos: { itens: 0, impostosItem: 0, totais: 0,
                                       transporte: 0, emitenteBasico: 0, identificacao: 0 },
                    itensTotal: 0, comIdFornecedor: 0,
                    /* Agregado das CAUSAS. "impostosItem: 5 de 7" nao diz o que fazer;
                       "icms.nAliq ausente em 2" diz. E o cruzamento CST x emitente
                       responde se o buraco e do Omie ou da natureza da nota. */
                    camposNulosFrequencia: {}, cstPorSituacao: {}, emitentesIncompletos: {},
                    /* ST contado sobre a base certa: so as notas em que ele se
                       aplica. Diluir no total daria um percentual que nao responde
                       pergunta nenhuma. */
                    stAplicaveis: 0, stCompletos: 0,
                    chavesVistas: { totais: {}, st: {}, icmsItem: {} } };
      if (detalhe) res.linhas = [];

      let creds;
      try { creds = getCredentials(u); }
      catch (e) { res.erroCredencial = (e && e.message) || String(e); out.unidades[u] = res; continue; }

      for (const ch of alvo) {
        if (Date.now() - t0 > ORCAMENTO_MS) {
          /* Nao cortar em silencio: uma amostra truncada sem aviso vira "cobertura
             medida" quando e "cobertura medida pela metade". */
          res.interrompidoPorTempo = true;
          break;
        }
        const r = await consultar(ch, creds);
        res.consultadas++;

        if (r.estado === 'sem_recebimento') {
          res.semRecebimento++;
          if (detalhe) res.linhas.push({ chave: ch, estado: 'sem_recebimento' });
        } else if (r.estado === 'erro') {
          res.erro++;
          if (detalhe) res.linhas.push({ chave: ch, estado: 'erro', motivo: r.motivo });
        } else {
          const c = conferirBlocos(r.data, ch);
          if (c.divergente) {
            res.divergente++;
            if (detalhe) res.linhas.push({ chave: ch, estado: 'divergente',
                                           chaveDevolvida: c.chaveDevolvida });
          } else {
            res.ok++;
            res.itensTotal += c.qtdItens;
            if (c.temIdFornecedor) res.comIdFornecedor++;
            for (const k of Object.keys(res.blocosCompletos)) {
              if (c.blocos[k]) res.blocosCompletos[k]++;
            }
            if (c.stAplicavel) {
              res.stAplicaveis++;
              if (c.blocos.impostoST) res.stCompletos++;
            }
            for (const grupo of Object.keys(res.chavesVistas)) {
              for (const campo of c.chavesVistas[grupo]) {
                res.chavesVistas[grupo][campo] = (res.chavesVistas[grupo][campo] || 0) + 1;
              }
            }
            for (const campo of c.camposNulos) {
              res.camposNulosFrequencia[campo] = (res.camposNulosFrequencia[campo] || 0) + 1;
            }
            /* CST agrupado so entre os itens: se todo "incompleto" cair em CST de
               isencao, nao ha buraco nenhum a tapar — a nota e assim. */
            const cst = c.cstPrimeiroItem == null ? '(ausente)' : c.cstPrimeiroItem;
            res.cstPorSituacao[cst] = (res.cstPorSituacao[cst] || 0) + 1;
            /* Emitente so entra na lista quando algo faltou: e ai que se ve se a
               falha se concentra em poucos fornecedores ou esta espalhada. */
            /* `=== false`, nao `!valor`: um bloco null e "nao se aplica a esta
               nota", e !null seria true — o ST apareceria como faltando em toda
               nota tributada normalmente, inventando um defeito inexistente em
               quase toda a amostra. */
            const faltando = Object.keys(c.blocos).filter(function (k) { return c.blocos[k] === false; });
            if (faltando.length && c.cnpjEmitente) {
              res.emitentesIncompletos[c.cnpjEmitente] =
                (res.emitentesIncompletos[c.cnpjEmitente] || 0) + 1;
            }
            if (detalhe) {
              res.linhas.push({ chave: ch, estado: 'ok', itens: c.qtdItens,
                                blocosFaltando: faltando, camposNulos: c.camposNulos,
                                cst: c.cstPrimeiroItem, emitente: c.cnpjEmitente,
                                stAplicavel: c.stAplicavel, stCompleto: c.blocos.impostoST });
            }
          }
        }
        await dorme(PAUSA_MS);
      }

      /* Percentual sobre o que foi DE FATO consultado, nunca sobre o sorteado —
         senao uma interrupcao por tempo apareceria como queda de cobertura. */
      res.taxaResposta = res.consultadas
        ? Math.round((res.ok / res.consultadas) * 100) + '%' : null;
      out.unidades[u] = res;
    }

    /* ---------- consolidado ---------- */
    const soma = { consultadas: 0, ok: 0, semRecebimento: 0, erro: 0, divergente: 0 };
    for (const u of Object.keys(out.unidades)) {
      const r = out.unidades[u];
      for (const k of Object.keys(soma)) soma[k] += r[k] || 0;
    }
    soma.taxaResposta = soma.consultadas
      ? Math.round((soma.ok / soma.consultadas) * 100) + '%' : null;
    out.consolidado = soma;

    out.tempos.omieMs = Date.now() - t0 - out.tempos.sharePointMs;

    out.leitura =
      'taxaResposta = das notas com chave, quantas o Omie detalha. ' +
      'blocosCompletos conta, DENTRE as que responderam, em quantas o bloco veio ' +
      'inteiro — taxa alta com bloco baixo significa endpoint que responde e nao ' +
      'serve. camposNulosFrequencia diz QUAL campo faltou, que e o que se conserta; ' +
      'cruze com cstPorSituacao antes de concluir: CST 40/41/50/51/60 e isencao/ ' +
      'nao-tributado/ST, e nessas a nota nao tem ICMS proprio — bloco incompleto ' +
      'ali e a natureza da operacao, nao buraco do Omie — MAS entao o quadro de ' +
      'ST precisa ter o que mostrar: compare stCompletos com stAplicaveis. Se ' +
      'stCompletos < stAplicaveis, a DANFE dessas notas sai sem imposto nenhum, ' +
      'e ai o caminho tem um furo de verdade. chavesVistas lista os campos que o ' +
      'Omie realmente devolveu, para o proximo passo se apoiar no que existe. ' +
      'emitentesIncompletos ' +
      'mostra se a falha se concentra em poucos fornecedores. divergente > 0 ' +
      'invalida o caminho: seria DANFE de outra nota. faltaEstruturalNoOmie ' +
      '(protocolo, endereco/IE do emitente, natureza da operacao) nao e medida ' +
      'porque o Omie nao tem: o que sai daqui e espelho fiel da nota, nao DANFE ' +
      'com valor fiscal.';

    out.timeMs = Date.now() - t0;
    context.res = { status: 200, headers: { 'Content-Type': 'application/json' }, body: out };
  } catch (err) {
    out.timeMs = Date.now() - t0;
    context.res = { status: 500, headers: { 'Content-Type': 'application/json' },
      body: Object.assign({ error: (err && err.message) || String(err) }, out) };
  }
};
