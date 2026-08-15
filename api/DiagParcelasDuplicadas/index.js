/**
 * /api/DiagParcelasDuplicadas  (GET) — ADMIN. Medicao, nao producao.
 *
 * PERGUNTA: quantos cards do quadro somam MAIS do que a nota vale?
 *
 * O caso que levantou isto: NF 14469 da Drogaria GW. A DANFE diz R$ 956,50. O
 * quadro tem DUAS contas a pagar de R$ 956,50 e o card mostra R$ 1.913,00 —
 * exatamente o dobro. E as duas linhas vem do Omie marcadas `001/001`: cada uma
 * se declara PARCELA UNICA. Duas parcelas unicas da mesma nota e contradicao,
 * nao parcelamento.
 *
 * POR QUE ISSO E MAIS GRAVE QUE ERRO DE TELA
 * O card e o que a pessoa aprova. Se ele mostra o dobro, a aprovacao autoriza o
 * dobro. E como a coluna do quadro e derivada do status da NOTA vinculada,
 * aprovar um card leva junto as duas contas — inclusive a que ainda nem venceu.
 *
 * COMO CLASSIFICA, E POR QUE NAO BASTA CONTAR LINHAS
 * Nota parcelada de verdade tambem tem varias linhas, e ela esta CERTA. O que
 * separa uma da outra e o denominador do NumeroParcela, que vem do proprio Omie:
 *
 *   001/003, 002/003, 003/003  -> parcelamento legitimo, soma = valor da nota
 *   001/001 e 001/001          -> SUSPEITO: as duas dizem ser unicas
 *   001/003 mas so 2 linhas    -> incompleto: falta parcela ou o sync perdeu uma
 *
 * A varredura e barata porque so le o SharePoint. Mas classificar nao prova:
 * "suspeito" e uma leitura do rotulo, nao do dinheiro. Por isso uma AMOSTRA vai
 * ao /contador/xml/ buscar o valor REAL da NF e comparar com a soma das linhas.
 * E a diferenca entre dizer "parece duplicado" e mostrar que 956,50 virou 1.913.
 *
 * SO LE. Nao altera, nao descarta, nao desvincula nada — e a decisao do que fazer
 * com os duplicados e do Rafael, nao minha.
 *
 * Query:
 *   ?amostra=8   quantos grupos suspeitos conferir contra a NF real (default 8)
 *   ?unidade=RJ  restringe a uma unidade
 */

require('isomorphic-fetch');
const { getGraphClient, resolveSiteId } = require('../shared/graph');
const { getCredentials } = require('../shared/omie');
const { LIST_DOCFIS, resolveListId, soDigitos } = require('../shared/documentosFiscais');
const { buscarXmlAutorizado } = require('../shared/recebimentoOmie');

const PAUSA_MS = 400;
const ORCAMENTO_MS = 38000;
function dorme(ms) { return new Promise(function (r) { setTimeout(r, ms); }); }

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

/* "001/003" -> {n:1, de:3}. Formato vazio ou estranho devolve null, e null NAO e
   tratado como 1/1: nao saber quantas parcelas existem e diferente de saber que
   ha uma so, e confundir os dois inventaria duplicidade onde nao ha. */
function parsearParcela(s) {
  const m = /^(\d+)\s*\/\s*(\d+)$/.exec(String(s || '').trim());
  if (!m) return null;
  return { n: Number(m[1]), de: Number(m[2]) };
}

/* Centavos inteiros, nunca float. 956.50 + 956.50 em ponto flutuante e a receita
   para uma diferenca de 0,01 aparecer como "cobra a mais" e me fazer perseguir um
   problema que nao existe. */
function centavos(v) { return Math.round(Number(v || 0) * 100); }

/* Funcao propria, e nao um trecho dentro do laco, porque e ELA que decide o que
   vira acusacao de duplicidade. Regra inline nao se testa, e uma regra que
   classifica dinheiro errado sem ninguem notar e pior do que nao ter regra. */
function classificarGrupo(g) {
  if (g.length === 1) return { tipo: 'unica' };

  const parsed = g.map(function (l) { return parsearParcela(l.parcela); });
  const semRotulo = parsed.filter(function (x) { return !x; }).length;
  const denominadores = {};
  for (const x of parsed) if (x) denominadores[x.de] = (denominadores[x.de] || 0) + 1;
  const listaDen = Object.keys(denominadores).map(Number);

  /* Codigos repetidos seriam bug NOSSO (a mesma conta gravada duas vezes).
     Codigos distintos significam contas distintas no Omie — e ai a duplicidade,
     se houver, e de la. Separar isso decide para quem vai o problema. */
  const codsUnicos = new Set(g.map(function (l) { return l.cod; })).size;
  const soma = g.reduce(function (s, l) { return s + centavos(l.valor); }, 0);
  const todosMesmoValor = new Set(g.map(function (l) { return centavos(l.valor); })).size === 1;

  let tipo;
  if (semRotulo === g.length) {
    /* Sem rotulo nenhum nao da para afirmar nada. Chutar "duplicado" aqui
       encheria a lista de falso positivo e faria o numero perder o sentido. */
    tipo = 'sem_rotulo';
  } else if (listaDen.length === 1 && listaDen[0] === 1) {
    /* TODAS dizem "de 1". Duas parcelas unicas da mesma nota nao existem. */
    tipo = 'suspeito_duplicado';
  } else if (listaDen.length === 1 && listaDen[0] === g.length) {
    tipo = 'parcelado_ok';      /* 3 linhas dizendo "de 3" */
  } else {
    tipo = 'parcelamento_incompleto';
  }

  return {
    tipo: tipo, soma: soma, todosMesmoValor: todosMesmoValor,
    codigosDistintos: codsUnicos,
    origemProvavel: codsUnicos === g.length ? 'contas distintas no Omie'
                                            : 'linha repetida no nosso sync'
  };
}

module.exports = async function (context, req) {
  const t0 = Date.now();
  const q = req.query || {};
  const filtroUnidade = String(q.unidade || '').toUpperCase();
  let amostra = parseInt(q.amostra, 10);
  if (!isFinite(amostra) || amostra < 0) amostra = 8;
  if (amostra > 20) amostra = 20;

  const out = { ok: true, timeMs: 0 };

  try {
    if (!(await isAdmin(req))) {
      context.res = { status: 403, headers: { 'Content-Type': 'application/json' },
        body: { error: 'Apenas admin' } };
      return;
    }

    const client = await getGraphClient();
    const siteId = await resolveSiteId(client);
    const idDoc = await resolveListId(client, siteId, LIST_DOCFIS);
    if (!idDoc) throw new Error('Lista ' + LIST_DOCFIS + ' nao existe');

    const campos = 'ChaveAcesso,NumeroNF,NumeroParcela,Valor,DataVencimento,StatusOmie,' +
                   'UnidadeOmie,CodigoLancamentoOmie,EmitenteNome,EmitenteCNPJ,Descartado';
    const linhas = [];
    let url = '/sites/' + siteId + '/lists/' + idDoc + '/items?$expand=' +
              encodeURIComponent('fields($select=' + campos + ')') + '&$top=999';
    let p = 0;
    while (url && p < 20) {
      const r = await client.api(url).get();
      for (const it of (r.value || [])) linhas.push(it);
      p++;
      const nl = r['@odata.nextLink'];
      url = nl ? nl.replace('https://graph.microsoft.com/v1.0', '') : null;
    }
    out.tempos = { sharePointMs: Date.now() - t0 };

    /* Agrupa por chave. So entram linhas do Omie: linha de SEFAZ ou lancamento
       manual e um documento inteiro, nao uma parcela, e misturar as duas
       contagens produziria "duplicidade" que e so origem diferente. */
    const grupos = {};
    let semChave = 0, descartadas = 0, semCodigoOmie = 0;
    for (const it of linhas) {
      const f = it.fields || {};
      if (String(f.Descartado || '') === 'Sim') { descartadas++; continue; }
      if (!f.CodigoLancamentoOmie) { semCodigoOmie++; continue; }
      const ch = soDigitos(f.ChaveAcesso);
      if (ch.length !== 44) { semChave++; continue; }
      const u = String(f.UnidadeOmie || '').toUpperCase();
      if (filtroUnidade && u !== filtroUnidade) continue;
      (grupos[ch] = grupos[ch] || []).push({
        id: it.id, cod: String(f.CodigoLancamentoOmie),
        parcela: String(f.NumeroParcela || ''),
        valor: Number(f.Valor || 0),
        venc: f.DataVencimento ? String(f.DataVencimento).slice(0, 10) : null,
        status: String(f.StatusOmie || ''), unidade: u,
        emitente: f.EmitenteNome || '', numeroNF: String(f.NumeroNF || '')
      });
    }

    const cls = { unica: 0, parcelado_ok: 0, suspeito_duplicado: 0,
                  parcelamento_incompleto: 0, sem_rotulo: 0 };
    const suspeitos = [];
    let valorEmSuspeita = 0;

    for (const ch of Object.keys(grupos)) {
      const g = grupos[ch];
      const c = classificarGrupo(g);
      cls[c.tipo]++;

      if (c.tipo === 'suspeito_duplicado') {
        valorEmSuspeita += c.soma;
        suspeitos.push({
          chave: ch, numeroNF: g[0].numeroNF, emitente: g[0].emitente,
          unidade: g[0].unidade, linhas: g.length,
          codigosDistintos: c.codigosDistintos,
          origemProvavel: c.origemProvavel,
          todosMesmoValor: c.todosMesmoValor,
          somaCentavos: c.soma,
          soma: (c.soma / 100).toFixed(2),
          detalhe: g.map(function (l) {
            return { cod: l.cod, parcela: l.parcela, valor: l.valor,
                     venc: l.venc, status: l.status };
          })
        });
      }
    }

    out.universo = {
      linhas: linhas.length, descartadas: descartadas, semChave: semChave,
      semCodigoOmie: semCodigoOmie, notasComChave: Object.keys(grupos).length
    };
    out.classificacao = cls;
    out.valorTotalEmSuspeita = (valorEmSuspeita / 100).toFixed(2);

    /* ---- a prova: comparar a soma com o valor REAL da NF ----
       Classificar pelo rotulo diz "parece duplicado". Buscar a nota no
       /contador/xml/ e comparar mostra o dinheiro. Sem este passo eu estaria
       apresentando uma suspeita como se fosse um achado. */
    suspeitos.sort(function (a, b) { return b.somaCentavos - a.somaCentavos; });
    const conferidos = [];
    for (const s of suspeitos.slice(0, amostra)) {
      if (Date.now() - t0 > ORCAMENTO_MS) { out.conferenciaInterrompida = true; break; }
      try {
        const doc = await buscarXmlAutorizado(s.chave, getCredentials(s.unidade));
        if (!doc) {
          conferidos.push({ chave: s.chave, numeroNF: s.numeroNF,
                            resultado: 'nota nao esta no Omie — nao da para comparar' });
        } else {
          const nf = centavos(doc.valor);
          const dif = s.somaCentavos - nf;
          conferidos.push({
            chave: s.chave, numeroNF: s.numeroNF, emitente: s.emitente,
            valorDaNF: (nf / 100).toFixed(2),
            somaDasLinhas: s.soma,
            diferenca: (dif / 100).toFixed(2),
            /* O veredito por nota, escrito: um numero solto na tabela seria lido
               de tres jeitos diferentes por tres pessoas. */
            resultado: dif === 0 ? 'CONFERE — a soma bate com a nota'
                     : dif > 0 ? 'COBRA A MAIS — o card mostra ' + (dif / 100).toFixed(2) +
                                 ' alem do que a nota vale'
                     : 'soma MENOR que a nota — pode faltar parcela',
            statusNaSefaz: doc.status === '00' ? 'autorizada'
                         : doc.status === '10' ? 'CANCELADA' : doc.status
          });
        }
      } catch (e) {
        conferidos.push({ chave: s.chave, resultado: 'erro: ' + (e && e.message) });
      }
      await dorme(PAUSA_MS);
    }
    out.conferidosContraANF = conferidos;
    out.cobrandoAMais = conferidos.filter(function (c) {
      return /COBRA A MAIS/.test(c.resultado || '');
    }).length;

    /* Lista completa fica por ultimo e cortada: o que decide e o numero e a
       prova, nao o despejo. */
    out.suspeitos = suspeitos.slice(0, 40);
    if (suspeitos.length > 40) out.suspeitosOmitidos = suspeitos.length - 40;

    out.leitura =
      'classificacao.suspeito_duplicado = grupos onde TODAS as linhas dizem "001/001": ' +
      'cada uma se declara parcela unica da mesma nota, o que nao existe. ' +
      'parcelado_ok = N linhas dizendo "de N", que e parcelamento legitimo e nao e ' +
      'problema. conferidosContraANF e a prova: compara a soma das linhas com o valor ' +
      'que a NF realmente tem, lido do XML autorizado. cobrandoAMais e o numero que ' +
      'importa — cada um desses e um card que pede aprovacao de mais dinheiro do que a ' +
      'nota vale. origemProvavel separa "o Omie tem duas contas" de "nosso sync gravou ' +
      'a mesma conta duas vezes", que levam a correcoes opostas.';

    out.timeMs = Date.now() - t0;
    context.res = { status: 200, headers: { 'Content-Type': 'application/json' }, body: out };
  } catch (err) {
    out.timeMs = Date.now() - t0;
    context.res = { status: 500, headers: { 'Content-Type': 'application/json' },
      body: Object.assign({ error: (err && err.message) || String(err) }, out) };
  }
};
