/**
 * /api/ListarDocumentosFiscais  (GET) — alimenta o quadro "NFs a Pagar".
 *
 * Devolve os documentos baixados da SEFAZ JA CLASSIFICADOS em colunas, cruzando
 * com a lista de Notas para saber o estagio real de cada um:
 *
 *   novas      documento da SEFAZ que ninguem lancou ainda
 *   lancadas   ja virou NF no sistema e esta aguardando aprovacao
 *   aprovadas  NF aprovada, ainda nao paga
 *   quitadas   NF marcada como processada/paga
 *
 * A COLUNA NAO E UM CAMPO GRAVADO: e derivada do status da nota vinculada. Assim
 * o quadro nunca discorda da fila de aprovacao — nao existe estado duplicado para
 * sair de sincronia.
 *
 * Tambem devolve o painel `sefaz` (ponteiro por CNPJ) para a tela mostrar quando
 * foi a ultima consulta e se alguma filial esta com erro.
 */

require('isomorphic-fetch');
const { getGraphClient } = require('../shared/graph');
const {
  LIST_DOCFIS, LIST_SEFAZ, LIST_NOTAS, resolveListId, soDigitos, lerConfigSefaz, lerCorteVencimento, lerCnpjsConfigurados
} = require('../shared/documentosFiscais');
/* Escopo mora em shared porque o detalhe da NF aplica a MESMA regra. Duas copias
   de regra de visibilidade divergem, e a que diverge e sempre a esquecida. */
const { montarEscopo, podeVer, todosItens } = require('../shared/escopoNF');
const { carregarNotas } = require('../shared/notas');

/* Traduz o status da nota para a coluna do quadro. */
function colunaPorStatus(status, processado) {
  const s = String(status || '').toLowerCase();
  if (processado) return 'quitadas';
  if (s === 'aprovada') return 'aprovadas';
  if (s === 'rejeitada') return null;   /* rejeitada volta a ser pendencia, nao card */
  return 'lancadas';
}

module.exports = async function (context, req) {
  const diag = { step: 'init', timeMs: 0 };
  const t0 = Date.now();

  try {
    const client = await getGraphClient();

    diag.step = 'site';
    const host = process.env.SHAREPOINT_SITE_HOSTNAME;
    const path = process.env.SHAREPOINT_SITE_PATH;
    const site = await client.api('/sites/' + host + ':' + path).get();
    const siteId = site.id;

    diag.step = 'listas';
    const idDoc = await resolveListId(client, siteId, LIST_DOCFIS);
    if (!idDoc) {
      /* A estrutura ainda nao foi criada. Nao e erro: e o estado inicial. */
      context.res = { status: 200, headers: { 'Content-Type': 'application/json' },
        body: { ok: true, precisaCriarEstrutura: true,
                mensagem: 'Rode /api/CriarListaDocumentosFiscais para criar o quadro.',
                colunas: { novas: [], lancadas: [], aprovadas: [], quitadas: [] },
                sefaz: [] } };
      return;
    }
    const idNotas = await resolveListId(client, siteId, LIST_NOTAS);
    const idSefaz = await resolveListId(client, siteId, LIST_SEFAZ);

    diag.step = 'ler';
    const docs = await todosItens(client, siteId, idDoc);
    /* As notas NAO podem ser lidas cruas. O SharePoint guarda cada coluna com um
       `name` interno que nem sempre e igual ao `displayName`, e nesta lista eles
       divergem (medido). Lendo cru, fields.Status vem undefined — e undefined em
       status nao explode: colunaPorStatus cai no default e joga TODO card
       vinculado em "Lancadas", inclusive os aprovados e os ja pagos. Um card na
       coluna errada nao parece defeito, parece o processo estar atrasado. */
    let notas = [];
    if (idNotas) {
      const r = await carregarNotas(client, siteId, idNotas);
      notas = r.notas.map(function (n) { return { id: n.id, fields: n.f }; });
      diag.notasColunasDivergentes = r.divergentes;
    }
    const ponteiros = idSefaz ? await todosItens(client, siteId, idSefaz, 2) : [];

    /* Unidade e diretoria vem do CADASTRO DE FORNECEDOR, pelo CNPJ do emitente.
       A SEFAZ nao tem essa informacao — ela e nossa, e e o que decide quem aprova.
       Documento cujo emitente nao esta cadastrado fica marcado (semCadastro) em vez
       de receber um palpite: atribuir diretoria errada mandaria a NF para o
       aprovador errado, que e pior do que admitir que falta cadastro. */
    const idForn = await resolveListId(client, siteId, 'PRONEP-NF-Fornecedores');
    const fornecedorPorCnpj = {};
    if (idForn) {
      for (const it of await todosItens(client, siteId, idForn)) {
        const f = it.fields || {};
        /* O SharePoint renomeou as colunas importadas do XLSX: field_2=documento,
           field_4=unidade, field_5=diretoria, field_7=ativo (mesmo mapa de
           ListarFornecedores — se um mudar, os dois mudam). */
        const doc = soDigitos(f.field_2);
        if (doc.length !== 14) continue;
        if (fornecedorPorCnpj[doc]) continue;   /* primeiro vence; duplicata e problema do cadastro */
        fornecedorPorCnpj[doc] = {
          razao: f.Title || '',
          fantasia: f.field_3 || '',
          unidade: f.field_4 || '',
          diretoria: f.field_5 || '',
          ativo: String(f.field_7 || '').toLowerCase() === 'sim'
        };
      }
    }
    diag.fornecedoresIndexados = Object.keys(fornecedorPorCnpj).length;

    /* Indexa notas por id e por chave de acesso. A chave permite reconhecer uma
       nota lancada ANTES de a coluna ChaveAcesso existir no documento. */
    const notaPorId = {};
    const notaPorChave = {};
    for (const n of notas) {
      const f = n.fields || {};
      notaPorId[String(n.id)] = n;
      const ch = soDigitos(f.ChaveAcesso);
      if (ch.length === 44) notaPorChave[ch] = n;
    }

    /* AGRUPAMENTO DE PARCELAS.
       Uma NF em 3x vira TRES contas a pagar no Omie, cada uma com seu vencimento
       e seu status. Sem agrupar, o quadro mostraria tres cards da mesma nota —
       exatamente a duplicidade que ele existe para evitar.
       Agrupa por chave de acesso; sem chave, por emitente + numero da NF. Linha da
       SEFAZ ou lancamento manual (sem CodigoLancamentoOmie) fica sozinha: ali cada
       linha ja e um documento, nao uma parcela. */
    const grupos = [];
    const porGrupo = {};
    let descartados = 0;
    for (const d of docs) {
      const f = d.fields || {};
      if (String(f.Descartado || '') === 'Sim') { descartados++; continue; }

      let chaveGrupo = null;
      if (f.CodigoLancamentoOmie) {
        const ch = soDigitos(f.ChaveAcesso);
        chaveGrupo = (ch.length === 44)
          ? 'ch:' + ch
          : 'nf:' + (f.UnidadeOmie || '') + '|' + soDigitos(f.CodigoClienteOmie) + '|' + (f.NumeroNF || '');
        /* Numero de NF vazio nao pode virar chave de grupo: juntaria contas de
           fornecedores diferentes num card so. */
        if (!f.NumeroNF && !soDigitos(f.ChaveAcesso)) chaveGrupo = null;
      }

      if (!chaveGrupo) { grupos.push({ principal: d, parcelas: [d] }); continue; }
      if (!porGrupo[chaveGrupo]) {
        porGrupo[chaveGrupo] = { principal: d, parcelas: [] };
        grupos.push(porGrupo[chaveGrupo]);
      }
      porGrupo[chaveGrupo].parcelas.push(d);
    }

    diag.step = 'escopo';
    const verComo = String((req.query && req.query.verComo) || '').trim();
    const escopo = await montarEscopo(client, siteId, req, verComo);

    const colunas = { novas: [], lancadas: [], aprovadas: [], quitadas: [] };
    let foraDoEscopo = 0;

    for (const grupo of grupos) {
      const d = grupo.principal;
      const f = d.fields || {};

      let nota = f.NotaItemId ? notaPorId[String(f.NotaItemId)] : null;
      if (!nota) {
        const ch = soDigitos(f.ChaveAcesso);
        if (ch.length === 44 && notaPorChave[ch]) nota = notaPorChave[ch];
      }

      const card = {
        id: d.id,
        chaveAcesso: f.ChaveAcesso || '',
        numeroNF: f.NumeroNF || '',
        serie: f.Serie || '',
        emitenteCNPJ: f.EmitenteCNPJ || '',
        emitenteNome: f.EmitenteNome || '',
        valor: f.Valor == null ? null : Number(f.Valor),
        dataEmissao: f.DataEmissao || null,
        dataVencimento: f.DataVencimento || null,
        cnpjDestino: f.CNPJDestino || '',
        origem: f.Origem || '',
        notaItemId: f.NotaItemId || null,
        vinculadoAuto: String(f.VinculadoAuto || '') === 'Sim'
      };

      /* Consolida as parcelas: o card mostra o TOTAL da nota e o PROXIMO
         vencimento em aberto — nao o da primeira parcela, que pode estar paga
         ha meses e faria a nota parecer vencida sem estar. */
      if (grupo.parcelas.length > 1 || f.CodigoLancamentoOmie) {
        const pcs = grupo.parcelas.map(function (p) {
          const pf = p.fields || {};
          return {
            id: p.id,
            numero: pf.NumeroParcela || '',
            valor: pf.Valor == null ? null : Number(pf.Valor),
            vencimento: pf.DataVencimento ? String(pf.DataVencimento).substring(0, 10) : null,
            status: pf.StatusOmie || '',
            paga: String(pf.StatusOmie || '').toUpperCase().indexOf('PAGO') >= 0,
            codigoBarras: pf.CodigoBarras || ''
          };
        }).sort(function (a, b) {
          return (a.vencimento || '9999') < (b.vencimento || '9999') ? -1 : 1;
        });

        card.parcelas = pcs;
        card.totalParcelas = pcs.length;
        card.parcelasPagas = pcs.filter(function (p) { return p.paga; }).length;
        card.valor = pcs.reduce(function (s, p) { return s + (Number(p.valor) || 0); }, 0);

        const emAberto = pcs.filter(function (p) { return !p.paga; });
        card.dataVencimento = (emAberto[0] || pcs[0] || {}).vencimento || null;
        card.statusOmie = emAberto.length ? (emAberto[0].status || '') : 'PAGO';
        card.unidadeOmie = f.UnidadeOmie || '';
        card.codigoBarras = (emAberto[0] || pcs[0] || {}).codigoBarras || '';
        card.todasPagas = emAberto.length === 0;

        /* ALERTA DE DUPLICIDADE.
           Medido em 14/08/2026 (DiagParcelasDuplicadas): 10 grupos do quadro tem
           duas contas que o Omie marcou "001/001" — cada uma se declarando parcela
           UNICA da mesma nota. Oito foram conferidas contra o XML autorizado e as
           oito somam EXATAMENTE o dobro do que a nota vale. R$ 7.882,77 de excesso,
           nas tres unidades, com sete fornecedores diferentes.

           Sem este aviso o card exibe o dobro em silencio, e quem aprova autoriza o
           dobro sem ter como desconfiar — o valor parece so um total.

           A regra e a MESMA da sonda, de proposito: todas as linhas dizendo "de 1".
           Nota parcelada de verdade traz 001/003, 002/003... e nao entra aqui.
           Linha sem rotulo tambem nao entra: nao saber quantas parcelas existem e
           diferente de saber que ha uma so, e acusar no escuro faria o aviso virar
           ruido que todo mundo aprende a ignorar.

           Isto NAO corrige nada — a origem e o Omie, que tem duas contas de fato
           (os codigos de lancamento sao distintos). O aviso existe para a conferencia
           acontecer antes do pagamento, nao depois. */
        if (pcs.length > 1) {
          const rotulos = pcs.map(function (p) {
            const m = /^(\d+)\s*\/\s*(\d+)$/.exec(String(p.numero || '').trim());
            return m ? Number(m[2]) : null;
          });
          const todasDizemUnica = rotulos.every(function (d) { return d === 1; });
          if (todasDizemUnica) {
            card.alertaDuplicidade = {
              contas: pcs.length,
              /* Centavos inteiros: somar 956.50 duas vezes em float produziria
                 uma diferenca de centavo, e centavo sobrando vira desconfianca
                 no aviso inteiro. */
              valorUnitario: pcs[0].valor,
              mesmoValor: new Set(pcs.map(function (p) {
                return Math.round(Number(p.valor || 0) * 100);
              })).size === 1,
              motivo: pcs.length + ' contas a pagar desta nota, cada uma marcada ' +
                      'no Omie como parcela unica (001/001). Isso nao e parcelamento.'
            };
          }
        }
      }

      const forn = fornecedorPorCnpj[soDigitos(f.EmitenteCNPJ)] || null;
      if (forn) {
        card.unidade = forn.unidade || '';
        card.diretoria = forn.diretoria || '';
        card.fornecedorCadastrado = true;
        card.fornecedorAtivo = forn.ativo;
        /* Cadastro existe mas esta incompleto: nao da para rotear a NF assim, e o
           sintoma e diferente de "fornecedor desconhecido" — a acao tambem e. */
        card.cadastroIncompleto = !forn.unidade || !forn.diretoria;
        if (forn.razao && !card.emitenteNome) card.emitenteNome = forn.razao;
      } else {
        card.fornecedorCadastrado = false;
        card.unidade = '';
        card.diretoria = '';
      }

      /* A unidade da empresa Omie e mais confiavel que a do cadastro de
         fornecedor: o cadastro diz onde o fornecedor costuma atender, o Omie diz
         qual filial de fato assumiu a conta. */
      if (f.UnidadeOmie) card.unidade = f.UnidadeOmie;

      /* PAGO NO OMIE E VERDADE FINANCEIRA e vem antes de tudo. Se o financeiro ja
         pagou, o card vai para Quitadas mesmo que a NF aqui nunca tenha sido
         lancada ou aprovada — o contrario deixaria uma conta paga parada em
         "Novas", convidando alguem a pagar de novo. */
      if (card.todasPagas) {
        if (nota) {
          card.notaItemId = String(nota.id);
          card.statusNota = (nota.fields || {}).Status || '';
        }
        if (podeVer(card, escopo)) colunas.quitadas.push(card); else foraDoEscopo++;
        continue;
      }

      if (!nota) {
        if (podeVer(card, escopo)) colunas.novas.push(card); else foraDoEscopo++;
        continue;
      }

      const nf = nota.fields || {};
      const alvo = colunaPorStatus(nf.Status, nf.Processado === true || nf.Processado === 'Sim');
      if (!alvo) {
        /* rejeitada: volta pra fila */
        if (podeVer(card, escopo)) colunas.novas.push(card); else foraDoEscopo++;
        continue;
      }

      card.notaItemId = String(nota.id);
      card.statusNota = nf.Status || '';
      card.aprovador = nf.AprovadorEmail || nf.Aprovador || '';
      /* Depois de lancada, a NOTA manda: alguem pode ter corrigido a unidade ou a
         diretoria no lancamento, e essa correcao e mais recente que o cadastro.
         Mas so sobrescreve o que a nota realmente tem — nota sem o campo nao apaga
         o que veio do fornecedor. */
      if (nf.Unidade) card.unidade = nf.Unidade;
      if (nf.Diretoria) card.diretoria = nf.Diretoria;
      /* Divergencia entre o cadastro e o que foi lancado: nao e erro, mas e o tipo
         de coisa que explica NF na fila do aprovador errado. */
      card.divergeDoCadastro = !!(forn && nf.Diretoria && forn.diretoria &&
                                  nf.Diretoria !== forn.diretoria);
      if (podeVer(card, escopo)) colunas[alvo].push(card); else foraDoEscopo++;
    }

    /* Vencimento mais proximo primeiro; sem vencimento vai para o fim. */
    for (const k of Object.keys(colunas)) {
      colunas[k].sort(function (a, b) {
        const va = a.dataVencimento || '9999-12-31';
        const vb = b.dataVencimento || '9999-12-31';
        return va < vb ? -1 : va > vb ? 1 : 0;
      });
    }

    /* A UNIDADE vem da configuracao, nao do apelido. Deduzir "RJ" da ultima
       palavra de "PRONEP RJ" funciona hoje e quebraria em silencio no dia em que
       alguem renomear para "PRONEP Rio de Janeiro" — e o filtro por unidade da
       tela depende disso casar com o UnidadeOmie dos cards. */
    const unidadePorCnpj = {};
    try {
      for (const c of lerCnpjsConfigurados()) {
        if (c.unidade) unidadePorCnpj[c.cnpj] = String(c.unidade).toUpperCase();
      }
    } catch (e) { /* sem config, a tela cai no apelido */ }

    const sefaz = ponteiros.map(function (p) {
      const f = p.fields || {};
      return {
        cnpj: f.CNPJ || '', apelido: f.Apelido || '',
        unidade: unidadePorCnpj[soDigitos(f.CNPJ)] || '',
        ultimoNSU: Number(f.UltimoNSU || 0), maxNSU: Number(f.MaxNSU || 0),
        ultimaConsulta: f.UltimaConsulta || null,
        cStat: f.UltimoCStat || '', motivo: f.UltimoMotivo || '',
        baixados: Number(f.Baixados || 0),
        /* 137 (sem documento novo) e 138 (documentos localizados) sao normais. */
        saudavel: ['137', '138', ''].indexOf(String(f.UltimoCStat || '')) >= 0
      };
    });

    diag.step = 'done';
    diag.timeMs = Date.now() - t0;
    context.res = {
      status: 200, headers: { 'Content-Type': 'application/json' },
      body: {
        ok: true,
        colunas: colunas,
        sefaz: sefaz,
        /* A tela precisa DIZER que a visao e parcial. Um quadro filtrado que se
           apresenta como completo faz o gestor concluir que nao ha nada a fazer. */
        escopo: {
          verTudo: escopo.verTudo,
          motivo: escopo.motivo || (escopo.verTudo ? '' : 'gestor'),
          minhasDiretorias: escopo.diretorias || [],
          lente: escopo.lente || null,
          ocultadosPorEscopo: foraDoEscopo
        },
        /* COBERTURA DE CHAVE, medida sobre os cards REAIS — nao sobre amostra.
           Decide se vale buscar o XML na SEFAZ: a chave e o unico jeito de
           consultar, e sem ela nao ha o que pedir. Uma amostra de 6 contas deu
           "17%", numero que nao sustenta decisao nenhuma; aqui a base e o quadro
           inteiro. Quebrado por unidade porque a mistura de perfis (RJ compra
           mais material, servicos concentram em outra) esconderia a diferenca. */
        coberturaChave: (function () {
          const porUnid = {};
          let com = 0, total = 0;
          for (const k of Object.keys(colunas)) {
            for (const c of colunas[k]) {
              const u = c.unidade || '(sem unidade)';
              if (!porUnid[u]) porUnid[u] = { com: 0, total: 0 };
              porUnid[u].total++; total++;
              if (soDigitos(c.chaveAcesso).length === 44) { porUnid[u].com++; com++; }
            }
          }
          for (const u of Object.keys(porUnid)) {
            porUnid[u].percentual = porUnid[u].total
              ? Math.round((porUnid[u].com / porUnid[u].total) * 100) + '%' : '—';
          }
          return {
            total: total, comChave: com, semChave: total - com,
            percentual: total ? Math.round((com / total) * 100) + '%' : '—',
            porUnidade: porUnid
          };
        })(),
        integracao: await lerConfigSefaz(client, siteId),
        corteVencimento: await lerCorteVencimento(client, siteId),
        totais: {
          novas: colunas.novas.length, lancadas: colunas.lancadas.length,
          aprovadas: colunas.aprovadas.length, quitadas: colunas.quitadas.length,
          descartados: descartados,
          /* Contagem so das NOVAS: em coluna posterior o cadastro ja nao bloqueia
             nada, e somar tudo inflaria um numero que serve para acao. */
          semCadastro: colunas.novas.filter(function (c) { return !c.fornecedorCadastrado; }).length,
          cadastroIncompleto: colunas.novas.filter(function (c) { return c.cadastroIncompleto; }).length
        },
        /* Vai na resposta de proposito: se um dia a lista de Notas ganhar colunas
           renomeadas novas, quem estiver depurando um card na coluna errada ve o
           motivo aqui em vez de suspeitar do fluxo de aprovacao. */
        notasColunasDivergentes: diag.notasColunasDivergentes || [],
        timeMs: diag.timeMs
      }
    };
  } catch (err) {
    diag.timeMs = Date.now() - t0;
    context.res = {
      status: 500, headers: { 'Content-Type': 'application/json' },
      body: Object.assign({ error: (err && err.message) || String(err) }, diag)
    };
  }
};
