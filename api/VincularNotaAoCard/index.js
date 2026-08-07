/**
 * /api/VincularNotaAoCard — liga um card do quadro a uma NF ja lancada, a mao.
 *
 *   GET  ?docId=<id>              lista as NFs candidatas (nao grava)
 *   POST { docId, notaId }        grava o vinculo
 *
 * POR QUE O AUTOMATICO NAO BASTA. O casamento automatico exige CNPJ **e** numero
 * da NF. Quando o numero foi digitado errado no lancamento, ele nunca casa — e
 * digitar errado e comum: medido na base da Pronep, Payfy 42010 contra 42012 no
 * Omie, Iberwan 18593 contra 18539, Control iD 4130719 contra 01430719. Em todos,
 * o VALOR batia exatamente. Sao a mesma conta com um digito trocado.
 *
 * Por isso as candidatas aqui sao buscadas por FORNECEDOR, nao por numero: se o
 * numero fosse confiavel, o automatico ja teria resolvido. Quem ordena e o valor,
 * que e o sinal que sobrevive ao erro de digitacao.
 *
 * NAO ESCOLHE SOZINHO. Devolve candidatas com a divergencia exposta (numero do
 * Omie x numero lancado) e quem decide e a pessoa. Vincular a nota errada faria o
 * card exibir o status de outra: uma conta em aberto apareceria como paga.
 *
 * NAO CORRIGE O NUMERO DA NOTA. Vincular e dizer "sao a mesma conta"; reescrever
 * o numero de uma NF ja aprovada e alterar registro fiscal depois da decisao, e
 * isso precisa de escolha explicita de quem responde pelo processo — nao de um
 * efeito colateral de um botao de vincular.
 */

require('isomorphic-fetch');
const { getGraphClient, resolveSiteId } = require('../shared/graph');
const { resolveListId, LIST_DOCFIS, LIST_NOTAS, soDigitos, vincularNota } =
      require('../shared/documentosFiscais');
const { montarEscopo, podeVer, todosItens } = require('../shared/escopoNF');
const { carregarNotas } = require('../shared/notas');
const { registrar: auditRegistrar } = require('../shared/auditLog');
const { getUser } = require('../shared/auth');

function numNorm(v) {
  const s = String(v == null ? '' : v).replace(/[^A-Za-z0-9]/g, '');
  return /^\d+$/.test(s) ? (s.replace(/^0+/, '') || '0') : s.toUpperCase();
}

function diasEntre(a, b) {
  const da = new Date(a), db = new Date(b);
  if (isNaN(da.getTime()) || isNaN(db.getTime())) return null;
  return Math.round(Math.abs(da - db) / 86400000);
}

/* Le o documento e devolve o card no formato que podeVer espera, junto das
   linhas irmas (parcelas da mesma NF), que precisam receber o mesmo vinculo. */
async function carregarDocumento(client, siteId, listDoc, docId) {
  const item = await client.api('/sites/' + siteId + '/lists/' + listDoc + '/items/' + docId)
    .expand('fields').get();
  const f = (item && item.fields) || {};
  return { id: String(item.id), f: f };
}

async function fornecedorPorCnpj(client, siteId, cnpj) {
  const idForn = await resolveListId(client, siteId, 'PRONEP-NF-Fornecedores');
  if (!idForn) return null;
  for (const it of await todosItens(client, siteId, idForn)) {
    const g = it.fields || {};
    if (soDigitos(g.field_2) === cnpj) {
      return { unidade: g.field_4 || '', diretoria: g.field_5 || '' };
    }
  }
  return null;
}

module.exports = async function (context, req) {
  const t0 = Date.now();
  const q = req.query || {};
  const body = req.body || {};
  const metodo = String(req.method || 'GET').toUpperCase();
  const docId = String(q.docId || body.docId || '');
  const verComo = String(q.verComo || body.verComo || '').trim();

  try {
    if (!docId) throw new Error('docId obrigatorio');

    const client = await getGraphClient();
    const siteId = await resolveSiteId(client);
    const listDoc = await resolveListId(client, siteId, LIST_DOCFIS);
    const listNotas = await resolveListId(client, siteId, LIST_NOTAS);
    if (!listDoc || !listNotas) throw new Error('Listas nao encontradas');

    const escopo = await montarEscopo(client, siteId, req, verComo);
    if (!escopo.autenticado) {
      context.res = { status: 401, headers: { 'Content-Type': 'application/json' },
        body: { error: 'Nao autenticado' } };
      return;
    }

    const doc = await carregarDocumento(client, siteId, listDoc, docId);
    const cnpj = soDigitos(doc.f.EmitenteCNPJ);
    const forn = cnpj ? await fornecedorPorCnpj(client, siteId, cnpj) : null;
    const card = {
      unidade: (forn && forn.unidade) || doc.f.UnidadeOmie || '',
      diretoria: (forn && forn.diretoria) || '',
      fornecedorCadastrado: !!forn,
      cadastroIncompleto: !!forn && !(forn.unidade && forn.diretoria)
    };
    /* 404 e nao 403: "existe mas voce nao pode ver" ja entrega que a NF existe. */
    if (!podeVer(card, escopo)) {
      context.res = { status: 404, headers: { 'Content-Type': 'application/json' },
        body: { error: 'Documento nao encontrado' } };
      return;
    }

    const { notas } = await carregarNotas(client, siteId, listNotas);
    const valorDoc = Number(doc.f.Valor || 0);
    const numDoc = numNorm(doc.f.NumeroNF);

    /* ---------------------------------------------------------- POST: vincula */
    if (metodo === 'POST') {
      const notaId = String(body.notaId || '');
      if (!notaId) throw new Error('notaId obrigatorio');

      const nota = notas.find(function (n) { return String(n.id) === notaId; });
      if (!nota) throw new Error('NF nao encontrada');

      /* A nota escolhida precisa ser do MESMO fornecedor. Sem esta trava, um
         docId valido mais um notaId qualquer ligariam contas de empresas
         diferentes — e o card passaria a exibir o status de uma nota alheia. */
      if (soDigitos(nota.f.CNPJFornecedor) !== cnpj) {
        context.res = { status: 400, headers: { 'Content-Type': 'application/json' },
          body: { error: 'A NF escolhida e de outro fornecedor. Vinculo recusado.',
                  cnpjDoCard: cnpj, cnpjDaNota: soDigitos(nota.f.CNPJFornecedor) } };
        return;
      }

      /* Todas as parcelas da mesma NF recebem o vinculo: o quadro elege uma linha
         como principal e e ela que precisa estar ligada — vincular so a clicada
         deixaria o card orfao conforme a ordem de leitura. */
      const todos = await todosItens(client, siteId, listDoc);
      const irmas = todos.filter(function (it) {
        const g = it.fields || {};
        if (String(g.Descartado || '') === 'Sim') return false;
        if (soDigitos(g.EmitenteCNPJ) !== cnpj) return false;
        return numNorm(g.NumeroNF) === numDoc;
      });
      const alvos = irmas.length ? irmas.map(function (i) { return String(i.id); }) : [String(doc.id)];

      for (const id of alvos) {
        await vincularNota(client, siteId, id, notaId, false);   /* false = manual */
      }

      let user = null;
      try { user = await getUser(req); } catch (e) { /* auditoria nao bloqueia */ }
      auditRegistrar(user, 'vincular_nf_manual',
        { tipo: 'documento_fiscal', id: docId, numero: doc.f.NumeroNF || '' },
        'sucesso',
        { notaId: notaId, numeroNaNota: nota.f.NumeroNF || '', numeroNoOmie: doc.f.NumeroNF || '',
          divergenciaDeNumero: numNorm(nota.f.NumeroNF) !== numDoc,
          linhasVinculadas: alvos.length }
      ).catch(function () {});

      context.res = { status: 200, headers: { 'Content-Type': 'application/json' },
        body: { ok: true, linhasVinculadas: alvos.length, notaId: notaId,
                mensagem: alvos.length + ' linha(s) vinculada(s) à NF #' + notaId,
                timeMs: Date.now() - t0 } };
      return;
    }

    /* --------------------------------------------------- GET: lista candidatas */
    const jaVinculadas = {};
    for (const it of await todosItens(client, siteId, listDoc)) {
      const g = it.fields || {};
      if (g.NotaItemId) jaVinculadas[String(g.NotaItemId)] = true;
    }

    const candidatas = notas
      .filter(function (n) { return soDigitos(n.f.CNPJFornecedor) === cnpj; })
      .map(function (n) {
        const valorNota = Number(n.f.Valor || 0);
        const numNota = numNorm(n.f.NumeroNF);
        return {
          notaId: n.id,
          numeroNF: n.f.NumeroNF || '',
          valor: valorNota,
          dataVencimento: n.f.DataVencimento || null,
          status: n.f.Status || '',
          unidade: n.f.Unidade || '',
          diretoria: n.f.Diretoria || '',
          jaVinculada: !!jaVinculadas[String(n.id)],
          /* Os tres sinais que a pessoa usa para decidir. O valor vem primeiro
             porque e o unico que sobrevive a um erro de digitacao no numero. */
          valorIgual: Math.abs(valorNota - valorDoc) < 0.01,
          numeroIgual: numNota === numDoc,
          diasDeVencimento: diasEntre(n.f.DataVencimento, doc.f.DataVencimento)
        };
      })
      .sort(function (a, b) {
        if (a.jaVinculada !== b.jaVinculada) return a.jaVinculada ? 1 : -1;
        if (a.valorIgual !== b.valorIgual) return a.valorIgual ? -1 : 1;
        if (a.numeroIgual !== b.numeroIgual) return a.numeroIgual ? -1 : 1;
        return (a.diasDeVencimento == null ? 9999 : a.diasDeVencimento)
             - (b.diasDeVencimento == null ? 9999 : b.diasDeVencimento);
      })
      .slice(0, 30);

    context.res = { status: 200, headers: { 'Content-Type': 'application/json' },
      body: {
        ok: true,
        card: { docId: docId, numeroNF: doc.f.NumeroNF || '', valor: valorDoc,
                dataVencimento: doc.f.DataVencimento || null,
                emitenteNome: doc.f.EmitenteNome || '', cnpj: cnpj },
        candidatas: candidatas,
        totalDoFornecedor: candidatas.length,
        timeMs: Date.now() - t0
      } };
  } catch (err) {
    context.res = { status: 500, headers: { 'Content-Type': 'application/json' },
      body: { error: (err && err.message) || String(err), timeMs: Date.now() - t0 } };
  }
};
