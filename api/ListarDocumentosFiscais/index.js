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
  LIST_DOCFIS, LIST_SEFAZ, LIST_NOTAS, resolveListId, soDigitos
} = require('../shared/documentosFiscais');

async function todosItens(client, siteId, listId, maxPaginas) {
  const out = [];
  let url = '/sites/' + siteId + '/lists/' + listId + '/items?expand=fields&$top=999';
  let p = 0;
  while (url && p < (maxPaginas || 20)) {
    const r = await client.api(url).get();
    out.push.apply(out, r.value || []);
    p++;
    const nl = r['@odata.nextLink'];
    url = nl ? nl.replace('https://graph.microsoft.com/v1.0', '') : null;
  }
  return out;
}

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
    const notas = idNotas ? await todosItens(client, siteId, idNotas) : [];
    const ponteiros = idSefaz ? await todosItens(client, siteId, idSefaz, 2) : [];

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

    const colunas = { novas: [], lancadas: [], aprovadas: [], quitadas: [] };
    let descartados = 0;

    for (const d of docs) {
      const f = d.fields || {};
      if (String(f.Descartado || '') === 'Sim') { descartados++; continue; }

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

      if (!nota) { colunas.novas.push(card); continue; }

      const nf = nota.fields || {};
      const alvo = colunaPorStatus(nf.Status, nf.Processado === true || nf.Processado === 'Sim');
      if (!alvo) { colunas.novas.push(card); continue; }   /* rejeitada: volta pra fila */

      card.notaItemId = String(nota.id);
      card.statusNota = nf.Status || '';
      card.aprovador = nf.AprovadorEmail || nf.Aprovador || '';
      card.unidade = nf.Unidade || '';
      card.diretoria = nf.Diretoria || '';
      colunas[alvo].push(card);
    }

    /* Vencimento mais proximo primeiro; sem vencimento vai para o fim. */
    for (const k of Object.keys(colunas)) {
      colunas[k].sort(function (a, b) {
        const va = a.dataVencimento || '9999-12-31';
        const vb = b.dataVencimento || '9999-12-31';
        return va < vb ? -1 : va > vb ? 1 : 0;
      });
    }

    const sefaz = ponteiros.map(function (p) {
      const f = p.fields || {};
      return {
        cnpj: f.CNPJ || '', apelido: f.Apelido || '',
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
        totais: {
          novas: colunas.novas.length, lancadas: colunas.lancadas.length,
          aprovadas: colunas.aprovadas.length, quitadas: colunas.quitadas.length,
          descartados: descartados
        },
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
