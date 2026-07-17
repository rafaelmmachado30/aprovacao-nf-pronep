/**
 * /api/DiagOrganizacaoNF  (GET) — ADMIN. READ-ONLY (nao move/deleta/renomeia nada).
 *
 * Mapa da organizacao dos PDFs no SharePoint pra planejar a limpeza do legado.
 * NAO altera estrutura. Estrutura canonica (mantida):
 *   - Pendentes : Notas Fiscais/Pendentes/{Unidade}/Diretoria {Diretoria}/arquivo   (depth 2)
 *   - Rejeitadas: Notas Fiscais/Rejeitadas/{Unidade}/Diretoria {Diretoria}/arquivo  (depth 2)
 *   - Aprovadas : Notas Fiscais/Notas Aprovadas/{Unidade}/{AAAA-MM-DD}/arquivo      (NAO mexer)
 *
 * Varre recursivamente cada raiz, casa arquivo x nota por VALOR (chave no nome) e reporta:
 *   - resumo por estado: total de notas, arquivos, notas sem arquivo, arquivos soltos/legado.
 *   - pastaLegadaNotasRejeitadas: conteudo de "Notas Fiscais/Notas Rejeitadas" (pasta antiga).
 *   - soltosPendentes / soltosRejeitadas: PDFs fora de {Unidade}/Diretoria {Diretoria}.
 *   - notasSemArquivo: nota cujo VALOR nao aparece em nenhum arquivo da raiz do seu estado.
 *   - orfaos: arquivo cujo valor nao bate com nenhuma nota do estado.
 *
 * Query: ?limite= (default 3000)  ?amostra= (max itens listados por bucket, default 50)
 */

require('isomorphic-fetch');
const { requireAdmin } = require('../shared/authz');
const { getGraphClient, resolveSiteId } = require('../shared/graph');

const LIST_NOTAS = 'PRONEP-NF-NotasFiscais';
const ROOT_PEND = 'Notas Fiscais/Pendentes';
const ROOT_REJ = 'Notas Fiscais/Rejeitadas';
const ROOT_REJ_LEGADO = 'Notas Fiscais/Notas Rejeitadas';
const ROOT_APROV = 'Notas Fiscais/Notas Aprovadas';

async function resolveListaNotas(client, siteId) {
  const lr = await client.api('/sites/' + siteId + '/lists').filter("displayName eq '" + LIST_NOTAS + "'").get();
  if (!lr.value || !lr.value.length) throw new Error('Lista ' + LIST_NOTAS + ' nao encontrada');
  const listId = lr.value[0].id;
  const cols = await client.api('/sites/' + siteId + '/lists/' + listId + '/columns').get();
  const inv = {};
  for (const c of (cols.value || [])) { if (c.displayName && c.name) inv[c.name] = c.displayName; }
  return { listId: listId, inv: inv };
}
function norm(item, inv) {
  const f = item.fields || {}; const out = { id: item.id };
  for (const [k, v] of Object.entries(f)) { if (inv[k]) out[inv[k]] = v; }
  return out;
}
function valorStrDe(v) {
  const n = (typeof v === 'number' ? v : Number(v)) || 0;
  return n > 0 ? n.toFixed(2).replace('.', ',') : '';
}
function valorDoNome(nome) {
  const m = String(nome || '').match(/_(\d{1,3}(?:\.\d{3})*,\d{2}|\d+,\d{2})(?=[_.]|$)/g);
  if (!m || !m.length) return '';
  return m[m.length - 1].replace(/^_/, '');
}

// Varre recursivamente uma raiz e devolve todos os ARQUIVOS com path relativo e profundidade.
// relPath = caminho sob a raiz (ex.: "SP/Diretoria Comercial/arquivo.pdf"); depth = nivel do arquivo.
async function walk(client, siteId, rootPath, maxDepth, budget) {
  const files = [];
  async function rec(relDir, depth) {
    if (depth > maxDepth || budget.n >= budget.max) return;
    const full = relDir ? rootPath + '/' + relDir : rootPath;
    let resp;
    try { resp = await client.api('/sites/' + siteId + '/drive/root:/' + full + ':/children').get(); }
    catch (e) { return; } // pasta inexistente
    for (const x of (resp.value || [])) {
      if (budget.n >= budget.max) return;
      const rel = relDir ? relDir + '/' + x.name : x.name;
      if (x.folder) { await rec(rel, depth + 1); }
      else if (x.file) {
        budget.n++;
        files.push({ name: x.name, relPath: rel, depth: (rel.split('/').length - 1), valor: valorDoNome(x.name), webUrl: x.webUrl });
      }
    }
  }
  await rec('', 0);
  return files;
}

module.exports = async function (context, req) {
  try {
    const authz = await requireAdmin(context, req);
    if (!authz) return;

    const limite = Math.min(5000, Math.max(1, parseInt((req.query && req.query.limite) || '3000', 10) || 3000));
    const amostra = Math.min(500, Math.max(1, parseInt((req.query && req.query.amostra) || '50', 10) || 50));
    const cut = function (arr) { return arr.slice(0, amostra); };

    const client = getGraphClient();
    const siteId = await resolveSiteId(client);
    const { listId, inv } = await resolveListaNotas(client, siteId);

    // Carrega todas as notas.
    const all = [];
    let url = '/sites/' + siteId + '/lists/' + listId + '/items?expand=fields&$top=1000';
    let pages = 0;
    while (url && pages < 40 && all.length < limite) {
      const r = await client.api(url).get();
      all.push.apply(all, (r.value || []));
      pages++;
      url = r['@odata.nextLink'] ? r['@odata.nextLink'].replace('https://graph.microsoft.com/v1.0', '') : null;
    }
    const notas = all.map(function (it) { return norm(it, inv); });
    const isPend = function (s) { return s === 'Lancada' || s === 'AguardandoN2'; };
    const pendentes = notas.filter(function (n) { return isPend(String(n.Status)); });
    const aprovadas = notas.filter(function (n) { return String(n.Status) === 'Aprovada'; });
    const rejeitadas = notas.filter(function (n) { return String(n.Status) === 'Rejeitada'; });

    // Varre as raizes (Aprovadas so pra contagem — nao mexemos nela).
    const budget = { n: 0, max: 8000 };
    const arqPend = await walk(client, siteId, ROOT_PEND, 3, budget);
    const arqRej = await walk(client, siteId, ROOT_REJ, 3, budget);
    const arqRejLeg = await walk(client, siteId, ROOT_REJ_LEGADO, 3, budget);
    const arqAprov = await walk(client, siteId, ROOT_APROV, 3, budget);

    // Indexa por valor por estado.
    function idxValor(arr) { const m = {}; for (const a of arr) { if (a.valor) (m[a.valor] = m[a.valor] || []).push(a); } return m; }

    // Analisa um estado: casa notas x arquivos por valor; separa soltos (depth<2) e orfaos.
    function analisa(notasEstado, arquivos, exigeSubpasta) {
      const idx = idxValor(arquivos);
      const valoresNotas = new Set();
      const semArquivo = [], multiplos = [];
      for (const n of notasEstado) {
        const v = valorStrDe(n.Valor); valoresNotas.add(v);
        const reg = { id: n.id, numero: String(n.NumeroNF || ''), valor: v, unidade: n.Unidade, diretoria: n.Diretoria, fornecedor: n.Fornecedor || n.CNPJFornecedor };
        const hits = v ? (idx[v] || []) : [];
        if (hits.length === 0) semArquivo.push(reg);
        else if (hits.length > 1) { reg.arquivos = hits.map(function (a) { return a.relPath; }); multiplos.push(reg); }
      }
      const orfaos = arquivos.filter(function (a) { return a.valor && !valoresNotas.has(a.valor); })
        .map(function (a) { return { arquivo: a.name, valor: a.valor, path: a.relPath }; });
      // solto = arquivo fora de {Unidade}/Diretoria {Diretoria} (depth != 2), so quando exige subpasta.
      const soltos = exigeSubpasta
        ? arquivos.filter(function (a) { return a.depth !== 2 || a.relPath.indexOf('/Diretoria ') < 0; })
            .map(function (a) { return { arquivo: a.name, path: a.relPath, depth: a.depth }; })
        : [];
      return { semArquivo: semArquivo, multiplos: multiplos, orfaos: orfaos, soltos: soltos };
    }

    const aPend = analisa(pendentes, arqPend, true);
    const aRej = analisa(rejeitadas, arqRej, true);

    context.res = {
      status: 200, headers: { 'Content-Type': 'application/json', 'Cache-Control': 'no-store' },
      body: {
        ok: true,
        _obs: 'READ-ONLY. Aprovadas nao e alvo de limpeza (estrutura por data mantida).',
        totais: {
          notas: { pendentes: pendentes.length, aprovadas: aprovadas.length, rejeitadas: rejeitadas.length },
          arquivos: { pendentes: arqPend.length, rejeitadas: arqRej.length, rejeitadasLegado: arqRejLeg.length, aprovadas: arqAprov.length },
          truncado: budget.n >= budget.max
        },
        pendentes: {
          notasSemArquivoCount: aPend.semArquivo.length, notasSemArquivo: cut(aPend.semArquivo),
          soltosCount: aPend.soltos.length, soltos: cut(aPend.soltos),
          orfaosCount: aPend.orfaos.length, orfaos: cut(aPend.orfaos),
          multiplosCount: aPend.multiplos.length, multiplos: cut(aPend.multiplos)
        },
        rejeitadas: {
          notasSemArquivoCount: aRej.semArquivo.length, notasSemArquivo: cut(aRej.semArquivo),
          soltosCount: aRej.soltos.length, soltos: cut(aRej.soltos),
          orfaosCount: aRej.orfaos.length, orfaos: cut(aRej.orfaos),
          multiplosCount: aRej.multiplos.length, multiplos: cut(aRej.multiplos)
        },
        pastaLegadaNotasRejeitadas: {
          existe: arqRejLeg.length > 0,
          totalArquivos: arqRejLeg.length,
          arquivos: cut(arqRejLeg.map(function (a) { return { arquivo: a.name, path: a.relPath, valor: a.valor }; }))
        }
      }
    };
  } catch (err) {
    context.log && context.log.error && context.log.error('DiagOrganizacaoNF:', err);
    context.res = { status: 500, headers: { 'Content-Type': 'application/json' },
      body: { error: (err && err.message) || String(err) } };
  }
};
