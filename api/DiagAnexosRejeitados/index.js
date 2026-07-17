/**
 * /api/DiagAnexosRejeitados  (GET) — ADMIN. READ-ONLY (nao move/deleta nada).
 *
 * Recuperacao do bug de match por NumeroNF. Casa por VALOR (chave unica no nome do
 * arquivo) e varre a subpasta {Unidade}/Diretoria {Diretoria} + o ROOT de Rejeitadas.
 * Reporta nos dois sentidos:
 *   - notasSemArquivo : nota rejeitada cujo VALOR nao tem arquivo em Rejeitadas.
 *   - arquivosOrfaos  : arquivo em Rejeitadas cujo valor nao bate com nenhuma rejeitada
 *                       (movido errado — candidato a restaurar, ex.: reembolso do Rafael).
 *   - conferem        : nota com exatamente 1 arquivo do mesmo valor (ok).
 *   - multiplos       : nota com >1 arquivo do mesmo valor (ambiguo, conferir).
 *
 * Query: ?unidade= ?diretoria= (filtra as notas)   ?limite= (default 3000)
 */

require('isomorphic-fetch');
const { requireAdmin } = require('../shared/authz');
const { getGraphClient, resolveSiteId } = require('../shared/graph');

const LIST_NOTAS = 'PRONEP-NF-NotasFiscais';
const ROOT_REJ = 'Notas Fiscais/Rejeitadas';

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
// Valor extraido do NOME (aceita ..._valor.pdf e ..._valor_APROVADA_... e milhar 9.915,94).
function valorDoNome(nome) {
  const m = String(nome || '').match(/_(\d{1,3}(?:\.\d{3})*,\d{2}|\d+,\d{2})(?=[_.]|$)/g);
  if (!m || !m.length) return '';
  return m[m.length - 1].replace(/^_/, '');
}

module.exports = async function (context, req) {
  try {
    const authz = await requireAdmin(context, req);
    if (!authz) return;

    const qUnid = String((req.query && req.query.unidade) || '').trim();
    const qDir = String((req.query && req.query.diretoria) || '').trim();
    const limite = Math.min(5000, Math.max(1, parseInt((req.query && req.query.limite) || '3000', 10) || 3000));

    const client = getGraphClient();
    const siteId = await resolveSiteId(client);
    const { listId, inv } = await resolveListaNotas(client, siteId);

    // Carrega TODAS as notas (pagina completa) e filtra rejeitadas.
    const all = [];
    let url = '/sites/' + siteId + '/lists/' + listId + '/items?expand=fields&$top=1000';
    let pages = 0;
    while (url && pages < 40 && all.length < limite) {
      const r = await client.api(url).get();
      all.push.apply(all, (r.value || []));
      pages++;
      url = r['@odata.nextLink'] ? r['@odata.nextLink'].replace('https://graph.microsoft.com/v1.0', '') : null;
    }
    const rejeitadas = all.map(function (it) { return norm(it, inv); })
      .filter(function (n) { return String(n.Status) === 'Rejeitada'; })
      .filter(function (n) { return !qUnid || n.Unidade === qUnid; })
      .filter(function (n) { return !qDir || n.Diretoria === qDir; });

    // Pastas a varrer: root + cada subpasta {Unidade}/Diretoria {Diretoria} das rejeitadas.
    const folders = new Set([ROOT_REJ]);
    for (const n of rejeitadas) {
      if (n.Unidade && n.Diretoria) folders.add(ROOT_REJ + '/' + n.Unidade + '/Diretoria ' + n.Diretoria);
    }
    // Indexa arquivos por valor (um arquivo pode aparecer em mais de uma pasta? nao — dedup por id).
    const arquivos = []; // { name, valor, folder, id }
    const vistos = new Set();
    for (const folder of folders) {
      try {
        const resp = await client.api('/sites/' + siteId + '/drive/root:/' + folder + ':/children').get();
        for (const x of (resp.value || [])) {
          if (!x.file || vistos.has(x.id)) continue;
          vistos.add(x.id);
          arquivos.push({ name: x.name, valor: valorDoNome(x.name), folder: folder, id: x.id });
        }
      } catch (e) { /* pasta inexistente */ }
    }
    const arqPorValor = {};
    for (const a of arquivos) { if (!a.valor) continue; (arqPorValor[a.valor] = arqPorValor[a.valor] || []).push(a); }

    // Sentido 1: cada nota rejeitada tem arquivo com o SEU valor?
    const conferem = [], notasSemArquivo = [], multiplos = [];
    const valoresDasNotas = new Set();
    for (const n of rejeitadas) {
      const v = valorStrDe(n.Valor);
      valoresDasNotas.add(v);
      const reg = { id: n.id, numero: String(n.NumeroNF || ''), valor: v, unidade: n.Unidade, diretoria: n.Diretoria, fornecedor: n.CNPJFornecedor || n.Fornecedor, lancadoPor: n.LancadoPor };
      const hits = v ? (arqPorValor[v] || []) : [];
      if (hits.length === 1) conferem.push(reg);
      else if (hits.length === 0) notasSemArquivo.push(reg);
      else { reg.arquivos = hits.map(function (a) { return a.name; }); multiplos.push(reg); }
    }

    // Sentido 2: arquivos cujo valor NAO bate com nenhuma nota rejeitada = movidos errado.
    const arquivosOrfaos = arquivos
      .filter(function (a) { return a.valor && !valoresDasNotas.has(a.valor); })
      .map(function (a) { return { arquivo: a.name, valor: a.valor, pasta: a.folder }; });

    context.res = {
      status: 200, headers: { 'Content-Type': 'application/json', 'Cache-Control': 'no-store' },
      body: {
        ok: true,
        totalRejeitadas: rejeitadas.length, totalArquivosRejeitados: arquivos.length,
        conferemCount: conferem.length,
        notasSemArquivoCount: notasSemArquivo.length, notasSemArquivo: notasSemArquivo,
        multiplosCount: multiplos.length, multiplos: multiplos,
        arquivosOrfaosCount: arquivosOrfaos.length, arquivosOrfaos: arquivosOrfaos
      }
    };
  } catch (err) {
    context.log && context.log.error && context.log.error('DiagAnexosRejeitados:', err);
    context.res = { status: 500, headers: { 'Content-Type': 'application/json' },
      body: { error: (err && err.message) || String(err) } };
  }
};
