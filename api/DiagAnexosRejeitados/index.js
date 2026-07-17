/**
 * /api/DiagAnexosRejeitados  (GET) — ADMIN. READ-ONLY (nao move/deleta nada).
 *
 * Diagnostico do bug de match por NumeroNF: lista as NFs REJEITADAS cujo PDF apontado
 * (pela logica antiga, match frouxo "_numero_") tem VALOR diferente do valor da nota —
 * sinal de que o anexo esta mislinkado (aponta pro arquivo de outra NF).
 *
 * Serve pra reconciliar manualmente no SharePoint. So LEITURA.
 *
 * Query: ?unidade=SP ?diretoria=Tecnologia (filtros opcionais) ?limite=500
 */

require('isomorphic-fetch');
const { requireAdmin } = require('../shared/authz');
const { getGraphClient, resolveSiteId } = require('../shared/graph');

const LIST_NOTAS = 'PRONEP-NF-NotasFiscais';

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
// Valor extraido do NOME do arquivo. Aceita "_valor_" (aprovado: ..._valor_APROVADA_...)
// e "_valor.pdf" (pendente/rejeitado: ..._valor.pdf, sem _ final).
function valorDoNome(nome) {
  const m = String(nome || '').match(/_(\d{1,3}(?:\.\d{3})*,\d{2}|\d+,\d{2})(?=[_.]|$)/g);
  if (!m || !m.length) return '';
  return m[m.length - 1].replace(/^_/, ''); // ultimo grupo (valor vem depois da UF)
}
// O nome contem esse valor? (seguido de "_" ou "." — nao um prefixo por acaso)
function nomeTemValor(nome, v) {
  return !!v && (String(nome).indexOf('_' + v + '_') >= 0 || String(nome).indexOf('_' + v + '.') >= 0);
}

module.exports = async function (context, req) {
  try {
    const authz = await requireAdmin(context, req);
    if (!authz) return;

    const qUnid = String((req.query && req.query.unidade) || '').trim();
    const qDir = String((req.query && req.query.diretoria) || '').trim();
    const limite = Math.min(2000, Math.max(1, parseInt((req.query && req.query.limite) || '500', 10) || 500));

    const client = getGraphClient();
    const siteId = await resolveSiteId(client);
    const { listId, inv } = await resolveListaNotas(client, siteId);

    // Carrega as notas REJEITADAS.
    const all = [];
    let url = '/sites/' + siteId + '/lists/' + listId + '/items?expand=fields&$top=500';
    let pages = 0;
    while (url && pages < 20 && all.length < limite) {
      const r = await client.api(url).get();
      all.push.apply(all, (r.value || []));
      pages++;
      url = r['@odata.nextLink'] ? r['@odata.nextLink'].replace('https://graph.microsoft.com/v1.0', '') : null;
    }
    const rejeitadas = all.map(function (it) { return norm(it, inv); })
      .filter(function (n) { return String(n.Status) === 'Rejeitada'; })
      .filter(function (n) { return !qUnid || n.Unidade === qUnid; })
      .filter(function (n) { return !qDir || n.Diretoria === qDir; });

    // Cache de listagem de pasta (evita re-listar a mesma pasta por nota).
    const cachePasta = {};
    async function listarPasta(folder) {
      if (cachePasta[folder] !== undefined) return cachePasta[folder];
      let arr = null;
      try {
        const resp = await client.api('/sites/' + siteId + '/drive/root:/' + folder + ':/children').get();
        arr = (resp.value || []).filter(function (x) { return x.file; });
      } catch (e) { arr = null; } // pasta inexistente
      cachePasta[folder] = arr;
      return arr;
    }

    const suspeitos = [], semArquivo = [], ok = [];
    for (const n of rejeitadas) {
      const numero = String(n.NumeroNF || '').trim();
      const valorNota = valorStrDe(n.Valor);
      const folder = (n.Unidade && n.Diretoria)
        ? 'Notas Fiscais/Rejeitadas/' + n.Unidade + '/Diretoria ' + n.Diretoria
        : 'Notas Fiscais/Rejeitadas';
      let files = await listarPasta(folder);
      if (!files) files = await listarPasta('Notas Fiscais/Rejeitadas'); // fallback raiz (legado)
      if (!files) { semArquivo.push({ id: n.id, numero: numero, valor: valorNota, unidade: n.Unidade, diretoria: n.Diretoria, motivo: 'pasta_inexistente' }); continue; }

      // Reproduz o match ANTIGO (frouxo por numero) pra ver o que a nota estava abrindo.
      const matched = numero ? files.filter(function (x) {
        return x.name && (x.name.startsWith(numero + '_') || x.name.indexOf('_' + numero + '_') >= 0);
      }) : [];
      if (!matched.length) { semArquivo.push({ id: n.id, numero: numero, valor: valorNota, unidade: n.Unidade, diretoria: n.Diretoria, motivo: 'sem_match_por_numero' }); continue; }

      // Ha algum arquivo com o VALOR da nota? (esse seria o correto)
      const certo = valorNota ? files.find(function (x) { return nomeTemValor(x.name, valorNota); }) : null;
      const apontado = matched[0]; // o que a logica antiga abriria (1o match frouxo)
      const valorApontado = valorDoNome(apontado.name);

      const registro = {
        id: n.id, numero: numero, valorNota: valorNota, unidade: n.Unidade, diretoria: n.Diretoria,
        fornecedor: n.CNPJFornecedor || n.Fornecedor, lancadoPor: n.LancadoPor,
        arquivoApontado: apontado.name, valorArquivoApontado: valorApontado,
        arquivoCorreto: certo ? certo.name : null
      };
      // MISLINK: o valor do arquivo apontado difere do valor da nota.
      if (valorNota && valorApontado && valorApontado !== valorNota) suspeitos.push(registro);
      else ok.push({ id: n.id, numero: numero, valor: valorNota, arquivo: apontado.name });
    }

    context.res = {
      status: 200, headers: { 'Content-Type': 'application/json', 'Cache-Control': 'no-store' },
      body: {
        ok: true, totalRejeitadas: rejeitadas.length,
        suspeitosMislink: suspeitos.length, suspeitos: suspeitos,
        semArquivoResolvido: semArquivo.length, semArquivo: semArquivo,
        conferemCount: ok.length
      }
    };
  } catch (err) {
    context.log && context.log.error && context.log.error('DiagAnexosRejeitados:', err);
    context.res = { status: 500, headers: { 'Content-Type': 'application/json' },
      body: { error: (err && err.message) || String(err) } };
  }
};
