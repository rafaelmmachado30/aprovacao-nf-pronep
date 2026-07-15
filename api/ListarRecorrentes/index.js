/**
 * /api/ListarRecorrentes  (GET)
 *
 * Estagio 1 do modulo de Contas Recorrentes.
 * Varre PRONEP-NF-NotasFiscais dos ultimos N meses, restringe a(s) diretoria(s) do
 * gestor logado, agrupa por CNPJ+Unidade e detecta recorrencia (fornecedor que
 * aparece em >= MIN_MESES dos ultimos JANELA_MESES). Calcula dia tipico de vencimento
 * e valor medio. Mescla com as DECISOES ja salvas (PRONEP-NF-Recorrentes).
 *
 * RBAC:
 *   - admin: ve todas as diretorias (ou ?diretoria= pra filtrar)
 *   - gestor: ve so a(s) diretoria(s) do(s) seu(s) grupo(s) gestor_*
 *   - demais: 403
 *
 * Query: ?diretoria=Tecnologia (opcional, admin) ?meses=6 (opcional)
 *
 * Resposta: { ok, scope, params, contas: [ ... ], diag }
 */

require('isomorphic-fetch');
const { getUser } = require('../shared/auth');
const { getUserRoles, ROLE_LABELS } = require('../shared/userRoles');
const { isAdminEmail } = require('../shared/authz');
const { lerDecisoes, chaveRecorrente, _norm } = require('../shared/recorrentes');
const { getGraphClient } = require('../shared/graph');

const LIST_NOTAS = 'PRONEP-NF-NotasFiscais';
const LIST_FORN = 'PRONEP-NF-Fornecedores';
const cache = { siteId: null, listNotasId: null, invColMap: null };
const _fornCache = { byDoc: null, ts: 0 };
const _FORN_TTL = 5 * 60 * 1000;

const JANELA_MESES_DEFAULT = 6;
const MIN_MESES = 3; // aparece em >= 3 dos ultimos N meses => recorrente


async function resolveSite(client) {
  if (cache.siteId && cache.listNotasId) return cache;
  const host = process.env.SHAREPOINT_SITE_HOSTNAME;
  const path = process.env.SHAREPOINT_SITE_PATH;
  if (!host || !path) throw new Error('SHAREPOINT_* incompletas');
  const siteResp = await client.api('/sites/' + host + ':' + path).get();
  cache.siteId = siteResp.id;
  const lists = await client.api('/sites/' + cache.siteId + '/lists').filter("displayName eq '" + LIST_NOTAS + "'").get();
  if (!lists.value || !lists.value.length) throw new Error('Lista ' + LIST_NOTAS + ' nao encontrada');
  cache.listNotasId = lists.value[0].id;
  return cache;
}

async function getInvColMap(client, siteId, listId) {
  if (cache.invColMap) return cache.invColMap;
  const resp = await client.api('/sites/' + siteId + '/lists/' + listId + '/columns').get();
  const inv = {};
  for (const c of (resp.value || [])) {
    if (!c.displayName || !c.name) continue;
    inv[c.name] = c.displayName;
  }
  cache.invColMap = inv;
  return inv;
}

function normalizeItem(item, invColMap) {
  const f = item.fields || {};
  const out = { id: item.id };
  for (const [internal, val] of Object.entries(f)) {
    const display = invColMap[internal];
    if (display) out[display] = val;
  }
  return out;
}

function readClientPrincipalRoles(req) {
  const header = req.headers && req.headers['x-ms-client-principal'];
  if (!header) return [];
  try { const p = JSON.parse(Buffer.from(header, 'base64').toString('utf-8')); return (p && p.userRoles) || []; }
  catch (e) { return []; }
}

// Index de fornecedores CNPJ(digitos) -> {razao,fantasia}
async function getFornecedoresIndex(client, siteId) {
  if (_fornCache.byDoc && (Date.now() - _fornCache.ts) < _FORN_TTL) return _fornCache.byDoc;
  const byDoc = {};
  try {
    const fl = await client.api('/sites/' + siteId + '/lists').filter("displayName eq '" + LIST_FORN + "'").get();
    if (fl.value && fl.value.length) {
      const flId = fl.value[0].id;
      const colResp = await client.api('/sites/' + siteId + '/lists/' + flId + '/columns').get();
      const inv = {};
      for (const c of (colResp.value || [])) { if (c.displayName && c.name) inv[c.name] = c.displayName; }
      let url = '/sites/' + siteId + '/lists/' + flId + '/items?expand=fields&$top=500';
      let pages = 0;
      while (url && pages < 30) {
        const r = await client.api(url).get();
        for (const it of (r.value || [])) {
          const f = {}; for (const [k, v] of Object.entries(it.fields || {})) { const d = inv[k]; if (d) f[d] = v; }
          const doc = String(f.Documento || f.CNPJ || f.field_2 || '').replace(/\D/g, '');
          if (doc && !byDoc[doc]) byDoc[doc] = { razao: f.Title || f.Razao || f.RazaoSocial || '', fantasia: f.NomeFantasia || f.field_3 || '' };
        }
        pages++;
        url = r['@odata.nextLink'] ? r['@odata.nextLink'].replace('https://graph.microsoft.com/v1.0', '') : null;
      }
    }
  } catch (e) { /* sem index -> usa CNPJ */ }
  _fornCache.byDoc = byDoc; _fornCache.ts = Date.now();
  return byDoc;
}

// Extrai uma Date a partir dos campos de data da NF (prioriza vencimento).
function dataDaNota(n) {
  const cands = [n.DataVencimento, n.Vencimento, n.DataEmissao, n.Emissao, n.LancadoEm];
  for (const c of cands) {
    if (!c) continue;
    const s = String(c).substring(0, 10);
    const d = new Date(s + 'T00:00:00');
    if (!isNaN(d.getTime())) return d;
  }
  return null;
}

function mediana(arr) {
  if (!arr.length) return null;
  const a = arr.slice().sort(function (x, y) { return x - y; });
  const m = Math.floor(a.length / 2);
  return a.length % 2 ? a[m] : Math.round((a[m - 1] + a[m]) / 2);
}

module.exports = async function (context, req) {
  const diag = { step: 'start' };
  try {
    const user = await getUser(req);
    if (!user) { context.res = { status: 401, body: { error: 'Nao autenticado' } }; return; }
    const userEmail = (user.email || '').toLowerCase();

    diag.step = 'roles';
    const claimsRoles = (user.claims && user.claims.roles) || [];
    const principalRoles = (user.source === 'easy-auth' ? readClientPrincipalRoles(req) : []) || [];
    const usefulPrincipalRoles = principalRoles.filter(function (r) { return r !== 'authenticated' && r !== 'anonymous'; });
    const graphRoles = await getUserRoles(user);
    const userRoles = Array.from(new Set([].concat(claimsRoles, usefulPrincipalRoles, graphRoles)));
    const isAdmin = userRoles.includes('administrador') || isAdminEmail(userEmail);
    const gestorRoles = userRoles.filter(function (r) { return r.indexOf('gestor') === 0; });
    const gestorLabels = gestorRoles.map(function (r) { return ROLE_LABELS[r]; }).filter(Boolean);

    if (!isAdmin && gestorLabels.length === 0) {
      context.res = { status: 403, body: { error: 'Apenas gestores ou admin acessam contas recorrentes' } };
      return;
    }

    // Escopo de diretorias
    const qDiretoria = (req.query && req.query.diretoria) || '';
    let scopeDiretorias = null; // null = todas
    if (qDiretoria && qDiretoria !== '__todas__') scopeDiretorias = [qDiretoria];
    else if (qDiretoria === '__todas__') scopeDiretorias = null; // forca todas
    else if (isAdmin) scopeDiretorias = null;                    // admin: default todas
    else if (gestorLabels.length) scopeDiretorias = gestorLabels; // gestor: sua(s) diretoria(s)
    else scopeDiretorias = null;
    const scopeNorm = scopeDiretorias ? scopeDiretorias.map(_norm) : null;
    diag.scope = { isAdmin, gestorLabels, scopeDiretorias };

    const janelaMeses = Math.max(2, Math.min(24, Number((req.query && req.query.meses) || JANELA_MESES_DEFAULT)));

    diag.step = 'graph';
    const client = getGraphClient();
    const { siteId, listNotasId } = await resolveSite(client);
    const invColMap = await getInvColMap(client, siteId, listNotasId);

    diag.step = 'fetch_notas';
    const all = [];
    let url = '/sites/' + siteId + '/lists/' + listNotasId + '/items?expand=fields&$top=500';
    let pages = 0;
    while (url && pages < 30) {
      const resp = await client.api(url).get();
      all.push.apply(all, (resp.value || []));
      pages++;
      url = resp['@odata.nextLink'] ? resp['@odata.nextLink'].replace('https://graph.microsoft.com/v1.0', '') : null;
    }
    diag.totalNotas = all.length;

    // Janela: do 1o dia do mes (hoje - janelaMeses) ate hoje
    const hoje = new Date();
    const limite = new Date(hoje.getFullYear(), hoje.getMonth() - (janelaMeses - 1), 1);

    diag.step = 'agrupar';
    const grupos = {}; // chave(cnpj|diretoria|unidade) -> stats
    const diretoriasSet = {}; // todas as diretorias vistas (pro seletor), antes do escopo
    for (const item of all) {
      const n = normalizeItem(item, invColMap);
      const diretoria = n.Diretoria || '';
      if (diretoria) diretoriasSet[diretoria] = true;
      if (scopeNorm && scopeNorm.indexOf(_norm(diretoria)) < 0) continue;
      const cnpj = String(n.CNPJFornecedor || '').replace(/\D/g, '');
      if (!cnpj) continue;
      const unidade = n.Unidade || '';
      const d = dataDaNota(n);
      if (!d || d < limite) continue;
      const mesKey = d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0');
      const valor = Number(n.Valor || 0) || 0;
      const chave = chaveRecorrente(cnpj, diretoria, unidade);
      let g = grupos[chave];
      if (!g) {
        g = grupos[chave] = {
          chave: chave, cnpj: cnpj, diretoria: diretoria, unidade: unidade,
          meses: {}, diasVenc: [], valores: [], ultimaData: null, ultimoValor: null, qtdNotas: 0
        };
      }
      g.meses[mesKey] = true;
      g.diasVenc.push(d.getDate());
      g.valores.push(valor);
      g.qtdNotas++;
      if (!g.ultimaData || d > g.ultimaData) { g.ultimaData = d; g.ultimoValor = valor; }
    }

    diag.step = 'decisoes';
    const decisoes = await lerDecisoes(client, siteId, null); // {} se lista nao existe ainda
    diag.listaDecisoesExiste = !!decisoes.listId;

    diag.step = 'fornecedores';
    const fornIdx = await getFornecedoresIndex(client, siteId);

    function nomeForn(cnpj) {
      const hit = fornIdx[cnpj];
      return (hit && (hit.fantasia || hit.razao)) || '';
    }

    diag.step = 'montar';
    const porChave = {};
    // 1) a partir dos grupos detectados
    for (const chave of Object.keys(grupos)) {
      const g = grupos[chave];
      const recMeses = Object.keys(g.meses).length;
      const qtd = g.qtdNotas;
      const dec = decisoes.porChave[chave];
      // Forca da recorrencia:
      //  - 'candidata': aparece em >=2 meses OU tem >=2 NFs (sinal de repeticao)
      //  - 'sugestao' : apareceu so 1x na janela (historico curto) -> gestor decide
      let status;
      if (dec && dec.ehRecorrente) status = 'confirmada';
      else if (dec && !dec.ehRecorrente) status = 'descartada';
      else if (recMeses >= MIN_MESES || recMeses >= 2 || qtd >= 2) status = 'candidata';
      else status = 'sugestao';
      porChave[chave] = {
        chave: chave, cnpj: g.cnpj, fornecedor: nomeForn(g.cnpj) || g.cnpj,
        diretoria: g.diretoria, unidade: g.unidade,
        recorrenciaMeses: recMeses, janelaMeses: janelaMeses, qtdNotas: g.qtdNotas,
        diaVencimento: (dec && dec.diaVencimento) || mediana(g.diasVenc),
        valorEstimado: (dec && dec.valorEstimado) || Math.round((g.valores.reduce(function (a, b) { return a + b; }, 0) / (g.valores.length || 1)) * 100) / 100,
        ultimaData: g.ultimaData ? g.ultimaData.toISOString().substring(0, 10) : null,
        ultimoValor: g.ultimoValor,
        status: status,
        ehRecorrente: dec ? dec.ehRecorrente : null,
        dataFim: (dec && dec.dataFim) ? String(dec.dataFim).substring(0, 10) : '',
        ativo: dec ? dec.ativo : true,
        temDecisao: !!dec
      };
    }
    // 2) decisoes confirmadas que NAO apareceram na janela (ex: conta esperada que nao chegou)
    for (const dec of decisoes.itens) {
      if (porChave[dec.chave]) continue;
      if (scopeNorm && scopeNorm.indexOf(_norm(dec.diretoria)) < 0) continue;
      if (!dec.ehRecorrente) continue; // descartadas fora da janela nao precisam aparecer
      porChave[dec.chave] = {
        chave: dec.chave, cnpj: dec.cnpj, fornecedor: dec.fornecedor || nomeForn(dec.cnpj) || dec.cnpj,
        diretoria: dec.diretoria, unidade: dec.unidade,
        recorrenciaMeses: 0, janelaMeses: janelaMeses, qtdNotas: 0,
        diaVencimento: dec.diaVencimento, valorEstimado: dec.valorEstimado,
        ultimaData: null, ultimoValor: null,
        status: 'confirmada', ehRecorrente: true,
        dataFim: dec.dataFim ? String(dec.dataFim).substring(0, 10) : '',
        ativo: dec.ativo, temDecisao: true, semMovimentoNaJanela: true
      };
    }

    const contas = Object.keys(porChave).map(function (k) { return porChave[k]; });

    // Ordena: confirmadas, candidatas (mais recorrentes primeiro), sugestoes, descartadas
    const ordemStatus = { confirmada: 0, candidata: 1, sugestao: 2, descartada: 3 };
    contas.sort(function (a, b) {
      const sa = ordemStatus[a.status] != null ? ordemStatus[a.status] : 9;
      const sb = ordemStatus[b.status] != null ? ordemStatus[b.status] : 9;
      if (sa !== sb) return sa - sb;
      if (b.recorrenciaMeses !== a.recorrenciaMeses) return b.recorrenciaMeses - a.recorrenciaMeses;
      return String(a.fornecedor).localeCompare(String(b.fornecedor));
    });

    diag.step = 'done';
    context.res = {
      status: 200, headers: { 'Content-Type': 'application/json', 'Cache-Control': 'no-store' },
      body: {
        ok: true,
        scope: { isAdmin, diretorias: scopeDiretorias },
        diretoriasDisponiveis: Object.keys(diretoriasSet).sort(function (a, b) { return a.localeCompare(b); }),
        params: { janelaMeses: janelaMeses, minMeses: MIN_MESES },
        listaDecisoesExiste: !!decisoes.listId,
        total: contas.length,
        contas: contas,
        diag: diag
      }
    };
  } catch (err) {
    context.log && context.log.error && context.log.error('ListarRecorrentes error:', err);
    context.res = { status: 500, headers: { 'Content-Type': 'application/json' },
      body: { error: (err && err.message) || String(err), diag: diag } };
  }
};
