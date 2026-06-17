/**
 * /api/ChecklistRecorrentes  (GET)  — Estagio 2.
 *
 * Para cada conta RECORRENTE confirmada (PRONEP-NF-Recorrentes, EhRecorrente=Sim,
 * Ativo, dentro da DataFim), cruza com as NFs do mes-alvo e calcula o status:
 *   - 'aprovada'   : NF do mes ja aprovada (ou integrada/processada)
 *   - 'lancada'    : NF do mes lancada, aguardando aprovacao
 *   - 'rejeitada'  : NF do mes rejeitada (precisa reenviar)
 *   - 'atrasada'   : sem NF e o vencimento esperado JA passou
 *   - 'risco'      : sem NF e o vencimento esperado entra no prazo D+5 (<=5 dias uteis)
 *   - 'aguardando' : sem NF e ainda nao chegou a hora
 *
 * RBAC: admin (todas) | gestor (sua(s) diretoria(s)) | demais 403.
 * Query: ?diretoria=Tecnologia ?mes=AAAA-MM (default: mes corrente, BRT)
 */

require('isomorphic-fetch');
const { getUser } = require('../shared/auth');
const { getUserRoles, ROLE_LABELS } = require('../shared/userRoles');
const { isAdminEmail } = require('../shared/authz');
const { lerDecisoes, chaveRecorrente, _norm } = require('../shared/recorrentes');
const { ClientSecretCredential } = require('@azure/identity');
const { Client } = require('@microsoft/microsoft-graph-client');
const { TokenCredentialAuthenticationProvider } =
  require('@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials');

const LIST_NOTAS = 'PRONEP-NF-NotasFiscais';
const LIST_FORN = 'PRONEP-NF-Fornecedores';
const cache = { siteId: null, listNotasId: null, invColMap: null };
const _fornCache = { byDoc: null, ts: 0 };
const _FORN_TTL = 5 * 60 * 1000;
const D5_DIAS_UTEIS = 5;

function getGraphClient() {
  const tenantId = process.env.AAD_TENANT_ID;
  const clientId = process.env.AAD_CLIENT_ID;
  const clientSecret = process.env.AAD_CLIENT_SECRET;
  if (!tenantId || !clientId || !clientSecret) throw new Error('AAD_* incompletas');
  const credential = new ClientSecretCredential(tenantId, clientId, clientSecret);
  const authProvider = new TokenCredentialAuthenticationProvider(credential, {
    scopes: ['https://graph.microsoft.com/.default']
  });
  return Client.initWithMiddleware({ authProvider });
}

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
  for (const c of (resp.value || [])) { if (c.displayName && c.name) inv[c.name] = c.displayName; }
  cache.invColMap = inv;
  return inv;
}

function normalizeItem(item, invColMap) {
  const f = item.fields || {};
  const out = { id: item.id };
  for (const [internal, val] of Object.entries(f)) { const d = invColMap[internal]; if (d) out[d] = val; }
  return out;
}

function readClientPrincipalRoles(req) {
  const header = req.headers && req.headers['x-ms-client-principal'];
  if (!header) return [];
  try { const p = JSON.parse(Buffer.from(header, 'base64').toString('utf-8')); return (p && p.userRoles) || []; }
  catch (e) { return []; }
}

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
  } catch (e) { /* ignore */ }
  _fornCache.byDoc = byDoc; _fornCache.ts = Date.now();
  return byDoc;
}

function dataVenc(n) {
  const cands = [n.DataVencimento, n.Vencimento];
  for (const c of cands) {
    if (!c) continue;
    const d = new Date(String(c).substring(0, 10) + 'T00:00:00');
    if (!isNaN(d.getTime())) return d;
  }
  return null;
}

// Conta dias uteis (seg-sex) de hoje (exclusivo) ate alvo (inclusivo).
// Retorna -1 se o alvo ja passou.
function diasUteisAte(hoje, alvo) {
  const d = new Date(hoje.getFullYear(), hoje.getMonth(), hoje.getDate());
  const fim = new Date(alvo.getFullYear(), alvo.getMonth(), alvo.getDate());
  if (fim < d) return -1;
  let count = 0;
  while (d < fim) {
    d.setDate(d.getDate() + 1);
    const dow = d.getDay();
    if (dow !== 0 && dow !== 6) count++;
  }
  return count;
}

module.exports = async function (context, req) {
  const diag = { step: 'start' };
  try {
    const user = await getUser(req);
    if (!user) { context.res = { status: 401, body: { error: 'Nao autenticado' } }; return; }
    const userEmail = (user.email || '').toLowerCase();

    const claimsRoles = (user.claims && user.claims.roles) || [];
    const principalRoles = (user.source === 'easy-auth' ? readClientPrincipalRoles(req) : []) || [];
    const usefulPrincipalRoles = principalRoles.filter(function (r) { return r !== 'authenticated' && r !== 'anonymous'; });
    const graphRoles = await getUserRoles(user);
    const userRoles = Array.from(new Set([].concat(claimsRoles, usefulPrincipalRoles, graphRoles)));
    const isAdmin = userRoles.includes('administrador') || isAdminEmail(userEmail);
    const gestorLabels = userRoles.filter(function (r) { return r.indexOf('gestor') === 0; })
      .map(function (r) { return ROLE_LABELS[r]; }).filter(Boolean);
    if (!isAdmin && gestorLabels.length === 0) {
      context.res = { status: 403, body: { error: 'Apenas gestores ou admin' } }; return;
    }

    const qDiretoria = (req.query && req.query.diretoria) || '';
    let scopeDiretorias = null;
    if (qDiretoria && qDiretoria !== '__todas__') scopeDiretorias = [qDiretoria];
    else if (qDiretoria === '__todas__') scopeDiretorias = null;
    else if (isAdmin) scopeDiretorias = null;
    else if (gestorLabels.length) scopeDiretorias = gestorLabels;
    const scopeNorm = scopeDiretorias ? scopeDiretorias.map(_norm) : null;

    // Mes-alvo (BRT). Default: mes corrente.
    const agoraBRT = new Date(Date.now() - 3 * 60 * 60 * 1000);
    let ano = agoraBRT.getUTCFullYear();
    let mes = agoraBRT.getUTCMonth(); // 0-based
    const qMes = (req.query && req.query.mes) || '';
    const mMes = qMes.match(/^(\d{4})-(\d{2})$/);
    if (mMes) { ano = Number(mMes[1]); mes = Number(mMes[2]) - 1; }
    const mesKeyAlvo = ano + '-' + String(mes + 1).padStart(2, '0');
    const ultimoDiaMes = new Date(ano, mes + 1, 0).getDate();
    const hojeBRT = new Date(agoraBRT.getUTCFullYear(), agoraBRT.getUTCMonth(), agoraBRT.getUTCDate());
    diag.mesAlvo = mesKeyAlvo;

    diag.step = 'decisoes';
    const client = getGraphClient();
    const { siteId, listNotasId } = await resolveSite(client);
    const decisoes = await lerDecisoes(client, siteId, null);
    if (!decisoes.listId) {
      context.res = { status: 200, headers: { 'Content-Type': 'application/json' },
        body: { ok: true, listaDecisoesExiste: false, mes: mesKeyAlvo, contas: [], resumo: {}, diretoriasDisponiveis: [] } };
      return;
    }

    // Filtra decisoes: recorrente=Sim, ativo, dentro da escopo e da DataFim
    const fimMesAlvo = new Date(ano, mes, ultimoDiaMes);
    const confirmadas = decisoes.itens.filter(function (d) {
      if (!d.ehRecorrente || !d.ativo) return false;
      if (scopeNorm && scopeNorm.indexOf(_norm(d.diretoria)) < 0) return false;
      if (d.dataFim) {
        const df = new Date(String(d.dataFim).substring(0, 10) + 'T00:00:00');
        if (!isNaN(df.getTime()) && df < fimMesAlvo) return false; // lembrete encerrado antes do mes
      }
      return true;
    });
    diag.confirmadas = confirmadas.length;

    diag.step = 'fetch_notas';
    const invColMap = await getInvColMap(client, siteId, listNotasId);
    const all = [];
    let url = '/sites/' + siteId + '/lists/' + listNotasId + '/items?expand=fields&$top=500';
    let pages = 0;
    while (url && pages < 30) {
      const resp = await client.api(url).get();
      all.push.apply(all, (resp.value || []));
      pages++;
      url = resp['@odata.nextLink'] ? resp['@odata.nextLink'].replace('https://graph.microsoft.com/v1.0', '') : null;
    }

    // Indexa NFs do mes-alvo por chave (cnpj|diretoria|unidade)
    diag.step = 'index_notas';
    const nfsPorChave = {};
    const diretoriasSet = {};
    for (const item of all) {
      const n = normalizeItem(item, invColMap);
      if (n.Diretoria) diretoriasSet[n.Diretoria] = true;
      const d = dataVenc(n);
      if (!d) continue;
      const mesKey = d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0');
      if (mesKey !== mesKeyAlvo) continue;
      const cnpj = String(n.CNPJFornecedor || '').replace(/\D/g, '');
      if (!cnpj) continue;
      const chave = chaveRecorrente(cnpj, n.Diretoria || '', n.Unidade || '');
      const reg = {
        id: n.id, numero: n.NumeroNF, status: n.Status, valor: Number(n.Valor || 0) || 0,
        vencimento: d.toISOString().substring(0, 10),
        integrado: (n.IntegradoOmie === true || n.IntegradoOmie === 'Sim'),
        processado: (n.Processado === true || n.Processado === 'Sim')
      };
      if (!nfsPorChave[chave]) nfsPorChave[chave] = [];
      nfsPorChave[chave].push(reg);
    }

    diag.step = 'fornecedores';
    const fornIdx = await getFornecedoresIndex(client, siteId);
    function nomeForn(cnpj, fallback) {
      const hit = fornIdx[cnpj];
      return (hit && (hit.fantasia || hit.razao)) || fallback || cnpj;
    }

    diag.step = 'montar';
    const contas = confirmadas.map(function (d) {
      const dia = d.diaVencimento && d.diaVencimento >= 1 && d.diaVencimento <= 31 ? Math.min(d.diaVencimento, ultimoDiaMes) : ultimoDiaMes;
      const vencEsperado = new Date(ano, mes, dia);
      const nfs = nfsPorChave[d.chave] || [];
      let nf = null;
      if (nfs.length) {
        // prioriza nao-rejeitada / mais recente
        const naoRej = nfs.filter(function (x) { return String(x.status) !== 'Rejeitada'; });
        const pool = naoRej.length ? naoRej : nfs;
        nf = pool.sort(function (a, b) { return String(b.vencimento).localeCompare(String(a.vencimento)); })[0];
      }
      let status, diasUteis = null;
      if (nf) {
        const s = String(nf.status || '');
        if (s === 'Aprovada') status = (nf.integrado || nf.processado) ? 'integrada' : 'aprovada';
        else if (s === 'Rejeitada') status = 'rejeitada';
        else status = 'lancada'; // Lancada / AguardandoN2 / outros
      } else {
        diasUteis = diasUteisAte(hojeBRT, vencEsperado);
        if (diasUteis < 0) status = 'atrasada';
        else if (diasUteis <= D5_DIAS_UTEIS) status = 'risco';
        else status = 'aguardando';
      }
      return {
        chave: d.chave, cnpj: d.cnpj,
        fornecedor: d.fornecedor || nomeForn(d.cnpj, ''),
        diretoria: d.diretoria, unidade: d.unidade,
        vencEsperado: vencEsperado.toISOString().substring(0, 10),
        diaVencimento: d.diaVencimento || null,
        valorEstimado: d.valorEstimado || null,
        diasUteisAteVenc: diasUteis,
        status: status,
        nf: nf ? { id: nf.id, numero: nf.numero, status: nf.status, valor: nf.valor, vencimento: nf.vencimento } : null
      };
    });

    // Ordena por urgencia
    const ordem = { atrasada: 0, risco: 1, rejeitada: 2, aguardando: 3, lancada: 4, aprovada: 5, integrada: 6 };
    contas.sort(function (a, b) {
      const oa = ordem[a.status] != null ? ordem[a.status] : 9;
      const ob = ordem[b.status] != null ? ordem[b.status] : 9;
      if (oa !== ob) return oa - ob;
      return String(a.vencEsperado).localeCompare(String(b.vencEsperado));
    });

    const resumo = { atrasada: 0, risco: 0, aguardando: 0, lancada: 0, aprovada: 0, integrada: 0, rejeitada: 0 };
    contas.forEach(function (c) { if (resumo[c.status] != null) resumo[c.status]++; });

    diag.step = 'done';
    context.res = {
      status: 200, headers: { 'Content-Type': 'application/json', 'Cache-Control': 'no-store' },
      body: {
        ok: true, listaDecisoesExiste: true, mes: mesKeyAlvo,
        scope: { isAdmin: isAdmin, diretorias: scopeDiretorias },
        diretoriasDisponiveis: Object.keys(diretoriasSet).sort(function (a, b) { return a.localeCompare(b); }),
        total: contas.length, resumo: resumo, contas: contas, diag: diag
      }
    };
  } catch (err) {
    context.log && context.log.error && context.log.error('ChecklistRecorrentes error:', err);
    context.res = { status: 500, headers: { 'Content-Type': 'application/json' },
      body: { error: (err && err.message) || String(err), diag: diag } };
  }
};
