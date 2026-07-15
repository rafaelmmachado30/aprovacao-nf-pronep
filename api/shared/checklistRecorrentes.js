/**
 * shared/checklistRecorrentes.js — Núcleo do Fechamento do Mês (reutilizável).
 *
 * Mesma lógica do endpoint /api/ChecklistRecorrentes, extraída para ser consumida também
 * pela SAN (tool fechamento_pendencias). Para cada conta recorrente confirmada, cruza com
 * as NFs do mês-alvo e devolve o status: atrasada | risco (D+5) | aguardando | lancada |
 * conciliada | aprovada | integrada.
 *
 * OBS: o endpoint mantém sua própria cópia por ora; se ambos divergirem, consolidar aqui.
 */

require('isomorphic-fetch');
const { lerDecisoes, chaveRecorrente, _norm } = require('./recorrentes');
const { lerConciliacoes, chaveConc } = require('./conciliacaoRecorrentes');

const LIST_FORN = 'PRONEP-NF-Fornecedores';
const D5_DIAS_UTEIS = 5;
const _fornCache = { byDoc: null, ts: 0 };
const _FORN_TTL = 5 * 60 * 1000;

function _normalizeItem(item, invColMap) {
  const f = item.fields || {};
  const out = { id: item.id };
  for (const [internal, val] of Object.entries(f)) { const d = invColMap[internal]; if (d) out[d] = val; }
  return out;
}
function _dataVenc(n) {
  for (const c of [n.DataVencimento, n.Vencimento]) {
    if (!c) continue;
    const d = new Date(String(c).substring(0, 10) + 'T00:00:00');
    if (!isNaN(d.getTime())) return d;
  }
  return null;
}
// Dias úteis (seg-sex) de hoje (exclusivo) até alvo (inclusivo). -1 se já passou.
function _diasUteisAte(hoje, alvo) {
  const d = new Date(hoje.getFullYear(), hoje.getMonth(), hoje.getDate());
  const fim = new Date(alvo.getFullYear(), alvo.getMonth(), alvo.getDate());
  if (fim < d) return -1;
  let count = 0;
  while (d < fim) { d.setDate(d.getDate() + 1); const dow = d.getDay(); if (dow !== 0 && dow !== 6) count++; }
  return count;
}
async function _fornecedoresIndex(client, siteId) {
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

/**
 * @param opts { scopeNorm?: string[]|null (diretorias normalizadas; null=todas), ano?, mes? (0-based) }
 * @returns { mes, total, resumo, contas }
 */
async function computar(client, siteId, listNotasId, invColMap, opts) {
  opts = opts || {};
  const scopeNorm = opts.scopeNorm || null;
  const agoraBRT = new Date(Date.now() - 3 * 60 * 60 * 1000);
  const ano = (opts.ano != null) ? opts.ano : agoraBRT.getUTCFullYear();
  const mes = (opts.mes != null) ? opts.mes : agoraBRT.getUTCMonth();
  const mesKeyAlvo = ano + '-' + String(mes + 1).padStart(2, '0');
  const ultimoDiaMes = new Date(ano, mes + 1, 0).getDate();
  const hojeBRT = new Date(agoraBRT.getUTCFullYear(), agoraBRT.getUTCMonth(), agoraBRT.getUTCDate());

  const decisoes = await lerDecisoes(client, siteId, null);
  if (!decisoes.listId) return { mes: mesKeyAlvo, total: 0, resumo: {}, contas: [], diretoriasDisponiveis: [], listaDecisoesExiste: false };
  let conciliacoes = {};
  try { conciliacoes = await lerConciliacoes(client, siteId, null); } catch (e) { conciliacoes = {}; }

  const fimMesAlvo = new Date(ano, mes, ultimoDiaMes);
  const confirmadas = decisoes.itens.filter(function (d) {
    if (!d.ehRecorrente || !d.ativo) return false;
    if (scopeNorm && scopeNorm.indexOf(_norm(d.diretoria)) < 0) return false;
    if (d.dataFim) {
      const df = new Date(String(d.dataFim).substring(0, 10) + 'T00:00:00');
      if (!isNaN(df.getTime()) && df < fimMesAlvo) return false;
    }
    return true;
  });

  const all = [];
  let url = '/sites/' + siteId + '/lists/' + listNotasId + '/items?expand=fields&$top=500';
  let pages = 0;
  while (url && pages < 30) {
    const resp = await client.api(url).get();
    all.push.apply(all, (resp.value || []));
    pages++;
    url = resp['@odata.nextLink'] ? resp['@odata.nextLink'].replace('https://graph.microsoft.com/v1.0', '') : null;
  }

  const nfsPorChave = {};
  const diretoriasSet = {};
  for (const item of all) {
    const n = _normalizeItem(item, invColMap);
    if (n.Diretoria) diretoriasSet[n.Diretoria] = true;
    const d = _dataVenc(n);
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

  const fornIdx = await _fornecedoresIndex(client, siteId);
  function nomeForn(cnpj, fallback) { const hit = fornIdx[cnpj]; return (hit && (hit.fantasia || hit.razao)) || fallback || cnpj; }

  const contas = confirmadas.map(function (d) {
    const dia = d.diaVencimento && d.diaVencimento >= 1 && d.diaVencimento <= 31 ? Math.min(d.diaVencimento, ultimoDiaMes) : ultimoDiaMes;
    const vencEsperado = new Date(ano, mes, dia);
    const nfs = (nfsPorChave[d.chave] || []).filter(function (x) { return String(x.status) !== 'Rejeitada'; });
    let nf = null;
    if (nfs.length) nf = nfs.sort(function (a, b) { return String(b.vencimento).localeCompare(String(a.vencimento)); })[0];
    let status, diasUteis = null;
    if (nf) {
      const s = String(nf.status || '');
      if (s === 'Aprovada') status = (nf.integrado || nf.processado) ? 'integrada' : 'aprovada';
      else status = 'lancada';
    } else {
      diasUteis = _diasUteisAte(hojeBRT, vencEsperado);
      if (diasUteis < 0) status = 'atrasada';
      else if (diasUteis <= D5_DIAS_UTEIS) status = 'risco';
      else status = 'aguardando';
    }
    const link = conciliacoes[chaveConc(d.chave, mesKeyAlvo)];
    if (link) { status = 'conciliada'; diasUteis = null; }
    return {
      chave: d.chave, cnpj: d.cnpj, fornecedor: d.fornecedor || nomeForn(d.cnpj, ''),
      diretoria: d.diretoria, unidade: d.unidade,
      vencEsperado: vencEsperado.toISOString().substring(0, 10),
      diaVencimento: d.diaVencimento || null,
      diasUteisAteVenc: diasUteis, valorEstimado: d.valorEstimado || null, status: status,
      nf: nf
        ? { id: nf.id, numero: nf.numero, status: nf.status, valor: nf.valor, vencimento: nf.vencimento }
        : (link ? { id: link.notaId, numero: link.numero, status: link.status, valor: link.valor, vencimento: null } : null),
      conciliada: link ? { numero: link.numero, por: link.por, em: link.em } : null
    };
  });

  const ordem = { atrasada: 0, risco: 1, aguardando: 3, lancada: 4, conciliada: 5, aprovada: 6, integrada: 7 };
  contas.sort(function (a, b) {
    const oa = ordem[a.status] != null ? ordem[a.status] : 9;
    const ob = ordem[b.status] != null ? ordem[b.status] : 9;
    if (oa !== ob) return oa - ob;
    return String(a.vencEsperado).localeCompare(String(b.vencEsperado));
  });
  const resumo = { atrasada: 0, risco: 0, aguardando: 0, lancada: 0, conciliada: 0, aprovada: 0, integrada: 0 };
  contas.forEach(function (c) { if (resumo[c.status] != null) resumo[c.status]++; });

  const diretoriasDisponiveis = Object.keys(diretoriasSet).sort(function (a, b) { return a.localeCompare(b); });
  return { mes: mesKeyAlvo, total: contas.length, resumo: resumo, contas: contas, diretoriasDisponiveis: diretoriasDisponiveis, listaDecisoesExiste: true };
}

module.exports = { computar };
