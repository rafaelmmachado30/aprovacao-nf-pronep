/**
 * /api/DiagIntegracaoOmie  (GET) — READ-ONLY. RBAC: admin ou financeiro_nf.
 *
 * Audita as NFs ja integradas no Omie pra achar as que PODEM ter recebido o PDF ERRADO
 * (bug do match so por numero, corrigido no #55). Para cada NF integrada:
 *   - correto : PDF resolvido por IDENTIDADE EXATA (acharPdfAlvo: URL aprovada -> num+valor)
 *   - porNumero: TODOS os PDFs aprovados (de qualquer NF) que a regra ANTIGA (so numero)
 *               casaria; o [0] e o "provavel anexado".
 *   - risco   : se porNumero tem >1 candidato, OU o provavel anexado != correto, OU nao ha correto.
 *
 * IMPORTANTE (perf): NAO varre o drive do SharePoint (isso estourava o timeout da function).
 * O universo de PDFs aprovados e montado a partir da PROPRIA lista de notas — cada nota
 * aprovada carrega o nome do seu PDF no campo UrlPDFAprovado. Isso captura exatamente as
 * colisoes de numero (que sao NFs diferentes com o mesmo numero) sem nenhuma chamada extra.
 *
 * NAO altera nada. Query: ?limite= (default 5000)  ?amostra= (itens listados, default 1000)
 *                        ?format=html (tabela legivel)
 */

require('isomorphic-fetch');
const { resolveAuthz } = require('../shared/authz');
const { getGraphClient, resolveSiteId } = require('../shared/graph');
const { acharPdfAlvo, urlDeCampo, nomeArquivoDeUrl } = require('../shared/pdfNota');

const LIST_NOTAS = 'PRONEP-NF-NotasFiscais';

async function resolveLista(client, siteId) {
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
function ehIntegrada(n) {
  return n.IntegradoOmie === true || n.IntegradoOmie === 'Sim' || n.IntegradoOmie === 'true';
}
function urlPdfDe(n) {
  return urlDeCampo(n.UrlPDFAprovadoStr) || urlDeCampo(n.UrlPDFAprovado) || urlDeCampo(n.UrlPDFStr) || urlDeCampo(n.UrlPDF) || '';
}

module.exports = async function (context, req) {
  try {
    const authz = await resolveAuthz(req);
    if (!authz) { context.res = { status: 401, body: { error: 'Nao autenticado' } }; return; }
    if (!authz.isAdmin && !authz.isFinanceiro) { context.res = { status: 403, body: { error: 'Acesso restrito a admin ou financeiro' } }; return; }

    const limite = Math.min(8000, Math.max(1, parseInt((req.query && req.query.limite) || '5000', 10) || 5000));
    const amostra = Math.min(2000, Math.max(1, parseInt((req.query && req.query.amostra) || '1000', 10) || 1000));

    let normalizaNumeroNF;
    try { normalizaNumeroNF = require('../shared/omie').normalizaNumeroNF; }
    catch (e) { normalizaNumeroNF = function (s) { return String(s || '').replace(/\D/g, '').replace(/^0+/, ''); }; }

    const client = await getGraphClient();
    const siteId = await resolveSiteId(client);
    const { listId, inv } = await resolveLista(client, siteId);

    // Le todas as NFs (paginado). So a lista — sem tocar no drive.
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

    // Universo de PDFs aprovados por unidade, montado da propria lista.
    // Cada entrada { id, name } imita um arquivo do drive pra reusar acharPdfAlvo/porNumero.
    const universoPorUnidade = {};
    for (const n of notas) {
      const u = urlPdfDe(n);
      const nome = u ? nomeArquivoDeUrl(u) : '';
      if (!nome) continue;
      const uni = n.Unidade || '';
      (universoPorUnidade[uni] = universoPorUnidade[uni] || []).push({ id: n.id, name: nome });
    }

    function candidatosPorNumero(files, numero) {
      const numFmt = normalizaNumeroNF(numero);
      if (!numFmt) return [];
      return (files || []).filter(function (x) {
        if (!x.name) return false;
        const base = x.name.replace(/\.pdf$/i, '');
        const tokens = base.split(/[_\-\s]+/);
        if (tokens.some(function (t) { return normalizaNumeroNF(t) === numFmt; })) return true;
        return normalizaNumeroNF(base).indexOf(numFmt) >= 0;
      });
    }

    const integradas = notas.filter(ehIntegrada);
    const suspeitos = [], ok = [];
    for (const n of integradas) {
      const files = universoPorUnidade[n.Unidade || ''] || [];
      const urlAprov = urlPdfDe(n);
      const achado = acharPdfAlvo(files, { url: urlAprov, numero: n.NumeroNF, valor: n.Valor });
      const correto = achado.target;
      const porNum = candidatosPorNumero(files, n.NumeroNF);
      const provavelAnexado = porNum[0] || null;

      const risco = (porNum.length > 1) || (!correto) || (provavelAnexado && correto && provavelAnexado.id !== correto.id);
      const reg = {
        id: n.id, numero: String(n.NumeroNF || ''), valor: n.Valor, unidade: n.Unidade, diretoria: n.Diretoria,
        fornecedor: n.CNPJFornecedor || n.Fornecedor,
        correto: correto ? correto.name : null, matchCorretoPor: achado.matchPor,
        provavelAnexadoAntigo: provavelAnexado ? provavelAnexado.name : null,
        candidatosPorNumero: porNum.length,
        motivo: !correto ? 'sem_pdf_correto_identificavel'
          : (porNum.length > 1 ? 'numero_ambiguo_multiplos_arquivos'
          : (provavelAnexado && provavelAnexado.id !== correto.id ? 'antigo_pegaria_outro_arquivo' : 'ok'))
      };
      if (risco) suspeitos.push(reg); else ok.push(reg);
    }

    const lista = suspeitos.slice(0, amostra);

    if ((req.query && req.query.format) === 'html') {
      const esc = function (s) { return String(s == null ? '' : s).replace(/[&<>"]/g, function (c) { return ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c]; }); };
      const motivoLabel = {
        sem_pdf_correto_identificavel: 'PDF correto NAO identificavel com seguranca',
        numero_ambiguo_multiplos_arquivos: 'Numero casa com VARIOS arquivos (ambiguo)',
        antigo_pegaria_outro_arquivo: 'Regra antiga anexaria OUTRO arquivo'
      };
      const linhas = lista.map(function (r) {
        return '<tr>' +
          '<td>' + esc(r.numero) + '</td>' +
          '<td>' + esc(r.unidade) + '</td>' +
          '<td>' + esc(r.diretoria) + '</td>' +
          '<td>' + esc(r.fornecedor) + '</td>' +
          '<td>' + esc(r.valor) + '</td>' +
          '<td class="ok">' + esc(r.correto || '—') + '</td>' +
          '<td class="warn">' + esc(r.provavelAnexadoAntigo || '—') + '</td>' +
          '<td>' + esc(motivoLabel[r.motivo] || r.motivo) + '</td>' +
          '</tr>';
      }).join('');
      const html = '<!doctype html><html lang="pt-BR"><head><meta charset="utf-8">' +
        '<meta name="viewport" content="width=device-width,initial-scale=1">' +
        '<title>Auditoria Integracao Omie</title><style>' +
        'body{font-family:system-ui,Segoe UI,Arial,sans-serif;margin:24px;color:#1f2937}' +
        'h1{font-size:20px;margin:0 0 4px}p{color:#4b5563;margin:4px 0}' +
        '.cards{display:flex;gap:12px;margin:16px 0}' +
        '.card{border:1px solid #e5e7eb;border-radius:10px;padding:12px 16px;min-width:120px}' +
        '.card b{display:block;font-size:24px}.card.sus b{color:#b45309}.card.ok b{color:#047857}' +
        'table{border-collapse:collapse;width:100%;font-size:13px;margin-top:8px}' +
        'th,td{border:1px solid #e5e7eb;padding:6px 8px;text-align:left;vertical-align:top}' +
        'th{background:#f9fafb}td.ok{color:#047857}td.warn{color:#b45309;font-weight:600}' +
        '.empty{padding:24px;background:#ecfdf5;border:1px solid #a7f3d0;border-radius:10px;color:#065f46}' +
        '</style></head><body>' +
        '<h1>Auditoria da Integracao Omie — PDFs anexados</h1>' +
        '<p>Read-only. Compara o PDF <b>correto</b> (identidade exata: URL aprovada + numero+valor) com o que a regra <b>antiga</b> (so numero) teria anexado. As linhas abaixo sao as que precisam de conferencia no Omie.</p>' +
        '<div class="cards">' +
        '<div class="card"><span>Integradas</span><b>' + integradas.length + '</b></div>' +
        '<div class="card sus"><span>Suspeitas</span><b>' + suspeitos.length + '</b></div>' +
        '<div class="card ok"><span>OK</span><b>' + ok.length + '</b></div>' +
        '</div>' +
        (lista.length
          ? ('<table><thead><tr><th>NF</th><th>Unidade</th><th>Diretoria</th><th>Fornecedor</th><th>Valor</th><th>PDF correto</th><th>Provavel anexado (antigo)</th><th>Motivo</th></tr></thead><tbody>' + linhas + '</tbody></table>')
          : '<div class="empty">Nenhuma integracao suspeita encontrada. Todos os PDFs anexados batem com a identidade exata.</div>') +
        '</body></html>';
      context.res = { status: 200, headers: { 'Content-Type': 'text/html; charset=utf-8', 'Cache-Control': 'no-store' }, body: html };
      return;
    }

    context.res = {
      status: 200, headers: { 'Content-Type': 'application/json', 'Cache-Control': 'no-store' },
      body: {
        ok: true,
        _obs: 'READ-ONLY. Universo montado da lista de notas (campo UrlPDFAprovado), sem varrer o drive. Suspeitos = conferir/corrigir o anexo no Omie.',
        totalNotas: notas.length,
        totalIntegradas: integradas.length,
        suspeitosCount: suspeitos.length,
        okCount: ok.length,
        suspeitos: lista
      }
    };
  } catch (err) {
    context.log && context.log.error && context.log.error('DiagIntegracaoOmie:', err);
    context.res = { status: 500, headers: { 'Content-Type': 'application/json' }, body: { error: (err && err.message) || String(err) } };
  }
};
