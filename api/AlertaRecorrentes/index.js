/**
 * /api/AlertaRecorrentes  (GET/POST) — Estagio 3.
 *
 * Disparado por cron do GitHub Actions (junto com o alerta da manha).
 * Detecta contas RECORRENTES confirmadas em situacao de risco e avisa, por e-mail,
 * QUEM confirmou a conta (confirmadoPor). Zero-spam: so envia se houver pendencia.
 *
 * Situacoes que geram alerta:
 *   - 'atrasada'  : venceu (vencimento esperado passou) e nao ha NF do mes
 *   - 'risco'     : vence em <= 5 dias uteis e nao ha NF do mes
 *   - 'rejeitada' : a NF do mes foi rejeitada (precisa reenviar)
 *
 * Auth: header X-Alerta-Secret == ALERTA_DIARIO_SECRET (mesmo secret do AlertaDiario).
 * Query:
 *   ?dry=1   -> nao envia, retorna o que enviaria (debug)
 *   ?mes=AAAA-MM (opcional, default mes corrente BRT)
 *
 * App Settings: AAD_*, SHAREPOINT_*, EMAIL_FROM_ADDRESS, ALERTA_DIARIO_SECRET,
 *   ALERTA_RECORRENTES_TESTE_EMAIL (opcional: manda tudo pra esse email).
 */

require('isomorphic-fetch');
const { getGraphClient } = require('../shared/graph');
const { lerDecisoes, chaveRecorrente } = require('../shared/recorrentes');

const LIST_NOTAS = 'PRONEP-NF-NotasFiscais';
const DEFAULT_FROM = 'datanalytics@pronep.com.br';
const SISTEMA_URL = 'https://purple-forest-09588fe10.7.azurestaticapps.net/';
const D5_DIAS_UTEIS = 5;
const _cache = { siteId: null, listNotasId: null, invColMap: null };

async function resolveSite(client) {
  if (_cache.siteId && _cache.listNotasId && _cache.invColMap) return _cache;
  const host = process.env.SHAREPOINT_SITE_HOSTNAME;
  const path = process.env.SHAREPOINT_SITE_PATH;
  const siteResp = await client.api('/sites/' + host + ':' + path).get();
  _cache.siteId = siteResp.id;
  const lists = await client.api('/sites/' + _cache.siteId + '/lists').filter("displayName eq '" + LIST_NOTAS + "'").get();
  if (!lists.value || !lists.value.length) throw new Error('Lista ' + LIST_NOTAS + ' nao encontrada');
  _cache.listNotasId = lists.value[0].id;
  const colResp = await client.api('/sites/' + _cache.siteId + '/lists/' + _cache.listNotasId + '/columns').get();
  _cache.invColMap = {};
  for (const c of (colResp.value || [])) { if (c.displayName && c.name) _cache.invColMap[c.name] = c.displayName; }
  return _cache;
}

function normalizeItem(item, invColMap) {
  const f = item.fields || {};
  const out = { id: item.id };
  for (const [internal, val] of Object.entries(f)) { const d = invColMap[internal]; if (d) out[d] = val; }
  return out;
}

function dataVenc(n) {
  for (const c of [n.DataVencimento, n.Vencimento]) {
    if (!c) continue;
    const d = new Date(String(c).substring(0, 10) + 'T00:00:00');
    if (!isNaN(d.getTime())) return d;
  }
  return null;
}

function diasUteisAte(hoje, alvo) {
  const d = new Date(hoje.getFullYear(), hoje.getMonth(), hoje.getDate());
  const fim = new Date(alvo.getFullYear(), alvo.getMonth(), alvo.getDate());
  if (fim < d) return -1;
  let count = 0;
  while (d < fim) { d.setDate(d.getDate() + 1); const dow = d.getDay(); if (dow !== 0 && dow !== 6) count++; }
  return count;
}

function escapeHtml(s) {
  return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}
function fmtBRL(v) {
  const n = Number(v || 0) || 0;
  return 'R$ ' + n.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}
function fmtDataBR(iso) {
  return iso ? String(iso).substring(0, 10).split('-').reverse().join('/') : '—';
}

function rotuloSituacao(status) {
  if (status === 'atrasada') return '⚠ ATRASADA (venceu sem lançamento)';
  if (status === 'risco') return '⚠ Risco D+5 (vence em breve)';
  if (status === 'rejeitada') return 'Rejeitada — precisa reenviar';
  return status;
}

function montarCorpoEmail(contas, mesKey) {
  const linhas = contas.map(function (c) {
    return '<tr>' +
      '<td style="padding:8px 10px;border-bottom:1px solid #eee">' + escapeHtml(c.fornecedor || c.cnpj) + '</td>' +
      '<td style="padding:8px 10px;border-bottom:1px solid #eee">' + escapeHtml(c.diretoria || '—') + ' / ' + escapeHtml(c.unidade || '—') + '</td>' +
      '<td style="padding:8px 10px;border-bottom:1px solid #eee">' + fmtDataBR(c.vencEsperado) + '</td>' +
      '<td style="padding:8px 10px;border-bottom:1px solid #eee;text-align:right">' + (c.valorEstimado ? fmtBRL(c.valorEstimado) : '—') + '</td>' +
      '<td style="padding:8px 10px;border-bottom:1px solid #eee;color:#92400E;font-weight:600">' + escapeHtml(rotuloSituacao(c.status)) + '</td>' +
      '</tr>';
  }).join('');
  return '' +
    '<div style="font-family:Segoe UI,Arial,sans-serif;color:#1f2937;max-width:680px">' +
    '<h2 style="color:#1B3A5B;margin-bottom:4px">Contas recorrentes — atenção</h2>' +
    '<p style="font-size:14px;color:#555;margin-top:0">As contas recorrentes abaixo precisam de atenção neste mês (' + escapeHtml(mesKey) + '). ' +
    'Lance/aprove antes do prazo D+5 do Financeiro para evitar renegociação.</p>' +
    '<table style="width:100%;border-collapse:collapse;font-size:13px;margin-top:10px">' +
    '<thead><tr style="background:#F4F7FA;text-align:left">' +
    '<th style="padding:8px 10px">Fornecedor</th><th style="padding:8px 10px">Dir./Un.</th>' +
    '<th style="padding:8px 10px">Venc. esperado</th><th style="padding:8px 10px;text-align:right">Valor estimado</th>' +
    '<th style="padding:8px 10px">Situação</th></tr></thead><tbody>' + linhas + '</tbody></table>' +
    '<p style="margin-top:18px"><a href="' + SISTEMA_URL + '" style="background:#1B3A5B;color:#fff;padding:10px 18px;border-radius:6px;text-decoration:none;font-size:14px">Abrir o sistema → Fechamento do Mês</a></p>' +
    '<p style="font-size:12px;color:#999;margin-top:16px">Você recebe este aviso porque confirmou estas contas como recorrentes. SAN · Sistema de Aprovação de NF — Pronep.</p>' +
    '</div>';
}

module.exports = async function (context, req) {
  const diag = { step: 'start', enviados: [], totalAlertas: 0 };
  try {
    // Auth por secret compartilhado (mesmo do AlertaDiario)
    const secret = (req.headers && req.headers['x-alerta-secret']) || (req.query && req.query.secret) || '';
    if (!process.env.ALERTA_DIARIO_SECRET || secret !== process.env.ALERTA_DIARIO_SECRET) {
      context.res = { status: 403, body: { error: 'Secret invalido' } };
      return;
    }
    if (process.env.ALERTA_RECORRENTES_DISABLED === 'true') {
      context.res = { status: 200, body: { ok: true, skipped: 'disabled' } };
      return;
    }
    const dry = (req.query && (req.query.dry === '1' || req.query.dry === 'true'));
    const testeEmail = process.env.ALERTA_RECORRENTES_TESTE_EMAIL || '';

    // Mes-alvo (BRT)
    const agoraBRT = new Date(Date.now() - 3 * 60 * 60 * 1000);
    let ano = agoraBRT.getUTCFullYear();
    let mes = agoraBRT.getUTCMonth();
    const qMes = (req.query && req.query.mes) || '';
    const mMes = qMes.match && qMes.match(/^(\d{4})-(\d{2})$/);
    if (mMes) { ano = Number(mMes[1]); mes = Number(mMes[2]) - 1; }
    const mesKey = ano + '-' + String(mes + 1).padStart(2, '0');
    const ultimoDiaMes = new Date(ano, mes + 1, 0).getDate();
    const hojeBRT = new Date(agoraBRT.getUTCFullYear(), agoraBRT.getUTCMonth(), agoraBRT.getUTCDate());
    diag.mesAlvo = mesKey;

    const client = getGraphClient();
    const { siteId, listNotasId, invColMap } = await resolveSite(client);

    diag.step = 'decisoes';
    const decisoes = await lerDecisoes(client, siteId, null);
    if (!decisoes.listId) { context.res = { status: 200, body: { ok: true, semLista: true } }; return; }
    const fimMesAlvo = new Date(ano, mes, ultimoDiaMes);
    const confirmadas = decisoes.itens.filter(function (d) {
      if (!d.ehRecorrente || !d.ativo) return false;
      if (d.dataFim) {
        const df = new Date(String(d.dataFim).substring(0, 10) + 'T00:00:00');
        if (!isNaN(df.getTime()) && df < fimMesAlvo) return false;
      }
      return true;
    });
    diag.confirmadas = confirmadas.length;

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
    const nfsPorChave = {};
    for (const item of all) {
      const n = normalizeItem(item, invColMap);
      const d = dataVenc(n);
      if (!d) continue;
      const mk = d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0');
      if (mk !== mesKey) continue;
      const cnpj = String(n.CNPJFornecedor || '').replace(/\D/g, '');
      if (!cnpj) continue;
      const chave = chaveRecorrente(cnpj, n.Diretoria || '', n.Unidade || '');
      (nfsPorChave[chave] = nfsPorChave[chave] || []).push({ status: n.Status });
    }

    diag.step = 'classificar';
    const porDestinatario = {}; // email -> [contas]
    for (const d of confirmadas) {
      const dia = (d.diaVencimento && d.diaVencimento >= 1 && d.diaVencimento <= 31) ? Math.min(d.diaVencimento, ultimoDiaMes) : ultimoDiaMes;
      const vencEsperado = new Date(ano, mes, dia);
      const nfs = nfsPorChave[d.chave] || [];
      let status = null;
      if (nfs.length) {
        const temNaoRej = nfs.some(function (x) { return String(x.status) !== 'Rejeitada'; });
        if (!temNaoRej) status = 'rejeitada'; // todas rejeitadas
        else status = null; // ja lancada/aprovada -> nao alerta
      } else {
        const du = diasUteisAte(hojeBRT, vencEsperado);
        if (du < 0) status = 'atrasada';
        else if (du <= D5_DIAS_UTEIS) status = 'risco';
        else status = null; // ainda nao chegou a hora
      }
      if (!status) continue;
      const dest = (d.confirmadoPor || '').toLowerCase().trim();
      if (!dest) continue;
      (porDestinatario[dest] = porDestinatario[dest] || []).push({
        fornecedor: d.fornecedor, cnpj: d.cnpj, diretoria: d.diretoria, unidade: d.unidade,
        vencEsperado: vencEsperado.toISOString().substring(0, 10), valorEstimado: d.valorEstimado, status: status
      });
    }

    diag.step = 'enviar';
    const fromAddress = process.env.EMAIL_FROM_ADDRESS || DEFAULT_FROM;
    for (const dest of Object.keys(porDestinatario)) {
      const contas = porDestinatario[dest];
      // ordena: atrasada, risco, rejeitada
      const ordem = { atrasada: 0, risco: 1, rejeitada: 2 };
      contas.sort(function (a, b) { return (ordem[a.status] || 9) - (ordem[b.status] || 9); });
      diag.totalAlertas += contas.length;
      const paraEmail = testeEmail || dest;
      const registro = { destinatarioReal: dest, enviadoPara: paraEmail, qtd: contas.length, status: contas.map(function (c) { return c.status; }) };
      if (dry) { diag.enviados.push(Object.assign({ dry: true }, registro)); continue; }
      try {
        await client.api('/users/' + fromAddress + '/sendMail').post({
          message: {
            subject: '⚠ Contas recorrentes precisando de atenção (' + contas.length + ')',
            body: { contentType: 'HTML', content: montarCorpoEmail(contas, mesKey) },
            toRecipients: [{ emailAddress: { address: paraEmail } }]
          },
          saveToSentItems: true
        });
        diag.enviados.push(registro);
      } catch (eMail) {
        diag.enviados.push(Object.assign({ erro: eMail.message }, registro));
      }
    }

    diag.step = 'done';
    context.res = {
      status: 200, headers: { 'Content-Type': 'application/json' },
      body: { ok: true, dry: !!dry, mes: mesKey, destinatarios: Object.keys(porDestinatario).length, totalAlertas: diag.totalAlertas, diag: diag }
    };
  } catch (err) {
    context.log && context.log.error && context.log.error('AlertaRecorrentes error:', err);
    context.res = { status: 500, headers: { 'Content-Type': 'application/json' }, body: { error: (err && err.message) || String(err), diag: diag } };
  }
};
