/**
 * AlertaDiario — HTTP Endpoint (disparado por GitHub Actions cron)
 *
 * Static Web Apps Managed Functions nao suporta TimerTrigger, entao usamos
 * cron do GitHub Actions (.github/workflows/alerta-diario.yml) que faz GET
 * neste endpoint nos horarios certos.
 *
 * URL: GET /api/AlertaDiario?tipo=manha|tarde|semanal
 *      Header: X-Alerta-Secret: <ALERTA_DIARIO_SECRET>
 *
 * Tipos:
 * - manha:   alerta da manha (seg-sex 9h BRT)
 * - tarde:   alerta da tarde (seg-qui 17h BRT)
 * - semanal: resumo semanal (sexta 17h BRT, substitui o "tarde")
 *
 * Pra cada gestor com NFs pendentes, monta email hibrido:
 *   - Template fixo: tabela com NFs ordenadas por vencimento, totais, D+5
 *   - Paragrafo IA: insight sobre concentracao, anomalias, tendencias
 *
 * Se gestor nao tem pendencias, NAO envia email (zero spam).
 *
 * App Settings:
 *   - AAD_TENANT_ID, AAD_CLIENT_ID, AAD_CLIENT_SECRET
 *   - SHAREPOINT_SITE_HOSTNAME, SHAREPOINT_SITE_PATH
 *   - EMAIL_FROM_ADDRESS (default datanalytics@pronep.com.br)
 *   - ANTHROPIC_API_KEY (primario, pro paragrafo de insight via Claude)
 *   - OPENAI_API_KEY (fallback automatico se Anthropic indisponivel)
 *   - ANTHROPIC_MODEL_HAIKU (opcional, default claude-haiku-4-5)
 *   - ALERTA_DIARIO_SECRET (obrigatorio - shared secret com GitHub Actions)
 *   - ALERTA_DIARIO_TESTE_EMAIL (opcional - envia tudo pra esse email; debug)
 *   - ALERTA_DIARIO_DISABLED ('true' pra pausar)
 */

require('isomorphic-fetch');
const { ClientSecretCredential } = require('@azure/identity');
const { Client } = require('@microsoft/microsoft-graph-client');
const { TokenCredentialAuthenticationProvider } = require('@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials');

const LIST_NOTAS = 'PRONEP-NF-NotasFiscais';
const DEFAULT_FROM = 'datanalytics@pronep.com.br';
const SISTEMA_URL = 'https://purple-forest-09588fe10.7.azurestaticapps.net/';

const _cache = { siteId: null, listNotasId: null, invColMap: null };

async function getGraphClient() {
  const credential = new ClientSecretCredential(
    process.env.AAD_TENANT_ID, process.env.AAD_CLIENT_ID, process.env.AAD_CLIENT_SECRET
  );
  const authProvider = new TokenCredentialAuthenticationProvider(credential, {
    scopes: ['https://graph.microsoft.com/.default']
  });
  return Client.initWithMiddleware({ authProvider });
}

async function resolveSiteAndList(client) {
  if (_cache.siteId && _cache.listNotasId && _cache.invColMap) return _cache;
  const host = process.env.SHAREPOINT_SITE_HOSTNAME;
  const path = process.env.SHAREPOINT_SITE_PATH;
  const siteResp = await client.api('/sites/' + host + ':' + path).get();
  _cache.siteId = siteResp.id;
  const listsResp = await client.api('/sites/' + _cache.siteId + '/lists').filter("displayName eq '" + LIST_NOTAS + "'").get();
  if (!listsResp.value || !listsResp.value.length) throw new Error('Lista nao encontrada');
  _cache.listNotasId = listsResp.value[0].id;
  const colsResp = await client.api('/sites/' + _cache.siteId + '/lists/' + _cache.listNotasId + '/columns').get();
  _cache.invColMap = {};
  for (const c of (colsResp.value || [])) {
    if (c.displayName && c.name) _cache.invColMap[c.name] = c.displayName;
  }
  return _cache;
}

// Carrega mapa CNPJ -> RazaoSocial da lista PRONEP-NF-Fornecedores (1x por instancia)
async function carregarMapFornecedores(client, siteId) {
  const lists = await client.api('/sites/' + siteId + '/lists').filter("displayName eq 'PRONEP-NF-Fornecedores'").get();
  if (!lists.value || !lists.value.length) return {};
  const listId = lists.value[0].id;
  const colsResp = await client.api('/sites/' + siteId + '/lists/' + listId + '/columns').get();
  const inv = {};
  for (const c of (colsResp.value || [])) {
    if (c.displayName && c.name) inv[c.name] = c.displayName;
  }
  // field_2 = documento, Title = razao social (segundo o padrao de import via XLSX)
  const all = [];
  let url = '/sites/' + siteId + '/lists/' + listId + '/items?expand=fields&$top=500';
  let pages = 0;
  while (url && pages < 30) {
    const resp = await client.api(url).get();
    all.push(...(resp.value || []));
    pages++;
    url = resp['@odata.nextLink']
      ? resp['@odata.nextLink'].replace('https://graph.microsoft.com/v1.0', '')
      : null;
  }
  const mapa = {};
  for (const it of all) {
    const f = it.fields || {};
    const out = {};
    for (const [k, v] of Object.entries(f)) {
      const display = inv[k] || k;
      out[display] = v;
    }
    const cnpj = String(out.documento || out.Documento || out.field_2 || out.CNPJ || '').replace(/\D/g, '');
    const razao = out.Title || out.razao || out.RazaoSocial || '';
    if (cnpj) mapa[cnpj] = razao;
  }
  return mapa;
}

function normalizeItem(item, invColMap) {
  const raw = item.fields || {};
  const out = { _id: item.id };
  for (const [k, v] of Object.entries(raw)) {
    const display = invColMap[k] || k;
    out[display] = v;
  }
  return out;
}

function fmtBRL(v) {
  if (!v) return 'R$ 0,00';
  try { return Number(v).toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' }); }
  catch (e) { return String(v); }
}

function fmtData(s) {
  if (!s) return '—';
  const d = String(s).substring(0, 10).split('-');
  if (d.length === 3) return d[2] + '/' + d[1] + '/' + d[0];
  return String(s);
}

function escapeHtml(s) {
  return String(s || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

// Calcula data BRT no formato YYYY-MM-DD
function hojeBRT() {
  const d = new Date(Date.now() - 3 * 60 * 60 * 1000);
  return d.toISOString().substring(0, 10);
}

// D+5 em dias uteis: simplificacao usa 5 dias corridos (manter compativel com sistema atual)
function diasAteVencer(vencimento) {
  if (!vencimento) return 999;
  const hoje = new Date(hojeBRT() + 'T00:00:00Z');
  const venc = new Date(String(vencimento).substring(0, 10) + 'T00:00:00Z');
  return Math.floor((venc - hoje) / (1000 * 60 * 60 * 24));
}

// =============================================================================
// Decide tipo do alerta — primeiro tenta query param, depois data/hora BRT
// =============================================================================
function decidirTipoAlerta(req) {
  const fromQuery = String((req.query && req.query.tipo) || '').toLowerCase();
  if (['manha','tarde','semanal'].includes(fromQuery)) return fromQuery;
  // Fallback: detecta pela hora BRT atual
  const brt = new Date(Date.now() - 3 * 60 * 60 * 1000);
  const dia = brt.getUTCDay();
  const hora = brt.getUTCHours();
  if (dia === 5 && hora >= 16) return 'semanal';
  if (hora >= 14) return 'tarde';
  return 'manha';
}

// =============================================================================
// =============================================================================
// Geracao de paragrafo de insight via Anthropic (com fallback OpenAI)
// =============================================================================
async function gerarInsight(contexto, tipo) {
  const promptSemanal = `Voce e a SOL, assistente IA do sistema de aprovacao de NF Pronep. Gere UM unico paragrafo curto (3-4 linhas, max 350 chars) de insight executivo sobre a semana do gestor. Foque em: padroes (concentracao em fornecedor/diretoria), comparacao com pendencias atuais, ou alerta acionavel. Tom profissional e direto, sem floreio. NUNCA repita os numeros que ja estao na tabela — adicione interpretacao.

Dados:
${JSON.stringify(contexto, null, 2)}

Responda com apenas o paragrafo, sem cabecalho.`;

  const promptDiario = `Voce e a SOL. Gere UM unico paragrafo curto (2-3 linhas, max 280 chars) de insight sobre as NFs pendentes do gestor.

PRIORIZE NESSA ORDEM (se aplicavel):
1) NFs VENCIDAS ou vencendo em <=5 dias — cite numero e fornecedor delas
2) Valores muito acima da media do fornecedor
3) Concentracao em 1 fornecedor

SE houver NFs em D+5, abra o paragrafo destacando-as por nome/numero. Tom direto, acionavel. Nao repita totais da tabela.

Dados:
${JSON.stringify(contexto, null, 2)}

Responda com apenas o paragrafo, sem cabecalho.`;

  const prompt = tipo === 'semanal' ? promptSemanal : promptDiario;

  // Tenta Anthropic primeiro
  if (process.env.ANTHROPIC_API_KEY) {
    try {
      let Anthropic;
      try { Anthropic = require('@anthropic-ai/sdk').default || require('@anthropic-ai/sdk'); } catch (e) {}
      if (Anthropic) {
        const client = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });
        const model = process.env.ANTHROPIC_MODEL_HAIKU || 'claude-haiku-4-5-20251001';
        const r = await client.messages.create({
          model: model,
          max_tokens: 200,
          temperature: 0.3,
          messages: [{ role: 'user', content: prompt }]
        });
        const txt = (r.content || []).filter(b => b.type === 'text').map(b => b.text).join('').trim();
        if (txt) return txt;
      }
    } catch (e) {
      console.error('[AlertaDiario] Anthropic falhou, caindo pra OpenAI:', e.message || e);
    }
  }

  // Fallback OpenAI
  if (!process.env.OPENAI_API_KEY) return null;
  let OpenAI;
  try { OpenAI = require('openai'); } catch (e) { return null; }
  try {
    const client = new OpenAI({ apiKey: process.env.OPENAI_API_KEY });
    const model = process.env.OPENAI_MODEL || 'gpt-4o-mini';
    const r = await client.chat.completions.create({
      model: model,
      messages: [{ role: 'user', content: prompt }],
      temperature: 0.3,
      max_tokens: 200
    });
    return (r.choices[0].message.content || '').trim();
  } catch (e) {
    return null;
  }
}

// =============================================================================
// Lista NFs pendentes e historico semanal
// =============================================================================
async function carregarNotas(client, siteId, listNotasId, invColMap) {
  const all = [];
  let url = '/sites/' + siteId + '/lists/' + listNotasId + '/items?expand=fields&$top=500';
  let pages = 0;
  while (url) {
    const resp = await client.api(url).get();
    all.push(...(resp.value || []));
    pages++;
    url = resp['@odata.nextLink']
      ? resp['@odata.nextLink'].replace('https://graph.microsoft.com/v1.0', '')
      : null;
    if (pages >= 30) break;
  }
  return all.map(it => normalizeItem(it, invColMap));
}

// Agrupa NFs por email do aprovador atual
function agruparPorAprovador(notas) {
  const grupos = {};
  for (const n of notas) {
    const email = String(n.AprovadorAtual || '').toLowerCase().trim();
    if (!email) continue;
    if (!grupos[email]) grupos[email] = [];
    grupos[email].push(n);
  }
  return grupos;
}

// =============================================================================
// Montagem do HTML do email
// =============================================================================
function buildEmailDiario(tipo, gestorEmail, notas, insight) {
  const isManha = tipo === 'manha';
  const isSemanal = tipo === 'semanal';
  const saudacao = isManha ? 'Bom dia' : (isSemanal ? 'Fechando a semana' : 'Boa tarde');
  const corHeader = isSemanal ? '#5E35B1' : '#1F4E79';
  const titulo = isSemanal
    ? 'Resumo da Semana — Aprovação de NF'
    : (isManha ? 'NFs pendentes pra você aprovar hoje' : 'Pendências do dia — atualização');
  const subtitulo = isSemanal
    ? 'Resumo executivo das últimas 5 dias úteis'
    : 'Sistema de Aprovação de NF · Pronep Life Care';

  // Ordena por vencimento ASC
  const ordenadas = notas.slice().sort((a, b) =>
    String(a.Vencimento || '').localeCompare(String(b.Vencimento || ''))
  );

  // Métricas
  const total = ordenadas.length;
  const totalValor = ordenadas.reduce((s, n) => s + (Number(n.Valor || 0)), 0);
  const d5 = ordenadas.filter(n => diasAteVencer(n.DataVencimento) <= 5);
  const vencidas = ordenadas.filter(n => diasAteVencer(n.DataVencimento) < 0);

  // Breakdown por unidade
  const porUnidade = {};
  for (const n of ordenadas) {
    const u = n.Unidade || '—';
    if (!porUnidade[u]) porUnidade[u] = { qtd: 0, total: 0 };
    porUnidade[u].qtd += 1;
    porUnidade[u].total += Number(n.Valor || 0);
  }

  // Top 10 NFs prioritarias pra mostrar na tabela
  const top = ordenadas.slice(0, 10);

  const linhasTop = top.map(n => {
    const d = diasAteVencer(n.DataVencimento);
    let badge = '';
    if (d < 0) badge = '<span style="background:#FFEBEE;color:#C62828;padding:2px 6px;border-radius:3px;font-size:10px;font-weight:600">VENCIDA</span>';
    else if (d <= 5) badge = '<span style="background:#FFF3E0;color:#E65100;padding:2px 6px;border-radius:3px;font-size:10px;font-weight:600">D+5</span>';
    return `<tr>
      <td style="padding:7px 10px;border-bottom:1px solid #ECEFF1;font-size:12px;color:#1F4E79"><b>${escapeHtml(n.NumeroNF || '—')}</b></td>
      <td style="padding:7px 10px;border-bottom:1px solid #ECEFF1;font-size:12px;color:#2C3E50">${escapeHtml(n._fornecedorNome || '—').substring(0, 30)}</td>
      <td style="padding:7px 10px;border-bottom:1px solid #ECEFF1;font-size:12px;color:#2C3E50;text-align:right"><b>${escapeHtml(fmtBRL(n.ValorTotal || n.Valor))}</b></td>
      <td style="padding:7px 10px;border-bottom:1px solid #ECEFF1;font-size:12px;color:#2C3E50">${escapeHtml(fmtData(n.DataVencimento))} ${badge}</td>
      <td style="padding:7px 10px;border-bottom:1px solid #ECEFF1;font-size:12px;color:#647883">${escapeHtml(n.Unidade || '—')}</td>
    </tr>`;
  }).join('');

  const restante = total > 10 ? `<tr><td colspan="5" style="padding:8px;text-align:center;font-size:11px;color:#647883">+ ${total - 10} NFs adicionais — abra o sistema pra ver todas</td></tr>` : '';

  const unidadesHtml = Object.entries(porUnidade).map(([u, dt]) =>
    `<tr><td style="padding:5px 12px;color:#647883;font-size:12px"><b>${escapeHtml(u)}</b></td>
         <td style="padding:5px 12px;color:#2C3E50;font-size:12px;text-align:right">${dt.qtd}</td>
         <td style="padding:5px 12px;color:#2C3E50;font-size:12px;text-align:right">${escapeHtml(fmtBRL(dt.total))}</td></tr>`
  ).join('');

  const insightBlock = insight ? `
    <div style="background:#E1F5FE;border-left:3px solid #27AAE1;padding:12px 16px;margin:18px 0;border-radius:4px">
      <div style="font-size:11px;color:#0277BD;font-weight:600;margin-bottom:4px;letter-spacing:0.5px;text-transform:uppercase">SAN — insight do dia</div>
      <div style="font-size:13px;color:#1F4E79;line-height:1.55">${escapeHtml(insight)}</div>
    </div>
  ` : '';

  const assunto = isSemanal
    ? `[Aprovação NF] Resumo semanal — ${total} pendente(s)`
    : `[Aprovação NF] ${isManha ? '☀' : '🌆'} ${total} NF${total > 1 ? 's' : ''} pra aprovar${d5.length > 0 ? ` · ${d5.length} em D+5` : ''}`;

  const alertaTopo = vencidas.length > 0
    ? `<div style="background:#FFEBEE;border-left:3px solid #C62828;padding:10px 14px;margin-bottom:14px;font-size:13px;color:#C62828"><b>⚠ Atenção:</b> ${vencidas.length} NF${vencidas.length > 1 ? 's vencidas' : ' vencida'} aguardando aprovação.</div>`
    : (d5.length > 0
      ? `<div style="background:#FFF8E1;border-left:3px solid #F9A825;padding:10px 14px;margin-bottom:14px;font-size:13px;color:#E65100"><b>Atenção:</b> ${d5.length} NF${d5.length > 1 ? 's vencendo' : ' vence'} em até 5 dias.</div>`
      : '');

  const corpo = `<!DOCTYPE html><html><body style="font-family:'Segoe UI',Roboto,Arial,sans-serif;background:#F4F8FB;margin:0;padding:0">
<table cellspacing="0" cellpadding="0" border="0" width="100%" style="background:#F4F8FB;padding:24px 0">
  <tr><td align="center">
    <table cellspacing="0" cellpadding="0" border="0" width="640" style="background:#fff;border-radius:8px;overflow:hidden;box-shadow:0 2px 8px rgba(31,78,121,0.1)">
      <tr><td style="background:linear-gradient(135deg,${corHeader} 0%,#27AAE1 100%);color:#fff;padding:20px 26px">
        <table cellspacing="0" cellpadding="0" border="0" width="100%"><tr>
          <td style="vertical-align:middle">
            <div style="font-size:11px;opacity:0.85;text-transform:uppercase;letter-spacing:1px">${escapeHtml(subtitulo)}</div>
            <div style="font-size:20px;font-weight:600;margin-top:4px">${escapeHtml(titulo)}</div>
          </td>
          <td style="vertical-align:middle;text-align:right;width:140px">
            <img src="https://purple-forest-09588fe10.7.azurestaticapps.net/pronep-logo.png" alt="Pronep" style="max-height:42px;max-width:130px;background:rgba(255,255,255,0.95);padding:6px 10px;border-radius:6px">
          </td>
        </tr></table>
      </td></tr>
      <tr><td style="padding:22px 26px;color:#2C3E50;font-size:14px;line-height:1.55">
        <p style="margin:0 0 12px"><b>${saudacao}!</b></p>
        ${alertaTopo}
        <table cellspacing="0" cellpadding="0" border="0" width="100%" style="margin-bottom:14px">
          <tr>
            <td style="background:#F4F8FB;padding:12px;border-radius:6px;text-align:center;width:33%">
              <div style="font-size:11px;color:#647883;text-transform:uppercase">Total Pendente</div>
              <div style="font-size:22px;font-weight:700;color:#1F4E79;margin-top:2px">${total}</div>
            </td>
            <td style="width:8px"></td>
            <td style="background:#F4F8FB;padding:12px;border-radius:6px;text-align:center;width:33%">
              <div style="font-size:11px;color:#647883;text-transform:uppercase">Valor Total</div>
              <div style="font-size:16px;font-weight:700;color:#1F4E79;margin-top:6px">${escapeHtml(fmtBRL(totalValor))}</div>
            </td>
            <td style="width:8px"></td>
            <td style="background:${d5.length > 0 ? '#FFF8E1' : '#F4F8FB'};padding:12px;border-radius:6px;text-align:center;width:33%">
              <div style="font-size:11px;color:${d5.length > 0 ? '#E65100' : '#647883'};text-transform:uppercase">Vencimento ≤ D+5</div>
              <div style="font-size:22px;font-weight:700;color:${d5.length > 0 ? '#E65100' : '#1F4E79'};margin-top:2px">${d5.length}</div>
            </td>
          </tr>
        </table>
        ${insightBlock}
        <h3 style="margin:18px 0 8px;color:#1F4E79;font-size:14px;font-weight:600">NFs prioritárias (por vencimento)</h3>
        <table cellspacing="0" cellpadding="0" border="0" width="100%" style="border-collapse:collapse">
          <thead>
            <tr style="background:#F4F8FB">
              <th style="padding:8px 10px;text-align:left;font-size:11px;color:#647883;text-transform:uppercase">NF</th>
              <th style="padding:8px 10px;text-align:left;font-size:11px;color:#647883;text-transform:uppercase">Fornecedor</th>
              <th style="padding:8px 10px;text-align:right;font-size:11px;color:#647883;text-transform:uppercase">Valor</th>
              <th style="padding:8px 10px;text-align:left;font-size:11px;color:#647883;text-transform:uppercase">Vencimento</th>
              <th style="padding:8px 10px;text-align:left;font-size:11px;color:#647883;text-transform:uppercase">Unid.</th>
            </tr>
          </thead>
          <tbody>${linhasTop}${restante}</tbody>
        </table>
        ${Object.keys(porUnidade).length > 1 ? `
        <h3 style="margin:22px 0 8px;color:#1F4E79;font-size:13px;font-weight:600">Por Unidade</h3>
        <table cellspacing="0" cellpadding="0" border="0" width="100%" style="background:#F4F8FB;border-radius:6px;border-collapse:collapse">
          <thead><tr><th style="padding:6px 12px;text-align:left;font-size:11px;color:#647883;text-transform:uppercase">Unid.</th>
          <th style="padding:6px 12px;text-align:right;font-size:11px;color:#647883;text-transform:uppercase">Qtd</th>
          <th style="padding:6px 12px;text-align:right;font-size:11px;color:#647883;text-transform:uppercase">Valor</th></tr></thead>
          <tbody>${unidadesHtml}</tbody>
        </table>` : ''}
        <p style="margin-top:24px;text-align:center">
          <a href="${SISTEMA_URL}?view=fila-aprovacao" style="background:#1F4E79;color:#fff;padding:12px 32px;border-radius:6px;text-decoration:none;font-weight:600;display:inline-block;font-size:14px">Abrir Fila de Aprovação</a>
        </p>
      </td></tr>
      <tr><td style="background:#F4F8FB;color:#647883;padding:14px 26px;font-size:11px;text-align:center;border-top:1px solid #DCE3E9">
        SAN — assistente IA · Pronep Life Care · automatizado, não responda
      </td></tr>
    </table>
  </td></tr></table></body></html>`;

  return { assunto, corpo };
}

async function enviarEmailViaGraph(graphClient, fromAddress, paraEmail, assunto, corpo) {
  const payload = {
    message: {
      subject: assunto,
      body: { contentType: 'HTML', content: corpo },
      toRecipients: [{ emailAddress: { address: paraEmail } }]
    },
    saveToSentItems: false
  };
  await graphClient.api('/users/' + fromAddress + '/sendMail').post(payload);
}

// =============================================================================
// HANDLER PRINCIPAL
// =============================================================================
module.exports = async function (context, req) {
  const stats = {
    iniciadoEm: new Date().toISOString(),
    tipoAlerta: null,
    aprovadoresProcessados: 0,
    emailsEnviados: 0,
    aprovadoresSemPendencias: 0,
    erros: []
  };

  // Auth por shared secret (necessario porque endpoint eh anonymous)
  const expectedSecret = process.env.ALERTA_DIARIO_SECRET;
  if (!expectedSecret) {
    context.res = { status: 500, body: { error: 'ALERTA_DIARIO_SECRET nao configurado' } };
    return;
  }
  const headerSecret = (req.headers && (req.headers['x-alerta-secret'] || req.headers['X-Alerta-Secret'])) || '';
  if (headerSecret !== expectedSecret) {
    context.res = { status: 401, body: { error: 'Unauthorized' } };
    return;
  }

  try {
    if (String(process.env.ALERTA_DIARIO_DISABLED || '').toLowerCase() === 'true') {
      context.log && context.log('AlertaDiario desabilitado por env var');
      context.res = { status: 200, body: { skipped: true, reason: 'ALERTA_DIARIO_DISABLED=true' } };
      return;
    }

    const tipo = decidirTipoAlerta(req);
    stats.tipoAlerta = tipo;
    context.log && context.log('AlertaDiario iniciando — tipo:', tipo);

    const client = await getGraphClient();
    const { siteId, listNotasId, invColMap } = await resolveSiteAndList(client);
    const todas = await carregarNotas(client, siteId, listNotasId, invColMap);
    // Carrega mapa CNPJ -> Razao Social pra preencher nome do fornecedor em cada NF
    let mapFornec = {};
    try { mapFornec = await carregarMapFornecedores(client, siteId); }
    catch (e) { context.log && context.log('WARN: nao conseguiu carregar fornecedores:', e.message); }
    // Resolve nome do fornecedor por CNPJ em cada nota
    for (const n of todas) {
      const cnpj = String(n.CNPJFornecedor || '').replace(/\D/g, '');
      n._fornecedorNome = (cnpj && mapFornec[cnpj]) || n.CNPJFornecedor || '—';
    }

    // Filtra pendentes
    const pendentes = todas.filter(n => ['Lancada', 'EmAprovacao', 'Pendente'].includes(String(n.Status || '')));

    // Agrupa por aprovador
    const grupos = agruparPorAprovador(pendentes);

    const fromAddress = process.env.EMAIL_FROM_ADDRESS || DEFAULT_FROM;
    const testeEmail = process.env.ALERTA_DIARIO_TESTE_EMAIL || '';

    for (const [aprovadorEmail, notas] of Object.entries(grupos)) {
      stats.aprovadoresProcessados++;
      if (!notas.length) {
        stats.aprovadoresSemPendencias++;
        continue;
      }
      try {
        // Contexto pro insight
        const contexto = {
          tipo: tipo,
          total: notas.length,
          totalValor: notas.reduce((s, n) => s + Number(n.Valor || 0), 0),
          fornecedores: [...new Set(notas.map(n => n._fornecedorNome).filter(Boolean))].slice(0, 10),
          vencendoEmD5: notas.filter(n => diasAteVencer(n.DataVencimento) <= 5).length,
          vencidas: notas.filter(n => diasAteVencer(n.DataVencimento) < 0).length,
          unidades: [...new Set(notas.map(n => n.Unidade).filter(Boolean))]
        };
        const insight = await gerarInsight(contexto, tipo);

        const { assunto, corpo } = buildEmailDiario(tipo, aprovadorEmail, notas, insight);
        const destinatario = testeEmail || aprovadorEmail;
        await enviarEmailViaGraph(client, fromAddress, destinatario, assunto, corpo);
        stats.emailsEnviados++;
        context.log && context.log('Email enviado para', destinatario, '· NFs:', notas.length);
      } catch (e) {
        stats.erros.push({ aprovador: aprovadorEmail, erro: e.message });
        context.log && context.log.error && context.log.error('Erro ao enviar pra', aprovadorEmail, e.message);
      }
    }

    stats.finalizadoEm = new Date().toISOString();
    stats.duracaoMs = Date.now() - new Date(stats.iniciadoEm).getTime();
    context.log && context.log('AlertaDiario concluido:', JSON.stringify(stats));
    context.res = { status: 200, headers: {'Content-Type':'application/json'}, body: stats };
  } catch (err) {
    stats.erros.push({ geral: err.message });
    context.log && context.log.error && context.log.error('AlertaDiario falhou:', err);
    context.res = { status: 500, headers: {'Content-Type':'application/json'}, body: stats };
  }
};
