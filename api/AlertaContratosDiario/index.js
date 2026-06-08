/**
 * AlertaContratosDiario — HTTP Endpoint (disparado por cron externo)
 *
 * Varre PRONEP-NF-Contratos, identifica contratos vencendo em 90, 60 ou 30
 * dias e envia email pro gestor da diretoria. Faz dedup via campo
 * `UltimoAlertaJanela` no contrato pra nao enviar 2x a mesma janela.
 *
 * URL: GET /api/AlertaContratosDiario?secret=<ALERTA_CONTRATOS_SECRET>
 *
 * Comportamento:
 *   - Pra cada contrato: identifica menor janela (90/60/30) ainda nao alertada
 *   - Tolerancia: aceita +/-2 dias da janela ideal (caso o cron falhe um dia)
 *   - Pula Status=Cancelado e DataFim nula
 *   - Resolve gestor via PRONEP-NF-Diretorias (field_3 = email)
 *   - Envia email institucional + push (se subscrito)
 *   - Grava `UltimoAlertaJanela=<90|60|30>` na coluna Observacoes pra dedup
 *
 * Pode ser disparado por:
 *   - GitHub Actions cron (workflow .github/workflows/alerta-contratos-diario.yml)
 *   - Scheduled-tasks do MCP
 *   - Admin manual (pra teste)
 *
 * App Settings:
 *   - AAD_TENANT_ID, AAD_CLIENT_ID, AAD_CLIENT_SECRET
 *   - SHAREPOINT_SITE_HOSTNAME, SHAREPOINT_SITE_PATH
 *   - ALERTA_CONTRATOS_SECRET (obrigatorio se quiser proteger - default sem secret)
 *   - ALERTA_CONTRATOS_TESTE_EMAIL (opcional - manda tudo pra esse email; debug)
 *   - ALERTA_CONTRATOS_DISABLED ('true' pra pausar)
 */

require('isomorphic-fetch');
const { ClientSecretCredential } = require('@azure/identity');
const { Client } = require('@microsoft/microsoft-graph-client');
const { TokenCredentialAuthenticationProvider } = require('@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials');

const LIST_CONTRATOS = 'PRONEP-NF-Contratos';
const LIST_DIRETORIAS = 'PRONEP-NF-Diretorias';
const SISTEMA_URL = 'https://purple-forest-09588fe10.7.azurestaticapps.net/';
const DEFAULT_FROM = 'datanalytics@pronep.com.br';
// OID do grupo AAD que tem acesso total a contratos (membros recebem TODOS os alertas).
// Mantem sincronizado com shared/userRoles.js GROUP_TO_ROLE pra role 'gestor_juridica'.
const GRUPO_JURIDICO_OID = '5aa9fc6b-900d-40eb-861d-8bbf72499da1';

// Janelas de alerta em dias (do MENOR pro MAIOR - prioriza mais urgente)
// Logica: o alerta dispara quando o contrato entra DENTRO da janela (dias <= janela).
// Cada janela so dispara UMA vez (controlado por dedup em Observacoes).
// Ex: contrato com 85 dias → alerta_90 hoje. Em 25 dias (60 dias restantes) → alerta_60.
// Em mais 30 (30 restantes) → alerta_30.
const JANELAS = [30, 60, 90];

function getGraphClient() {
  const credential = new ClientSecretCredential(
    process.env.AAD_TENANT_ID, process.env.AAD_CLIENT_ID, process.env.AAD_CLIENT_SECRET
  );
  const authProvider = new TokenCredentialAuthenticationProvider(credential, { scopes: ['https://graph.microsoft.com/.default'] });
  return Client.initWithMiddleware({ authProvider });
}

function diasParaVencer(dataFim) {
  if (!dataFim) return null;
  const hoje = new Date(new Date().getTime() - 3*60*60*1000);
  const hojeStr = hoje.toISOString().substring(0,10);
  const hj = new Date(hojeStr + 'T00:00:00Z');
  const fim = new Date(String(dataFim).substring(0,10) + 'T00:00:00Z');
  return Math.round((fim.getTime() - hj.getTime()) / (24*60*60*1000));
}

function fmtBRL(v) {
  if (v == null || v === '') return '—';
  try { return Number(v).toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' }); }
  catch (e) { return String(v); }
}

function fmtData(s) {
  if (!s) return '—';
  const d = String(s).substring(0,10).split('-');
  return d.length === 3 ? d[2] + '/' + d[1] + '/' + d[0] : String(s);
}

function escapeHtml(s) {
  return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

// Identifica a JANELA aplicavel pra um contrato baseado em diasParaVencer + dedup em Observacoes
// Logica: pega a MAIS ESPECIFICA (menor janela) que se aplica E que ainda nao foi enviada.
// Ex: contrato com 25 dias e ja foi enviado alerta_30: nao envia nada (proxima seria so apos vencer).
// Ex: contrato com 25 dias sem nenhum alerta: envia alerta_30 (a mais urgente aplicavel).
function escolherJanelaAplicavel(diasFalta, observacoes) {
  if (diasFalta == null || diasFalta < 0) return null;  // vencido nao alerta (ou alerta separado)
  if (diasFalta > 90) return null;                       // longe demais da janela
  const obs = String(observacoes || '');
  const ja = {
    30: /_alerta_30=/.test(obs),
    60: /_alerta_60=/.test(obs),
    90: /_alerta_90=/.test(obs)
  };
  // Pega a janela mais urgente APLICAVEL (dias <= janela) e ainda nao enviada
  for (const j of JANELAS) {  // [30, 60, 90] - mais urgente primeiro
    if (diasFalta <= j && !ja[j]) return j;
  }
  return null;
}

function buildEmailContrato(contrato, janela) {
  const fornecedor = escapeHtml(contrato.fornecedor || '(sem fornecedor)');
  const titulo = escapeHtml(contrato.titulo || contrato.nomeArquivo || '(sem titulo)');
  const diretoria = escapeHtml(contrato.diretoria || '—');
  const unidade = escapeHtml(contrato.unidade || '—');
  const dataFim = fmtData(contrato.dataFim);
  const dataInicio = fmtData(contrato.dataInicio);
  const valor = fmtBRL(contrato.valor);
  const dias = contrato.diasParaVencer != null ? contrato.diasParaVencer : '?';

  // Cores e urgencia por janela
  let cor = '#3B82F6', icone = '⏰', urgenciaTxt = 'Atencao antecipada';
  if (janela === 60) { cor = '#F59E0B'; icone = '⚠'; urgenciaTxt = 'Inicie a renegociacao'; }
  if (janela === 30) { cor = '#DC2626'; icone = '🚨'; urgenciaTxt = 'URGENTE - decisao necessaria'; }

  const assunto = '[Contrato] Vencendo em ' + janela + ' dias — ' + (contrato.fornecedor || '');
  const linkContrato = SISTEMA_URL + 'api/AbrirContrato?id=' + encodeURIComponent(contrato.id);
  const linkSistema = SISTEMA_URL + '?view=contratos';

  const html =
    '<!DOCTYPE html><html lang="pt-BR"><head><meta charset="UTF-8"><style>' +
    'body{font-family:"Segoe UI",Arial,sans-serif;color:#2C3E50;line-height:1.55;margin:0;padding:0;background:#F4F8FB}' +
    '.wrap{max-width:640px;margin:24px auto;background:#fff;border-radius:8px;overflow:hidden;box-shadow:0 4px 16px rgba(31,78,121,0.1)}' +
    '.hdr{background:' + cor + ';color:#fff;padding:18px 24px;display:flex;align-items:center;gap:12px}' +
    '.hdr .icone{font-size:28px}' +
    '.hdr .titulo{font-size:18px;font-weight:700;margin:0}' +
    '.hdr .sub{font-size:12px;opacity:.85;margin-top:2px}' +
    '.body{padding:24px}' +
    '.urgencia{background:#FFF4E5;border-left:4px solid ' + cor + ';padding:12px 16px;border-radius:4px;margin-bottom:18px;font-size:13px;color:#5D4037}' +
    '.urgencia b{color:' + cor + '}' +
    '.tabela{width:100%;border-collapse:collapse;font-size:13px;margin-bottom:18px}' +
    '.tabela th{text-align:left;padding:8px 12px;background:#F4F8FB;color:#1F4E79;font-size:11px;text-transform:uppercase;letter-spacing:0.4px;width:35%}' +
    '.tabela td{padding:8px 12px;border-bottom:1px solid #ECEFF1}' +
    '.cta{text-align:center;margin:24px 0 8px}' +
    '.cta a{background:#1F4E79;color:#fff;text-decoration:none;padding:12px 28px;border-radius:6px;font-weight:600;font-size:14px;display:inline-block;margin:0 4px}' +
    '.cta a.sec{background:#fff;color:#1F4E79;border:2px solid #1F4E79}' +
    '.foot{padding:14px 24px;background:#F4F8FB;text-align:center;font-size:11px;color:#647883;border-top:1px solid #E0E7EE}' +
    '</style></head><body>' +
    '<div class="wrap">' +
    '<div class="hdr"><div class="icone">' + icone + '</div><div><div class="titulo">Contrato vencendo em ' + janela + ' dias</div><div class="sub">' + urgenciaTxt + '</div></div></div>' +
    '<div class="body">' +
    '<div class="urgencia"><b>' + fornecedor + '</b> tem contrato vencendo em <b>' + dias + ' dia' + (dias === 1 ? '' : 's') + '</b> (' + dataFim + ').</div>' +
    '<table class="tabela">' +
    '<tr><th>Fornecedor</th><td>' + fornecedor + '</td></tr>' +
    '<tr><th>Contrato</th><td style="font-size:12px;word-break:break-word">' + titulo + '</td></tr>' +
    '<tr><th>Diretoria</th><td>' + diretoria + '</td></tr>' +
    '<tr><th>Unidade</th><td>' + unidade + '</td></tr>' +
    '<tr><th>Vigência</th><td>' + dataInicio + ' → <b style="color:' + cor + '">' + dataFim + '</b></td></tr>' +
    '<tr><th>Valor do contrato</th><td>' + valor + '</td></tr>' +
    '<tr><th>Status atual</th><td>' + escapeHtml(contrato.status || '—') + '</td></tr>' +
    '</table>' +
    '<div class="cta">' +
    '<a href="' + linkContrato + '" target="_blank">📄 Abrir PDF do Contrato</a>' +
    '<a href="' + linkSistema + '" target="_blank" class="sec">Ver no Sistema</a>' +
    '</div>' +
    '</div>' +
    '<div class="foot">Você recebeu este alerta porque é gestor da diretoria <b>' + diretoria + '</b>.<br>Pronep Life Care · Sistema de Aprovação de Notas Fiscais e Gestão de Contratos</div>' +
    '</div></body></html>';

  return { assunto, html };
}

async function enviarEmail(client, destinatario, assunto, html) {
  const from = process.env.EMAIL_FROM_ADDRESS || DEFAULT_FROM;
  const teste = process.env.ALERTA_CONTRATOS_TESTE_EMAIL;
  const to = teste || destinatario;
  await client.api('/users/' + from + '/sendMail').post({
    message: {
      subject: assunto,
      body: { contentType: 'HTML', content: html },
      toRecipients: [{ emailAddress: { address: to } }],
      from: { emailAddress: { address: from } }
    },
    saveToSentItems: 'true'
  });
  return { ok: true, sentTo: to, viaTeste: !!teste };
}

// Envia push notification (PWA/browser) pro gestor - best-effort, falha silencia
async function enviarPushAlerta(client, siteId, gestorEmail, contrato, janela) {
  let pushNotif;
  try { pushNotif = require('../shared/pushNotif'); } catch (e) { return { ok: false, error: 'pushNotif nao carregado' }; }
  if (!pushNotif.configurarWebPush()) return { ok: false, skipped: true, reason: 'VAPID nao configurado' };
  const valorFmt = contrato.valor ? Number(contrato.valor).toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' }) : '';
  const icone = janela === 30 ? '🚨' : janela === 60 ? '⚠' : '⏰';
  const payload = {
    title: icone + ' Contrato vencendo em ' + janela + ' dias',
    body: (contrato.fornecedor || '') + (valorFmt ? ' · ' + valorFmt : '') + ' · vence em ' + (contrato.diasParaVencer || '?') + 'd',
    tag: 'contrato-' + contrato.id,
    url: '/?view=contratos',
    evento: 'contrato_vencendo',
    contratoId: contrato.id,
    timestamp: Date.now()
  };
  try {
    return await pushNotif.enviarPushPraEmail(client, siteId, gestorEmail, payload);
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

async function gravarDedup(client, siteId, listContratosId, colMapContratos, contratoId, observacoesAtuais, janela) {
  const colObs = (colMapContratos && colMapContratos['Observacoes']) || 'Observacoes';
  const hoje = new Date(new Date().getTime() - 3*60*60*1000).toISOString().substring(0,10);
  const novaLinha = '_alerta_' + janela + '=' + hoje;
  const obs = observacoesAtuais || '';
  // Evita duplicacao se ja tinha
  if (obs.indexOf('_alerta_' + janela + '=') >= 0) return;
  const nova = obs ? (obs + '\n' + novaLinha) : novaLinha;
  await client.api('/sites/' + siteId + '/lists/' + listContratosId + '/items/' + contratoId + '/fields')
    .patch({ [colObs]: nova });
}

async function carregarContratos(client, siteId, listContratosId, invColMap) {
  const todos = [];
  let nextUrl = '/sites/' + siteId + '/lists/' + listContratosId + '/items?expand=fields&$top=999';
  let pages = 0;
  while (nextUrl && pages < 50) {
    pages++;
    const r = await client.api(nextUrl).get();
    for (const it of (r.value || [])) {
      const raw = it.fields || {};
      const f = {};
      for (const [k, v] of Object.entries(raw)) {
        const display = invColMap[k] || k;
        f[display] = v;
      }
      todos.push({
        id: it.id,
        titulo: f.Title || '',
        fornecedor: f.Fornecedor || '',
        diretoria: f.Diretoria || '',
        unidade: f.Unidade || 'CORPORATIVO',
        dataInicio: f.DataInicio ? String(f.DataInicio).substring(0,10) : null,
        dataFim: f.DataFim ? String(f.DataFim).substring(0,10) : null,
        valor: f.ValorContrato != null ? Number(f.ValorContrato) : null,
        status: f.Status || '',
        observacoes: f.Observacoes || ''
      });
    }
    const next = r['@odata.nextLink'];
    if (next) {
      const idx = next.indexOf('/v1.0/');
      nextUrl = idx >= 0 ? next.substring(idx + 5) : null;
    } else nextUrl = null;
  }
  return todos;
}

// Resolve emails dos membros do grupo Juridico (acesso total a contratos).
// Esses emails sao adicionados como destinatarios extras em TODOS os alertas.
async function obterEmailsJuridico(client) {
  try {
    const resp = await client.api('/groups/' + GRUPO_JURIDICO_OID + '/members')
      .select('mail,userPrincipalName,displayName').top(100).get();
    const emails = [];
    for (const m of (resp.value || [])) {
      const email = (m.mail || m.userPrincipalName || '').toLowerCase().trim();
      if (email && email.indexOf('@') >= 0 && emails.indexOf(email) < 0) emails.push(email);
    }
    return emails;
  } catch (e) {
    // Falha silenciosa - Juridico nao recebe extras nessa execucao mas o resto roda
    return [];
  }
}

async function carregarGestoresPorDiretoria(client, siteId, listDirId) {
  const map = {};
  if (!listDirId) return map;
  const r = await client.api('/sites/' + siteId + '/lists/' + listDirId + '/items?expand=fields&$top=200').get();
  for (const it of (r.value || [])) {
    const f = it.fields || {};
    const email = String(f.field_3 || '').toLowerCase().trim();
    const titulo = String(f.Title || '');
    const partes = titulo.split('|');
    const dir = (partes[1] || '').trim();
    if (!dir || !email) continue;
    if (!map[dir]) map[dir] = [];
    if (map[dir].indexOf(email) < 0) map[dir].push(email);
  }
  return map;
}

module.exports = async function (context, req) {
  const inicio = Date.now();
  try {
    if (process.env.ALERTA_CONTRATOS_DISABLED === 'true') {
      context.res = { status: 200, body: { ok: true, skipped: true, motivo: 'ALERTA_CONTRATOS_DISABLED=true' } };
      return;
    }
    // Validacao de secret (se configurado)
    const secretEnv = process.env.ALERTA_CONTRATOS_SECRET;
    if (secretEnv) {
      const secretReq = (req.query && req.query.secret) || (req.headers && req.headers['x-alerta-secret']);
      if (secretReq !== secretEnv) {
        context.res = { status: 401, body: { error: 'secret invalido' } };
        return;
      }
    }
    const dryRun = String((req.query && req.query.dryRun) || '') === 'true';

    const client = getGraphClient();
    // Resolve site e listas
    const host = process.env.SHAREPOINT_SITE_HOSTNAME;
    const path = process.env.SHAREPOINT_SITE_PATH;
    const siteResp = await client.api('/sites/' + host + ':' + path).get();
    const siteId = siteResp.id;
    const lists = await client.api('/sites/' + siteId + '/lists').get();
    let listContratosId = null, listDirId = null;
    for (const l of (lists.value || [])) {
      if (l.displayName === LIST_CONTRATOS) listContratosId = l.id;
      if (l.displayName === LIST_DIRETORIAS) listDirId = l.id;
    }
    if (!listContratosId) {
      context.res = { status: 500, body: { error: 'Lista ' + LIST_CONTRATOS + ' nao encontrada' } };
      return;
    }
    // Cols
    const colsResp = await client.api('/sites/' + siteId + '/lists/' + listContratosId + '/columns').get();
    const colMap = {}, invColMap = {};
    for (const c of (colsResp.value || [])) {
      if (c.displayName && c.name) { colMap[c.displayName] = c.name; invColMap[c.name] = c.displayName; }
    }
    // Contratos + gestores + emails do Juridico (recebem TODOS os alertas)
    const contratos = await carregarContratos(client, siteId, listContratosId, invColMap);
    const gestores = await carregarGestoresPorDiretoria(client, siteId, listDirId);
    const emailsJuridico = await obterEmailsJuridico(client);

    const stats = {
      totalContratos: contratos.length, candidatos: 0, enviados: 0, semGestor: 0,
      jaAlertados: 0, foraDeJanela: 0, vencidos: 0, semDataFim: 0, cancelados: 0, erros: 0,
      breakdown: { dentro_30: 0, dentro_60: 0, dentro_90: 0, alem_90: 0 }
    };
    const enviados = [];
    const erros = [];
    const semGestor = [];

    for (const c of contratos) {
      if (c.status === 'Cancelado') { stats.cancelados++; continue; }
      const dias = diasParaVencer(c.dataFim);
      if (dias == null) { stats.semDataFim++; continue; }
      if (dias < 0) { stats.vencidos++; continue; }
      // Breakdown de distribuicao
      if (dias <= 30) stats.breakdown.dentro_30++;
      else if (dias <= 60) stats.breakdown.dentro_60++;
      else if (dias <= 90) stats.breakdown.dentro_90++;
      else stats.breakdown.alem_90++;

      const janela = escolherJanelaAplicavel(dias, c.observacoes);
      if (!janela) {
        // Pode ser fora de janela (>90d) OU ja foi alertado (todas as janelas do nivel ja enviadas)
        if (dias > 90) stats.foraDeJanela++;
        else stats.jaAlertados++;
        continue;
      }
      stats.candidatos++;
      // Resolve gestor da diretoria + adiciona Juridico (acesso total) em TODOS os alertas.
      // Dedup pra evitar enviar duplicado se Juridico tambem for gestor da diretoria.
      const gestoresDir = gestores[c.diretoria] || [];
      const destinatarios = [];
      for (const e of gestoresDir) if (destinatarios.indexOf(e) < 0) destinatarios.push(e);
      for (const e of emailsJuridico) if (destinatarios.indexOf(e) < 0) destinatarios.push(e);
      if (!destinatarios.length) {
        stats.semGestor++;
        semGestor.push({ contratoId: c.id, fornecedor: c.fornecedor, diretoria: c.diretoria });
        continue;
      }
      const contratoComDias = Object.assign({}, c, { diasParaVencer: dias });
      const { assunto, html } = buildEmailContrato(contratoComDias, janela);
      if (dryRun) {
        enviados.push({ contratoId: c.id, fornecedor: c.fornecedor, diretoria: c.diretoria, gestores: destinatarios, janela: janela, dias: dias, dryRun: true });
        continue;
      }
      // Envia pra cada gestor da diretoria (email + push best-effort por email)
      let sucesso = false;
      for (const email of destinatarios) {
        try {
          await enviarEmail(client, email, assunto, html);
          sucesso = true;
        } catch (e) {
          stats.erros++;
          erros.push({ contratoId: c.id, gestor: email, erro: 'email: ' + e.message });
        }
        // Push notification (in-app) - falha silenciosa
        try {
          await enviarPushAlerta(client, siteId, email, contratoComDias, janela);
        } catch (e) { /* push e best-effort */ }
      }
      if (sucesso) {
        stats.enviados++;
        try {
          await gravarDedup(client, siteId, listContratosId, colMap, c.id, c.observacoes, janela);
        } catch (e) {
          erros.push({ contratoId: c.id, gestor: 'dedup', erro: 'falha gravar dedup: ' + e.message });
        }
        enviados.push({ contratoId: c.id, fornecedor: c.fornecedor, diretoria: c.diretoria, gestores: destinatarios, janela: janela, dias: dias });
      }
    }

    context.res = {
      status: 200,
      headers: { 'Content-Type': 'application/json' },
      body: {
        ok: true,
        tempoMs: Date.now() - inicio,
        dryRun,
        stats,
        enviados: enviados.slice(0, 100),
        semGestor: semGestor.slice(0, 50),
        erros: erros.slice(0, 50)
      }
    };
  } catch (err) {
    context.log && context.log.error && context.log.error('AlertaContratosDiario:', err);
    context.res = { status: 500, body: { error: err.message, stack: (err.stack || '').split('\n').slice(0, 8) } };
  }
};
