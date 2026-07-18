/**
 * Modulo compartilhado de envio de email + Teams
 * Usado por PostNota, AprovarNota, RejeitarNota e a Function EnviarNotificacao
 *
 * Chamada in-process (sem HTTP loop) pra evitar problemas de roteamento interno.
 */

require('isomorphic-fetch');
const jwt = require('jsonwebtoken');
const { ClientSecretCredential } = require('@azure/identity');
const { Client } = require('@microsoft/microsoft-graph-client');
const { TokenCredentialAuthenticationProvider } =
  require('@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials');
const { enviarTeamsAtividade } = require('./teamsActivity');

const DEFAULT_FROM = 'datanalytics@pronep.com.br';

async function getGraphClient() {
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

// Gera links assinados pra aprovar/rejeitar via email
function gerarLinks(itemId, aprovadorEmail) {
  const secret = process.env.LINK_APROVACAO_SECRET;
  if (!secret) return null;
  const base = 'https://purple-forest-09588fe10.7.azurestaticapps.net';
  // C7: validade reduzida de 7d para 48h (menor janela se o link vazar) + jti unico
  // por token (identificador rastreavel) + algoritmo HS256 explicito.
  const crypto = require('crypto');
  const opts = { expiresIn: '48h', algorithm: 'HS256' };
  const tokenAprovar = jwt.sign({ itemId, aprovador: aprovadorEmail, action: 'aprovar', jti: crypto.randomUUID() }, secret, opts);
  const tokenRejeitar = jwt.sign({ itemId, aprovador: aprovadorEmail, action: 'rejeitar', jti: crypto.randomUUID() }, secret, opts);
  return {
    aprovar: `${base}/api/AprovacaoViaLink?token=${encodeURIComponent(tokenAprovar)}`,
    rejeitar: `${base}/api/AprovacaoViaLink?token=${encodeURIComponent(tokenRejeitar)}`
  };
}

function escapeHtml(s) {
  return String(s || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function fmtBRL(v) {
  if (v === null || v === undefined || v === '') return '—';
  try { return Number(v).toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' }); }
  catch (e) { return String(v); }
}

function fmtData(s) {
  if (!s) return '—';
  const d = String(s).substring(0,10);
  const partes = d.split('-');
  if (partes.length === 3) return `${partes[2]}/${partes[1]}/${partes[0]}`;
  return d;
}

function buildEmail(evento, dados, links) {
  const numero = dados.numero || '';
  const fornecedor = dados.fornecedor || '';
  const valor = fmtBRL(dados.valor);
  const vencimento = fmtData(dados.vencimento);
  const unidade = dados.unidade || '';
  const diretoria = dados.diretoria || '';
  const aprovador = dados.aprovador || '';
  const motivo = dados.motivo || '';
  const urlPDF = dados.urlPDF || '';
  const submitter = dados.submitter || '';

  let assunto = '', corHeader = '#1F4E79', titulo = '', acao = '';

  if (evento === 'lancada') {
    assunto = `[Aprovação NF] Nova NF ${numero ? 'NF '+numero : ''} aguardando sua aprovação`;
    corHeader = '#1F4E79'; titulo = 'Nova Nota Fiscal para aprovação';
    acao = `<p><b>Ação necessária:</b> você é o aprovador responsável pela diretoria <b>${escapeHtml(diretoria)}</b> da unidade <b>${escapeHtml(unidade)}</b>. Acesse o sistema para aprovar ou rejeitar.</p>`;
  } else if (evento === 'aprovada') {
    assunto = `[Aprovação NF] Sua NF ${numero ? 'NF '+numero : ''} foi APROVADA`;
    corHeader = '#2E7D32'; titulo = '✓ Nota Fiscal Aprovada';
    acao = `<p>Sua NF foi aprovada por <b>${escapeHtml(aprovador)}</b>. O PDF com watermark foi arquivado em <b>Notas Aprovadas</b>.</p>`;
  } else if (evento === 'rejeitada') {
    assunto = `[Aprovação NF] Sua NF ${numero ? 'NF '+numero : ''} foi REJEITADA`;
    corHeader = '#C62828'; titulo = '✕ Nota Fiscal Rejeitada';
    acao = `<p>Sua NF foi rejeitada por <b>${escapeHtml(aprovador)}</b>.</p>
            <p><b>Motivo:</b> ${escapeHtml(motivo)}</p>
            <p>Faça as correções e reenvie a NF pelo sistema.</p>`;
  } else {
    assunto = `[Aprovação NF] ${evento}`; titulo = `Evento ${evento}`;
  }

  const corpo = `<!DOCTYPE html><html><body style="font-family:'Segoe UI',Roboto,Arial,sans-serif;background:#F4F8FB;margin:0;padding:0">
<table cellspacing="0" cellpadding="0" border="0" width="100%" style="background:#F4F8FB;padding:24px 0">
  <tr><td align="center">
    <table cellspacing="0" cellpadding="0" border="0" width="600" style="background:#fff;border-radius:8px;overflow:hidden;box-shadow:0 2px 8px rgba(31,78,121,0.1)">
      <tr><td style="background:${corHeader};color:#fff;padding:18px 24px;font-size:18px;font-weight:600">Pronep — Aprovação de NF</td></tr>
      <tr><td style="padding:24px;color:#2C3E50;font-size:14px;line-height:1.55">
        <h2 style="margin:0 0 14px 0;color:${corHeader};font-size:20px">${titulo}</h2>
        ${acao}
        <table cellspacing="0" cellpadding="0" border="0" width="100%" style="margin-top:18px;background:#F4F8FB;border-radius:6px">
          <tr><td style="padding:14px 16px">
            <table cellspacing="0" cellpadding="6" border="0" width="100%" style="font-size:13px">
              <tr><td style="color:#647883;width:140px"><b>NF</b></td><td style="color:#2C3E50">${escapeHtml(numero)}</td></tr>
              <tr><td style="color:#647883"><b>Fornecedor</b></td><td style="color:#2C3E50">${escapeHtml(fornecedor)}</td></tr>
              <tr><td style="color:#647883"><b>Valor</b></td><td style="color:#2C3E50"><b>${escapeHtml(valor)}</b></td></tr>
              <tr><td style="color:#647883"><b>Vencimento</b></td><td style="color:#2C3E50">${escapeHtml(vencimento)}</td></tr>
              <tr><td style="color:#647883"><b>Unidade · Diretoria</b></td><td style="color:#2C3E50">${escapeHtml(unidade)} · ${escapeHtml(diretoria)}</td></tr>
              ${submitter ? `<tr><td style="color:#647883"><b>Lançado por</b></td><td style="color:#2C3E50">${escapeHtml(submitter)}</td></tr>` : ''}
            </table>
          </td></tr>
        </table>
        ${links && evento === 'lancada' ? `
        ${urlPDF ? `
        <p style="margin-top:20px;text-align:center">
          <a href="${escapeHtml(urlPDF)}" target="_blank" style="background:#1F4E79;color:#fff;padding:13px 32px;border-radius:6px;text-decoration:none;font-weight:600;display:inline-block;font-size:15px">📄 Ver PDF da NF</a>
        </p>
        <p style="margin-top:6px;text-align:center;font-size:12px;color:#647883">Recomendamos visualizar o PDF antes de aprovar ou rejeitar.</p>
        ` : ''}
        <p style="margin-top:18px;text-align:center">
          <a href="${links.aprovar}" style="background:#2E7D32;color:#fff;padding:14px 32px;border-radius:6px;text-decoration:none;font-weight:600;display:inline-block;margin:0 6px;font-size:15px">✓ Aprovar</a>
          <a href="${links.rejeitar}" style="background:#C62828;color:#fff;padding:14px 32px;border-radius:6px;text-decoration:none;font-weight:600;display:inline-block;margin:0 6px;font-size:15px">✕ Rejeitar</a>
        </p>
        <p style="margin-top:10px;text-align:center;font-size:12px;color:#647883">Click direto aqui aprova ou rejeita. Token expira em 7 dias.</p>
        <p style="margin-top:20px;text-align:center">
          <a href="https://purple-forest-09588fe10.7.azurestaticapps.net/" style="background:#fff;color:${corHeader};padding:10px 24px;border-radius:6px;text-decoration:none;font-weight:600;display:inline-block;border:1px solid ${corHeader}">Ou abrir o Sistema</a>
        </p>` : `
        <p style="margin-top:24px;text-align:center">
          <a href="https://purple-forest-09588fe10.7.azurestaticapps.net/" style="background:${corHeader};color:#fff;padding:12px 28px;border-radius:6px;text-decoration:none;font-weight:600;display:inline-block">Abrir Sistema</a>
        </p>
        ${urlPDF ? `<p style="margin-top:14px;text-align:center;font-size:13px"><a href="${escapeHtml(urlPDF)}" target="_blank" style="color:${corHeader}">📄 Ver PDF arquivado</a></p>` : ''}`}
        
      </td></tr>
      <tr><td style="background:#F4F8FB;color:#647883;padding:14px 24px;font-size:11px;text-align:center;border-top:1px solid #DCE3E9">
        Sistema de Aprovação de NF · Pronep Life Care · automatizado, não responda
      </td></tr>
    </table>
  </td></tr></table></body></html>`;

  return { assunto, corpo };
}

// Envia email via Graph (Mail.Send Application)
async function enviarEmail(evento, destinatarios, dados, cc) {
  const fromAddress = process.env.EMAIL_FROM_ADDRESS || DEFAULT_FROM;
  // Gera links assinados se for evento 'lancada' e tivermos itemId
  let links = null;
  if (evento === 'lancada' && dados && dados.itemId && destinatarios[0]) {
    links = gerarLinks(dados.itemId, destinatarios[0]);
  }
  const { assunto, corpo } = buildEmail(evento, dados || {}, links);
  const client = await getGraphClient();
  const mailPayload = {
    message: {
      subject: assunto,
      body: { contentType: 'HTML', content: corpo },
      toRecipients: destinatarios.map(e => ({ emailAddress: { address: e } })),
      ccRecipients: (cc || []).map(e => ({ emailAddress: { address: e } }))
    },
    saveToSentItems: true
  };
  await client.api(`/users/${fromAddress}/sendMail`).post(mailPayload);
  return { ok: true, from: fromAddress, to: destinatarios };
}

// Envia Adaptive Card via Teams Incoming Webhook (se configurado)
// Resolve URL do webhook do aprovador especifico (com fallback pro canal central)
function resolverWebhookUrl(aprovadorEmail) {
  // App Setting TEAMS_WEBHOOKS deve ser um JSON tipo {"email":"url", ...}
  const webhooksJson = process.env.TEAMS_WEBHOOKS;
  if (webhooksJson) {
    try {
      const webhooks = JSON.parse(webhooksJson);
      const url = webhooks[(aprovadorEmail || '').toLowerCase()];
      if (url) return { url, scope: '1on1', for: aprovadorEmail };
    } catch (e) { /* JSON invalido, continua */ }
  }
  // Fallback: canal central (TEAMS_WEBHOOK_URL)
  const canalUrl = process.env.TEAMS_WEBHOOK_URL;
  if (canalUrl) return { url: canalUrl, scope: 'canal', for: 'canal' };
  return null;
}

async function enviarTeams(evento, dados, destinatariosEmail) {
  // Escolhe webhook por aprovador (1:1) ou cai pro canal central
  const aprovadorAlvo = (destinatariosEmail && destinatariosEmail[0]) || dados.aprovador || '';
  const webhookInfo = resolverWebhookUrl(aprovadorAlvo);
  if (!webhookInfo) return { ok: false, skipped: true, reason: 'Nenhum webhook configurado (nem TEAMS_WEBHOOKS por usuario nem TEAMS_WEBHOOK_URL canal)' };
  const webhook = webhookInfo.url;

  const numero = dados.numero || '';
  const fornecedor = dados.fornecedor || '';
  const valor = fmtBRL(dados.valor);
  const vencimento = fmtData(dados.vencimento);
  const unidade = dados.unidade || '';
  const diretoria = dados.diretoria || '';
  const motivo = dados.motivo || '';

  let titulo = '', cor = 'default';
  if (evento === 'lancada')   { titulo = '📬 Nova NF para aprovação'; cor = 'attention'; }
  if (evento === 'aprovada')  { titulo = '✓ NF aprovada'; cor = 'good'; }
  if (evento === 'rejeitada') { titulo = '✕ NF rejeitada'; cor = 'warning'; }

  const facts = [
    { title: 'NF', value: String(numero) },
    { title: 'Fornecedor', value: String(fornecedor) },
    { title: 'Valor', value: valor },
    { title: 'Vencimento', value: vencimento },
    { title: 'Unidade', value: String(unidade) },
    { title: 'Diretoria', value: String(diretoria) }
  ];
  if (evento === 'rejeitada' && motivo) facts.push({ title: 'Motivo', value: String(motivo) });

  // Gera links assinados se for evento 'lancada' (botoes interativos no card)
  let actions = [];
  if (dados.urlPDF) {
    actions.push({ type: 'Action.OpenUrl', title: '📄 Ver PDF', url: dados.urlPDF });
  }
  actions.push({ type: 'Action.OpenUrl', title: 'Abrir Sistema',
    url: 'https://purple-forest-09588fe10.7.azurestaticapps.net/' });
  if (evento === 'lancada' && dados.itemId && destinatariosEmail && destinatariosEmail[0]) {
    const links = gerarLinks(dados.itemId, destinatariosEmail[0]);
    if (links) {
      actions = [];
      if (dados.urlPDF) {
        actions.push({ type: 'Action.OpenUrl', title: '📄 Ver PDF', url: dados.urlPDF });
      }
      actions.push(
        { type: 'Action.OpenUrl', title: '✓ Aprovar', url: links.aprovar, style: 'positive' },
        { type: 'Action.OpenUrl', title: '✕ Rejeitar', url: links.rejeitar, style: 'destructive' },
        { type: 'Action.OpenUrl', title: 'Abrir Sistema',
          url: 'https://purple-forest-09588fe10.7.azurestaticapps.net/' }
      );
    }
  }

  const card = {
    type: 'message',
    attachments: [{
      contentType: 'application/vnd.microsoft.card.adaptive',
      content: {
        $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
        type: 'AdaptiveCard', version: '1.4',
        body: [
          { type: 'TextBlock', size: 'Large', weight: 'Bolder', color: cor, text: titulo },
          { type: 'FactSet', facts: facts }
        ],
        actions: actions
      }
    }]
  };
  const r = await fetch(webhook, {
    method: 'POST', headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(card)
  });
  return { ok: r.ok, status: r.status, scope: webhookInfo.scope, sentTo: webhookInfo.for };
}

// Funcao alto-nivel chamada por outras Functions in-process
async function notificar(evento, destinatarios, dados, cc) {
  const result = { email: null, teamsAtividade: null, teamsWebhook: null, push: null };
  if (Array.isArray(destinatarios) && destinatarios.length > 0) {
    try { result.email = await enviarEmail(evento, destinatarios, dados, cc); }
    catch (e) { result.email = { ok: false, error: e.message, statusCode: e.statusCode, body: e.body }; }
  }

  // Caminho oficial: sendActivityNotification via Graph (1:1 no Teams do aprovador)
  const aprovadorAlvo = (destinatarios && destinatarios[0]) || (dados && dados.aprovador) || '';
  if (aprovadorAlvo) {
    try { result.teamsAtividade = await enviarTeamsAtividade(evento, dados || {}, aprovadorAlvo); }
    catch (e) { result.teamsAtividade = { ok: false, error: e.message, body: e.body, statusCode: e.statusCode }; }
  } else {
    result.teamsAtividade = { ok: false, skipped: true, reason: 'Sem aprovador para resolver userId' };
  }

  // Fallback: se o caminho Graph nao tiver funcionado (Teams App nao registrada ainda
  // ou permissao pendente) E houver webhook de canal central configurado, posta la.
  // Webhook 1:1 (TEAMS_WEBHOOKS por email) continua disponivel mas DEPRECIADO no codigo.
  if (!result.teamsAtividade || !result.teamsAtividade.ok) {
    try { result.teamsWebhook = await enviarTeams(evento, dados || {}, destinatarios); }
    catch (e) { result.teamsWebhook = { ok: false, error: e.message }; }
  } else {
    result.teamsWebhook = { ok: true, skipped: true, reason: 'Atividade Graph enviou com sucesso' };
  }

  // Push notification: dispara pra todos os destinatarios que tiverem subscription registrada.
  // Best-effort — falhas nao bloqueiam o resto. Se VAPID nao configurado, vira no-op.
  try {
    result.push = await enviarPushParaDestinatarios(evento, destinatarios, dados);
  } catch (e) {
    result.push = { ok: false, error: e.message };
  }

  return result;
}

// Envia push pra cada destinatario (best-effort). Reutiliza graph client local.
async function enviarPushParaDestinatarios(evento, destinatarios, dados) {
  if (!Array.isArray(destinatarios) || destinatarios.length === 0) {
    return { ok: true, skipped: true, reason: 'sem destinatarios' };
  }
  let pushNotif;
  try { pushNotif = require('./pushNotif'); }
  catch (e) { return { ok: false, error: 'pushNotif module nao carregado: ' + e.message }; }

  if (!pushNotif.configurarWebPush()) {
    return { ok: false, skipped: true, reason: 'VAPID nao configurado' };
  }

  // Constroi payload do push
  const numero = (dados && dados.numero) || '';
  const valor = (dados && dados.valor) || 0;
  const forn = (dados && dados.fornecedor) || '';
  const nfId = (dados && dados.nfId) || (dados && dados.id) || '';
  const valorFmt = typeof valor === 'number'
    ? valor.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' })
    : valor;

  let title = 'Aprovacao NF Pronep';
  let body = '';
  if (evento === 'lancada') {
    title = 'Nova NF para aprovar';
    body = (forn ? forn + ' · ' : '') + 'NF ' + numero + ' · ' + valorFmt;
  } else if (evento === 'aprovada') {
    title = 'NF aprovada';
    body = 'NF ' + numero + ' foi aprovada · ' + valorFmt;
  } else if (evento === 'rejeitada') {
    title = 'NF rejeitada';
    body = 'NF ' + numero + ' foi rejeitada' + (dados && dados.motivo ? ' · ' + dados.motivo : '');
  } else {
    body = 'NF ' + numero;
  }

  const payload = {
    title: title,
    body: body,
    tag: 'nf-' + nfId,                    // permite "substituir" notificacao da mesma NF
    url: '/?view=fila-aprovacao' + (nfId ? '&nf=' + encodeURIComponent(nfId) : ''),
    evento: evento,
    nfId: nfId,
    timestamp: Date.now()
  };

  // Conecta no Graph (mesmo padrao das outras notifs)
  const { ClientSecretCredential } = require('@azure/identity');
  const { Client } = require('@microsoft/microsoft-graph-client');
  const { TokenCredentialAuthenticationProvider } =
    require('@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials');

  const credential = new ClientSecretCredential(
    process.env.AAD_TENANT_ID, process.env.AAD_CLIENT_ID, process.env.AAD_CLIENT_SECRET
  );
  const authProvider = new TokenCredentialAuthenticationProvider(credential, {
    scopes: ['https://graph.microsoft.com/.default']
  });
  const client = Client.initWithMiddleware({ authProvider });
  const host = process.env.SHAREPOINT_SITE_HOSTNAME;
  const path = process.env.SHAREPOINT_SITE_PATH;
  const siteResp = await client.api(`/sites/${host}:${path}`).get();
  const siteId = siteResp.id;

  const resultados = {};
  for (const email of destinatarios) {
    if (!email) continue;
    try {
      resultados[email] = await pushNotif.enviarPushPraEmail(client, siteId, email, payload);
    } catch (e) {
      resultados[email] = { error: e.message };
    }
  }
  return { ok: true, porEmail: resultados };
}

module.exports = { notificar, enviarEmail, enviarTeams, buildEmail, gerarLinks };
