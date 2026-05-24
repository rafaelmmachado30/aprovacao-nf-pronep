/**
 * Modulo compartilhado de envio de email + Teams
 * Usado por PostNota, AprovarNota, RejeitarNota e a Function EnviarNotificacao
 *
 * Chamada in-process (sem HTTP loop) pra evitar problemas de roteamento interno.
 */

require('isomorphic-fetch');
const { ClientSecretCredential } = require('@azure/identity');
const { Client } = require('@microsoft/microsoft-graph-client');
const { TokenCredentialAuthenticationProvider } =
  require('@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials');

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

function buildEmail(evento, dados) {
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
        <p style="margin-top:24px;text-align:center">
          <a href="https://purple-forest-09588fe10.7.azurestaticapps.net/" style="background:${corHeader};color:#fff;padding:12px 28px;border-radius:6px;text-decoration:none;font-weight:600;display:inline-block">Abrir Sistema</a>
        </p>
        ${urlPDF ? `<p style="margin-top:16px;font-size:13px;color:#647883;text-align:center">PDF arquivado: <a href="${escapeHtml(urlPDF)}" style="color:${corHeader}">abrir no SharePoint</a></p>` : ''}
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
  const { assunto, corpo } = buildEmail(evento, dados || {});
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
async function enviarTeams(evento, dados) {
  const webhook = process.env.TEAMS_WEBHOOK_URL;
  if (!webhook) return { ok: false, skipped: true, reason: 'TEAMS_WEBHOOK_URL nao setado' };
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
        actions: [
          { type: 'Action.OpenUrl', title: 'Abrir Sistema',
            url: 'https://purple-forest-09588fe10.7.azurestaticapps.net/' }
        ]
      }
    }]
  };
  const r = await fetch(webhook, {
    method: 'POST', headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(card)
  });
  return { ok: r.ok, status: r.status };
}

// Funcao alto-nivel chamada por outras Functions in-process
async function notificar(evento, destinatarios, dados, cc) {
  const result = { email: null, teams: null };
  if (Array.isArray(destinatarios) && destinatarios.length > 0) {
    try { result.email = await enviarEmail(evento, destinatarios, dados, cc); }
    catch (e) { result.email = { ok: false, error: e.message, statusCode: e.statusCode, body: e.body }; }
  }
  try { result.teams = await enviarTeams(evento, dados); }
  catch (e) { result.teams = { ok: false, error: e.message }; }
  return result;
}

module.exports = { notificar, enviarEmail, enviarTeams, buildEmail };
