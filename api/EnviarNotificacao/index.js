/**
 * Sistema de Aprovacao de NF - EnviarNotificacao
 *
 * POST /api/EnviarNotificacao
 * Body JSON: { evento, notaId, destinatarios: ["email1","email2"], cc, dados }
 *
 *   evento: 'lancada' | 'aprovada' | 'rejeitada'
 *   notaId: ID da lista SP (string)
 *   destinatarios: array de e-mails do destinatario principal
 *   cc: array opcional de e-mails em copia
 *   dados: { numero, fornecedor, valor, vencimento, unidade, diretoria, aprovador, motivo, urlPDF }
 *
 * Envia:
 *  - E-mail via Graph (Mail.Send) em nome do EMAIL_FROM_ADDRESS
 *  - Teams Adaptive Card via TEAMS_WEBHOOK_URL (se configurado)
 *
 * App Settings opcionais:
 *   EMAIL_FROM_ADDRESS  (default: datanalytics@pronep.com.br)
 *   TEAMS_WEBHOOK_URL   (se nao setar, pula envio Teams)
 *
 * Permissoes Graph (Application):
 *   Mail.Send  (com admin consent)
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

// Constroi assunto + corpo HTML do email
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

  let assunto = '';
  let corpo = '';
  let corHeader = '#1F4E79';
  let titulo = '';
  let acao = '';

  if (evento === 'lancada') {
    assunto = `[Aprovação NF] Nova NF ${numero ? 'NF '+numero : ''} aguardando sua aprovação`;
    corHeader = '#1F4E79';
    titulo = 'Nova Nota Fiscal para aprovação';
    acao = `<p><b>Ação necessária:</b> você é o aprovador responsável pela diretoria <b>${escapeHtml(diretoria)}</b> da unidade <b>${escapeHtml(unidade)}</b>. Acesse o sistema para aprovar ou rejeitar.</p>`;
  } else if (evento === 'aprovada') {
    assunto = `[Aprovação NF] Sua NF ${numero ? 'NF '+numero : ''} foi APROVADA`;
    corHeader = '#2E7D32';
    titulo = '✓ Nota Fiscal Aprovada';
    acao = `<p>Sua NF foi aprovada por <b>${escapeHtml(aprovador)}</b>. O PDF com watermark foi arquivado em <b>Notas Aprovadas</b>.</p>`;
  } else if (evento === 'rejeitada') {
    assunto = `[Aprovação NF] Sua NF ${numero ? 'NF '+numero : ''} foi REJEITADA`;
    corHeader = '#C62828';
    titulo = '✕ Nota Fiscal Rejeitada';
    acao = `<p>Sua NF foi rejeitada por <b>${escapeHtml(aprovador)}</b>.</p>
            <p><b>Motivo:</b> ${escapeHtml(motivo)}</p>
            <p>Faça as correções necessárias e reenvie a NF pelo sistema.</p>`;
  } else {
    assunto = `[Aprovação NF] Evento ${evento}`;
    titulo = `Evento ${evento}`;
  }

  corpo = `<!DOCTYPE html><html><body style="font-family:'Segoe UI',Roboto,Arial,sans-serif;background:#F4F8FB;margin:0;padding:0">
<table cellspacing="0" cellpadding="0" border="0" width="100%" style="background:#F4F8FB;padding:24px 0">
  <tr><td align="center">
    <table cellspacing="0" cellpadding="0" border="0" width="600" style="background:#fff;border-radius:8px;overflow:hidden;box-shadow:0 2px 8px rgba(31,78,121,0.1)">
      <tr><td style="background:${corHeader};color:#fff;padding:18px 24px;font-size:18px;font-weight:600">
        Pronep — Aprovação de NF
      </td></tr>
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
              <tr><td style="color:#647883"><b>Unidade / Diretoria</b></td><td style="color:#2C3E50">${escapeHtml(unidade)} · ${escapeHtml(diretoria)}</td></tr>
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
        Sistema de Aprovação de NF · Pronep Life Care · automatizado, não responda este e-mail
      </td></tr>
    </table>
  </td></tr>
</table>
</body></html>`;

  return { assunto, corpo };
}

// Constroi Adaptive Card pro Teams
function buildTeamsCard(evento, dados) {
  const numero = dados.numero || '';
  const fornecedor = dados.fornecedor || '';
  const valor = fmtBRL(dados.valor);
  const vencimento = fmtData(dados.vencimento);
  const unidade = dados.unidade || '';
  const diretoria = dados.diretoria || '';
  const motivo = dados.motivo || '';

  let titulo = '';
  let cor = 'default';
  if (evento === 'lancada')   { titulo = '📬 Nova NF para aprovação'; cor = 'attention'; }
  if (evento === 'aprovada')  { titulo = '✓ NF aprovada'; cor = 'good'; }
  if (evento === 'rejeitada') { titulo = '✕ NF rejeitada'; cor = 'warning'; }

  const facts = [
    { title: 'NF',          value: numero },
    { title: 'Fornecedor',  value: fornecedor },
    { title: 'Valor',       value: valor },
    { title: 'Vencimento',  value: vencimento },
    { title: 'Unidade',     value: unidade },
    { title: 'Diretoria',   value: diretoria }
  ];
  if (evento === 'rejeitada' && motivo) facts.push({ title: 'Motivo', value: motivo });

  return {
    type: 'message',
    attachments: [{
      contentType: 'application/vnd.microsoft.card.adaptive',
      content: {
        $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
        type: 'AdaptiveCard',
        version: '1.4',
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
}

module.exports = async function (context, req) {
  const diag = { step: 'start' };
  try {
    diag.step = 'parse_body';
    const body = req.body || {};
    const { evento, notaId, destinatarios, cc, dados } = body;
    if (!evento)         return errResp(context, 400, 'evento obrigatorio');
    if (!destinatarios || !destinatarios.length) return errResp(context, 400, 'destinatarios obrigatorio');
    diag.evento = evento;
    diag.destinatarios = destinatarios;

    const fromAddress = process.env.EMAIL_FROM_ADDRESS || DEFAULT_FROM;
    diag.from = fromAddress;

    diag.step = 'build_email';
    const { assunto, corpo } = buildEmail(evento, dados || {});

    diag.step = 'send_email';
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
    let emailResult = 'sent';
    try {
      await client.api(`/users/${fromAddress}/sendMail`).post(mailPayload);
      diag.emailEnviado = true;
    } catch (mailErr) {
      diag.emailErro = {
        message: mailErr.message,
        statusCode: mailErr.statusCode,
        body: mailErr.body
      };
      emailResult = 'falhou';
    }

    diag.step = 'send_teams';
    const teamsWebhook = process.env.TEAMS_WEBHOOK_URL;
    let teamsResult = 'skipped';
    if (teamsWebhook) {
      const card = buildTeamsCard(evento, dados || {});
      try {
        const r = await fetch(teamsWebhook, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(card)
        });
        if (r.ok) {
          teamsResult = 'sent';
          diag.teamsEnviado = true;
        } else {
          teamsResult = 'falhou (' + r.status + ')';
          diag.teamsErro = `HTTP ${r.status}`;
        }
      } catch (te) {
        teamsResult = 'falhou (network)';
        diag.teamsErro = te.message;
      }
    }

    context.res = {
      status: 200,
      headers: { 'Content-Type': 'application/json' },
      body: {
        ok: true,
        evento,
        email: emailResult,
        teams: teamsResult,
        diag
      }
    };
  } catch (err) {
    context.log && context.log.error && context.log.error('EnviarNotificacao error:', err);
    context.res = {
      status: 500,
      headers: { 'Content-Type': 'application/json' },
      body: {
        error: (err && err.message) || String(err),
        statusCode: err && err.statusCode,
        graphBody: err && err.body,
        diag
      }
    };
  }
};

function errResp(context, status, msg) {
  context.res = {
    status,
    headers: { 'Content-Type': 'application/json' },
    body: { error: msg }
  };
}
