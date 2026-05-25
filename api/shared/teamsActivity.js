/**
 * Envio de notificacao Teams 1:1 via Graph API (sendActivityNotification)
 *
 * Substitui o webhook depreciated do Power Automate. Caminho oficial Microsoft:
 *   POST /users/{userId}/teamwork/sendActivityNotification
 *
 * Requer:
 *   - Teams App registrada no tenant com webApplicationInfo.id = nosso App Reg
 *   - Permissao Application "TeamsActivity.Send" + admin consent
 *
 * Documentacao: https://learn.microsoft.com/graph/teams-send-activityfeednotifications
 */

require('isomorphic-fetch');
const { ClientSecretCredential } = require('@azure/identity');
const { Client } = require('@microsoft/microsoft-graph-client');
const { TokenCredentialAuthenticationProvider } =
  require('@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials');

const BASE_URL = 'https://purple-forest-09588fe10.7.azurestaticapps.net';

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

/**
 * Resolve userId (UUID) a partir do email do aprovador
 */
async function resolverUserId(client, email) {
  if (!email) throw new Error('Email do aprovador vazio');
  // /users/{email} aceita UPN diretamente
  const user = await client.api(`/users/${encodeURIComponent(email)}`).select('id,mail,userPrincipalName').get();
  if (!user || !user.id) throw new Error(`Usuario nao encontrado: ${email}`);
  return user.id;
}

function fmtBRL(v) {
  if (v === null || v === undefined || v === '') return '';
  try { return Number(v).toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' }); }
  catch (e) { return String(v); }
}

/**
 * Mapeia evento do sistema para o activityType declarado no manifest da Teams App
 */
function mapearEvento(evento) {
  if (evento === 'lancada')   return 'approvalRequired';
  if (evento === 'aprovada')  return 'nfAprovada';
  if (evento === 'rejeitada') return 'nfRejeitada';
  return null;
}

/**
 * Monta o payload do sendActivityNotification
 *
 * @param {string} evento - 'lancada' | 'aprovada' | 'rejeitada'
 * @param {object} dados  - { numero, fornecedor, valor, aprovador, motivo, itemId }
 */
function buildPayload(evento, dados) {
  const activityType = mapearEvento(evento);
  if (!activityType) return null;

  const numero = String(dados.numero || '—');
  const fornecedor = String(dados.fornecedor || '');
  const valor = fmtBRL(dados.valor);
  const aprovador = String(dados.aprovador || 'Sistema NF');
  const motivo = String(dados.motivo || '');

  // Topic: o titulo que aparece na activity feed (sino do Teams)
  // webUrl: pra onde clicar leva o usuario
  let topicValue, previewText;
  if (evento === 'lancada') {
    topicValue = `Nova NF ${numero} aguardando sua aprovacao`;
    previewText = `${fornecedor}${valor ? ' — ' + valor : ''}`;
  } else if (evento === 'aprovada') {
    topicValue = `Sua NF ${numero} foi APROVADA`;
    previewText = `Aprovada por ${aprovador}`;
  } else {
    topicValue = `Sua NF ${numero} foi REJEITADA`;
    previewText = motivo ? `Motivo: ${motivo}` : `Rejeitada por ${aprovador}`;
  }

  // templateParameters: substituidos no templateText declarado no manifest
  const templateParameters = [
    { name: 'actor',    value: aprovador },
    { name: 'nfNumber', value: numero }
  ];

  return {
    topic: {
      source: 'text',
      value: topicValue,
      webUrl: BASE_URL + '/'
    },
    activityType,
    previewText: { content: previewText },
    templateParameters
  };
}

/**
 * Envia notificacao 1:1 via Graph sendActivityNotification.
 * Retorna { ok: bool, status?, error?, sentTo? }
 */
async function enviarTeamsAtividade(evento, dados, aprovadorEmail) {
  const payload = buildPayload(evento, dados);
  if (!payload) return { ok: false, skipped: true, reason: `Evento '${evento}' sem mapeamento` };
  if (!aprovadorEmail) return { ok: false, skipped: true, reason: 'Email do aprovador vazio' };

  const client = await getGraphClient();
  const userId = await resolverUserId(client, aprovadorEmail);

  try {
    await client.api(`/users/${userId}/teamwork/sendActivityNotification`).post(payload);
    return { ok: true, sentTo: aprovadorEmail, userId, activityType: payload.activityType };
  } catch (err) {
    return {
      ok: false,
      status: err.statusCode,
      error: err.message,
      body: err.body,
      sentTo: aprovadorEmail,
      activityType: payload.activityType
    };
  }
}

module.exports = { enviarTeamsAtividade, buildPayload, mapearEvento };
