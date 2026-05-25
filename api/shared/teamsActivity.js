/**
 * Envio de notificacao Teams 1:1 via Graph API (sendActivityNotification)
 * Substitui o webhook depreciated do Power Automate.
 *
 * Garante que a Teams App esteja instalada no Teams pessoal do aprovador
 * antes de enviar a notificacao (requerido pelo sendActivityNotification).
 *
 * Permissoes Application requeridas:
 *   - TeamsActivity.Send
 *   - TeamsAppInstallation.ReadWriteForUser.All
 *   - AppCatalog.Read.All
 *   - User.Read.All
 *
 * Doc: https://learn.microsoft.com/graph/teams-send-activityfeednotifications
 */

require('isomorphic-fetch');
const { ClientSecretCredential } = require('@azure/identity');
const { Client } = require('@microsoft/microsoft-graph-client');
const { TokenCredentialAuthenticationProvider } =
  require('@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials');

const BASE_URL = 'https://purple-forest-09588fe10.7.azurestaticapps.net';
const TEAMS_APP_ID = process.env.TEAMS_APP_ID || '5c52fce3-bf34-4b8c-a624-06defd5f85f6';
const TEAMS_ENTITY_ID = 'aprovacao-nf-home';

// Microsoft exige topic.webUrl no formato https://teams.microsoft.com/l/...
function buildTeamsDeepLink(itemId) {
  const subEntityId = itemId ? 'nf-' + itemId : 'home';
  const context = JSON.stringify({ subEntityId: subEntityId });
  const webUrl = BASE_URL + '/';
  return 'https://teams.microsoft.com/l/entity/' + TEAMS_APP_ID + '/' + TEAMS_ENTITY_ID
    + '?webUrl=' + encodeURIComponent(webUrl)
    + '&context=' + encodeURIComponent(context);
}

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

async function resolverUserId(client, email) {
  if (!email) throw new Error('Email do aprovador vazio');
  const user = await client.api('/users/' + encodeURIComponent(email)).select('id,mail,userPrincipalName').get();
  if (!user || !user.id) throw new Error('Usuario nao encontrado: ' + email);
  return user.id;
}

// Cache em memoria pra nao consultar o catalogo a cada notif
let _catalogAppIdCache = null;
async function getCatalogAppId(client) {
  if (_catalogAppIdCache) return _catalogAppIdCache;
  const r = await client.api('/appCatalogs/teamsApps')
    .filter("externalId eq '" + TEAMS_APP_ID + "'")
    .select('id,displayName,externalId,distributionMethod')
    .get();
  if (!r || !r.value || r.value.length === 0) {
    throw new Error('Teams App nao encontrada no catalogo do tenant. externalId=' + TEAMS_APP_ID);
  }
  _catalogAppIdCache = r.value[0].id;
  return _catalogAppIdCache;
}

async function garantirAppInstalada(client, userId, catalogAppId) {
  // Lista apps instaladas pro user e procura a nossa
  try {
    const installed = await client.api('/users/' + userId + '/teamwork/installedApps')
      .expand('teamsApp')
      .get();
    if (installed && installed.value) {
      const ja = installed.value.find(function (x) {
        return x.teamsApp && x.teamsApp.id === catalogAppId;
      });
      if (ja) return 'already-installed';
    }
  } catch (e) {
    // Se a leitura falhar, tenta instalar mesmo assim
  }
  // Instala
  await client.api('/users/' + userId + '/teamwork/installedApps').post({
    'teamsApp@odata.bind': 'https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/' + catalogAppId
  });
  return 'just-installed';
}

function fmtBRL(v) {
  if (v === null || v === undefined || v === '') return '';
  try { return Number(v).toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' }); }
  catch (e) { return String(v); }
}

function mapearEvento(evento) {
  if (evento === 'lancada')   return 'approvalRequired';
  if (evento === 'aprovada')  return 'nfAprovada';
  if (evento === 'rejeitada') return 'nfRejeitada';
  return null;
}

function buildPayload(evento, dados) {
  const activityType = mapearEvento(evento);
  if (!activityType) return null;

  const numero = String(dados.numero || '-');
  const fornecedor = String(dados.fornecedor || '');
  const valor = fmtBRL(dados.valor);
  const aprovador = String(dados.aprovador || 'Sistema NF');
  const motivo = String(dados.motivo || '');

  let topicValue, previewText;
  if (evento === 'lancada') {
    topicValue = 'Nova NF ' + numero + ' aguardando sua aprovacao';
    previewText = fornecedor + (valor ? ' - ' + valor : '');
  } else if (evento === 'aprovada') {
    topicValue = 'Sua NF ' + numero + ' foi APROVADA';
    previewText = 'Aprovada por ' + aprovador;
  } else {
    topicValue = 'Sua NF ' + numero + ' foi REJEITADA';
    previewText = motivo ? 'Motivo: ' + motivo : 'Rejeitada por ' + aprovador;
  }

  return {
    topic: {
      source: 'text',
      value: topicValue,
      webUrl: buildTeamsDeepLink(dados.itemId)
    },
    activityType: activityType,
    previewText: { content: previewText }
  };
}

async function enviarTeamsAtividade(evento, dados, aprovadorEmail) {
  const payload = buildPayload(evento, dados);
  if (!payload) return { ok: false, skipped: true, reason: 'Evento sem mapeamento: ' + evento };
  if (!aprovadorEmail) return { ok: false, skipped: true, reason: 'Email do aprovador vazio' };

  let installStatus = 'not-attempted';
  let userId = null;
  try {
    const client = await getGraphClient();
    userId = await resolverUserId(client, aprovadorEmail);

    // Garante que a Teams App esta instalada pro aprovador antes de notificar
    try {
      const catalogAppId = await getCatalogAppId(client);
      installStatus = await garantirAppInstalada(client, userId, catalogAppId);
    } catch (instErr) {
      installStatus = 'install-error: ' + (instErr.message || String(instErr));
      // Continua mesmo assim: pode ser que ja esteja instalada e a leitura falhou
    }

    await client.api('/users/' + userId + '/teamwork/sendActivityNotification').post(payload);
    return {
      ok: true,
      sentTo: aprovadorEmail,
      userId: userId,
      activityType: payload.activityType,
      installStatus: installStatus
    };
  } catch (err) {
    return {
      ok: false,
      status: err.statusCode,
      error: err.message,
      body: err.body,
      sentTo: aprovadorEmail,
      userId: userId,
      activityType: payload.activityType,
      installStatus: installStatus
    };
  }
}

module.exports = {
  enviarTeamsAtividade,
  buildPayload,
  buildTeamsDeepLink,
  mapearEvento,
  getCatalogAppId,
  garantirAppInstalada
};
