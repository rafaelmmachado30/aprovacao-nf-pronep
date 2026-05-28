/**
 * Helper de Web Push Notifications.
 *
 * Funcoes:
 *  - configurarWebPush()        — inicializa lib web-push com VAPID keys
 *  - listarSubscriptionsPorEmail(client, siteId, email)
 *  - salvarSubscription(client, siteId, email, subscription, userAgent)
 *  - removerSubscription(client, siteId, endpoint)
 *  - enviarPushPraEmail(client, siteId, email, payload) — usa todas as subscriptions do user
 *
 * IMPORTANTE: requer env vars VAPID_PUBLIC_KEY, VAPID_PRIVATE_KEY, VAPID_SUBJECT.
 * Se nao configuradas, as funcoes viram no-op (loga warn e segue sem disparar).
 *
 * SP list: PRONEP-NF-PushSubscriptions
 *   Columns: Title (email), Endpoint (text/url), P256DH (text), Auth (text), UserAgent (text)
 */

const webpush = require('web-push');

const LIST_SUBSCRIPTIONS = 'PRONEP-NF-PushSubscriptions';

let _configurado = false;
let _configurarTentou = false;

function configurarWebPush() {
  if (_configurado) return true;
  if (_configurarTentou) return false;
  _configurarTentou = true;

  const pub = process.env.VAPID_PUBLIC_KEY;
  const priv = process.env.VAPID_PRIVATE_KEY;
  const subject = process.env.VAPID_SUBJECT || 'mailto:datanalytics@pronep.com.br';

  if (!pub || !priv) {
    console.warn('[pushNotif] VAPID_PUBLIC_KEY ou VAPID_PRIVATE_KEY nao configurados. Push desabilitado.');
    return false;
  }
  try {
    webpush.setVapidDetails(subject, pub, priv);
    _configurado = true;
    return true;
  } catch (e) {
    console.error('[pushNotif] Falha ao configurar web-push:', e.message || e);
    return false;
  }
}

// Resolve internal name da coluna pelo displayName (cache simples)
let _colMapCache = null;
async function getColMap(client, siteId, listName) {
  if (_colMapCache && _colMapCache._key === siteId + '|' + listName) return _colMapCache.map;
  const lists = await client.api(`/sites/${siteId}/lists`).filter(`displayName eq '${listName}'`).get();
  if (!lists.value || !lists.value.length) throw new Error(`Lista '${listName}' nao encontrada`);
  const listId = lists.value[0].id;
  const cols = await client.api(`/sites/${siteId}/lists/${listId}/columns`).get();
  const map = { _listId: listId };
  for (const col of (cols.value || [])) {
    if (col.displayName && col.name) map[col.displayName] = col.name;
  }
  _colMapCache = { _key: siteId + '|' + listName, map };
  return map;
}

// Le todas as subscriptions de um email
async function listarSubscriptionsPorEmail(client, siteId, email) {
  if (!email) return [];
  try {
    const map = await getColMap(client, siteId, LIST_SUBSCRIPTIONS);
    const listId = map._listId;
    // Filtra pela coluna Title (= email)
    const resp = await client
      .api(`/sites/${siteId}/lists/${listId}/items?expand=fields&$top=100`)
      .filter(`fields/Title eq '${email.toLowerCase()}'`)
      .header('Prefer', 'HonorNonIndexedQueriesWarningMayFailRandomly')
      .get();
    const items = resp.value || [];
    const subs = [];
    for (const item of items) {
      const f = item.fields || {};
      const endpoint = f[map['Endpoint']] || f.Endpoint || '';
      const p256dh = f[map['P256DH']] || f.P256DH || '';
      const auth = f[map['Auth']] || f.Auth || '';
      if (endpoint && p256dh && auth) {
        subs.push({ spId: item.id, endpoint, keys: { p256dh, auth } });
      }
    }
    return subs;
  } catch (e) {
    console.warn('[pushNotif] Erro listando subscriptions de ' + email + ': ' + (e.message || e));
    return [];
  }
}

async function salvarSubscription(client, siteId, email, subscription, userAgent) {
  if (!email || !subscription || !subscription.endpoint) {
    throw new Error('email + subscription.endpoint obrigatorios');
  }
  const map = await getColMap(client, siteId, LIST_SUBSCRIPTIONS);
  const listId = map._listId;

  // Verifica se ja existe (mesma combinacao email + endpoint) — evita duplicata
  try {
    const existing = await client
      .api(`/sites/${siteId}/lists/${listId}/items?expand=fields&$top=20`)
      .filter(`fields/Title eq '${email.toLowerCase()}'`)
      .header('Prefer', 'HonorNonIndexedQueriesWarningMayFailRandomly')
      .get();
    const items = existing.value || [];
    for (const item of items) {
      const f = item.fields || {};
      const endpoint = f[map['Endpoint']] || f.Endpoint || '';
      if (endpoint === subscription.endpoint) {
        // Ja existe — atualiza keys (podem ter mudado por re-permissao)
        const fields = {};
        if (map['P256DH']) fields[map['P256DH']] = subscription.keys && subscription.keys.p256dh || '';
        if (map['Auth']) fields[map['Auth']] = subscription.keys && subscription.keys.auth || '';
        if (map['UserAgent']) fields[map['UserAgent']] = (userAgent || '').slice(0, 250);
        await client.api(`/sites/${siteId}/lists/${listId}/items/${item.id}/fields`).update(fields);
        return { id: item.id, atualizado: true };
      }
    }
  } catch (e) {
    // Se falhar consulta, segue criando — o pior cenario eh duplicata
  }

  // Cria novo
  const fields = {
    Title: email.toLowerCase()
  };
  if (map['Endpoint'])  fields[map['Endpoint']]  = subscription.endpoint;
  if (map['P256DH'])    fields[map['P256DH']]    = (subscription.keys && subscription.keys.p256dh) || '';
  if (map['Auth'])      fields[map['Auth']]      = (subscription.keys && subscription.keys.auth) || '';
  if (map['UserAgent']) fields[map['UserAgent']] = (userAgent || '').slice(0, 250);

  const created = await client.api(`/sites/${siteId}/lists/${listId}/items`).post({ fields });
  return { id: created.id, criado: true };
}

async function removerSubscription(client, siteId, endpoint) {
  if (!endpoint) return { removidos: 0 };
  const map = await getColMap(client, siteId, LIST_SUBSCRIPTIONS);
  const listId = map._listId;
  // Lista por endpoint — pode ser caro mas sao poucos registros
  const resp = await client
    .api(`/sites/${siteId}/lists/${listId}/items?expand=fields&$top=200`)
    .get();
  let removidos = 0;
  for (const item of (resp.value || [])) {
    const f = item.fields || {};
    const ep = f[map['Endpoint']] || f.Endpoint || '';
    if (ep === endpoint) {
      try {
        await client.api(`/sites/${siteId}/lists/${listId}/items/${item.id}`).delete();
        removidos++;
      } catch (e) {}
    }
  }
  return { removidos };
}

// Envia push pra um email — busca todas suas subscriptions e envia em paralelo.
// Subscriptions inativas (410 Gone) sao removidas automaticamente.
async function enviarPushPraEmail(client, siteId, email, payload) {
  if (!configurarWebPush()) {
    console.warn('[pushNotif] Push desabilitado (sem VAPID keys). Skip.');
    return { enviados: 0, erros: 0, removidos: 0, skip: 'no_vapid' };
  }
  const subs = await listarSubscriptionsPorEmail(client, siteId, email);
  if (!subs.length) return { enviados: 0, erros: 0, removidos: 0, skip: 'no_subscriptions' };

  const payloadStr = typeof payload === 'string' ? payload : JSON.stringify(payload);
  let enviados = 0, erros = 0, removidos = 0;

  await Promise.all(subs.map(async (s) => {
    try {
      await webpush.sendNotification({
        endpoint: s.endpoint,
        keys: s.keys
      }, payloadStr, { TTL: 60 * 60 * 24 }); // TTL 24h
      enviados++;
    } catch (e) {
      // 404/410 = subscription invalida/expirada — remove
      const status = e && (e.statusCode || e.status);
      if (status === 404 || status === 410) {
        try {
          await removerSubscription(client, siteId, s.endpoint);
          removidos++;
        } catch (e2) {}
      } else {
        erros++;
        console.warn('[pushNotif] Erro enviando push: ' + (e.message || e) + ' (status ' + status + ')');
      }
    }
  }));

  return { enviados, erros, removidos };
}

module.exports = {
  configurarWebPush,
  listarSubscriptionsPorEmail,
  salvarSubscription,
  removerSubscription,
  enviarPushPraEmail,
  LIST_SUBSCRIPTIONS
};
