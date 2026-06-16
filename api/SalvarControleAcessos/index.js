/**
 * /api/SalvarControleAcessos (POST) — ADMIN ONLY
 *
 * Body (por diretoria):  { diretoria: 'Tecnologia', emails: ['a@x','b@x'] }
 * Body (mapa inteiro):   { mapa: { 'Tecnologia': ['a@x'], 'RH': [...] } }
 *
 * Salva no item Config 'acessoContratos'. Modelo complementar: definir uma diretoria
 * aqui passa a sobrepor o fallback (aprovador de NF) APENAS para aquela diretoria.
 * Para REMOVER a config explicita de uma diretoria (voltar ao fallback), envie emails: [].
 */

require('isomorphic-fetch');
const { requireAdmin } = require('../shared/authz');
const { lerMapaAcessos, salvarMapaAcessos, resolveConfigListId } = require('../shared/acessoContratos');
const { registrar: auditRegistrar } = require('../shared/auditLog');
const { ClientSecretCredential } = require('@azure/identity');
const { Client } = require('@microsoft/microsoft-graph-client');
const { TokenCredentialAuthenticationProvider } =
  require('@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials');

function getGraphClient() {
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

async function resolveSite(client) {
  const host = process.env.SHAREPOINT_SITE_HOSTNAME;
  const path = process.env.SHAREPOINT_SITE_PATH;
  if (!host || !path) throw new Error('SHAREPOINT_* incompletas');
  const siteResp = await client.api('/sites/' + host + ':' + path).get();
  return siteResp.id;
}

// Normaliza lista de diretorias: trim, sem vazios, sem duplicatas (case-insensitive).
function normDirs(arr) {
  const out = [];
  const seen = new Set();
  for (const d of (Array.isArray(arr) ? arr : [])) {
    const v = String(d || '').trim();
    const k = v.toLowerCase();
    if (v && !seen.has(k)) { seen.add(k); out.push(v); }
  }
  return out;
}

module.exports = async function (context, req) {
  try {
    const authz = await requireAdmin(context, req);
    if (!authz) return;

    const body = req.body || {};
    const client = getGraphClient();
    const siteId = await resolveSite(client);
    const cfgId = await resolveConfigListId(client, siteId);
    if (!cfgId) {
      context.res = { status: 500, body: { error: "Lista 'PRONEP-NF-Config' nao encontrada" } };
      return;
    }

    const mapaAtual = await lerMapaAcessos(client, siteId, cfgId);
    let novoMapa;

    if (body.mapa && typeof body.mapa === 'object') {
      // Substitui o mapa inteiro (pasta -> diretorias com acesso)
      novoMapa = {};
      for (const d of Object.keys(body.mapa)) {
        const ds = normDirs(body.mapa[d]);
        if (ds.length) novoMapa[String(d).trim()] = ds;
      }
    } else if (body.diretoria) {
      // Atualiza apenas uma pasta. body.diretoria = pasta; body.diretorias = diretorias liberadas.
      novoMapa = Object.assign({}, mapaAtual);
      const folder = String(body.diretoria).trim();
      const ds = normDirs(body.diretorias);
      if (ds.length) novoMapa[folder] = ds;
      else delete novoMapa[folder]; // lista vazia = remove config explicita (volta ao fallback)
    } else {
      context.res = { status: 400, body: { error: 'Envie { diretoria, diretorias } ou { mapa }' } };
      return;
    }

    const r = await salvarMapaAcessos(client, siteId, cfgId, novoMapa);

    auditRegistrar(authz.user, 'config_update',
      { tipo: 'acesso_contratos', id: r.itemId },
      'sucesso',
      { action: r.action, diretoria: body.diretoria || '(mapa completo)', acessos: novoMapa }
    ).catch(function () {});

    context.res = {
      status: 200,
      headers: { 'Content-Type': 'application/json' },
      body: { ok: true, itemId: r.itemId, action: r.action, acessos: novoMapa }
    };
  } catch (err) {
    context.log && context.log.error && context.log.error('SalvarControleAcessos:', err);
    context.res = { status: 500, headers: { 'Content-Type': 'application/json' },
      body: { error: (err && err.message) || String(err) } };
  }
};
