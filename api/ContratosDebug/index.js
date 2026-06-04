/**
 * /api/ContratosDebug — diagnostico isolado das peças do modulo de Contratos.
 * Anonymous, retorna info sobre o que carrega/funciona/falha.
 *
 * Query params:
 *   ?step=env|require|graph|site|listar
 *   ?pasta=/GERENCIA DE PROJETOS E TI/CORPORATIVO  (so usado em step=listar)
 */
module.exports = async function (context, req) {
  const out = {
    timestamp: new Date().toISOString(),
    step: (req.query && req.query.step) || 'all',
    env: {},
    requires: {},
    graph: null,
    site: null,
    listar: null
  };

  // 1. Env vars
  out.env = {
    hasTenant: !!process.env.AAD_TENANT_ID,
    hasClient: !!process.env.AAD_CLIENT_ID,
    hasSecret: !!process.env.AAD_CLIENT_SECRET,
    hasNFHostname: !!process.env.SHAREPOINT_SITE_HOSTNAME,
    hasNFPath: !!process.env.SHAREPOINT_SITE_PATH,
    hasContratosHostname: !!process.env.SHAREPOINT_CONTRATOS_HOSTNAME,
    hasContratosPath: !!process.env.SHAREPOINT_CONTRATOS_PATH,
    contratosHostname: process.env.SHAREPOINT_CONTRATOS_HOSTNAME || null,
    contratosPath: process.env.SHAREPOINT_CONTRATOS_PATH || null,
    hasAnthropic: !!process.env.ANTHROPIC_API_KEY
  };

  // 2. Requires
  const tries = ['../shared/contratos', '../shared/auth', 'pdf-parse', 'mammoth', '@anthropic-ai/sdk', '@azure/identity', '@microsoft/microsoft-graph-client'];
  for (const m of tries) {
    try {
      const r = require(m);
      out.requires[m] = { ok: true, keys: Object.keys(r || {}).slice(0, 10) };
    } catch (e) {
      out.requires[m] = { ok: false, error: e.message };
    }
  }

  // 3. Graph client (lazy)
  if (out.env.hasTenant && out.env.hasClient && out.env.hasSecret) {
    try {
      const { ClientSecretCredential } = require('@azure/identity');
      const { Client } = require('@microsoft/microsoft-graph-client');
      const { TokenCredentialAuthenticationProvider } = require('@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials');
      const credential = new ClientSecretCredential(
        process.env.AAD_TENANT_ID,
        process.env.AAD_CLIENT_ID,
        process.env.AAD_CLIENT_SECRET
      );
      const authProvider = new TokenCredentialAuthenticationProvider(credential, {
        scopes: ['https://graph.microsoft.com/.default']
      });
      const client = require('@microsoft/microsoft-graph-client').Client.initWithMiddleware({ authProvider });
      out.graph = { instanciado: true };

      // 4. Site de Contratos
      const host = process.env.SHAREPOINT_CONTRATOS_HOSTNAME || 'pronepadmin.sharepoint.com';
      const path = process.env.SHAREPOINT_CONTRATOS_PATH || '/sites/CONTRATOS-SERVICOS-CONTRATOS';
      try {
        const siteResp = await client.api('/sites/' + host + ':' + path).get();
        out.site = { ok: true, siteId: siteResp.id, displayName: siteResp.displayName, webUrl: siteResp.webUrl };
        const driveResp = await client.api('/sites/' + siteResp.id + '/drive').get();
        out.site.driveId = driveResp.id;
      } catch (e) {
        out.site = { ok: false, error: e.message, statusCode: e.statusCode, code: e.code, body: e.body };
      }

      // 5. Listar pasta especifica
      if (out.site && out.site.ok) {
        const pasta = (req.query && req.query.pasta) || '/CONTRATOS/CONTRATOS E DOCUMENTOS - PRESTADORES';
        try {
          const enc = encodeURIComponent(pasta).replace(/%2F/g, '/');
          const url = '/drives/' + out.site.driveId + '/root:' + enc + ':/children?$top=200';
          const resp = await client.api(url).get();
          out.listar = {
            ok: true,
            pasta,
            url,
            count: (resp.value || []).length,
            sample: (resp.value || []).slice(0, 5).map(function(it){
              return {
                name: it.name,
                isFolder: !!it.folder,
                isFile: !!it.file,
                size: it.size || 0,
                ext: it.file ? ((it.name || '').split('.').pop().toLowerCase()) : null
              };
            })
          };
        } catch (e) {
          out.listar = { ok: false, error: e.message, statusCode: e.statusCode, body: e.body };
        }
      }

      // 6. Site do sistema NF (pra garantir criacao de lista)
      try {
        const siteNF = await client.api('/sites/' + process.env.SHAREPOINT_SITE_HOSTNAME + ':' + process.env.SHAREPOINT_SITE_PATH).get();
        out.siteNF = { ok: true, siteId: siteNF.id };
        const lists = await client.api('/sites/' + siteNF.id + "/lists").filter("displayName eq 'PRONEP-NF-Contratos'").get();
        out.siteNF.contratosListaExiste = (lists.value && lists.value.length > 0);
        if (out.siteNF.contratosListaExiste) out.siteNF.contratosListaId = lists.value[0].id;
      } catch (e) {
        out.siteNF = { ok: false, error: e.message };
      }
    } catch (e) {
      out.graph = { instanciado: false, error: e.message, stack: (e.stack || '').split('\n').slice(0, 5) };
    }
  }

  context.res = {
    status: 200,
    headers: { 'Content-Type': 'application/json' },
    body: out
  };
};
