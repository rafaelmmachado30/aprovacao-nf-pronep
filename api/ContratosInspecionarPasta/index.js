/**
 * /api/ContratosInspecionarPasta?path=/CONTRATOS/.../NOME_PASTA
 *
 * Dump completo dos metadados que o Graph API expoe pra uma pasta.
 * Usado pra descobrir como detectar pastas marcadas (cor vermelha = cancelado).
 *
 * Retorna:
 *   - driveItem completo (com expand de listItem)
 *   - listItem (com fields) — onde costuma estar a tag de cor
 *   - children: id+name+listItem.fields de cada filho (pra comparar)
 */
module.exports = async function (context, req) {
  // C5: dump arbitrario de pasta do SharePoint — restrito a admin.
  const { requireAdmin } = require('../shared/authz');
  if (!(await requireAdmin(context, req))) return;
  try {
    const pasta = (req.query && req.query.path) || '';
    if (!pasta) {
      context.res = { status: 400, body: { error: 'Passe ?path=...' } };
      return;
    }

    const { ClientSecretCredential } = require('@azure/identity');
    const { Client } = require('@microsoft/microsoft-graph-client');
    const { TokenCredentialAuthenticationProvider } = require('@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials');
    const credential = new ClientSecretCredential(
      process.env.AAD_TENANT_ID, process.env.AAD_CLIENT_ID, process.env.AAD_CLIENT_SECRET
    );
    const authProvider = new TokenCredentialAuthenticationProvider(credential, {
      scopes: ['https://graph.microsoft.com/.default']
    });
    const client = Client.initWithMiddleware({ authProvider });

    const host = process.env.SHAREPOINT_CONTRATOS_HOSTNAME || 'pronepadmin.sharepoint.com';
    const sitePath = process.env.SHAREPOINT_CONTRATOS_PATH || '/sites/CONTRATOS-SERVICOS-CONTRATOS';
    const site = await client.api('/sites/' + host + ':' + sitePath).get();
    const drive = await client.api('/sites/' + site.id + '/drive').get();
    const driveId = drive.id;

    const enc = encodeURIComponent(pasta).replace(/%2F/g, '/');

    // Pega item da pasta com expand do listItem (onde os campos custom estao)
    const item = await client.api('/drives/' + driveId + '/root:' + enc + ':?expand=listItem($expand=fields)').get();

    // Pega filhos (1 nivel) com expand tambem pra comparar
    const children = await client.api('/drives/' + driveId + '/root:' + enc + ':/children?$top=200&expand=listItem($expand=fields)').get();

    // Filtra os campos NAO-padrao do listItem.fields (provavel onde a cor vive)
    const camposPadrao = new Set(['@odata.etag', 'Title', 'ContentType', 'Modified', 'Created',
      'AuthorLookupId', 'EditorLookupId', '_UIVersionString', 'Attachments', 'Edit',
      'LinkTitleNoMenu', 'LinkTitle', 'ItemChildCount', 'FolderChildCount', '_ComplianceFlags',
      '_ComplianceTag', '_ComplianceTagWrittenTime', '_ComplianceTagUserId', '_IsRecord',
      'AppAuthorLookupId', 'AppEditorLookupId', 'id', 'FileSizeDisplay']);

    function camposCustom(fields) {
      if (!fields) return {};
      const out = {};
      for (const k of Object.keys(fields)) {
        if (!camposPadrao.has(k)) out[k] = fields[k];
      }
      return out;
    }

    const itemFields = (item.listItem && item.listItem.fields) || {};

    const filhosResumo = (children.value || []).slice(0, 25).map(function(c) {
      return {
        name: c.name,
        isFolder: !!c.folder,
        isFile: !!c.file,
        listItemFields: c.listItem && c.listItem.fields ? c.listItem.fields : null,
        camposCustom: c.listItem && c.listItem.fields ? camposCustom(c.listItem.fields) : null
      };
    });

    context.res = {
      status: 200,
      headers: { 'Content-Type': 'application/json' },
      body: {
        ok: true,
        pasta,
        item: {
          id: item.id,
          name: item.name,
          webUrl: item.webUrl,
          createdDateTime: item.createdDateTime,
          lastModifiedDateTime: item.lastModifiedDateTime,
          listItemId: item.listItem && item.listItem.id,
          allListItemFields: itemFields,
          camposCustom: camposCustom(itemFields)
        },
        children: filhosResumo
      }
    };
  } catch (err) {
    context.res = {
      status: 500,
      headers: { 'Content-Type': 'application/json' },
      body: { error: err.message, statusCode: err.statusCode, body: err.body, stack: (err.stack || '').split('\n').slice(0, 6) }
    };
  }
};
