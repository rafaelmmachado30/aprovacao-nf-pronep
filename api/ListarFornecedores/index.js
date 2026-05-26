/**
 * Sistema de Aprovacao de NF - ListarFornecedores
 *
 * Le a lista PRONEP-NF-Fornecedores no SharePoint via Graph API.
 * Suporta paginacao (nextLink), filtros e ordenacao por busca.
 *
 * Query params:
 *   ?q=texto       -> filtra por Title/NomeFantasia/Documento
 *   ?ativo=Sim     -> filtra por status ativo
 *   ?limit=50      -> limita resultado (default: 5000, max: 5000)
 *
 * App Settings exigidas:
 *   - AAD_CLIENT_ID, AAD_CLIENT_SECRET, AAD_TENANT_ID
 *   - SHAREPOINT_SITE_HOSTNAME (ex: pronepadmin.sharepoint.com)
 *   - SHAREPOINT_SITE_PATH (ex: /sites/Aprovacao-NotasFiscaisServicos)
 *
 * Permissoes Graph (Application, com admin consent):
 *   - Sites.Read.All  ou  Sites.ReadWrite.All
 */

require('isomorphic-fetch');
const { getUser } = require('../shared/auth');
const { ClientSecretCredential } = require('@azure/identity');
const { Client } = require('@microsoft/microsoft-graph-client');
const { TokenCredentialAuthenticationProvider } =
  require('@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials');

const LIST_NAME = 'PRONEP-NF-Fornecedores';

// Cache do site/list id em memoria (Functions reaproveita instancia)
const cache = { siteId: null, listId: null };

async function getGraphClient() {
  const tenantId = process.env.AAD_TENANT_ID;
  const clientId = process.env.AAD_CLIENT_ID;
  const clientSecret = process.env.AAD_CLIENT_SECRET;
  if (!tenantId || !clientId || !clientSecret) {
    throw new Error('App Settings AAD_* incompletas');
  }
  const credential = new ClientSecretCredential(tenantId, clientId, clientSecret);
  const authProvider = new TokenCredentialAuthenticationProvider(credential, {
    scopes: ['https://graph.microsoft.com/.default']
  });
  return Client.initWithMiddleware({ authProvider });
}

async function resolveSiteAndList(client) {
  if (cache.siteId && cache.listId) return cache;

  const host = process.env.SHAREPOINT_SITE_HOSTNAME;
  const path = process.env.SHAREPOINT_SITE_PATH;
  if (!host || !path) {
    throw new Error('App Settings SHAREPOINT_SITE_HOSTNAME e SHAREPOINT_SITE_PATH obrigatorias');
  }

  // Resolve site
  const siteResp = await client
    .api(`/sites/${host}:${path}`)
    .get();
  cache.siteId = siteResp.id;

  // Resolve list por nome (displayName)
  const listsResp = await client
    .api(`/sites/${cache.siteId}/lists`)
    .filter(`displayName eq '${LIST_NAME}'`)
    .get();
  if (!listsResp.value || !listsResp.value.length) {
    throw new Error(`Lista '${LIST_NAME}' nao encontrada no site`);
  }
  cache.listId = listsResp.value[0].id;
  return cache;
}

module.exports = async function (context, req) {
  const diag = { step: 'start' };
  try {
    const q     = (req.query && req.query.q) ? String(req.query.q).toLowerCase() : '';
    const ativo = (req.query && req.query.ativo) ? String(req.query.ativo) : '';
    const limit = Math.min(parseInt((req.query && req.query.limit) || '5000', 10), 5000);

    diag.step = 'graph_client';
    const client = await getGraphClient();

    diag.step = 'resolve_site';
    const { siteId, listId } = await resolveSiteAndList(client);
    diag.siteId = siteId; diag.listId = listId;

    diag.step = 'fetch_items';
    // Pega items com fields. Lista grande -> paginar (todas as paginas, sem cortar)
    const all = [];
    let url = `/sites/${siteId}/lists/${listId}/items?expand=fields&$top=500`;
    let pages = 0;
    while (url) {
      const resp = await client.api(url).get();
      all.push(...(resp.value || []));
      pages++;
      url = resp['@odata.nextLink']
        ? resp['@odata.nextLink'].replace('https://graph.microsoft.com/v1.0','')
        : null;
      // proteção: max 30 paginas (15k items)
      if (pages >= 30) break;
    }
    diag.pages = pages;
    diag.totalItems = all.length;

    diag.step = 'transform';
    // SharePoint renomeou as colunas importadas via XLSX como field_1, field_2, etc
    // Mapeamento descoberto via debug:
    //   Title=Razao, field_1=TipoDoc, field_2=Doc, field_3=Fantasia, field_4=Unidade,
    //   field_5=Diretoria, field_6=UF, field_7=Ativo, field_8=Telefone, field_9=Email,
    //   field_10=Cidade, field_11=CEP, field_12=ApareceNoHist, field_13=QtdNFs
    let fornecedores = all.map(item => {
      const f = item.fields || {};
      return {
        id: item.id,
        razao:             f.Title    || '',
        tipoDocumento:     f.field_1  || '',
        documento:         f.field_2  || '',
        nomeFantasia:      f.field_3  || '',
        unidade:           f.field_4  || '',
        diretoria:         f.field_5  || '',
        uf:                f.field_6  || '',
        ativo:             String(f.field_7 || '').toLowerCase() === 'sim',
        telefone:          f.field_8  || '',
        email:             f.field_9  || '',
        cidade:            f.field_10 || '',
        cep:               f.field_11 || '',
        apareceNoHistorico: String(f.field_12 || '').toLowerCase() === 'sim',
        qtdNFsHistorico:   parseInt(f.field_13 || '0', 10) || 0
      };
    });

    // Filtros server-side
    if (ativo === 'Sim') fornecedores = fornecedores.filter(x => x.ativo);
    if (ativo === 'Nao') fornecedores = fornecedores.filter(x => !x.ativo);
    if (q) {
      fornecedores = fornecedores.filter(x =>
        (x.razao || '').toLowerCase().includes(q) ||
        (x.nomeFantasia || '').toLowerCase().includes(q) ||
        (x.documento || '').toLowerCase().includes(q)
      );
    }
    // Aplica limite no final (depois de filtros)
    const totalFiltrado = fornecedores.length;
    if (limit && fornecedores.length > limit) {
      fornecedores = fornecedores.slice(0, limit);
    }

    context.res = {
      status: 200,
      headers: { 'Content-Type': 'application/json', 'Cache-Control': 'no-store' },
      body: {
        total: fornecedores.length,
        totalFiltrado: totalFiltrado,
        totalAntesFiltros: all.length,
        diag,
        fornecedores
      }
    };
  } catch (err) {
    context.log && context.log.error && context.log.error('ListarFornecedores error:', err);
    context.res = {
      status: 500,
      headers: { 'Content-Type': 'application/json' },
      body: {
        error: (err && err.message) || String(err),
        code: err && err.code,
        statusCode: err && err.statusCode,
        body: err && err.body,
        diag
      }
    };
  }
};
