/**
 * /api/AdminLimparBase
 *
 * APAGA TODAS AS NFs da base e remove os PDFs das pastas Pendentes/Aprovadas/Rejeitadas.
 *
 * **PROTECAO TRIPLA:**
 *  1. Usuario deve ter role 'administrador' (grupo PRONEP-NF-Admin no Entra ID)
 *  2. Body deve conter { confirmacao: "LIMPAR" } literal
 *  3. Method POST obrigatorio
 *
 * NAO MEXE:
 *  - PRONEP-NF-Fornecedores
 *  - PRONEP-NF-Diretorias
 *  - PRONEP-NF-Config
 *
 * Retorna:
 *  200 { ok, removidos: { notas, pdfsPendentes, pdfsAprovadas, pdfsRejeitadas } }
 *  401 { error: "Nao autenticado" }
 *  403 { error: "Sem permissao - precisa ser administrador" }
 *  400 { error: "Confirmacao invalida" }
 *  500 { error, diag }
 */

require('isomorphic-fetch');
const { getUser } = require('../shared/auth');
const { ClientSecretCredential } = require('@azure/identity');
const { Client } = require('@microsoft/microsoft-graph-client');
const { TokenCredentialAuthenticationProvider } =
  require('@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials');

const LIST_NOTAS = 'PRONEP-NF-NotasFiscais';
const ADMIN_GROUP_ID = '480a1595-bdc3-492a-9ef2-317f148a237e';
// IMPORTANTE: a pasta "Notas Fiscais/Notas Aprovadas" foi reaproveitada de uma estrutura
// pre-existente do SharePoint da Pronep e contem historico legacy da empresa. ENQUANTO
// nao migrarmos pra uma pasta nova exclusiva do projeto (ex: "Aprovados"), ela NAO entra
// na limpeza — apagariamos historico real.
// Quando migrar pra pasta exclusiva, reincluir o caminho neste array.
const PASTAS = [
  'Notas Fiscais/Pendentes',
  // 'Notas Fiscais/Notas Aprovadas', // DESABILITADO ate migracao da pasta legacy
  'Notas Fiscais/Rejeitadas'
];

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

// Verifica se o user tem role administrador via Graph (transitiveMemberOf)
async function ehAdministrador(client, userKey) {
  try {
    const result = await client
      .api(`/users/${encodeURIComponent(userKey)}/transitiveMemberOf`)
      .select('id')
      .top(200)
      .get();
    const items = (result && result.value) || [];
    return items.some(g => g && g.id && String(g.id).toLowerCase() === ADMIN_GROUP_ID.toLowerCase());
  } catch (e) {
    return false;
  }
}

async function resolveSiteAndList(client) {
  const host = process.env.SHAREPOINT_SITE_HOSTNAME;
  const path = process.env.SHAREPOINT_SITE_PATH;
  if (!host || !path) throw new Error('SHAREPOINT_* incompletas');
  const siteResp = await client.api(`/sites/${host}:${path}`).get();
  const siteId = siteResp.id;
  const lists = await client.api(`/sites/${siteId}/lists`).filter(`displayName eq '${LIST_NOTAS}'`).get();
  if (!lists.value || !lists.value.length) throw new Error(`Lista '${LIST_NOTAS}' nao encontrada`);
  const listId = lists.value[0].id;
  return { siteId, listId };
}

// Apaga todos os items da lista (com paginacao)
async function apagarTodasNotas(client, siteId, listId, context) {
  let removidos = 0;
  let erros = 0;
  let nextLink = `/sites/${siteId}/lists/${listId}/items?$top=100&$select=id`;
  while (nextLink) {
    const resp = await client.api(nextLink).get();
    const items = (resp.value || []);
    for (const item of items) {
      try {
        await client.api(`/sites/${siteId}/lists/${listId}/items/${item.id}`).delete();
        removidos++;
      } catch (e) {
        erros++;
        if (context && context.log) context.log.warn('Erro ao deletar item ' + item.id + ': ' + (e.message || e));
      }
    }
    // Pagina
    nextLink = resp['@odata.nextLink'] || null;
    if (nextLink && nextLink.startsWith('https://graph.microsoft.com/v1.0')) {
      nextLink = nextLink.substring('https://graph.microsoft.com/v1.0'.length);
    }
  }
  return { removidos, erros };
}

// Recursivamente apaga TODOS os arquivos PDF de uma pasta SP (e subpastas)
async function apagarArquivosDaPasta(client, siteId, pastaPath, context) {
  let removidos = 0;
  let erros = 0;
  try {
    const resp = await client.api(`/sites/${siteId}/drive/root:/${pastaPath}:/children?$top=200`).get();
    const items = (resp.value || []);
    for (const item of items) {
      if (item.file) {
        // E um arquivo - delete
        try {
          await client.api(`/sites/${siteId}/drive/items/${item.id}`).delete();
          removidos++;
        } catch (e) {
          erros++;
          if (context && context.log) context.log.warn('Erro ao deletar file ' + item.name + ': ' + (e.message || e));
        }
      } else if (item.folder) {
        // E uma subpasta - recurse
        const subPath = pastaPath + '/' + item.name;
        const sub = await apagarArquivosDaPasta(client, siteId, subPath, context);
        removidos += sub.removidos;
        erros += sub.erros;
        // Apos esvaziar a subpasta, deletamos ela tambem (mas so se nao for raiz das 3 principais)
        try {
          await client.api(`/sites/${siteId}/drive/items/${item.id}`).delete();
        } catch (e) {
          // Pode falhar se a pasta ja sumiu, ignora
        }
      }
    }
  } catch (e) {
    // Pasta nao existe — ok, retorna 0
    if (context && context.log) context.log.warn('Pasta nao encontrada (ok): ' + pastaPath + ' — ' + (e.message || e));
  }
  return { removidos, erros };
}

module.exports = async function (context, req) {
  const diag = { step: 'start' };
  try {
    // 1. Autenticacao
    diag.step = 'auth';
    const user = await getUser(req);
    if (!user) {
      context.res = { status: 401, headers: { 'Content-Type': 'application/json' },
        body: { error: 'Nao autenticado', authError: req._authError || null } };
      return;
    }
    diag.user = user.email;

    // 2. Confirmacao do body
    diag.step = 'confirm';
    const body = req.body || {};
    if (body.confirmacao !== 'LIMPAR') {
      context.res = { status: 400, headers: { 'Content-Type': 'application/json' },
        body: { error: 'Confirmacao invalida. Body deve conter { confirmacao: "LIMPAR" } (exato).' } };
      return;
    }

    // 3. Cliente Graph + verificacao de role admin
    diag.step = 'graph_client';
    const client = getGraphClient();

    diag.step = 'check_admin';
    const userKey = (user.oid && /^[0-9a-f-]{32,36}$/i.test(user.oid)) ? user.oid : user.email;
    const isAdmin = await ehAdministrador(client, userKey);
    if (!isAdmin) {
      context.res = { status: 403, headers: { 'Content-Type': 'application/json' },
        body: { error: 'Sem permissao. Esta operacao eh restrita a administradores (grupo PRONEP-NF-Admin).' } };
      return;
    }
    diag.adminConfirmado = true;

    // 4. Resolver site e lista
    diag.step = 'resolve_site';
    const { siteId, listId } = await resolveSiteAndList(client);
    diag.siteId = siteId;

    // 5. Apagar todas as NFs
    diag.step = 'delete_notas';
    const notasResult = await apagarTodasNotas(client, siteId, listId, context);
    diag.notasRemovidas = notasResult.removidos;
    diag.notasErros = notasResult.erros;

    // 6. Apagar PDFs das 3 pastas
    diag.step = 'delete_pdfs';
    const pdfsPorPasta = {};
    for (const pasta of PASTAS) {
      const r = await apagarArquivosDaPasta(client, siteId, pasta, context);
      pdfsPorPasta[pasta] = r;
    }
    diag.pdfsPorPasta = pdfsPorPasta;

    // 7. Log do evento (auditoria minima — quem limpou e quando)
    if (context.log) {
      context.log('[AdminLimparBase] EXECUTADO por ' + user.email + ' as ' + new Date().toISOString() +
        ' | notas removidas: ' + notasResult.removidos +
        ' | erros notas: ' + notasResult.erros);
    }

    context.res = {
      status: 200,
      headers: { 'Content-Type': 'application/json' },
      body: {
        ok: true,
        executadoPor: user.email,
        executadoEm: new Date().toISOString(),
        removidos: {
          notas: notasResult.removidos,
          notasErros: notasResult.erros,
          pdfsPendentes: (pdfsPorPasta['Notas Fiscais/Pendentes'] || {}).removidos || 0,
          pdfsRejeitadas: (pdfsPorPasta['Notas Fiscais/Rejeitadas'] || {}).removidos || 0
        },
        preservados: {
          pastaNotasAprovadas: 'NAO TOCADA — contem historico legacy da Pronep. Sera incluida quando migrarmos pra pasta exclusiva do projeto.'
        }
      }
    };
  } catch (err) {
    if (context.log && context.log.error) context.log.error('AdminLimparBase erro:', err);
    context.res = {
      status: 500,
      headers: { 'Content-Type': 'application/json' },
      body: { error: (err && err.message) || String(err), diag }
    };
  }
};
