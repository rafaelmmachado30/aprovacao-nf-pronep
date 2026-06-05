/**
 * /api/MigrarColunasURL — one-shot: cria colunas UrlPDFAprovadoStr e UrlPDFStr
 * como text multi-line na lista PRONEP-NF-NotasFiscais (resolve o bug das colunas
 * UrlPDFAprovado/UrlPDF originais que estao com tipo nao reconhecido pelo Graph).
 *
 * Idempotente: se a coluna ja existe, retorna sucesso sem recriar.
 *
 * RBAC: anonymous mas com Sites.Manage.All do App Reg. Recomendado rodar apenas 1x.
 */
require('isomorphic-fetch');
const { ClientSecretCredential } = require('@azure/identity');
const { Client } = require('@microsoft/microsoft-graph-client');
const { TokenCredentialAuthenticationProvider } = require('@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials');

module.exports = async function (context, req) {
  const out = { timestamp: new Date().toISOString(), colunasCriadas: [], colunasJaExistiam: [], erros: [] };
  try {
    const credential = new ClientSecretCredential(
      process.env.AAD_TENANT_ID, process.env.AAD_CLIENT_ID, process.env.AAD_CLIENT_SECRET
    );
    const authProvider = new TokenCredentialAuthenticationProvider(credential, {
      scopes: ['https://graph.microsoft.com/.default']
    });
    const client = Client.initWithMiddleware({ authProvider });

    const siteResp = await client.api('/sites/' + process.env.SHAREPOINT_SITE_HOSTNAME + ':' + process.env.SHAREPOINT_SITE_PATH).get();
    const siteId = siteResp.id;

    const listas = await client.api('/sites/' + siteId + '/lists').filter("displayName eq 'PRONEP-NF-NotasFiscais'").get();
    if (!listas.value || !listas.value.length) {
      out.erro = 'Lista PRONEP-NF-NotasFiscais nao encontrada';
      context.res = { status: 404, body: out }; return;
    }
    const listId = listas.value[0].id;

    const cols = await client.api('/sites/' + siteId + '/lists/' + listId + '/columns').get();
    const existentes = (cols.value || []).map(c => c.name);

    const novasColunas = [
      { name: 'UrlPDFAprovadoStr', displayName: 'UrlPDFAprovadoStr' },
      { name: 'UrlPDFStr',        displayName: 'UrlPDFStr' }
    ];

    for (const col of novasColunas) {
      if (existentes.indexOf(col.name) !== -1) {
        out.colunasJaExistiam.push(col.name);
        continue;
      }
      try {
        const created = await client.api('/sites/' + siteId + '/lists/' + listId + '/columns').post({
          name: col.name,
          displayName: col.displayName,
          description: 'URL do PDF (string) — substitui ' + col.name.replace('Str', '') + ' que estava com tipo invalido no Graph',
          text: { allowMultipleLines: true, appendChangesToExistingText: false, linesForEditing: 3 }
        });
        out.colunasCriadas.push({ name: created.name, displayName: created.displayName });
      } catch (e) {
        out.erros.push({ coluna: col.name, msg: e.message, statusCode: e.statusCode, body: e.body });
      }
    }

    out.ok = out.erros.length === 0;
    context.res = { status: 200, headers: { 'Content-Type': 'application/json' }, body: out };
  } catch (e) {
    out.erro = e.message;
    out.stack = (e.stack || '').split('\n').slice(0, 6);
    context.res = { status: 500, headers: { 'Content-Type': 'application/json' }, body: out };
  }
};
