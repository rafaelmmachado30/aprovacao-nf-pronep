/**
 * /api/CriarListaContratos
 *
 * Endpoint dedicado pra criar (ou validar) a lista PRONEP-NF-Contratos.
 * NAO chama Claude. NAO faz crawl. Operacao 100% sobre SharePoint.
 *
 * Comportamento:
 *   1. Procura a lista pelo displayName
 *   2. Se nao existe, cria via Graph (sem colunas iniciais alem das default)
 *   3. Apos achar/criar, valida que todas as colunas esperadas existem
 *   4. Cria as colunas que faltam (uma a uma — pra isolar falha de schema)
 *   5. Retorna diagnostico detalhado por etapa
 *
 * RBAC: admin only.
 *
 * Custo Claude: ZERO.
 */

require('isomorphic-fetch');
const { getGraphClient } = require('../shared/graph');

const LIST_NAME = 'PRONEP-NF-Contratos';

// Schema completo das colunas esperadas. Cada entrada: { name, def }
// def usa a sintaxe do Graph (text, dateTime, number, etc).
const COLUNAS_ESPERADAS = [
  { name: 'Diretoria',          def: { text: {} } },
  { name: 'Unidade',            def: { text: {} } },
  { name: 'Fornecedor',         def: { text: {} } },
  { name: 'CNPJFornecedor',     def: { text: {} } },
  { name: 'DataInicio',         def: { dateTime: { displayAs: 'standard', format: 'dateOnly' } } },
  { name: 'DataFim',            def: { dateTime: { displayAs: 'standard', format: 'dateOnly' } } },
  { name: 'Status',             def: { text: {} } },
  { name: 'Situacao',           def: { text: {} } },
  { name: 'CaminhoSharepoint',  def: { text: { allowMultipleLines: true } } },
  { name: 'DriveItemId',        def: { text: {} } },
  { name: 'LeituraIAStatus',    def: { text: {} } },
  { name: 'LeituraIATexto',     def: { text: { allowMultipleLines: true } } },
  { name: 'ValorContrato',      def: { number: {} } },
  { name: 'Observacoes',        def: { text: { allowMultipleLines: true } } },
  { name: 'PathRelativoSP',     def: { text: { allowMultipleLines: true } } },
  { name: 'UltimaLeitura',      def: { dateTime: { displayAs: 'standard', format: 'dateTime' } } },
  { name: 'NomeArquivo',        def: { text: {} } },
  { name: 'TamanhoArquivo',     def: { number: {} } }
];

function readClientPrincipal(req) {
  const header = req.headers && req.headers['x-ms-client-principal'];
  if (!header) return null;
  try { return JSON.parse(Buffer.from(header, 'base64').toString('utf-8')); } catch (e) { return null; }
}

async function isAdmin(req) {
  const p = readClientPrincipal(req);
  const roles = (p && p.userRoles) || [];
  if (roles.includes('administrador') || roles.includes('admin')) return true;
  try {
    const { getUser } = require('../shared/auth');
    const user = await getUser(req);
    if (!user || !user.oid) return false;
    const { getUserRoles } = require('../shared/userRoles');
    const userRoles = await getUserRoles(user);
    return (userRoles || []).includes('administrador');
  } catch (e) {
    return false;
  }
}

module.exports = async function (context, req) {
  const diag = {
    step: 'init',
    listName: LIST_NAME,
    site: null,
    listaJaExistia: null,
    listaCriada: false,
    listIdAposGarantir: null,
    colunasExistentes: [],
    colunasCriadas: [],
    colunasComFalha: [],
    erros: [],
    timeMs: 0
  };
  const t0 = Date.now();
  try {
    if (!(await isAdmin(req))) {
      context.res = { status: 403, headers: { 'Content-Type': 'application/json' },
        body: { error: 'Apenas admin' } };
      return;
    }

    const client = await getGraphClient();

    // ETAPA 1: resolver site
    diag.step = 'resolve_site';
    const host = process.env.SHAREPOINT_SITE_HOSTNAME;
    const path = process.env.SHAREPOINT_SITE_PATH;
    if (!host || !path) throw new Error('SHAREPOINT_SITE_HOSTNAME/PATH nao configurados');
    const siteResp = await client.api('/sites/' + host + ':' + path).get();
    const siteId = siteResp.id;
    diag.site = { host, path, siteId, webUrl: siteResp.webUrl };

    // ETAPA 2: procurar lista
    diag.step = 'search_list';
    let listId = null;
    try {
      const lists = await client.api('/sites/' + siteId + '/lists')
        .filter("displayName eq '" + LIST_NAME + "'").get();
      if (lists.value && lists.value.length) {
        listId = lists.value[0].id;
        diag.listaJaExistia = true;
        diag.listIdAposGarantir = listId;
      } else {
        diag.listaJaExistia = false;
      }
    } catch (eSearch) {
      // Filter as vezes da problema. Tenta lista geral e filtra manualmente
      diag.erros.push({ step: 'search_list_filter', error: eSearch.message, body: eSearch.body });
      try {
        const allLists = await client.api('/sites/' + siteId + '/lists').get();
        const found = (allLists.value || []).find(function(l){ return l.displayName === LIST_NAME; });
        if (found) {
          listId = found.id;
          diag.listaJaExistia = true;
          diag.listIdAposGarantir = listId;
          diag.erros.push({ step: 'search_list_fallback', message: 'achou via listagem completa' });
        } else {
          diag.listaJaExistia = false;
        }
      } catch (eListAll) {
        diag.erros.push({ step: 'search_list_listall', error: eListAll.message });
        throw new Error('Nao consigo listar listas do site: ' + eListAll.message);
      }
    }

    // ETAPA 3: se nao existe, cria (sem colunas extras - so default)
    if (!listId) {
      diag.step = 'create_list';
      try {
        const newList = await client.api('/sites/' + siteId + '/lists').post({
          displayName: LIST_NAME,
          list: { template: 'genericList' }
        });
        listId = newList.id;
        diag.listaCriada = true;
        diag.listIdAposGarantir = listId;
      } catch (eCreate) {
        diag.erros.push({
          step: 'create_list',
          error: eCreate.message,
          graphCode: eCreate.code,
          graphStatusCode: eCreate.statusCode,
          graphBody: eCreate.body,
          requestId: eCreate.requestId,
          stack: (eCreate.stack || '').split('\n').slice(0, 5)
        });
        // Retorna 500 com diagnostico claro
        diag.timeMs = Date.now() - t0;
        context.res = { status: 500, headers: { 'Content-Type': 'application/json' },
          body: Object.assign({ error: 'Falha ao criar lista: ' + eCreate.message }, diag) };
        return;
      }
    }

    // ETAPA 4: listar colunas atuais
    diag.step = 'list_columns';
    let colsResp;
    try {
      colsResp = await client.api('/sites/' + siteId + '/lists/' + listId + '/columns').get();
    } catch (eCols) {
      diag.erros.push({ step: 'list_columns', error: eCols.message });
      throw eCols;
    }
    const colsExistentes = new Set();
    for (const c of (colsResp.value || [])) {
      if (c.displayName) colsExistentes.add(c.displayName);
      if (c.name) colsExistentes.add(c.name);
    }
    diag.colunasExistentes = Array.from(colsExistentes).filter(function(n){
      // Filtra colunas default do SP (Title, ID, etc) pra economizar payload
      return !['ID','Title','ContentType','Modified','Created','Author','Editor','_UIVersionString','Attachments','Edit','LinkTitleNoMenu','LinkTitle','DocIcon','ItemChildCount','FolderChildCount','_ComplianceFlags','_ComplianceTag','_ComplianceTagWrittenTime','_ComplianceTagUserId','AppAuthor','AppEditor','ContentTypeId'].includes(n);
    });

    // ETAPA 5: criar colunas que faltam (uma a uma — isola falha de schema)
    diag.step = 'create_missing_columns';
    for (const col of COLUNAS_ESPERADAS) {
      if (colsExistentes.has(col.name)) continue;
      try {
        const payload = Object.assign({ name: col.name }, col.def);
        await client.api('/sites/' + siteId + '/lists/' + listId + '/columns').post(payload);
        diag.colunasCriadas.push(col.name);
      } catch (eCol) {
        diag.colunasComFalha.push({
          coluna: col.name,
          error: eCol.message,
          graphCode: eCol.code,
          graphStatusCode: eCol.statusCode,
          graphBody: eCol.body
        });
      }
    }

    diag.step = 'done';
    diag.timeMs = Date.now() - t0;
    context.res = {
      status: 200,
      headers: { 'Content-Type': 'application/json' },
      body: Object.assign({
        ok: true,
        mensagem: diag.listaCriada
          ? 'Lista ' + LIST_NAME + ' CRIADA com sucesso. Colunas adicionadas: ' + diag.colunasCriadas.length
          : (diag.listaJaExistia
              ? 'Lista ' + LIST_NAME + ' ja existia. Colunas adicionadas agora: ' + diag.colunasCriadas.length + ' (' + diag.colunasComFalha.length + ' falharam)'
              : 'Estado inesperado'),
        sharepointUrl: (diag.site && diag.site.webUrl ? diag.site.webUrl + '/Lists/' + LIST_NAME : null)
      }, diag)
    };
  } catch (err) {
    diag.timeMs = Date.now() - t0;
    context.log && context.log.error && context.log.error('CriarListaContratos:', err);
    context.res = {
      status: 500,
      headers: { 'Content-Type': 'application/json' },
      body: Object.assign({ error: err.message, stack: (err.stack || '').split('\n').slice(0, 8) }, diag)
    };
  }
};
