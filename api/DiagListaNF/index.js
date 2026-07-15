/**
 * /api/DiagListaNF — diagnostico da lista PRONEP-NF-NotasFiscais.
 *
 * Mostra:
 *   - colMap (displayName -> internalName)
 *   - colTypes (internalName -> tipo detectado)
 *   - existencia da coluna UrlPDFAprovado (com tipo e descricao)
 *   - opcionalmente testa PATCH em uma NF (param ?testarPatch=ID_NF)
 *
 * Pra cravar por que o AprovarNota nao esta gravando UrlPDFAprovado.
 */
require('isomorphic-fetch');
const { getGraphClient } = require('../shared/graph');

module.exports = async function (context, req) {
  // C5: diagnostico de lista — restrito a admin.
  const { requireAdmin } = require('../shared/authz');
  if (!(await requireAdmin(context, req))) return;
  const out = { timestamp: new Date().toISOString() };
  try {
    const client = getGraphClient();

    const siteResp = await client.api('/sites/' + process.env.SHAREPOINT_SITE_HOSTNAME + ':' + process.env.SHAREPOINT_SITE_PATH).get();
    const siteId = siteResp.id;
    out.siteId = siteId;

    const listas = await client.api('/sites/' + siteId + '/lists').filter("displayName eq 'PRONEP-NF-NotasFiscais'").get();
    if (!listas.value || !listas.value.length) {
      out.erro = 'Lista PRONEP-NF-NotasFiscais nao encontrada';
      context.res = { status: 200, body: out }; return;
    }
    const listId = listas.value[0].id;
    out.listId = listId;

    const cols = await client.api('/sites/' + siteId + '/lists/' + listId + '/columns').get();
    out.totalColunas = cols.value.length;

    const interesse = ['UrlPDFAprovado', 'UrlPDF', 'urlPDFAprovado', 'urlPDF', 'Status', 'AprovadoEm'];
    out.colunasRelevantes = (cols.value || [])
      .filter(c => interesse.indexOf(c.displayName) !== -1 || interesse.indexOf(c.name) !== -1)
      .map(c => c);  // DUMP COMPLETO — pra ver TODAS as propriedades da coluna no SP

    // Procura ESPECIFICAMENTE qualquer coluna que pareca ser UrlPDFAprovado
    out.suspeitasUrlPDFAprovado = (cols.value || [])
      .filter(c => /url.*pdf.*aprovad/i.test(c.displayName || '') || /url.*pdf.*aprovad/i.test(c.name || ''))
      .map(c => ({
        displayName: c.displayName,
        name: c.name,
        hyperlink: !!c.hyperlinkOrPicture,
        text: !!c.text
      }));

    // Se passou ?criarColunaNova=1&testarPatch=ID, cria coluna hyperlink limpa via Graph
    // e tenta PATCH nela — confirma se o Graph aceita coluna criada por ele mesmo.
    if (req.query && req.query.criarColunaNova) {
      // Tenta criar como TEXT (Multi-line). URLs sao strings, nao precisa ser hyperlink real.
      const nomeColTeste = 'TesteUrlStr_' + Date.now();
      try {
        const created = await client.api('/sites/' + siteId + '/lists/' + listId + '/columns').post({
          name: nomeColTeste,
          displayName: nomeColTeste,
          text: { allowMultipleLines: true, appendChangesToExistingText: false, linesForEditing: 3 }
        });
        out.colunaTeste = { nome: nomeColTeste, internal: created.name, criada: true, tipo: 'text-multiline' };
        if (req.query.testarPatch) {
          const testUrl = 'https://exemplo.com/teste-' + Date.now() + '.pdf';
          const itemId = req.query.testarPatch;
          const patch = {};
          patch[created.name] = testUrl;  // string simples — text aceita
          try {
            await client.api('/sites/' + siteId + '/lists/' + listId + '/items/' + itemId + '/fields').patch(patch);
            out.colunaTeste.patchSucesso = true;
            const item = await client.api('/sites/' + siteId + '/lists/' + listId + '/items/' + itemId + '?expand=fields').get();
            out.colunaTeste.valorGravado = (item.fields || {})[created.name];
          } catch (eP) {
            out.colunaTeste.patchErro = { msg: eP.message, status: eP.statusCode, body: eP.body };
          }
        }
        out.colunaTeste.aviso = 'COLUNA DE TESTE TEXT CRIADA. Apague manualmente no SP.';
      } catch (e) {
        out.colunaTeste = { erro: e.message, statusCode: e.statusCode, body: e.body, tentou: 'text-multiline' };
      }
    }

    // Se passou ?testarPatch=ID, tenta gravar URL de teste na NF e mostra o resultado
    if (req.query && req.query.testarPatch) {
      const itemId = req.query.testarPatch;
      out.testarPatch = { itemId };
      const testUrl = 'https://exemplo.com/teste-' + Date.now() + '.pdf';
      // Tenta 3 estrategias:
      // A) objeto Hyperlink {Url, Description} no campo displayName 'UrlPDFAprovado'
      try {
        await client.api('/sites/' + siteId + '/lists/' + listId + '/items/' + itemId + '/fields')
          .patch({ UrlPDFAprovado: { Url: testUrl, Description: 'teste' } });
        out.testarPatch.A_objetoNoDisplayName = 'sucesso';
      } catch (e) {
        out.testarPatch.A_objetoNoDisplayName = { erro: e.message, statusCode: e.statusCode, body: e.body };
      }
      // B) string simples no campo displayName
      try {
        await client.api('/sites/' + siteId + '/lists/' + listId + '/items/' + itemId + '/fields')
          .patch({ UrlPDFAprovado: testUrl });
        out.testarPatch.B_stringNoDisplayName = 'sucesso';
      } catch (e) {
        out.testarPatch.B_stringNoDisplayName = { erro: e.message, statusCode: e.statusCode, body: e.body };
      }
      // C) usando internal name (achar primeiro)
      const colInternal = (cols.value || []).find(c => c.displayName === 'UrlPDFAprovado');
      if (colInternal) {
        const patch = {};
        patch[colInternal.name] = { Url: testUrl, Description: 'teste internal' };
        try {
          await client.api('/sites/' + siteId + '/lists/' + listId + '/items/' + itemId + '/fields').patch(patch);
          out.testarPatch.C_objetoNoInternalName = { sucesso: true, internal: colInternal.name };
        } catch (e) {
          out.testarPatch.C_objetoNoInternalName = { erro: e.message, statusCode: e.statusCode, body: e.body, internal: colInternal.name };
        }
      } else {
        out.testarPatch.C_objetoNoInternalName = 'coluna UrlPDFAprovado nao tem displayName matching';
      }

      // Le item depois pra ver o que ficou
      try {
        const item = await client.api('/sites/' + siteId + '/lists/' + listId + '/items/' + itemId + '?expand=fields').get();
        const f = item.fields || {};
        out.testarPatch.itemDepois = {
          UrlPDFAprovado: f.UrlPDFAprovado,
          tipoUrlPDFAprovado: typeof f.UrlPDFAprovado
        };
        // Tambem pelos campos de colInternal
      } catch (e) { out.testarPatch.itemDepoisErro = e.message; }
    }
  } catch (e) {
    out.erro = e.message;
    out.stack = (e.stack || '').split('\n').slice(0, 6);
  }
  context.res = { status: 200, headers: { 'Content-Type': 'application/json' }, body: out };
};
