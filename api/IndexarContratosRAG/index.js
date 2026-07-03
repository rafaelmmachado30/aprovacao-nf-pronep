/**
 * /api/IndexarContratosRAG (GET|POST) — ADMIN ONLY. Fase 1 da base de conhecimento.
 *
 * Indexa os contratos COMERCIAIS em LOTES (o front chama em loop, driblando o timeout
 * de 4min do SWA). Enumera pela LISTA PRONEP-NF-Contratos (Diretoria='Comercial' +
 * DriveItemId), baixa o arquivo do drive dos contratos, extrai o texto COMPLETO,
 * quebra em trechos, gera embeddings e grava shards JSON em _RAG/comercial/.
 *
 * OBS: "Comercial" NAO e uma pasta literal — e uma diretoria derivada do caminho
 * (classificarPath). Por isso enumeramos pela lista ja sincronizada, nao por pasta.
 *
 * Query: ?offset=0&limit=10   (offset=0 recria o manifest de arquivos)
 * Resposta: { ok, total, offset, processado, arquivosNoBatch, chunksNoBatch, done, next }
 */

require('isomorphic-fetch');
const { requireAdmin } = require('../shared/authz');
const {
  getGraphClient, garantirListaContratos, getContratoColMap,
  resolveContratosSite, extrairTexto
} = require('../shared/contratos');
const { embed, chunkTexto, vecToB64, salvarShard, lerJson, RAG_FOLDER } = require('../shared/ragContratos');

// Subpastas/documentos que NAO sao contratos (poluem o RAG). Excluidos do indice.
const EXCLUIR_SUBPASTA = /glosa|recurso de glosas/i;

// Apaga os shards antigos (part-*.json) antes de reindexar do zero (offset 0).
async function _limparShards(client, driveId) {
  try {
    const enc = encodeURIComponent(RAG_FOLDER).replace(/%2F/g, '/');
    const resp = await client.api('/drives/' + driveId + '/root:/' + enc + ':/children').select('id,name').top(999).get();
    for (const x of (resp.value || [])) {
      if (/^part-.*\.json$/i.test(x.name || '')) {
        try { await client.api('/drives/' + driveId + '/items/' + x.id).delete(); } catch (e) { /* ignora */ }
      }
    }
  } catch (e) { /* pasta ainda nao existe */ }
}

const EXTS_OK = ['pdf', 'docx'];
function _norm(s) {
  return String(s || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '').trim();
}
function _ext(nome) {
  const m = /\.([a-z0-9]+)$/i.exec(String(nome || ''));
  return m ? m[1].toLowerCase() : '';
}

module.exports = async function (context, req) {
  const diag = { step: 'start' };
  try {
    const authz = await requireAdmin(context, req);
    if (!authz) return;

    const offset = Math.max(0, parseInt((req.query && req.query.offset) || '0', 10) || 0);
    const limit = Math.min(20, Math.max(1, parseInt((req.query && req.query.limit) || '10', 10) || 10));

    const client = getGraphClient();
    diag.step = 'resolve';
    const listInfo = await garantirListaContratos(client);       // { siteId (app), listId }
    const appSiteId = listInfo.siteId;
    const listId = listInfo.listId;
    const colMap = await getContratoColMap(client, appSiteId, listId);
    const contr = await resolveContratosSite(client);            // { driveId } dos arquivos
    const driveId = contr.driveId;
    const col = function (d) { return colMap[d] || d; };

    // Manifest de arquivos: no offset 0 (re)cria enumerando a lista (Comercial + DriveItemId).
    diag.step = 'manifest';
    // No inicio (offset 0), limpa os shards antigos pra reindexar do zero (evita
    // sobrar chunk de indexacao anterior, ex.: os "Recurso de Glosas" ja indexados).
    if (offset === 0) await _limparShards(client, driveId);

    let manifest = offset === 0 ? null : await lerJson(client, driveId, '_files.json');
    if (!manifest || !Array.isArray(manifest.files)) {
      const files = [];
      let url = '/sites/' + appSiteId + '/lists/' + listId + '/items?expand=fields&$top=500';
      let pages = 0;
      while (url && pages < 40) {
        const resp = await client.api(url).get();
        for (const it of (resp.value || [])) {
          const f = it.fields || {};
          const diretoria = _norm(f[col('Diretoria')]);
          if (diretoria !== 'comercial') continue;
          const driveItemId = f[col('DriveItemId')] || '';
          if (!driveItemId) continue;
          const nome = f[col('NomeArquivo')] || f[col('Title')] || '';
          const ext = _ext(nome);
          if (EXTS_OK.indexOf(ext) < 0) continue;
          // Exclui documentos que nao sao contratos (Recurso de Glosas etc.)
          const caminho = String(f[col('PathRelativoSP')] || '') + ' ' + String(f[col('CaminhoSharepoint')] || '') + ' ' + String(nome);
          if (EXCLUIR_SUBPASTA.test(caminho)) continue;
          files.push({
            listItemId: it.id,
            driveItemId: driveItemId,
            nome: nome,
            ext: ext,
            fornecedor: f[col('Fornecedor')] || '',
            subpasta: f[col('PathRelativoSP')] || f[col('CaminhoSharepoint')] || '',
            webUrl: f[col('CaminhoSharepoint')] || ''
          });
        }
        pages++;
        url = resp['@odata.nextLink'] ? resp['@odata.nextLink'].replace('https://graph.microsoft.com/v1.0', '') : null;
      }
      manifest = { total: files.length, files: files, criadoEm: new Date().toISOString() };
      await salvarShard(client, driveId, '_files.json', manifest);
    }

    const total = manifest.files.length;
    const slice = manifest.files.slice(offset, offset + limit);

    diag.step = 'process_batch';
    const chunksOut = [];
    const arquivosOk = [];
    for (const f of slice) {
      try {
        const texto = await extrairTexto(client, driveId, f.driveItemId, f.ext, {});
        if (!texto || texto.length < 40) continue; // imagem/scan sem texto
        const pedacos = chunkTexto(texto);
        if (!pedacos.length) continue;
        const vecs = await embed(pedacos);
        for (let i = 0; i < pedacos.length; i++) {
          chunksOut.push({
            contratoId: f.listItemId, contratoNome: f.nome,
            diretoria: 'Comercial', subpasta: f.subpasta, fornecedor: f.fornecedor,
            webUrl: f.webUrl, chunkIdx: i, texto: pedacos[i], vec: vecToB64(vecs[i])
          });
        }
        arquivosOk.push(f.nome);
      } catch (e) {
        diag.ultimoErro = { arquivo: f.nome, erro: (e && e.message) || String(e) };
      }
    }

    diag.step = 'save_shard';
    const shardNome = 'part-' + String(offset).padStart(5, '0') + '.json';
    await salvarShard(client, driveId, shardNome, { offset: offset, gerado: new Date().toISOString(), chunks: chunksOut });

    const processado = Math.min(offset + limit, total);
    context.res = {
      status: 200, headers: { 'Content-Type': 'application/json', 'Cache-Control': 'no-store' },
      body: {
        ok: true, total: total, offset: offset, processado: processado,
        arquivosNoBatch: arquivosOk.length, chunksNoBatch: chunksOut.length,
        done: processado >= total, next: processado < total ? processado : null,
        ultimoErro: diag.ultimoErro || null
      }
    };
  } catch (err) {
    context.log && context.log.error && context.log.error('IndexarContratosRAG:', err);
    context.res = { status: 500, headers: { 'Content-Type': 'application/json' },
      body: { error: (err && err.message) || String(err), diag: diag } };
  }
};
