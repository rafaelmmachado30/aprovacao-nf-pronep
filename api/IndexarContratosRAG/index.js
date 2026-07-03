/**
 * /api/IndexarContratosRAG (GET|POST) — ADMIN ONLY. Fase 1 da base de conhecimento.
 *
 * Indexa a pasta Comercial em LOTES (o front chama em loop, driblando o timeout de
 * 4min do SWA). Cada lote: pega os proximos N contratos, extrai texto, quebra em
 * trechos, gera embeddings e grava um shard JSON em _RAG/comercial/part-<offset>.json.
 *
 * Query: ?offset=0&limit=10   (offset=0 recria o manifest de arquivos)
 * Resposta: { ok, total, offset, processado, chunksNoBatch, done, next }
 *
 * Idempotente: rodar de novo com o mesmo offset sobrescreve o mesmo shard.
 */

require('isomorphic-fetch');
const { requireAdmin } = require('../shared/authz');
const {
  getGraphClient, resolveContratosSite, crawlPasta, extrairTexto,
  eRelevantePraContrato, ROOT_FOLDER_PATH
} = require('../shared/contratos');
const {
  embed, chunkTexto, vecToB64, salvarShard, lerJson
} = require('../shared/ragContratos');

const COMERCIAL_PATH = ROOT_FOLDER_PATH + '/Comercial';
const EXTS_OK = ['pdf', 'docx', 'doc'];

module.exports = async function (context, req) {
  const diag = { step: 'start' };
  try {
    const authz = await requireAdmin(context, req);
    if (!authz) return;

    const offset = Math.max(0, parseInt((req.query && req.query.offset) || '0', 10) || 0);
    const limit = Math.min(20, Math.max(1, parseInt((req.query && req.query.limit) || '10', 10) || 10));

    const client = getGraphClient();
    diag.step = 'resolve_site';
    const site = await resolveContratosSite(client);
    const driveId = site.driveId;

    // Manifest de arquivos: no offset 0 (re)cria varrendo a pasta Comercial.
    diag.step = 'manifest';
    let manifest = offset === 0 ? null : await lerJson(client, driveId, '_files.json');
    if (!manifest || !Array.isArray(manifest.files)) {
      const crawl = await crawlPasta(client, driveId, COMERCIAL_PATH, { maxArquivos: 5000 });
      const files = crawl
        .filter(function (f) {
          const ext = String(f.ext || '').toLowerCase();
          return EXTS_OK.indexOf(ext) >= 0 && eRelevantePraContrato(f.nome).relevante;
        })
        .map(function (f) {
          return { id: f.id, nome: f.nome, ext: String(f.ext || '').toLowerCase(), ancestors: f.ancestors || [], webUrl: f.webUrl || '' };
        });
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
        const texto = await extrairTexto(client, driveId, f.id, f.ext, {});
        if (!texto || texto.length < 40) continue; // provavelmente imagem/scan sem texto
        const pedacos = chunkTexto(texto);
        if (!pedacos.length) continue;
        const vecs = await embed(pedacos);
        const anc = f.ancestors || [];
        const subpasta = anc.join(' / ');
        const fornecedor = anc.length ? anc[anc.length - 1] : '';
        for (let i = 0; i < pedacos.length; i++) {
          chunksOut.push({
            contratoId: f.id, contratoNome: f.nome,
            diretoria: 'Comercial', subpasta: subpasta, fornecedor: fornecedor,
            webUrl: f.webUrl, chunkIdx: i, texto: pedacos[i], vec: vecToB64(vecs[i])
          });
        }
        arquivosOk.push(f.nome);
      } catch (e) {
        // pula arquivo problematico (protegido/scan/corrompido) sem parar o lote
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
