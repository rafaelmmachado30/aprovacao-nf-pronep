/**
 * /api/CurarContratos (GET|POST) — ADMIN ONLY. Fase 1 da NOVA arquitetura (ingestao curada).
 *
 * Cura as operadoras PILOTO (Amil, Bradesco, CABESP) a partir do TEXTO JA EXTRAIDO no
 * indice RAG (_RAG/comercial/part-*.json). Reusar o indice evita os 404 de DriveItemId
 * desatualizado e os PDFs sem texto (que nem entram no indice).
 *
 * DESEMPENHO: a fase prep remonta o texto do piloto e SALVA num unico arquivo
 * (_curado_textos.json). Cada request de processamento so LE esse arquivo (pequeno) e
 * cura 1 documento — nunca recarrega o indice de ~65 shards (o que estourava o timeout).
 *
 * Fluxo: prep=1 (limpa curado/ + monta textos, sem IA) -> offset/limit=1 em loop.
 * Query: ?prep=1 | ?offset=0&limit=1 | ?debug=1&limit=1
 */

require('isomorphic-fetch');
const { requireAdmin } = require('../shared/authz');
const { getGraphClient, resolveContratosSite } = require('../shared/contratos');
const rag = require('../shared/ragContratos');
const { salvarShard, lerJson } = rag;
const { curar, ultimoErroCuradoria, MODEL } = require('../shared/curadoriaContrato');
const { nomeCurado, montarMarkdown, salvarCurado, limparCurado } = require('../shared/fonteUnica');

const PILOTO = /amil|bradesco|cabesp/i;
const EXCLUIR = /glosa|recurso de glosas/i;
const TEXTOS_PILOTO = '_curado_textos.json';
const MAX_TEXTO_JOIN = 40000; // curar corta em 32k; guardamos um pouco mais de margem

// Agrupa os chunks do indice por contrato (so os do piloto) e remonta o texto.
function _montarTextos(chunks) {
  const map = {};
  for (const c of (chunks || [])) {
    const hay = String(c.fornecedor || '') + ' ' + String(c.subpasta || '') + ' ' + String(c.contratoNome || '');
    if (!PILOTO.test(hay) || EXCLUIR.test(hay)) continue;
    const id = c.contratoId;
    if (!id) continue;
    if (!map[id]) map[id] = { contratoId: id, nome: c.contratoNome || '', fornecedor: c.fornecedor || '', subpasta: c.subpasta || '', webUrl: c.webUrl || '', _chunks: [] };
    map[id]._chunks.push({ idx: (c.chunkIdx != null ? c.chunkIdx : 0), texto: c.texto || '' });
  }
  return Object.keys(map).map(function (id) {
    const e = map[id];
    const cs = e._chunks.sort(function (a, b) { return a.idx - b.idx; });
    let t = cs.map(function (x) { return x.texto; }).join('\n');
    if (t.length > MAX_TEXTO_JOIN) t = t.slice(0, MAX_TEXTO_JOIN);
    return { contratoId: e.contratoId, nome: e.nome, fornecedor: e.fornecedor, subpasta: e.subpasta, webUrl: e.webUrl, texto: t };
  });
}

module.exports = async function (context, req) {
  const diag = { step: 'start', modelo: MODEL };
  try {
    const authz = await requireAdmin(context, req);
    if (!authz) return;

    const offset = Math.max(0, parseInt((req.query && req.query.offset) || '0', 10) || 0);
    const limit = Math.min(8, Math.max(1, parseInt((req.query && req.query.limit) || '1', 10) || 1));
    const prep = req.query && (req.query.prep === '1' || req.query.prep === 'true');

    const client = getGraphClient();
    diag.step = 'resolve';
    const contr = await resolveContratosSite(client);
    const driveId = contr.driveId;

    // ===== FASE PREP: limpa curado/ + remonta e salva os textos do piloto (sem IA).
    if (prep) {
      diag.step = 'prep';
      await limparCurado(client, driveId);
      const chunks = await rag.carregarIndice(client, driveId, true);
      const files = _montarTextos(chunks);
      await salvarShard(client, driveId, TEXTOS_PILOTO, { total: files.length, files: files, criadoEm: new Date().toISOString() });
      context.res = { status: 200, headers: { 'Content-Type': 'application/json', 'Cache-Control': 'no-store' },
        body: { ok: true, prepared: true, total: files.length, next: 0, fonte: 'indice_rag' } };
      return;
    }

    // ===== Processamento: le SO o arquivo de textos (nao recarrega o indice).
    diag.step = 'ler_textos';
    let bundle = await lerJson(client, driveId, TEXTOS_PILOTO);
    if (!bundle || !Array.isArray(bundle.files)) {
      // Sem prep previo: monta na hora (uma vez) e segue.
      const chunks = await rag.carregarIndice(client, driveId, false);
      bundle = { total: 0, files: _montarTextos(chunks) };
      bundle.total = bundle.files.length;
      await salvarShard(client, driveId, TEXTOS_PILOTO, bundle);
    }
    const total = bundle.files.length;
    const slice = bundle.files.slice(offset, offset + limit);

    // ===== DEBUG (nao-destrutivo).
    if (req.query && (req.query.debug === '1' || req.query.debug === 'true')) {
      const amostra = [];
      for (const f of slice.slice(0, Math.min(limit, 2))) {
        const info = { nome: f.nome, fornecedor: f.fornecedor, textoLen: (f.texto || '').length };
        try {
          const canon = await curar(f.texto, 'nativo');
          info.curadoOk = !!canon;
          info.curadoErro = canon ? null : ultimoErroCuradoria();
          if (canon) info.amostra = { doc_tipo: canon.doc_tipo, operadora: canon.operadora && canon.operadora.nome, estado: canon.estado_uf, diarias: (canon.diarias || []).length, clausulas: (canon.clausulas || []).length };
        } catch (e) { info.erro = (e && e.message) || String(e); }
        amostra.push(info);
      }
      context.res = { status: 200, headers: { 'Content-Type': 'application/json', 'Cache-Control': 'no-store' },
        body: { ok: true, debug: true, modelo: MODEL, totalManifest: total, amostra: amostra } };
      return;
    }

    // ===== Cura o lote (1 por request).
    diag.step = 'process_batch';
    const curadosOut = [];
    const ignorados = [];
    const resultados = await Promise.all(slice.map(async function (f) {
      try {
        const texto = f.texto || '';
        if (texto.length < 60) return { nome: f.nome, motivo: 'sem_texto_indice' };
        const canonico = await curar(texto, 'nativo');
        if (!canonico) return { nome: f.nome, motivo: 'curadoria_falhou', erro: ultimoErroCuradoria() };
        if (canonico.doc_tipo && ['glosa', 'outro'].indexOf(canonico.doc_tipo) >= 0) return { nome: f.nome, motivo: 'doc_tipo_' + canonico.doc_tipo };
        const md = montarMarkdown(canonico, { arquivoOrigem: f.nome, webUrl: f.webUrl });
        const nome = nomeCurado(canonico, f.fornecedor || f.nome);
        await salvarCurado(client, driveId, nome, md);
        return { nome: f.nome, curado: nome, ok: true };
      } catch (e) {
        diag.ultimoErro = { arquivo: f.nome, erro: (e && e.message) || String(e) };
        return { nome: f.nome, motivo: 'erro', erro: (e && e.message) || String(e) };
      }
    }));
    for (const r of resultados) {
      if (r && r.ok) curadosOut.push(r.curado);
      else if (r) ignorados.push({ nome: r.nome, motivo: r.motivo, erro: r.erro });
    }

    const processado = Math.min(offset + limit, total);
    context.res = { status: 200, headers: { 'Content-Type': 'application/json', 'Cache-Control': 'no-store' },
      body: {
        ok: true, modelo: MODEL, total: total, offset: offset, processado: processado,
        arquivosNoBatch: slice.length, curadosNoBatch: curadosOut.length,
        curados: curadosOut, ignorados: ignorados,
        done: processado >= total, next: processado < total ? processado : null,
        ultimoErro: diag.ultimoErro || null
      } };
  } catch (err) {
    context.log && context.log.error && context.log.error('CurarContratos:', err);
    context.res = { status: 500, headers: { 'Content-Type': 'application/json' },
      body: { error: (err && err.message) || String(err), diag: diag } };
  }
};
