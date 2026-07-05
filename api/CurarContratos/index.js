/**
 * /api/CurarContratos (GET|POST) — ADMIN ONLY. Fase 1 da NOVA arquitetura (ingestao curada).
 *
 * Cura as operadoras PILOTO (Amil, Bradesco, CABESP) a partir do TEXTO JA EXTRAIDO no
 * indice RAG (_RAG/comercial/part-*.json). Reusar o indice evita os 404 de DriveItemId
 * desatualizado e os PDFs sem texto (que nem entram no indice). Para cada contrato,
 * remonta o texto dos seus chunks, extrai o modelo CANONICO (curadoriaContrato) e grava
 * a FONTE UNICA curada (Markdown+YAML) em _RAG/curado/. Dessa fonte os trilhos
 * (relacional/vetorial/grafo) sao derivados depois — sem perda e regeneravel.
 *
 * Roda em LOTES (o front chama em loop). prep=1 limpa curado/ + monta o manifesto (sem IA).
 *
 * Query: ?prep=1 | ?offset=0&limit=1 | ?debug=1&limit=1 (nao-destrutivo)
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
const MANIFEST_PILOTO = '_curado_files.json';
const MAX_TEXTO_JOIN = 60000; // texto remontado por contrato (curar depois corta em 32k)

// Agrupa os chunks do indice por contrato, so os que casam com o piloto.
function _agruparPiloto(chunks) {
  const map = {};
  for (const c of (chunks || [])) {
    const hay = String(c.fornecedor || '') + ' ' + String(c.subpasta || '') + ' ' + String(c.contratoNome || '');
    if (!PILOTO.test(hay)) continue;
    if (EXCLUIR.test(hay)) continue;
    const id = c.contratoId;
    if (!id) continue;
    if (!map[id]) map[id] = { contratoId: id, nome: c.contratoNome || '', fornecedor: c.fornecedor || '', subpasta: c.subpasta || '', webUrl: c.webUrl || '', chunks: [] };
    map[id].chunks.push({ idx: (c.chunkIdx != null ? c.chunkIdx : 0), texto: c.texto || '' });
  }
  return map;
}
function _textoDoContrato(entry) {
  const cs = (entry.chunks || []).slice().sort(function (a, b) { return a.idx - b.idx; });
  let t = cs.map(function (x) { return x.texto; }).join('\n');
  if (t.length > MAX_TEXTO_JOIN) t = t.slice(0, MAX_TEXTO_JOIN);
  return t;
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

    // ===== FASE PREP: limpa curado/ + monta manifesto a partir do indice RAG.
    diag.step = 'manifest';
    if (prep) {
      await limparCurado(client, driveId);
      const chunks = await rag.carregarIndice(client, driveId, true);
      const map = _agruparPiloto(chunks);
      const files = Object.keys(map).map(function (id) {
        const e = map[id];
        return { contratoId: e.contratoId, nome: e.nome, fornecedor: e.fornecedor, subpasta: e.subpasta, webUrl: e.webUrl };
      });
      const man = { total: files.length, files: files, criadoEm: new Date().toISOString() };
      await salvarShard(client, driveId, MANIFEST_PILOTO, man);
      context.res = { status: 200, headers: { 'Content-Type': 'application/json', 'Cache-Control': 'no-store' },
        body: { ok: true, prepared: true, total: man.total, next: 0, fonte: 'indice_rag' } };
      return;
    }

    // Manifesto (monta se faltar). Indice carregado (cache) p/ remontar o texto.
    let manifest = await lerJson(client, driveId, MANIFEST_PILOTO);
    const chunks = await rag.carregarIndice(client, driveId, false);
    const map = _agruparPiloto(chunks);
    if (!manifest || !Array.isArray(manifest.files)) {
      const files = Object.keys(map).map(function (id) {
        const e = map[id];
        return { contratoId: e.contratoId, nome: e.nome, fornecedor: e.fornecedor, subpasta: e.subpasta, webUrl: e.webUrl };
      });
      manifest = { total: files.length, files: files };
    }

    const total = manifest.files.length;
    const slice = manifest.files.slice(offset, offset + limit);

    // ===== DEBUG (nao-destrutivo).
    if (req.query && (req.query.debug === '1' || req.query.debug === 'true')) {
      const amostra = [];
      for (const f of slice.slice(0, Math.min(limit, 2))) {
        const info = { nome: f.nome, fornecedor: f.fornecedor };
        try {
          const e = map[f.contratoId];
          const texto = e ? _textoDoContrato(e) : '';
          info.textoLen = texto.length;
          const canon = await curar(texto, 'nativo');
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

    // ===== Processa o lote.
    diag.step = 'process_batch';
    const curadosOut = [];
    const ignorados = [];
    const resultados = await Promise.all(slice.map(async function (f) {
      try {
        const e = map[f.contratoId];
        const texto = e ? _textoDoContrato(e) : '';
        if (!texto || texto.length < 60) return { nome: f.nome, motivo: 'sem_texto_indice' };
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
