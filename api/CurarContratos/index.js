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

const EXCLUIR = /glosa|recurso de glosas/i;
const TEXTOS_PILOTO = '_curado_textos.json';
const MAX_TEXTO_JOIN = 40000; // curar corta em 32k; guardamos um pouco mais de margem

// Deriva a UF (SP/RJ/ES) do caminho/pasta da unidade. Usado como FALLBACK quando o
// conteudo do contrato nao trouxe o estado (a pasta e obrigatoria e confiavel na Pronep).
function _estadoDaPasta(caminho) {
  const m = /(?:^|[^a-z])(SP|RJ|ES)(?:[^a-z]|$)/i.exec(String(caminho || ''));
  return m ? m[1].toUpperCase() : null;
}

// Agrupa os chunks do indice por contrato e remonta o texto. filtro = regex opcional
// (ex.: operadoras especificas). Sem filtro = TODO o acervo Comercial indexado.
function _montarTextos(chunks, filtro) {
  const map = {};
  for (const c of (chunks || [])) {
    const hay = String(c.fornecedor || '') + ' ' + String(c.subpasta || '') + ' ' + String(c.contratoNome || '');
    if (EXCLUIR.test(hay)) continue;
    if (filtro && !filtro.test(hay)) continue;
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
    // Auth: secret compartilhado (runner do GitHub Actions, server-side) OU admin (navegador).
    const secret = process.env.CURADORIA_SECRET;
    const hdrSecret = (req.headers && (req.headers['x-curadoria-secret'] || req.headers['X-Curadoria-Secret'])) || '';
    const viaSecret = !!(secret && hdrSecret && hdrSecret === secret);
    if (!viaSecret) {
      const authz = await requireAdmin(context, req);
      if (!authz) return;
    }

    const offset = Math.max(0, parseInt((req.query && req.query.offset) || '0', 10) || 0);
    const limit = Math.min(8, Math.max(1, parseInt((req.query && req.query.limit) || '1', 10) || 1));
    const prep = req.query && (req.query.prep === '1' || req.query.prep === 'true');
    const modeloOverride = (req.query && req.query.model) || null; // 'sonnet'|'opus'|'haiku'
    // Filtro opcional por operadora(s): ?operadoras=amil,bradesco. Sem isso = TODO o Comercial.
    let filtro = null;
    const ops = (req.query && req.query.operadoras) || '';
    if (ops.trim()) { try { filtro = new RegExp(ops.split(',').map(function (s) { return s.trim(); }).filter(Boolean).join('|'), 'i'); } catch (e) { filtro = null; } }

    const client = getGraphClient();
    diag.step = 'resolve';
    const contr = await resolveContratosSite(client);
    const driveId = contr.driveId;

    // ===== FASE PREP: limpa curado/ + remonta e salva os textos do piloto (sem IA).
    if (prep) {
      diag.step = 'prep';
      await limparCurado(client, driveId);
      const chunks = await rag.carregarIndice(client, driveId, true);
      const files = _montarTextos(chunks, filtro);
      await salvarShard(client, driveId, TEXTOS_PILOTO, { total: files.length, files: files, criadoEm: new Date().toISOString() });
      context.res = { status: 200, headers: { 'Content-Type': 'application/json', 'Cache-Control': 'no-store' },
        body: { ok: true, prepared: true, total: files.length, next: 0, fonte: 'indice_rag' } };
      return;
    }

    // ===== MODO LISTA (runner background): devolve os textos preparados (meta + texto)
    // para o runner do GitHub Actions curar FORA do teto de 45s (Sonnet/Opus).
    if (req.query && (req.query.lista === '1' || req.query.lista === 'true')) {
      diag.step = 'lista';
      let bundle = await lerJson(client, driveId, TEXTOS_PILOTO);
      if (!bundle || !Array.isArray(bundle.files)) {
        const chunks = await rag.carregarIndice(client, driveId, false);
        bundle = { total: 0, files: _montarTextos(chunks, filtro) };
        bundle.total = bundle.files.length;
        await salvarShard(client, driveId, TEXTOS_PILOTO, bundle);
      }
      context.res = { status: 200, headers: { 'Content-Type': 'application/json', 'Cache-Control': 'no-store' },
        body: { ok: true, total: bundle.files.length, files: bundle.files } };
      return;
    }

    // ===== MODO SALVAR (runner background): recebe o canonico JA curado (a IA rodou no
    // runner, sem teto de tempo) e faz o pos-processo do app: skip glosa/outro, deriva UF
    // pela pasta, monta markdown, nomeia e faz upload. Rapido (<45s) — so Graph, sem IA.
    if (req.query && (req.query.salvar === '1' || req.query.salvar === 'true')) {
      diag.step = 'salvar';
      const body = req.body || {};
      const f = body.file || {};
      const canonico = body.canonico || null;
      if (!canonico) {
        context.res = { status: 400, headers: { 'Content-Type': 'application/json' }, body: { ok: false, motivo: 'sem_canonico' } };
        return;
      }
      if (canonico.doc_tipo && ['glosa', 'outro'].indexOf(canonico.doc_tipo) >= 0) {
        context.res = { status: 200, headers: { 'Content-Type': 'application/json' }, body: { ok: true, skipped: true, motivo: 'doc_tipo_' + canonico.doc_tipo } };
        return;
      }
      if (!canonico.estado_uf) {
        const ufPasta = _estadoDaPasta(String(f.subpasta || '') + ' ' + String(f.webUrl || '') + ' ' + String(f.nome || ''));
        if (ufPasta) {
          canonico.estado_uf = ufPasta;
          canonico.proveniencia = canonico.proveniencia || {};
          canonico.proveniencia.estado_origem = 'pasta';
        }
      }
      const md = montarMarkdown(canonico, { arquivoOrigem: f.nome, webUrl: f.webUrl });
      const nome = nomeCurado(canonico, f.fornecedor || f.nome, f.contratoId);
      await salvarCurado(client, driveId, nome, md);
      context.res = { status: 200, headers: { 'Content-Type': 'application/json' }, body: { ok: true, curado: nome } };
      return;
    }

    // ===== Processamento: le SO o arquivo de textos (nao recarrega o indice).
    diag.step = 'ler_textos';
    let bundle = await lerJson(client, driveId, TEXTOS_PILOTO);
    if (!bundle || !Array.isArray(bundle.files)) {
      // Sem prep previo: monta na hora (uma vez) e segue.
      const chunks = await rag.carregarIndice(client, driveId, false);
      bundle = { total: 0, files: _montarTextos(chunks, filtro) };
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
          const canon = await curar(f.texto, 'nativo', { model: modeloOverride });
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
        const canonico = await curar(texto, 'nativo', { model: modeloOverride });
        if (!canonico) return { nome: f.nome, motivo: 'curadoria_falhou', erro: ultimoErroCuradoria() };
        if (canonico.doc_tipo && ['glosa', 'outro'].indexOf(canonico.doc_tipo) >= 0) return { nome: f.nome, motivo: 'doc_tipo_' + canonico.doc_tipo };
        // UF: conteudo primeiro; se omisso, deriva da PASTA da unidade (regra Pronep:
        // contrato mora em SP/RJ/ES). Registra a origem na proveniencia.
        if (!canonico.estado_uf) {
          const ufPasta = _estadoDaPasta(String(f.subpasta || '') + ' ' + String(f.webUrl || '') + ' ' + String(f.nome || ''));
          if (ufPasta) {
            canonico.estado_uf = ufPasta;
            canonico.proveniencia = canonico.proveniencia || {};
            canonico.proveniencia.estado_origem = 'pasta';
          }
        }
        const md = montarMarkdown(canonico, { arquivoOrigem: f.nome, webUrl: f.webUrl });
        const nome = nomeCurado(canonico, f.fornecedor || f.nome, f.contratoId);
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
