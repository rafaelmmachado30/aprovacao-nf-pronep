/**
 * shared/ragContratos.js — Base de conhecimento (RAG) dos contratos.
 *
 * Fase 1: indexa a pasta Comercial. Extrai texto (reusa contratos.js), quebra em
 * trechos, gera embeddings (OpenAI text-embedding-3-small, dimensoes reduzidas p/
 * caber num indice leve) e guarda shards JSON no drive dos CONTRATOS, em _RAG/comercial/.
 *
 * Consulta (Fase 2, via SAN): emb-eda a pergunta, faz similaridade de cosseno sobre
 * os chunks (carregados e cacheados em memoria) e devolve os trechos mais relevantes
 * com a origem (contrato) pra citacao. Respeita o escopo de diretorias do usuario.
 *
 * SEM nova infra: armazenamento em arquivos no SharePoint + busca brute-force na Function.
 */

require('isomorphic-fetch');

const EMBED_MODEL = process.env.RAG_EMBED_MODEL || 'text-embedding-3-small';
const EMBED_DIMS = Number(process.env.RAG_EMBED_DIMS || 512);
const RAG_FOLDER = '_RAG/comercial'; // relativo a raiz do drive dos contratos

let _openai = null;
function getOpenAI() {
  if (_openai) return _openai;
  const key = process.env.OPENAI_API_KEY;
  if (!key) throw new Error('OPENAI_API_KEY nao configurada (necessaria para embeddings do RAG)');
  const OpenAI = require('openai');
  const Ctor = OpenAI.default || OpenAI.OpenAI || OpenAI;
  _openai = new Ctor({ apiKey: key });
  return _openai;
}

// Gera embeddings p/ um array de textos. Retorna array de number[].
async function embed(texts, dims) {
  const cli = getOpenAI();
  const resp = await cli.embeddings.create({
    model: EMBED_MODEL,
    input: texts,
    dimensions: dims || EMBED_DIMS
  });
  return (resp.data || []).map(function (d) { return d.embedding; });
}

// Quebra o texto em trechos com sobreposicao (sem cortar palavra quando possivel).
function chunkTexto(texto, maxChars, overlap) {
  maxChars = maxChars || 2200;
  overlap = overlap || 250;
  const t = String(texto || '').replace(/\s+/g, ' ').trim();
  if (!t) return [];
  if (t.length <= maxChars) return [t];
  const chunks = [];
  let i = 0;
  while (i < t.length) {
    let end = Math.min(i + maxChars, t.length);
    if (end < t.length) {
      const sp = t.lastIndexOf(' ', end);
      if (sp > i + maxChars * 0.6) end = sp;
    }
    const piece = t.slice(i, end).trim();
    if (piece) chunks.push(piece);
    if (end >= t.length) break;
    i = end - overlap;
    if (i < 0) i = 0;
  }
  return chunks;
}

// Vetores como base64 de Float32 (compacto no JSON).
function vecToB64(vec) {
  const f = new Float32Array(vec);
  return Buffer.from(f.buffer, f.byteOffset, f.byteLength).toString('base64');
}
function b64ToVec(b64) {
  const buf = Buffer.from(b64, 'base64');
  // copia p/ ArrayBuffer alinhado (evita erro de offset no Float32Array)
  const ab = buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength);
  return new Float32Array(ab);
}
function cosine(a, b) {
  let dot = 0, na = 0, nb = 0;
  const n = Math.min(a.length, b.length);
  for (let i = 0; i < n; i++) { dot += a[i] * b[i]; na += a[i] * a[i]; nb += b[i] * b[i]; }
  if (!na || !nb) return 0;
  return dot / (Math.sqrt(na) * Math.sqrt(nb));
}

// ---- Armazenamento (drive dos contratos) ----
function _encPath(p) { return encodeURIComponent(p).replace(/%2F/g, '/'); }

async function salvarShard(client, driveId, nome, obj) {
  const up = '/drives/' + driveId + '/root:/' + _encPath(RAG_FOLDER + '/' + nome) + ':/content';
  const body = Buffer.from(JSON.stringify(obj));
  return await client.api(up).header('Content-Type', 'application/json').put(body);
}

async function lerJson(client, driveId, nome) {
  try {
    const meta = await client.api('/drives/' + driveId + '/root:/' + _encPath(RAG_FOLDER + '/' + nome))
      .select('id,@microsoft.graph.downloadUrl').get();
    const url = meta['@microsoft.graph.downloadUrl'];
    if (url) { const r = await fetch(url); if (r.ok) return await r.json(); }
    const ab = await client.api('/drives/' + driveId + '/items/' + meta.id + '/content').responseType('arraybuffer').get();
    return JSON.parse(Buffer.from(ab).toString('utf-8'));
  } catch (e) { return null; }
}

async function listarShards(client, driveId) {
  try {
    const resp = await client.api('/drives/' + driveId + '/root:/' + _encPath(RAG_FOLDER) + ':/children')
      .select('name').top(999).get();
    return (resp.value || []).map(function (x) { return x.name; })
      .filter(function (n) { return /^part-.*\.json$/i.test(n); });
  } catch (e) { return []; }
}

// Carrega e cacheia (memoria do processo) todos os chunks p/ busca.
const _idxCache = { ts: 0, chunks: null };
const _IDX_TTL = 10 * 60 * 1000;
async function carregarIndice(client, driveId, force) {
  if (!force && _idxCache.chunks && (Date.now() - _idxCache.ts) < _IDX_TTL) return _idxCache.chunks;
  const shards = await listarShards(client, driveId);
  const all = [];
  for (const s of shards) {
    const obj = await lerJson(client, driveId, s);
    if (obj && Array.isArray(obj.chunks)) all.push.apply(all, obj.chunks);
  }
  _idxCache.chunks = all; _idxCache.ts = Date.now();
  return all;
}

// Busca semantica. diretoriasPermitidas = array de diretorias (lowercase) OU null (todas).
async function buscar(client, driveId, pergunta, opts) {
  opts = opts || {};
  const topK = opts.topK || 8;
  const permit = opts.diretoriasPermitidas || null;
  const chunks = await carregarIndice(client, driveId, false);
  if (!chunks.length) return [];
  const qvec = (await embed([pergunta]))[0];
  const scored = [];
  for (const c of chunks) {
    if (permit && permit.indexOf(String(c.diretoria || '').toLowerCase()) < 0) continue;
    const v = b64ToVec(c.vec);
    scored.push({ score: cosine(qvec, v), c: c });
  }
  scored.sort(function (a, b) { return b.score - a.score; });
  return scored.slice(0, topK).map(function (s) {
    return {
      score: s.score,
      contratoId: s.c.contratoId, contratoNome: s.c.contratoNome,
      diretoria: s.c.diretoria, subpasta: s.c.subpasta, fornecedor: s.c.fornecedor,
      webUrl: s.c.webUrl, texto: s.c.texto
    };
  });
}

module.exports = {
  EMBED_MODEL, EMBED_DIMS, RAG_FOLDER,
  getOpenAI, embed, chunkTexto, vecToB64, b64ToVec, cosine,
  salvarShard, lerJson, listarShards, carregarIndice, buscar
};
