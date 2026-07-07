/**
 * /api/CarregarBaseCurada (GET|POST) — ADMIN ONLY. Trilho RELACIONAL (Supabase).
 *
 * Le a fonte unica curada (_RAG/curado/*.md), extrai o canonico e popula as tabelas
 * do Postgres/Supabase. O relacional e DERIVADO — regeneravel do zero a qualquer hora.
 *
 * Roda em LOTES (o front chama em loop) p/ nao estourar o timeout de 45s do SWA:
 *  - prep=1: TRUNCATE CASCADE (reload limpo) + conta as fontes. Sem inserts.
 *  - offset/limit: le as fontes do lote e insere.
 *
 * Query: ?ping=1 (testa conexao) | ?prep=1 | ?offset=0&limit=10
 */

require('isomorphic-fetch');
const { requireAdmin } = require('../shared/authz');
const { getGraphClient, resolveContratosSite } = require('../shared/contratos');
const pg = require('../shared/pgContratos');
const loader = require('../shared/loaderCurado');
const { CURADO_FOLDER } = require('../shared/fonteUnica');

function _encPath(p) { return encodeURIComponent(p).replace(/%2F/g, '/'); }
function _canonicoDeMd(md) {
  try { const m = /```json\s*([\s\S]*?)```/i.exec(String(md || '')); if (!m) return null; return JSON.parse(m[1]); }
  catch (e) { return null; }
}

// Lista os .md curados COM downloadUrl, ordenados por nome (offset estavel entre lotes).
async function _listarCurados(client, driveId) {
  const resp = await client.api('/drives/' + driveId + '/root:/' + _encPath(CURADO_FOLDER) + ':/children')
    .select('id,name,@microsoft.graph.downloadUrl').top(999).get();
  return (resp.value || [])
    .filter(function (x) { return /\.md$/i.test(x.name || ''); })
    .map(function (x) { return { name: x.name, id: x.id, url: x['@microsoft.graph.downloadUrl'] || null }; })
    .sort(function (a, b) { return a.name < b.name ? -1 : (a.name > b.name ? 1 : 0); });
}

module.exports = async function (context, req) {
  try {
    const authz = await requireAdmin(context, req);
    if (!authz) return;

    if (req.query && (req.query.ping === '1' || req.query.ping === 'true')) {
      const p = await pg.ping();
      context.res = { status: 200, headers: { 'Content-Type': 'application/json', 'Cache-Control': 'no-store' }, body: p };
      return;
    }
    if (!pg.getPool()) {
      context.res = { status: 400, headers: { 'Content-Type': 'application/json' },
        body: { ok: false, error: 'SUPABASE_DB_URL nao configurada no Function App.' } };
      return;
    }

    const offset = Math.max(0, parseInt((req.query && req.query.offset) || '0', 10) || 0);
    const limit = Math.min(25, Math.max(1, parseInt((req.query && req.query.limit) || '10', 10) || 10));
    const prep = req.query && (req.query.prep === '1' || req.query.prep === 'true');

    const client = getGraphClient();
    const site = await resolveContratosSite(client);
    const driveId = site.driveId;

    // ===== PREP: TRUNCATE (reload limpo) + conta fontes.
    if (prep) {
      const conn = await pg.getPool().connect();
      try { await loader.limparTudo(conn); } finally { conn.release(); }
      const files = await _listarCurados(client, driveId);
      context.res = { status: 200, headers: { 'Content-Type': 'application/json', 'Cache-Control': 'no-store' },
        body: { ok: true, prepared: true, total: files.length, next: 0 } };
      return;
    }

    // ===== Lote: le as fontes do slice e insere.
    const files = await _listarCurados(client, driveId);
    const total = files.length;
    if (!total) {
      context.res = { status: 200, headers: { 'Content-Type': 'application/json' },
        body: { ok: true, total: 0, done: true, aviso: 'Nenhuma fonte curada encontrada. Rode "Curar contratos (IA)" antes.' } };
      return;
    }
    const slice = files.slice(offset, offset + limit);
    const conn = await pg.getPool().connect();
    const carregados = [];
    const pulados = [];
    try {
      for (const f of slice) {
        try {
          let txt = null;
          if (f.url) { const r = await fetch(f.url); if (r.ok) txt = await r.text(); }
          if (txt == null) {
            const ab = await client.api('/drives/' + driveId + '/items/' + f.id + '/content').responseType('arraybuffer').get();
            txt = Buffer.from(ab).toString('utf-8');
          }
          const canon = _canonicoDeMd(txt);
          if (!canon) { pulados.push({ nome: f.name, motivo: 'md_sem_canonico' }); continue; }
          canon._arquivo = f.name;
          const r = await loader.carregar(conn, canon, { arquivoOrigem: f.name });
          if (r && r.ok) carregados.push({ operadora: r.operadora, estado: r.estado, id: r.contratoId });
          else pulados.push({ nome: f.name, motivo: (r && r.motivo) || 'desconhecido' });
        } catch (e) {
          pulados.push({ nome: f.name, motivo: (e && e.message) || String(e) });
        }
      }
    } finally {
      conn.release();
    }

    const processado = Math.min(offset + limit, total);
    context.res = { status: 200, headers: { 'Content-Type': 'application/json', 'Cache-Control': 'no-store' },
      body: {
        ok: true, total: total, offset: offset, processado: processado,
        carregadosNoBatch: carregados.length, puladosNoBatch: pulados.length, pulados: pulados,
        done: processado >= total, next: processado < total ? processado : null
      } };
  } catch (err) {
    context.log && context.log.error && context.log.error('CarregarBaseCurada:', err);
    context.res = { status: 500, headers: { 'Content-Type': 'application/json' },
      body: { error: (err && err.message) || String(err) } };
  }
};
