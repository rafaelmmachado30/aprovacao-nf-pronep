/**
 * /api/CarregarBaseCurada (GET|POST) — ADMIN ONLY. Trilho RELACIONAL (Supabase).
 *
 * Le a fonte unica curada (_RAG/curado/*.md), extrai o canonico e popula as tabelas
 * do Postgres/Supabase. O relacional e DERIVADO — regeneravel do zero a qualquer hora.
 *
 * Como so ha ~dezenas de contratos no piloto e a fonte ja esta curada (sem chamada de
 * IA aqui), roda de uma vez. offset 0 faz TRUNCATE CASCADE (reload limpo).
 *
 * Query: ?ping=1 (so testa conexao) | ?wipe=0 (nao limpa antes)
 * Resposta: { ok, carregados, pulados, detalhes }
 */

require('isomorphic-fetch');
const { requireAdmin } = require('../shared/authz');
const { getGraphClient, resolveContratosSite } = require('../shared/contratos');
const { carregarCurados } = require('../shared/ragContratos');
const pg = require('../shared/pgContratos');
const loader = require('../shared/loaderCurado');

module.exports = async function (context, req) {
  try {
    const authz = await requireAdmin(context, req);
    if (!authz) return;

    // Teste de conectividade rapido.
    if (req.query && (req.query.ping === '1' || req.query.ping === 'true')) {
      const p = await pg.ping();
      context.res = { status: 200, headers: { 'Content-Type': 'application/json', 'Cache-Control': 'no-store' }, body: p };
      return;
    }
    if (!pg.getPool()) {
      context.res = { status: 400, headers: { 'Content-Type': 'application/json' },
        body: { ok: false, error: 'SUPABASE_DB_URL nao configurada no Function App (Configuration > Application settings).' } };
      return;
    }

    const client = getGraphClient();
    const site = await resolveContratosSite(client);
    const itens = await carregarCurados(client, site.driveId, true); // force = pega o estado atual
    if (!itens || !itens.length) {
      context.res = { status: 200, headers: { 'Content-Type': 'application/json' },
        body: { ok: true, carregados: 0, pulados: 0, aviso: 'Nenhuma fonte curada encontrada. Rode "Curar piloto" antes.' } };
      return;
    }

    const wipe = !(req.query && (req.query.wipe === '0' || req.query.wipe === 'false'));
    const pool = pg.getPool();
    const conn = await pool.connect();
    const carregados = [];
    const pulados = [];
    try {
      if (wipe) await loader.limparTudo(conn);
      for (const canon of itens) {
        try {
          const r = await loader.carregar(conn, canon, { arquivoOrigem: canon._arquivo });
          if (r && r.ok) carregados.push({ operadora: r.operadora, estado: r.estado, id: r.contratoId, arquivo: canon._arquivo });
          else pulados.push({ arquivo: canon._arquivo, motivo: (r && r.motivo) || 'desconhecido' });
        } catch (e) {
          pulados.push({ arquivo: canon._arquivo, motivo: (e && e.message) || String(e) });
        }
      }
    } finally {
      conn.release();
    }

    context.res = { status: 200, headers: { 'Content-Type': 'application/json', 'Cache-Control': 'no-store' },
      body: { ok: true, fontes: itens.length, carregados: carregados.length, pulados: pulados.length,
              detalhes: { carregados: carregados, pulados: pulados } } };
  } catch (err) {
    context.log && context.log.error && context.log.error('CarregarBaseCurada:', err);
    context.res = { status: 500, headers: { 'Content-Type': 'application/json' },
      body: { error: (err && err.message) || String(err) } };
  }
};
