/**
 * shared/pgContratos.js — Conexao Postgres (Supabase) para o trilho RELACIONAL.
 *
 * Pool unico reutilizado entre invocacoes da Function. Connection string vem do
 * secret SUPABASE_DB_URL (ou DATABASE_URL). Supabase exige SSL.
 *
 * Expoe:
 *  - getPool(): Pool | null (null se secret ausente)
 *  - query(sql, params): resultado
 *  - runSelectSeguro(sql): SELECT read-only com guarda (bloqueia escrita/DDL)
 *  - ping(): { ok, versao } — teste de conectividade
 */

let _pool = null;
let _Pool = null;

function _connString() {
  return process.env.SUPABASE_DB_URL || process.env.DATABASE_URL || '';
}

function getPool() {
  const cs = _connString();
  if (!cs) return null;
  if (_pool) return _pool;
  if (!_Pool) { _Pool = require('pg').Pool; }
  _pool = new _Pool({
    connectionString: cs,
    ssl: { rejectUnauthorized: false }, // Supabase usa cert gerenciado
    max: 3,                              // Functions: poucas conexoes
    idleTimeoutMillis: 30000,
    connectionTimeoutMillis: 10000,
    statement_timeout: 15000            // corta query travada (ms)
  });
  return _pool;
}

async function query(sql, params) {
  const pool = getPool();
  if (!pool) throw new Error('SUPABASE_DB_URL nao configurada');
  return await pool.query(sql, params || []);
}

async function ping() {
  try {
    const r = await query('select version() as v');
    return { ok: true, versao: (r.rows && r.rows[0] && r.rows[0].v) || null };
  } catch (e) { return { ok: false, erro: (e && e.message) || String(e) }; }
}

// Guarda de leitura: aceita 1 comando SELECT/WITH; bloqueia escrita/DDL/multi-statement.
const _PROIBIDO = /\b(insert|update|delete|drop|alter|create|grant|revoke|truncate|copy|vacuum|reindex|call|do|merge|comment|set|reset)\b/i;
function _validarSelect(sql) {
  const s = String(sql || '').trim().replace(/;+\s*$/, ''); // tira ; final
  if (!s) return { ok: false, erro: 'SQL vazio' };
  if (s.indexOf(';') >= 0) return { ok: false, erro: 'Apenas UM comando SELECT e permitido (sem ;).' };
  if (!/^(select|with)\b/i.test(s)) return { ok: false, erro: 'So SELECT/WITH e permitido.' };
  if (_PROIBIDO.test(s)) return { ok: false, erro: 'Comando de escrita/DDL bloqueado (base somente leitura pela SAN).' };
  return { ok: true, sql: s };
}

async function runSelectSeguro(sql, maxLinhas) {
  const v = _validarSelect(sql);
  if (!v.ok) return { erro: v.erro };
  const lim = Math.min(500, Math.max(1, maxLinhas || 200));
  // Envolve em subquery com LIMIT rigido + transacao read-only.
  const wrapped = 'select * from (' + v.sql + ') _q limit ' + lim;
  const pool = getPool();
  if (!pool) return { erro: 'SUPABASE_DB_URL nao configurada' };
  const client = await pool.connect();
  try {
    await client.query('begin transaction read only');
    await client.query('set local statement_timeout = 15000');
    const r = await client.query(wrapped);
    await client.query('commit');
    return { colunas: (r.fields || []).map(function (f) { return f.name; }), linhas: r.rows, total: r.rowCount };
  } catch (e) {
    try { await client.query('rollback'); } catch (_) { /* ignora */ }
    return { erro: (e && e.message) || String(e) };
  } finally {
    client.release();
  }
}

module.exports = { getPool, query, ping, runSelectSeguro };
