/**
 * shared/loaderCurado.js — Deriva o trilho RELACIONAL a partir do canonico curado.
 *
 * Recebe UM canonico (o mesmo JSON gravado na fonte unica) e insere nas tabelas do
 * Supabase (operadora, contrato, reajuste, aditivo, diaria, procedimento, regra_matmed,
 * clausula, risco, proveniencia). O relacional e DERIVADO — pode ser regenerado do zero.
 *
 * Estrategia: o endpoint faz TRUNCATE ... CASCADE no inicio (reload limpo) e depois
 * chama carregar() por contrato. operadora e upsert (UNIQUE nome); o resto e insert.
 *
 * Regras: enums sao saneados; datas so entram se ISO (YYYY-MM-DD); numeros toleram
 * virgula decimal; estado invalido/null => contrato NAO entra (integridade) e retorna skip.
 */

const UF = ['SP', 'RJ', 'ES'];

function _num(v) {
  if (v == null) return null;
  if (typeof v === 'number') return isFinite(v) ? v : null;
  let s = String(v).replace(/[^\d.,-]/g, '').trim();
  if (!s) return null;
  if (s.indexOf(',') >= 0 && s.indexOf('.') >= 0) s = s.replace(/\./g, '').replace(',', '.'); // 1.234,56
  else if (s.indexOf(',') >= 0) s = s.replace(',', '.');
  const n = parseFloat(s);
  return isFinite(n) ? n : null;
}
function _date(v) {
  if (!v) return null;
  const s = String(v).trim();
  const m = /^(\d{4}-\d{2}-\d{2})/.exec(s);
  return m ? m[1] : null;
}
function _enum(v, allowed, fallback) {
  if (v == null) return fallback != null ? fallback : null;
  const s = String(v).trim();
  for (const a of allowed) { if (a.toLowerCase() === s.toLowerCase()) return a; }
  return fallback != null ? fallback : null;
}
function _bool(v) { return v === true || v === 'true' || v === 'sim'; }
function _txt(v) { return (v == null || v === '') ? null : String(v); }
function _jsonb(v) { return v == null ? null : JSON.stringify(v); }

// TRUNCATE de reload limpo (chamar 1x no inicio da carga).
async function limparTudo(client) {
  await client.query('truncate operadora, contrato, reajuste, aditivo, diaria, procedimento, regra_matmed, clausula, risco, proveniencia, clausula_chunk restart identity cascade');
}

async function _upsertOperadora(client, op) {
  op = op || {};
  const nome = _txt(op.nome) || 'Desconhecida';
  const r = await client.query(
    'insert into operadora (nome, nome_completo, cnpj, registro_ans, segmento) values ($1,$2,$3,$4,$5) ' +
    'on conflict (nome) do update set cnpj = coalesce(excluded.cnpj, operadora.cnpj), ' +
    'registro_ans = coalesce(excluded.registro_ans, operadora.registro_ans), ' +
    'segmento = coalesce(excluded.segmento, operadora.segmento) returning id',
    [nome, _txt(op.nome_completo), _txt(op.cnpj), _txt(op.registro_ans), _txt(op.segmento)]
  );
  return r.rows[0].id;
}

async function _insProveniencia(client, canon, meta) {
  const prov = canon.proveniencia || {};
  const r = await client.query(
    'insert into proveniencia (arquivo_origem, web_url, metodo_extracao, selo_confianca, criado_em) ' +
    'values ($1,$2,$3,$4, now()) returning id',
    [
      _txt((meta && meta.arquivoOrigem)) || _txt(canon._arquivo) || 'curado',
      _txt(meta && meta.webUrl),
      _enum(prov.metodo_extracao, ['nativo', 'OCR'], 'nativo'),
      _enum(prov.selo_confianca, ['CONFIRMADO', 'PARCIAL', 'PENDENTE'], 'PENDENTE')
    ]
  );
  return r.rows[0].id;
}

// Carrega 1 canonico. Retorna { ok, contratoId } ou { skip, motivo }.
async function carregar(client, canon, meta) {
  canon = canon || {};
  // UF e OBRIGATORIA (regra Pronep: contrato mora na pasta da unidade SP/RJ/ES).
  // A curadoria ja deriva a UF da pasta quando o conteudo e omisso; se ainda faltar,
  // e um contrato mal arquivado — pula e sinaliza, em vez de gravar sem unidade.
  const estado = _enum(canon.estado_uf, UF, null);
  if (!estado) return { skip: true, motivo: 'sem_uf (verificar pasta da unidade)', arquivo: canon._arquivo };

  const provId = await _insProveniencia(client, canon, meta);
  const opId = await _upsertOperadora(client, canon.operadora);
  const ct = canon.contrato || {};

  const cr = await client.query(
    'insert into contrato (operadora_id, estado_uf, numero, cnpj_pronep_contratada, objeto, data_assinatura, ' +
    'inicio_vigencia, fim_vigencia, prazo, status, tem_liminar, qualidade_redacao_score, qualidade_redacao_notas, ' +
    'completude, drive_item_id, proveniencia_id) ' +
    'values ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,$12,$13,$14,$15,$16) returning id',
    [
      opId, estado, _txt(ct.numero), _txt(ct.cnpj_pronep_contratada), _txt(ct.objeto), _date(ct.data_assinatura),
      _date(ct.inicio_vigencia), _date(ct.fim_vigencia), _enum(ct.prazo, ['determinado', 'indeterminado'], null),
      _enum(ct.status, ['VIGENTE', 'HISTORICO'], 'VIGENTE'), _bool(ct.tem_liminar),
      (ct.qualidade_redacao_score != null ? Math.max(0, Math.min(5, parseInt(ct.qualidade_redacao_score, 10) || 0)) : null),
      _txt(ct.qualidade_redacao_notas), _jsonb(canon.completude), _txt(meta && meta.driveItemId), provId
    ]
  );
  const contratoId = cr.rows[0].id;

  for (const r of (canon.reajustes || [])) {
    await client.query(
      'insert into reajuste (contrato_id, numero_rp, tipo, indice, percentual, inicio_vigencia, data_base, proximo_reajuste_previsto, proveniencia_id) ' +
      'values ($1,$2,$3,$4,$5,$6,$7,$8,$9)',
      [contratoId, _txt(r.numero_rp), _enum(r.tipo, ['GERAL', 'pontual'], 'pontual'), _txt(r.indice), _num(r.percentual),
       _date(r.inicio_vigencia), _date(r.data_base), _date(r.proximo_reajuste_previsto), provId]
    );
  }
  for (const a of (canon.aditivos || [])) {
    await client.query(
      'insert into aditivo (contrato_id, objeto, inicio_vigencia, tipo, melhora_para_pronep, melhora_justificativa, proveniencia_id) ' +
      'values ($1,$2,$3,$4,$5,$6,$7)',
      [contratoId, _txt(a.objeto), _date(a.inicio_vigencia), _enum(a.tipo, ['geral', 'pontual'], null),
       _enum(a.melhora_para_pronep, ['sim', 'nao', 'neutro'], null), _txt(a.melhora_justificativa), provId]
    );
  }
  for (const d of (canon.diarias || [])) {
    await client.query(
      'insert into diaria (contrato_id, descricao, valor_diaria, inclusos, exclusos, vigencia_inicio, vigencia_fim, confianca_preco, proveniencia_id) ' +
      'values ($1,$2,$3,$4,$5,$6,$7,$8,$9)',
      [contratoId, _txt(d.descricao), _num(d.valor_diaria), _jsonb(d.inclusos || []), _jsonb(d.exclusos || []),
       _date(d.vigencia_inicio), _date(d.vigencia_fim), _enum(d.confianca_preco, ['NATIVO', 'OCR-VALIDADO', 'OCR-A-VALIDAR'], 'OCR-A-VALIDAR'), provId]
    );
  }
  for (const p of (canon.procedimentos || [])) {
    await client.query(
      'insert into procedimento (contrato_id, codigo_tuss, descricao, valor, unidade, confianca_preco, proveniencia_id) ' +
      'values ($1,$2,$3,$4,$5,$6,$7)',
      [contratoId, _txt(p.codigo_tuss), _txt(p.descricao), _num(p.valor), _txt(p.unidade),
       _enum(p.confianca_preco, ['NATIVO', 'OCR-VALIDADO', 'OCR-A-VALIDAR'], 'OCR-A-VALIDAR'), provId]
    );
  }
  for (const m of (canon.matmed || [])) {
    await client.query(
      'insert into regra_matmed (contrato_id, base, operador, percentual, categoria, proveniencia_id) values ($1,$2,$3,$4,$5,$6)',
      [contratoId, _enum(m.base, ['Simpro', 'Brasindice'], null), _enum(m.operador, ['deflator', 'acrescimo'], null),
       _num(m.percentual), _txt(m.categoria), provId]
    );
  }
  for (const cl of (canon.clausulas || [])) {
    await client.query(
      'insert into clausula (contrato_id, tipo, texto_literal, paragrafo_ref, proveniencia_id) values ($1,$2,$3,$4,$5)',
      [contratoId, _txt(cl.tipo), _txt(cl.texto_literal), _txt(cl.paragrafo_ref), provId]
    );
  }
  for (const rk of (canon.riscos || [])) {
    await client.query(
      'insert into risco (contrato_id, descricao, categoria, severidade, proveniencia_id) values ($1,$2,$3,$4,$5)',
      [contratoId, _txt(rk.descricao), _enum(rk.categoria, ['financeiro', 'juridico', 'operacional'], null),
       _enum(rk.severidade, ['alto', 'medio', 'baixo'], null), provId]
    );
  }
  return { ok: true, contratoId: contratoId, operadora: canon.operadora && canon.operadora.nome, estado: estado };
}

module.exports = { limparTudo, carregar };
