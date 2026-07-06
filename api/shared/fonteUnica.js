/**
 * shared/fonteUnica.js — Fonte única curada por contrato (Markdown + YAML).
 *
 * A partir do modelo canônico (curadoriaContrato), gera um .md com:
 *  - frontmatter YAML (campos-chave, leitura humana + parse rápido);
 *  - corpo Markdown (diárias, matmed, cláusulas literais, riscos, proveniência);
 *  - bloco ```json``` com o canônico COMPLETO — é dele que os derivados (relacional,
 *    grafo, vetor) são regenerados de forma determinística.
 *
 * Guarda em _RAG/curado/ no drive dos contratos (versionado pelo histórico do SP).
 */

require('isomorphic-fetch');

const CURADO_FOLDER = '_RAG/curado';

function _slug(s) {
  return String(s || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '')
    .replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '').slice(0, 60) || 'contrato';
}
function nomeCurado(canonico, fallback, id) {
  const op = _slug((canonico && canonico.operadora && canonico.operadora.nome) || fallback || 'contrato');
  const uf = _slug((canonico && canonico.estado_uf) || 'na');
  const num = _slug((canonico && canonico.contrato && canonico.contrato.numero) || '');
  // Sufixo com o id garante unicidade (evita 2 contratos gerarem o mesmo arquivo).
  const sufixo = id != null ? String(id).replace(/[^a-z0-9]/gi, '') : '';
  return [op, uf, num].filter(Boolean).join('_') + (sufixo ? ('_' + sufixo) : '') + '.md';
}

// YAML-escapa um escalar simples.
function _y(v) {
  if (v === null || v === undefined) return 'null';
  if (typeof v === 'number' || typeof v === 'boolean') return String(v);
  return JSON.stringify(String(v)); // string entre aspas (YAML aceita)
}

function montarMarkdown(canonico, meta) {
  meta = meta || {};
  const c = canonico || {};
  const op = c.operadora || {};
  const ct = c.contrato || {};
  const prov = c.proveniencia || {};
  const L = [];
  // Frontmatter
  L.push('---');
  L.push('doc_tipo: ' + _y(c.doc_tipo || 'contrato'));
  L.push('operadora: ' + _y(op.nome));
  L.push('operadora_cnpj: ' + _y(op.cnpj));
  L.push('registro_ans: ' + _y(op.registro_ans));
  L.push('estado_uf: ' + _y(c.estado_uf));
  L.push('numero: ' + _y(ct.numero));
  L.push('status: ' + _y(ct.status));
  L.push('inicio_vigencia: ' + _y(ct.inicio_vigencia));
  L.push('fim_vigencia: ' + _y(ct.fim_vigencia));
  L.push('tem_liminar: ' + _y(ct.tem_liminar));
  L.push('qualidade_redacao_score: ' + _y(ct.qualidade_redacao_score));
  L.push('selo_confianca: ' + _y(prov.selo_confianca));
  L.push('arquivo_origem: ' + _y(meta.arquivoOrigem));
  L.push('web_url: ' + _y(meta.webUrl));
  L.push('curado_em: ' + _y(new Date().toISOString()));
  L.push('---');
  L.push('');
  L.push('# ' + (op.nome || 'Contrato') + ' — ' + (c.estado_uf || '') + (ct.numero ? (' (nº ' + ct.numero + ')') : ''));
  L.push('');
  L.push('**Objeto:** ' + (ct.objeto || '—'));
  L.push('');
  // Reajustes
  if ((c.reajustes || []).length) {
    L.push('## Reajustes');
    c.reajustes.forEach(function (r) {
      L.push('- **' + (r.tipo || '?') + '** · índice: ' + (r.indice || '—') + ' · % : ' + (r.percentual != null ? r.percentual : '—') +
        ' · data-base: ' + (r.data_base || '—') + ' · próximo: ' + (r.proximo_reajuste_previsto || '—'));
    });
    L.push('');
  }
  // Diárias
  if ((c.diarias || []).length) {
    L.push('## Diárias / Pacotes');
    c.diarias.forEach(function (d) {
      L.push('- **' + (d.descricao || '—') + '** — R$ ' + (d.valor_diaria != null ? d.valor_diaria : '—') +
        ' [' + (d.confianca_preco || '—') + ']');
      if ((d.inclusos || []).length) L.push('  - inclui: ' + d.inclusos.join(', '));
      if ((d.exclusos || []).length) L.push('  - exclui: ' + d.exclusos.join(', '));
    });
    L.push('');
  }
  // MatMed
  if ((c.matmed || []).length) {
    L.push('## Regras MatMed');
    c.matmed.forEach(function (m) {
      L.push('- ' + (m.categoria || '—') + ': ' + (m.base || '—') + ' ' + (m.operador || '') + ' ' + (m.percentual != null ? m.percentual + '%' : ''));
    });
    L.push('');
  }
  // Cláusulas
  if ((c.clausulas || []).length) {
    L.push('## Cláusulas');
    c.clausulas.forEach(function (cl) {
      L.push('### ' + (cl.tipo || 'cláusula') + (cl.paragrafo_ref ? (' (' + cl.paragrafo_ref + ')') : ''));
      L.push('> ' + String(cl.texto_literal || '').replace(/\n/g, '\n> '));
      L.push('');
    });
  }
  // Riscos
  if ((c.riscos || []).length) {
    L.push('## Riscos');
    c.riscos.forEach(function (r) { L.push('- [' + (r.severidade || '?') + '/' + (r.categoria || '?') + '] ' + (r.descricao || '')); });
    L.push('');
  }
  // Canônico completo (fonte dos derivados)
  L.push('## Dados estruturados (canônico)');
  L.push('```json');
  L.push(JSON.stringify(canonico, null, 2));
  L.push('```');
  return L.join('\n');
}

function _enc(p) { return encodeURIComponent(p).replace(/%2F/g, '/'); }

async function salvarCurado(client, driveId, nome, conteudo) {
  const up = '/drives/' + driveId + '/root:/' + _enc(CURADO_FOLDER + '/' + nome) + ':/content';
  return await client.api(up).header('Content-Type', 'text/markdown').put(Buffer.from(conteudo, 'utf-8'));
}

async function limparCurado(client, driveId) {
  try {
    const resp = await client.api('/drives/' + driveId + '/root:/' + _enc(CURADO_FOLDER) + ':/children').select('id,name').top(999).get();
    const alvos = (resp.value || []).filter(function (x) { return /\.md$/i.test(x.name || ''); });
    await Promise.all(alvos.map(async function (x) {
      try { await client.api('/drives/' + driveId + '/items/' + x.id).delete(); } catch (e) { /* ignora */ }
    }));
  } catch (e) { /* pasta ainda nao existe */ }
}

module.exports = { CURADO_FOLDER, nomeCurado, montarMarkdown, salvarCurado, limparCurado };
