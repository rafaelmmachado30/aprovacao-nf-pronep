/**
 * /api/CurarContratos (GET|POST) — ADMIN ONLY. Fase 1 da NOVA arquitetura (ingestao curada).
 *
 * Le os contratos das operadoras PILOTO (Amil, Bradesco, CABESP), extrai o modelo
 * CANONICO (curadoriaContrato) e grava a FONTE UNICA curada (Markdown+YAML) em
 * _RAG/curado/ no drive dos contratos. Dessa fonte unica os trilhos (relacional/
 * vetorial/grafo) sao derivados depois — sem perda e regeneravel.
 *
 * Roda em LOTES (o front chama em loop, driblando o timeout do SWA). offset=0 recria
 * o manifest do piloto e limpa a pasta curado/.
 *
 * Query: ?offset=0&limit=4   |   ?debug=1&limit=1 (nao-destrutivo)
 * Resposta: { ok, total, offset, processado, arquivosNoBatch, curadosNoBatch, done, next }
 */

require('isomorphic-fetch');
const { requireAdmin } = require('../shared/authz');
const {
  getGraphClient, garantirListaContratos, getContratoColMap,
  resolveContratosSite, extrairTexto
} = require('../shared/contratos');
const { salvarShard, lerJson } = require('../shared/ragContratos');
const { curar, ultimoErroCuradoria, MODEL } = require('../shared/curadoriaContrato');
const { nomeCurado, montarMarkdown, salvarCurado, limparCurado } = require('../shared/fonteUnica');

// Operadoras do piloto (por Fornecedor). CABESP tolera acento/variacoes.
const PILOTO = /amil|bradesco|cabesp/i;
const EXCLUIR_SUBPASTA = /glosa|recurso de glosas/i;
const EXTS_OK = ['pdf', 'docx'];

const MANIFEST_PILOTO = '_curado_files.json';

function _norm(s) { return String(s || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '').trim(); }
function _ext(nome) { const m = /\.([a-z0-9]+)$/i.exec(String(nome || '')); return m ? m[1].toLowerCase() : ''; }

module.exports = async function (context, req) {
  const diag = { step: 'start', modelo: MODEL };
  try {
    const authz = await requireAdmin(context, req);
    if (!authz) return;

    const offset = Math.max(0, parseInt((req.query && req.query.offset) || '0', 10) || 0);
    const limit = Math.min(8, Math.max(1, parseInt((req.query && req.query.limit) || '4', 10) || 4));

    const client = getGraphClient();
    diag.step = 'resolve';
    const listInfo = await garantirListaContratos(client);
    const appSiteId = listInfo.siteId;
    const listId = listInfo.listId;
    const colMap = await getContratoColMap(client, appSiteId, listId);
    const contr = await resolveContratosSite(client);
    const driveId = contr.driveId;
    const col = function (d) { return colMap[d] || d; };

    // ===== Manifest do piloto (offset 0 recria e limpa curado/).
    diag.step = 'manifest';
    if (offset === 0) await limparCurado(client, driveId);
    let manifest = offset === 0 ? null : await lerJson(client, driveId, MANIFEST_PILOTO);
    if (!manifest || !Array.isArray(manifest.files)) {
      const files = [];
      let url = '/sites/' + appSiteId + '/lists/' + listId + '/items?expand=fields&$top=500';
      let pages = 0;
      while (url && pages < 40) {
        const resp = await client.api(url).get();
        for (const it of (resp.value || [])) {
          const f = it.fields || {};
          if (_norm(f[col('Diretoria')]) !== 'comercial') continue;
          const driveItemId = f[col('DriveItemId')] || '';
          if (!driveItemId) continue;
          const nome = f[col('NomeArquivo')] || f[col('Title')] || '';
          const ext = _ext(nome);
          if (EXTS_OK.indexOf(ext) < 0) continue;
          const fornecedor = f[col('Fornecedor')] || '';
          const caminho = String(f[col('PathRelativoSP')] || '') + ' ' + String(f[col('CaminhoSharepoint')] || '') + ' ' + String(nome);
          if (EXCLUIR_SUBPASTA.test(caminho)) continue;
          // Piloto: casa pelo Fornecedor OU pelo caminho (pasta da operadora).
          if (!PILOTO.test(fornecedor) && !PILOTO.test(caminho)) continue;
          files.push({
            listItemId: it.id, driveItemId: driveItemId, nome: nome, ext: ext,
            fornecedor: fornecedor,
            subpasta: f[col('PathRelativoSP')] || f[col('CaminhoSharepoint')] || '',
            webUrl: f[col('CaminhoSharepoint')] || ''
          });
        }
        pages++;
        url = resp['@odata.nextLink'] ? resp['@odata.nextLink'].replace('https://graph.microsoft.com/v1.0', '') : null;
      }
      manifest = { total: files.length, files: files, criadoEm: new Date().toISOString() };
      await salvarShard(client, driveId, MANIFEST_PILOTO, manifest);
    }

    const total = manifest.files.length;
    const slice = manifest.files.slice(offset, offset + limit);

    // ===== MODO DEBUG (nao-destrutivo).
    if (req.query && (req.query.debug === '1' || req.query.debug === 'true')) {
      const amostra = [];
      for (const f of slice.slice(0, Math.min(limit, 2))) {
        const info = { nome: f.nome, fornecedor: f.fornecedor, ext: f.ext };
        try {
          const ex = await extrairTexto(client, driveId, f.driveItemId, f.ext, {});
          const texto = (ex && ex.texto) || '';
          info.textoLen = texto.length;
          const canonico = await curar(texto, 'nativo');
          info.curadoOk = !!canonico;
          info.curadoErro = canonico ? null : ultimoErroCuradoria();
          if (canonico) info.amostraCanonico = { doc_tipo: canonico.doc_tipo, operadora: canonico.operadora && canonico.operadora.nome, estado: canonico.estado_uf, status: canonico.contrato && canonico.contrato.status, diarias: (canonico.diarias || []).length, clausulas: (canonico.clausulas || []).length };
        } catch (e) { info.erro = (e && e.message) || String(e); }
        amostra.push(info);
      }
      context.res = { status: 200, headers: { 'Content-Type': 'application/json', 'Cache-Control': 'no-store' },
        body: { ok: true, debug: true, modelo: MODEL, hasAnthropicKey: !!process.env.ANTHROPIC_API_KEY, hasOpenAIKey: !!process.env.OPENAI_API_KEY, totalManifest: total, amostra: amostra } };
      return;
    }

    // ===== Processa o lote EM PARALELO: extrai texto -> cura -> grava fonte unica.
    diag.step = 'process_batch';
    const curadosOut = [];
    const ignorados = [];
    const resultados = await Promise.all(slice.map(async function (f) {
      try {
        const ex = await extrairTexto(client, driveId, f.driveItemId, f.ext, {});
        const texto = (ex && ex.texto) || '';
        if (!texto || texto.length < 60) return { nome: f.nome, motivo: 'sem_texto' };
        const canonico = await curar(texto, 'nativo');
        if (!canonico) return { nome: f.nome, motivo: 'curadoria_falhou', erro: ultimoErroCuradoria() };
        // So contrato/aditivo viram fonte unica; glosa/outro sao ruido (nao entram).
        if (canonico.doc_tipo && ['glosa', 'outro'].indexOf(canonico.doc_tipo) >= 0) {
          return { nome: f.nome, motivo: 'doc_tipo_' + canonico.doc_tipo };
        }
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
