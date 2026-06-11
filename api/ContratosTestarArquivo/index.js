/**
 * /api/ContratosTestarArquivo?driveItemId=X
 * OU
 * /api/ContratosTestarArquivo?pasta=/caminho/relativo/no/SP
 *
 * Processa UM unico arquivo isoladamente:
 *   - se driveItemId: usa o ID direto
 *   - se pasta: lista os files dessa pasta e usa o PRIMEIRO PDF
 *
 * Baixa, extrai texto, chama Claude, retorna o resultado SEM persistir.
 * Diagnostico granular: cada step com tempo/sucesso/erro.
 *
 * Anonymous pra facilitar diagnostico (somente admin no SincronizarContratos).
 */

module.exports = async function (context, req) {
  // C5: diagnostico — restrito a admin.
  const { requireAdmin } = require('../shared/authz');
  if (!(await requireAdmin(context, req))) return;
  const diag = { steps: [], inicio: Date.now() };
  function logStep(nome, extra) {
    diag.steps.push(Object.assign({ step: nome, ms: Date.now() - diag.inicio }, extra || {}));
  }

  let contratos;
  try {
    contratos = require('../shared/contratos');
    logStep('require_contratos');
  } catch (e) {
    context.res = { status: 500, headers: { 'Content-Type': 'application/json' },
      body: { error: 'falha ao carregar shared/contratos: ' + e.message, stack: (e.stack || '').split('\n').slice(0, 8) } };
    return;
  }

  try {
    const driveItemId = (req.query && req.query.driveItemId) || '';
    const pastaRel = (req.query && req.query.pasta) || '';

    logStep('inicio', { driveItemId: !!driveItemId, pastaRel });

    const client = contratos.getGraphClient();
    logStep('graph_client');

    const { driveId } = await contratos.resolveContratosSite(client);
    logStep('site_resolvido', { driveId: !!driveId });

    let arq;
    if (driveItemId) {
      // Pega metadata do arquivo direto
      const item = await client.api('/drives/' + driveId + '/items/' + driveItemId).get();
      arq = {
        nome: item.name,
        id: item.id,
        size: item.size,
        ext: (item.name || '').split('.').pop().toLowerCase(),
        webUrl: item.webUrl
      };
    } else if (pastaRel) {
      const listing = await contratos.listarPasta(client, driveId, pastaRel);
      logStep('pasta_listada', { folders: 0, files: listing.files.length });
      if (!listing.files.length) {
        // Tenta primeira subpasta
        const subs = listing.folders || [];
        if (subs.length) {
          logStep('descendo_subpasta', { sub: subs[0].name });
          const sub = await contratos.listarPasta(client, driveId, subs[0].path);
          if (sub.files.length) {
            arq = sub.files[0];
            arq.ext = arq.ext || (arq.name || '').split('.').pop().toLowerCase();
          }
        }
      } else {
        arq = listing.files[0];
        arq.ext = arq.ext || (arq.name || '').split('.').pop().toLowerCase();
      }
    } else {
      context.res = { status: 400, body: { error: 'Passe driveItemId ou pasta' } };
      return;
    }

    if (!arq) {
      context.res = { status: 404, body: { error: 'Arquivo nao encontrado', diag } };
      return;
    }
    logStep('arquivo_encontrado', { nome: arq.nome, size: arq.size, ext: arq.ext });

    // Baixar + extrair texto
    let ext;
    try {
      ext = await contratos.extrairTexto(client, driveId, arq.id, arq.ext, { tamanhoMaxMB: 10 });
      logStep('texto_extraido', { vazio: ext.vazio, len: (ext.texto || '').length, erro: ext.erro });
    } catch (e) {
      logStep('texto_falhou', { erro: e.message });
      context.res = {
        status: 200,
        headers: { 'Content-Type': 'application/json' },
        body: { ok: false, etapa: 'extracao_texto', erro: e.message, arquivo: { nome: arq.nome, size: arq.size }, diag }
      };
      return;
    }

    if (ext.vazio) {
      context.res = {
        status: 200,
        headers: { 'Content-Type': 'application/json' },
        body: {
          ok: true,
          arquivo: { nome: arq.nome, size: arq.size },
          textoVazio: true,
          motivo: 'PDF sem texto (provavelmente scaneado, requer OCR)',
          erro: ext.erro,
          diag
        }
      };
      return;
    }

    // Chamar Claude
    const vig = await contratos.extrairVigenciaInteligente(ext.texto);
    logStep('vigencia_extraida', {
      dataInicio: vig.dataInicio,
      dataFim: vig.dataFim,
      confidence: vig.confidence,
      naoEncontrou: !!vig.naoEncontrou,
      modelo: vig._modelo,
      tokensIn: vig._tokensIn,
      tokensOut: vig._tokensOut
    });

    context.res = {
      status: 200,
      headers: { 'Content-Type': 'application/json' },
      body: {
        ok: true,
        arquivo: { nome: arq.nome, size: arq.size, ext: arq.ext, webUrl: arq.webUrl },
        textoExtraido: {
          tamanho: ext.texto.length,
          inicio: ext.texto.slice(0, 800),
          fim: ext.texto.slice(-400)
        },
        vigencia: vig,
        statusCalculado: contratos.calcularStatus(vig.dataFim, vig.indeterminado),
        diasParaVencer: contratos.calcularDiasParaVencer(vig.dataFim),
        diag
      }
    };
  } catch (err) {
    context.res = {
      status: 500,
      headers: { 'Content-Type': 'application/json' },
      body: {
        ok: false,
        error: err.message,
        statusCode: err.statusCode,
        body: err.body,
        stack: (err.stack || '').split('\n').slice(0, 8),
        diag
      }
    };
  }
};
