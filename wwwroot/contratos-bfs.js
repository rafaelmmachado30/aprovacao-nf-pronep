/**
 * contratos-bfs.js — Orquestrador BFS pra sincronizar contratos do SharePoint
 * descendo subpasta-por-subpasta. Evita timeout do SWA Gateway ao manter cada
 * chamada de Sincronizar em escopo pequeno (uma pasta de prestador por vez).
 *
 * Carregado como <script src="/contratos-bfs.js"></script> no index.html.
 * Expoe: window.contratosBfs.iniciar()
 *
 * Depende de funcoes globais do index.html:
 *   - openModal(html), closeModal()
 *   - toast(msg, tipo)
 *   - carregarContratosReais()  (atualiza a tela depois)
 */
(function() {
  'use strict';

  var ROOT = '/CONTRATOS/CONTRATOS E DOCUMENTOS - PRESTADORES';

  // Diretorias VARRIDAS pelo BFS — somente as que têm contratos de PRESTADORES (custos pagos).
  // Comercial e Financeira foram REMOVIDAS porque contêm contratos de VÍNCULO com a Pronep
  // (Pronep prestando serviço pra outros), não cobranças que recebemos. Decisão Rafa 04/06/2026.
  // Pra incluir no futuro: descomenta as 2 linhas abaixo.
  var DIRETORIAS = [
    { id: '/OUVIDORIA',                                  label: 'Ouvidoria' },
    { id: '/QUALIDADE',                                  label: 'Qualidade' },
    { id: '/PACIENTES PARTICULARES',                     label: 'Pacientes Particulares' },
    { id: '/GERÊNCIA DE RH',                             label: 'RH' },
    // { id: '/DIRETORIA COMERCIAL',                     label: 'Comercial' },   // vinculo prestador-Pronep
    // { id: '/DIRETORIA FINANCEIRA',                    label: 'Financeira' },  // vinculo prestador-Pronep
    { id: '/DIRETORIA DE OPERAÇÕES',                     label: 'Operacoes' },
    { id: '/DIRETORIA DE SUPRIMENTOS E LOGÍSTICA',       label: 'Suprimentos' },
    { id: '/JURÍDICO',                                   label: 'Juridico' },
    { id: '/GERÊNCIA DE PROJETOS E TI',                  label: 'Tecnologia' }
  ];

  // Pastas estruturais dentro de prestador — quando uma subpasta SO contem
  // essas, a pasta-pai eh considerada "folha de processamento" (chama o
  // Sincronizar nela e deixa o crawler do backend resolver).
  var ESTRUTURAIS = ['CONTRATOS', 'DOCUMENTOS', 'ADITIVOS', 'PROPOSTAS', 'NDAS', 'OUTROS'];

  var state = { cancelado: false, erros: [], errosConsecutivos: 0 };

  // CIRCUIT BREAKER — protege contra cenarios catastroficos como
  // "persistir is not defined" que custariam tokens em centenas de arquivos.
  var LIMITE_ERROS_TOTAIS = 20;       // aborta tudo se passar disso
  var LIMITE_ERROS_CONSECUTIVOS = 5;  // aborta se tiver 5 falhas em sequencia
  var lastErrorRate = 0;

  function normalizeStr(s) {
    return String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '').toUpperCase().trim();
  }

  function eEstrutural(nome) {
    var n = normalizeStr(nome);
    for (var i = 0; i < ESTRUTURAIS.length; i++) {
      if (ESTRUTURAIS[i] === n) return true;
    }
    return false;
  }

  function souFolhaProcessavel(listagem) {
    if (listagem.pdfs && listagem.pdfs.length > 0) return true;
    if (!listagem.subpastas || !listagem.subpastas.length) return false;
    for (var i = 0; i < listagem.subpastas.length; i++) {
      if (!eEstrutural(listagem.subpastas[i].nome)) return false;
    }
    return true;
  }

  async function listarSubpastas(path) {
    var r = await fetch('/api/ListarSubpastasContratos?path=' + encodeURIComponent(path), { credentials: 'include' });
    if (!r.ok) throw new Error('HTTP ' + r.status);
    return await r.json();
  }

  async function sincronizarPasta(subpathRel, maxArquivos) {
    var params = new URLSearchParams();
    params.set('pasta', subpathRel.normalize('NFC'));
    params.set('maxArquivos', String(maxArquivos || 10));
    var r = await fetch('/api/SincronizarContratos?' + params.toString(), { credentials: 'include' });
    if (!r.ok) {
      var msg = 'HTTP ' + r.status;
      try { var j = await r.json(); msg = j.error || msg; } catch (e) {}
      throw new Error(msg);
    }
    return await r.json();
  }

  function renderProgresso(elId, idxDir, totalDir, labelDir, msg, acumulado, tempoSeg) {
    var pct = Math.round((idxDir / totalDir) * 100);
    var el = document.getElementById(elId);
    if (!el) return;
    el.innerHTML =
      '<div style="background:#EFF6FF;border-left:3px solid #3B82F6;color:#1E3A8A;padding:14px;border-radius:4px">' +
      '<div style="font-weight:600;margin-bottom:6px">Diretoria ' + (idxDir + 1) + '/' + totalDir + ': <span style="color:#1F4E79">' + labelDir + '</span></div>' +
      '<div style="height:10px;background:#DBEAFE;border-radius:5px;overflow:hidden;margin-bottom:8px">' +
      '<div style="height:100%;width:' + pct + '%;background:#3B82F6;transition:width .3s"></div></div>' +
      '<div style="font-size:12px;color:#475569">' + msg + '</div>' +
      '<div style="font-size:12px;color:#475569;margin-top:6px">Acumulado: <b>' + acumulado.novos + '</b> novos · ' + acumulado.atualizados + ' atualizados · ' + acumulado.pulados + ' já gravados · ' + acumulado.erros + ' erros · ' + acumulado.batches + ' chamadas · ' + tempoSeg + 's</div>' +
      '<div style="margin-top:10px"><button class="btn btn-ghost" style="font-size:12px;padding:4px 10px" onclick="window.contratosBfs.cancelar()">⏹ Cancelar</button></div>' +
      '</div>';
  }

  function renderFinal(elId, acumulado, tempoSeg, cancelado) {
    var min = Math.floor(tempoSeg / 60);
    var seg = tempoSeg % 60;
    var errosHTML = '';
    if (state.erros.length > 0) {
      var itens = state.erros.slice(0, 40).map(function(e) {
        return '<div style="margin-bottom:6px;padding:6px;background:#fff;border-radius:3px;font-size:11px"><b>' + (e.diretoria || '?') + '</b>' + (e.pasta ? ' · <span style="color:#6B7280">' + e.pasta.split('/').slice(-2).join('/') + '</span>' : '') + (e.arquivo ? ' · <i>' + e.arquivo + '</i>' : '') + '<br>' + (e.error || e.erro || '?') + '</div>';
      }).join('');
      errosHTML = '<details style="margin-top:12px;background:#FEF2F2;border:1px solid #FCA5A5;border-radius:4px">' +
        '<summary style="padding:10px;cursor:pointer;font-weight:600;color:#7F1D1D">⚠ Ver detalhes dos ' + state.erros.length + ' erros</summary>' +
        '<div style="padding:10px;font-family:monospace;font-size:11px;max-height:300px;overflow-y:auto">' + itens +
        (state.erros.length > 40 ? '<div style="color:#6B7280">...e mais ' + (state.erros.length - 40) + ' erros</div>' : '') +
        '</div></details>';
    }
    var el = document.getElementById(elId);
    if (!el) return;
    el.innerHTML =
      '<div class="alert alert-success" style="padding:14px;border-radius:4px;display:block">' +
      '<div style="font-weight:600;font-size:16px;margin-bottom:10px">' + (cancelado ? '⏹ Sincronização cancelada' : '✓ Sincronização concluída') + '</div>' +
      '<div style="display:block;margin-bottom:4px"><b>Diretorias processadas:</b> ' + acumulado.diretorias + '/' + DIRETORIAS.length + '</div>' +
      '<div style="display:block;margin-bottom:4px"><b>Prestadores processados:</b> ' + acumulado.prestadores + '</div>' +
      '<div style="display:block;margin-bottom:4px"><b>Novos contratos:</b> ' + acumulado.novos + '</div>' +
      '<div style="display:block;margin-bottom:4px"><b>Atualizados:</b> ' + acumulado.atualizados + '</div>' +
      '<div style="display:block;margin-bottom:4px"><b>Já gravados (pulados):</b> ' + acumulado.pulados + '</div>' +
      '<div style="display:block;margin-bottom:4px"><b>Erros:</b> ' + acumulado.erros + '</div>' +
      '<div style="display:block;margin-bottom:4px"><b>Total de chamadas:</b> ' + acumulado.batches + '</div>' +
      '<div style="display:block;margin-bottom:4px"><b>Tempo total:</b> ' + min + 'm ' + seg + 's</div>' +
      errosHTML +
      '<div style="margin-top:14px"><button class="btn btn-primary" onclick="closeModal(); if(typeof carregarContratosReais===\'function\') carregarContratosReais();">Fechar e atualizar lista</button></div>' +
      '</div>';
  }

  // Pre-flight: chama Sincronizar com dryRun=true E maxArquivos=1 numa pasta CONHECIDA
  // (TOTVS SA, ja tem dados) pra validar que o pipeline funciona ANTES de queimar tokens.
  async function preflightCheck() {
    try {
      var params = new URLSearchParams();
      params.set('pasta', '/GERÊNCIA DE PROJETOS E TI/CORPORATIVO/TOTVS SA/CONTRATOS');
      params.set('maxArquivos', '1');
      params.set('dryRun', 'true');
      var r = await fetch('/api/SincronizarContratos?' + params.toString(), { credentials: 'include' });
      if (!r.ok) {
        var msg = 'HTTP ' + r.status;
        try { var j = await r.json(); msg = j.error || msg; } catch (e) {}
        return { ok: false, motivo: 'Sincronizar retornou ' + msg };
      }
      var data = await r.json();
      // Resposta valida? Verifica campos esperados
      if (!data || typeof data.totalEncontrados !== 'number') {
        return { ok: false, motivo: 'Resposta do Sincronizar sem campos esperados (totalEncontrados)' };
      }
      // Algum erro no resultado individual?
      if ((data.resultados || []).some(function(r){ return r.erro; })) {
        var erro = (data.resultados || []).find(function(r){ return r.erro; });
        return { ok: false, motivo: 'Sincronizar gerou erro mesmo em dry-run: ' + (erro && erro.erro) };
      }
      return { ok: true, exemplo: data };
    } catch (e) {
      return { ok: false, motivo: 'Exception no preflight: ' + e.message };
    }
  }

  async function iniciar() {
    if (!confirm('Vai varrer as 10 diretorias da Pronep no SharePoint via BFS (subpasta por subpasta). Cada chamada é pequena pra evitar timeout. Tempo estimado 30-90 min, custo $1.50 a $4 em tokens Claude.\n\nVai rodar um pre-flight check primeiro (custo ~0). Confirma?')) return;

    state.cancelado = false;
    state.erros = [];
    state.errosConsecutivos = 0;

    var html =
      '<h3 style="margin-top:0">⚡ Sincronizar TUDO (BFS)</h3>' +
      '<div style="background:#FEF3C7;border-left:3px solid #F59E0B;padding:12px;border-radius:4px;margin-bottom:14px">' +
      '<b>Estratégia BFS:</b> mapeia subpastas primeiro, depois processa cada pasta de prestador em uma chamada pequena. Evita timeout em diretorias grandes (Jurídico, Tecnologia).<br>' +
      '<b>Circuit breaker:</b> aborta automaticamente após ' + LIMITE_ERROS_CONSECUTIVOS + ' erros consecutivos ou ' + LIMITE_ERROS_TOTAIS + ' erros totais.' +
      '</div>' +
      '<div id="bfs-progresso"><div class="muted small">Rodando pre-flight check...</div></div>';
    openModal(html);

    // === PRE-FLIGHT CHECK ===
    var preflight = await preflightCheck();
    if (!preflight.ok) {
      document.getElementById('bfs-progresso').innerHTML =
        '<div class="alert alert-error" style="padding:14px;border-radius:4px">' +
        '<b>❌ Pre-flight check FALHOU</b><br>' +
        'O backend retornou erro no teste inicial (sem custo de tokens). NÃO vou iniciar a sincronização real.<br><br>' +
        '<b>Motivo:</b> ' + preflight.motivo + '<br><br>' +
        'Corrija o backend antes de tentar de novo. Comum ser deploy desatualizado ou syntax error em SincronizarContratos.' +
        '<div style="margin-top:14px"><button class="btn btn-secondary" onclick="closeModal()">Fechar</button></div>' +
        '</div>';
      return;
    }
    // Pre-flight OK — segue
    document.getElementById('bfs-progresso').innerHTML = '<div class="alert alert-success" style="padding:10px;border-radius:4px">✓ Pre-flight passou. Iniciando varredura...</div>';
    await new Promise(function(r){ setTimeout(r, 800); });

    var totalStart = Date.now();
    var acumulado = { novos: 0, atualizados: 0, pulados: 0, erros: 0, batches: 0, diretorias: 0, prestadores: 0 };

    for (var i = 0; i < DIRETORIAS.length; i++) {
      if (state.cancelado) break;
      var dir = DIRETORIAS[i];
      var tempoSeg = Math.round((Date.now() - totalStart) / 1000);

      renderProgresso('bfs-progresso', i, DIRETORIAS.length, dir.label, 'Mapeando estrutura...', acumulado, tempoSeg);

      // BFS pra encontrar folhas processaveis
      var fila = [ROOT + dir.id.normalize('NFC')];
      var folhas = [];
      var iteracoesMap = 0;

      while (fila.length > 0 && !state.cancelado && iteracoesMap < 500) {
        iteracoesMap++;
        var pasta = fila.shift();
        try {
          tempoSeg = Math.round((Date.now() - totalStart) / 1000);
          renderProgresso('bfs-progresso', i, DIRETORIAS.length, dir.label, 'Mapeando ' + pasta.split('/').slice(-2).join('/') + ' (' + folhas.length + ' folhas) ...', acumulado, tempoSeg);
          var listagem = await listarSubpastas(pasta);
          if (souFolhaProcessavel(listagem)) {
            folhas.push(pasta);
          } else {
            for (var s = 0; s < (listagem.subpastas || []).length; s++) {
              fila.push(listagem.subpastas[s].path);
            }
          }
        } catch (e) {
          state.erros.push({ diretoria: dir.label, pasta: pasta, error: 'mapear: ' + e.message });
        }
      }

      // Processa cada folha (chamada pequena, escopo de prestador)
      for (var f = 0; f < folhas.length; f++) {
        if (state.cancelado) break;
        // CIRCUIT BREAKER global
        if (state.erros.length >= LIMITE_ERROS_TOTAIS) {
          state.cancelado = true;
          state.erros.push({ diretoria: dir.label, error: 'CIRCUIT_BREAKER: ' + LIMITE_ERROS_TOTAIS + ' erros totais atingidos - abortando pra nao queimar tokens' });
          break;
        }
        if (state.errosConsecutivos >= LIMITE_ERROS_CONSECUTIVOS) {
          state.cancelado = true;
          state.erros.push({ diretoria: dir.label, error: 'CIRCUIT_BREAKER: ' + LIMITE_ERROS_CONSECUTIVOS + ' erros consecutivos - algo esta errado, abortando' });
          break;
        }
        var folha = folhas[f];
        tempoSeg = Math.round((Date.now() - totalStart) / 1000);
        renderProgresso('bfs-progresso', i, DIRETORIAS.length, dir.label,
          'Prestador ' + (f + 1) + '/' + folhas.length + ': ' + folha.split('/').slice(-2).join('/'),
          acumulado, tempoSeg);
        acumulado.batches++;

        var tentativas = 0;
        while (tentativas < 10 && !state.cancelado) {
          tentativas++;
          try {
            var subpath = folha.startsWith(ROOT) ? folha.slice(ROOT.length) : folha;
            var resp = await sincronizarPasta(subpath, 10);
            var stats = resp.stats || {};
            acumulado.novos += (stats.novos || 0);
            acumulado.atualizados += (stats.atualizados || 0);
            acumulado.pulados += (stats.pulados || 0);
            acumulado.erros += (stats.erros || 0);
            var teveErroIndividual = false;
            for (var k = 0; k < (resp.resultados || []).length; k++) {
              var r2 = resp.resultados[k];
              if (r2.erro) {
                state.erros.push({ diretoria: dir.label, pasta: folha, arquivo: r2.nome, erro: r2.erro });
                teveErroIndividual = true;
              }
            }
            // Reset consecutivos somente se nada de errado nessa chamada
            if (!teveErroIndividual && (stats.erros || 0) === 0) {
              state.errosConsecutivos = 0;
            } else {
              state.errosConsecutivos++;
            }
            var restantes = resp.restantes || 0;
            if (restantes === 0) break;
          } catch (e) {
            acumulado.erros++;
            state.errosConsecutivos++;
            state.erros.push({ diretoria: dir.label, pasta: folha, error: 'sincronizar: ' + e.message });
            break;
          }
        }
        acumulado.prestadores++;
      }

      acumulado.diretorias++;
    }

    var tempoSegFinal = Math.round((Date.now() - totalStart) / 1000);
    renderFinal('bfs-progresso', acumulado, tempoSegFinal, state.cancelado);
  }

  function cancelar() {
    state.cancelado = true;
  }

  window.contratosBfs = { iniciar: iniciar, cancelar: cancelar };
})();
