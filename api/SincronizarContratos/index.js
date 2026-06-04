/**
 * /api/SincronizarContratos — varre o SharePoint de contratos, extrai vigencia
 * via Claude, e persiste em PRONEP-NF-Contratos.
 *
 * RBAC: admin only (operacao pesada).
 *
 * Query/Body params:
 *   - pasta (string): subpath relativo dentro de "/CONTRATOS/CONTRATOS E DOCUMENTOS - PRESTADORES".
 *                     Exemplo: "/GERENCIA DE PROJETOS E TI/CORPORATIVO"
 *                     Default: raiz inteira.
 *   - dryRun (bool):  se true, NAO persiste. Apenas mostra o que seria gravado.
 *   - recursivo (bool): default true. Se false, processa apenas a pasta indicada (sem subpastas).
 *   - forcarReleitura (bool): default false. Se true, processa arquivos ja na lista.
 *   - maxArquivos (int): default 50. Limite de seguranca por chamada.
 *   - garantirLista (bool): default true. Cria a lista PRONEP-NF-Contratos se nao existir.
 *
 * Retorna:
 *   {
 *     ok: true,
 *     pasta: '...',
 *     totalEncontrados: 12,
 *     processados: 12,
 *     novos: 8,
 *     atualizados: 0,
 *     pulados: 4,
 *     erros: [],
 *     dryRun: false,
 *     listaCriada: false,
 *     resultados: [ { nome, path, diretoria, unidade, fornecedor, vigencia, status, persistido } ]
 *   }
 */

// IMPORTANTE: requires sao lazy (dentro do handler) pra evitar crash no carregamento
// caso algum modulo grande quebre no startup (pdf-parse tem historia de bug com
// arquivo de teste interno em algumas configs).
let contratos, getUser;
function carregarDeps() {
  if (!contratos) contratos = require('../shared/contratos');
  if (!getUser) getUser = require('../shared/auth').getUser;
}

function readClientPrincipal(req) {
  const header = req.headers && req.headers['x-ms-client-principal'];
  if (!header) return null;
  try { return JSON.parse(Buffer.from(header, 'base64').toString('utf-8')); } catch (e) { return null; }
}

function readClientPrincipalRoles(req) {
  const p = readClientPrincipal(req);
  return (p && p.userRoles) || [];
}

async function isAdmin(req) {
  const roles = readClientPrincipalRoles(req);
  if (roles.includes('administrador') || roles.includes('admin')) return true;
  // Fallback: getUserRoles (consulta grupos AAD)
  try {
    carregarDeps();
    const user = await getUser(req);
    if (!user || !user.oid) return false;
    const { getUserRoles } = require('../shared/userRoles');
    const userRoles = await getUserRoles(user);
    return (userRoles || []).includes('administrador');
  } catch (e) {
    return false;
  }
}

function ctxErr(context, status, msg, extra) {
  context.res = {
    status: status,
    headers: { 'Content-Type': 'application/json' },
    body: Object.assign({ error: msg }, extra || {})
  };
}

module.exports = async function (context, req) {
  const inicio = Date.now();
  // Carga lazy ANTES do try pra retornar erro especifico se shared/contratos falhar
  try {
    carregarDeps();
  } catch (eDep) {
    context.res = { status: 500, headers: { 'Content-Type': 'application/json' },
      body: { error: 'falha ao carregar deps: ' + eDep.message, stack: (eDep.stack || '').split('\n').slice(0, 8) } };
    return;
  }
  try {
    // RBAC
    if (!(await isAdmin(req))) {
      return ctxErr(context, 403, 'Apenas admin pode rodar a sincronizacao');
    }

    const params = Object.assign({}, req.query || {}, req.body || {});
    const pastaRel = params.pasta || '';
    const dryRun = String(params.dryRun || '') === 'true' || params.dryRun === true;
    const recursivo = !(String(params.recursivo || '') === 'false' || params.recursivo === false);
    const forcarReleitura = String(params.forcarReleitura || '') === 'true' || params.forcarReleitura === true;
    const maxArquivos = Math.min(parseInt(params.maxArquivos || '50') || 50, 200);
    const garantirLista = !(String(params.garantirLista || '') === 'false' || params.garantirLista === false);

    const pastaCompleta = (contratos.ROOT_FOLDER_PATH +
                          (pastaRel ? (pastaRel.startsWith('/') ? pastaRel : '/' + pastaRel) : ''))
                         .replace(/\/+$/, '');

    const client = contratos.getGraphClient();

    // 1. Garante lista PRONEP-NF-Contratos (no site do sistema NF)
    let siteNF, listIdContratos, listaCriada = false;
    if (garantirLista) {
      try {
        const r = await contratos.garantirListaContratos(client);
        siteNF = r.siteId;
        listIdContratos = r.listId;
        listaCriada = r.criada;
      } catch (eLista) {
        return ctxErr(context, 500, 'Falha ao garantir lista PRONEP-NF-Contratos: ' + eLista.message, {
          step: 'garantirLista',
          graphStatusCode: eLista.statusCode,
          graphBody: eLista.body,
          graphCode: eLista.code,
          stack: (eLista.stack || '').split('\n').slice(0, 8)
        });
      }
    }

    // 2. Resolve site de Contratos + driveId
    const { contratoSite, driveId } = await contratos.resolveContratosSite(client);

    // 3. Lista arquivos da pasta alvo
    const progress = [];
    let arquivos;
    if (recursivo) {
      // Crawl COMPLETO da arvore (sem aplicar maxArquivos aqui) — assim NUNCA
      // perdemos pastas porque o limite foi atingido na fase de listagem.
      // O `maxArquivos` se aplica somente aos NAO-EXISTENTES no loop de processamento.
      arquivos = await contratos.crawlPasta(client, driveId, pastaCompleta, {
        maxDepth: 6,
        maxArquivos: 5000,
        onProgress: function(ev){ if (progress.length < 50) progress.push(ev); }
      });
    } else {
      const listing = await contratos.listarPasta(client, driveId, pastaCompleta);
      arquivos = listing.files.map(function(f){
        return Object.assign({}, f, { ancestors: [pastaCompleta.split('/').pop()] });
      });
    }

    const totalEncontrados = arquivos.length;
    if (totalEncontrados === 0) {
      context.res = {
        status: 200,
        headers: { 'Content-Type': 'application/json' },
        body: {
          ok: true,
          pasta: pastaCompleta,
          totalEncontrados: 0,
          mensagem: 'Nenhum PDF/DOCX encontrado nessa pasta',
          progress
        }
      };
      return;
    }

    // 4. Carrega lista atual pra dedup (busca por DriveItemId)
    let colMap = {};
    let itensExistentes = {};
    if (!dryRun && garantirLista) {
      colMap = await contratos.getContratoColMap(client, siteNF, listIdContratos);
      const existentes = await client.api('/sites/' + siteNF + '/lists/' + listIdContratos + '/items?expand=fields&$top=999').get();
      const colDriveId = colMap['DriveItemId'] || 'DriveItemId';
      for (const it of (existentes.value || [])) {
        const f = it.fields || {};
        const did = f[colDriveId] || f.DriveItemId || '';
        if (did) itensExistentes[did] = it.id;
      }
    }

    // 5. Filtra: separa arquivos a processar (novos OU forcar) dos pulados (ja gravados)
    const stats = { novos: 0, atualizados: 0, pulados: 0, erros: 0 };
    const resultados = [];
    let aFiltrados;
    if (forcarReleitura) {
      aFiltrados = arquivos;
    } else {
      // Conta pulados (ja existentes) ANTES de cortar pra maxArquivos
      aFiltrados = arquivos.filter(function(a){ return !itensExistentes[a.id]; });
      const jaExistentes = arquivos.length - aFiltrados.length;
      stats.pulados = jaExistentes;
      // Adiciona os pulados aos resultados pra rastreabilidade (sem dados pesados)
      for (const arq of arquivos) {
        if (itensExistentes[arq.id]) {
          resultados.push({
            nome: arq.nome,
            path: arq.path,
            persistido: false,
            motivo: 'ja_existe',
            spItemId: itensExistentes[arq.id]
          });
        }
      }
    }

    // 6. Limita aos NAO-EXISTENTES por maxArquivos
    const aProcessar = aFiltrados.slice(0, maxArquivos);
    const restantes = aFiltrados.length - aProcessar.length;

    // 7. Processa cada arquivo
    for (const arq of aProcessar) {
      try {

        // Classifica path -> diretoria/unidade/fornecedor
        const classif = contratos.classificarPath(arq.ancestors);

        // Extrai texto
        const ext = await contratos.extrairTexto(client, driveId, arq.id, arq.ext);
        if (ext.vazio && !ext.erro) {
          // PDF scaneado provavelmente — sinaliza
          resultados.push({
            nome: arq.nome,
            path: arq.path,
            diretoria: classif.diretoria,
            unidade: classif.unidade,
            fornecedor: classif.fornecedor,
            vigencia: { naoEncontrou: true, motivo: 'PDF sem texto (provavelmente scaneado, precisa OCR)' },
            status: 'SemVigencia',
            persistido: false,
            obsLeitura: 'pdf_scaneado'
          });
          if (!dryRun && garantirLista) {
            await persistir(client, siteNF, listIdContratos, colMap, arq, classif, { naoEncontrou: true, motivo: 'PDF scaneado' }, itensExistentes[arq.id]);
            stats.novos++;
          }
          continue;
        }

        // Extrai vigencia via Claude
        const vig = await contratos.extrairVigenciaInteligente(ext.texto);
        const status = contratos.calcularStatus(vig.dataFim, vig.indeterminado);

        resultados.push({
          nome: arq.nome,
          path: arq.path,
          diretoria: classif.diretoria,
          unidade: classif.unidade,
          fornecedor: classif.fornecedor,
          vigencia: vig,
          status: status,
          persistido: !dryRun,
          tokensIn: vig._tokensIn || 0,
          tokensOut: vig._tokensOut || 0,
          modelo: vig._modelo
        });

        if (!dryRun && garantirLista) {
          await persistir(client, siteNF, listIdContratos, colMap, arq, classif, vig, itensExistentes[arq.id]);
          if (itensExistentes[arq.id]) stats.atualizados++; else stats.novos++;
        }
      } catch (e) {
        // Caso especial: arquivo muito grande — grava entry vazia pra usuario
        // ver no sistema e editar manualmente, em vez de simplesmente "errar".
        if (e.message && e.message.indexOf('arquivo_grande_demais') === 0) {
          try {
            const classif = contratos.classificarPath(arq.ancestors);
            const vig = { naoEncontrou: true, motivo: 'arquivo grande demais (' + (arq.size ? Math.round(arq.size/1024/1024) + 'MB' : '?') + ') - extracao manual necessaria' };
            resultados.push({
              nome: arq.nome,
              path: arq.path,
              diretoria: classif.diretoria,
              unidade: classif.unidade,
              fornecedor: classif.fornecedor,
              vigencia: vig,
              status: 'SemVigencia',
              persistido: !dryRun,
              obsLeitura: 'arquivo_grande'
            });
            if (!dryRun && garantirLista) {
              await persistir(client, siteNF, listIdContratos, colMap, arq, classif, vig, itensExistentes[arq.id]);
              if (itensExistentes[arq.id]) stats.atualizados++; else stats.novos++;
            }
          } catch (e2) {
            stats.erros++;
            resultados.push({ nome: arq.nome, path: arq.path, erro: 'arq grande + falha persist: ' + e2.message });
          }
        } else {
          stats.erros++;
          resultados.push({
            nome: arq.nome,
            path: arq.path,
            erro: e.message
          });
        }
      }
    }

    context.res = {
      status: 200,
      headers: { 'Content-Type': 'application/json' },
      body: {
        ok: true,
        pasta: pastaCompleta,
        totalEncontrados,
        processados: aProcessar.length,
        restantes,
        stats,
        dryRun,
        recursivo,
        listaCriada,
        tempoMs: Date.now() - inicio,
        resultados,
        progress: progress.slice(0, 20)
      }
    };
  } catch (err) {
    context.log && context.log.error && context.log.error('SincronizarContratos error:', err);
    ctxErr(context, 500, err.message, { stack: (err.stack || '').split('\n').slice(0, 8) });
  }
};

async function persistir(client, siteId, listId, colMap, arq, classif, vig, spItemIdExistente) {
  const fields = {};
  const set = function(displayName, val) {
    if (val === undefined || val === null || val === '') return;
    const internal = colMap[displayName] || displayName;
    fields[internal] = val;
  };
  // Title = nome do arquivo (truncado)
  set('Title', (arq.nome || '').slice(0, 255));
  set('Diretoria', classif.diretoria);
  set('Unidade', classif.unidade);
  set('Fornecedor', classif.fornecedor || vig.fornecedorIdentificado || '');
  set('NomeArquivo', arq.nome);
  set('TamanhoArquivo', arq.size || 0);
  set('PathRelativoSP', arq.path);
  set('DriveItemId', arq.id);
  set('CaminhoSharepoint', arq.webUrl);
  set('UltimaLeitura', new Date().toISOString());

  // Vigencia
  if (vig.dataInicio) set('DataInicio', vig.dataInicio + 'T00:00:00Z');
  if (vig.dataFim) set('DataFim', vig.dataFim + 'T00:00:00Z');
  if (vig.valorContrato) set('ValorContrato', vig.valorContrato);

  // Status calculado
  set('Status', require('../shared/contratos').calcularStatus(vig.dataFim, vig.indeterminado));

  // Diagnostico da leitura IA
  let leituraStatus = 'auto_alto';
  if (vig.naoEncontrou) leituraStatus = 'nao_encontrou';
  else if (vig.confidence === 'baixo') leituraStatus = 'auto_baixo';
  else if (vig.indeterminado) leituraStatus = 'indeterminado';
  set('LeituraIAStatus', leituraStatus);

  const trechos = [];
  if (vig.trecho) trechos.push(vig.trecho);
  if (vig.motivo) trechos.push('Motivo: ' + vig.motivo);
  if (vig._modelo) trechos.push('Modelo: ' + vig._modelo);
  if (vig._escalacao) trechos.push('Escalado pra Sonnet apos Haiku reportar baixo confidence');
  set('LeituraIATexto', trechos.join('\n').slice(0, 30000));

  if (spItemIdExistente) {
    await client.api('/sites/' + siteId + '/lists/' + listId + '/items/' + spItemIdExistente + '/fields')
      .patch(fields);
  } else {
    await client.api('/sites/' + siteId + '/lists/' + listId + '/items')
      .post({ fields: fields });
  }
}
vo,
        listaCriada,
        tempoMs: Date.now() - inicio,
        resultados,
        progress: progress.slice(0, 20)
      }
    };
  } catch (err) {
    context.log && context.log.error && context.log.error('SincronizarContratos error:', err);
    ctxErr(context, 500, err.message, { stack: (err.stack || '').split('\n').slice(0, 8) });
  }
};
;
  }
};
