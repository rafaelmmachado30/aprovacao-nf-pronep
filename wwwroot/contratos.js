/* contratos.js — funcoes da tela de Contratos, extraidas do index.html (P1.3).
   Carregado como <script src> ANTES do script inline (igual contratos-bfs.js).
   Sao funcoes GLOBAIS (onclick depende disso). Dependem de globais do index em
   runtime: escHtml, openModal, closeModal, toast, carregarContratosReais,
   confirmarAcao, mostrarView, window.contratosBfs. */

function renderContratosView() {
  const stats = window._contratosStats || {};
  const arvore = window._contratosArvore || [];
  const scope = window._contratosScope || {};
  const f = window._filtroContratos;

  // KPIs
  const kpiCards = [
    { label: 'Total', v: stats.total || 0, cor: 'var(--azul-escuro)' },
    { label: 'Ativos', v: stats.ativos || 0, cor: 'var(--verde)' },
    { label: 'Vencendo 90d', v: stats.vencendo90 || 0, cor: '#3B82F6' },
    { label: 'Vencendo 60d', v: stats.vencendo60 || 0, cor: '#F59E0B' },
    { label: 'Vencendo 30d', v: stats.vencendo30 || 0, cor: '#DC2626' },
    { label: 'Vencidos', v: stats.vencidos || 0, cor: '#7F1D1D' },
    { label: 'Sem vigência', v: stats.semVigencia || 0, cor: '#6B7280' }
  ];
  const kpiHTML = kpiCards.map(k => `
    <div class="kpi-card" style="border-left:4px solid ${k.cor};padding:12px;background:#fff;border-radius:6px;flex:1;min-width:120px;box-shadow:0 1px 3px rgba(0,0,0,.06)">
      <div class="small muted" style="font-size:11px;text-transform:uppercase">${k.label}</div>
      <div style="font-size:24px;font-weight:700;color:${k.cor}">${k.v}</div>
    </div>
  `).join('');

  const escopo = Array.isArray(scope.diretoriasGestorOf) ? ('Diretorias: ' + scope.diretoriasGestorOf.join(', ')) : 'Acesso total (Admin)';

  return `
    <h2>Contratos</h2>
    <div class="view-sub">Acervo de contratos vigentes e historicos da Pronep. Vigencia extraida automaticamente por IA. <b>${escopo}</b>.</div>

    <div style="display:flex;gap:8px;flex-wrap:wrap;margin-bottom:14px">
      ${kpiHTML}
    </div>

    <div class="filter-bar filter-stack-mobile" style="display:flex;gap:12px;align-items:end;flex-wrap:wrap;margin-bottom:12px">
      <div style="flex:1;min-width:200px">
        <div class="label">Buscar (fornecedor, nome do arquivo)</div>
        <input id="fct-busca" type="text" value="${(f.busca || '').replace(/"/g,'&quot;')}" placeholder="Ex: TOTVS, Aile..." oninput="aplicaFiltroContratos()" style="width:100%">
      </div>
      <div style="min-width:140px">
        <div class="label">Status</div>
        <select id="fct-status" onchange="aplicaFiltroContratos()">
          <option value="TODOS"${f.status==='TODOS'?' selected':''}>Todos</option>
          <option value="Ativo"${f.status==='Ativo'?' selected':''}>Ativo</option>
          <option value="Vencendo30"${f.status==='Vencendo30'?' selected':''}>Vencendo 30d</option>
          <option value="Vencendo60"${f.status==='Vencendo60'?' selected':''}>Vencendo 60d</option>
          <option value="Vencendo90"${f.status==='Vencendo90'?' selected':''}>Vencendo 90d</option>
          <option value="Vencido"${f.status==='Vencido'?' selected':''}>Vencido</option>
          <option value="Cancelado"${f.status==='Cancelado'?' selected':''}>Cancelado</option>
          <option value="SemVigencia"${f.status==='SemVigencia'?' selected':''}>Sem vigência</option>
        </select>
      </div>
      <div style="display:flex;gap:6px">
        <button class="btn btn-secondary" onclick="abrirSincronizarContratos()">Sincronizar pasta</button>
        <button class="btn btn-primary" onclick="abrirSincronizarTudo()" title="Varre as 10 diretorias da Pronep em sequencia">⚡ Sincronizar TUDO</button>
        ${(window._contratosScope && window._contratosScope.isAdmin) ? '<button class="btn btn-secondary" onclick="limparDuplicatasContratos()" title="Remove contratos duplicados (mesmo arquivo repetido), mantendo 1 de cada">🧹 Limpar duplicatas</button>' : ''}
        ${(window._contratosScope && window._contratosScope.isAdmin) ? '<button class="btn btn-secondary" onclick="indexarComercialRAG(this)" title="Cria a base de conhecimento (RAG) dos contratos Comerciais para a IA consultar">🧠 Indexar Comercial (IA)</button>' : ''}
        ${(window._contratosScope && window._contratosScope.isAdmin) ? '<button class="btn btn-secondary" onclick="curarPilotoContratos(this)" title="Ingestão curada: lê os contratos Comerciais do índice, extrai o modelo canônico e grava a fonte única (Markdown+YAML)">📚 Curar contratos (IA)</button>' : ''}
        ${(window._contratosScope && window._contratosScope.isAdmin) ? '<button class="btn btn-secondary" onclick="carregarBaseRelacional(this)" title="Deriva o banco relacional (Supabase) a partir da fonte única curada — trilho de números/agregações/temporal">🗄️ Carregar base relacional (Supabase)</button>' : ''}
      </div>
      <div id="rag-index-progress" class="small muted" style="margin-top:8px"></div>
      <div id="curado-progress" class="small muted" style="margin-top:8px"></div>
      <div id="relacional-progress" class="small muted" style="margin-top:8px"></div>
    </div>

    <div id="contratos-arvore">
      ${renderArvoreContratosV2(arvore)}
    </div>
  `;
}

// ===== Arvore de contratos ESPELHANDO o SharePoint (aninhada por subpasta real) =====
// Usa o caminho salvo (pathSP) pra reconstruir a hierarquia: Diretoria -> Unidade ->
// Prestador -> VIGENTE/HISTORICO -> arquivos. Cada contrato dentro da sua pasta.
function _ctNorm(s) { return String(s || '').toLowerCase().normalize('NFD').replace(/\p{Diacritic}/gu, '').replace(/[^a-z0-9]/g, ''); }
function _ctSubpastas(c) {
  const ROOTN = 'contratosedocumentosprestadores';
  let segs = String(c.pathSP || '').split('/').filter(Boolean);
  const i = segs.findIndex(function (s) { return _ctNorm(s) === ROOTN; });
  if (i >= 0) segs = segs.slice(i + 1);
  if (segs.length) segs = segs.slice(0, -1); // remove o nome do arquivo
  if (segs.length) segs = segs.slice(1);     // remove a diretoria-raw (o topo ja e a diretoria)
  return segs;
}
function _ctContar(node) { let n = node.contratos.length; for (const k in node.filhos) n += _ctContar(node.filhos[k]); return n; }
function _ctRenderPasta(node, idPath) {
  const keys = Object.keys(node.filhos).sort(function (a, b) { return a.localeCompare(b, 'pt'); });
  let html = '';
  for (const k of keys) {
    const ch = node.filhos[k];
    const key = idPath + '|' + k;
    const id = 'ctp-' + key.replace(/[^A-Za-z0-9|]/g, '_');
    html += '<details id="' + id + '" data-det-key="' + escAttr(key) + '" style="margin-top:6px">' +
      '<summary style="padding:6px 10px;cursor:pointer;font-weight:500;color:#374151;background:#F1F5F9;border-radius:4px">' +
      iconePasta('#3B82F6') + ' ' + escHtml(k) + ' <span class="small muted">(' + _ctContar(ch) + ')</span></summary>' +
      '<div style="padding:6px 6px 6px 14px">' + _ctRenderPasta(ch, key) + '</div>' +
    '</details>';
  }
  if (node.contratos.length) {
    html += '<div style="padding:6px 0;display:flex;flex-direction:column;gap:6px">' +
      node.contratos.map(function (c) { return renderCardContrato(c); }).join('') + '</div>';
  }
  return html;
}
function renderArvoreContratosV2(arvore) {
  if (!arvore || !arvore.length) {
    return '<div class="card" style="text-align:center;padding:30px"><div class="muted">Nenhum contrato encontrado com esses filtros.</div></div>';
  }
  return arvore.map(function (dir) {
    const todos = [];
    (dir.unidades || []).forEach(function (u) { (u.fornecedores || []).forEach(function (fr) { (fr.contratos || []).forEach(function (c) { todos.push(c); }); }); });
    const root = { nome: dir.diretoria, contratos: [], filhos: {} };
    todos.forEach(function (c) {
      let node = root;
      _ctSubpastas(c).forEach(function (seg) { if (!node.filhos[seg]) node.filhos[seg] = { nome: seg, contratos: [], filhos: {} }; node = node.filhos[seg]; });
      node.contratos.push(c);
    });
    const idDir = 'ctp-' + String(dir.diretoria).replace(/[^A-Za-z0-9]/g, '_');
    return '<details id="' + idDir + '" data-det-key="' + escAttr(dir.diretoria) + '" style="background:#fff;border:1px solid var(--cinza-borda);border-radius:6px;margin-bottom:10px">' +
      '<summary style="padding:12px 14px;cursor:pointer;font-weight:600;color:var(--azul-escuro);background:#F8FAFC">' +
      iconePasta('#3B82F6') + ' ' + escHtml(dir.diretoria) + ' <span class="small muted">(' + (dir.total || todos.length) + ')</span></summary>' +
      '<div style="padding:8px 14px 14px">' + _ctRenderPasta(root, dir.diretoria) + '</div>' +
    '</details>';
  }).join('');
}

// Limpa contratos duplicados (admin) — dry-run + loop em lotes com barra de progresso.
async function limparDuplicatasContratos() {
  var dry;
  try {
    const rr = await fetch('/api/LimparContratosDuplicados', { credentials: 'include' });
    dry = await rr.json();
    if (!rr.ok || !dry.ok) { toast('Falha ao verificar duplicatas: ' + (dry && dry.error || ('HTTP ' + rr.status)), 'error'); return; }
  } catch (e) { toast('Erro de rede: ' + (e.message || e), 'error'); return; }

  if (!dry.duplicatasIdentificadas) { toast('Nenhuma duplicata encontrada. ✓', 'success'); return; }

  const ok = await confirmarAcao({
    titulo: 'Limpar contratos duplicados?',
    mensagem: 'Encontrei <b>' + dry.duplicatasIdentificadas + '</b> contratos duplicados (de ' + dry.totalItens + ' no total).\n\n' +
      'Vou remover os repetidos, mantendo <b>1 de cada</b> arquivo (o mais recente, preservando edições). Isso não pode ser desfeito.',
    tipo: 'danger', confirmLabel: 'Sim, limpar', cancelLabel: 'Cancelar'
  });
  if (!ok) return;

  const total = dry.duplicatasIdentificadas;
  openModal(
    '<div style="text-align:center;padding:22px 14px">' +
    '<div style="display:inline-block;width:46px;height:46px;border:5px solid #E5E7EB;border-top-color:#1F4E79;border-radius:50%;animation:spinLD .8s linear infinite;margin-bottom:14px"></div>' +
    '<h3 style="margin:0 0 6px;color:#1F4E79;font-size:18px">Limpando duplicatas...</h3>' +
    '<div style="height:10px;background:#E5E7EB;border-radius:5px;overflow:hidden;margin:14px 0 8px"><div id="ld-bar" style="height:100%;width:0%;background:#1F4E79;transition:width .3s"></div></div>' +
    '<div id="ld-status" style="font-size:13px;color:#374151">Iniciando...</div>' +
    '</div><style>@keyframes spinLD{to{transform:rotate(360deg)}}</style>'
  );

  let removidosAcum = 0;
  let guarda = 0;
  let erroFinal = null;
  while (guarda++ < 200) {
    let r;
    try {
      const resp = await fetch('/api/LimparContratosDuplicados?aplicar=true', { credentials: 'include' });
      r = await resp.json();
      if (!resp.ok || !r.ok) { erroFinal = (r && r.error) || ('HTTP ' + resp.status); break; }
    } catch (e) { erroFinal = e.message || String(e); break; }
    removidosAcum += (r.removidos || 0);
    const pct = total > 0 ? Math.min(100, Math.round(removidosAcum / total * 100)) : 100;
    const bar = document.getElementById('ld-bar'); if (bar) bar.style.width = pct + '%';
    const st = document.getElementById('ld-status'); if (st) st.textContent = 'Removidos ' + removidosAcum + ' de ' + total + ' (faltam ' + Math.max(0, (r.restantes || 0)) + ')';
    // Termina quando nao ha mais o que remover (ou o lote nao removeu nada — evita loop infinito)
    if ((r.restantes || 0) <= 0 || (r.removidos || 0) === 0) break;
  }

  try { closeModal(); } catch (e) {}
  if (erroFinal) { toast('Limpeza interrompida: ' + erroFinal + ' (removidos ' + removidosAcum + '). Tente de novo.', 'error'); }
  else { toast('Limpeza concluída: ' + removidosAcum + ' duplicata(s) removida(s). ✓', 'success'); }
  // Recarrega a lista de contratos pra refletir a limpeza
  window._contratosCarregados = false;
  if (typeof carregarContratosReais === 'function') carregarContratosReais();
}

function renderArvoreContratos(arvore) {
  if (!arvore || !arvore.length) {
    return '<div class="card" style="text-align:center;padding:30px"><div class="muted">Nenhum contrato encontrado com esses filtros.</div></div>';
  }
  // ID estavel pra detalhes: serve pra preservar estado open/closed entre re-renders
  function detId(parts) { return 'det-' + parts.join('|').replace(/[^A-Za-z0-9|]/g, '_'); }
  return arvore.map(function(dir) {
    var idDir = detId([dir.diretoria]);
    return `
      <details id="${idDir}" data-det-key="${dir.diretoria}" style="background:#fff;border:1px solid var(--cinza-borda);border-radius:6px;margin-bottom:10px">
        <summary style="padding:12px 14px;cursor:pointer;font-weight:600;color:var(--azul-escuro);background:#F8FAFC">
          ${iconePasta('#3B82F6')} ${dir.diretoria} <span class="small muted">(${dir.total})</span>
        </summary>
        <div style="padding:8px 14px 14px">
          ${(dir.unidades || []).map(function(un){
            var idUn = detId([dir.diretoria, un.unidade]);
            return `
              <details id="${idUn}" data-det-key="${dir.diretoria}|${un.unidade}" style="margin-top:8px">
                <summary style="padding:6px 10px;cursor:pointer;font-weight:500;color:#374151;background:#F1F5F9;border-radius:4px">
                  ${iconePasta('#3B82F6')} ${un.unidade} <span class="small muted">(${un.total})</span>
                </summary>
                <div style="padding:8px 6px 6px;display:flex;flex-direction:column;gap:8px">
                  ${(un.fornecedores || []).map(function(forn){
                    var contratos = forn.contratos || [];
                    var qtdCancelados = contratos.filter(function(c){ return c.status === 'Cancelado'; }).length;
                    var qtdTotal = contratos.length;
                    var todosCancelados = qtdTotal > 0 && qtdCancelados === qtdTotal;
                    var parcial = qtdCancelados > 0 && qtdCancelados < qtdTotal;
                    // Cor da pasta + badge: VERDE (todos ativos), AMARELO (parcial), VERMELHO (todos cancelados)
                    var corPasta, statusLabel, corStatus;
                    if (todosCancelados) {
                      corPasta = '#DC2626'; statusLabel = 'Cancelado'; corStatus = '#DC2626';
                    } else if (parcial) {
                      corPasta = '#F59E0B'; statusLabel = 'Ativo Parcialmente ('+(qtdTotal-qtdCancelados)+'/'+qtdTotal+')'; corStatus = '#F59E0B';
                    } else {
                      corPasta = '#3B82F6'; statusLabel = 'Ativo'; corStatus = '#10B981';
                    }
                    // Acao do botao: se TODOS cancelados → Reativar todos. Caso contrario (ativo OU parcial) → Cancelar todos.
                    var acaoBatch = todosCancelados ? 'reativar' : 'cancelar';
                    var labelBatch = todosCancelados ? '↻ Reativar todos' : '✕ Cancelar todos';
                    var corBtn = todosCancelados ? '#10B981' : '#DC2626';
                    // Encoda nome do fornecedor pra atributo HTML (escape de aspas)
                    var fEsc = String(forn.fornecedor).replace(/"/g,'&quot;').replace(/&/g,'&amp;');
                    var dEsc = String(dir.diretoria).replace(/"/g,'&quot;').replace(/&/g,'&amp;');
                    var uEsc = String(un.unidade).replace(/"/g,'&quot;').replace(/&/g,'&amp;');
                    var idForn = detId([dir.diretoria, un.unidade, forn.fornecedor]);
                    var fornKey = dir.diretoria + '|' + un.unidade + '|' + forn.fornecedor;
                    return `
                      <details id="${idForn}" data-det-key="${fornKey.replace(/"/g,'&quot;')}" style="padding:0;background:#FAFAFA;border-radius:4px">
                        <summary style="padding:8px 12px;cursor:pointer;font-weight:600;color:#1F2937;list-style:none;display:flex;align-items:center;gap:6px">
                          <span style="font-size:11px;color:#6B7280;width:10px">▶</span>${iconePasta(corPasta)} ${escHtml(forn.fornecedor)} <span class="small muted">(${contratos.length})</span>
                          <span style="flex:1"></span>
                          <span style="background:${corStatus};color:#fff;padding:2px 10px;border-radius:12px;font-size:11px;font-weight:600">${statusLabel}</span>
                          <button type="button"
                            onclick="event.stopPropagation();event.preventDefault();cancelarFornecedorClick(this)"
                            data-diretoria="${dEsc}" data-unidade="${uEsc}" data-fornecedor="${fEsc}" data-acao="${acaoBatch}"
                            style="background:#fff;border:1px solid ${corBtn};color:${corBtn};padding:3px 10px;border-radius:4px;cursor:pointer;font-size:11px;font-weight:600">
                            ${labelBatch}
                          </button>
                        </summary>
                        <div style="padding:6px 10px 10px;display:flex;flex-direction:column;gap:6px">
                          ${contratos.map(function(c){ return renderCardContrato(c); }).join('')}
                        </div>
                      </details>
                    `;
                  }).join('')}
                </div>
              </details>
            `;
          }).join('')}
        </div>
      </details>
    `;
  }).join('');
}

// Ícone SVG de pasta colorida — substitui o emoji 📁/📍/🏢
function iconePasta(cor) {
  return '<svg style="display:inline-block;vertical-align:middle;width:18px;height:14px" viewBox="0 0 24 18" fill="' + cor + '">' +
    '<path d="M2 0h7l2 2h11a2 2 0 0 1 2 2v12a2 2 0 0 1-2 2H2a2 2 0 0 1-2-2V2a2 2 0 0 1 2-2z"/>' +
    '</svg>';
}

// Cancelar/Reativar TODOS os contratos de um fornecedor de uma vez.
// Disparado pelo botao "Cancelar todos" / "Reativar todos" no summary do fornecedor na arvore.
// Mostra modal de progresso, preserva estado open/closed da arvore, e da scroll de volta
// pra pasta tocada apos o reload.
async function cancelarFornecedorClick(btn) {
  var diretoria = btn.getAttribute('data-diretoria');
  var unidade = btn.getAttribute('data-unidade');
  var fornecedor = btn.getAttribute('data-fornecedor');
  var acao = btn.getAttribute('data-acao');
  var ehCancelar = acao === 'cancelar';
  var ok = await confirmarAcao({
    titulo: ehCancelar ? 'Cancelar todos os contratos?' : 'Reativar todos os contratos?',
    mensagem:
      (ehCancelar ? 'Você vai marcar como CANCELADOS' : 'Você vai REATIVAR') + ' todos os contratos de:\n\n' +
      '<b>' + fornecedor + '</b>\n' +
      diretoria + ' / ' + unidade + '\n\n' +
      'Essa ação afeta TODOS os documentos da pasta.',
    tipo: ehCancelar ? 'danger' : 'success',
    confirmLabel: ehCancelar ? 'Sim, cancelar todos' : 'Sim, reativar todos',
    cancelLabel: 'Cancelar'
  });
  if (!ok) return;

  // 1. Captura estado atual da arvore (quais <details> estao abertos) pra restaurar depois
  var detsAbertos = [];
  try {
    var todosDets = document.querySelectorAll('details[data-det-key]');
    for (var i = 0; i < todosDets.length; i++) {
      if (todosDets[i].open) detsAbertos.push(todosDets[i].getAttribute('data-det-key'));
    }
    // Garante que o caminho ate a pasta tocada vai ficar aberto
    var keyForn = diretoria + '|' + unidade + '|' + fornecedor;
    var keyUn = diretoria + '|' + unidade;
    var keyDir = diretoria;
    if (detsAbertos.indexOf(keyForn) < 0) detsAbertos.push(keyForn);
    if (detsAbertos.indexOf(keyUn)   < 0) detsAbertos.push(keyUn);
    if (detsAbertos.indexOf(keyDir)  < 0) detsAbertos.push(keyDir);
  } catch (e) { /* nao critico */ }

  // 2. Modal de progresso (spinner + texto)
  var cor = acao === 'cancelar' ? '#DC2626' : '#10B981';
  var verboIng = acao === 'cancelar' ? 'Cancelando' : 'Reativando';
  var html =
    '<div style="text-align:center;padding:24px 12px">' +
    '  <div style="margin-bottom:18px">' +
    '    <div style="display:inline-block;width:48px;height:48px;border:5px solid #E5E7EB;border-top-color:' + cor + ';border-radius:50%;animation:spinFornecedor 0.8s linear infinite"></div>' +
    '  </div>' +
    '  <h3 style="margin:0 0 8px;color:' + cor + ';font-size:18px">' + verboIng + ' contratos...</h3>' +
    '  <div style="color:#6B7280;font-size:14px;margin-bottom:4px"><b>' + fornecedor + '</b></div>' +
    '  <div style="color:#9CA3AF;font-size:12px">' + diretoria + ' / ' + unidade + '</div>' +
    '  <div id="cancelarFornStatus" style="margin-top:18px;padding:10px;background:#F9FAFB;border-radius:6px;font-size:13px;color:#374151">Processando lote no SharePoint...</div>' +
    '</div>' +
    '<style>@keyframes spinFornecedor { to { transform: rotate(360deg); } }</style>';
  if (typeof openModal === 'function') openModal(html);

  try {
    var r = await fetch('/api/CancelarFornecedor', {
      method: 'POST',
      credentials: 'include',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ diretoria: diretoria, unidade: unidade, fornecedor: fornecedor, acao: acao })
    });
    var j = await r.json();
    if (!r.ok) {
      if (typeof closeModal === 'function') closeModal();
      alert('Erro: ' + (j.error || ('HTTP ' + r.status)));
      return;
    }

    // Atualiza modal pra "concluido" antes de recarregar
    var statusEl = document.getElementById('cancelarFornStatus');
    if (statusEl) {
      var msgOk = acao === 'cancelar'
        ? ('✓ ' + j.atualizados + ' de ' + j.total + ' contratos cancelados')
        : ('✓ ' + j.atualizados + ' de ' + j.total + ' contratos reativados');
      statusEl.innerHTML = '<span style="color:#10B981;font-weight:600">' + msgOk + '</span><br><span style="color:#6B7280">Atualizando árvore...</span>';
    }

    // Recarrega a arvore preservando estado
    if (typeof carregarContratosReais === 'function') {
      await carregarContratosReais();
    }

    // 3. Re-expande os details que estavam abertos + a pasta tocada
    try {
      for (var k = 0; k < detsAbertos.length; k++) {
        var dets = document.querySelectorAll('details[data-det-key="' + detsAbertos[k].replace(/"/g,'\\"') + '"]');
        for (var d = 0; d < dets.length; d++) dets[d].open = true;
      }
      // Scroll suave ate a pasta tocada
      var alvoSel = 'details[data-det-key="' + (diretoria + '|' + unidade + '|' + fornecedor).replace(/"/g,'\\"') + '"]';
      var alvo = document.querySelector(alvoSel);
      if (alvo) {
        alvo.scrollIntoView({ behavior: 'smooth', block: 'center' });
        // Pisca a pasta tocada pra ajudar a localizar
        alvo.style.transition = 'background-color 1.4s';
        var corOriginal = alvo.style.backgroundColor;
        alvo.style.backgroundColor = acao === 'cancelar' ? '#FEE2E2' : '#D1FAE5';
        setTimeout(function(){ alvo.style.backgroundColor = corOriginal; }, 1400);
      }
    } catch (e) { /* nao critico */ }

    if (typeof closeModal === 'function') closeModal();
    if (typeof toast === 'function') {
      toast(acao === 'cancelar'
        ? ('Cancelados ' + j.atualizados + ' de ' + j.total + ' contratos')
        : ('Reativados ' + j.atualizados + ' de ' + j.total + ' contratos'), 'sucesso');
    }
  } catch (e) {
    if (typeof closeModal === 'function') closeModal();
    alert('Falha: ' + e.message);
  }
}

function renderCardContrato(c) {
  const corStatus = {
    'Ativo':        '#10B981',
    'Vencendo90':   '#3B82F6',
    'Vencendo60':   '#F59E0B',
    'Vencendo30':   '#DC2626',
    'Vencido':      '#7F1D1D',
    'Cancelado':    '#DC2626',
    'SemVigencia':  '#9CA3AF',
    'Indeterminado':'#10B981'
  };
  const cor = corStatus[c.status] || '#6B7280';
  const vigenciaStr = c.dataInicio || c.dataFim
    ? (c.dataInicio || '?') + ' → ' + (c.dataFim || (c.status === 'Cancelado' ? 'Cancelado' : 'Indeterminado'))
    : (c.status === 'Cancelado' ? 'Contrato cancelado' : 'Sem vigência identificada');
  const dias = c.diasParaVencer;
  const diasTxt = (dias !== null && dias !== undefined && c.status !== 'Cancelado')
    ? (dias < 0 ? ' (vencido há ' + Math.abs(dias) + 'd)' : ' (' + dias + 'd pra vencer)')
    : '';
  const valor = c.valor ? ' · ' + Number(c.valor).toLocaleString('pt-BR', {style:'currency',currency:'BRL'}) : '';
  const fornecImg = (c.fornecedor || c.title || '').slice(0, 50);
  const isCancelado = c.status === 'Cancelado';

  return `
    <div style="padding:10px;background:#fff;border:1px solid #E5E7EB;border-left:4px solid ${cor};border-radius:4px">
      <div style="display:flex;justify-content:space-between;align-items:flex-start;gap:10px;flex-wrap:wrap">
        <div style="flex:1;min-width:240px">
          <div style="font-weight:600;color:#1F2937">${escHtml(c.title || c.nomeArquivo || 'Sem nome')}</div>
          <div class="small muted" style="margin-top:4px">${vigenciaStr}${diasTxt}${valor}</div>
        </div>
        <div style="display:flex;gap:6px;align-items:center;flex-wrap:wrap">
          <span style="background:${cor};color:#fff;padding:2px 10px;border-radius:12px;font-size:11px;font-weight:600">${c.status || 'SemVigencia'}</span>
          <button class="btn btn-ghost" onclick="abrirContratoPDF('${c.id}')" title="Ver PDF" style="padding:4px 8px;font-size:12px">📄 PDF</button>
          <button class="btn btn-ghost" onclick='editarVigenciaContrato(${escAttr(JSON.stringify({id:c.id, title:c.title||c.nomeArquivo, dataInicio:c.dataInicio, dataFim:c.dataFim, valor:c.valor, observacoes:c.observacoes||"", status:c.status}))})' title="Editar vigência" style="padding:4px 8px;font-size:12px">✎ Editar</button>
          ${isCancelado
            ? '<button class="btn btn-ghost" onclick="alterarStatusContrato(\'' + c.id + '\', \'Ativo\')" style="padding:4px 8px;font-size:12px">↻ Reativar</button>'
            : '<button class="btn btn-ghost" onclick="alterarStatusContrato(\'' + c.id + '\', \'Cancelado\')" style="padding:4px 8px;font-size:12px;color:#DC2626">✕ Cancelar</button>'
          }
        </div>
      </div>
    </div>
  `;
}

function aplicaFiltroContratos() {
  const f = window._filtroContratos;
  f.busca = document.getElementById('fct-busca').value;
  f.status = document.getElementById('fct-status').value;
  carregarContratosReais();
}

function abrirContratoPDF(id) {
  window.open('/api/AbrirContrato?id=' + encodeURIComponent(id), '_blank');
}

function editarVigenciaContrato(c) {
  const html = `
    <h3 style="margin-top:0">Editar vigência do contrato</h3>
    <div class="small muted" style="margin-bottom:14px">${(c.title || '').replace(/</g,'&lt;')}</div>
    <div style="margin-bottom:10px">
      <div class="label">Data de início (assinatura)</div>
      <input id="evc-di" type="date" value="${c.dataInicio || ''}" style="width:100%">
    </div>
    <div style="margin-bottom:10px">
      <div class="label">Data fim (deixe em branco se indeterminado)</div>
      <input id="evc-df" type="date" value="${c.dataFim || ''}" style="width:100%">
    </div>
    <div style="margin-bottom:10px">
      <div class="label">Valor do contrato (R$, opcional)</div>
      <input id="evc-valor" type="number" step="0.01" min="0" value="${c.valor != null ? c.valor : ''}" placeholder="ex: 20171.10" style="width:100%">
    </div>
    <div style="margin-bottom:10px">
      <div class="label">Observações (opcional)</div>
      <textarea id="evc-obs" rows="3" style="width:100%;resize:vertical">${(c.observacoes || '').replace(/</g,'&lt;')}</textarea>
    </div>
    <div class="small muted" style="margin-bottom:14px">Status será recalculado automaticamente com base na data fim. Se quiser forçar um status específico (ex: Cancelado), use os botões do card.</div>
    <div style="display:flex;gap:10px;justify-content:flex-end">
      <button class="btn btn-secondary" onclick="closeModal()">Cancelar</button>
      <button class="btn btn-primary" onclick="salvarVigenciaContrato('${c.id}')">Salvar</button>
    </div>
    <div id="evc-resultado" style="margin-top:14px"></div>
  `;
  openModal(html);
}

async function salvarVigenciaContrato(id) {
  const dataInicio = document.getElementById('evc-di').value || '';
  const dataFim = document.getElementById('evc-df').value || '';
  const observacoes = document.getElementById('evc-obs').value || '';
  const valorRaw = document.getElementById('evc-valor').value;
  const valor = valorRaw !== '' && !isNaN(parseFloat(valorRaw)) ? parseFloat(valorRaw) : null;
  const resDiv = document.getElementById('evc-resultado');
  resDiv.innerHTML = '<div class="muted small">Salvando...</div>';
  mostrarSalvandoOverlay('Salvando vigência do contrato...');
  try {
    const r = await fetch('/api/AtualizarStatusContrato', {
      method: 'POST',
      credentials: 'include',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        id: id,
        dataInicio: dataInicio || null,
        dataFim: dataFim || null,
        valor: valor,
        observacoes: observacoes,
        status: 'auto'
      })
    });
    const data = await r.json();
    if (!r.ok) {
      resDiv.innerHTML = '<div class="alert alert-error">' + (data.error || 'erro') + '</div>';
      return;
    }
    if (typeof toast === 'function') toast('Vigência atualizada — status: ' + data.status, 'success');
    closeModal();
    carregarContratosReais();
  } catch (e) {
    resDiv.innerHTML = '<div class="alert alert-error">' + e.message + '</div>';
  } finally {
    esconderSalvandoOverlay();
  }
}

async function alterarStatusContrato(id, novoStatus) {
  const acao = novoStatus === 'Cancelado' ? 'cancelar' : 'reativar';
  const ehCancel = acao === 'cancelar';
  const okStatus = await confirmarAcao({
    titulo: ehCancel ? 'Cancelar este contrato?' : 'Reativar este contrato?',
    mensagem: ehCancel
      ? 'Você vai marcar este contrato como <b>Cancelado</b>.\n\nO documento continuará no SharePoint, apenas a sinalização vai mudar.'
      : 'Você vai <b>reativar</b> este contrato.\n\nO status volta a ser calculado pela data de vigência.',
    tipo: ehCancel ? 'danger' : 'success',
    confirmLabel: ehCancel ? 'Sim, cancelar' : 'Sim, reativar',
    cancelLabel: 'Cancelar'
  });
  if (!okStatus) return;
  try {
    const r = await fetch('/api/AtualizarStatusContrato', {
      method: 'POST',
      credentials: 'include',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ id: id, status: novoStatus })
    });
    const data = await r.json();
    if (!r.ok) {
      if (typeof toast === 'function') toast('Erro: ' + (data.error || r.status), 'error');
      return;
    }
    if (typeof toast === 'function') toast('Contrato ' + (novoStatus === 'Cancelado' ? 'cancelado' : 'reativado'), 'success');
    carregarContratosReais();
  } catch (e) {
    if (typeof toast === 'function') toast('Erro: ' + e.message, 'error');
  }
}

// ============================================================================
// ORQUESTRADOR: SincronizarTudo — varre as 10 diretorias-raiz em sequencia
// Roda no front (escapa do timeout do SWA Functions) com progresso em tempo real.
// ============================================================================
const DIRETORIAS_PRONEP = [
  { id: '/OUVIDORIA',                                  label: 'Ouvidoria' },
  { id: '/QUALIDADE',                                  label: 'Qualidade' },
  { id: '/PACIENTES PARTICULARES',                     label: 'Pacientes Particulares' },
  { id: '/GERÊNCIA DE RH',                             label: 'RH' },
  { id: '/DIRETORIA COMERCIAL',                        label: 'Comercial' },
  { id: '/DIRETORIA FINANCEIRA',                       label: 'Financeira' },
  { id: '/DIRETORIA DE OPERAÇÕES',                     label: 'Operações' },
  { id: '/DIRETORIA DE SUPRIMENTOS E LOGÍSTICA',       label: 'Suprimentos' },
  { id: '/JURÍDICO',                                   label: 'Jurídico' },
  { id: '/GERÊNCIA DE PROJETOS E TI',                  label: 'Tecnologia' }
];

window._syncCancelado = false;
window._syncErros = [];

async function sincronizarTudoOrquestrado() {
  const okSync = await confirmarAcao({
    titulo: 'Sincronizar TODAS as diretorias?',
    mensagem:
      'O sistema vai varrer as diretorias da Pronep no SharePoint, ler todos os PDFs novos e extrair vigência via IA.\n\n' +
      '<b>Tempo:</b> 30 a 90 minutos\n' +
      '<b>Custo estimado:</b> $1.50 a $4 em tokens Claude\n\n' +
      'Recomendado rodar fora do horário de pico.',
    tipo: 'warning',
    confirmLabel: 'Sim, sincronizar',
    cancelLabel: 'Cancelar'
  });
  if (!okSync) return;

  window._syncCancelado = false;
  window._syncErros = [];
  const modal = document.getElementById('modal-content');
  if (modal) modal.scrollTop = 0;

  const totalStart = Date.now();
  const acumulado = { novos: 0, atualizados: 0, pulados: 0, erros: 0, batches: 0, diretorias: 0, totalEncontrados: 0 };
  const porDiretoria = [];

  const resDiv = document.getElementById('syc-resultado');
  function atualiza(html) { if (resDiv) resDiv.innerHTML = html; }

  function renderProgresso(idxDir, nomeDir, batchInfo) {
    const pct = Math.round((idxDir / DIRETORIAS_PRONEP.length) * 100);
    const tempoSeg = Math.round((Date.now() - totalStart) / 1000);
    atualiza(`
      <div class="alert" style="background:#EFF6FF;border-left:3px solid #3B82F6;color:#1E3A8A;padding:14px;border-radius:4px">
        <div style="font-weight:600;margin-bottom:6px">Sincronizando diretoria ${idxDir + 1}/${DIRETORIAS_PRONEP.length}: <span style="color:var(--azul-escuro)">${nomeDir}</span></div>
        <div style="height:10px;background:#DBEAFE;border-radius:5px;overflow:hidden;margin-bottom:8px">
          <div style="height:100%;width:${pct}%;background:#3B82F6;transition:width .3s"></div>
        </div>
        <div class="small muted">${batchInfo || ''}</div>
        <div class="small muted" style="margin-top:6px">Acumulado: <b>${acumulado.novos}</b> novos · ${acumulado.atualizados} atualizados · ${acumulado.pulados} pulados · ${acumulado.erros} erros · ${acumulado.batches} chamadas · ${tempoSeg}s</div>
        <div style="margin-top:10px"><button class="btn btn-ghost" style="font-size:12px;padding:4px 10px" onclick="window._syncCancelado = true">⏹ Cancelar (termina o batch atual)</button></div>
      </div>
    `);
  }

  for (let i = 0; i < DIRETORIAS_PRONEP.length; i++) {
    if (window._syncCancelado) break;
    const dir = DIRETORIAS_PRONEP[i];
    let tentativas = 0;
    let pastaCompletaMsg = '';

    while (true) {
      if (window._syncCancelado) break;
      tentativas++;
      renderProgresso(i, dir.label, `Batch ${tentativas} dessa diretoria... ${pastaCompletaMsg}`);
      acumulado.batches++;

      try {
        const params = new URLSearchParams();
        // IMPORTANTE: normaliza pra Unicode NFC (composto) — alguns paths com acentos
        // podem vir em NFD do JavaScript e o Graph API rejeita.
        params.set('pasta', dir.id.normalize('NFC'));
        params.set('maxArquivos', '30');
        const r = await fetch('/api/SincronizarContratos?' + params.toString(), { credentials: 'include' });
        const data = await r.json();
        if (!r.ok) {
          acumulado.erros++;
          window._syncErros.push({ diretoria: dir.label, pasta: dir.id, statusHttp: r.status, error: data.error || 'sem mensagem', step: data.step, graphStatusCode: data.graphStatusCode, graphBody: data.graphBody });
          pastaCompletaMsg = `<b style="color:#DC2626">Erro:</b> ${data.error || ('HTTP ' + r.status)}`;
          break;
        }
        const s = data.stats || {};
        acumulado.novos += (s.novos || 0);
        acumulado.atualizados += (s.atualizados || 0);
        acumulado.pulados += (s.pulados || 0);
        acumulado.erros += (s.erros || 0);
        acumulado.totalEncontrados += (data.totalEncontrados || 0);

        // Captura erros individuais de arquivos
        for (const r2 of (data.resultados || [])) {
          if (r2.erro) {
            window._syncErros.push({ diretoria: dir.label, pasta: dir.id, arquivo: r2.nome, erro: r2.erro });
          }
        }

        const restantes = data.restantes || 0;
        const totalEncBatch = data.totalEncontrados || 0;
        pastaCompletaMsg = `Batch ${tentativas}: ${totalEncBatch} encontrados na árvore, ${s.novos || 0} novos, ${s.pulados || 0} já gravados, ${restantes} ainda a processar`;

        // Termina se NAO ha mais arquivos novos a processar nessa diretoria
        if (restantes === 0) break;
        if (tentativas > 20) break;
      } catch (e) {
        acumulado.erros++;
        window._syncErros.push({ diretoria: dir.label, pasta: dir.id, error: 'rede: ' + e.message });
        pastaCompletaMsg = `<b style="color:#DC2626">Erro de rede:</b> ${e.message}`;
        break;
      }
    }

    acumulado.diretorias++;
    porDiretoria.push({ label: dir.label, encontrados: acumulado.totalEncontrados });
  }

  const tempoTotal = Math.round((Date.now() - totalStart) / 1000);
  const min = Math.floor(tempoTotal / 60);
  const seg = tempoTotal % 60;

  const errosHTML = (window._syncErros && window._syncErros.length)
    ? `<details style="margin-top:12px;background:#FEF2F2;border:1px solid #FCA5A5;border-radius:4px">
         <summary style="padding:10px;cursor:pointer;font-weight:600;color:#7F1D1D">⚠ Ver detalhes dos ${window._syncErros.length} erros</summary>
         <div style="padding:10px;font-family:monospace;font-size:11px;max-height:300px;overflow-y:auto">
           ${window._syncErros.slice(0, 30).map(function(e){
             return '<div style="margin-bottom:8px;padding:6px;background:#fff;border-radius:3px"><b>' + (e.diretoria || '?') + '</b> ' + (e.arquivo ? '· <span style="color:#6B7280">' + e.arquivo + '</span>' : '') + '<br>' + (e.error || e.erro || '?') + (e.step ? ' (step: ' + e.step + ')' : '') + (e.graphStatusCode ? ' (graph: ' + e.graphStatusCode + ')' : '') + '</div>';
           }).join('')}
           ${window._syncErros.length > 30 ? '<div class="muted">...e mais ' + (window._syncErros.length - 30) + ' erros</div>' : ''}
         </div>
       </details>`
    : '';

  atualiza(`
    <div class="alert alert-success" style="padding:14px;border-radius:4px;display:block">
      <div style="font-weight:600;font-size:16px;margin-bottom:10px">${window._syncCancelado ? '⏹ Sincronização cancelada' : '✓ Sincronização concluída'}</div>
      <div style="display:block;margin-bottom:4px"><b>Diretorias processadas:</b> ${acumulado.diretorias}/${DIRETORIAS_PRONEP.length}</div>
      <div style="display:block;margin-bottom:4px"><b>Novos contratos:</b> ${acumulado.novos}</div>
      <div style="display:block;margin-bottom:4px"><b>Atualizados:</b> ${acumulado.atualizados}</div>
      <div style="display:block;margin-bottom:4px"><b>Já existentes (pulados):</b> ${acumulado.pulados}</div>
      <div style="display:block;margin-bottom:4px"><b>Erros:</b> ${acumulado.erros}</div>
      <div style="display:block;margin-bottom:4px"><b>Total de batches:</b> ${acumulado.batches}</div>
      <div style="display:block;margin-bottom:4px"><b>Total de arquivos encontrados:</b> ${acumulado.totalEncontrados}</div>
      <div style="display:block;margin-bottom:4px"><b>Tempo total:</b> ${min}m ${seg}s</div>
      ${errosHTML}
      <div style="margin-top:14px"><button class="btn btn-primary" onclick="closeModal(); carregarContratosReais();">Fechar e atualizar lista</button></div>
    </div>
  `);
}

function abrirSincronizarTudo() {
  openModal(`
    <h3 style="margin-top:0">Sincronizar TODAS as diretorias da Pronep</h3>
    <div style="background:#FEF3C7;border-left:3px solid #F59E0B;padding:12px;border-radius:4px;margin-bottom:14px">
      <b>Atenção:</b> Esta operação vai varrer as 10 diretorias do SharePoint, ler todos os PDFs novos e extrair vigência via IA.
      <ul style="margin:6px 0 0 18px;padding:0">
        <li>Duração: <b>30 a 90 minutos</b></li>
        <li>Custo estimado: <b>$1.50 a $4 em tokens Claude</b></li>
        <li>Arquivos já gravados são pulados (não há custo extra)</li>
        <li>Pode cancelar a qualquer momento (termina o batch atual)</li>
      </ul>
    </div>
    <div id="syc-resultado" style="margin-bottom:14px"></div>
    <div style="display:flex;gap:10px;justify-content:flex-end">
      <button class="btn btn-secondary" onclick="closeModal()">Fechar</button>
      <button class="btn btn-primary" onclick="window.contratosBfs.iniciar()">▶ Iniciar sincronização completa (BFS)</button>
    </div>
  `);
}

function abrirSincronizarContratos() {
  const html = `
    <h3 style="margin-top:0">Sincronizar contratos do SharePoint</h3>
    <div class="small muted" style="margin-bottom:14px">Indique o caminho da pasta dentro de <span class="mono">/CONTRATOS/CONTRATOS E DOCUMENTOS - PRESTADORES</span>. Deixe vazio para sincronizar tudo (cuidado, varios PDFs podem custar tempo + tokens).</div>
    <div style="margin-bottom:10px">
      <div class="label">Pasta (subpath)</div>
      <input id="syc-pasta" type="text" placeholder="Ex: /GERÊNCIA DE PROJETOS E TI/CORPORATIVO/TOTVS SA" style="width:100%" />
    </div>
    <div style="display:flex;gap:10px;align-items:center;margin-bottom:10px">
      <div style="flex:1">
        <div class="label">Máximo de arquivos</div>
        <input id="syc-max" type="number" value="30" min="1" max="100" style="width:100%">
      </div>
      <div style="flex:1">
        <label><input id="syc-dryrun" type="checkbox"> Dry-run (só simula, não grava)</label>
      </div>
    </div>
    <div style="margin-bottom:14px">
      <label><input id="syc-force" type="checkbox"> Forçar releitura (re-processa arquivos já gravados — use ao corrigir mapeamento de diretoria ou regravar todos)</label>
    </div>
    <div style="display:flex;gap:10px;justify-content:flex-end;flex-wrap:wrap">
      <button class="btn btn-secondary" onclick="closeModal()">Cancelar</button>
      <button class="btn btn-secondary" onclick="executarSincronizarContratos()" title="Chamada direta. Pode dar timeout em pastas grandes.">Sincronizar (rápido)</button>
      <button class="btn btn-primary" onclick="executarSincronizarBFS()" title="Quebra em sub-pastas pequenas. Evita timeout em escopos grandes.">⚡ Sincronizar via BFS</button>
    </div>
    <div style="margin-top:8px;font-size:11px;color:#6B7280">
      <b>Quando usar BFS?</b> Pra varrer uma diretoria/unidade inteira com muitos prestadores. Quebra em chamadas pequenas — sem timeout.
    </div>
    <div id="syc-resultado" style="margin-top:14px"></div>
  `;
  openModal(html);
}
