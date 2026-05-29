/**
 * SOL — Assistente IA do Sistema de Aprovacao de NF Pronep
 *
 * Wrapper sobre OpenAI com tool use (function calling). Define as tools que a
 * SOL pode chamar e o system prompt que rege seu comportamento.
 *
 * IMPORTANTE - Seguranca:
 *  - A SOL recebe um `user` (do getUser/auth) e SO consulta dados que esse user
 *    tem permissao de ver. Read-only tools chamam ListarNotas com o mesmo principal.
 *  - Acoes destrutivas (aprovar, rejeitar) NAO sao executadas pela SOL. As tools
 *    "propor_aprovacao" e "propor_rejeicao" apenas RETORNAM metadata pro frontend
 *    mostrar um card de confirmacao. O usuario clica e o frontend chama
 *    AprovarNota/RejeitarNota normalmente (com o mesmo auth).
 *
 * Dependencias: openai SDK (^4)
 *
 * App Settings exigidas:
 *   OPENAI_API_KEY - chave da OpenAI
 *   OPENAI_MODEL   - opcional, default 'gpt-4o-mini'
 */

require('isomorphic-fetch');
const OpenAI = require('openai');

const DEFAULT_MODEL = process.env.OPENAI_MODEL || 'gpt-4o-mini';

// =============================================================================
// CLIENTE OpenAI
// =============================================================================
let _openai = null;
function getOpenAI() {
  if (_openai) return _openai;
  const apiKey = process.env.OPENAI_API_KEY;
  if (!apiKey) throw new Error('OPENAI_API_KEY nao configurada');
  _openai = new OpenAI({ apiKey: apiKey });
  return _openai;
}

// =============================================================================
// SYSTEM PROMPT
// =============================================================================
function buildSystemPrompt(user, viewAtual) {
  const hoje = new Date();
  const brt = new Date(hoje.getTime() - 3 * 60 * 60 * 1000);
  const dataHoje = brt.getUTCFullYear() + '-' + String(brt.getUTCMonth()+1).padStart(2,'0') + '-' + String(brt.getUTCDate()).padStart(2,'0');

  // Extrai primeiro nome pro tratamento amigavel
  const firstName = (user.name || user.email || '').split(/[\s.@]/)[0]
    .replace(/[^A-Za-zÀ-ÿ]/g, '')
    .replace(/^./, c => c.toUpperCase());

  return [
    'Voce eh a SOL — assistente IA do sistema de Aprovacao de Notas Fiscais da Pronep Life Care.',
    '',
    'IDENTIDADE:',
    '  - Nome: SOL',
    '  - Tom: profissional, objetivo, calorosa. Usa portugues brasileiro coloquial mas correto.',
    '  - Sem emojis exceto se o usuario usar primeiro.',
    '  - Respostas curtas e diretas. Sem rodeios, sem floreio.',
    '',
    'CONTEXTO:',
    '  - Usuario logado: ' + user.name + ' (' + user.email + ')',
    '  - Primeiro nome (use SEMPRE este pra se dirigir ao usuario): ' + firstName,
    '  - Data de hoje (BRT): ' + dataHoje,
    '  - Tela atual: ' + (viewAtual || 'fila-aprovacao'),
    '',
    'CAPACIDADES (via tools):',
    '  - listar_fila: lista as NFs que o usuario tem PRA APROVAR (status Lancada/Aprovacao)',
    '  - listar_aprovadas: NFs ja aprovadas, com filtros de periodo',
    '  - detalhes_nf: detalhes completos de uma NF especifica pelo id',
    '  - agregar_por_fornecedor: soma e conta NFs por fornecedor (escopo fila ou aprovadas)',
    '  - detectar_anomalia: compara uma NF com a media historica do fornecedor (alerta se outlier)',
    '  - propor_aprovacao: prepara uma AC ÇÃO de aprovacao (NAO executa - retorna pro frontend confirmar)',
    '  - propor_rejeicao: prepara uma AÇÃO de rejeicao (idem)',
    '',
    'REGRAS DE OURO:',
    '  1. NUNCA aprove ou rejeite sem o usuario confirmar. Use propor_aprovacao / propor_rejeicao.',
    '  2. Quando o usuario pedir relatorio da fila, SEMPRE ordene por vencimento crescente (mais proximo primeiro).',
    '  3. Destaque as NFs vencendo em ate 5 dias uteis (D+5) — sao prioridade.',
    '  4. Se o usuario pedir algo fora do dominio (NF, fornecedor, aprovacao), recuse educadamente.',
    '  5. Numeros monetarios: formate R$ 15.000,00 (ponto milhar, virgula decimal).',
    '  6. Datas: formato dd/mm/aaaa nas respostas (mas use ISO YYYY-MM-DD nas tools).',
    '  7. Quando listar NFs, use formato markdown table compacto.',
    '  8. Se nao tem dados pra responder, fale isso direto. Nao invente.',
    '  9. Pra acoes destrutivas, SEMPRE confirme o numero, fornecedor e valor da NF antes de propor.',
    '  10. Se NAO encontrar a NF que o usuario pediu (numero invalido, fora do escopo dele, etc), responda de forma AMIGAVEL e EMPATICA, sempre usando o PRIMEIRO NOME do usuario (campo "Primeiro nome" no contexto acima). Exemplo: \'Oi ' + firstName + ', nao encontrei a NF 1234 que voce me pediu. Pode confirmar o numero?\'. NUNCA seja seca tipo \'NF nao encontrada\' — sempre humanize.',
    '  12. Quando for chamar o usuario pelo nome (em saudacoes ou respostas amigaveis), use SEMPRE o primeiro nome (' + firstName + ') — nunca o nome completo, nunca o email.',
    '  11. Quando for propor uma acao destrutiva (aprovacao/rejeicao), redija sua resposta como UMA UNICA pergunta direta, do tipo \'Posso seguir com a aprovacao da NF X (Fornecedor Y, R$ Z)?\'. Nao escreva texto longo antes do card — o frontend ja mostra um modal com os dados, sua mensagem deve ser curta e direta.',
    '',
    'EXEMPLOS DE INTERAÇÃO:',
    '  User: "Liste minha fila"',
    '    → chama listar_fila, depois responde com tabela ordenada por vencimento.',
    '  User: "SOL, aprove a NF 1234"',
    '    → chama detalhes_nf(numero=1234), confirma os dados, chama propor_aprovacao(id).',
    '    Responde: "Encontrei a NF 1234 do fornecedor X, R$ Y, venc. dd/mm. Quer aprovar?"',
    '    NAO chama nenhuma tool de execucao real.',
    '  User: "Quanto vou liberar este mes?"',
    '    → chama agregar_por_fornecedor(escopo=fila), soma total, responde com numero.'
  ].join('\n');
}

// =============================================================================
// TOOL SCHEMAS (formato OpenAI Function Calling)
// =============================================================================
const TOOLS = [
  {
    type: 'function',
    function: {
      name: 'listar_fila',
      description: 'Lista as NFs pendentes de aprovacao do usuario atual, ordenadas por vencimento (mais proximo primeiro). Aplica RBAC automaticamente — usuario so ve o que tem permissao.',
      parameters: {
        type: 'object',
        properties: {
          unidade: { type: 'string', enum: ['SP','RJ','ES','TODAS'], description: 'Filtrar por unidade (default TODAS)' },
          apenas_d5: { type: 'boolean', description: 'Se true, mostra apenas NFs vencendo em D+5 ou antes' }
        }
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'listar_aprovadas',
      description: 'Lista NFs ja aprovadas pelo usuario ou no escopo dele. Suporta filtro de periodo.',
      parameters: {
        type: 'object',
        properties: {
          periodo: { type: 'string', enum: ['hoje','ontem','esta_semana','este_mes','mes_passado','todos'], description: 'Janela temporal de aprovados (default este_mes)' },
          unidade: { type: 'string', enum: ['SP','RJ','ES','TODAS'] }
        }
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'detalhes_nf',
      description: 'Retorna detalhes completos de uma NF pelo id interno (spListItemId) ou pelo NumeroNF. Use quando o usuario perguntar sobre uma NF especifica.',
      parameters: {
        type: 'object',
        properties: {
          id: { type: 'string', description: 'spListItemId (preferido)' },
          numero: { type: 'string', description: 'NumeroNF (fallback - pode ter mais de uma)' }
        }
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'agregar_por_fornecedor',
      description: 'Agrega NFs (count, sum) agrupadas por fornecedor. Util pra "quanto tenho do fornec X" ou "quais sao meus top fornecedores".',
      parameters: {
        type: 'object',
        properties: {
          escopo: { type: 'string', enum: ['fila','aprovadas','ambos'], description: 'Default fila' },
          top_n: { type: 'integer', description: 'Limitar aos top N por valor total (default 10)' }
        }
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'detectar_anomalia',
      description: 'Verifica se uma NF tem valor fora da curva comparado ao historico do mesmo fornecedor. Retorna se eh outlier (>2x ou <0.5x da mediana) e detalhes.',
      parameters: {
        type: 'object',
        properties: {
          id: { type: 'string', description: 'spListItemId da NF' }
        },
        required: ['id']
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'propor_aprovacao',
      description: 'Prepara uma proposta de aprovacao. NAO executa — retorna metadata pro frontend mostrar card de confirmacao. Use depois de confirmar os dados da NF com o usuario.',
      parameters: {
        type: 'object',
        properties: {
          id: { type: 'string', description: 'spListItemId da NF a aprovar' }
        },
        required: ['id']
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'propor_rejeicao',
      description: 'Prepara uma proposta de rejeicao. NAO executa — retorna metadata pro frontend mostrar card de confirmacao. Use depois de confirmar com o usuario.',
      parameters: {
        type: 'object',
        properties: {
          id: { type: 'string', description: 'spListItemId da NF a rejeitar' },
          motivo: { type: 'string', description: 'Motivo da rejeicao (obrigatorio)' }
        },
        required: ['id','motivo']
      }
    }
  }
];

// =============================================================================
// TOOL EXECUTORS — implementacao das tools que a SOL pode chamar
// =============================================================================
// Recebem: (args, ctx) onde ctx = { user, graphClient, siteId, listNotasId, colMap, invColMap, listFornecId, listDirId }
// Retornam: string (JSON serializado) que a OpenAI vai ler

const { ClientSecretCredential } = require('@azure/identity');
const { Client } = require('@microsoft/microsoft-graph-client');
const { TokenCredentialAuthenticationProvider } = require('@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials');

const LIST_NOTAS = 'PRONEP-NF-NotasFiscais';
const LIST_FORNECEDORES = 'PRONEP-NF-Fornecedores';
const LIST_DIRETORIAS = 'PRONEP-NF-Diretorias';

// Cache de site/list em memoria
const _cache = { siteId: null, listNotasId: null, listFornecId: null, listDirId: null, colMap: null, invColMap: null, fornCache: null, fornCacheAt: 0 };

async function getGraphClient() {
  const tenantId = process.env.AAD_TENANT_ID;
  const clientId = process.env.AAD_CLIENT_ID;
  const clientSecret = process.env.AAD_CLIENT_SECRET;
  if (!tenantId || !clientId || !clientSecret) throw new Error('AAD_* incompletas');
  const credential = new ClientSecretCredential(tenantId, clientId, clientSecret);
  const authProvider = new TokenCredentialAuthenticationProvider(credential, {
    scopes: ['https://graph.microsoft.com/.default']
  });
  return Client.initWithMiddleware({ authProvider });
}

async function resolveSiteAndLists(client) {
  if (_cache.siteId && _cache.listNotasId) return _cache;
  const host = process.env.SHAREPOINT_SITE_HOSTNAME;
  const path = process.env.SHAREPOINT_SITE_PATH;
  const siteResp = await client.api('/sites/' + host + ':' + path).get();
  _cache.siteId = siteResp.id;
  const listsResp = await client.api('/sites/' + _cache.siteId + '/lists').get();
  for (const l of (listsResp.value || [])) {
    if (l.displayName === LIST_NOTAS) _cache.listNotasId = l.id;
    if (l.displayName === LIST_FORNECEDORES) _cache.listFornecId = l.id;
    if (l.displayName === LIST_DIRETORIAS) _cache.listDirId = l.id;
  }
  // colMap (displayName -> internalName) e invColMap (internalName -> displayName)
  const colsResp = await client.api('/sites/' + _cache.siteId + '/lists/' + _cache.listNotasId + '/columns').get();
  _cache.colMap = {}; _cache.invColMap = {};
  for (const c of (colsResp.value || [])) {
    if (c.displayName && c.name) {
      _cache.colMap[c.displayName] = c.name;
      _cache.invColMap[c.name] = c.displayName;
    }
  }
  return _cache;
}

// Le item bruto e normaliza pra fields { Title, Status, ValorTotal, ... } usando invColMap
function normalizeItem(item, invColMap) {
  const raw = item.fields || {};
  const out = { _id: item.id };
  for (const [k, v] of Object.entries(raw)) {
    const display = invColMap[k] || k;
    out[display] = v;
  }
  return out;
}

// Aplica RBAC: gestor ve onde AprovadorAtual = email; admin/financeiro ve tudo
function filtraRBAC(notas, user, isAdmin) {
  if (isAdmin) return notas;
  const email = (user.email || '').toLowerCase();
  return notas.filter(n => {
    const apv = String(n.AprovadorAtual || '').toLowerCase();
    return apv === email;
  });
}

async function tool_listar_fila(args, ctx) {
  const { client, siteId, listNotasId, invColMap } = ctx.gr;
  const top = 500;
  // Status pendente = "Lancada" ou em fluxo de aprovacao
  let url = '/sites/' + siteId + '/lists/' + listNotasId + '/items?expand=fields&$top=' + top;
  const resp = await client.api(url).get();
  let notas = (resp.value || []).map(it => normalizeItem(it, invColMap));
  // Filtra: status Lancada (pendente) ou EmAprovacao
  notas = notas.filter(n => ['Lancada','EmAprovacao','Pendente'].includes(String(n.Status || '')));
  // Filtra RBAC (admin/financeiro veem tudo, resto so o seu)
  notas = filtraRBAC(notas, ctx.user, ctx.isAdmin);
  if (args.unidade && args.unidade !== 'TODAS') {
    notas = notas.filter(n => String(n.Unidade || '') === args.unidade);
  }
  // Ordena por vencimento ASC
  notas.sort((a, b) => String(a.Vencimento || '').localeCompare(String(b.Vencimento || '')));
  if (args.apenas_d5) {
    const hoje = new Date(new Date().getTime() - 3*60*60*1000);
    const d5 = new Date(hoje.getTime() + 5*24*60*60*1000).toISOString().substring(0,10);
    notas = notas.filter(n => String(n.Vencimento || '').substring(0,10) <= d5);
  }
  // Limita resposta pra nao estourar contexto
  notas = notas.slice(0, 50);
  return {
    total: notas.length,
    notas: notas.map(n => ({
      id: n._id,
      numero: n.NumeroNF,
      fornecedor: n.Fornecedor,
      cnpj: n.FornecedorCNPJ,
      valor: Number(n.ValorTotal || n.Valor || 0),
      vencimento: String(n.Vencimento || '').substring(0,10),
      unidade: n.Unidade,
      diretoria: n.Diretoria,
      status: n.Status,
      lancadoPor: n.LancadoPor,
      aprovadorAtual: n.AprovadorAtual
    }))
  };
}

async function tool_listar_aprovadas(args, ctx) {
  const { client, siteId, listNotasId, invColMap } = ctx.gr;
  const periodo = args.periodo || 'este_mes';
  const resp = await client.api('/sites/' + siteId + '/lists/' + listNotasId + '/items?expand=fields&$top=2000').get();
  let notas = (resp.value || []).map(it => normalizeItem(it, invColMap));
  notas = notas.filter(n => String(n.Status || '') === 'Aprovada');
  // Periodo
  const agora = new Date(new Date().getTime() - 3*60*60*1000);
  const hojeStr = agora.toISOString().substring(0,10);
  function dateNDaysAgo(n) {
    return new Date(agora.getTime() - n*24*60*60*1000).toISOString().substring(0,10);
  }
  if (periodo === 'hoje') notas = notas.filter(n => String(n.AprovadoEm || '').substring(0,10) === hojeStr);
  else if (periodo === 'ontem') notas = notas.filter(n => String(n.AprovadoEm || '').substring(0,10) === dateNDaysAgo(1));
  else if (periodo === 'esta_semana') notas = notas.filter(n => String(n.AprovadoEm || '').substring(0,10) >= dateNDaysAgo(7));
  else if (periodo === 'este_mes') {
    const inicioMes = hojeStr.substring(0,7) + '-01';
    notas = notas.filter(n => String(n.AprovadoEm || '').substring(0,10) >= inicioMes);
  } else if (periodo === 'mes_passado') {
    const d = new Date(agora);
    d.setUTCMonth(d.getUTCMonth() - 1);
    const inicio = d.toISOString().substring(0,7) + '-01';
    const fim = hojeStr.substring(0,7) + '-01';
    notas = notas.filter(n => {
      const a = String(n.AprovadoEm || '').substring(0,10);
      return a >= inicio && a < fim;
    });
  }
  // RBAC: usuario so ve oque aprovou ou esta no escopo dele
  notas = filtraRBAC(notas, ctx.user, ctx.isAdmin);
  if (args.unidade && args.unidade !== 'TODAS') {
    notas = notas.filter(n => String(n.Unidade || '') === args.unidade);
  }
  notas.sort((a, b) => String(b.AprovadoEm || '').localeCompare(String(a.AprovadoEm || '')));
  notas = notas.slice(0, 50);
  return {
    total: notas.length,
    periodo: periodo,
    notas: notas.map(n => ({
      id: n._id,
      numero: n.NumeroNF,
      fornecedor: n.Fornecedor,
      valor: Number(n.ValorTotal || n.Valor || 0),
      vencimento: String(n.Vencimento || '').substring(0,10),
      aprovadoEm: String(n.AprovadoEm || '').substring(0,10),
      unidade: n.Unidade
    }))
  };
}

async function tool_detalhes_nf(args, ctx) {
  const { client, siteId, listNotasId, invColMap } = ctx.gr;
  let item = null;
  if (args.id) {
    try {
      item = await client.api('/sites/' + siteId + '/lists/' + listNotasId + '/items/' + args.id + '?expand=fields').get();
    } catch (e) { /* nao achou pelo id, tenta por numero */ }
  }
  if (!item && args.numero) {
    // Pagina por todas as NFs ate achar (max 30 paginas = 15k items)
    const alvoNum = String(args.numero).replace(/\D/g, '');
    let url = '/sites/' + siteId + '/lists/' + listNotasId + '/items?expand=fields&$top=500';
    let pages = 0;
    while (url && pages < 30 && !item) {
      const resp = await client.api(url).get();
      for (const it of (resp.value || [])) {
        const n = normalizeItem(it, invColMap);
        const numStr = String(n.NumeroNF || '');
        // Match exato OU normalizado (so digitos) pra tolerar formatacao tipo "NF-48" vs "48"
        if (numStr === String(args.numero) || numStr.replace(/\D/g,'') === alvoNum) {
          item = { id: n._id, fields: it.fields };
          break;
        }
      }
      pages++;
      url = resp['@odata.nextLink']
        ? resp['@odata.nextLink'].replace('https://graph.microsoft.com/v1.0', '')
        : null;
    }
  }
  if (!item) return { erro: 'NF nao encontrada' };
  const n = normalizeItem(item, invColMap);
  // RBAC check
  if (!ctx.isAdmin) {
    const email = (ctx.user.email || '').toLowerCase();
    const apv = String(n.AprovadorAtual || '').toLowerCase();
    const lan = String(n.LancadoPor || '').toLowerCase();
    if (apv !== email && lan !== email) {
      return { erro: 'Voce nao tem permissao pra ver essa NF' };
    }
  }
  return {
    id: n._id,
    numero: n.NumeroNF,
    fornecedor: n.Fornecedor,
    cnpj: n.FornecedorCNPJ,
    valor: Number(n.ValorTotal || n.Valor || 0),
    vencimento: String(n.Vencimento || '').substring(0,10),
    unidade: n.Unidade,
    diretoria: n.Diretoria,
    status: n.Status,
    aprovadorAtual: n.AprovadorAtual,
    lancadoPor: n.LancadoPor,
    aprovadoEm: n.AprovadoEm,
    descricao: n.Descricao,
    serie: n.Serie
  };
}

async function tool_agregar_por_fornecedor(args, ctx) {
  const escopo = args.escopo || 'fila';
  const topN = args.top_n || 10;
  let notas = [];
  if (escopo === 'fila' || escopo === 'ambos') {
    const r = await tool_listar_fila({ unidade: 'TODAS' }, ctx);
    notas = notas.concat(r.notas);
  }
  if (escopo === 'aprovadas' || escopo === 'ambos') {
    const r = await tool_listar_aprovadas({ periodo: 'este_mes' }, ctx);
    notas = notas.concat(r.notas);
  }
  const grupos = {};
  for (const n of notas) {
    const k = n.fornecedor || '(sem nome)';
    if (!grupos[k]) grupos[k] = { fornecedor: k, qtd: 0, total: 0 };
    grupos[k].qtd += 1;
    grupos[k].total += Number(n.valor || 0);
  }
  let lista = Object.values(grupos);
  lista.sort((a,b) => b.total - a.total);
  lista = lista.slice(0, topN);
  return { escopo: escopo, fornecedores: lista };
}

async function tool_detectar_anomalia(args, ctx) {
  // Pega a NF
  const detalhe = await tool_detalhes_nf({ id: args.id }, ctx);
  if (detalhe.erro) return detalhe;
  // Pega historico do mesmo fornecedor (CNPJ)
  const { client, siteId, listNotasId, invColMap } = ctx.gr;
  const resp = await client.api('/sites/' + siteId + '/lists/' + listNotasId + '/items?expand=fields&$top=2000').get();
  let historico = (resp.value || []).map(it => normalizeItem(it, invColMap));
  historico = historico.filter(n => {
    const cnpj1 = String(n.FornecedorCNPJ || '').replace(/\D/g, '');
    const cnpj2 = String(detalhe.cnpj || '').replace(/\D/g, '');
    return cnpj1 && cnpj1 === cnpj2 && n._id !== detalhe.id;
  });
  if (historico.length < 3) {
    return {
      nf: detalhe,
      conclusao: 'historico_insuficiente',
      detalhe: 'Fornecedor com menos de 3 NFs historicas — nao da pra comparar com confianca.',
      qtd_historico: historico.length
    };
  }
  const valores = historico.map(n => Number(n.ValorTotal || n.Valor || 0)).filter(v => v > 0).sort((a,b) => a - b);
  const mediana = valores[Math.floor(valores.length / 2)];
  const media = valores.reduce((s,v) => s+v, 0) / valores.length;
  const valor = detalhe.valor;
  let conclusao = 'normal';
  if (valor > mediana * 2) conclusao = 'alto';
  else if (valor < mediana * 0.5) conclusao = 'baixo';
  return {
    nf: detalhe,
    conclusao: conclusao,
    valor_atual: valor,
    mediana_historica: Math.round(mediana * 100) / 100,
    media_historica: Math.round(media * 100) / 100,
    qtd_historico: valores.length,
    razao_vs_mediana: Math.round((valor / mediana) * 100) / 100
  };
}

// PROPOR — retorna metadata pro frontend, NAO executa
async function tool_propor_aprovacao(args, ctx) {
  const detalhe = await tool_detalhes_nf({ id: args.id }, ctx);
  if (detalhe.erro) return detalhe;
  return {
    acao_proposta: 'aprovacao',
    id: detalhe.id,
    confirmar_no_frontend: true,
    dados: detalhe
  };
}

async function tool_propor_rejeicao(args, ctx) {
  const detalhe = await tool_detalhes_nf({ id: args.id }, ctx);
  if (detalhe.erro) return detalhe;
  return {
    acao_proposta: 'rejeicao',
    id: detalhe.id,
    motivo: args.motivo,
    confirmar_no_frontend: true,
    dados: detalhe
  };
}

const TOOL_IMPL = {
  listar_fila: tool_listar_fila,
  listar_aprovadas: tool_listar_aprovadas,
  detalhes_nf: tool_detalhes_nf,
  agregar_por_fornecedor: tool_agregar_por_fornecedor,
  detectar_anomalia: tool_detectar_anomalia,
  propor_aprovacao: tool_propor_aprovacao,
  propor_rejeicao: tool_propor_rejeicao
};

// =============================================================================
// LOOP PRINCIPAL: conversa + tool use
// =============================================================================
/**
 * Executa um turno da SOL. Pode envolver multiplas iteracoes (chama tool, le
 * resultado, decide proximo passo).
 *
 * @param {Array} history  Mensagens anteriores no formato OpenAI (system NAO incluido aqui — eh adicionado)
 * @param {string} userMessage  Nova mensagem do usuario
 * @param {Object} user  { email, name, oid } do auth
 * @param {Object} opts  { viewAtual, isAdmin, model, maxIter }
 * @returns {Object} { resposta, acoes_propostas: [], tokensUsed }
 */
async function runSol(history, userMessage, user, opts) {
  opts = opts || {};
  const model = opts.model || DEFAULT_MODEL;
  const maxIter = opts.maxIter || 8;
  const viewAtual = opts.viewAtual || 'fila-aprovacao';
  const isAdmin = !!opts.isAdmin;

  const openai = getOpenAI();
  const client = await getGraphClient();
  const { siteId, listNotasId, invColMap } = await resolveSiteAndLists(client);

  const ctx = {
    user: user,
    isAdmin: isAdmin,
    gr: { client, siteId, listNotasId, invColMap }
  };

  // Monta mensagens
  const messages = [
    { role: 'system', content: buildSystemPrompt(user, viewAtual) }
  ];
  // Adiciona historico previo (filtra so role+content e tool_calls/tool responses)
  for (const h of (history || [])) {
    if (!h || !h.role) continue;
    if (h.role === 'user' || h.role === 'assistant') {
      messages.push({ role: h.role, content: h.content || '' });
    }
  }
  messages.push({ role: 'user', content: userMessage });

  const acoesPropostas = [];
  let totalTokens = 0;
  let resposta = '';

  for (let i = 0; i < maxIter; i++) {
    const completion = await openai.chat.completions.create({
      model: model,
      messages: messages,
      tools: TOOLS,
      tool_choice: 'auto',
      temperature: 0.2,
      max_tokens: 1500
    });
    totalTokens += (completion.usage && completion.usage.total_tokens) || 0;

    const msg = completion.choices[0].message;
    messages.push(msg);

    // Se nao tem tool calls, eh a resposta final
    if (!msg.tool_calls || msg.tool_calls.length === 0) {
      resposta = msg.content || '';
      break;
    }

    // Executa cada tool call
    for (const tc of msg.tool_calls) {
      const fnName = tc.function && tc.function.name;
      let args = {};
      try { args = JSON.parse(tc.function.arguments || '{}'); } catch (e) {}
      const impl = TOOL_IMPL[fnName];
      let result;
      if (!impl) {
        result = { erro: 'tool desconhecida: ' + fnName };
      } else {
        try {
          result = await impl(args, ctx);
        } catch (e) {
          result = { erro: 'falha ao executar ' + fnName + ': ' + (e.message || String(e)) };
        }
      }
      // Se for uma proposta de acao, guarda pra retornar pro frontend
      if (result && result.confirmar_no_frontend) {
        acoesPropostas.push(result);
      }
      messages.push({
        role: 'tool',
        tool_call_id: tc.id,
        content: JSON.stringify(result)
      });
    }
  }

  return { resposta: resposta, acoes_propostas: acoesPropostas, tokens: totalTokens, model: model };
}

module.exports = { runSol, DEFAULT_MODEL };
