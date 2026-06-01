/**
 * SOL — Assistente IA do Sistema de Aprovacao de NF Pronep
 *
 * Wrapper dual-provider sobre Anthropic Claude (primario) e OpenAI (fallback)
 * com tool use (function calling). Define as tools que a SOL pode chamar e o
 * system prompt que rege seu comportamento.
 *
 * IMPORTANTE - Seguranca:
 *  - A SOL recebe um `user` (do getUser/auth) e SO consulta dados que esse user
 *    tem permissao de ver. Read-only tools chamam ListarNotas com o mesmo principal.
 *  - Acoes destrutivas (aprovar, rejeitar) NAO sao executadas pela SOL. As tools
 *    "propor_aprovacao" e "propor_rejeicao" apenas RETORNAM metadata pro frontend
 *    mostrar um card de confirmacao. O usuario clica e o frontend chama
 *    AprovarNota/RejeitarNota normalmente (com o mesmo auth).
 *
 * PROVIDERS:
 *  - Anthropic Claude (primario): claude-haiku-4-5 default, escala pra
 *    claude-sonnet-4-6 se a SOL chamar 3+ tools na mesma turn (consulta complexa).
 *  - OpenAI (fallback): gpt-4o-mini. Roda automaticamente se Anthropic falhar
 *    ou ANTHROPIC_API_KEY nao estiver setada.
 *
 * Dependencias: @anthropic-ai/sdk ^0.32, openai ^4
 *
 * App Settings exigidas (pelo menos uma):
 *   ANTHROPIC_API_KEY - chave Anthropic (preferido)
 *   OPENAI_API_KEY    - chave OpenAI (fallback)
 *
 * App Settings opcionais:
 *   ANTHROPIC_MODEL_HAIKU  - default 'claude-haiku-4-5-20251001'
 *   ANTHROPIC_MODEL_SONNET - default 'claude-sonnet-4-6'
 *   OPENAI_MODEL           - default 'gpt-4o-mini'
 */

require('isomorphic-fetch');
// SDKs sao carregados LAZY (dentro dos getters) pra que erro de require de
// um nao crashe a Function inteira. Se @anthropic-ai/sdk nao tiver no
// node_modules, getAnthropic() retorna null e cai no fallback OpenAI.

const MODEL_HAIKU = process.env.ANTHROPIC_MODEL_HAIKU || 'claude-haiku-4-5-20251001';
const MODEL_SONNET = process.env.ANTHROPIC_MODEL_SONNET || 'claude-sonnet-4-6';
const DEFAULT_MODEL = process.env.ANTHROPIC_MODEL || MODEL_HAIKU;
const OPENAI_MODEL = process.env.OPENAI_MODEL || 'gpt-4o-mini';

// =============================================================================
// CLIENTES (Anthropic primario, OpenAI fallback)
// =============================================================================
let _anthropic = null;
let _anthropicLoadError = null;
function getAnthropic() {
  if (_anthropic) return _anthropic;
  if (_anthropicLoadError) return null; // ja tentou e falhou, nao tenta de novo
  const apiKey = process.env.ANTHROPIC_API_KEY;
  if (!apiKey) return null;
  try {
    const mod = require('@anthropic-ai/sdk');
    const Anthropic = mod.default || mod.Anthropic || mod;
    _anthropic = new Anthropic({ apiKey: apiKey });
    return _anthropic;
  } catch (e) {
    _anthropicLoadError = e;
    console.error('[SOL] Erro ao carregar @anthropic-ai/sdk:', e && e.message);
    return null;
  }
}

let _openai = null;
let _openaiLoadError = null;
function getOpenAI() {
  if (_openai) return _openai;
  if (_openaiLoadError) return null;
  const apiKey = process.env.OPENAI_API_KEY;
  if (!apiKey) return null;
  try {
    const OpenAI = require('openai');
    _openai = new OpenAI({ apiKey: apiKey });
    return _openai;
  } catch (e) {
    _openaiLoadError = e;
    console.error('[SOL] Erro ao carregar openai:', e && e.message);
    return null;
  }
}

function getLoadErrors() {
  return {
    anthropic: _anthropicLoadError ? { message: _anthropicLoadError.message, type: _anthropicLoadError.constructor && _anthropicLoadError.constructor.name } : null,
    openai: _openaiLoadError ? { message: _openaiLoadError.message, type: _openaiLoadError.constructor && _openaiLoadError.constructor.name } : null
  };
}

// =============================================================================
// SYSTEM PROMPT
// =============================================================================
// =============================================================================
// ESCOPO POR VIEW: SAN se comporta diferente em cada aba
// =============================================================================
const VIEW_SCOPES = {
  'fila-aprovacao': {
    titulo: 'Fila de Aprovacao',
    foco: 'analisar fila pendente, propor aprovar/rejeitar, detectar anomalias e responder duvidas sobre as NFs do usuario.',
    tools: ['listar_fila','listar_aprovadas','detalhes_nf','agregar_por_fornecedor','detectar_anomalia','propor_aprovacao','propor_rejeicao']
  },
  'aprovadas': {
    titulo: 'Notas Aprovadas',
    foco: 'RELATORIOS pro financeiro, agregacoes por fornecedor/periodo/diretoria, abrir PDFs de NFs aprovadas e propor marcar como processado. Aqui NAO HA aprovar/rejeitar — essas NFs ja passaram.',
    tools: ['listar_aprovadas','detalhes_nf','agregar_por_fornecedor','abrir_nf','propor_marcar_processado']
  },
  'lancamento': {
    titulo: 'Lancamento de NF',
    foco: 'APENAS ORIENTAR o usuario sobre como preencher o formulario de lancamento de NF. Voce NAO tem acesso a dados aqui — sem tools. Explica campos, mascaras, validacoes, upload de PDF.',
    tools: []
  },
  'fornecedores': { titulo: 'Fornecedores', foco: 'apenas orientacao sobre como usar a tela de cadastro de fornecedores. Sem tools.', tools: [] },
  'dashboard':    { titulo: 'Dashboard',    foco: 'apenas orientacao sobre os indicadores e graficos exibidos. Sem tools.', tools: [] },
  'rejeitadas':   { titulo: 'Notas Rejeitadas', foco: 'apenas orientacao sobre o que aparece aqui. Sem tools.', tools: [] },
  'minhas-nfs':   { titulo: 'Minhas NFs', foco: 'orientacao + leitura de aprovadas. Tools de leitura ok, sem acoes.', tools: ['listar_aprovadas','detalhes_nf'] },
  'rejeitadas-minhas': { titulo: 'Minhas NFs Rejeitadas', foco: 'apenas orientacao. Sem tools.', tools: [] },
  'configuracoes': { titulo: 'Configuracoes', foco: 'apenas orientacao sobre opcoes de admin/usuario. Sem tools.', tools: [] }
};

function getViewScope(viewAtual) {
  return VIEW_SCOPES[viewAtual || 'fila-aprovacao'] || VIEW_SCOPES['fila-aprovacao'];
}

function getToolsForView(viewAtual) {
  var scope = getViewScope(viewAtual);
  var allowed = scope.tools || [];
  if (allowed.length === 0) return [];
  return TOOLS.filter(function(t){ return allowed.indexOf(t.function.name) >= 0; });
}

function buildSystemPrompt(user, viewAtual) {
  const hoje = new Date();
  const brt = new Date(hoje.getTime() - 3 * 60 * 60 * 1000);
  const dataHoje = brt.getUTCFullYear() + '-' + String(brt.getUTCMonth()+1).padStart(2,'0') + '-' + String(brt.getUTCDate()).padStart(2,'0');

  // Extrai primeiro nome pro tratamento amigavel
  const firstName = (user.name || user.email || '').split(/[\s.@]/)[0]
    .replace(/[^A-Za-zÀ-ÿ]/g, '')
    .replace(/^./, c => c.toUpperCase());

  const scope = getViewScope(viewAtual);
  const semTools = (scope.tools || []).length === 0;

  // Prompt comum (identidade + contexto + regras universais)
  const base = [
    'Voce eh a SAN (Sistema de Aprovacao de Notas) — assistente IA do sistema de Aprovacao de Notas Fiscais da Pronep Life Care.',
    '',
    'IDENTIDADE:',
    '  - Nome: SAN',
    '  - Tom: profissional, objetivo, calorosa. Portugues brasileiro coloquial mas correto.',
    '  - Sem emojis exceto se o usuario usar primeiro.',
    '  - Respostas curtas e diretas. Sem rodeios, sem floreio.',
    '',
    'CONTEXTO:',
    '  - Usuario logado: ' + user.name + ' (' + user.email + ')',
    '  - Primeiro nome (use SEMPRE este pra se dirigir ao usuario): ' + firstName,
    '  - Data de hoje (BRT): ' + dataHoje,
    '  - Tela atual: ' + scope.titulo + ' (id: ' + (viewAtual || 'fila-aprovacao') + ')',
    '  - Foco nesta tela: ' + scope.foco,
    '',
    'REGRAS UNIVERSAIS:',
    '  1. Se NAO encontrar dado pedido, responda de forma amigavel usando o primeiro nome: \'Oi ' + firstName + ', nao encontrei...\'. NUNCA seja seca.',
    '  2. Chame o usuario sempre pelo primeiro nome (' + firstName + ') — nunca nome completo nem email.',
    '  3. Numeros monetarios: R$ 15.000,00 (ponto milhar, virgula decimal).',
    '  4. Datas: dd/mm/aaaa nas respostas (ISO YYYY-MM-DD nas tools).',
    '  5. Tabelas markdown compactas pra listar NFs.',
    '  6. Se nao tem dados pra responder, fale direto. NAO INVENTE.',
    '  7. Se o usuario pedir algo fora do dominio do sistema (NF, fornecedor, aprovacao), recuse educadamente.'
  ];

  // === REGRAS DE COMPORTAMENTO ESPECIFICAS DA VIEW ===
  let specific = [];
  if (viewAtual === 'fila-aprovacao') {
    specific = [
      '',
      'REGRAS DESTA TELA (Fila de Aprovacao):',
      '  - Ordene listas por vencimento crescente (mais proximo primeiro).',
      '  - Destaque NFs vencendo em ate 5 dias uteis (D+5) — prioridade.',
      '  - NUNCA aprove ou rejeite sem confirmar. Use propor_aprovacao / propor_rejeicao.',
      '  - CRITICO: quando o usuario disser \'aprove a NF X\' ou \'rejeite a NF X\', voce DEVE em UMA UNICA turn: (1) chamar detalhes_nf({numero: X}), (2) chamar propor_aprovacao({id: <id_do_passo_1>}) ou propor_rejeicao(id, motivo). PROIBIDO terminar a resposta com \'?\' tipo \'quer aprovar?\'. PROIBIDO esperar \'sim\'. O frontend abre AUTOMATICAMENTE o modal de confirmacao — esse modal EH a confirmacao do usuario. Sua resposta em texto deve ter no maximo UMA frase: \'Encontrei a NF X. Abrindo a confirmacao...\'',
      '',
      'EXEMPLOS:',
      '  User: "Liste minha fila" → chama listar_fila, responde com tabela por vencimento.',
      '  User: "SAN, aprove a NF 1234" → detalhes_nf(numero=1234) + propor_aprovacao(id=...) na mesma turn.',
      '  User: "Quanto vou liberar este mes?" → agregar_por_fornecedor(escopo=fila), soma, responde.'
    ];
  } else if (viewAtual === 'aprovadas') {
    specific = [
      '',
      'REGRAS DESTA TELA (Notas Aprovadas):',
      '  - Foco em RELATORIOS pro financeiro. Use tabelas markdown e agregue por fornecedor, diretoria, periodo ou unidade conforme o usuario pedir.',
      '  - Aqui as NFs ja foram aprovadas — NAO ha propor_aprovacao/rejeicao. Se o usuario pedir pra aprovar, oriente a ir pra Fila de Aprovacao.',
      '  - Para abrir PDF: use a tool abrir_nf. O frontend renderiza um BOTAO clicavel "Abrir PDF" na sua resposta — voce NAO abre nada automaticamente. Diga algo tipo: "Pronto, clica no botao abaixo pra abrir o PDF." NUNCA diga "abri o PDF" — quem abre eh o usuario clicando no botao.',
      '  - Para marcar como processado: use a tool propor_marcar_processado (chame detalhes_nf primeiro pra obter o id). O frontend abre modal de confirmacao automaticamente.',
      '  - CRITICO: quando o usuario disser "abre/ver PDF da NF X", chame detalhes_nf(numero=X) + abrir_nf(id=...) na mesma turn. Resposta em texto: BREVE tipo "Abrindo o PDF da NF X..." Sem perguntas.',
      '  - CRITICO: quando o usuario disser "marca a NF X como processada", chame detalhes_nf(numero=X) + propor_marcar_processado(id=...). Resposta breve: "Abrindo a confirmacao pra marcar a NF X como processada..."',
      '',
      'EXEMPLOS:',
      '  User: "Quanto liberei este mes?" → listar_aprovadas(periodo=este_mes), soma total.',
      '  User: "Top 5 fornecedores aprovados" → agregar_por_fornecedor(escopo=aprovadas, top_n=5).',
      '  User: "Abre a NF 1234" → detalhes_nf(numero=1234) + abrir_nf(id=...) na mesma turn. Resposta: "Pronto, clica no botao abaixo pra abrir o PDF da NF 1234."',
      '  User: "Marca a NF 1234 como processada" → detalhes_nf(numero=1234) + propor_marcar_processado(id=...).'
    ];
  } else if (viewAtual === 'lancamento') {
    specific = [
      '',
      'REGRAS DESTA TELA (Lancamento de NF) — MUITO IMPORTANTE:',
      '  - Voce NAO TEM TOOLS aqui. Sua unica funcao eh ORIENTAR o usuario a preencher o formulario.',
      '  - Conhecimento dos campos do form de lancamento:',
      '    * Fornecedor: campo de BUSCA — digite o nome ou CNPJ. Se o fornecedor nao existe, ha um botao "+" pra cadastrar.',
      '    * Categoria/Diretoria: aparecem automaticamente se o fornecedor for de uma unica diretoria. Se o fornecedor for multi-diretoria, o usuario escolhe.',
      '    * Unidade: SP, RJ, ES (radio). Bloqueia automaticamente se o fornecedor atende apenas uma.',
      '    * Numero da NF: como veio impresso na nota (ex: 36534, NF-2026-045).',
      '    * Chave de Acesso: 44 digitos da NF-e (codigo de barras). Obrigatorio pra notas eletronicas. O sistema usa pra detectar duplicidade.',
      '    * Valor: digite o valor em R$ — o sistema aplica mascara automatica (R$ 15.000,00).',
      '    * Vencimento: data de vencimento do boleto/pagamento. Formato dd/mm/aaaa.',
      '    * PDF da NF: arrastar ou clicar pra anexar. So aceita PDF. Tamanho maximo geralmente 10MB.',
      '    * Solicitante: quem demandou o servico (text livre).',
      '  - Validacoes que o usuario pode encontrar:',
      '    * Duplicidade: o sistema checa chave de acesso (44 digitos) e numero+fornecedor. Se ja existe e nao foi rejeitada, BLOQUEIA o envio.',
      '    * Vencimento vencido: alerta se a data ja passou.',
      '    * Fornecedor incompleto: precisa ter CNPJ ou CPF antes de aceitar a NF.',
      '  - Se o usuario pedir pra abrir/aprovar/listar NFs, oriente educadamente que essa tela so faz lancamento — pra acoes ou consultas, ir pra Fila de Aprovacao ou Notas Aprovadas.'
    ];
  } else if (semTools) {
    specific = [
      '',
      'REGRAS DESTA TELA (' + scope.titulo + '):',
      '  - Voce NAO TEM TOOLS aqui. Apenas tira duvidas de NAVEGACAO e USABILIDADE do sistema.',
      '  - Se o usuario pedir uma acao especifica em NF (aprovar, abrir, ver), oriente a ir pra Fila de Aprovacao ou Notas Aprovadas.'
    ];
  }

  return base.concat(specific).join('\n');
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
      description: 'Retorna detalhes completos de uma NF. Quando o usuario disser "NF X" ou "nota X", X eh o NumeroNF (que ele digitou no lancamento). USE SEMPRE o parametro "numero" pra esses casos. Use "id" apenas quando voce ja tem o spListItemId interno de uma chamada anterior (ex: tool listar_fila retornou n.id).',
      parameters: {
        type: 'object',
        properties: {
          id: { type: 'string', description: 'spListItemId interno do SharePoint (item.id retornado por listar_fila). NAO use isso pro numero que o usuario falou.' },
          numero: { type: 'string', description: 'NumeroNF — eh o que o usuario digitou ao lancar a NF (ex: "36534", "NF-2026-045"). USE ESTE quando o usuario disser "NF X".' }
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
  },
  {
    type: 'function',
    function: {
      name: 'abrir_nf',
      description: 'Abre o PDF de uma NF aprovada em nova aba do navegador. Use quando o usuario disser "abrir NF X", "ver PDF da NF X", "me mostra o PDF da X". NAO precisa de confirmacao — eh acao nao-destrutiva (so visualiza).',
      parameters: {
        type: 'object',
        properties: {
          id: { type: 'string', description: 'spListItemId da NF (de listar_aprovadas/detalhes_nf)' },
          numero: { type: 'string', description: 'NumeroNF — usar se o usuario disse o numero da nota e voce ainda nao tem o id' }
        }
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'propor_marcar_processado',
      description: 'Propoe MARCAR uma NF aprovada como PROCESSADA (financeiro liberou pra integracao). NAO executa — retorna metadata pro frontend mostrar card de confirmacao. Use quando o usuario disser "marca a NF X como processada", "processei a NF X", "checa o pago da NF X". Sempre chame detalhes_nf antes pra obter id, fornecedor, valor.',
      parameters: {
        type: 'object',
        properties: {
          id: { type: 'string', description: 'spListItemId da NF (obrigatorio)' }
        },
        required: ['id']
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

// Helper: carrega mapa CNPJ -> RazaoSocial uma vez (cacheado no ctx)
async function carregarMapFornecedoresParaSol(client, siteId) {
  const lists = await client.api('/sites/' + siteId + '/lists').filter("displayName eq 'PRONEP-NF-Fornecedores'").get();
  if (!lists.value || !lists.value.length) return {};
  const listId = lists.value[0].id;
  const colsResp = await client.api('/sites/' + siteId + '/lists/' + listId + '/columns').get();
  const inv = {};
  for (const c of (colsResp.value || [])) {
    if (c.displayName && c.name) inv[c.name] = c.displayName;
  }
  const all = [];
  let url = '/sites/' + siteId + '/lists/' + listId + '/items?expand=fields&$top=500';
  let pages = 0;
  while (url && pages < 30) {
    const resp = await client.api(url).get();
    all.push(...(resp.value || []));
    pages++;
    url = resp['@odata.nextLink']
      ? resp['@odata.nextLink'].replace('https://graph.microsoft.com/v1.0', '')
      : null;
  }
  const mapa = {};
  for (const it of all) {
    const f = it.fields || {};
    const out = {};
    for (const [k, v] of Object.entries(f)) {
      const display = inv[k] || k;
      out[display] = v;
    }
    const cnpj = String(out.documento || out.Documento || out.field_2 || out.CNPJ || '').replace(/\D/g, '');
    const razao = out.Title || out.razao || out.RazaoSocial || '';
    if (cnpj) mapa[cnpj] = razao;
  }
  return mapa;
}

async function tool_detalhes_nf(args, ctx) {
  const { client, siteId, listNotasId, invColMap } = ctx.gr;
  let item = null;
  // Se passou id, tenta GET direto. Pode ser tanto spListItemId quanto NumeroNF
  // (a SOL nem sempre acerta a distincao, entao tratamos os dois)
  if (args.id) {
    try {
      item = await client.api('/sites/' + siteId + '/lists/' + listNotasId + '/items/' + args.id + '?expand=fields').get();
    } catch (e) { /* nao achou pelo id, tenta por numero */ }
  }
  // FALLBACK CRUCIAL: se nao passou numero mas passou id, usa id como numero tambem
  // (resolve o caso da SOL passar "36534" como id quando na verdade eh NumeroNF)
  const numeroAlvo = args.numero || args.id;
  if (!item && numeroAlvo) {
    // Pagina por todas as NFs ate achar (max 30 paginas = 15k items)
    const alvoNum = String(numeroAlvo).replace(/\D/g, '');
    let url = '/sites/' + siteId + '/lists/' + listNotasId + '/items?expand=fields&$top=500';
    let pages = 0;
    while (url && pages < 30 && !item) {
      const resp = await client.api(url).get();
      for (const it of (resp.value || [])) {
        const n = normalizeItem(it, invColMap);
        const numStr = String(n.NumeroNF || '');
        // Match exato OU normalizado (so digitos) pra tolerar formatacao tipo "NF-48" vs "48"
        if (numStr === String(numeroAlvo) || numStr.replace(/\D/g,'') === alvoNum) {
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
  // Resolve nome do fornecedor via CNPJ (lookup na lista Fornecedores 1x por sessao)
  let fornNome = '';
  try {
    if (!ctx._fornMap) {
      ctx._fornMap = await carregarMapFornecedoresParaSol(ctx.gr.client, ctx.gr.siteId);
    }
    const cnpj = String(n.CNPJFornecedor || '').replace(/\D/g, '');
    fornNome = (cnpj && ctx._fornMap[cnpj]) || String(n.CNPJFornecedor || '');
  } catch (e) { fornNome = String(n.CNPJFornecedor || ''); }

  return {
    id: n._id,
    numero: n.NumeroNF,
    fornecedor: fornNome || '(sem nome)',
    cnpj: n.CNPJFornecedor || '',
    valor: Number(n.Valor || 0),
    vencimento: String(n.DataVencimento || '').substring(0,10),
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

async function tool_abrir_nf(args, ctx) {
  // Resolve id se vier numero
  let id = args.id;
  if (!id && args.numero) {
    const det = await tool_detalhes_nf({ numero: args.numero }, ctx);
    if (det && det.erro) return det;
    if (det && det.id) id = det.id;
  }
  if (!id) return { erro: 'precisa de id ou numero da NF' };
  return {
    acao_imediata: 'abrir_pdf',
    id: id,
    url: '/api/AbrirPdfDaNota?id=' + encodeURIComponent(id),
    confirmar_no_frontend: false
  };
}

async function tool_propor_marcar_processado(args, ctx) {
  if (!args.id) return { erro: 'id obrigatorio' };
  const det = await tool_detalhes_nf({ id: args.id }, ctx);
  if (det && det.erro) return det;
  return {
    confirmar_no_frontend: true,
    tipo: 'processado',
    id: det.id,
    numero: det.numero,
    fornecedor: det.fornecedor,
    valor: det.valor
  };
}

const TOOL_IMPL = {
  listar_fila: tool_listar_fila,
  listar_aprovadas: tool_listar_aprovadas,
  detalhes_nf: tool_detalhes_nf,
  agregar_por_fornecedor: tool_agregar_por_fornecedor,
  detectar_anomalia: tool_detectar_anomalia,
  propor_aprovacao: tool_propor_aprovacao,
  propor_rejeicao: tool_propor_rejeicao,
  abrir_nf: tool_abrir_nf,
  propor_marcar_processado: tool_propor_marcar_processado
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
// =============================================================================
// Converte TOOLS do formato OpenAI pro formato Anthropic
// =============================================================================
function toolsParaAnthropic(tools) {
  return tools.map(t => ({
    name: t.function.name,
    description: t.function.description,
    input_schema: t.function.parameters || { type: 'object', properties: {} }
  }));
}

// =============================================================================
// LOOP PRINCIPAL: tenta Anthropic primeiro, cai pra OpenAI se falhar
// =============================================================================
async function runSol(history, userMessage, user, opts) {
  opts = opts || {};
  const maxIter = opts.maxIter || 8;
  const viewAtual = opts.viewAtual || 'fila-aprovacao';
  const isAdmin = !!opts.isAdmin;

  const client = await getGraphClient();
  const { siteId, listNotasId, invColMap } = await resolveSiteAndLists(client);
  const ctx = { user, isAdmin, gr: { client, siteId, listNotasId, invColMap } };
  const systemPrompt = buildSystemPrompt(user, viewAtual);

  // Tenta Anthropic primeiro
  const anthropic = getAnthropic();
  let anthropicError = null;
  if (anthropic) {
    try {
      return await runSolAnthropic(anthropic, history, userMessage, systemPrompt, ctx, maxIter, opts.model, getToolsForView(viewAtual));
    } catch (e) {
      anthropicError = {
        message: e.message || String(e),
        status: e.status || (e.response && e.response.status),
        type: e.constructor && e.constructor.name,
        stack: (e.stack || '').split('\n').slice(0, 5).join(' | ')
      };
      console.error('[SOL] Anthropic falhou:', anthropicError);
      // Cai pro fallback OpenAI
    }
  } else {
    anthropicError = { message: 'ANTHROPIC_API_KEY nao configurada ou Anthropic SDK nao carregou' };
  }

  // Fallback OpenAI
  const openai = getOpenAI();
  if (!openai) {
    const loadErrs = getLoadErrors();
    throw new Error('Nenhum provider IA disponivel. anthropic_error=' + JSON.stringify(anthropicError) + ' load_errors=' + JSON.stringify(loadErrs));
  }
  const result = await runSolOpenAI(openai, history, userMessage, systemPrompt, ctx, maxIter, getToolsForView(viewAtual));
  // Anexa o erro do Anthropic no response pra debug
  result.anthropic_error = anthropicError;
  return result;
}

// =============================================================================
// Anthropic — formato com content blocks + tool_use/tool_result
// =============================================================================
async function runSolAnthropic(client, history, userMessage, systemPrompt, ctx, maxIter, modelOverride, viewTools) {
  // Defesa em profundidade: se vier um modelo OpenAI (gpt-*) por engano, ignora e usa o default Anthropic.
  const safeOverride = (modelOverride && !/^gpt[-_]/i.test(String(modelOverride))) ? modelOverride : null;
  let model = safeOverride || DEFAULT_MODEL;
  // Tools filtradas por view (escopo da Fase C). Se viewTools nao foi passado, mantem todas (compat).
  const effectiveTools = Array.isArray(viewTools) ? viewTools : TOOLS;
  const anthropicTools = toolsParaAnthropic(effectiveTools);

  // Anthropic messages: alterna user/assistant. system vai como param separado.
  const messages = [];
  for (const h of (history || [])) {
    if (!h || !h.role) continue;
    if (h.role === 'user' || h.role === 'assistant') {
      messages.push({ role: h.role, content: String(h.content || '') });
    }
  }
  messages.push({ role: 'user', content: userMessage });

  const acoesPropostas = [];
  const toolCallsDebug = [];
  let totalTokens = 0;
  let resposta = '';
  let toolCallsCount = 0;

  for (let i = 0; i < maxIter; i++) {
    // Escalation: se ja chamou >3 tools nesta turn, sobe pra Sonnet
    if (toolCallsCount >= 3 && model === MODEL_HAIKU) {
      model = MODEL_SONNET;
      console.log('[SOL] Escalando Haiku → Sonnet apos', toolCallsCount, 'tool calls');
    }

    const completion = await client.messages.create({
      model: model,
      max_tokens: 2000,
      temperature: 0.2,
      system: systemPrompt,
      messages: messages,
      ...(anthropicTools.length > 0 ? { tools: anthropicTools } : {})
    });

    totalTokens += (completion.usage && (completion.usage.input_tokens + completion.usage.output_tokens)) || 0;

    // Adiciona resposta do assistant ao historico (preservando content blocks)
    messages.push({ role: 'assistant', content: completion.content });

    // Se stop_reason !== 'tool_use', terminou — extrai texto
    if (completion.stop_reason !== 'tool_use') {
      const textBlocks = (completion.content || []).filter(b => b.type === 'text');
      resposta = textBlocks.map(b => b.text).join('\n').trim();
      break;
    }

    // Executa as tools chamadas
    const toolUseBlocks = (completion.content || []).filter(b => b.type === 'tool_use');
    const toolResults = [];

    for (const tu of toolUseBlocks) {
      const fnName = tu.name;
      const args = tu.input || {};
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
      toolCallsCount++;
      toolCallsDebug.push({
        tool: fnName,
        args: args,
        resultSummary: result && result.erro ? { erro: result.erro } :
                       (result && Array.isArray(result.notas) ? { totalNotas: result.notas.length, primeiras3: result.notas.slice(0,3).map(n => ({ id: n.id, numero: n.numero })) } :
                       (result && result.id ? { id: result.id, numero: result.numero, fornecedor: result.fornecedor } :
                       { tipo: typeof result }))
      });
      if (result && (result.confirmar_no_frontend || result.acao_imediata)) {
        acoesPropostas.push(result);
      }
      toolResults.push({
        type: 'tool_result',
        tool_use_id: tu.id,
        content: JSON.stringify(result)
      });
    }

    // Continua a conversa com os resultados
    messages.push({ role: 'user', content: toolResults });
  }

  return { resposta: resposta, acoes_propostas: acoesPropostas, tokens: totalTokens, model: model, tool_calls_debug: toolCallsDebug, provider: 'anthropic' };
}

// =============================================================================
// OpenAI — formato function calling (fallback caso Anthropic falhe)
// =============================================================================
async function runSolOpenAI(openai, history, userMessage, systemPrompt, ctx, maxIter, viewTools) {
  const effectiveTools = Array.isArray(viewTools) ? viewTools : TOOLS;
  const messages = [{ role: 'system', content: systemPrompt }];
  for (const h of (history || [])) {
    if (!h || !h.role) continue;
    if (h.role === 'user' || h.role === 'assistant') {
      messages.push({ role: h.role, content: h.content || '' });
    }
  }
  messages.push({ role: 'user', content: userMessage });

  const acoesPropostas = [];
  const toolCallsDebug = [];
  let totalTokens = 0;
  let resposta = '';

  for (let i = 0; i < maxIter; i++) {
    const completion = await openai.chat.completions.create({
      model: OPENAI_MODEL,
      messages: messages,
      tools: effectiveTools.length > 0 ? effectiveTools : undefined,
      tool_choice: effectiveTools.length > 0 ? 'auto' : undefined,
      temperature: 0.2,
      max_tokens: 1500
    });
    totalTokens += (completion.usage && completion.usage.total_tokens) || 0;
    const msg = completion.choices[0].message;
    messages.push(msg);
    if (!msg.tool_calls || msg.tool_calls.length === 0) {
      resposta = msg.content || '';
      break;
    }
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
      toolCallsDebug.push({
        tool: fnName,
        args: args,
        resultSummary: result && result.erro ? { erro: result.erro } :
                       (result && Array.isArray(result.notas) ? { totalNotas: result.notas.length, primeiras3: result.notas.slice(0,3).map(n => ({ id: n.id, numero: n.numero })) } :
                       (result && result.id ? { id: result.id, numero: result.numero, fornecedor: result.fornecedor } :
                       { tipo: typeof result }))
      });
      if (result && (result.confirmar_no_frontend || result.acao_imediata)) {
        acoesPropostas.push(result);
      }
      messages.push({
        role: 'tool',
        tool_call_id: tc.id,
        content: JSON.stringify(result)
      });
    }
  }

  return { resposta: resposta, acoes_propostas: acoesPropostas, tokens: totalTokens, model: OPENAI_MODEL, tool_calls_debug: toolCallsDebug, provider: 'openai-fallback' };
}

module.exports = { runSol, DEFAULT_MODEL, getAnthropic, getOpenAI, getLoadErrors };
