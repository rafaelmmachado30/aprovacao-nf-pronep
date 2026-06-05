/**
 * shared/contratos.js — Modulo de Controle de Contratos da Pronep.
 *
 * Crawler recursivo do SharePoint "CONTRATOS-SERVICOS-CONTRATOS" + extracao
 * de vigencia (DataInicio/DataFim) via Claude Haiku com escalacao pra Sonnet
 * quando a confianca for baixa. Persiste em PRONEP-NF-Contratos.
 *
 * Estrutura do SP:
 *   /Shared Documents/CONTRATOS/CONTRATOS E DOCUMENTOS - PRESTADORES/
 *     <Diretoria>/
 *       <Unidade ou subpasta>/
 *         <Prestador>/
 *           contrato.pdf, aditivos, etc.
 *
 * Mapeamento de diretorias do SP -> diretorias do sistema NF:
 *   - DIRETORIA COMERCIAL       -> Comercial
 *   - DIRETORIA DE OPERACOES    -> Operacoes
 *   - DIRETORIA DE SUPRIMENTOS  -> Suprimentos
 *   - DIRETORIA FINANCEIRA      -> Financeira
 *   - GERENCIA DE PROJETOS E TI -> Tecnica  (TI esta sob Tecnica no NF)
 *   - GERENCIA DE RH            -> RH
 *   - JURIDICO                  -> Juridica
 *   - OUVIDORIA                 -> Ouvidoria
 *   - PACIENTES PARTICULARES    -> Particulares
 *   - QUALIDADE                 -> Qualidade
 *
 * Env vars necessarias:
 *   SHAREPOINT_CONTRATOS_HOSTNAME (default: pronepadmin.sharepoint.com)
 *   SHAREPOINT_CONTRATOS_PATH     (default: /sites/CONTRATOS-SERVICOS-CONTRATOS)
 *   ANTHROPIC_API_KEY             (ja usado pelo SAN)
 */

require('isomorphic-fetch');
const { ClientSecretCredential } = require('@azure/identity');
const { Client } = require('@microsoft/microsoft-graph-client');
const { TokenCredentialAuthenticationProvider } =
  require('@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials');
const Anthropic = require('@anthropic-ai/sdk');
const pdfParse = require('pdf-parse');

const LIST_CONTRATOS = 'PRONEP-NF-Contratos';
const ROOT_FOLDER_PATH = '/CONTRATOS/CONTRATOS E DOCUMENTOS - PRESTADORES';

// Cache de site/list pra economizar chamadas
const cache = {
  contratoSite: null,
  contratoListId: null,
  contratoColMap: null,
  driveId: null
};

// ============================================================================
// GRAPH CLIENT (SITE DE CONTRATOS)
// ============================================================================

function getGraphClient() {
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

async function resolveContratosSite(client) {
  if (cache.contratoSite && cache.driveId) return cache;
  const host = process.env.SHAREPOINT_CONTRATOS_HOSTNAME || 'pronepadmin.sharepoint.com';
  const path = process.env.SHAREPOINT_CONTRATOS_PATH || '/sites/CONTRATOS-SERVICOS-CONTRATOS';
  const siteResp = await client.api('/sites/' + host + ':' + path).get();
  cache.contratoSite = siteResp.id;
  const driveResp = await client.api('/sites/' + siteResp.id + '/drive').get();
  cache.driveId = driveResp.id;
  return cache;
}

// ============================================================================
// LISTA DE CONTROLE (PRONEP-NF-Contratos)
// ============================================================================

/**
 * Cria a lista PRONEP-NF-Contratos no SITE DO SISTEMA NF (nao no de contratos).
 * Operacao idempotente — se ja existe, retorna sem criar.
 */
async function garantirListaContratos(client) {
  // Reutiliza o site do sistema NF (mesmo das outras listas)
  const host = process.env.SHAREPOINT_SITE_HOSTNAME;
  const path = process.env.SHAREPOINT_SITE_PATH;
  if (!host || !path) throw new Error('SHAREPOINT_SITE_HOSTNAME/PATH nao configurados');

  const siteResp = await client.api('/sites/' + host + ':' + path).get();
  const siteId = siteResp.id;

  // Procura se ja existe
  const lists = await client.api('/sites/' + siteId + '/lists')
    .filter("displayName eq '" + LIST_CONTRATOS + "'").get();
  if (lists.value && lists.value.length) {
    cache.contratoListId = lists.value[0].id;
    return { siteId: siteId, listId: cache.contratoListId, criada: false };
  }

  // Cria a lista
  const newList = await client.api('/sites/' + siteId + '/lists').post({
    displayName: LIST_CONTRATOS,
    list: { template: 'genericList' },
    columns: [
      { name: 'Diretoria',          text: {} },
      { name: 'Unidade',            text: {} },
      { name: 'Fornecedor',         text: {} },
      { name: 'CNPJFornecedor',     text: {} },
      { name: 'DataInicio',         dateTime: { displayAs: 'standard', format: 'dateOnly' } },
      { name: 'DataFim',            dateTime: { displayAs: 'standard', format: 'dateOnly' } },
      { name: 'Status',             text: {} },
      { name: 'CaminhoSharepoint',  text: { allowMultipleLines: true } },
      { name: 'DriveItemId',        text: {} },
      { name: 'LeituraIAStatus',    text: {} },
      { name: 'LeituraIATexto',     text: { allowMultipleLines: true } },
      { name: 'ValorContrato',      number: {} },
      { name: 'Observacoes',        text: { allowMultipleLines: true } },
      { name: 'PathRelativoSP',     text: { allowMultipleLines: true } },
      { name: 'UltimaLeitura',      dateTime: { displayAs: 'standard', format: 'dateTime' } },
      { name: 'NomeArquivo',        text: {} },
      { name: 'TamanhoArquivo',     number: {} }
    ]
  });
  cache.contratoListId = newList.id;
  return { siteId: siteId, listId: newList.id, criada: true };
}

async function getContratoColMap(client, siteId, listId) {
  if (cache.contratoColMap) return cache.contratoColMap;
  const resp = await client.api('/sites/' + siteId + '/lists/' + listId + '/columns').get();
  const map = {};
  for (const c of (resp.value || [])) {
    if (c.displayName && c.name) map[c.displayName] = c.name;
  }
  cache.contratoColMap = map;
  return map;
}

// ============================================================================
// CRAWLER DO SHAREPOINT DE CONTRATOS
// ============================================================================

/**
 * Lista o conteudo de uma pasta do SP de contratos.
 * Retorna { folders: [], files: [] }.
 */
async function listarPasta(client, driveId, folderPath) {
  // folderPath formato: "/CONTRATOS/CONTRATOS E DOCUMENTOS - PRESTADORES/DIRETORIA X/..."
  let url;
  if (folderPath === '/' || folderPath === '') {
    url = '/drives/' + driveId + '/root/children?$top=200';
  } else {
    const enc = encodeURIComponent(folderPath).replace(/%2F/g, '/');
    url = '/drives/' + driveId + '/root:' + enc + ':/children?$top=200';
  }
  const resp = await client.api(url).get();
  const folders = [];
  const files = [];
  for (const it of (resp.value || [])) {
    if (it.folder) {
      folders.push({
        name: it.name,
        id: it.id,
        path: folderPath.replace(/\/$/, '') + '/' + it.name,
        childCount: it.folder.childCount || 0,
        webUrl: it.webUrl
      });
    } else if (it.file) {
      const ext = (it.name || '').split('.').pop().toLowerCase();
      // Politica: apenas PDFs (Word editavel fica fora do padrao Pronep)
      if (ext === 'pdf') {
        files.push({
          name: it.name,
          id: it.id,
          path: folderPath.replace(/\/$/, '') + '/' + it.name,
          size: it.size || 0,
          ext: ext,
          webUrl: it.webUrl,
          lastModified: it.lastModifiedDateTime
        });
      }
    }
  }
  return { folders, files };
}

/**
 * Faz crawl recursivo de uma pasta. Retorna lista plana de arquivos com metadata.
 * EARLY-STOP: para assim que `maxArquivos` for atingido (evita varrer dezenas de
 * pastas grandes quando o consumidor so precisa de N arquivos).
 *
 * @param {string} pastaRaiz - caminho relativo no drive
 * @param {object} opts - { maxDepth?, maxArquivos?, onProgress? }
 */
async function crawlPasta(client, driveId, pastaRaiz, opts) {
  opts = opts || {};
  const maxDepth = opts.maxDepth || 8;
  const maxArquivos = opts.maxArquivos || 999999;
  const onProgress = opts.onProgress || function(){};
  const resultado = [];
  let pastasVisitadas = 0;

  async function recurse(path, depth, ancestors) {
    if (depth > maxDepth) return;
    if (resultado.length >= maxArquivos) return;  // early-stop
    let listing;
    try {
      listing = await listarPasta(client, driveId, path);
      pastasVisitadas++;
    } catch (e) {
      onProgress({ tipo: 'erro_pasta', path, erro: e.message });
      return;
    }
    onProgress({ tipo: 'pasta', path, files: listing.files.length, folders: listing.folders.length, visitadas: pastasVisitadas });
    // Arquivos primeiro
    for (const f of listing.files) {
      if (resultado.length >= maxArquivos) return;
      resultado.push({
        nome: f.name,
        id: f.id,
        path: f.path,
        size: f.size,
        ext: f.ext,
        webUrl: f.webUrl,
        lastModified: f.lastModified,
        ancestors: ancestors.slice()
      });
    }
    // Recursivo
    for (const sub of listing.folders) {
      if (resultado.length >= maxArquivos) return;
      await recurse(sub.path, depth + 1, ancestors.concat([sub.name]));
    }
  }

  // ancestors iniciais: tudo entre ROOT_FOLDER_PATH e pastaRaiz.
  // Ex: pastaRaiz = "/CONTRATOS/CONTRATOS E DOCUMENTOS - PRESTADORES/GERÊNCIA DE PROJETOS E TI/CORPORATIVO/TOTVS SA"
  //     ROOT       = "/CONTRATOS/CONTRATOS E DOCUMENTOS - PRESTADORES"
  //     iniciais   = ['GERÊNCIA DE PROJETOS E TI', 'CORPORATIVO', 'TOTVS SA']
  // Sem isso, ao descer pra subpasta CONTRATOS o ancestors fica ['CONTRATOS'] e perde a hierarquia.
  let relRaiz = pastaRaiz;
  if (relRaiz.startsWith(ROOT_FOLDER_PATH)) relRaiz = relRaiz.slice(ROOT_FOLDER_PATH.length);
  const ancestorsIniciais = relRaiz.split('/').filter(function(s){ return s && s.length > 0; });

  await recurse(pastaRaiz, 0, ancestorsIniciais);
  return resultado;
}

// ============================================================================
// EXTRACAO DE TEXTO
// ============================================================================

async function baixarArquivo(client, driveId, itemId, opts) {
  opts = opts || {};
  const tamanhoMaxMB = opts.tamanhoMaxMB || 30;
  // Pega item SEM select — @microsoft.graph.downloadUrl eh "instance annotation"
  // e $select a oculta. Sem select, ela vem incluida automaticamente.
  const item = await client.api('/drives/' + driveId + '/items/' + itemId).get();
  const size = item.size || 0;
  if (size > tamanhoMaxMB * 1024 * 1024) {
    throw new Error('arquivo_grande_demais: ' + Math.round(size/1024/1024) + 'MB (max ' + tamanhoMaxMB + 'MB)');
  }
  const downloadUrl = item['@microsoft.graph.downloadUrl'];
  if (downloadUrl) {
    const fetch = require('isomorphic-fetch');
    const resp = await fetch(downloadUrl);
    if (!resp.ok) throw new Error('download HTTP ' + resp.status);
    const arrayBuffer = await resp.arrayBuffer();
    return Buffer.from(arrayBuffer);
  }
  // Fallback: /content endpoint via Graph client (faz 302 -> downloadUrl)
  try {
    const buf = await client.api('/drives/' + driveId + '/items/' + itemId + '/content').responseType('arraybuffer').get();
    return Buffer.from(buf);
  } catch (e) {
    throw new Error('Graph nao retornou downloadUrl e fallback /content falhou: ' + e.message);
  }
}

async function extrairTextoPDF(buffer) {
  try {
    const data = await pdfParse(buffer, { max: 30 }); // primeiras 30 paginas
    return { texto: data.text || '', paginas: data.numpages || 0, vazio: !(data.text && data.text.trim()) };
  } catch (e) {
    return { texto: '', paginas: 0, vazio: true, erro: e.message };
  }
}

// DOCX: usaremos mammoth se disponivel; senao retorna vazio sinalizando que precisa OCR/manual
async function extrairTextoDOCX(buffer) {
  try {
    const mammoth = require('mammoth');
    const result = await mammoth.extractRawText({ buffer: buffer });
    return { texto: result.value || '', vazio: !(result.value && result.value.trim()) };
  } catch (e) {
    return { texto: '', vazio: true, erro: 'mammoth nao instalado ou falhou: ' + e.message };
  }
}

async function extrairTexto(client, driveId, itemId, ext, opts) {
  const buf = await baixarArquivo(client, driveId, itemId, opts);
  if (ext === 'pdf') return await extrairTextoPDF(buf);
  // Politica Pronep: so PDFs. Outros formatos ignorados.
  return { texto: '', vazio: true, erro: 'formato_nao_suportado' };
}

// ============================================================================
// EXTRACAO DE VIGENCIA VIA CLAUDE
// ============================================================================

function getAnthropic() {
  const apiKey = process.env.ANTHROPIC_API_KEY;
  if (!apiKey) throw new Error('ANTHROPIC_API_KEY nao configurada');
  return new Anthropic({ apiKey: apiKey });
}

const PROMPT_VIGENCIA = [
  'Voce e um analista juridico EXTRAINDO A VIGENCIA DE UM CONTRATO.',
  '',
  'CONCEITOS CRITICOS — NAO CONFUNDA:',
  '• "VIGENCIA DO CONTRATO" = periodo em que o contrato esta valido executando (ex: "Vigencia de 12 meses", "Prazo de 24 meses prorrogavel", "Vigora ate 31/12/2027").',
  '• "VALIDADE DA PROPOSTA" = prazo limite pra ACEITAR uma proposta comercial (ex: "Proposta valida ate 31/08/2020"). ISSO NAO E VIGENCIA. Se o texto so menciona validade da proposta, retorne { "naoEncontrou": true, "motivo": "documento eh proposta sem vigencia contratual" }.',
  '',
  'O QUE BUSCAR (palavras-chave de vigencia REAL):',
  '• "Vigencia", "Prazo de vigencia", "Prazo do contrato"',
  '• "Vigora a partir de... pelo prazo de X meses/anos"',
  '• "Por prazo determinado de N (numero) anos"',
  '• "Validade dos servicos de X meses, renovavel automaticamente"',
  '• "Prazo indeterminado", "Renovacao automatica"',
  '',
  'REGRAS DE OUTPUT:',
  '1. Vigencia explicita com DataInicio e DataFim absolutas: { "dataInicio": "YYYY-MM-DD", "dataFim": "YYYY-MM-DD", "indeterminado": false, "confidence": "alto" }',
  '2. Vigencia relativa "N meses a contar da assinatura DD/MM/YYYY": calcule DataFim e retorne com confidence "alto".',
  '3. Vigencia indeterminada ou renovacao automatica: { "dataInicio": "YYYY-MM-DD"|null, "dataFim": null, "indeterminado": true, "confidence": "alto" }.',
  '4. Documento eh PROPOSTA COMERCIAL sem clausula de vigencia: { "naoEncontrou": true, "motivo": "proposta sem vigencia contratual explicita" }. NAO use validade da proposta como vigencia.',
  '5. Documento eh ADITIVO/NDA/aprovacao interna sem vigencia propria: { "naoEncontrou": true, "motivo": "documento sem clausula propria de vigencia (ex: aditivo, NDA, aprovacao interna)" }.',
  '6. Ambiguidade real (multiplas vigencias citadas, datas conflitantes): { "confidence": "baixo", ... }.',
  '7. SEMPRE inclua "trecho" com a citacao EXATA da clausula onde achou (max 400 chars). Se naoEncontrou=true, trecho fica null ou vazio.',
  '',
  'FORMATO DE SAIDA (JSON estrito, sem markdown):',
  '{',
  '  "dataInicio": "YYYY-MM-DD" | null,',
  '  "dataFim": "YYYY-MM-DD" | null,',
  '  "indeterminado": boolean,',
  '  "naoEncontrou": boolean,',
  '  "confidence": "alto" | "baixo",',
  '  "trecho": "string com a clausula original ou null",',
  '  "motivo": "se naoEncontrou=true, explique brevemente",',
  '  "valorContrato": number | null (valor mensal ou total se explicitado, em BRL),',
  '  "fornecedorIdentificado": "nome do fornecedor mencionado no contrato" | null,',
  '  "tipoDocumento": "contrato" | "proposta" | "aditivo" | "nda" | "outro"',
  '}',
  '',
  'APENAS O JSON. SEM EXPLICACOES ADICIONAIS, SEM MARKDOWN.'
].join('\n');

async function extrairVigenciaIA(texto, opts) {
  opts = opts || {};
  const anthropic = getAnthropic();
  // Limita o texto pra economizar tokens: pega comeco + fim (vigencia costuma estar nesses lugares)
  const textoLimitado = texto.length > 20000
    ? texto.slice(0, 12000) + '\n\n[...TRUNCADO...]\n\n' + texto.slice(-8000)
    : texto;

  const modelo = opts.modelo || 'claude-haiku-4-5-20251001';
  let resposta;
  try {
    resposta = await anthropic.messages.create({
      model: modelo,
      max_tokens: 600,
      messages: [{
        role: 'user',
        content: PROMPT_VIGENCIA + '\n\nTEXTO DO CONTRATO:\n' + textoLimitado
      }]
    });
  } catch (e) {
    return { erro: 'anthropic_error: ' + e.message, modelo };
  }
  const respText = (resposta.content && resposta.content[0] && resposta.content[0].text) || '';
  let parsed;
  try {
    // Remove fences markdown caso o modelo escape
    const limpo = respText.replace(/^```json\s*/i, '').replace(/```\s*$/, '').trim();
    parsed = JSON.parse(limpo);
  } catch (e) {
    return { erro: 'parse_error', raw: respText.slice(0, 500), modelo };
  }
  parsed._modelo = modelo;
  parsed._tokensIn = (resposta.usage && resposta.usage.input_tokens) || 0;
  parsed._tokensOut = (resposta.usage && resposta.usage.output_tokens) || 0;
  return parsed;
}

/**
 * Extrai vigencia com escalacao Haiku -> Sonnet quando Haiku reportar confidence baixo.
 */
async function extrairVigenciaInteligente(texto) {
  if (!texto || !texto.trim()) {
    return { naoEncontrou: true, motivo: 'texto vazio (PDF scaneado sem OCR ou arquivo corrompido)', _modelo: 'none' };
  }
  let r = await extrairVigenciaIA(texto, { modelo: 'claude-haiku-4-5-20251001' });
  if (r && r.confidence === 'baixo' && !r.naoEncontrou) {
    // Escala pra Sonnet
    const r2 = await extrairVigenciaIA(texto, { modelo: 'claude-sonnet-4-6' });
    if (r2 && !r2.erro) {
      r2._escalacao = true;
      return r2;
    }
  }
  return r;
}

// ============================================================================
// CLASSIFICACAO DE STATUS (ATIVO/VENCENDO/VENCIDO)
// ============================================================================

function calcularStatus(dataFim, indeterminado) {
  if (indeterminado) return 'Ativo';
  if (!dataFim) return 'SemVigencia';
  const hoje = new Date();
  hoje.setHours(0,0,0,0);
  const fim = new Date(dataFim);
  fim.setHours(0,0,0,0);
  const diff = Math.floor((fim - hoje) / (1000 * 60 * 60 * 24));
  if (diff < 0) return 'Vencido';
  if (diff <= 30) return 'Vencendo30';
  if (diff <= 60) return 'Vencendo60';
  if (diff <= 90) return 'Vencendo90';
  return 'Ativo';
}

function calcularDiasParaVencer(dataFim) {
  if (!dataFim) return null;
  const hoje = new Date();
  hoje.setHours(0,0,0,0);
  const fim = new Date(dataFim);
  fim.setHours(0,0,0,0);
  return Math.floor((fim - hoje) / (1000 * 60 * 60 * 24));
}

// ============================================================================
// MAPEAMENTO DE METADATA (Diretoria, Unidade, Fornecedor) A PARTIR DO PATH
// ============================================================================

// Mapa SEM acentos — usamos normalizeStr() na hora de comparar.
// Acentos no SP podem vir como Unicode composto (Ê = U+00CA) OU decomposto (E + U+0302).
// Pra evitar falha de match, normalizamos ambos pra ASCII puro.
const MAPA_DIRETORIA = {
  'DIRETORIA COMERCIAL': 'Comercial',
  'DIRETORIA DE OPERACOES': 'Operacoes',
  'DIRETORIA DE SUPRIMENTOS E LOGISTICA': 'Suprimentos',
  'DIRETORIA FINANCEIRA': 'Financeira',
  'GERENCIA DE PROJETOS E TI': 'Tecnologia',
  'GERENCIA DE RH': 'RH',
  'JURIDICO': 'Juridica',
  'OUVIDORIA': 'Ouvidoria',
  'PACIENTES PARTICULARES': 'Particulares',
  'QUALIDADE': 'Qualidade'
};

// Pastas "estruturais" que existem dentro do prestador mas NAO sao o nome do prestador.
const SUBPASTAS_ESTRUTURAIS = new Set(['CONTRATOS', 'DOCUMENTOS', 'ADITIVOS', 'PROPOSTAS', 'NDAS', 'OUTROS']);

function normalizeStr(s) {
  return String(s || '')
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')  // tira acentos (NFD + remove combining chars)
    .toUpperCase()
    .replace(/\s+/g, ' ').trim();
}

const UNIDADES_VALIDAS = ['SP', 'RJ', 'ES', 'CORPORATIVO'];

function classificarPath(ancestors) {
  // ancestors = ['GERÊNCIA DE PROJETOS E TI', 'CORPORATIVO', 'TOTVS SA', 'CONTRATOS']
  // Queremos: diretoria='Tecnica', unidade='CORPORATIVO', fornecedor='TOTVS SA'
  let diretoria = '';
  let unidade = 'CORPORATIVO';
  let fornecedor = '';

  for (let i = 0; i < ancestors.length; i++) {
    const aNorm = normalizeStr(ancestors[i]);
    if (MAPA_DIRETORIA[aNorm]) {
      diretoria = MAPA_DIRETORIA[aNorm];
      // proximo nivel pode ser unidade (CORPORATIVO|SP|RJ|ES) ou prestador direto
      const nextNorm = normalizeStr(ancestors[i+1]);
      let idxFornecedor;
      if (UNIDADES_VALIDAS.includes(nextNorm)) {
        unidade = nextNorm;
        idxFornecedor = i + 2;
      } else {
        idxFornecedor = i + 1;
      }
      // O proximo ancestor depois da unidade EH o fornecedor, EXCETO se for
      // uma subpasta estrutural (CONTRATOS, DOCUMENTOS, etc.) — nesse caso
      // pega o anterior.
      const candidato = ancestors[idxFornecedor] || '';
      if (SUBPASTAS_ESTRUTURAIS.has(normalizeStr(candidato))) {
        // Sao raras essas — o fornecedor real esta um nivel acima
        fornecedor = ancestors[idxFornecedor - 1] || candidato;
      } else {
        fornecedor = candidato;
      }
      break;
    }
  }
  // Fallback robusto: percorre ancestors de tras pra frente, pulando subpastas estruturais
  if (!fornecedor && ancestors.length) {
    for (let i = ancestors.length - 1; i >= 0; i--) {
      const aNorm = normalizeStr(ancestors[i]);
      if (!SUBPASTAS_ESTRUTURAIS.has(aNorm) && !MAPA_DIRETORIA[aNorm] && !UNIDADES_VALIDAS.includes(aNorm)) {
        fornecedor = ancestors[i];
        break;
      }
    }
  }
  return { diretoria, unidade, fornecedor };
}

// ============================================================================
// FILTRO DE RELEVANCIA — economiza tokens Claude pulando arquivos que claramente
// NAO sao contratos comerciais (docs cadastrais, certidoes, emails, etc).
// ============================================================================
const PATTERNS_IGNORAR_CONTRATO = [
  /documentos[_ ]unificados/i,
  /(^|[ _-])cnpj([_ .-]|$)/i,             // "CNPJ_", " CNPJ ", "-CNPJ.pdf"
  /cart[aã]o[_ -]?cnpj/i,                 // "CARTÃO CNPJ" / "cartao CNPJ"
  /(^|[ _-])contrato[_ ]social/i,
  /estatuto[_ -]?social/i,                // "estatuto e ata consolidado"
  /\bestatuto[_ -]?e[_ -]?ata/i,
  /inscri[cç][aã]o[_ -]?municipal/i,
  /inscri[cç][aã]o[_ -]?estadual/i,
  /licenciamento[_ -]integrado/i,
  /licenciamento[_ -]prefeitura/i,
  /(^|[ _-])alvar[aá]([_ .-]|$)/i,
  /(^|[ _-])certid[aã]o([_ .-]|$)/i,
  /(^|[ _-])comprovante[_ -]/i,
  /\bemail[_ -]/i,
  /(^|[ _-])rg([_ .-]|$)/i,
  /(^|[ _-])cpf([_ .-]|$)/i,
  /procura[cç][aã]o[_ -]?simples/i
];

function eRelevantePraContrato(nomeArquivo) {
  const nome = String(nomeArquivo || '');
  for (const re of PATTERNS_IGNORAR_CONTRATO) {
    if (re.test(nome)) return { relevante: false, motivo: 'nome_filtrado: ' + re.toString() };
  }
  return { relevante: true };
}

module.exports = {
  getGraphClient,
  resolveContratosSite,
  garantirListaContratos,
  getContratoColMap,
  listarPasta,
  crawlPasta,
  extrairTexto,
  extrairTextoPDF,
  extrairVigenciaIA,
  extrairVigenciaInteligente,
  calcularStatus,
  calcularDiasParaVencer,
  classificarPath,
  normalizeStr,
  eRelevantePraContrato,
  PATTERNS_IGNORAR_CONTRATO,
  ROOT_FOLDER_PATH,
  LIST_CONTRATOS
};
