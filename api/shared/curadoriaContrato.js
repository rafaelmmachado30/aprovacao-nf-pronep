/**
 * shared/curadoriaContrato.js — Curadoria/extração canônica de um contrato (Fase 1).
 *
 * Lê o texto de UM documento e devolve o modelo canônico (Fase 0) em JSON, com:
 *  - classificação do documento (contrato | aditivo | glosa | outro) p/ filtrar ruído;
 *  - VIGENTE vs HISTORICO;
 *  - operadora/estado/vigência SEMPRE do CONTEÚDO (nunca da pasta/nome);
 *  - selo de confiança e confianca_preco (NATIVO p/ texto nativo).
 *
 * Zero invenção: campo não encontrado = null. Anthropic (Sonnet p/ qualidade) com
 * fallback OpenAI. Best-effort: retorna null se ambos falharem (nunca lança).
 */

require('isomorphic-fetch');

// Haiku por padrao: Sonnet estoura os 45s fixos do gateway SWA em contratos grandes.
// Para reprocessar com Sonnet (mais qualidade) quando houver fila/durable, basta setar
// ANTHROPIC_MODEL_CURADORIA com o id do Sonnet.
const MODEL = process.env.ANTHROPIC_MODEL_CURADORIA || process.env.ANTHROPIC_MODEL_HAIKU || 'claude-haiku-4-5-20251001';
const OPENAI_MODEL = process.env.OPENAI_MODEL || 'gpt-4o-mini';
const MAX_TEXTO = 28000; // enxuto p/ caber no timeout do gateway (SWA ~45s) por documento

const INSTRUCAO =
  'Voce e um curador juridico de contratos de operadoras de saude (home care). Le o documento e ' +
  'devolve SOMENTE um JSON valido (sem markdown, sem texto fora do JSON) no schema abaixo. ' +
  'REGRAS: (1) Zero invencao: campo nao encontrado = null; nunca chute. (2) operadora, estado (SP/RJ/ES) e ' +
  'vigencia vem do CONTEUDO do documento, JAMAIS inferidos de pasta/nome de arquivo. (3) Classifique doc_tipo ' +
  'em contrato | aditivo | glosa | outro (recurso de glosa, cadastral, proposta sem vigencia = "glosa"/"outro"). ' +
  '(4) status = VIGENTE se ainda vigente na data de hoje, senao HISTORICO. (5) So reajuste GERAL da tabela define ' +
  'data_base; pontual e observacao. (6) qualidade_redacao_score de 0 a 5 (clareza/completude juridica). ' +
  'SCHEMA: {' +
  '"doc_tipo":"contrato|aditivo|glosa|outro",' +
  '"operadora":{"nome":"","cnpj":null,"registro_ans":null,"segmento":null},' +
  '"estado_uf":"SP|RJ|ES|null",' +
  '"contrato":{"numero":null,"cnpj_pronep_contratada":null,"objeto":"","data_assinatura":null,"inicio_vigencia":null,"fim_vigencia":null,"prazo":"determinado|indeterminado|null","status":"VIGENTE|HISTORICO","tem_liminar":false,"qualidade_redacao_score":0,"qualidade_redacao_notas":""},' +
  '"completude":{"tem_lgpd":false,"tem_rescisao":false,"tem_reajuste":false,"tem_sla":false},' +
  '"reajustes":[{"numero_rp":null,"tipo":"GERAL|pontual","indice":null,"percentual":null,"data_base":null,"proximo_reajuste_previsto":null}],' +
  '"aditivos":[{"objeto":"","inicio_vigencia":null,"tipo":"pontual|geral","melhora_para_pronep":"sim|nao|neutro","melhora_justificativa":""}],' +
  '"diarias":[{"descricao":"","valor_diaria":null,"inclusos":[],"exclusos":[],"vigencia_inicio":null,"vigencia_fim":null}],' +
  '"procedimentos":[{"codigo_tuss":null,"descricao":"","valor":null,"unidade":null}],' +
  '"matmed":[{"base":"Simpro|Brasindice","operador":"deflator|acrescimo","percentual":null,"categoria":""}],' +
  '"clausulas":[{"tipo":"vigencia|reajuste|glosa|rescisao|faturamento|LGPD|reembolso|exclusividade|liminar","texto_literal":"","paragrafo_ref":null}],' +
  '"riscos":[{"descricao":"","categoria":"financeiro|juridico|operacional","severidade":"alto|medio|baixo"}]' +
  '}. Responda APENAS o JSON.';

function _parseJson(txt) {
  if (!txt) return null;
  let s = String(txt).trim().replace(/^```(json)?/i, '').replace(/```$/i, '').trim();
  const a = s.indexOf('{'), b = s.lastIndexOf('}');
  if (a >= 0 && b > a) s = s.slice(a, b + 1);
  try { return JSON.parse(s); } catch (e) { return null; }
}

// Ids dos modelos (mesmos strings usados pela SAN).
const MODEL_HAIKU = process.env.ANTHROPIC_MODEL_HAIKU || 'claude-haiku-4-5-20251001';
const MODEL_SONNET = process.env.ANTHROPIC_MODEL_SONNET || 'claude-sonnet-4-6';
const MODEL_OPUS = process.env.ANTHROPIC_MODEL_OPUS || 'claude-opus-4-8';
const FORCE_MODEL = process.env.ANTHROPIC_MODEL_CURADORIA || null; // se setado, manda
const SMALL_CHARS = 8000; // ate aqui, Sonnet cabe nos 45s do gateway

function _resolveModelo(alias) {
  if (!alias) return null;
  const a = String(alias).toLowerCase();
  if (a === 'haiku') return MODEL_HAIKU;
  if (a === 'sonnet') return MODEL_SONNET;
  if (a === 'opus') return MODEL_OPUS;
  return alias; // ja e um id completo
}

// Escada de modelos: do melhor que cabe -> fallback pra baixo (resiliencia).
// override (?model=) tem prioridade; senao, decide pelo TAMANHO do texto (tempo).
function _escadaModelos(len, override) {
  const desc = [MODEL_OPUS, MODEL_SONNET, MODEL_HAIKU];
  const alvo = _resolveModelo(override) || FORCE_MODEL;
  if (alvo) { const i = desc.indexOf(alvo); return i >= 0 ? desc.slice(i) : [alvo, MODEL_HAIKU]; }
  if (len <= SMALL_CHARS) return [MODEL_SONNET, MODEL_HAIKU]; // pequeno: tenta Sonnet
  return [MODEL_HAIKU];                                      // grande: so Haiku cabe no tempo
}

async function _viaAnthropic(texto, model) {
  const key = process.env.ANTHROPIC_API_KEY;
  if (!key) return null;
  const mod = require('@anthropic-ai/sdk');
  const Anthropic = mod.default || mod.Anthropic || mod;
  const client = new Anthropic({ apiKey: key });
  const maxTokens = (model === MODEL_HAIKU) ? 4000 : 3000; // limita geracao dos lentos
  const resp = await client.messages.create({
    model: model, max_tokens: maxTokens, temperature: 0,
    system: INSTRUCAO,
    messages: [{ role: 'user', content: 'Documento:\n\n' + texto.slice(0, MAX_TEXTO) }]
  });
  const t = (resp.content || []).filter(function (b) { return b.type === 'text'; }).map(function (b) { return b.text; }).join('');
  return _parseJson(t);
}

async function _viaOpenAI(texto) {
  const key = process.env.OPENAI_API_KEY;
  if (!key) return null;
  const OpenAI = require('openai');
  const Ctor = OpenAI.default || OpenAI.OpenAI || OpenAI;
  const client = new Ctor({ apiKey: key });
  const resp = await client.chat.completions.create({
    model: OPENAI_MODEL, temperature: 0, response_format: { type: 'json_object' },
    messages: [{ role: 'system', content: INSTRUCAO }, { role: 'user', content: 'Documento (JSON):\n\n' + texto.slice(0, MAX_TEXTO) }]
  });
  const t = resp.choices && resp.choices[0] && resp.choices[0].message && resp.choices[0].message.content;
  return _parseJson(t);
}

let _ultimoErro = null;
function ultimoErroCuradoria() { return _ultimoErro; }

// Retorna o objeto canonico (ou null). metodo = 'nativo'|'OCR' (define selo/confianca_preco).
// opts.model = 'haiku'|'sonnet'|'opus'|<id> (override; senao escada por tamanho).
async function curar(texto, metodo, opts) {
  _ultimoErro = null;
  opts = opts || {};
  if (!texto || texto.length < 60) { _ultimoErro = 'texto curto'; return null; }

  const escada = _escadaModelos(texto.length, opts.model);
  let canonico = null;
  let modeloUsado = null;
  const erros = [];
  for (const m of escada) {
    try {
      const c = await _viaAnthropic(texto, m);
      if (c) { canonico = c; modeloUsado = m; break; }
      erros.push(m + ': JSON invalido');
    } catch (e) { erros.push(m + ': ' + ((e && e.message) || String(e))); }
  }
  if (!canonico) {
    try { canonico = await _viaOpenAI(texto); if (canonico) modeloUsado = OPENAI_MODEL; else erros.push('openai: JSON invalido'); }
    catch (e) { erros.push('openai: ' + ((e && e.message) || String(e))); }
  }
  if (!canonico) { _ultimoErro = erros.join(' | '); return null; }

  // Selo/confianca a partir do metodo de extracao (nativo = texto digital).
  const nativo = (metodo || 'nativo') === 'nativo';
  const conf = nativo ? 'NATIVO' : 'OCR-A-VALIDAR';
  (canonico.diarias || []).forEach(function (d) { if (!d.confianca_preco) d.confianca_preco = conf; });
  (canonico.procedimentos || []).forEach(function (p) { if (!p.confianca_preco) p.confianca_preco = conf; });
  canonico.proveniencia = canonico.proveniencia || {};
  canonico.proveniencia.metodo_extracao = metodo || 'nativo';
  canonico.proveniencia.modelo_curadoria = modeloUsado;
  // Uma fonte so (sem cruzamento ainda) => PARCIAL; se doc nao-nativo => PENDENTE.
  canonico.proveniencia.selo_confianca = nativo ? 'PARCIAL' : 'PENDENTE';
  return canonico;
}

module.exports = { curar, ultimoErroCuradoria, MODEL };
