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

const MODEL = process.env.ANTHROPIC_MODEL_CURADORIA || process.env.ANTHROPIC_MODEL_SONNET || 'claude-sonnet-4-6';
const OPENAI_MODEL = process.env.OPENAI_MODEL || 'gpt-4o-mini';
const MAX_TEXTO = 45000; // Sonnet aguenta bem; cobre a maioria dos contratos

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

async function _viaAnthropic(texto) {
  const key = process.env.ANTHROPIC_API_KEY;
  if (!key) return null;
  const mod = require('@anthropic-ai/sdk');
  const Anthropic = mod.default || mod.Anthropic || mod;
  const client = new Anthropic({ apiKey: key });
  const resp = await client.messages.create({
    model: MODEL, max_tokens: 4000, temperature: 0,
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
async function curar(texto, metodo) {
  _ultimoErro = null;
  if (!texto || texto.length < 60) { _ultimoErro = 'texto curto'; return null; }
  let canonico = null;
  try { canonico = await _viaAnthropic(texto); if (!canonico) _ultimoErro = 'anthropic: JSON invalido'; }
  catch (e) { _ultimoErro = 'anthropic: ' + ((e && e.message) || String(e)); }
  if (!canonico) {
    try { canonico = await _viaOpenAI(texto); if (!canonico) _ultimoErro = (_ultimoErro || '') + ' | openai: JSON invalido'; }
    catch (e) { _ultimoErro = (_ultimoErro || '') + ' | openai: ' + ((e && e.message) || String(e)); }
  }
  if (!canonico) return null;

  // Selo/confianca a partir do metodo de extracao (nativo = texto digital).
  const nativo = (metodo || 'nativo') === 'nativo';
  const conf = nativo ? 'NATIVO' : 'OCR-A-VALIDAR';
  (canonico.diarias || []).forEach(function (d) { if (!d.confianca_preco) d.confianca_preco = conf; });
  (canonico.procedimentos || []).forEach(function (p) { if (!p.confianca_preco) p.confianca_preco = conf; });
  canonico.proveniencia = canonico.proveniencia || {};
  canonico.proveniencia.metodo_extracao = metodo || 'nativo';
  // Uma fonte so (sem cruzamento ainda) => PARCIAL; se doc nao-nativo => PENDENTE.
  canonico.proveniencia.selo_confianca = nativo ? 'PARCIAL' : 'PENDENTE';
  return canonico;
}

module.exports = { curar, ultimoErroCuradoria, MODEL };
