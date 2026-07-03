/**
 * shared/fichaContrato.js — Extrai uma FICHA ESTRUTURADA de um contrato (Fase 3 do RAG).
 *
 * Roda 1x por contrato NA INDEXACAO (barato, Haiku). Le o texto e devolve um JSON
 * normalizado com os campos que importam pra comparar/ranquear negociacao/rentabilidade.
 * Com essas fichas, a SAN compara a CARTEIRA inteira (coisa que a busca por trechos
 * nao faz). Fallback OpenAI se a Anthropic falhar; retorna null se ambos falharem.
 */

require('isomorphic-fetch');

const MODEL_HAIKU = process.env.ANTHROPIC_MODEL_HAIKU || 'claude-haiku-4-5-20251001';
const OPENAI_MODEL = process.env.OPENAI_MODEL || 'gpt-4o-mini';
const MAX_TEXTO = 14000; // trunca p/ controlar custo

const CAMPOS = [
  'fornecedor', 'objeto', 'valorMensal', 'valorTotal', 'moeda',
  'vigenciaInicio', 'vigenciaFim', 'renovacaoAutomatica',
  'reajusteIndice', 'reajusteFrequencia',
  'prazoPagamentoDias', 'multaRescisao', 'avisoPrevioDias', 'jurosMora',
  'sla', 'exclusividade', 'garantias', 'riscos', 'resumo'
];

const INSTRUCAO =
  'Voce extrai dados de contratos comerciais e devolve SOMENTE um objeto JSON valido (sem texto fora do JSON, sem markdown). ' +
  'Campos (use null quando nao encontrar; NAO invente):\n' +
  '- fornecedor (string), objeto (string curta: o que o contrato contrata)\n' +
  '- valorMensal (number|null, R$), valorTotal (number|null, R$), moeda (ex "BRL")\n' +
  '- vigenciaInicio (YYYY-MM-DD|null), vigenciaFim (YYYY-MM-DD|null), renovacaoAutomatica (true|false|null)\n' +
  '- reajusteIndice (ex "IPCA","IGPM","INPC"|null), reajusteFrequencia (ex "anual"|null)\n' +
  '- prazoPagamentoDias (number|null: prazo p/ pagar apos NF/fatura)\n' +
  '- multaRescisao (string curta descrevendo a multa|null), avisoPrevioDias (number|null), jurosMora (string|null)\n' +
  '- sla (string curta de nivel de servico/penalidade|null), exclusividade (true|false|null)\n' +
  '- garantias (string curta|null), riscos (string curta: pontos de atencao pra Pronep|null)\n' +
  '- resumo (1 frase objetiva do contrato)\n' +
  'Responda APENAS o JSON.';

function _parseJson(txt) {
  if (!txt) return null;
  let s = String(txt).trim();
  // remove cercas de codigo
  s = s.replace(/^```(json)?/i, '').replace(/```$/i, '').trim();
  // pega do primeiro { ao ultimo }
  const a = s.indexOf('{'), b = s.lastIndexOf('}');
  if (a >= 0 && b > a) s = s.slice(a, b + 1);
  try { return JSON.parse(s); } catch (e) { return null; }
}

function _normalizar(obj) {
  if (!obj || typeof obj !== 'object') return null;
  const out = {};
  for (const k of CAMPOS) out[k] = (obj[k] === undefined ? null : obj[k]);
  return out;
}

async function _viaAnthropic(texto) {
  const key = process.env.ANTHROPIC_API_KEY;
  if (!key) return null;
  const mod = require('@anthropic-ai/sdk');
  const Anthropic = mod.default || mod.Anthropic || mod;
  const client = new Anthropic({ apiKey: key });
  const resp = await client.messages.create({
    model: MODEL_HAIKU,
    max_tokens: 900,
    temperature: 0,
    system: INSTRUCAO,
    messages: [{ role: 'user', content: 'Contrato:\n\n' + texto.slice(0, MAX_TEXTO) }]
  });
  const t = (resp.content || []).filter(function (b) { return b.type === 'text'; }).map(function (b) { return b.text; }).join('');
  return _normalizar(_parseJson(t));
}

async function _viaOpenAI(texto) {
  const key = process.env.OPENAI_API_KEY;
  if (!key) return null;
  const OpenAI = require('openai');
  const Ctor = OpenAI.default || OpenAI.OpenAI || OpenAI;
  const client = new Ctor({ apiKey: key });
  const resp = await client.chat.completions.create({
    model: OPENAI_MODEL,
    temperature: 0,
    response_format: { type: 'json_object' },
    messages: [
      { role: 'system', content: INSTRUCAO },
      { role: 'user', content: 'Contrato:\n\n' + texto.slice(0, MAX_TEXTO) }
    ]
  });
  const t = resp.choices && resp.choices[0] && resp.choices[0].message && resp.choices[0].message.content;
  return _normalizar(_parseJson(t));
}

// Retorna a ficha normalizada ou null (best-effort — nunca lanca).
async function extrairFicha(texto) {
  if (!texto || texto.length < 40) return null;
  try { const a = await _viaAnthropic(texto); if (a) return a; } catch (e) { /* fallback */ }
  try { const o = await _viaOpenAI(texto); if (o) return o; } catch (e) { /* desiste */ }
  return null;
}

module.exports = { extrairFicha, CAMPOS };
