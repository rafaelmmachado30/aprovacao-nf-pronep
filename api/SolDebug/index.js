/**
 * SolDebug — endpoint de diagnostico do SOL.
 *
 * GET /api/SolDebug                  — info basica de env, requires e module load
 * GET /api/SolDebug?test=anthropic   — testa uma chamada simples ao Anthropic
 * GET /api/SolDebug?test=runsol      — chama runSol() inteiro com user fake (mesmo path do SolChat)
 *
 * Anonymous, mas nao expoe valores sensiveis (so booleanos).
 */

module.exports = async function (context, req) {
  const out = {
    timestamp: new Date().toISOString(),
    env: {
      hasAnthropicKey: !!process.env.ANTHROPIC_API_KEY,
      hasOpenAIKey: !!process.env.OPENAI_API_KEY,
      anthropicModelHaiku: process.env.ANTHROPIC_MODEL_HAIKU || '(default)',
      anthropicModelSonnet: process.env.ANTHROPIC_MODEL_SONNET || '(default)',
      openaiModel: process.env.OPENAI_MODEL || '(default)',
      anthropicKeyPrefix: (process.env.ANTHROPIC_API_KEY || '').slice(0, 12) + '...',
      openaiKeyPrefix: (process.env.OPENAI_API_KEY || '').slice(0, 8) + '...'
    },
    requires: {},
    sol: {},
    testCall: null,
    runSolTest: null
  };

  try {
    const mod = require('@anthropic-ai/sdk');
    out.requires.anthropic = {
      ok: true,
      hasDefault: !!mod.default,
      hasAnthropic: !!mod.Anthropic,
      keys: Object.keys(mod).slice(0, 10)
    };
  } catch (e) {
    out.requires.anthropic = { ok: false, error: e.message, stack: (e.stack || '').split('\n').slice(0, 5) };
  }

  try {
    const OpenAI = require('openai');
    out.requires.openai = { ok: true, type: typeof OpenAI };
  } catch (e) {
    out.requires.openai = { ok: false, error: e.message, stack: (e.stack || '').split('\n').slice(0, 5) };
  }

  try {
    const sol = require('../shared/sol');
    const anth = sol.getAnthropic && sol.getAnthropic();
    const oai = sol.getOpenAI && sol.getOpenAI();
    out.sol = {
      loaded: true,
      defaultModel: sol.DEFAULT_MODEL,
      hasGetAnthropic: typeof sol.getAnthropic === 'function',
      hasGetOpenAI: typeof sol.getOpenAI === 'function',
      anthropicClient: !!anth,
      openaiClient: !!oai,
      loadErrors: sol.getLoadErrors ? sol.getLoadErrors() : '(funcao nao disponivel)'
    };
  } catch (e) {
    out.sol = { loaded: false, error: e.message, stack: (e.stack || '').split('\n').slice(0, 8) };
  }

  const testParam = (req.query && req.query.test) || '';

  if (testParam === 'anthropic') {
    try {
      const mod = require('@anthropic-ai/sdk');
      const Anthropic = mod.default || mod.Anthropic || mod;
      const client = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });
      const model = process.env.ANTHROPIC_MODEL_HAIKU || 'claude-haiku-4-5-20251001';
      const resp = await client.messages.create({
        model: model,
        max_tokens: 50,
        messages: [{ role: 'user', content: 'Diga apenas: ok' }]
      });
      out.testCall = { ok: true, model: model, stop_reason: resp.stop_reason, content: resp.content, usage: resp.usage };
    } catch (e) {
      out.testCall = { ok: false, error: e.message, status: e.status, type: e.constructor && e.constructor.name, stack: (e.stack || '').split('\n').slice(0, 8) };
    }
  }

  if (testParam === 'runsol') {
    try {
      const sol = require('../shared/sol');
      const fakeUser = { email: 'rafael.machado@pronep.com.br', name: 'Rafael Machado', oid: '00000000-0000-0000-0000-000000000000' };
      const t0 = Date.now();
      const result = await sol.runSol([], 'Diga apenas: ok', fakeUser, { viewAtual: 'fila-aprovacao', isAdmin: true, maxIter: 2 });
      out.runSolTest = {
        ok: true,
        elapsedMs: Date.now() - t0,
        provider: result.provider,
        model: result.model,
        resposta: result.resposta,
        tokens: result.tokens,
        tool_calls_count: (result.tool_calls_debug || []).length,
        anthropic_error: result.anthropic_error
      };
    } catch (e) {
      out.runSolTest = {
        ok: false,
        error: e.message,
        type: e.constructor && e.constructor.name,
        status: e.status,
        stack: (e.stack || '').split('\n').slice(0, 15)
      };
    }
  }

  context.res = { status: 200, headers: { 'Content-Type': 'application/json' }, body: out };
};
