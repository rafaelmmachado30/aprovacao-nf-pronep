/**
 * /api/ContratosPing — endpoint MINIMO de teste.
 * Sem requires no top-level (alem do basico). Testa lazy load passo-a-passo.
 *
 * Query: ?step=ping|pdfparse|anthropic|contratos|all
 */
module.exports = async function (context, req) {
  const step = (req.query && req.query.step) || 'all';
  const log = [];
  function add(nome, fn) {
    const t0 = Date.now();
    try {
      const r = fn();
      log.push({ nome, ok: true, ms: Date.now() - t0, keys: r && typeof r === 'object' ? Object.keys(r).slice(0, 6) : null });
    } catch (e) {
      log.push({ nome, ok: false, ms: Date.now() - t0, error: e.message, stack: (e.stack || '').split('\n').slice(0, 4) });
    }
  }
  log.push({ nome: 'ping', ok: true, time: new Date().toISOString() });
  if (step === 'pdfparse' || step === 'all') add('require_pdfparse', function(){ return require('pdf-parse'); });
  if (step === 'anthropic' || step === 'all') add('require_anthropic', function(){ return require('@anthropic-ai/sdk'); });
  if (step === 'mammoth' || step === 'all') add('require_mammoth', function(){ return require('mammoth'); });
  if (step === 'auth' || step === 'all') add('require_auth', function(){ return require('../shared/auth'); });
  if (step === 'contratos' || step === 'all') add('require_contratos', function(){ return require('../shared/contratos'); });

  context.res = {
    status: 200,
    headers: { 'Content-Type': 'application/json' },
    body: { ok: true, step, log }
  };
};
