/**
 * SOL — endpoint de chat com a IA.
 *
 * POST body:
 *   {
 *     message: "string",        // mensagem do usuario
 *     history: [...],           // historico previo da conversa (role: user|assistant, content: string)
 *     view: "fila-aprovacao" | "aprovadas"
 *   }
 *
 * Retorna:
 *   200 { resposta: "markdown", acoes_propostas: [...], tokens: N, model: "..." }
 *   401 { error: "..." }
 *   500 { error: "...", detail: "..." }
 *
 * Auth: getUser (Easy Auth ou Teams JWT). Acoes destrutivas (aprovar/rejeitar)
 * NAO sao executadas aqui — sao apenas propostas pro frontend confirmar.
 */

require('isomorphic-fetch');
const { getUser } = require('../shared/auth');
const { runSol } = require('../shared/sol');
const { salvar: salvarHistorico } = require('../shared/solHistorico');

// Detecta se o usuario tem perfil admin/financeiro (ve tudo) pelos grupos Entra ID
// Replica logica de MeusGrupos — verifica se o oid eh membro de grupos privilegiados.
// Pra simplicidade, no MVP: se o email termina em uma whitelist OU o usuario tem
// claim de grupo admin, considera admin. Em prod, melhor consultar Graph.
function detectAdminFromHeaders(req, user) {
  // Header X-Sol-Admin: 'true' setado pelo frontend depois de checar MeusGrupos
  const headerAdmin = req.headers && (req.headers['x-sol-admin'] || req.headers['X-Sol-Admin']);
  if (String(headerAdmin || '').toLowerCase() === 'true') return true;
  // Fallback: emails admin hardcoded (mesma logica do front)
  const adminEmails = (process.env.SOL_ADMIN_EMAILS || '').toLowerCase().split(',').map(s => s.trim()).filter(Boolean);
  if (adminEmails.includes((user.email || '').toLowerCase())) return true;
  return false;
}

module.exports = async function (context, req) {
  const start = Date.now();
  try {
    const user = await getUser(req);
    if (!user) {
      context.res = { status: 401, body: { error: 'Nao autenticado' } };
      return;
    }

    const body = req.body || {};
    const message = String(body.message || '').trim();
    if (!message) {
      context.res = { status: 400, body: { error: 'message obrigatorio' } };
      return;
    }
    if (message.length > 2000) {
      context.res = { status: 400, body: { error: 'message excede 2000 caracteres' } };
      return;
    }

    const history = Array.isArray(body.history) ? body.history.slice(-20) : []; // ultima 20 msgs
    const viewAtual = String(body.view || 'fila-aprovacao');
    const isAdmin = detectAdminFromHeaders(req, user);

    if (!process.env.OPENAI_API_KEY) {
      context.res = { status: 500, body: { error: 'OPENAI_API_KEY nao configurada no Azure SWA' } };
      return;
    }

    // NAO passar 'model' aqui — runSol() decide internamente baseado no provider:
    // Anthropic usa ANTHROPIC_MODEL_HAIKU, OpenAI (fallback) usa OPENAI_MODEL.
    // Passar 'gpt-4o-mini' como override quebra o Anthropic com 404 (model not found).
    const result = await runSol(history, message, user, {
      viewAtual: viewAtual,
      isAdmin: isAdmin,
      maxIter: 8
    });

    const elapsedMs = Date.now() - start;

    // Persiste o turno no historico SP (best-effort, nao bloqueia a resposta)
    // Salva user msg + assistant response em paralelo (Promise.allSettled).
    Promise.allSettled([
      salvarHistorico(user, 'user', message, {}),
      salvarHistorico(user, 'assistant', result.resposta || '', { tokensUsed: result.tokens || 0 })
    ]).catch(function(e){ /* silencia */ });

    context.res = {
      status: 200,
      headers: { 'Content-Type': 'application/json' },
      body: {
        resposta: result.resposta || '',
        acoes_propostas: result.acoes_propostas || [],
        tokens: result.tokens || 0,
        model: result.model,
        provider: result.provider,
        elapsedMs: elapsedMs,
        tool_calls_debug: result.tool_calls_debug || [],
        anthropic_error: result.anthropic_error || null
      }
    };
  } catch (err) {
    context.log && context.log.error && context.log.error('SolChat error:', err);
    context.res = {
      status: 500,
      headers: { 'Content-Type': 'application/json' },
      body: {
        error: 'Erro interno da SOL',
        detail: (err && err.message) || String(err),
        stack: (err && err.stack || '').split('\n').slice(0, 8)
      }
    };
  }
};
