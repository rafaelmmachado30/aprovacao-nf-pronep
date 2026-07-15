/**
 * /api/Aquecer  (GET)
 *
 * Endpoint minimo de WARM-UP contra o cold start do Static Web Apps.
 * Pingado por um cron do GitHub Actions em horario comercial (aquecer.yml).
 * Manter o container "quente" evita os ~40s de primeira carga de manha.
 *
 * DELIBERADAMENTE MINIMO (o warm-up ja quebrou producao antes - "tela branca"):
 *   - NAO carrega deps pesadas (nada de pdf/mupdf/IA/Teams).
 *   - NAO tem efeito colateral (nao envia nada, nao grava nada).
 *   - So mantem a instancia viva e, de brinde, aquece o cache de siteId
 *     do shared/graph (best-effort; se o Graph falhar, ainda retorna 200).
 *
 * Auth: mesmo padrao dos crons existentes (header X-Alerta-Secret == App Setting
 * ALERTA_DIARIO_SECRET). Rota liberada como anonymous no staticwebapp.config.json.
 */

const { getGraphClient, resolveSiteId } = require('../shared/graph');

module.exports = async function (context, req) {
  const t0 = Date.now();

  // Auth por shared secret (endpoint eh anonymous na rota)
  const expectedSecret = process.env.ALERTA_DIARIO_SECRET;
  if (!expectedSecret) {
    context.res = { status: 500, headers: { 'Content-Type': 'application/json' }, body: { error: 'ALERTA_DIARIO_SECRET nao configurado' } };
    return;
  }
  const headerSecret = (req.headers && (req.headers['x-alerta-secret'] || req.headers['X-Alerta-Secret'])) || '';
  if (headerSecret !== expectedSecret) {
    context.res = { status: 401, headers: { 'Content-Type': 'application/json' }, body: { error: 'Unauthorized' } };
    return;
  }

  // Aquece o cache de siteId do shared/graph. Best-effort: nunca derruba o warm-up.
  let siteWarm = false;
  try {
    const client = getGraphClient();
    await resolveSiteId(client);
    siteWarm = true;
  } catch (e) {
    context.log && context.log.warn && context.log.warn('Aquecer: warm do siteId falhou (ok):', (e && e.message) || e);
  }

  context.res = {
    status: 200,
    headers: { 'Content-Type': 'application/json', 'Cache-Control': 'no-store' },
    body: { ok: true, siteWarm: siteWarm, ms: Date.now() - t0 }
  };
};
