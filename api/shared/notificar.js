/**
 * Helper compartilhado: dispara notificacao via chamada HTTP interna ao /api/EnviarNotificacao
 * Nao bloqueia o fluxo principal (try/catch absorve erros)
 */
require('isomorphic-fetch');

async function notificar(evento, destinatarios, dados, cc) {
  try {
    // Detecta URL base. Em SWA, WEBSITE_HOSTNAME ou WEBSITE_SITE_NAME ajudam
    const host = process.env.WEBSITE_HOSTNAME || 'localhost';
    const base = host.startsWith('http') ? host : `https://${host}`;
    const url = `${base}/api/EnviarNotificacao`;

    // Para chamadas internas de Function pra Function no SWA, podemos pular o auth
    // pois o host eh interno. Mas a function tem authLevel anonymous, entao OK.
    const resp = await fetch(url, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        evento,
        destinatarios: Array.isArray(destinatarios) ? destinatarios : [destinatarios],
        cc: cc || [],
        dados: dados || {}
      })
    });
    return resp.ok;
  } catch (e) {
    return false;
  }
}

module.exports = { notificar };
