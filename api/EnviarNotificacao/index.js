/**
 * Enviar Notificação (Email + Teams)
 *
 * POST /api/EnviarNotificacao
 * Body JSON: { evento: 'lancada'|'aprovada'|'rejeitada', notaId, destinatarios }
 *
 * Status: ESQUELETO - implementar na Sprint 4.
 *
 * TODO:
 *  1. Renderizar template baseado no evento
 *  2. Enviar e-mail via Microsoft Graph (Mail.Send) na caixa nf-aprovacoes@pronep.com.br
 *  3. Enviar Adaptive Card para o canal Teams via Incoming Webhook (TEAMS_WEBHOOK_URL)
 *     - Card para "lancada" tem botões Action.Http "Aprovar" / "Rejeitar" → chamam back direto
 */

module.exports = async function (context, req) {
  context.log('=== EnviarNotificacao invoked ===');
  context.res = {
    status: 501,
    body: { error: 'Not implemented yet — Sprint 4', _ver_arquivo: 'api/EnviarNotificacao/index.js' }
  };
};
