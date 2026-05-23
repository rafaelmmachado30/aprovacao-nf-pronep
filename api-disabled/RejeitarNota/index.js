/**
 * Rejeitar Nota Fiscal
 *
 * POST /api/RejeitarNota
 * Body JSON: { id, motivo, observacao }
 *
 * Status: ESQUELETO - implementar na Sprint 3.
 *
 * TODO:
 *  1. Validar role (gestor da diretoria correspondente)
 *  2. Mover PDF de Notas Pendentes para Notas Rejeitadas/{Unidade}/Diretoria {Diretoria}
 *  3. Atualizar lista: Status=REJEITADA, MotivoRejeicao, RejeitadoPor, RejeitadoEm
 *  4. Notificar quem lançou via EnviarNotificacao
 */

module.exports = async function (context, req) {
  context.log('=== RejeitarNota invoked ===');
  context.res = {
    status: 501,
    body: { error: 'Not implemented yet — Sprint 3', _ver_arquivo: 'api/RejeitarNota/index.js' }
  };
};
