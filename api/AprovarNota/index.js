/**
 * Aprovar Nota Fiscal
 *
 * POST /api/AprovarNota
 * Body JSON: { id: "NF-2026-XXXX" }
 *
 * Status: ESQUELETO - implementar na Sprint 3.
 *
 * TODO em sequência:
 *  1. Validar role do usuário (precisa ser gestor da diretoria correspondente)
 *  2. Buscar a NF na lista 'NotasFiscais' (Graph API)
 *  3. Verificar se ainda está PENDENTE
 *  4. Baixar o PDF de Notas Pendentes
 *  5. Aplicar watermark com pdf-lib (3 linhas, azul #0000FF, alpha 0.3):
 *       "APROVADO"
 *       "dd/MM/yyyy HH:mm:ss"  (timezone America/Sao_Paulo)
 *       "Por: {prefixo-do-email}"
 *     Estilo idêntico ao PDF Blocks atual.
 *  6. Garantir pasta /Notas Aprovadas/{Unidade}/{AAAA-MM-DD} (criar se não existir)
 *  7. Subir o PDF watermark em Notas Aprovadas
 *  8. Deletar o PDF original de Notas Pendentes
 *  9. Atualizar o item: Status=APROVADA, AprovadoPor, AprovadoEm
 * 10. Chamar EnviarNotificacao para informar quem lançou
 */

module.exports = async function (context, req) {
  context.log('=== AprovarNota invoked ===');
  context.res = {
    status: 501,
    body: { error: 'Not implemented yet — Sprint 3', _ver_arquivo: 'api/AprovarNota/index.js' }
  };
};
