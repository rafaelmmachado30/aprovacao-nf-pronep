/**
 * Lançar Nova Nota Fiscal
 *
 * POST /api/PostNota
 * Body multipart/form-data:
 *   - fornecedorCNPJ (string)
 *   - numero, serie, valor, vencimento (campos da NF)
 *   - negociadoCom, negociadoComEmail (opcional, quando vencimento < D+5)
 *   - file (binary, PDF)
 *
 * Status: ESQUELETO - implementar na Sprint 2.
 *
 * TODO em sequência:
 *  1. Parse multipart e validar campos (negociar usar 'parse-multipart-data')
 *  2. Validar tamanho (<= 6 MB = 6291456 bytes)
 *  3. Calcular SHA-256 do PDF
 *  4. Extrair chave NFS-e do PDF (regex de 44 dígitos no texto)
 *  5. Anti-duplicidade: consultar SharePoint List 'NotasFiscais' por
 *      ChaveAcesso || HashArquivo || (CNPJ+Numero+Serie). Se bater, 409.
 *  6. Validar vencimento >= D+5. Fora disso, exigir negociadoCom preenchido.
 *  7. Resolver unidade/diretoria do fornecedor (lookup na lista Fornecedores)
 *  8. Determinar aprovador via lista Diretorias (matriz Unidade x Diretoria)
 *  9. Salvar PDF em SharePoint/Notas Fiscais/Notas Pendentes/{Unidade}/Diretoria {Diretoria}
 * 10. Criar item na lista 'NotasFiscais' em status PENDENTE
 * 11. Chamar EnviarNotificacao para notificar o gestor por e-mail + Teams
 * 12. Retornar {id, caminhoSharePoint, aprovador}
 */

module.exports = async function (context, req) {
  context.log('=== PostNota invoked ===');
  context.res = {
    status: 501,
    body: { error: 'Not implemented yet — Sprint 2', _ver_arquivo: 'api/PostNota/index.js' }
  };
};
