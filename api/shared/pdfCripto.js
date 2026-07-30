// Remove a protecao de PDFs cifrados pra que o pdf-lib consiga carimba-los.
//
// PROBLEMA: boletos de banco (Itau, entre outros) vem com criptografia de PERMISSOES —
// abrem sem senha, mas tem restricao de edicao. O pdf-lib carrega esses arquivos e marca
// isEncrypted=true, porem NAO decifra os streams: salvar com o watermark produziria um
// arquivo corrompido. Por isso o AprovarNota/RejeitarNota arquivavam o original sem
// carimbo — e o Financeiro estornava a NF por "falta o carimbo de aprovado".
//
// SOLUCAO: o mupdf (WASM, ja usado no CaixaEntrada pra unir PDFs) decifra de verdade.
// Reescrevemos o arquivo sem criptografia e o pdf-lib carimba pelo caminho normal, com
// o mesmo visual de sempre.
//
// So funciona quando o PDF abre sem senha de USUARIO. PDF com senha de abertura nao tem
// como ser carimbado — nesse caso devolve null e quem chama avisa o aprovador.

/**
 * @param {Buffer} pdfBuffer PDF possivelmente cifrado
 * @returns {Promise<Buffer|null>} PDF sem criptografia, ou null se nao foi possivel
 */
async function removerProtecao(pdfBuffer) {
  try {
    // import() dinamico: o mupdf e ESM-only e carrega WASM (pesado). Fica fora do caminho
    // comum — so entra em PDF cifrado, que e minoria.
    const mupdf = await import('mupdf');
    const doc = mupdf.Document.openDocument(new Uint8Array(pdfBuffer), 'application/pdf');
    // Senha de ABERTURA: nem o mupdf passa sem ela. Nao ha carimbo possivel.
    if (doc.needsPassword && doc.needsPassword()) return null;
    const saida = doc.saveToBuffer('encrypt=none').asUint8Array();
    if (!saida || !saida.length) return null;
    return Buffer.from(saida);
  } catch (e) {
    return null;
  }
}

module.exports = { removerProtecao };
