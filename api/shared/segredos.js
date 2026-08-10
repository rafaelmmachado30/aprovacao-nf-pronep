/**
 * shared/segredos.js — cifra e decifra segredos curtos (AES-256-GCM).
 *
 * PARA QUE, SE O CONTAINER JA E PRIVADO
 * A senha do certificado passa a morar num blob porque a Function nao consegue
 * escrever numa App Setting — e sem isso a renovacao anual continuaria dependendo do
 * portal do Azure, que era exatamente o que se queria tirar do caminho.
 *
 * Sendo honesto sobre o ganho: a protecao aqui e modesta e vale dizer qual e. A
 * chave da conta de Storage esta numa App Setting, entao quem tem o portal alcanca o
 * blob. Cifrar com uma chave SEPARADA (CONFIG_CRYPTO_KEY) faz com que ler o blob nao
 * seja suficiente, e evita a senha em texto puro em repouso, em backup e em qualquer
 * listagem. Nao e cofre; e uma camada a mais no mesmo perimetro.
 *
 * GCM e nao CBC porque GCM autentica: se o blob for alterado, decifrar falha em vez
 * de devolver bytes silenciosamente errados.
 *
 * A chave nunca entra no repositorio nem em log. Para gerar uma:
 *   node -e "console.log(require('crypto').randomBytes(32).toString('base64'))"
 * e colar direto no portal do Azure, em Application settings.
 */

const crypto = require('crypto');

const PREFIXO = 'v1';

function lerChave() {
  const b64 = process.env.CONFIG_CRYPTO_KEY;
  if (!b64) {
    throw new Error('App Setting CONFIG_CRYPTO_KEY nao configurada. Gere com: ' +
      'node -e "console.log(require(\'crypto\').randomBytes(32).toString(\'base64\'))"');
  }
  let chave;
  try {
    chave = Buffer.from(String(b64).trim(), 'base64');
  } catch (e) {
    throw new Error('CONFIG_CRYPTO_KEY nao esta em base64 valido');
  }
  if (chave.length !== 32) {
    throw new Error('CONFIG_CRYPTO_KEY precisa ter 32 bytes em base64 (tem ' +
                    chave.length + ')');
  }
  return chave;
}

/** True se da para cifrar — a tela usa isto para avisar antes de o usuario digitar. */
function temChave() {
  try { lerChave(); return true; } catch (e) { return false; }
}

/** @returns {string} "v1.<iv>.<tag>.<cifrado>", tudo em base64 */
function cifrar(texto) {
  const chave = lerChave();
  const iv = crypto.randomBytes(12);
  const c = crypto.createCipheriv('aes-256-gcm', chave, iv);
  const ct = Buffer.concat([c.update(String(texto), 'utf8'), c.final()]);
  return [PREFIXO, iv.toString('base64'), c.getAuthTag().toString('base64'),
          ct.toString('base64')].join('.');
}

/** @returns {string} o texto original; lanca se o pacote foi mexido ou a chave mudou */
function decifrar(pacote) {
  const partes = String(pacote || '').split('.');
  if (partes.length !== 4 || partes[0] !== PREFIXO) {
    throw new Error('Segredo em formato desconhecido');
  }
  const chave = lerChave();
  const iv = Buffer.from(partes[1], 'base64');
  const tag = Buffer.from(partes[2], 'base64');
  const ct = Buffer.from(partes[3], 'base64');
  const d = crypto.createDecipheriv('aes-256-gcm', chave, iv);
  d.setAuthTag(tag);
  try {
    return Buffer.concat([d.update(ct), d.final()]).toString('utf8');
  } catch (e) {
    /* Mensagem propria: a do OpenSSL ("unable to authenticate data") nao ajuda quem
       le o log, e o motivo real quase sempre e a chave ter sido trocada. */
    throw new Error('Nao consegui decifrar o segredo. A CONFIG_CRYPTO_KEY foi trocada ' +
                    'depois que ele foi gravado? Nesse caso, grave a senha de novo.');
  }
}

module.exports = { cifrar, decifrar, temChave };
