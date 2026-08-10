/**
 * scripts/verificar-segredos.js — prova shared/segredos.js.
 *   node scripts/verificar-segredos.js
 *
 * Usa chaves geradas na hora. Nenhum segredo real entra aqui.
 */

'use strict';

const crypto = require('crypto');
const s = require('../api/shared/segredos');

let falhas = 0;
function ok(cond, nome, detalhe) {
  if (cond) { console.log('  ok   ' + nome); return; }
  falhas++;
  console.log('  FALHA ' + nome + (detalhe ? '  -> ' + detalhe : ''));
}

console.log('\n1. sem chave configurada');
delete process.env.CONFIG_CRYPTO_KEY;
ok(s.temChave() === false, 'temChave() false');
try {
  s.cifrar('x');
  ok(false, 'deveria ter lancado');
} catch (e) {
  ok(/CONFIG_CRYPTO_KEY nao configurada/.test(e.message),
     'o erro ensina a gerar a chave', e.message.slice(0, 70));
}

console.log('\n2. chave de tamanho errado');
process.env.CONFIG_CRYPTO_KEY = crypto.randomBytes(16).toString('base64');
try {
  s.cifrar('x');
  ok(false, 'deveria ter lancado');
} catch (e) {
  ok(/32 bytes/.test(e.message), 'exige 32 bytes', e.message);
}

console.log('\n3. ida e volta');
process.env.CONFIG_CRYPTO_KEY = crypto.randomBytes(32).toString('base64');
ok(s.temChave() === true, 'temChave() true');
const senha = 'SenhaComAcento-çãõ!@#$%1234';
const pacote = s.cifrar(senha);
ok(pacote.indexOf('v1.') === 0, 'formato versionado');
ok(pacote.split('.').length === 4, 'quatro partes');
ok(pacote.indexOf(senha) < 0, 'a senha nao aparece no pacote');
ok(s.decifrar(pacote) === senha, 'decifrou identico, inclusive acento');

console.log('\n4. IV aleatorio');
ok(s.cifrar(senha) !== s.cifrar(senha),
   'cifrar a mesma senha duas vezes da pacotes diferentes');

console.log('\n5. pacote adulterado');
const partes = pacote.split('.');
const ct = Buffer.from(partes[3], 'base64');
ct[0] ^= 0xFF;
partes[3] = ct.toString('base64');
try {
  s.decifrar(partes.join('.'));
  ok(false, 'GCM deveria ter recusado');
} catch (e) {
  ok(/Nao consegui decifrar/.test(e.message), 'GCM detectou a adulteracao');
}

console.log('\n6. chave trocada depois de gravar');
process.env.CONFIG_CRYPTO_KEY = crypto.randomBytes(32).toString('base64');
try {
  s.decifrar(pacote);
  ok(false, 'deveria ter falhado');
} catch (e) {
  ok(/trocada/.test(e.message), 'a mensagem aponta a troca de chave');
}

console.log('\n7. formato desconhecido');
try {
  s.decifrar('lixo');
  ok(false, 'deveria ter falhado');
} catch (e) {
  ok(/formato desconhecido/.test(e.message), 'recusou');
}

console.log(falhas ? '\n' + falhas + ' falha(s)\n' : '\nsegredos verificado\n');
process.exit(falhas ? 1 : 0);
