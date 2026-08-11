/**
 * scripts/verificar-cert-a1.js — prova shared/certA1.js com certificados de teste
 * gerados na hora e jogados fora no fim.
 *
 *   node scripts/verificar-cert-a1.js
 *
 * NAO usa e nao le nenhum certificado da Pronep. Gera dois descartaveis: um no
 * formato atual e um no formato legado (RC2), que e justamente o caso que produz a
 * mensagem incompreensivel do OpenSSL.
 *
 * Sem openssl no PATH este script SAI COM ERRO em vez de "pular". Verificador que
 * passa sem verificar nada foi o que derrubou o login em producao neste projeto.
 */

'use strict';

const { execFileSync } = require('child_process');
const fs = require('fs');
const os = require('os');
const path = require('path');

const certA1 = require('../api/shared/certA1');

let falhas = 0;
function ok(cond, nome, detalhe) {
  if (cond) { console.log('  ok   ' + nome); return; }
  falhas++;
  console.log('  FALHA ' + nome + (detalhe ? '  -> ' + detalhe : ''));
}

function openssl(args) {
  return execFileSync('openssl', args, { stdio: ['ignore', 'pipe', 'pipe'] });
}

const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'certa1-'));
const SENHA = 'senha-de-teste-descartavel';

try {
  openssl(['version']);
} catch (e) {
  console.error('openssl nao encontrado no PATH. Este verificador precisa dele para ' +
                'gerar os certificados de teste — sem isso ele nao verifica nada.');
  process.exit(2);
}

(async function () {
  const key = path.join(dir, 'k.pem');
  const crt = path.join(dir, 'c.pem');
  const atual = path.join(dir, 'atual.pfx');
  const legado = path.join(dir, 'legado.pfx');

  /* 400 dias para o teste de dias restantes nao depender da data de hoje. */
  openssl(['req', '-x509', '-newkey', 'rsa:2048', '-keyout', key, '-out', crt,
           '-days', '400', '-nodes', '-subj', '/CN=CERTIFICADO DE TESTE DESCARTAVEL']);

  openssl(['pkcs12', '-export', '-out', atual, '-inkey', key, '-in', crt,
           '-passout', 'pass:' + SENHA,
           '-keypbe', 'AES-256-CBC', '-certpbe', 'AES-256-CBC', '-macalg', 'sha256']);

  let temLegado = true;
  try {
    openssl(['pkcs12', '-export', '-out', legado, '-inkey', key, '-in', crt,
             '-passout', 'pass:' + SENHA, '-legacy']);
  } catch (e) {
    temLegado = false;
  }

  const bufAtual = fs.readFileSync(atual);

  console.log('\n1. certificado no formato atual, senha certa');
  const r1 = await certA1.inspecionar(bufAtual, SENHA);
  ok(r1.ok === true, 'aceitou', JSON.stringify(r1));
  ok(r1.titular === 'CERTIFICADO DE TESTE DESCARTAVEL', 'leu o titular', r1.titular);
  ok(/^\d{4}-\d{2}-\d{2}$/.test(r1.validadeFim || ''), 'leu a validade', r1.validadeFim);
  ok(r1.diasRestantes > 390 && r1.diasRestantes <= 400,
     'dias restantes proximos de 400', String(r1.diasRestantes));
  ok(!r1.aviso, 'sem aviso de validade ilegivel', r1.aviso);

  console.log('\n2. senha errada');
  const r2 = await certA1.inspecionar(bufAtual, 'senha-errada');
  ok(r2.ok === false, 'recusou');
  ok(r2.causa === 'senha', 'classificou como senha', r2.causa + ': ' + r2.mensagem);
  ok(!/openssl|PKCS12|mac verify/i.test(r2.mensagem),
     'mensagem sem jargao do OpenSSL', r2.mensagem);

  console.log('\n3. senha em branco');
  const r3 = await certA1.inspecionar(bufAtual, '');
  ok(r3.ok === false && r3.causa === 'senha', 'recusou pedindo a senha', JSON.stringify(r3));

  console.log('\n4. arquivo que nao e certificado');
  const pdfFalso = Buffer.concat([Buffer.from('%PDF-1.7\n'), Buffer.alloc(500, 0x41)]);
  const r4 = await certA1.inspecionar(pdfFalso, SENHA);
  ok(r4.ok === false && r4.causa === 'arquivo', 'recusou pelo cabecalho', JSON.stringify(r4));

  /* Os dois casos que a regra antiga ("comeca com 0x30 0x82") errava. Ambos vieram de
     um certificado real recusado em producao, nao de imaginacao. */
  console.log('\n4b. .cer sem chave privada (comeca com 0x30 0x82, e NAO serve)');
  const cer = path.join(dir, 'so-cert.cer');
  openssl(['x509', '-in', crt, '-outform', 'DER', '-out', cer]);
  const bufCer = fs.readFileSync(cer);
  ok(bufCer[0] === 0x30 && bufCer[1] === 0x82, 'o .cer realmente comeca com 30 82');
  ok(certA1.pareceCertificado(bufCer) === false,
     'recusado: nao tem o OID do pkcs-12');
  const r4b = await certA1.inspecionar(bufCer, SENHA);
  ok(r4b.ok === false && r4b.causa === 'arquivo', 'inspecionar tambem recusa',
     JSON.stringify(r4b));

  console.log('\n4c. PKCS#12 em BER, comprimento indefinido (0x30 0x80)');
  /* Reescreve o comprimento do SEQUENCE externo para a forma indefinida, que e o
     que a certificadora emitiu no A1 real. O conteudo continua o mesmo arquivo. */
  const ber = Buffer.from(bufAtual);
  ber[1] = 0x80;
  ok(ber[0] === 0x30 && ber[1] === 0x80, 'fixture em 30 80');
  ok(certA1.pareceCertificado(ber) === true,
     'ACEITO: comprimento indefinido e PKCS#12 valido');

  console.log('\n5. arquivo minusculo');
  const r5 = await certA1.inspecionar(Buffer.from([0x30, 0x82, 0x01]), SENHA);
  ok(r5.ok === false && r5.causa === 'arquivo', 'recusou por tamanho', JSON.stringify(r5));

  console.log('\n6. formato legado (RC2) — o erro que assusta');
  if (!temLegado) {
    falhas++;
    console.log('  FALHA nao consegui gerar o .pfx legado com este openssl; ' +
                'este caso ficou sem prova');
  } else {
    const r6 = await certA1.inspecionar(fs.readFileSync(legado), SENHA);
    ok(r6.ok === false, 'recusou', JSON.stringify(r6));
    ok(r6.causa === 'formato-legado', 'classificou como formato antigo', r6.causa);
    ok(/reexporte/i.test(r6.mensagem), 'mandou reexportar', r6.mensagem);
    ok(/keypbe AES-256-CBC/.test(r6.mensagem), 'deu o comando pronto');
  }

  console.log('\n7. classificacao isolada');
  ok(certA1.classificar('Unsupported PKCS12 PFX data').causa === 'formato-legado', 'PKCS12 -> legado');
  ok(certA1.classificar('mac verify failure').causa === 'senha', 'mac verify -> senha');
  ok(certA1.classificar('coisa nova').causa === 'outro', 'desconhecido -> outro');
  ok(certA1.pareceCertificado(bufAtual) === true, 'pareceCertificado no arquivo bom');
  ok(certA1.pareceCertificado(Buffer.from('nao')) === false, 'pareceCertificado no lixo');

  fs.rmSync(dir, { recursive: true, force: true });
  ok(!fs.existsSync(dir), 'apagou os certificados de teste');

  console.log(falhas ? '\n' + falhas + ' falha(s)\n' : '\ncertA1 verificado\n');
  process.exit(falhas ? 1 : 0);
})().catch(function (e) {
  try { fs.rmSync(dir, { recursive: true, force: true }); } catch (x) { /* ok */ }
  console.error('erro inesperado:', e && e.message);
  process.exit(1);
});
