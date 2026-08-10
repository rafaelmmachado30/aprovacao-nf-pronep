/**
 * scripts/verificar-config-cert.js — prova /api/ConfigCertificadoSefaz.
 *   node scripts/verificar-config-cert.js
 *
 * Dubla APENAS o que sai da maquina: Blob Storage, Graph (auth/auditoria) e o
 * isomorphic-fetch. shared/segredos e shared/certA1 rodam de verdade — se a senha
 * aparecesse em texto puro no que vai para o Storage, o teste 5 acusaria.
 *
 * Os certificados sao gerados na hora com openssl e apagados no fim. Nenhum
 * certificado da Pronep e lido.
 */

'use strict';

const { execFileSync } = require('child_process');
const fs = require('fs');
const os = require('os');
const path = require('path');
const crypto = require('crypto');

let falhas = 0;
function ok(cond, nome, detalhe) {
  if (cond) { console.log('  ok   ' + nome); return; }
  falhas++;
  console.log('  FALHA ' + nome + (detalhe ? '  -> ' + detalhe : ''));
}

const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'cfgcert-'));
const SENHA = 'senha-descartavel-123';
const CNPJ = '00092929000198';

try { execFileSync('openssl', ['version'], { stdio: 'ignore' }); }
catch (e) {
  console.error('openssl nao encontrado; sem ele este verificador nao verifica nada.');
  process.exit(2);
}

/* ---------------------------------------------------------------- dubles */
const blobFalso = { arquivos: {}, falharEm: null };
const auditoria = [];
let usuarioEhAdmin = true;

function instalarDuble(caminhoRelativo, modulo) {
  const p = require.resolve(caminhoRelativo);
  require.cache[p] = { id: p, filename: p, loaded: true, exports: modulo };
}

/* isomorphic-fetch pode nao estar instalado fora do deploy; nada aqui usa rede. */
try { require.resolve('isomorphic-fetch'); }
catch (e) {
  const p = path.join(dir, 'isomorphic-fetch.js');
  fs.writeFileSync(p, 'module.exports = {};');
  require.cache[require.resolve(p)] = { id: p, filename: p, loaded: true, exports: {} };
  const Module = require('module');
  const originalResolve = Module._resolveFilename;
  Module._resolveFilename = function (pedido) {
    if (pedido === 'isomorphic-fetch') return p;
    return originalResolve.apply(this, arguments);
  };
}

instalarDuble('../api/shared/auth', {
  getUser: function () { return { oid: 'oid-teste', email: 'rafael.machado@pronep.com.br', name: 'Rafael' }; }
});

instalarDuble('../api/shared/authz', {
  requireAdmin: async function (context, req) {
    if (!usuarioEhAdmin) {
      context.res = { status: 403, headers: {}, body: { error: 'Acesso restrito a administradores' } };
      return null;
    }
    return { isAdmin: true };
  }
});

instalarDuble('../api/shared/auditLog', {
  registrar: async function (user, acao, objeto, resultado, detalhes) {
    auditoria.push({ user: user, acao: acao, objeto: objeto, resultado: resultado, detalhes: detalhes });
    return { ok: true };
  }
});

instalarDuble('../api/shared/documentosFiscais', {
  lerCnpjsConfigurados: function () {
    return [{ cnpj: CNPJ, apelido: 'Matriz RJ', unidade: 'RJ', diretoria: null }];
  }
});

instalarDuble('../api/shared/blobCert', {
  listar: async function () { return Object.keys(blobFalso.arquivos); },
  lerTexto: async function (nome) {
    const v = blobFalso.arquivos[nome];
    return v == null ? null : (Buffer.isBuffer(v) ? v.toString('utf8') : v);
  },
  lerPfx: async function (cnpj) {
    const v = blobFalso.arquivos[cnpj + '.pfx'];
    if (!v) throw new Error('Certificado ' + cnpj + '.pfx nao existe no container');
    return v;
  },
  gravarPfx: async function (cnpj, buf) {
    if (blobFalso.falharEm === 'pfx') throw new Error('Storage caiu no arquivo');
    blobFalso.arquivos[cnpj + '.pfx'] = buf;
    return true;
  },
  gravarTexto: async function (nome, txt) {
    if (blobFalso.falharEm && nome.indexOf(blobFalso.falharEm) >= 0) {
      throw new Error('Storage caiu em ' + nome);
    }
    blobFalso.arquivos[nome] = txt;
    return true;
  }
});

const endpoint = require('../api/ConfigCertificadoSefaz/index.js');

function chamar(metodo, body) {
  const context = { log: function () {}, res: null };
  return endpoint(context, { method: metodo, body: body || null, headers: {} })
    .then(function () { return context.res; });
}

(async function () {
  const key = path.join(dir, 'k.pem');
  const crt = path.join(dir, 'c.pem');
  const atual = path.join(dir, 'atual.pfx');
  const legado = path.join(dir, 'legado.pfx');

  execFileSync('openssl', ['req', '-x509', '-newkey', 'rsa:2048', '-keyout', key, '-out', crt,
    '-days', '400', '-nodes', '-subj', '/CN=PRONEP TESTE DESCARTAVEL'], { stdio: 'ignore' });
  execFileSync('openssl', ['pkcs12', '-export', '-out', atual, '-inkey', key, '-in', crt,
    '-passout', 'pass:' + SENHA, '-keypbe', 'AES-256-CBC', '-certpbe', 'AES-256-CBC',
    '-macalg', 'sha256'], { stdio: 'ignore' });
  let temLegado = true;
  try {
    execFileSync('openssl', ['pkcs12', '-export', '-out', legado, '-inkey', key, '-in', crt,
      '-passout', 'pass:' + SENHA, '-legacy'], { stdio: 'ignore' });
  } catch (e) { temLegado = false; }

  const pfxB64 = fs.readFileSync(atual).toString('base64');

  process.env.SEFAZ_CERT_STORAGE = 'DefaultEndpointsProtocol=https;AccountName=x;AccountKey=' +
    Buffer.from('chave-falsa').toString('base64') + ';';
  process.env.CONFIG_CRYPTO_KEY = crypto.randomBytes(32).toString('base64');
  delete process.env['SEFAZ_CERT_' + CNPJ + '_SENHA'];

  console.log('\n1. quem nao e admin nao passa');
  usuarioEhAdmin = false;
  const r1 = await chamar('POST', { cnpj: CNPJ, pfxBase64: pfxB64, senha: SENHA });
  ok(r1.status === 403, 'devolveu 403', String(r1.status));
  ok(Object.keys(blobFalso.arquivos).length === 0, 'nao gravou nada');
  usuarioEhAdmin = true;

  console.log('\n2. GET com nada configurado ainda');
  const r2 = await chamar('GET');
  ok(r2.status === 200, '200');
  ok(r2.body.cnpjs.length === 1, 'listou o CNPJ de SEFAZ_CNPJS');
  ok(r2.body.cnpjs[0].temArquivo === false, 'diz que nao tem arquivo');
  ok(r2.body.cnpjs[0].validadeFim === null, 'sem validade');
  ok(r2.body.prontoParaGravar === true, 'pronto para gravar (tem CONFIG_CRYPTO_KEY)');

  console.log('\n3. CNPJ fora de SEFAZ_CNPJS');
  const r3 = await chamar('POST', { cnpj: '11111111000111', pfxBase64: pfxB64, senha: SENHA });
  ok(r3.status === 400, '400', String(r3.status));
  ok(/SEFAZ_CNPJS/.test(r3.body.error), 'explica onde cadastrar');

  console.log('\n4. senha errada nao grava nada');
  const r4 = await chamar('POST', { cnpj: CNPJ, pfxBase64: pfxB64, senha: 'errada' });
  ok(r4.status === 422, '422', String(r4.status));
  ok(r4.body.causa === 'senha', 'causa senha', r4.body.causa);
  ok(Object.keys(blobFalso.arquivos).length === 0, 'container continua vazio',
     Object.keys(blobFalso.arquivos).join(','));

  console.log('\n5. envio valido');
  const r5 = await chamar('POST', { cnpj: CNPJ, pfxBase64: pfxB64, senha: SENHA });
  ok(r5.status === 200, '200', JSON.stringify(r5.body));
  ok(r5.body.titular === 'PRONEP TESTE DESCARTAVEL', 'devolveu o titular', r5.body.titular);
  ok(/^\d{4}-\d{2}-\d{2}$/.test(r5.body.validadeFim || ''), 'devolveu a validade', r5.body.validadeFim);
  ok(r5.body.diasRestantes > 390, 'dias restantes', String(r5.body.diasRestantes));

  const nomes = Object.keys(blobFalso.arquivos).sort();
  ok(nomes.join(',') === CNPJ + '.meta.json,' + CNPJ + '.pfx,' + CNPJ + '.senha',
     'gravou arquivo, senha e meta', nomes.join(','));

  const pacoteSenha = blobFalso.arquivos[CNPJ + '.senha'];
  ok(pacoteSenha.indexOf(SENHA) < 0, 'A SENHA NAO ESTA EM TEXTO PURO no Storage');
  ok(pacoteSenha.indexOf('v1.') === 0, 'senha cifrada no formato versionado');
  const segredos = require('../api/shared/segredos');
  ok(segredos.decifrar(pacoteSenha) === SENHA, 'e decifravel de volta');

  const respostaCrua = JSON.stringify(r5.body);
  ok(respostaCrua.indexOf(SENHA) < 0, 'a resposta HTTP nao contem a senha');

  const meta = JSON.parse(blobFalso.arquivos[CNPJ + '.meta.json']);
  ok(meta.atualizadoPor === 'rafael.machado@pronep.com.br', 'meta registra quem trocou');
  ok(!!meta.atualizadoEm, 'meta registra quando');
  ok(JSON.stringify(meta).indexOf(SENHA) < 0, 'a senha nao esta no meta.json');

  ok(auditoria.length === 1, 'auditou uma vez', String(auditoria.length));
  ok(auditoria[0].acao === 'sefaz.certificado.atualizado', 'acao correta', auditoria[0].acao);
  ok(auditoria[0].detalhes && auditoria[0].detalhes.cnpj === CNPJ,
     'detalhes no 5o argumento (nao no 4o)', JSON.stringify(auditoria[0].detalhes));
  ok(JSON.stringify(auditoria[0]).indexOf(SENHA) < 0, 'a senha nao entrou na auditoria');

  console.log('\n6. GET depois do envio');
  const r6 = await chamar('GET');
  const c6 = r6.body.cnpjs[0];
  ok(c6.temArquivo === true, 've o arquivo');
  ok(c6.temSenhaNoBlob === true, 've a senha no blob');
  ok(c6.titular === 'PRONEP TESTE DESCARTAVEL', 've o titular');
  ok(c6.diasRestantes > 390, 'calcula os dias restantes', String(c6.diasRestantes));
  ok(JSON.stringify(r6.body).indexOf(SENHA) < 0, 'o GET nao devolve a senha');

  console.log('\n7. trocar SO a senha, sem enviar arquivo');
  const r7ruim = await chamar('POST', { cnpj: CNPJ, senha: 'outra-errada' });
  ok(r7ruim.status === 422, 'senha que nao abre o arquivo guardado e recusada',
     String(r7ruim.status));
  const r7 = await chamar('POST', { cnpj: CNPJ, senha: SENHA });
  ok(r7.status === 200 && r7.body.arquivoTrocado === false,
     'aceita a senha correta sem trocar o arquivo', JSON.stringify(r7.body));
  ok(/arquivo continua o mesmo/.test(r7.body.mensagem), 'mensagem diz o que mudou');

  console.log('\n8. senha sem arquivo nenhum gravado');
  const guardado = blobFalso.arquivos;
  blobFalso.arquivos = {};
  const r8 = await chamar('POST', { cnpj: CNPJ, senha: SENHA });
  ok(r8.status === 404, '404 pedindo o arquivo', String(r8.status));
  ok(/Envie o arquivo/.test(r8.body.error), 'mensagem clara');
  blobFalso.arquivos = guardado;

  console.log('\n9. arquivo grande demais');
  const grande = Buffer.alloc(70 * 1024, 0x41);
  grande[0] = 0x30; grande[1] = 0x82;
  const r9 = await chamar('POST', { cnpj: CNPJ, pfxBase64: grande.toString('base64'), senha: SENHA });
  ok(r9.status === 413, '413', String(r9.status));
  ok(/9 KB/.test(r9.body.error), 'diz o tamanho esperado');

  console.log('\n10. formato legado');
  if (!temLegado) {
    falhas++;
    console.log('  FALHA nao gerei o .pfx legado; caso sem prova');
  } else {
    const r10 = await chamar('POST', {
      cnpj: CNPJ, pfxBase64: fs.readFileSync(legado).toString('base64'), senha: SENHA });
    ok(r10.status === 422 && r10.body.causa === 'formato-legado', 'recusou com causa clara',
       JSON.stringify(r10.body));
    ok(/keypbe AES-256-CBC/.test(r10.body.error), 'deu o comando de reexportacao');
  }

  console.log('\n11. sem CONFIG_CRYPTO_KEY');
  const chaveGuardada = process.env.CONFIG_CRYPTO_KEY;
  delete process.env.CONFIG_CRYPTO_KEY;
  const r11 = await chamar('POST', { cnpj: CNPJ, pfxBase64: pfxB64, senha: SENHA });
  ok(r11.status === 503, '503', String(r11.status));
  ok(/CONFIG_CRYPTO_KEY/.test(r11.body.error), 'diz qual App Setting falta');
  const r11get = await chamar('GET');
  ok(r11get.body.prontoParaGravar === false, 'o GET avisa que nao esta pronto');
  ok(r11get.body.avisos.some(function (a) { return /CONFIG_CRYPTO_KEY/.test(a); }),
     'aviso na lista');
  process.env.CONFIG_CRYPTO_KEY = chaveGuardada;

  console.log('\n12. Storage falha na segunda gravacao');
  blobFalso.arquivos = {};
  blobFalso.falharEm = '.senha';
  const r12 = await chamar('POST', { cnpj: CNPJ, pfxBase64: pfxB64, senha: SENHA });
  ok(r12.status === 502, '502', String(r12.status));
  ok(/Ja tinha gravado: arquivo/.test(r12.body.error),
     'diz exatamente o que ficou pela metade', r12.body.error);
  ok(/de novo/.test(r12.body.error), 'diz que repetir resolve');
  ok(r12.body.gravado.join(',') === 'arquivo', 'lista o que foi gravado');
  blobFalso.falharEm = null;

  console.log('\n13. metodo nao suportado');
  const r13 = await chamar('DELETE', {});
  ok(r13.status === 405, '405', String(r13.status));

  fs.rmSync(dir, { recursive: true, force: true });
  ok(!fs.existsSync(dir), 'apagou os certificados de teste');

  console.log(falhas ? '\n' + falhas + ' falha(s)\n' : '\nConfigCertificadoSefaz verificado\n');
  process.exit(falhas ? 1 : 0);
})().catch(function (e) {
  try { fs.rmSync(dir, { recursive: true, force: true }); } catch (x) { /* ok */ }
  console.error('erro inesperado:', e && e.stack);
  process.exit(1);
});
