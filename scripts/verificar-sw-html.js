/**
 * Verifica as guardas de integridade de HTML do service worker.
 *
 * POR QUE ESTE ARQUIVO EXISTE
 * Em 24/08/2026 o sistema abriu TOTALMENTE EM BRANCO para o Rafael. O DevTools mostrou o
 * documento terminando depois do </head>: sem <body>, altura zero, console VAZIO — sem body
 * nao ha script para rodar nem erro para reportar. O arquivo no servidor estava intacto; o
 * cache do navegador guardava um HTML truncado e o servia indefinidamente.
 *
 * Nao reimplementa nada: EXTRAI as funcoes do wwwroot/sw.js e as exercita com um Cache API
 * de mentira. Se a regra mudar no SW, o teste acompanha ou acusa.
 *
 * Rodar: node scripts/verificar-sw-html.js
 */

const fs = require('fs');
const src = fs.readFileSync('wwwroot/sw.js', 'utf8');

const armazem = new Map();
function respostaFalsa(txt, ct) {
  return {
    _txt: txt, status: 200, ok: true,
    headers: { get: () => ct || 'text/html; charset=utf-8' },
    clone() { return respostaFalsa(txt, ct); },
    async text() { if (txt === null) throw new Error('stream abortado'); return txt; }
  };
}
global.Response = function (body, init) { return respostaFalsa(body, (init&&init.headers&&init.headers['Content-Type'])); };
global.caches = {
  async open() { return { async put(k, v) { armazem.set(k, v); }, async delete(k) { return armazem.delete(k); } }; },
  async match(k) { return armazem.get(k) || undefined; }
};
global.console = console;

/* Pega so os tres helpers, sem executar o resto do SW (que usa self/addEventListener). */
global.SHELL_CACHE = 'teste-shell';   // definido acima do trecho extraido, no sw.js real
const ini = src.indexOf('const HTML_TERMINA');
const fim = src.indexOf('// Instalacao');
eval(src.slice(ini, fim) + '\nmodule.exports={htmlCompleto,guardarHtmlSeCompleto,lerHtmlDoCacheSeCompleto};');
const H = module.exports;

let falhas = 0;
function ok(cond, nome, extra) {
  console.log((cond ? '  ok   ' : '  FALHA ') + nome + (cond ? '' : '  -> ' + extra));
  if (!cond) falhas++;
}

(async () => {
  const COMPLETO = '<!DOCTYPE html><html lang="pt-BR"><head><title>x</title></head><body><div id="app"></div><script>1</script></body></html>';
  const TRUNCADO = '<!DOCTYPE html><html lang="pt-BR"><head><title>x</title></head>';   // o caso do Rafael
  const SEM_BODY = '<!DOCTYPE html><html><head></head></html>';

  console.log('\n1. htmlCompleto');
  ok(await H.htmlCompleto(respostaFalsa(COMPLETO)) !== null, 'HTML completo e aceito');
  ok(await H.htmlCompleto(respostaFalsa(TRUNCADO)) === null, 'HTML truncado (sem </html>) e RECUSADO');
  ok(await H.htmlCompleto(respostaFalsa(SEM_BODY)) === null, 'HTML sem <body> e RECUSADO');
  ok(await H.htmlCompleto(respostaFalsa(null)) === null, 'stream abortado e RECUSADO');
  ok(await H.htmlCompleto(respostaFalsa(COMPLETO + '\n\n')) !== null, 'espacos no fim nao invalidam');

  console.log('\n2. guardarHtmlSeCompleto');
  armazem.clear();
  ok(await H.guardarHtmlSeCompleto('/index.html', respostaFalsa(COMPLETO)) === true, 'grava o completo');
  ok(armazem.has('/index.html'), 'ficou no cache');
  ok(await H.guardarHtmlSeCompleto('/index.html', respostaFalsa(TRUNCADO)) === false, 'NAO grava o truncado');
  const guardado = await armazem.get('/index.html').text();
  ok(guardado === COMPLETO, 'o completo NAO foi sobrescrito pelo truncado');

  console.log('\n3. lerHtmlDoCacheSeCompleto');
  armazem.clear();
  armazem.set('/index.html', respostaFalsa(TRUNCADO));   // simula cache JA envenenado
  ok(await H.lerHtmlDoCacheSeCompleto('/index.html') === null, 'cache envenenado nao e servido');
  ok(!armazem.has('/index.html'), 'e o envenenado foi APAGADO (nao sobrevive a visitas futuras)');
  armazem.set('/index.html', respostaFalsa(COMPLETO));
  ok(await H.lerHtmlDoCacheSeCompleto('/index.html') !== null, 'cache bom e servido');
  armazem.clear();
  ok(await H.lerHtmlDoCacheSeCompleto('/index.html') === null, 'cache vazio devolve null (cai no offline.html)');

  console.log(falhas ? '\n' + falhas + ' falha(s)\n' : '\nTodos os testes passaram\n');
  process.exit(falhas ? 1 : 0);
})();
