#!/usr/bin/env node
/**
 * Verifica a sintaxe do JavaScript de wwwroot/index.html.
 *
 * POR QUE EXISTE, E POR QUE ASSIM: o verificador anterior filtrava blocos com
 *   /type=["'](?!text\/javascript)/.test(bloco)
 * testando o BLOCO INTEIRO — corpo incluso. O corpo tem `type="email"`, entao o
 * unico bloco que importa (545 KB de codigo) era descartado em silencio, e a
 * mensagem "5 blocos compilam" contava apenas <script src> vazios. O resultado
 * foi um `/*` sem fechamento chegar em producao e derrubar o login.
 *
 * Verificacao que nao verifica e pior do que verificacao nenhuma: ela produz
 * confianca. Aqui o teste do atributo olha SO os atributos, e o script FALHA se
 * nao encontrar nenhum bloco com codigo — silencio nao passa por sucesso.
 */
const fs = require('fs');
const vm = require('vm');
const path = process.argv[2] || 'wwwroot/index.html';
const html = fs.readFileSync(path, 'utf8');

const re = /<script([^>]*)>([\s\S]*?)<\/script>/g;
let m, comCodigo = 0, verificados = 0;
const erros = [];

while ((m = re.exec(html))) {
  const attrs = m[1] || '';
  const corpo = m[2] || '';
  const linhaTag = html.slice(0, m.index).split('\n').length;

  if (/\bsrc\s*=/.test(attrs)) continue;                       // arquivo externo
  const tipo = /\btype\s*=\s*["']([^"']+)["']/.exec(attrs);    // SO os atributos
  if (tipo && !/^(text\/javascript|module|application\/javascript)$/i.test(tipo[1])) continue;
  if (!corpo.trim()) continue;

  comCodigo++;
  try {
    new vm.Script(corpo, { filename: path + ' (script na linha ' + linhaTag + ')' });
    verificados++;
  } catch (e) {
    /* A linha do erro e relativa ao bloco: soma a linha da tag para dar a do arquivo. */
    const rel = /:(\d+)\n/.exec('\n' + (e.stack || '').split('\n').slice(0, 2).join('\n') + '\n');
    erros.push('  ' + path + ': bloco iniciado na linha ' + linhaTag + ' — ' + e.message
      + (rel ? ' (por volta da linha ' + (linhaTag + Number(rel[1])) + ' do arquivo)' : ''));
  }
}

if (!comCodigo) {
  console.error('FALHA: nenhum bloco <script> com codigo foi encontrado em ' + path + '.');
  console.error('Isso quase certamente e defeito DESTE verificador, nao do arquivo.');
  process.exit(2);
}
if (erros.length) {
  console.error('FALHA de sintaxe em ' + erros.length + ' bloco(s):');
  erros.forEach(function (e) { console.error(e); });
  process.exit(1);
}
console.log('OK: ' + verificados + ' bloco(s) com codigo compilam (' + path + ')');
