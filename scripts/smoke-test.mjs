#!/usr/bin/env node
// Smoke test de sintaxe — rede de seguranca minima contra "tela branca".
//
// O que faz (sem executar nada, so PARSE):
//   1. `node --check` em todos os arquivos .js sob api/ (fora de node_modules)
//   2. Extrai cada <script> inline (sem src) do wwwroot/index.html e valida a sintaxe
//
// Sai com codigo != 0 se QUALQUER arquivo tiver erro de sintaxe.
// Nao instala nada, nao acessa a rede, nao toca em producao.
//
// Uso: node scripts/smoke-test.mjs

import { spawnSync } from 'node:child_process';
import { readFileSync, readdirSync, writeFileSync, rmSync, mkdtempSync } from 'node:fs';
import { join, relative } from 'node:path';
import { tmpdir } from 'node:os';
import { fileURLToPath } from 'node:url';

const ROOT = join(fileURLToPath(new URL('.', import.meta.url)), '..');
const API_DIR = join(ROOT, 'api');
const INDEX_HTML = join(ROOT, 'wwwroot', 'index.html');

let errors = 0;
let checked = 0;

/** Lista recursiva de .js em `dir`, pulando node_modules. */
function listJs(dir) {
  const out = [];
  for (const entry of readdirSync(dir, { withFileTypes: true })) {
    if (entry.name === 'node_modules' || entry.name === '.git') continue;
    const full = join(dir, entry.name);
    if (entry.isDirectory()) out.push(...listJs(full));
    else if (entry.name.endsWith('.js')) out.push(full);
  }
  return out;
}

/** Roda `node --check` num arquivo; retorna true se OK. */
function nodeCheck(file, label) {
  checked++;
  const res = spawnSync(process.execPath, ['--check', file], { encoding: 'utf8' });
  if (res.status !== 0) {
    errors++;
    console.error(`\n  ✗ ERRO DE SINTAXE: ${label}`);
    console.error((res.stderr || '').split('\n').slice(0, 6).map((l) => '    ' + l).join('\n'));
    return false;
  }
  return true;
}

// --- 1. Backend: api/**/*.js ---------------------------------------------
console.log('== 1. Backend (api/**/*.js) ==');
const jsFiles = listJs(API_DIR);
for (const f of jsFiles) nodeCheck(f, relative(ROOT, f));
console.log(`   ${jsFiles.length} arquivos backend verificados.`);

// --- 2. Front: <script> inline do index.html -----------------------------
console.log('\n== 2. Front (<script> inline de wwwroot/index.html) ==');
const html = readFileSync(INDEX_HTML, 'utf8');
// captura <script ...>...</script> IGNORANDO os que tem atributo src
const re = /<script(?![^>]*\bsrc=)[^>]*>([\s\S]*?)<\/script>/gi;
const tmp = mkdtempSync(join(tmpdir(), 'smoke-'));
let idx = 0;
let m;
while ((m = re.exec(html)) !== null) {
  const code = m[1];
  if (!code.trim()) continue;
  idx++;
  // numero da linha onde o bloco comeca, pra mensagem util
  const line = html.slice(0, m.index).split('\n').length;
  const tmpFile = join(tmp, `inline-${idx}.js`);
  writeFileSync(tmpFile, code);
  nodeCheck(tmpFile, `index.html <script> inline #${idx} (linha ~${line})`);
}
rmSync(tmp, { recursive: true, force: true });
console.log(`   ${idx} bloco(s) <script> inline verificado(s).`);

// --- Resultado -----------------------------------------------------------
console.log('\n' + '─'.repeat(48));
if (errors === 0) {
  console.log(`✓ SMOKE TEST OK — ${checked} unidade(s) sem erro de sintaxe.`);
  process.exit(0);
} else {
  console.error(`✗ SMOKE TEST FALHOU — ${errors} erro(s) em ${checked} unidade(s).`);
  process.exit(1);
}
