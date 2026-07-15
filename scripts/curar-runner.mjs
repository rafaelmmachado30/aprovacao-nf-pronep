// Curadoria em background (P1.4) — roda no runner do GitHub Actions, SEM o teto de
// ~45s do gateway do Static Web Apps. Isso destrava Sonnet/Opus e tira a dependencia
// do navegador aberto.
//
// Fluxo:
//   1. POST ?prep=1        (app) — monta os textos do piloto a partir do indice RAG
//   2. GET  ?lista=1       (app) — devolve os textos (meta + texto)
//   3. curar(texto, ...)   (AQUI, no runner) — chama a Anthropic (Sonnet/Opus), sem cap
//   4. POST ?salvar=1      (app) — pos-processo (UF/markdown/nome) + upload do curado
//
// A logica de negocio (prompt, escada de modelos, markdown, UF, upload) continua no app;
// o runner so orquestra e paga o custo de tempo da IA.
//
// Env exigidas: BASE_URL, CURADORIA_SECRET, ANTHROPIC_API_KEY.
// Opcionais: MODEL (haiku|sonnet|opus, default sonnet), OPERADORAS (filtro),
//            SKIP_PREP (1 = nao refaz o prep), LIMIT (n = cura so os N primeiros).

import { createRequire } from 'module';
import { resolve } from 'path';

const require = createRequire(import.meta.url);
const { curar, ultimoErroCuradoria } = require(resolve('api/shared/curadoriaContrato.js'));

const BASE_URL = (process.env.BASE_URL || '').replace(/\/$/, '');
const SECRET = process.env.CURADORIA_SECRET || '';
const MODEL = process.env.MODEL || 'sonnet';
const OPERADORAS = process.env.OPERADORAS || '';
const SKIP_PREP = process.env.SKIP_PREP === '1';
const LIMIT = process.env.LIMIT ? parseInt(process.env.LIMIT, 10) : null;

if (!BASE_URL) { console.error('BASE_URL nao definido'); process.exit(1); }
if (!SECRET) { console.error('CURADORIA_SECRET nao definido'); process.exit(1); }
if (!process.env.ANTHROPIC_API_KEY) { console.error('ANTHROPIC_API_KEY nao definido'); process.exit(1); }

const HDRS = { 'X-Curadoria-Secret': SECRET, 'Accept': 'application/json' };
const qOps = OPERADORAS ? ('&operadoras=' + encodeURIComponent(OPERADORAS)) : '';

async function api(path, opts) {
  const res = await fetch(BASE_URL + path, opts);
  const text = await res.text();
  let json = null;
  try { json = JSON.parse(text); } catch (e) { /* nao-json */ }
  if (!res.ok) throw new Error('HTTP ' + res.status + ' em ' + path + ': ' + text.slice(0, 300));
  return json;
}

async function main() {
  const t0 = Date.now();

  if (!SKIP_PREP) {
    console.log('[prep] montando textos do piloto...');
    const prep = await api('/api/CurarContratos?prep=1' + qOps, { method: 'POST', headers: HDRS });
    console.log('[prep] ok — total no manifesto: ' + (prep && prep.total));
  } else {
    console.log('[prep] pulado (SKIP_PREP=1)');
  }

  console.log('[lista] buscando textos...');
  const lista = await api('/api/CurarContratos?lista=1' + qOps, { method: 'GET', headers: HDRS });
  let files = (lista && lista.files) || [];
  if (LIMIT != null) files = files.slice(0, LIMIT);
  console.log('[lista] ' + files.length + ' documento(s) para curar com modelo="' + MODEL + '"');

  const resumo = { ok: 0, skipped: 0, semTexto: 0, falha: 0 };
  for (let i = 0; i < files.length; i++) {
    const f = files[i];
    const rot = '[' + (i + 1) + '/' + files.length + '] ' + (f.nome || f.contratoId || '?');
    const texto = f.texto || '';
    if (texto.length < 60) { console.log(rot + ' — sem texto no indice, pulado'); resumo.semTexto++; continue; }
    try {
      const canonico = await curar(texto, 'nativo', { model: MODEL });
      if (!canonico) { console.log(rot + ' — curadoria falhou: ' + ultimoErroCuradoria()); resumo.falha++; continue; }
      const meta = { contratoId: f.contratoId, nome: f.nome, fornecedor: f.fornecedor, subpasta: f.subpasta, webUrl: f.webUrl };
      const saved = await api('/api/CurarContratos?salvar=1', {
        method: 'POST',
        headers: Object.assign({ 'Content-Type': 'application/json' }, HDRS),
        body: JSON.stringify({ file: meta, canonico: canonico })
      });
      if (saved && saved.skipped) { console.log(rot + ' — ignorado (' + saved.motivo + ')'); resumo.skipped++; }
      else { console.log(rot + ' — curado: ' + (saved && saved.curado)); resumo.ok++; }
    } catch (e) {
      console.log(rot + ' — ERRO: ' + ((e && e.message) || e));
      resumo.falha++;
    }
  }

  const secs = Math.round((Date.now() - t0) / 1000);
  console.log('\n=== FIM (' + secs + 's) — curados=' + resumo.ok + ' ignorados=' + resumo.skipped +
    ' semTexto=' + resumo.semTexto + ' falhas=' + resumo.falha + ' ===');
  // Nao falha o job por falhas pontuais de curadoria (best-effort); falha so se NADA curou.
  if (resumo.ok === 0 && files.length > 0) process.exit(1);
}

main().catch(function (e) { console.error('Fatal:', (e && e.stack) || e); process.exit(1); });
