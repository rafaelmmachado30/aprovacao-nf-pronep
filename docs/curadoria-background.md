# Curadoria em background (P1.4)

Tira a curadoria de contratos de dentro do request HTTP (teto de ~45s do Static Web
Apps) e move para o **runner do GitHub Actions**, que não tem esse limite. Resultado:
destrava **Sonnet/Opus** (antes só Haiku cabia em docs grandes) e elimina a
dependência do **navegador aberto** dirigindo o loop.

## Como funciona

```
GitHub Actions (runner, sem cap de tempo)
  1. POST /api/CurarContratos?prep=1     -> app monta os textos (indice RAG)
  2. GET  /api/CurarContratos?lista=1    -> app devolve textos (meta + texto)
  3. curar(texto, {model})               -> runner chama a Anthropic (Sonnet/Opus)
  4. POST /api/CurarContratos?salvar=1   -> app faz UF/markdown/nome + upload
```

A lógica de negócio (prompt, escada de modelos, markdown, regra de UF pela pasta,
upload) continua **no app** (`shared/curadoriaContrato.js`, `shared/fonteUnica.js`).
O runner só orquestra e absorve o tempo da IA. O `curar()` é reusado direto (o
workflow roda `npm ci` em `api/` e requer o módulo).

## Autenticação

`/api/CurarContratos` é rota `anonymous` no `staticwebapp.config.json`, mas o
endpoint se protege: aceita **`X-Curadoria-Secret`** (runner) OU sessão **admin**
(navegador). Sem um dos dois → 401/403. Mesmo padrão dos crons de alerta.

O fluxo antigo pelo navegador (`?prep`/`?offset&limit`) continua funcionando como
fallback (via sessão admin).

## Setup (uma vez)

1. **App Setting no Azure** (Static Web App → Configuration):
   - `CURADORIA_SECRET` = um valor secreto forte (ex.: `openssl rand -hex 24`)
2. **Repository Secrets no GitHub** (Settings → Secrets and variables → Actions):
   - `CURADORIA_SECRET` = **o mesmo valor** do App Setting acima
   - `ANTHROPIC_API_KEY` = a chave da Anthropic (a curadoria roda no runner)

> Opcional: se as App Settings do app usam IDs de modelo customizados
> (`ANTHROPIC_MODEL_SONNET`/`_OPUS`/`_HAIKU`), replicar como env no workflow para o
> runner usar os mesmos IDs. Sem isso, o runner usa os defaults do módulo (atuais).

## Como rodar

GitHub → **Actions → "Curar Contratos (background)" → Run workflow**. Inputs:
- **model**: `sonnet` (default) | `opus` | `haiku`
- **operadoras**: filtro opcional (ex.: `amil,bradesco`); vazio = todo o Comercial
- **limit**: curar só os N primeiros (teste); vazio = todos

Recomendado no primeiro uso: `model=sonnet`, `limit=2` para validar ponta a ponta
antes de rodar o acervo inteiro.

## Custo e limites

- **Infra: ~zero.** Usa minutos do GitHub Actions + o custo de LLM que já se pagaria.
- Job vai até 6h (`timeout-minutes: 330`) — folga de sobra para Sonnet/Opus.
- Sequencial por doc (amigável a rate limit e com log claro por documento).

## Se um dia a escala exigir (refinamento futuro)

Migrar para um Function App dedicado com Storage Queue (Premium, "Always On"). Mais
robusto e 100% server-side, mas com custo (~US$150+/mês) e infra/deploy próprios.
O caminho atual (runner) resolve o caso prático a custo ~zero.
