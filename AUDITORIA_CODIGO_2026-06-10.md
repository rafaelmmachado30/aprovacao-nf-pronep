# Auditoria de Código — Sistema Aprovação NF
**Data:** 10/06/2026 · **Escopo:** api/ (62 functions), wwwroot/, config, repo · **Modo:** somente leitura (nada foi alterado)

---

## 🔴 CRÍTICOS (corrigir antes de qualquer outra coisa)

### C1. Working tree local desatualizado — risco de regredir produção ⚠️ URGENTE
Os 4 arquivos modificados (`api/PostNota/index.js`, `api/SincronizarContratos/index.js`, `api/shared/sol.js`, `wwwroot/index.html`) **não são trabalho novo: são versões ANTIGAS** (mtime 5–9/jun, anteriores ao commit `d7aa76a` de 10/jun). O diff local **desfaz o fix P0 do PostNota** e a feature de Rejeitadas. Pior: `api/shared/sol.js` local está **truncado** — termina em `modu` (o `module.exports` foi cortado), o que quebraria toda a SAN em runtime.
**Correção:** `git restore api/PostNota/index.js api/SincronizarContratos/index.js api/shared/sol.js wwwroot/index.html`. Até lá, NÃO rodar `04_deploy_azure.bat` nem `git add -A`.

### C2. Toda a API aceita acesso anônimo — base do RBAC é forjável
Todas as 50 functions têm `authLevel: anonymous` e `staticwebapp.config.json` libera `/api/*` para `anonymous` + `authenticated`. O header `x-ms-client-principal` (base do `getUser`) só é confiável quando o SWA exige `authenticated` — como aceita anônimo, **qualquer pessoa na internet pode forjar o header em base64 e se passar por qualquer e-mail/role**.
**Correção (maior alavancagem de todas):** mudar `/api/*` para `allowedRoles: ["authenticated"]` no `staticwebapp.config.json`. Avaliar exceção apenas para `AprovacaoViaLink` e endpoints de alerta (que usam secret próprio).

### C3. `AbrirPdfDaNota` sem autenticação — IDOR vaza qualquer PDF de NF
`api/AbrirPdfDaNota/index.js` — recebe `?id=` e redireciona ao PDF sem nenhum `getUser` nem RBAC. Enumerando IDs, qualquer um baixa todos os PDFs de NF de todas as diretorias.
**Correção:** exigir auth + aplicar o mesmo escopo do `ListarNotas`.

### C4. `ConfigUpdate` sem checagem de admin
`api/ConfigUpdate/index.js:103-105` — comentário admite "frontend filtra a UI". Qualquer autenticado (ou anônimo, via C2) altera `gestorMasterGlobal`/`valorLimite`, redirecionando aprovações de 2º nível.
**Correção:** exigir role `administrador` via `getUserRoles`.

### C5. Endpoints admin/migração/debug sem auth
`MigrarColunasURL`, `ContratosInspecionarPasta` (dump de qualquer pasta SP via `?path=`), `OmieDebug` (expõe prefixos de app keys + dispara chamadas reais à Omie), `SolDebug`, `DiagListaNF`, `DiagListaContratos`, `ContratosDebug`, `ContratosTestarArquivo`, `EnviarNotificacao` (relay de e-mail/Teams para destinatários arbitrários).
**Correção:** gate de admin em todos; remover os de migração one-shot já executados.

### C6. XSS no frontend — falta de escape generalizada
`index.html` tem apenas uma função `esc()` local; quase todo o resto injeta dados da API via `innerHTML` sem escape. Vetores reais: chat SAN `renderAcao` (~linha 1149, resposta de LLM), `md()` linha 1104 (aceita link `javascript:`), auditoria (~3073), tabelas de NF/fila/aprovadas/rejeitadas, fornecedores (~5114). Nome de fornecedor com `<img onerror>` executa script na tela do aprovador.
**Correção:** criar `escHtml()` global e aplicar em TODO valor de API interpolado; validar esquema de URL no `md()`; migrar `onclick` com dados serializados para `dataset` + `addEventListener`.

### C7. `AprovacaoViaLink` — link de e-mail aprova sem identidade real
Fabrica um `x-ms-client-principal` sintético a partir do JWT (HS256, validade 7 dias, sem uso-único real). Link encaminhado/vazado = qualquer pessoa aprova/rejeita. Se a NF for reaberta, o link velho volta a funcionar.
**Correção:** exigir sessão autenticada + conferir `getUser(req).email === token.aprovador`; `jti` persistido e invalidado após uso; expiração 48h.

---

## 🟠 ALTOS

| # | Achado | Local | Correção |
|---|--------|-------|----------|
| A1 | Anti-duplicidade lê só 1ª página (999 itens) e não é atômico — NF duplicada escapa | `PostNota:316-319` | `$filter` por ChaveAcesso/Hash indexados + re-check pós-criação |
| A2 | Fallback "pega o PDF mais recente" pode carimbar/mover/**deletar o PDF errado** | `AprovarNota:394-397`, `RejeitarNota:275-277` | retornar 404 em vez de fallback |
| ~~A3~~ | ~~Auto-aprovação~~ — **DESCARTADO**: fluxo legítimo. O aprovador de Tecnologia lança e aprova a própria NF na ausência de outra pessoa, para a NF seguir o padrão e chegar ao financeiro. **Manter funcionando.** | `AprovarNota:283-292` | nenhuma ação |
| A4 | Admin hardcoded `rafael.machado@...` em 5+ arquivos; membro do grupo `administrador` não consegue aprovar | `AprovarNota:283`, `RejeitarNota:239`, `ListarNotas:143`, etc. | centralizar em env var/grupo AD |
| A5 | Bug de campo: `sol.js:862` filtra por `FornecedorCNPJ` mas a lista grava `CNPJFornecedor` → detecção de anomalia **sempre** retorna "histórico insuficiente"; fila da SAN sai com fornecedor `undefined` | `shared/sol.js:655,862` | usar `CNPJFornecedor` |
| A6 | Tools da SAN sem paginação: `$top=2000` num único GET (Graph trunca em ~999) — NFs somem silenciosamente da fila/relatórios | `sol.js:631/671/1090` | loop `@odata.nextLink` (como AlertaDiario faz) |
| A7 | `ListarNotas` baixa a base inteira (até 15k itens) + recarrega lista de fornecedores **a cada request** | `ListarNotas:124-218` | `$filter`+`$select` no Graph; cache TTL 5min |
| A8 | `escopoDoGestor` e mapa de aprovadores usam const **hardcoded**, ignorando o mapa dinâmico carregado da API — gestor pode ver/aprovar escopo errado | `index.html:4336,4237` | consumir `aprovadoresReaisDinamico` |
| A9 | `logout()` é cosmético — não chama `/.auth/logout`, cookie permanece (máquina compartilhada re-loga sozinha) | `index.html:2247` | redirecionar para `/.auth/logout` |
| A10 | Secret em query string no workflow de contratos (vaza em logs); endpoint aceita anônimo se env ausente | `alerta-contratos-diario.yml:37`, `AlertaContratosDiario:296-300` | header + secret obrigatório + `timingSafeEqual` |
| A11 | `pdf-parse@1.1.1` abandonado desde ~2018, usado em 3 arquivos | `api/package.json` | migrar p/ `unpdf`/`pdfjs-dist` |
| A12 | `x-sol-admin: true` enviado pelo front dá visão admin à SAN — forjável | `SolChat:32-37` | derivar de `getUserRoles`, nunca de header |

---

## 🟡 MÉDIOS

**Backend**
- Aprovação não atômica (sem eTag/If-Match): UI + link de e-mail simultâneos duplicam watermark (`AprovarNota:294/452`).
- Sem compensação em falha parcial: PDF movido mas PATCH falha → NF órfã (`AprovarNota:432-452`).
- `RejeitarNota:313` grava `AprovadoEm` na rejeição (deveria ser `RejeitadoEm` — sol.js espera esse campo e nunca acha).
- Duplicata CNPJ compara sem normalizar máscara (`PostNota:348`).
- `AdminLimparBase` é stub que retorna `ok:true` sem fazer nada — e o front ainda chama.
- Financeiro não consegue `escopo=todas` em rejeitadas na SAN (`sol.js:1094`).
- Sem retry/backoff em `fetch` direto (download PDF, Omie com rate-limit 60/min, webhooks Teams).
- Token Graph renegociado a cada request — `getGraphClient()` copiado em **36 arquivos**, sem singleton.
- Awaits sequenciais: `AlertaDiario:545-573` (IA + e-mail em série por gestor — estoura timeout 30s do SWA com ~10 gestores).
- Falta `$select` generalizada (payloads cheios em todos os `expand=fields`).
- OData injection: e-mail concatenado em `$filter` sem escapar `'` (`pushNotif.js:73,105`).
- HTML injection nas páginas do `AprovacaoViaLink` (interpola `e.message`/`aprovador` sem escape).
- `unidade`/`diretoria` entram em caminho SharePoint sem validar enum (`PostNota`) — criação de pastas arbitrárias.

**Frontend**
- Filtro da Fila re-renderiza a view inteira a cada tecla (perde foco do cursor, sem debounce) — Aprovadas/Fornecedores já fazem certo (`index.html:6071/5977`).
- `vendor/xlsx.full.min.js` (864 KB) carregado síncrono no `<head>` para todos — carregar sob demanda.
- Se a API falha, o app exibe as 11 NFs **mock** como se fossem reais, sem aviso.
- SW: versionamento manual de cache (`CACHE_VERSION` fixo) — vendor velho pode ser servido pós-deploy.
- Sem `navigationFallback` no staticwebapp.config.json (deep-link com path → 404).

**Config/Repo**
- `Azure/static-web-apps-deploy@latest` não pinado (supply chain); Node version indefinida (sem `engines` nem `apiRuntime`).
- `host.json` mínimo: sem timeout/retry/sampling de App Insights.
- `aplicar-fix-*.bat` fazem `git add` + push automático pra main (deploy sem revisão); deploy duplo (.bat + CI) pode divergir.
- PDF 4,2MB e xlsx de fornecedores soltos na raiz sem gitignore — um `git add -A` os commitaria.

---

## 🟢 BAIXOS
- Timezone: UTC-3 hardcoded em ~8 lugares; `contratos.js:503-525` usa hora do servidor (UTC) sem ajuste → entre 21h-00h BRT o status do contrato diverge ("Vencido" 3h antes).
- Race na auto-numeração de NF (`PostNota:281`, já documentada em comentário).
- Dead code: `Hello`, `HelloAdmin`, `ContratosPing`, `*Debug*`, `Diag*`, `Migrar*` (já executados), `api-disabled/` (rastreado no git).
- Arquivos `^C` e `main` (0 bytes) na raiz — lixo de terminal, deletar.
- `.env.exemplo` desatualizado (faltam 5 variáveis do `.env` real).
- IDs duplicados no DOM (`omie-step-2` 2x); duplicação `renderOpen`/`renderPanel` no chat.
- SDKs IA antigos (`@anthropic-ai/sdk 0.32`, `openai 4.65`); `isomorphic-fetch` obsoleto (Node 18+ tem fetch nativo) em 43 arquivos.
- `sort` por campo inexistente `n.Vencimento` (é `DataVencimento`) — lista do e-mail de alerta sai desordenada (`sol.js:642`, `AlertaDiario:281`).

---

## ✅ Pontos positivos
`ListarNotas` e tools da SAN reaplicam RBAC server-side; `shared/email.js` usa `escapeHtml`; validação JWT do Teams correta (JWKS/audience/issuer); `AuditLog` exige admin; estorno segregado a admin/financeiro; SW com `index.html` network-first; `.env` nunca foi commitado; tokens só via GitHub Secrets; sem dependência com CVE ativa conhecida.

---

## Plano de ação sugerido (ordem de execução)

**Hoje (15 min):**
1. C1 — `git restore` dos 4 arquivos regressivos + deletar `^C`, `main`, `aplicar-fix-*.bat`.

**Sprint imediato (segurança, ~1-2 dias):**
2. C2 — `/api/*` → `authenticated` no staticwebapp.config.json (1 linha, fecha C2/C5 parcialmente).
3. C3, C4, C5 — auth/role nos endpoints expostos; remover migrações e debugs.
4. C6 — `escHtml()` global no front.
5. C7 + A12 — refazer confiança do link de aprovação e do `x-sol-admin`.

**Sprint seguinte (corretude, ~2-3 dias):**
6. A1, A2, A3 — anti-duplicidade paginado, remover fallback de PDF, bloquear auto-aprovação.
7. A5, A6 — bugs de campo e paginação da SAN.
8. A8, A9 — mapa dinâmico de aprovadores e logout real.

**Backlog (qualidade/performance):**
9. `shared/graph.js` unificado (mata 36 duplicações + cache de token/colMap de uma vez).
10. `$filter`/`$select` no ListarNotas; xlsx sob demanda; debounce na Fila.
11. Pinagem de action, `apiRuntime: node:20`, host.json, migração do pdf-parse.
12. Helper `hojeBRT()` único; extrair `utils.js` do monolito de 7,7k linhas.
