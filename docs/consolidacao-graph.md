# Consolidação do `shared/graph.js`

Registro do esforço de eliminar o boilerplate do Microsoft Graph
(`getGraphClient` + resolução de site/lista) duplicado em dezenas de endpoints,
movendo tudo para `api/shared/graph.js`.

## Por que

No início, **52 endpoints** definiam a própria cópia de `getGraphClient()` e
~38 tinham a própria resolução de site/lista. Isso multiplicava pontos de falha
e fazia qualquer correção (retry, timeout, throttling 429) precisar ser repetida
em dezenas de arquivos. Bônus observado após a migração: o `siteId` passou a ser
resolvido **uma vez por processo** e compartilhado, melhorando a latência.

## O módulo `api/shared/graph.js`

| Função | Retorno | Uso |
|---|---|---|
| `getGraphClient()` | `Client` Graph (síncrono) | autenticação por app (client credentials) |
| `resolveSiteId(client)` | `siteId` (string) | cacheado no processo |
| `resolveList(client, siteId, nome)` | `listId` (string) | cacheado por nome de lista |
| `resolveSiteAndList(client, nome)` | `{ siteId, listId }` | atalho site + lista |

Env vars (App Settings), as mesmas de sempre: `AAD_TENANT_ID`, `AAD_CLIENT_ID`,
`AAD_CLIENT_SECRET`, `SHAREPOINT_SITE_HOSTNAME`, `SHAREPOINT_SITE_PATH`.

Nota: `getGraphClient` é **síncrono** (`Client.initWithMiddleware` não é async).
Chamadas antigas `await getGraphClient()` continuam válidas.

## Placar

**CONCLUÍDA: 56 endpoints migrados.** Todo endpoint que fala com o Graph usa o
`shared/graph.js`. Único não migrado: `ContratosDebug` (intencional — ver abaixo).

| Lote | Endpoints | Padrão |
|---|---|---|
| 1 | ListarDiretorias, ListarFornecedores, ConfigGet, ListarNotas, ListarGestoresFinanceiro | piloto (completo) |
| 2 | ConfigUpdate, SalvarDiretoria, AdicionarFornecedor, EditarFornecedor, MarcarProcessado (completo); AbrirPdfDaNota, AlertaDiario, RejeitarNota (só getGraphClient) | misto |
| 3 | AprovacaoViaLink, CriarListaContratos, CriarListaRecorrentes, DiagListaContratos, ListarMembrosGrupo, MigrarColunaAlinhamentoFinanceiro, MigrarColunaNFAuto, PushUnsubscribe, ReatribuirPendentes, AlertaContratosDiario | só getGraphClient |
| 4 | PushSubscribe, ConciliarRecorrente, GetControleAcessos, GetControleAcessosTelas, LimparContratosDuplicados, RelatorioAcessosContratos, SalvarControleAcessos, SalvarControleAcessosTelas, CaixaEntrada, SalvarRecorrente | completo (resolveSiteId) |
| 5 | AlertaRecorrentes, ChecklistRecorrentes, IntegrarOmie, ListarRecorrentes, MarcarContratosNotificados, MarcarNFsRejeitadasVistas | só getGraphClient (bundle) |
| 6 | AbrirContrato, AtualizarStatusContrato, CancelarFornecedor, ListarContratos, SanNotificacoesPendentes, PostNota, AprovarNota (multi-lista/críticos); MeusGrupos, DiagListaNF, MigrarColunasURL, ContratosInspecionarPasta (client inline) | só getGraphClient |

## O que ficou de fora (intencional, opcional)

### `ContratosDebug` — NÃO migrar
É um endpoint de diagnóstico que testa `require()` dos módulos do Graph (tem um
array `tries` e constrói o client pra provar que os módulos carregam). Migrar
removeria o propósito dele. Deixar como está.

### 6 módulos `shared/` com `getGraphClient` próprio — deixar por ora
Infraestrutura central (usada por vários endpoints), blast radius alto e ganho
marginal. Se um dia, migrar **um a um** com revisão dedicada — não em lote.
- `sol.js`, `email.js`, `auditLog.js`, `solHistorico.js`, `teamsActivity.js`,
  `contratos.js`

### Refinamento opcional: helper multi-lista no shared
5 endpoints (contratos) resolvem 2+ listas de uma vez com resolvers locais
(`resolveSiteELists`/`EListas`), preservados na migração (só o `getGraphClient`
foi trocado). Se quiser eliminar essa duplicação de resolução de lista, dá pra
adicionar um helper `resolveSiteAndLists(client, [nomes])` no `shared/graph.js`
e migrar esses resolvers depois.

## Processo de migração (receita validada)

1. Branch nova a partir do `main` (empilhar quando houver dependência).
2. Trocar `require`s do Graph + `getGraphClient` local pelo `require('../shared/graph')`.
3. Se o resolver local for `{siteId,listId}` puro → usar `resolveSiteAndList(client, NOME)`.
   Se retornar bundle → manter o resolver local, migrar só o `getGraphClient`.
4. Remover caches locais de siteId/listId que ficarem órfãos.
5. Verificar: `node scripts/smoke-test.mjs`, `node -e "require('./<ep>')"` de cada
   endpoint, e grep por símbolos órfãos (`ClientSecretCredential`, etc.).
6. Commit por lote, PR, smoke test roda sobre o PR, validar no preview.

## Regra de deploy

`shared/graph.js` já está no `main`/produção. Qualquer endpoint migrado só passa
a *usar* um módulo já existente — sem risco de "module not found". Mas um endpoint
migrado NUNCA deve ir a produção sem o `shared/graph.js` estar lá antes.
