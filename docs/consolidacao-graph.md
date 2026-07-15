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

**39 de 52 migrados (75%)** — lotes 1 a 4, já em produção.

| Lote | Endpoints | Padrão |
|---|---|---|
| 1 | ListarDiretorias, ListarFornecedores, ConfigGet, ListarNotas, ListarGestoresFinanceiro | piloto (completo) |
| 2 | ConfigUpdate, SalvarDiretoria, AdicionarFornecedor, EditarFornecedor, MarcarProcessado (completo); AbrirPdfDaNota, AlertaDiario, RejeitarNota (só getGraphClient) | misto |
| 3 | AprovacaoViaLink, CriarListaContratos, CriarListaRecorrentes, DiagListaContratos, ListarMembrosGrupo, MigrarColunaAlinhamentoFinanceiro, MigrarColunaNFAuto, PushUnsubscribe, ReatribuirPendentes, AlertaContratosDiario | só getGraphClient |
| 4 | PushSubscribe, ConciliarRecorrente, GetControleAcessos, GetControleAcessosTelas, LimparContratosDuplicados, RelatorioAcessosContratos, SalvarControleAcessos, SalvarControleAcessosTelas, CaixaEntrada, SalvarRecorrente | completo (resolveSiteId) |

## Restam 13, por dificuldade

### Cat.2 — resolver "bundle" (só trocar `getGraphClient`, mantendo resolver local)
O resolver local retorna mais que `{siteId,listId}` (ex.: `invColMap`, `driveId`),
então NÃO é drop-in do `resolveSiteAndList`. Migra-se só o `getGraphClient`.
- `AlertaRecorrentes` → `{siteId, listNotasId, invColMap}`
- `ChecklistRecorrentes` → `{siteId, listNotasId}`
- `IntegrarOmie` → `{siteId, driveId, listNotasId}`
- `ListarRecorrentes` → `{siteId, listNotasId}`
- `MarcarContratosNotificados` → `{siteId, listId, colMap}`
- `MarcarNFsRejeitadasVistas` → `{siteId, listId, colMap}`

### Multi-lista — precisam de helper novo no shared
Resolvem 2+ listas de uma vez (`resolveSiteELists`/`resolveSiteEListas`).
Opção A: adicionar um helper no `shared/graph.js` que resolve várias listas de
uma vez. Opção B: migrar só o `getGraphClient` e deixar o resolver local.
- `ListarContratos`, `AbrirContrato`, `AtualizarStatusContrato`,
  `CancelarFornecedor`, `SanNotificacoesPendentes`

### Críticos — branch isolada e revisão dedicada
- `PostNota` — `resolveSiteAndDrive` (upload de PDF para o drive)
- `AprovarNota` — grava watermark APROVADO e MOVE o PDF de pasta; é o coração do
  fluxo. Migrar só o `getGraphClient` primeiro; não tocar na lógica de arquivo.

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
