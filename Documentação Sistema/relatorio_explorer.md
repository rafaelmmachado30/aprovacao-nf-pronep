Perfeito! Agora tenho o mapeamento completo. Vou compilar o relatório:

---

# RELATÓRIO TÉCNICO: SISTEMA DE APROVAÇÃO DE NF PRONEP

## Documento de Mapeamento Completo para Reprodução em Outro Tenant/SharePoint

**Data:** 28 de maio de 2026  
**Versão do Sistema:** 1.0.2  
**Status:** Production-ready (Sprint 4 fechamento)  
**Escopo:** Mapeamento 100% técnico — arquitetura, configuração, dependências, permissões

---

## 1. ESTRUTURA GERAL DO PROJETO

Localização: `C:\Pronep\Aprovacao_NF\`

```
C:\Pronep\Aprovacao_NF\
├── LEIA-ME.md                          Visão geral + roadmap
├── GUIA_ENTRA_ID.md                    Setup 12 grupos + App Registration
├── GUIA_AZURE_SWA.md                   Criar Static Web App + deploy
├── GUIA_SHAREPOINT.md                  Criar 3 listas + estrutura pastas
├── SETUP_TIPOS_COLUNA_SHAREPOINT.md   Converter tipos de colunas
├── SETUP_TEAMS_SSO.md                  Configurar Teams SSO
├── SETUP_TEAMS_ACTIVITY.md             Notificações Teams 1:1
├── SETUP_PUSH_NOTIFICATIONS.md         Push notifications nativas
├── ROADMAP_SPRINT_4_FECHAMENTO.md     Status final + backlog
├── .env.exemplo                        Template variáveis (copiar → .env)
├── .gitignore                          Exclui .env, node_modules, logs
├── instalar.bat                        npm install + SWA CLI
├── 04_deploy_azure.bat                 Deploy wwwroot + api via SWA CLI
├── wwwroot/                            Frontend SPA (HTML/CSS/JS vanilla)
│   ├── index.html                      (~9000 linhas) — protótipo completo refatorado
│   ├── staticwebapp.config.json        Rotas, Easy Auth, headers, MIME types
│   ├── manifest.webmanifest            PWA manifest
│   ├── sw.js                           Service Worker (PWA offline)
│   ├── pronep-logo.png, favicon.*      Assets
│   ├── sem-acesso.html                 403 fallback (sem permissão)
│   ├── vendor/                         Libs locais (Chart.js, Teams SDK, XLSX)
│   │   ├── chart.umd.min.js            Chart.js para gráficos
│   │   ├── teams-js.min.js             Microsoft Teams JS SDK
│   │   └── xlsx.full.min.js            SheetJS (import XLSX)
│   └── ...
├── api/                                Azure Functions (Node.js v22)
│   ├── host.json                       Configuração runtime (ExtensionBundle v4)
│   ├── package.json                    Dependencies (10 pacotes)
│   ├── shared/                         Módulos compartilhados
│   │   ├── auth.js                     Easy Auth + Teams SSO token validation
│   │   ├── email.js                    Email Graph API + Teams cards
│   │   ├── teamsActivity.js            Teams sendActivityNotification
│   │   ├── pushNotif.js                Web Push (VAPID keys)
│   │   └── notificar.js                Re-export de email.js
│   ├── Hello/                          GET /api/Hello — health check
│   ├── MeusGrupos/                     GET /api/MeusGrupos — mapeia grupos → roles
│   ├── ListarNotas/                    GET /api/ListarNotas — RBAC server-side
│   ├── PostNota/                       POST /api/PostNota — lança NF + validação
│   ├── AprovarNota/                    POST /api/AprovarNota — aprova + watermark
│   ├── RejeitarNota/                   POST /api/RejeitarNota — rejeita + arquivo
│   ├── AprovacaoViaLink/               GET /api/AprovacaoViaLink — approve via JWT
│   ├── ListarFornecedores/             GET /api/ListarFornecedores — CRUD read
│   ├── AdicionarFornecedor/            POST /api/AdicionarFornecedor
│   ├── EditarFornecedor/               PATCH /api/EditarFornecedor
│   ├── ListarDiretorias/               GET /api/ListarDiretorias — matriz unidade x diretoria
│   ├── ListarGestoresFinanceiro/       GET /api/ListarGestoresFinanceiro
│   ├── AbrirPdfDaNota/                 GET /api/AbrirPdfDaNota — download PDF assinado
│   ├── MarcarProcessado/               PATCH /api/MarcarProcessado
│   ├── PushSubscribe/                  POST /api/PushSubscribe
│   ├── PushUnsubscribe/                POST /api/PushUnsubscribe
│   ├── PushPublicKey/                  GET /api/PushPublicKey — retorna VAPID_PUBLIC_KEY
│   ├── ConfigGet/                      GET /api/ConfigGet — lê PRONEP-NF-Config
│   ├── ConfigUpdate/                   PATCH /api/ConfigUpdate — escreve em Config
│   └── node_modules/                   (gerado por npm install)
├── api-disabled/                       Código antigo (deprecated)
│   ├── PostNota/                       v0.0.1 sketch (substituído por versão funcional)
│   └── package.json.original           Template histórico
├── teams-app/                          Teams App manifest + ícones
│   ├── manifest.json                   v1.16 schema — Personal Tab + Activities
│   ├── color.png                       Ícone colorido 192x192
│   ├── outline.png                     Ícone outline 32x32
│   ├── aprovacao-nf-teams.zip          Pacote v1.0.0
│   ├── aprovacao-nf-teams-v1.0.1.zip   Pacote v1.0.1 (com SSO)
│   └── aprovacao-nf-teams-v1.0.2.zip   Pacote v1.0.2 (com Activities)
├── .github/workflows/                  GitHub Actions CI/CD
│   └── azure-static-web-apps.yml       Trigger: push main / PR → deploy SWA
├── lib/                                Helpers (vazio — futuro)
├── logs/                               (runtime logs — local development)
├── backups/                            (snapshots SharePoint — futuro)
└── outputs/                            (arquivos PDF de documentação)
    └── Pronep_Aprovacao_NF_Arquitetura_v11.pdf
```

---

## 2. AZURE FUNCTIONS — MAPEAMENTO COMPLETO

### 2.1 Funções Implementadas (23 total)

| # | Nome | HTTP | Método | O que faz | Auth | Lista SharePoint | Permissões Graph |
|---|---|---|---|---|---|---|---|
| 1 | **Hello** | `/api/Hello` | GET | Health check — valida env vars (AAD_TENANT_ID, SHAREPOINT_SITE_HOSTNAME) | Anonymous | — | — |
| 2 | **MeusGrupos** | `/api/MeusGrupos` | GET | Lê Easy Auth header → obtém token Application via Client Credentials → Graph `/users/{userKey}/transitiveMemberOf` → mapeia GUIDs de grupo para roles | Requer x-ms-client-principal (Easy Auth) | — | **GroupMember.Read.All** |
| 3 | **ListarNotas** | `/api/ListarNotas` | GET | Le `PRONEP-NF-NotasFiscais` com filtros (?status=Lancada\|Aprovada\|Rejeitada, ?unidade=RJ\|SP\|ES, ?diretoria=...) + RBAC server-side (admin/financeiro→ve tudo, gestor→onde AprovadorAtual=email, submitter→onde LancadoPor=email) | Requer autenticação (Easy Auth ou Teams token) | `PRONEP-NF-NotasFiscais` | **Sites.ReadWrite.All** (Application) |
| 4 | **PostNota** | `/api/PostNota` | POST | Recebe NF em JSON (CNPJ, numero, serie, valor, vencimento, unidade, diretoria, fileBase64) → valida CNPJ (Brasilapi) → calcula SHA-256 do PDF → deteta duplicidade (hash + CNPJ+Numero+Serie) → extrai ChaveAcesso NF do PDF (regex 44 dígitos) → resolve aprovador via lista Diretorias (Unidade x Diretoria) → sobe PDF em `/Notas Fiscais/Notas Pendentes/{Unidade}/Diretoria {Diretoria}/` → cria item na lista NotasFiscais status PENDENTE → envia email + Teams card | Requer autenticação | `PRONEP-NF-NotasFiscais`, `PRONEP-NF-Diretorias`, `PRONEP-NF-Fornecedores` | **Sites.ReadWrite.All**, **Mail.Send** |
| 5 | **AprovarNota** | `/api/AprovarNota` | POST | Body: {id} → valida role (gestor da diretoria) → baixa PDF de Pendentes → aplica watermark "APROVADO" + timestamp + email aprovador (3 linhas azuis) via pdf-lib → sobe em `/Notas Aprovadas/{Unidade}/{AAAA-MM-DD}/` → deleta original de Pendentes → Status=Aprovada na lista → email + Teams card | Requer autenticação + role gestor | `PRONEP-NF-NotasFiscais`, `PRONEP-NF-Config` (multi-nível) | **Sites.ReadWrite.All**, **Mail.Send**, **TeamsActivity.Send** |
| 6 | **RejeitarNota** | `/api/RejeitarNota` | POST | Body: {id, motivo, observacao} → valida role → move PDF de Pendentes → `/Rejeitadas/{Unidade}/Diretoria {Diretoria}/` → Status=Rejeitada, MotivoRejeicao → email + Teams | Requer autenticação + role | `PRONEP-NF-NotasFiscais` | **Sites.ReadWrite.All**, **Mail.Send**, **TeamsActivity.Send** |
| 7 | **AprovacaoViaLink** | `/api/AprovacaoViaLink?token=...` | GET | Valida JWT assinado (7 dias) com {itemId, aprovador, action} → executa AprovarNota ou RejeitarNota sem exigir login via UI → utilitário pra links nos emails | Anonymous (valida JWT) | `PRONEP-NF-NotasFiscais` | **Sites.ReadWrite.All**, **Mail.Send**, **TeamsActivity.Send** |
| 8 | **ListarFornecedores** | `/api/ListarFornecedores` | GET | Query: ?q=texto (filtro), ?ativo=Sim (status), ?limit=50 (default 5000) → le `PRONEP-NF-Fornecedores` com paginação (nextLink) → retorna array com colMap normalizado | Requer autenticação | `PRONEP-NF-Fornecedores` | **Sites.Read.All** ou **Sites.ReadWrite.All** |
| 9 | **AdicionarFornecedor** | `/api/AdicionarFornecedor` | POST | Body: {CNPJ, RazaoSocial, NomeFantasia, UnidadeAtendimento, DiretoriaPadrao, Categoria, Ativo} → insere em `PRONEP-NF-Fornecedores` com validation → retorna item criado | Requer autenticação | `PRONEP-NF-Fornecedores` | **Sites.ReadWrite.All** |
| 10 | **EditarFornecedor** | `/api/EditarFornecedor` | PATCH | Body: {id, ...campos} → atualiza item em Fornecedores | Requer autenticação | `PRONEP-NF-Fornecedores` | **Sites.ReadWrite.All** |
| 11 | **ListarDiretorias** | `/api/ListarDiretorias` | GET | Le `PRONEP-NF-Diretorias` (matriz 27 linhas: 3 unidades × 9 diretorias) → retorna com colMap normalizado pra frontend popular dropdowns | Requer autenticação | `PRONEP-NF-Diretorias` | **Sites.Read.All** |
| 12 | **ListarGestoresFinanceiro** | `/api/ListarGestoresFinanceiro` | GET | Le membros do grupo Entra ID definido em `GESTOR_FINANCEIRO_GROUP_ID` env var (default hardcoded) → retorna email + displayName pra combobox "Negociar com" | Requer autenticação | — | **GroupMember.Read.All**, **User.Read.All** |
| 13 | **AbrirPdfDaNota** | `/api/AbrirPdfDaNota?itemId=...&tipo=original\|aprovado` | GET | Resolve URL do PDF no drive SharePoint (campo UrlPDF ou UrlPDFAprovado) → gera link assinado (SAS token válido 1h) → redireciona ou inline (Content-Disposition) | Requer autenticação + RBAC | `PRONEP-NF-NotasFiscais` | **Sites.ReadWrite.All** |
| 14 | **MarcarProcessado** | `/api/MarcarProcessado` | PATCH | Body: {id, processado: true/false} → toggle coluna "Processado" na lista NotasFiscais (marca como "ya visto" pra auditoria) | Requer autenticação | `PRONEP-NF-NotasFiscais` | **Sites.ReadWrite.All** |
| 15 | **PushSubscribe** | `/api/PushSubscribe` | POST | Body: {subscription: {endpoint, keys: {p256dh, auth}}, userAgent} → salva em `PRONEP-NF-PushSubscriptions` com Title=email do user (Easy Auth) | Requer autenticação | `PRONEP-NF-PushSubscriptions` | **Sites.ReadWrite.All** |
| 16 | **PushUnsubscribe** | `/api/PushUnsubscribe` | POST | Body: {endpoint} → remove subscription de PushSubscriptions | Requer autenticação | `PRONEP-NF-PushSubscriptions` | **Sites.ReadWrite.All** |
| 17 | **PushPublicKey** | `/api/PushPublicKey` | GET | Retorna {publicKey: process.env.VAPID_PUBLIC_KEY} em JSON pra frontend registrar ServiceWorker | Anonymous | — | — |
| 18 | **ConfigGet** | `/api/ConfigGet` | GET | Le `PRONEP-NF-Config` lista (1 item Title="global") → retorna campo ConfigJson (JSON blob) com config do sistema (ex: multiNivel.habilitado) | Requer autenticação | `PRONEP-NF-Config` | **Sites.Read.All** |
| 19 | **ConfigUpdate** | `/api/ConfigUpdate` | PATCH | Body: {configJson: {...}} → atualiza item global em Config — requer role administrador | Requer autenticação + role admin | `PRONEP-NF-Config` | **Sites.ReadWrite.All** |
| ~~20~~ | ~~**EnviarNotificacao**~~ | ~~(deprecated)~~ | — | Substituído por chamadas in-process em PostNota, AprovarNota, RejeitarNota (shared/email.js + shared/teamsActivity.js) | — | — | — |

### 2.2 Padrão de Desenvolvimento

**Todos os índices de funções seguem:**
```javascript
module.exports = async function (context, req) {
  try {
    // Step 1: validate user (getUser from shared/auth)
    const user = await getUser(req);
    if (!user) { context.res = {status: 401, body: {error: '...'}};  return; }
    
    // Step 2: get Graph client
    const client = await getGraphClient(); // ClientSecretCredential + GraphClient
    
    // Step 3: resolve site + list
    const {siteId, listId} = await resolveSiteAndList(client);
    
    // Step 4: business logic
    // ...
    
    // Step 5: return result
    context.res = {status: 200, body: {...}};
  } catch (e) {
    context.log.error('erro:', e);
    context.res = {status: 500, body: {error: e.message, diag: {...}}};
  }
};
```

**Cache em memória (reaproveitada entre invocações):**
- `cache = { siteId, listId, colMap, colTypes, colMapCachedAt }`
- TTL: 5 minutos (CACHE_TTL_MS = 5 * 60 * 1000)
- Evita múltiplas chamadas Graph pra mesmos metadados

**Normalização de colunas:**
- Graph API retorna `internalName` (ex: `OData__3wfZy5i`) ≠ `displayName` (ex: "Status")
- Código monta colMap bidireccional: `{displayName → internalName, internalName → displayName}`
- Aplica nos payloads Graph e na normalização de resposta

---

## 3. FRONTEND — SPA VANILLA (wwwroot/index.html)

### 3.1 Características

| Aspecto | Descrição |
|---|---|
| **Tipo** | Single Page Application (SPA) vanilla HTML/CSS/JS (sem framework) |
| **Linhas** | ~9000 (completo com CSS inline + JS embedded) |
| **Localização** | `C:\Pronep\Aprovacao_NF\wwwroot\index.html` |
| **Modo de execução** | Servido como estático pelo SWA; executa 100% no browser |
| **Estado** | Session storage (localStorage pra config do usuário) |
| **Roteamento** | Hash-based (#dashboard, #fila, #aprovadas, #meus-dados, etc.) |

### 3.2 Views (Abas do Menu)

1. **Dashboard** (`#dashboard`) — KPIs, gráficos
   - Cards: NFs aprovadas (mês), Valor aprovado (R$), NFs pendentes, Duplicidades evitadas
   - Gráficos: valor por diretoria (Chart.js), distribuição por unidade (pie), evolução mensal (line)
   - Mapa visual de Aprovadores (matriz por diretoria/unidade)

2. **Nova NF** (`#nova-nf`) — Formulário de lançamento
   - Campo: CNPJ/Fornecedor (combobox autocomplete)
   - Número, Série, Valor, Vencimento, Unidade (dropdown SP/RJ/ES), Diretoria (auto-resolvida)
   - Upload PDF com drag-and-drop
   - Preview PDF (inline embed ou link)
   - Botão "Lançar" chama POST /api/PostNota

3. **Fila de Aprovação** (`#fila`) — Notas pendentes
   - Tabela/Cards mobile: ID, Fornecedor, Valor, Vencimento, Status, Ações
   - Filtros: Status, Unidade, Diretoria, Data lançamento
   - Botões "Visualizar", "Aprovar", "Rejeitar"
   - RBAC: só gestores veem NFs da sua diretoria

4. **Notas Aprovadas** (`#aprovadas`)
   - Tabela: ID, Fornecedor, Valor, Aprovador, Data aprovação
   - Link "Abrir PDF" chama AbrirPdfDaNota?tipo=aprovado

5. **Notas Rejeitadas** (`#rejeitadas`)
   - Tabela: ID, Fornecedor, Valor, Motivo rejeição
   - Link "Re-lançar" (futuro)

6. **Meus Dados / Fornecedores** (`#fornecedores`)
   - CRUD: Listar, Adicionar, Editar, Deletar (soft)
   - Import XLSX/CSV com validação linha-a-linha
   - Combobox search (autocomplete por CNPJ/razao social)

7. **Configurações** (`#configuracoes`)
   - Push notifications: botão "Ativar notificações" (ServiceWorker registration)
   - Botão "Testar notificação" (dispara push local)
   - Info do usuário, grupos, roles

8. **Auditoria** (`#auditoria`) — Histórico
   - Tabela: Data, Evento, Usuario, Item, Detalhes
   - Filtros por data range, usuário, tipo evento

### 3.3 Tecnologias Embarcadas

| Lib | Versão | Localização | Uso |
|---|---|---|---|
| **Chart.js** | (local minified) | `/vendor/chart.umd.min.js` | Gráficos (linha, pizza, barra) no dashboard |
| **Teams JS SDK** | (local minified) | `/vendor/teams-js.min.js` | Detectar execução dentro Teams, chamar getAuthToken() SSO |
| **SheetJS** | (local minified) | `/vendor/xlsx.full.min.js` | Parse XLSX/CSV no browser pra importação Fornecedores |
| **pdf-lib** | — | Backend apenas (api/AprovarNota) | Aplicar watermark ao PDF |

### 3.4 Autenticação no Frontend

1. **Easy Auth (Browser)**
   - URL `/login` redireciona a `/.auth/login/aad` (configado no staticwebapp.config.json)
   - SWA injeta `x-ms-client-principal` header (base64 encoded principal JSON)
   - Frontend decodifica via JavaScript pra exibir nome + email
   - Automaticamente attachado a todas as requests `fetch()` ao `/api/*`

2. **Teams SSO (iframe da Personal Tab)**
   - Frontend chama `microsoftTeams.authentication.getAuthToken({scopes: [...]})`
   - Teams retorna MSAL token do user
   - Frontend passa via header custom `X-Teams-Token` ou `Authorization: Bearer`
   - Backend (shared/auth.js) valida contra JWKS Microsoft

### 3.5 Progressive Web App (PWA)

| Arquivo | Conteúdo |
|---|---|
| `manifest.webmanifest` | name: "Aprovacao NF", display: "standalone", icons, theme-color |
| `sw.js` | Service Worker — cache-first pra assets, network-first pra /api/* |
| Apple touch icon | `/apple-touch-icon.png` (pra "Agregar à tela de início") |

---

## 4. CONFIGURAÇÃO AZURE STATIC WEB APPS (wwwroot/staticwebapp.config.json)

```json
{
  "auth": {
    "identityProviders": {
      "azureActiveDirectory": {
        "registration": {
          "openIdIssuer": "https://login.microsoftonline.com/4b30645b-0888-45c0-9481-712bde435ffd/v2.0",
          "clientIdSettingName": "AAD_CLIENT_ID",
          "clientSecretSettingName": "AAD_CLIENT_SECRET"
        }
      }
    }
  },
  "routes": [
    {route: "/login", rewrite: "/.auth/login/aad"},
    {route: "/logout", redirect: "/.auth/logout"},
    {route: "/sem-acesso.html", allowedRoles: ["anonymous", "authenticated"]},
    {route: "/vendor/*", allowedRoles: ["anonymous", "authenticated"]},
    {route: "/manifest.webmanifest", allowedRoles: ["anonymous", "authenticated"], headers: {..., "Cache-Control": "no-cache"}},
    {route: "/sw.js", allowedRoles: ["anonymous", "authenticated"], headers: {..., "Service-Worker-Allowed": "/"}},
    {route: "/api/*", allowedRoles: ["anonymous", "authenticated"]},
    {route: "/*", allowedRoles: ["anonymous", "authenticated"]}
  ],
  "responseOverrides": {
    "403": {rewrite: "/sem-acesso.html", statusCode: 403}
  },
  "globalHeaders": {
    "Content-Security-Policy": "frame-ancestors 'self' https://teams.microsoft.com ...",
    "X-Content-Type-Options": "nosniff",
    "Referrer-Policy": "strict-origin-when-cross-origin"
  },
  "mimeTypes": {
    ".json": "application/json",
    ".webmanifest": "application/manifest+json"
  }
}
```

**Fluxo:**
1. Usuário acessa `https://aprovacao-nf-pronep.azurestaticapps.net`
2. SWA vê que nenhuma role setada → redireciona `/login` → `/.auth/login/aad`
3. Browser é levado pro login Microsoft (Entra ID)
4. Pós-login, volta pra URL original com cookie de sessão + header x-ms-client-principal
5. Frontend decodifica o header, extrai email + nome
6. Fetch automático pra `/api/MeusGrupos` retorna roles mapeadas
7. Frontend renderiza menu baseado em roles

**Headers de segurança:**
- CSP: permite frames do Teams (`frame-ancestors 'self' https://teams.microsoft.com`)
- X-Content-Type-Options: nosniff (bloqueia MIME sniffing)
- Referrer-Policy: strict-origin-when-cross-origin (não vaza query params em navegação)

---

## 5. GITHUB ACTIONS CI/CD (.github/workflows/azure-static-web-apps.yml)

```yaml
name: Azure Static Web Apps CI/CD

on:
  push:
    branches: [main]
  pull_request:
    types: [opened, synchronize, reopened, closed]
    branches: [main]

jobs:
  build_and_deploy_job:
    if: github.event_name == 'push' || (github.event_name == 'pull_request' && github.event.action != 'closed')
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4
      - name: Build And Deploy
        uses: Azure/static-web-apps-deploy@latest
        with:
          azure_static_web_apps_api_token: ${{ secrets.AZURE_STATIC_WEB_APPS_API_TOKEN }}
          repo_token: ${{ secrets.GITHUB_TOKEN }}
          action: "upload"
          app_location: "/wwwroot"        # Frontend estático
          api_location: "/api"            # Azure Functions
          output_location: ""             # Sem build step

  close_pull_request_job:
    if: github.event_name == 'pull_request' && github.event.action == 'closed'
    runs-on: ubuntu-latest
    steps:
      - uses: Azure/static-web-apps-deploy@latest
        with:
          azure_static_web_apps_api_token: ${{ secrets.AZURE_STATIC_WEB_APPS_API_TOKEN }}
          action: "close"
```

**Secrets necesários no GitHub:**
- `AZURE_STATIC_WEB_APPS_API_TOKEN` — obtém em Azure Portal > SWA > Gerenciar token de implantação

**Fluxo:**
1. Push pra main → GitHub Actions roda automaticamente
2. Checkout repo
3. `Azure/static-web-apps-deploy` faz upload de `/wwwroot` + `/api` pro SWA
4. SWA compila Functions + publica HTML + pronto em ~2-3 min
5. PR merges desencadeiam novo deploy

---

## 6. MICROSOFT TEAMS APP (teams-app/manifest.json)

```json
{
  "$schema": "https://developer.microsoft.com/en-us/json-schemas/teams/v1.16/MicrosoftTeams.schema.json",
  "manifestVersion": "1.16",
  "version": "1.0.2",
  "id": "5c52fce3-bf34-4b8c-a624-06defd5f85f6",
  "packageName": "br.com.pronep.aprovacao-nf",
  "developer": {
    "name": "Pronep Life Care",
    "websiteUrl": "https://www.pronep.com.br",
    "privacyUrl": "https://www.pronep.com.br/privacidade",
    "termsOfUseUrl": "https://www.pronep.com.br/termos"
  },
  "icons": {
    "color": "color.png",
    "outline": "outline.png"
  },
  "name": {
    "short": "Aprovacao NF",
    "full": "Pronep - Aprovacao de NF"
  },
  "description": {
    "short": "Notificacoes de NF para aprovadores",
    "full": "Sistema de Aprovacao de Notas Fiscais da Pronep Life Care. Envia notificacoes 1:1 no Teams quando uma NF aguarda aprovacao, foi aprovada ou rejeitada."
  },
  "accentColor": "#1F4E79",
  "staticTabs": [
    {
      "entityId": "aprovacao-nf-home",
      "name": "Aprovacao NF",
      "contentUrl": "https://purple-forest-09588fe10.7.azurestaticapps.net/",
      "websiteUrl": "https://purple-forest-09588fe10.7.azurestaticapps.net/",
      "scopes": ["personal"]
    }
  ],
  "permissions": ["identity"],
  "validDomains": ["purple-forest-09588fe10.7.azurestaticapps.net"],
  "webApplicationInfo": {
    "id": "85f7fa68-a241-4caf-803f-9991bd1f0eee",
    "resource": "api://purple-forest-09588fe10.7.azurestaticapps.net/85f7fa68-a241-4caf-803f-9991bd1f0eee"
  },
  "activities": {
    "activityTypes": [
      {
        "type": "approvalRequired",
        "description": "Nova NF aguardando aprovacao",
        "templateText": "NF {nfNumber} aguardando sua aprovacao"
      },
      {
        "type": "nfAprovada",
        "description": "NF aprovada",
        "templateText": "Sua NF {nfNumber} foi APROVADA"
      },
      {
        "type": "nfRejeitada",
        "description": "NF rejeitada",
        "templateText": "Sua NF {nfNumber} foi REJEITADA"
      }
    ]
  }
}
```

**Campos críticos:**
- `id` (externalId) — GUID da Teams App (usado em filtros `externalId eq '5c52fce3-...'`)
- `contentUrl` / `websiteUrl` — domínio do SWA (ex: `purple-forest-09588fe10.7.azurestaticapps.net`)
- `webApplicationInfo.id` — Client ID da App Registration (85f7fa68-...)
- `webApplicationInfo.resource` — Application ID URI com domínio + Client ID
- `validDomains` — lista branca de domínios (impedindo iframe cross-origin)
- `activities` — templates de notificações (sendActivityNotification enche os placeholders `{nfNumber}`)

**Como usar:**
1. Zipar: `Compress-Archive -Path .\manifest.json,.\color.png,.\outline.png -DestinationPath aprovacao-nf-teams.zip`
2. Upload no Teams Admin Center: Manage apps → + Upload new app
3. Aguardar processamento (~30s)
4. Permitir na política de apps se necessário

---

## 7. VARIÁVEIS DE AMBIENTE (Azure SWA App Settings)

Todas as funções dependem dessas variáveis. Devem ser configuradas no Azure Portal: SWA > Configuration > Application settings.

### 7.1 Variáveis Obrigatórias (todas)

| Variável | Valor de exemplo | Origem | Descrição |
|---|---|---|---|
| **AAD_TENANT_ID** | `4b30645b-0888-45c0-9481-712bde435ffd` | Entra ID | Tenant ID (mesmo do Analytics) |
| **AAD_CLIENT_ID** | `85f7fa68-a241-4caf-803f-9991bd1f0eee` | App Reg "Pronep Aprovacao NF SWA" | Application (client) ID |
| **AAD_CLIENT_SECRET** | (gerado, ~36 chars) | App Reg → Certificados e segredos | Client Secret (NUNCA commit, só em Azure) |
| **SHAREPOINT_SITE_HOSTNAME** | `pronepadmin.sharepoint.com` | URL SharePoint | Host sem /sites/... |
| **SHAREPOINT_SITE_PATH** | `/sites/Aprovacao-NotasFiscaisServicos` | URL SharePoint | Caminho relativo (com /sites/) |

### 7.2 Variáveis Opcionais (com defaults)

| Variável | Default | Descrição |
|---|---|---|
| **APP_ID_URI** | `api://purple-forest-09588fe10.7.azurestaticapps.net/{CLIENT_ID}` | Teams SSO — Application ID URI (com domínio SWA) |
| **TEAMS_APP_ID** | `5c52fce3-bf34-4b8c-a624-06defd5f85f6` | Manifest externalId (pra sendActivityNotification achar o catalog app) |
| **EMAIL_FROM_ADDRESS** | `datanalytics@pronep.com.br` | Remetente dos emails transacionais (Mail.Send Graph) |
| **LINK_APROVACAO_SECRET** | (gerado, ~32 chars) | JWT secret pra links assinados (7 dias) em emails — OBRIGATÓRIO se usar AprovacaoViaLink |
| **VAPID_PUBLIC_KEY** | (gerado) | Push Notifications — chave pública (gerada uma única vez via Python script) |
| **VAPID_PRIVATE_KEY** | (gerado, **SECRETO**) | Push Notifications — chave privada (NUNCA commit, só em Azure) |
| **VAPID_SUBJECT** | `mailto:datanalytics@pronep.com.br` | Push Notifications — email de contato VAPID |
| **GESTOR_FINANCEIRO_GROUP_ID** | (hardcoded no código) | Entra ID — GUID do grupo "Gestores Financeiro" (pra ListarGestoresFinanceiro) |
| **ADMIN_GROUP_ID** | (vazio por default) | Entra ID — opcional, valida role admin em ConfigUpdate |

---

## 8. SHAREPOINT LISTS — ESTRUTURA

### 8.1 Listas Necessárias (criar no site `Aprovacao-NotasFiscaisServicos`)

#### Lista 1: `PRONEP-NF-NotasFiscais` (crítica)

| Display Name | Tipo interno | Tipo SharePoint | Obrigatório | Índice | Observação |
|---|---|---|---|---|---|
| **ChaveAcesso** | ChaveAcesso | Linha de texto (44 chars) | Sim | ✓ | NFS-e key (regex 44 dígitos) — key primária de duplicidade |
| **HashArquivo** | HashArquivo | Linha de texto (64 chars) | Sim | ✓ | SHA-256 do PDF — uniqueness validation |
| **NumeroNF** | NumeroNF | Linha de texto | Sim | | Permite zeros à esquerda (ex: 000123) |
| **Serie** | Serie | Linha de texto | Sim | | |
| **Fornecedor** | Fornecedor | Lookup (de lista Fornecedores) | Sim | | |
| **Valor** | Valor | **Moeda (R$)** | Sim | | Tipo correto: Currency (já migrado via SETUP_TIPOS_COLUNA) |
| **Vencimento** | Vencimento | Data | Sim | | |
| **UnidadeAtendimento** | UnidadeAtendimento | Escolha (SP\|RJ\|ES) | Sim | | |
| **DiretoriaPadrao** | DiretoriaPadrao | Lookup (de Diretorias) | Sim | | |
| **Status** | Status | Escolha (LANCADA\|APROVADA\|REJEITADA\|AUTO_REJEITADA) | Sim | | |
| **MotivoRejeicao** | MotivoRejeicao | Várias linhas de texto | Não | | Preenchido se Status=REJEITADA |
| **UrlPDF** | UrlPDF | **Hiperlink** | Sim | | Tipo correto (via SETUP_TIPOS_COLUNA) |
| **UrlPDFAprovado** | UrlPDFAprovado | **Hiperlink** | Não | | Preenchido após aprovação (com watermark) |
| **LancadoPor** | LancadoPor | Texto | Sim | | Email do submitter |
| **LancadoEm** | LancadoEm | Data/Hora | Sim | | Timestamp |
| **AprovadorAtual** | AprovadorAtual | Texto | Sim | | Email do gestor (resolvido via matriz Diretorias) |
| **AprovadoPor** | AprovadoPor | Texto | Não | | Email do aprovador (post-approval) |
| **AprovadoEm** | AprovadoEm | Data/Hora | Não | | Timestamp de aprovação |
| **ForaPrazo** | ForaPrazo | Sim/Não | Sim | | Vencimento < D+5 (auto-calc ou flag no lançamento) |
| **NegociouCom** | NegociouCom | Texto | Não | | Email do gestor financeiro (pra negociação D+5) |
| **NegociouComEmail** | NegociouComEmail | Texto | Não | | Idem (redundância pra query) |
| **Processado** | Processado | Sim/Não | Não | | Flag do usuário: "marcada como vista" (auditoria) |
| **CaminhoSharePoint** | CaminhoSharePoint | Hiperlink | Não | | Path completo no drive (/Notas Pendentes/...) |
| **Descricao** | Descricao | Várias linhas de texto | Não | | Campo livre |

#### Lista 2: `PRONEP-NF-Fornecedores`

| Display Name | Tipo Interno | Tipo SharePoint | Obrigatório | Índice |
|---|---|---|---|---|
| **CNPJ** (ou Documento) | Documento | Linha de texto | Sim | ✓ (unique) |
| **RazaoSocial** | RazaoSocial | Linha de texto | Sim | |
| **NomeFantasia** | NomeFantasia | Linha de texto | Não | |
| **UnidadeAtendimento** | UnidadeAtendimento | Escolha (SP\|RJ\|ES) | Sim | |
| **DiretoriaPadrao** | DiretoriaPadrao | Linha de texto | Sim | |
| **Categoria** | Categoria | Escolha (Serviços\|Materiais\|Locação\|Reembolso\|Medicamentos\|Outros) | Sim | |
| **Ativo** | Ativo | Sim/Não | Sim | |

#### Lista 3: `PRONEP-NF-Diretorias` (matriz unidade × diretoria)

| Display Name | Tipo Interno | Tipo SharePoint | Observação |
|---|---|---|---|
| **Unidade** | Unidade | Escolha (SP\|RJ\|ES) | |
| **Diretoria** | Diretoria | Linha de texto | Ex: "Suprimentos", "Técnica", "Financeira", ... |
| **GrupoEntraID** | GrupoEntraID | Linha de texto | GUID do grupo Entra ID (ex: `01d540d1-8596-...`) |
| **GestorEmail** | GestorEmail | Linha de texto | Email do aprovador atual (ex: bruno.hioka@pronep.com.br) |

**Dados de exemplo (27 linhas totais):**
```
SP | Suprimentos      | 01d540d1-8596-... | bruno.hioka@pronep.com.br
SP | Técnica          | fc3a375b-329c-... | nicolle.boas@pronep.com.br
SP | Financeira       | 6b77405b-ba89-... | henrico.molina@pronep.com.br
... (mais 24 linhas combinando 3 unidades × 9 diretorias)
```

#### Lista 4: `PRONEP-NF-PushSubscriptions` (Web Push)

| Display Name | Tipo Interno | Tipo SharePoint | Observação |
|---|---|---|---|
| **Title** | Title | Linha de texto | Email do usuário (coluna padrão do SharePoint) |
| **Endpoint** | Endpoint | **Várias linhas de texto** | URL do service worker push endpoint (200-400 chars) |
| **P256DH** | P256DH | Linha de texto | Chave pública da subscription |
| **Auth** | Auth | Linha de texto | Auth secret da subscription |
| **UserAgent** | UserAgent | Linha de texto | Browser user agent (debug) |

#### Lista 5: `PRONEP-NF-Config` (system configuration)

| Display Name | Tipo Interno | Tipo SharePoint | Observação |
|---|---|---|---|
| **Title** | Title | Linha de texto | Sempre "global" (1 item) |
| **ConfigJson** | ConfigJson | Várias linhas de texto | JSON blob: `{multiNivel: {habilitado: false, ...}}` |

### 8.2 Estrutura de Pastas no Drive (Documentos)

Já existente (criada pelo Power Automate antigo) — o novo sistema apenas reaproveita:

```
Notas Fiscais/
├── Notas Pendentes/
│   ├── SP/
│   │   ├── Diretoria Suprimentos/
│   │   ├── Diretoria Técnica/
│   │   ├── Diretoria Financeira/
│   │   ├── Diretoria RH-DP/
│   │   ├── Diretoria Tecnologia/
│   │   ├── Diretoria Fiscal-Contábil/
│   │   ├── Diretoria Jurídica/
│   │   ├── Diretoria Administrativa/
│   │   └── Diretoria Qualidade/
│   ├── RJ/ (mesmas 9 diretorias)
│   └── ES/ (mesmas 9 diretorias)
├── Notas Aprovadas/
│   ├── SP/{AAAA-MM-DD}/ (criadas dinamicamente por data)
│   ├── RJ/{AAAA-MM-DD}/
│   └── ES/{AAAA-MM-DD}/
└── Notas Rejeitadas/
    ├── SP/ (9 diretorias)
    ├── RJ/ (9 diretorias)
    └── ES/ (9 diretorias)
```

---

## 9. GRUPO ENTRA ID & MAPEAMENTO DE ROLES

### 9.1 Os 13 Grupos (criados via GUIA_ENTRA_ID.md)

| # | Nome do Grupo | GUID Real (Placeholder) | Role Mapeada | Membros Iniciais |
|---|---|---|---|---|
| 1 | `PRONEP-NF-Submitter` | `01d540d1-8596-42d0-9a20-de5c361c7c96` | submitter | Todos operacional (RH-DP dinâmico) |
| 2 | `PRONEP-NF-Admin` | `480a1595-bdc3-492a-9ef2-317f148a237e` | administrador | Rafael Machado |
| 3 | `PRONEP-Financeiro-Gestao` | `c2a73d16-4659-4b3c-93a1-0c0fbfaaaa96` | financeiro_nf | Sandra Ferreira, Vitor Costa, Monica Pires, Bruno Mendes |
| 4 | `PRONEP-NF-Gestor-Suprimentos` | `2d9f5bcf-2ae0-494e-957b-a1c69016664d` | gestor_suprimentos | Bruno Hioka |
| 5 | `PRONEP-NF-Gestor-Financeira` | `6b77405b-ba89-47ee-af21-58ec19bb3ff7` | gestor_financeira | Henrico Molina |
| 6 | `PRONEP-NF-Gestor-Tecnologia` | `a7826b5c-7c29-4a24-836b-a7432aa941ec` | gestor_tecnologia | Rafael Machado |
| 7 | `PRONEP-NF-Gestor-Qualidade` | `a6711877-8746-4ca5-a955-c15980c7e90d` | gestor_qualidade | Sabrina Fernandes |
| 8 | `PRONEP-NF-Gestor-RH-DP` | `13a544d8-3dde-4820-9695-c492e58a2782` | gestor_rh_dp | Janilene Santos |
| 9 | `PRONEP-NF-Gestor-Fiscal-Contabil` | `b9272d98-3e26-4e2d-aae6-ff9057f57e5c` | gestor_fiscal_contabil | Janilene Santos |
| 10 | `PRONEP-NF-Gestor-Juridica` | `5aa9fc6b-900d-40eb-861d-8bbf72499da1` | gestor_juridica | Rafaella Santos |
| 11 | `PRONEP-NF-Gestor-Administrativa` | `4f28f31b-b704-4615-961b-a9ca0898cea8` | gestor_administrativa | Rafaella Santos |
| 12 | `PRONEP-NF-Gestor-Tecnica-SP` | `fc3a375b-329c-4d9c-81be-06180f0598af` | gestor_tecnica_sp | Nicolle Boas |
| 13 | `PRONEP-NF-Gestor-Tecnica-RJES` | `334eb19b-c138-4551-8e45-a36ca4e32e48` | gestor_tecnica_rjes | Vitor Amaral |

### 9.2 GROUP_TO_ROLE (hardcoded em api/MeusGrupos/index.js)

```javascript
const GROUP_TO_ROLE = {
  '01d540d1-8596-42d0-9a20-de5c361c7c96': 'submitter',
  '480a1595-bdc3-492a-9ef2-317f148a237e': 'administrador',
  'c2a73d16-4659-4b3c-93a1-0c0fbfaaaa96': 'financeiro_nf',
  '2d9f5bcf-2ae0-494e-957b-a1c69016664d': 'gestor_suprimentos',
  '6b77405b-ba89-47ee-af21-58ec19bb3ff7': 'gestor_financeira',
  'a7826b5c-7c29-4a24-836b-a7432aa941ec': 'gestor_tecnologia',
  'a6711877-8746-4ca5-a955-c15980c7e90d': 'gestor_qualidade',
  '13a544d8-3dde-4820-9695-c492e58a2782': 'gestor_rh_dp',
  'b9272d98-3e26-4e2d-aae6-ff9057f57e5c': 'gestor_fiscal_contabil',
  '5aa9fc6b-900d-40eb-861d-8bbf72499da1': 'gestor_juridica',
  '4f28f31b-b704-4615-961b-a9ca0898cea8': 'gestor_administrativa',
  'fc3a375b-329c-4d9c-81be-06180f0598af': 'gestor_tecnica_sp',
  '334eb19b-c138-4551-8e45-a36ca4e32e48': 'gestor_tecnica_rjes'
};
```

**Ao criar em novo tenant:**
1. Criar os 13 grupos com os MESMOS nomes
2. Anotar os GUIDs verdadeiros
3. Substituir os GUIDs no código acima
4. Preencher a lista Diretorias com os GUIDs reais

---

## 10. SHARED MODULES (api/shared/)

### 10.1 auth.js

**Responsabilidade:** Extrair usuário autenticado de 2 fontes.

```javascript
async getUser(req) -> { email, name, oid, source, claims? }
```

**Lógica:**
1. Tenta Easy Auth (header `x-ms-client-principal` base64)
   - Decodifica JSON → principal.userDetails (email), principal.userId (SWA internal ID)
   - Extrai claims.oid (Entra ID real) se existente
2. Tenta Teams SSO Bearer (header `X-Teams-Token` ou `Authorization`)
   - Valida JWT contra JWKS Microsoft
   - Confere audience (`APP_ID_URI` ou `CLIENT_ID`)
   - Confere issuer (2 opções: `login.microsoftonline.com/.../v2.0` ou `sts.windows.net/.../`)
3. Retorna `{ email, name, oid, source: 'easy-auth'|'teams-sso', claims }`

**Erro:** retorna `null` (Function checa e retorna 401)

### 10.2 email.js

**Responsabilidade:** Construir + enviar emails transacionais + Teams cards.

```javascript
async notificar(evento, destinatarios, dados, links?) 
  -> { ok, emailIds[], teamsResultado? }
```

**Fluxo:**
1. Constrói corpo HTML baseado em `evento` (lancada|aprovada|rejeitada)
2. Se `evento === 'lancada'`: inclui botões "Aprovar" + "Rejeitar" (links JWT assinados)
3. Chama Graph API `POST /users/{email}/sendMail` (Mail.Send)
4. Chama `enviarTeamsAtividade(...)` se TEAMS_WEBHOOK_URL ou sendActivityNotification

**Variáveis consumidas:**
- `AAD_TENANT_ID`, `AAD_CLIENT_ID`, `AAD_CLIENT_SECRET` (Client Credentials)
- `EMAIL_FROM_ADDRESS` (remetente)
- `LINK_APROVACAO_SECRET` (JWT secret, 7 dias validade)
- `TEAMS_WEBHOOK_URL` (deprecated: canais do Teams via webhook)

### 10.3 teamsActivity.js

**Responsabilidade:** Enviar notificações 1:1 via Graph sendActivityNotification.

```javascript
async enviarTeamsAtividade(email, tipoAtividade, dados)
  -> { ok, resultado }
```

**Fluxo:**
1. Resolve `userId` do email via Graph `/users/{email}`
2. Obtém `catalogAppId` do Teams app (filtra por `externalId`)
3. Garante que a app está instalada no Teams pessoal do user
   - Se não: `POST /users/{userId}/teamwork/installedApps` (auto-install)
4. Envia `POST /users/{userId}/teamwork/sendActivityNotification`
   - Payload: `{ topic, activityType, resource, resourceData: {data} }`
   - Teams renderiza a notificação no sino com ícone + texto template

**Variáveis consumidas:**
- `AAD_TENANT_ID`, `AAD_CLIENT_ID`, `AAD_CLIENT_SECRET`
- `TEAMS_APP_ID` (externalId do manifest)

**Permissões Graph necessárias:**
- `TeamsActivity.Send` (enviar notificações)
- `TeamsAppInstallation.ReadWriteForUser.All` (auto-install app)
- `AppCatalog.Read.All` (descobrir app no catálogo)
- `User.Read.All` (resolver userId do email)

### 10.4 pushNotif.js

**Responsabilidade:** Gerenciar Web Push Notifications (VAPID).

```javascript
configurarWebPush() -> bool  // inicializa lib web-push
async enviarPushPraEmail(client, siteId, email, payload) -> resultado[]
```

**Fluxo:**
1. Initializa `web-push` com VAPID keys (VAPID_PUBLIC_KEY + VAPID_PRIVATE_KEY)
2. Lista todas as subscriptions do email em `PRONEP-NF-PushSubscriptions`
3. Pra cada subscription:
   - Chama `webpush.sendNotification(subscription, JSON.stringify(payload))`
   - Retorna resultado ou erro
4. Lida com casos especiais (subscrição expirada, endpoint inválido)

**Variáveis consumidas:**
- `VAPID_PUBLIC_KEY` (compartilhada com frontend)
- `VAPID_PRIVATE_KEY` (**SECRETO** — nunca no código)
- `VAPID_SUBJECT` (email de contato VAPID)

---

## 11. LISTAS SHAREPOINT REFERENCIADAS NO CÓDIGO

**Resumo de constantes hardcoded (procurar por `LIST_` ou `const.*= '.*-NF-'`):**

| Constante | Display Name | Arquivo(s) | Uso |
|---|---|---|---|
| `LIST_NOTAS` | `PRONEP-NF-NotasFiscais` | ListarNotas, PostNota, AprovarNota, RejeitarNota, AprovacaoViaLink, MarcarProcessado, AbrirPdfDaNota | CRUD notas fiscais |
| `LIST_DIRETORIAS` | `PRONEP-NF-Diretorias` | PostNota, ListarDiretorias | Matriz aprovadores |
| `LIST_FORNECEDORES` | `PRONEP-NF-Fornecedores` | ListarFornecedores, AdicionarFornecedor, EditarFornecedor, PostNota | CRUD fornecedores |
| `LIST_SUBSCRIPTIONS` | `PRONEP-NF-PushSubscriptions` | pushNotif.js, PushSubscribe, PushUnsubscribe | Web Push subscriptions |
| `LIST_CONFIG` | `PRONEP-NF-Config` | AprovarNota, ConfigGet, ConfigUpdate | System config (multiNivel, etc.) |

---

## 12. PACKAGE.JSON DO API (api/package.json)

```json
{
  "name": "pronep-aprovacao-nf-api",
  "version": "1.0.1",
  "description": "API functions para o Sistema de Aprovacao de NF da Pronep",
  "scripts": {
    "test": "echo \"No tests\""
  },
  "dependencies": {
    "@azure/identity": "^4.0.0",
    "@microsoft/microsoft-graph-client": "^3.0.7",
    "isomorphic-fetch": "^3.0.0",
    "jsonwebtoken": "^9.0.2",
    "jwks-rsa": "^3.2.2",
    "pdf-lib": "^1.17.1",
    "web-push": "^3.6.7"
  }
}
```

**Dependências explicadas:**
- `@azure/identity` — ClientSecretCredential (App Registration auth)
- `@microsoft/microsoft-graph-client` — Graph API client + auth middleware
- `isomorphic-fetch` — fetch() pra Node.js (compat com browser fetch)
- `jsonwebtoken` — gera + valida JWT (links assinados, Teams tokens)
- `jwks-rsa` — obtém chaves públicas do JWKS Microsoft (validar tokens)
- `pdf-lib` — aplica watermarks aos PDFs (AprovarNota)
- `web-push` — envia push notifications (VAPID)

---

## 13. DOCUMENTAÇÃO MARKDOWN EXISTENTE — RESUMOS

### LEIA-ME.md
Visão geral do projeto, estrutura de pastas, status Sprint 0-5, quem é quem (matriz aprovadores), contato. **Essencial pra onboarding.**

### GUIA_ENTRA_ID.md
**Passo a passo (5 etapas):**
1. Criar App Registration "Pronep Aprovacao NF SWA" + client secret
2. Criar 13 grupos no Entra ID (especificar nome exato + membros)
3. Colar GUIDs reais em `api/MeusGrupos/index.js`
4. Confirmar tenant ID em staticwebapp.config.json
5. Próximo passo: GUIA_AZURE_SWA.md

**Diagnósticos:** login 403 (usuário sem grupo), GetRoles retorna roles vazio (grupos não sendo emitidos no token — checar Token configuration), AADSTS50011 (redirect URI errada).

### GUIA_AZURE_SWA.md
**8 etapas:**
1. Criar SWA no Portal Azure (nome: `aprovacao-nf-pronep`, Free tier)
2. Pegar deployment token (SWA > Visão geral > Gerenciar token de implantação)
3. Configurar redirect URI na App Reg (`https://aprovacao-nf-pronep.azurestaticapps.net/.auth/login/aad/callback`)
4. App Settings do SWA: `AAD_CLIENT_ID`, `AAD_CLIENT_SECRET`
5. Adicionar permissões Graph (Sites.Selected, Mail.Send)
6. Instalar SWA CLI (via `instalar.bat`)
7. Deploy: `.\04_deploy_azure.bat` (powered by `$env:DEPLOYMENT_TOKEN`)
8. Testar: acessar URL, logar, ver protótipo

**Custo:** Free tier SWA cobre tudo (unlimited static hosting + 100K requests/mês pra Functions).

### GUIA_SHAREPOINT.md
**Criar 3 listas + estrutura:**
1. `Fornecedores` — CNPJ (unique), RazaoSocial, UnidadeAtendimento, DiretoriaPadrao, Categoria, Ativo
2. `NotasFiscais` — ChaveAcesso (unique), HashArquivo (unique), NumeroNF, Valor, Status, LancadoPor, AprovadorAtual, etc.
3. `Diretorias` — Matriz 27 linhas (3 unidades × 9 diretorias) com GrupoEntraID + email gestor

**Pastas já existentes** (reaproveitadas do Power Automate antigo):
- Notas Pendentes/{Unidade}/Diretoria {Diretoria}/ (PDF aqui até aprovação)
- Notas Aprovadas/{Unidade}/{AAAA-MM-DD}/ (PDF com watermark)
- Notas Rejeitadas/{Unidade}/Diretoria {Diretoria}/ (PDF rejeitado)

**Sites.Selected permission:** Usar Graph Explorer pra autorizar App Reg no site específico (POST /sites/{site-id}/permissions).

### SETUP_TIPOS_COLUNA_SHAREPOINT.md
**Débito técnico (migração de tipos):**
1. **Valor** → Currency (R$) — Graph detecta automaticamente, mas interface UI clássica cria como Número
2. **NumeroNF** → Single line of text (não Número — pra preservar zeros à esquerda)
3. **UrlPDF** / **UrlPDFAprovado** → Hyperlink or Picture (não Texto)

Backend já suporta todos os 3 tipos via `formatByType()`. Depois que migrar no SharePoint, redeploy não necessário.

### SETUP_TEAMS_SSO.md
**3 ações pra Teams SSO dentro da Personal Tab:**
1. Expose an API na App Reg:
   - Application ID URI: `api://{SWA_DOMAIN}/{CLIENT_ID}`
   - Scope: `access_as_user`
2. Autorizar 3 clientes Teams:
   - Desktop/mobile: `1fec8e78-bce4-4aaf-ab1b-5451cc387264`
   - Web: `5e3ce6c0-2b1f-4285-8d4b-75ee78787346`
   - Office web: `4765445b-32c6-49b0-83e6-1d93765276ca`
3. Atualizar manifest: zipar + upload Teams Admin Center

Frontend chama `microsoftTeams.authentication.getAuthToken({scopes: [...]})` → pega MSAL token do Teams → passa via `X-Teams-Token` header → backend valida.

### SETUP_TEAMS_ACTIVITY.md
**3 ações pra sendActivityNotification:**
1. Adicionar 3 permissões Application na App Reg:
   - `TeamsActivity.Send`
   - `TeamsAppInstallation.ReadWriteForUser.All`
   - `AppCatalog.Read.All`
   - (grant admin consent)
2. Anotar Client ID da App Reg
3. Zipar teams-app/ + upload Teams Admin Center
   - Manifest: `webApplicationInfo.id` = Client ID
   - File: `aprovacao-nf-teams-v1.0.2.zip`
   - Resultado: Teams app visível em Manage apps com status **Allowed**

Backend (teamsActivity.js) faz a notificação 1:1 ao aprovador quando NF é lançada.

### SETUP_PUSH_NOTIFICATIONS.md
**5 ações (VAPID setup):**
1. Gerar VAPID keys (uma única vez):
   ```python
   python3 -c "from cryptography.hazmat.primitives.asymmetric import ec; ..."
   ```
   → VAPID_PUBLIC_KEY, VAPID_PRIVATE_KEY, VAPID_SUBJECT
2. Adicionar 3 variáveis em Azure SWA App Settings
3. Criar lista `PRONEP-NF-PushSubscriptions` no SharePoint
   - Columns: Title (email), Endpoint (multi-line), P256DH, Auth, UserAgent
4. Permissões Graph já existentes (Sites.ReadWrite.All)
5. Usuários ativam manualmente em **Configurações** → "Ativar notificações"

Push é **canal extra** (não substitui email nem Teams). Só funciona em PWA instalada (desktop/mobile).

### ROADMAP_SPRINT_4_FECHAMENTO.md
**Estado:** tudo pronto pra production (Sprint 4 = fechamento).

**Implementado:**
- Lançamento NF + validação CNPJ + duplicidade
- Aprovação com watermark + archive
- Email + Teams notifications
- Dashboard + KPIs
- CRUD Fornecedores (import XLSX)
- SSO Teams + Personal Tab
- Push notifications (Web Push + VAPIR)

**Pendentes (backlog nível baixo):**
- PWA (instalável mobile) — ~1h30
- Testar JWT Teams silent SSO — ~15min
- Multi-aprovador / cadeia de aprovação (N1 → N2)

**Setups manuais SharePoint:**
- Converter coluna Valor → Currency ✓
- Converter NumeroNF → Text ✓
- Converter UrlPDF/UrlPDFAprovado → Hyperlink ✓

---

## 14. ARQUIVOS CRÍTICOS — QUICK REFERENCE

| Arquivo | Linhas | Propósito |
|---|---|---|
| `C:\Pronep\Aprovacao_NF\wwwroot\index.html` | ~9000 | Frontend SPA (100% nele) |
| `C:\Pronep\Aprovacao_NF\api\MeusGrupos\index.js` | ~250 | Mapeia grupos Entra ID → roles (HARDCODE GUIDs aqui) |
| `C:\Pronep\Aprovacao_NF\api\shared\auth.js` | ~170 | Easy Auth + Teams SSO validation |
| `C:\Pronep\Aprovacao_NF\api\shared\email.js` | ~350 | Email + Teams cards |
| `C:\Pronep\Aprovacao_NF\api\PostNota\index.js` | ~400 | Lança NF (anti-duplicidade, roteamento) |
| `C:\Pronep\Aprovacao_NF\api\AprovarNota\index.js` | ~450 | Aprova + watermark (pdf-lib) |
| `C:\Pronep\Aprovacao_NF\wwwroot\staticwebapp.config.json` | ~125 | Rotas, Easy Auth, headers, CSP |
| `C:\Pronep\Aprovacao_NF\teams-app\manifest.json` | ~66 | Teams App config (update `contentUrl`, `id`, `webApplicationInfo`) |
| `C:\Pronep\Aprovacao_NF\.github\workflows\azure-static-web-apps.yml` | ~46 | CI/CD GitHub Actions |
| `.env.exemplo` | ~40 | Template de variáveis (copiar → .env) |

---

## 15. CHECKLIST DE REPRODUÇÃO EM NOVO TENANT

### **Fase 0 — Preparação (1 dia)**
- [ ] Solicitar acesso Azure Admin (subscription)
- [ ] Solicitar acesso Teams Admin Center
- [ ] Solicitar acesso SharePoint Admin Center
- [ ] Clonar/copiar repositório GitHub
- [ ] Ler LEIA-ME.md + GUIA_ENTRA_ID.md

### **Fase 1 — Entra ID (1-2 horas)**
- [ ] Criar App Registration "Pronep Aprovacao NF SWA"
- [ ] Criar client secret (copiar valor → .env)
- [ ] Criar 13 grupos Entra ID (anotar GUIDs)
- [ ] Editar `api/MeusGrupos/index.js` com GUIDs reais
- [ ] Configurar Token configuration (emitir claim "groups")

### **Fase 2 — Azure (30-45 min)**
- [ ] Criar Static Web App (name: `aprovacao-nf-pronep`)
- [ ] Pegar deployment token
- [ ] Configurar redirect URI no App Reg
- [ ] Adicionar App Settings (AAD_CLIENT_ID, AAD_CLIENT_SECRET, SHAREPOINT_*)
- [ ] Adicionar permissões Graph (Sites.Selected, Mail.Send, TeamsActivity.Send, etc.)
- [ ] Testar `/api/Hello` (verificar env vars)

### **Fase 3 — SharePoint (2 horas)**
- [ ] Criar site `Aprovacao-NotasFiscaisServicos` (se novo)
- [ ] Criar 5 listas (NotasFiscais, Fornecedores, Diretorias, Config, PushSubscriptions)
- [ ] Criar estrutura de pastas (Notas Pendentes, Aprovadas, Rejeitadas)
- [ ] Popular lista Diretorias com matriz (27 linhas)
- [ ] Converter tipos de colunas (Value → Currency, etc.)
- [ ] Autorizar App Reg via Sites.Selected (Graph Explorer)

### **Fase 4 — Teams (1-2 horas)**
- [ ] Editar `teams-app/manifest.json` (contentUrl, id, webApplicationInfo)
- [ ] Adicionar 3 permissões Graph (TeamsActivity.Send, TeamsAppInstallation.*, AppCatalog.Read.All)
- [ ] Zipar + upload Teams Admin Center
- [ ] Permitir app na política de apps global/específica
- [ ] Testar Teams SSO (abrir Personal Tab)

### **Fase 5 — Deploy (30 min)**
- [ ] Instalar Node + SWA CLI (`instalar.bat`)
- [ ] Configurar `DEPLOYMENT_TOKEN` no PowerShell
- [ ] Rodar `04_deploy_azure.bat`
- [ ] Acessar URL final, logar, testar views

### **Fase 6 — Setup Opcional (1-2 horas)**
- [ ] Gerar VAPID keys (push notifications)
- [ ] Configurar VAPIR no Azure + lista PushSubscriptions
- [ ] Testar link de aprovação via email
- [ ] Testar Teams activity notifications
- [ ] Configurar GitHub Actions secret (AZURE_STATIC_WEB_APPS_API_TOKEN)

**Tempo total:** 1 semana (~30 horas de trabalho, podendo reduzir com paralelização)

---

## 16. TROUBLESHOOTING COMUM

| Sintoma | Causa | Solução |
|---|---|---|
| Login redireciona pro login da Microsoft mas volta com 403 | Redirect URI errado | Verificar App Reg → Autenticação → Redirect URI deve ser exato: `https://aprovacao-nf-pronep.azurestaticapps.net/.auth/login/aad/callback` |
| `/api/Hello` retorna 500 (AAD_TENANT_ID missing) | App Settings não configuradas | Abrir Azure Portal → SWA → Configuration → verificar todos os AAD_* e SHAREPOINT_* estão lá |
| MeusGrupos retorna `roles: []` | Grupos não emitidos no token | Ir App Reg → Token configuration → "Adicionar declaração de grupos" → marcar "ID do grupo" |
| PostNota falha com "Lista PRONEP-NF-NotasFiscais nao encontrada" | Nome da lista errado ou site incorreto | Verificar SHAREPOINT_SITE_HOSTNAME e SHAREPOINT_SITE_PATH (sem trailing slash) |
| PDF approvement falha "pdf-lib error" | PDF pode estar corrompido ou muito grande | Checar tamanho do PDF (máx 6MB esperado), tentar PDF diferente |
| Teams app não aparece em Manage apps | Manifest.json com erro ou não ziped corretamente | Re-zipar: `Compress-Archive -Path manifest.json,color.png,outline.png ...` |
| Notificações Teams 1:1 não aparecem | Permissões Graph incompletas | Verificar se `TeamsActivity.Send`, `TeamsAppInstallation.ReadWriteForUser.All`, `AppCatalog.Read.All` têm green check |
| Push notifications "VAPID keys not configured" | VAPID_PUBLIC_KEY / VAPID_PRIVATE_KEY ausentes | Gerar via Python script, colar em Azure App Settings, restart SWA |
| Emails não chegam | EMAIL_FROM_ADDRESS errado ou sem permissão Mail.Send | Verificar account, testar com account@pronep.com.br, adicionar permissão Mail.Send na App Reg |

---

## 17. DIAGRAMA DE FLUXO ALTO NÍVEL

```
[User browser]
    ↓ login
[Azure Static Web Apps Easy Auth]
    ↓ x-ms-client-principal header
[Frontend SPA (index.html)]
    ↓ fetch /api/MeusGrupos (getUser + MeusGrupos)
[Azure Functions (Node.js)]
    ↓ Graph API (Client Credentials)
[Microsoft Graph API]
    ├→ /users/{userKey}/transitiveMemberOf (mapping grupos)
    ├→ /sites/{siteId}/lists (PRONEP-NF-NotasFiscais, etc.)
    ├→ /users/{email}/sendMail (notifications)
    └→ /users/{userId}/teamwork/sendActivityNotification (Teams 1:1)
[SharePoint Lists + Drive]
    ├→ Metadata (columns, display names)
    └→ PDFs (/Notas Pendentes/, /Notas Aprovadas/, etc.)

[User in Teams]
    ↓ Teams Personal Tab opens
[Frontend (iframe) + Teams JS SDK]
    ↓ getAuthToken SSO
[Teams MSAL token] → [Backend validates via JWKS]
    ↓ /api/MeusGrupos (Teams SSO path)
[Azure Functions (com X-Teams-Token)]
    ↓ Graph API
[SharePoint Lists]
```

---

## CONCLUSÃO

Este sistema é **100% reproduzível** em outro tenant/SharePoint desde que se siga:

1. **Entra ID:** 13 grupos + App Reg com permissões Graph corretas
2. **Azure:** SWA + App Settings (18 variáveis, 9 obrigatórias + 9 opcionais)
3. **SharePoint:** 5 listas + estrutura de pastas
4. **Teams:** Manifest manifest + upload admin center
5. **GitHub:** Secret de deploy (CI/CD)

**Não há "estado oculto"** — tudo está em código, configuração ou lista SharePoint. Basta seguir o checklist de reprodução e os guias MD.

**Tempos estimados:**
- Setup infra (AAD + Azure + SharePoint): 3-4 dias
- Testes + troubleshooting: 1-2 dias
- Go-live: 1 dia (treinamento + cutover Power Automate antigo)

**Custos:**
- Azure Static Web Apps Free: $0 (até 100K requests/mês)
- SharePoint Online: já incluído (M365)
- Entra ID: já incluído (M365)
- Teams: já incluído (M365)

---

**Fim do Relatório Técnico.**

Qualquer dúvida sobre mapeamento, permissões ou passo a passo, favor consultar os MDs (LEIA-ME.md, GUIA_*.md) ou este documento.