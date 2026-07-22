# Padrão de Acessos — Entra ID + Azure Static Web Apps

Modelo de referência para autenticação e autorização de sistemas internos Pronep,
extraído da implementação do **Sistema de Aprovação de NF**. Serve como base a copiar
em sistemas futuros — com uma folha do que é fixo e do que se troca por projeto.

> **Referências de origem** (neste repo): [`GUIA_ENTRA_ID.md`](../GUIA_ENTRA_ID.md),
> [`GUIA_AZURE_SWA.md`](../GUIA_AZURE_SWA.md), [`GUIA_SHAREPOINT.md`](../GUIA_SHAREPOINT.md),
> [`SETUP_TEAMS_SSO.md`](../SETUP_TEAMS_SSO.md), [`wwwroot/staticwebapp.config.json`](../wwwroot/staticwebapp.config.json),
> e os módulos `api/shared/{auth,authz,graph,userRoles}.js`.

| | |
|---|---|
| **Tenant** | `4b30645b-0888-45c0-9481-712bde435ffd` (único em todos os sistemas Pronep) |
| **Identidade** | Microsoft Entra ID |
| **Hospedagem** | Azure Static Web Apps (plano **Standard** — obrigatório, ver seção 8) |
| **Dados** | SharePoint Online via Microsoft Graph (app-only) |

---

## 1. Arquitetura de referência

O usuário nunca fala direto com os dados. Cada requisição atravessa a porta de
autenticação do SWA, passa pela resolução de papéis e só então chega às funções e ao
SharePoint.

```
[Colaborador @pronep.com.br]              browser OU aba pessoal do Teams
        │
        ▼
[Easy Auth do Static Web App]             porta: bloqueia quem não tem sessão Entra ID
        │                                 (bloco "auth" do staticwebapp.config.json)
        ▼
[Microsoft Entra ID]                      valida login → token com oid + claim de grupos
        │
        ▼
[Resolução de papéis]                     grupos de segurança → papéis da aplicação
        │                                 (GetRoles / shared/authz.js)
        ▼
[Frontend estático + Azure Functions]     wwwroot/ via CDN · api/* exige "authenticated"
        │
        ▼
[SharePoint Online via Microsoft Graph]   app-only com Sites.Selected (um site específico)
```

---

## 2. Os cinco blocos do padrão

Todo sistema novo reusa estes cinco componentes. Cada um é uma peça configurável do
Azure/Entra, com um artefato de código correspondente no repositório.

### Bloco 1 · Identidade — App Registration dedicada
Uma App Registration **separada por sistema** (nunca compartilhada), no tenant único
da Pronep. É a identidade da aplicação — dona do login, do client secret e das
permissões de Graph.
- Emite o claim de **grupos de segurança** (Token configuration → Group ID) para o backend enxergar os papéis.
- Guarda o **client secret** (validade 2 anos) — só no App Settings do SWA e no `.env` local (nunca no git).
- Recebe as permissões de aplicativo do Graph (`Sites.Selected`, `Mail.Send`) com consentimento de admin.

### Bloco 2 · Papéis — Grupos de segurança do Entra ID
Cada papel do sistema é um **grupo de segurança atribuído**. O grupo é a fonte de
verdade de "quem pode o quê" — a gestão de acesso vira gestão de membros no Entra,
sem tocar em código.
- Nomenclatura padronizada: `PRONEP-<SISTEMA>-<Papel>` (ver seção 3).
- O **Object ID (GUID)** de cada grupo é o que amarra grupo → papel no código.
- Um grupo base de acesso mínimo (ex.: `-Submitter`) evita o 403 de "usuário sem nenhum grupo".

### Bloco 3 · Porta de entrada — Static Web App com Easy Auth
Plano **Standard** (~US$ 9/mês — o bloco `auth` customizado é recusado no deploy do
Free), região East US 2 (o conteúdo estático é CDN global). O bloco `auth` do
`staticwebapp.config.json` liga o Easy Auth ao Entra ID e o roteamento decide o que é
público e o que exige sessão.
- `/api/*` exige **`authenticated`** — fecha spoofing do header `x-ms-client-principal`.
- Rotas de **cron/webhook** ficam anônimas mas autenticam por secret próprio no endpoint.
- Deploy manual via SWA CLI com **deployment token** (sem GitHub Actions obrigatório).

### Bloco 4 · Autorização — Resolução de papéis centralizada
O SWA **não** popula os papéis a partir dos grupos automaticamente. Um módulo
compartilhado (`shared/userRoles.js` + `shared/authz.js`) consulta os grupos via Graph
e entrega flags prontas: `isAdmin`, `isGestor`, etc.
- Mapa `GROUP_TO_ROLE`: GUID do grupo → papel da aplicação.
- Merge de **3 fontes** de papéis num só array (ver seção 4).
- Guardas reutilizáveis: `requireAdmin(context, req)` no topo de cada endpoint sensível.

### Bloco 5 · Dados — Microsoft Graph app-only, escopo mínimo
O acesso ao SharePoint é **app-only** (client credentials), autenticado pela mesma App
Registration. Cliente Graph centralizado em `shared/graph.js`, com site e listas
resolvidos e cacheados no processo.
- `Sites.Selected` em vez de `Sites.ReadWrite.All` — acesso concedido a **um site específico**, não a todo o tenant.
- Segredos via env: `AAD_TENANT_ID` / `AAD_CLIENT_ID` / `AAD_CLIENT_SECRET`.
- Site resolvido por `SHAREPOINT_SITE_HOSTNAME` + `SHAREPOINT_SITE_PATH`.

### Bloco opcional · SSO dentro do Teams
Se o sistema tiver uma aba no Teams, adiciona-se um caminho de autenticação por token
MSAL. Só ligar quando houver app do Teams — o browser continua no Easy Auth normal.
- Expose an API + scope `access_as_user` + 3 clientes Teams autorizados (GUIDs fixos da Microsoft).
- Backend valida o JWT contra o JWKS do tenant (`shared/auth.js`).
- Token viaja no header custom **`X-Teams-Token`** — o SWA sobrescreve o `Authorization` padrão.

---

## 3. Convenção de nomenclatura de grupos

Formato único: `PRONEP-<SISTEMA>-<Papel>`. Deixa claro, só de bater o olho no Entra,
a qual sistema e a qual papel cada grupo pertence.

| Papel funcional | Nome do grupo (exemplo NF) | Papel no código |
|---|---|---|
| Acesso mínimo / submissão | `PRONEP-NF-Submitter` | `submitter` |
| Administrador do sistema | `PRONEP-NF-Admin` | `administrador` |
| TI / suporte técnico | `PRONEP-NF-TI` | `ti` |
| Gestor de uma diretoria | `PRONEP-NF-Gestor-<Área>` | `gestor_<area>` |
| Papel com recorte regional | `PRONEP-NF-Gestor-Tecnica-SP` | `gestor_tecnica_sp` |

**Regra de ouro:** ao criar um grupo, **copie o Object ID na hora** — é ele, não o
nome, que entra no `GROUP_TO_ROLE`. Mantenha o mapa sincronizado nos dois lugares onde
ele existe (`userRoles.js` e `MeusGrupos`).

---

## 4. Modelo de autorização — merge de três fontes

Um usuário pode chegar por caminhos diferentes, e cada caminho carrega os papéis de um
jeito. A autorização une as três fontes num único conjunto e remove os papéis default
inúteis do SWA.

- **Fonte A · Claims do Teams SSO** — do JWT validado (`user.claims.roles`), quando o acesso vem da aba do Teams.
- **Fonte B · Principal do Easy Auth** — os `userRoles` do `x-ms-client-principal` no browser, descartando `authenticated` e `anonymous`.
- **Fonte C · Graph `transitiveMemberOf`** (canônica) — todos os grupos do usuário via Graph, mapeados pelos GUIDs conhecidos. Cacheada 5 min por usuário.

```js
// shared/authz.js — resolveAuthz(req)
roles = unique([ ...claimsRoles, ...principalRoles, ...graphRoles ])

// flags derivadas, prontas para os endpoints:
isAdmin      = roles.has('administrador') || ADMIN_EMAILS.has(email)
isGestor     = roles.some(r => r.startsWith('gestor'))
isFinanceiro = roles.has('financeiro_nf')
```

---

## 5. Roteamento — `staticwebapp.config.json`

O arquivo de rotas é a segunda camada de defesa. Padrão: tudo em `/api/*` fechado,
exceções explícitas e nomeadas para páginas públicas e crons que se autenticam por secret.

```jsonc
// nega por padrão — fecha spoofing de x-ms-client-principal
{ "route": "/api/*",            "allowedRoles": ["authenticated"] }

// exceções nomeadas: cron/webhook autenticado por secret no endpoint
{ "route": "/api/AlertaDiario", "allowedRoles": ["anonymous","authenticated"] }
{ "route": "/api/Aquecer",      "allowedRoles": ["anonymous","authenticated"] }

// páginas públicas mínimas (login, sem-acesso, service worker)
{ "route": "/sem-acesso.html",  "allowedRoles": ["anonymous","authenticated"] }
{ "route": "/*",                "allowedRoles": ["anonymous","authenticated"] }

// 403 renderiza a página amigável de "sem acesso"
"responseOverrides": { "403": { "rewrite": "/sem-acesso.html", "statusCode": 403 } }
```

---

## 6. O que é fixo e o que se troca por sistema

Folha de parâmetros a preencher no início de cada projeto novo.

| Parâmetro | Status | Valor / origem |
|---|---|---|
| Tenant ID (`MS_TENANT_ID`) | **Fixo** | `4b30645b-0888-45c0-9481-712bde435ffd` — o mesmo em todos |
| App Registration + Client ID | Trocar | Criar uma nova por sistema — `<SISTEMA> SWA` |
| Client Secret | Trocar | Novo, validade 2 anos — só no App Settings + `.env` |
| Grupos de segurança + GUIDs | Trocar | Novos grupos `PRONEP-<SISTEMA>-*` e seus Object IDs |
| Nome e URL do SWA | Trocar | `<sistema>-pronep.azurestaticapps.net` |
| Redirect URI | Trocar | `https://<url>/.auth/login/aad/callback` |
| Site SharePoint + permissão | Trocar | Site do sistema + `Sites.Selected` concedido a ele |
| Mapa `GROUP_TO_ROLE` | Trocar | Reescrever GUID → papel para os novos grupos |
| Estrutura do código (`shared/*`) | **Reusar** | `auth.js`, `authz.js`, `graph.js`, `userRoles.js` |

---

## 7. Checklist de implantação

A ordem importa: o SWA precisa existir antes do redirect URI, e o site precisa ser
conhecido antes de conceder `Sites.Selected`.

- [ ] **Criar a App Registration** dedicada no tenant Pronep, contas só deste diretório. Ativar claim de grupos (Token configuration → Group ID em ID, Access e SAML) e gerar o client secret de 2 anos.
- [ ] **Criar os grupos de segurança** seguindo `PRONEP-<SISTEMA>-<Papel>`. Copiar o Object ID de cada grupo na hora e adicionar os membros iniciais.
- [ ] **Provisionar o Static Web App** (plano **Standard**, East US 2, origem "Outro"). Pegar o deployment token e guardar só no `.env` local. Marcar **"Tokens de ID"** na App Registration (Autenticação → Concessão implícita).
- [ ] **Configurar o redirect URI** na App Registration, agora que a URL do SWA existe: `https://<url>/.auth/login/aad/callback`.
- [ ] **Preencher os App Settings** do SWA: `AAD_CLIENT_ID`, `AAD_CLIENT_SECRET`, `AAD_TENANT_ID`, `SHAREPOINT_*`. Nunca versionar esses valores.
- [ ] **Conceder as permissões de Graph**: `Sites.Selected` + `Mail.Send`, com consentimento de admin. Depois, autorizar o site específico via Graph Explorer ou `Grant-PnPAzureADAppSitePermission`.
- [ ] **Colar os GUIDs** no `GROUP_TO_ROLE` e ajustar as flags de papel. Manter o mapa sincronizado entre `userRoles.js` e `MeusGrupos`.
- [ ] **Fazer o primeiro deploy** via SWA CLI e testar o fluxo de login ponta a ponta.
- [ ] **(Opcional) Ligar o Teams SSO**: Expose an API, scope `access_as_user`, 3 clientes Teams, manifest da app.

---

## 8. Decisões de segurança e armadilhas

O que custou tempo para descobrir no sistema de NF — registrado aqui para não se repetir.

- **[CRÍTICO] Segredos nunca no git — nem nos guias.** Client secret e deployment token vivem só no App Settings do SWA e no `.env` (no `.gitignore`). Documentação e `.md` apontam para onde o valor está, jamais o valor em si.
- **Escopo mínimo de Graph: `Sites.Selected`.** Preferir a `Sites.ReadWrite.All`. Custa um passo extra (autorizar o site), mas limita o dano de um secret vazado a um único site em vez de todo o SharePoint do tenant.
- **`/api/*` fechado por padrão.** Sem a regra `authenticated` em `/api/*`, um atacante pode forjar o header `x-ms-client-principal` e se passar por qualquer usuário. Exceções só nomeadas e autenticadas por secret próprio.
- **SWA sobrescreve o `Authorization: Bearer`.** Bug não documentado: o Easy Auth troca o header `Authorization` pelo token interno do runtime. No SSO do Teams, mande o token MSAL no header custom `X-Teams-Token` — o SWA não toca em headers não-padrão.
- **`principal.userId` não é o `oid` do Entra.** O `userId` do Easy Auth é um ID interno do SWA. O `oid` real, usado para consultar o Graph, vem do claim `objectidentifier`. Confundir os dois faz a resolução de papéis devolver vazio.
- **[CORRIGIDO na implantação do Sobreaviso, 07/2026] Plano Free NÃO serve para este padrão.** A autenticação customizada (bloco `auth` com App Registration própria — claim de grupos, tenant único, secret) **só existe no Standard**; o deploy no Free falha com *"The 'auth' configuration ... is only supported on the Standard SKU"*. Free só oferece provedores genéricos, sem grupos e aceitando conta de qualquer tenant. Junto com o Standard, marcar **"Tokens de ID"** na App Registration (Autenticação) — sem isso o callback devolve `401`.
- **O hostname SharePoint do tenant é `pronepadmin.sharepoint.com`** (não `pronep.sharepoint.com`). Usar outro valor gera `Invalid hostname for this tenancy` no Graph. Vale para `SHAREPOINT_SITE_HOSTNAME` de todos os sistemas.
- **`Sites.Selected` com papel `write` não cria listas.** Se a API do sistema cria as próprias listas (padrão Sobreaviso: `/api/Setup`), conceda o site com papel **`manage`**. Dica: se o `PATCH` da permissão devolver 404, faça um novo `POST` com o papel desejado — o Graph atualiza a concessão existente do mesmo app.
- **SWA CLI não roda em Mac Apple Silicon** (binário de deploy é x64 → erro `spawn -86`). Alternativa adotada: deploy via GitHub Actions (`Azure/static-web-apps-deploy@v1` com o deployment token em secret), que também dá deploy repetível por push.

---

*Padrão extraído do Sistema de Aprovação de NF · Pronep Life Care · base de referência para sistemas futuros.*
