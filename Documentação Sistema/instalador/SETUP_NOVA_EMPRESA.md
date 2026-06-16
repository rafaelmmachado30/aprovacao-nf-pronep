# Sistema de Aprovação de NF — Setup para Nova Empresa

Este código foi originalmente desenvolvido para a Pronep Life Care e contém referências hardcoded à infraestrutura dela. Este guia lista TUDO que precisa ser configurado/substituído para rodar em outra empresa.

---

## 1. Pré-requisitos

| Item | Necessário para |
|---|---|
| Tenant Microsoft 365 / Entra ID | Auth + grupos + Graph API |
| Site SharePoint Online dedicado | Armazenamento de listas + PDFs |
| Azure Subscription | Static Web App (hospedagem) |
| Conta Anthropic Claude API | SAN (IA conversacional) + extração de vigências de contratos |
| Conta OpenAI (opcional) | Fallback se Anthropic falhar |
| Git + Node.js 22 + Azure Functions Core Tools | Build local e deploy |

---

## 2. Estrutura SharePoint

Criar um site SharePoint dedicado (ex: `https://<empresa>.sharepoint.com/sites/Aprovacao-NF`) com as seguintes **listas**:

### 2.1 Listas principais (criar manualmente OU deixar o sistema criar)
- `PRONEP-NF-NotasFiscais` — fila de NFs (criar manual; ver schema em `PRONEP-NF-NotasFiscais.xlsx`)
- `PRONEP-NF-Fornecedores` — base de fornecedores (criar manual; ver `PRONEP-NF-Fornecedores.xlsx`)
- `PRONEP-NF-Diretorias` — mapeamento de gestores aprovadores (criar manual; ver `PRONEP-NF-Diretorias.xlsx` e `.csv`)
- `PRONEP-NF-Config` — config geral
- `PRONEP-NF-PushSubscriptions` — push notifications
- `PRONEP-NF-AuditLog` — auditoria
- `PRONEP-NF-SOL-Conversas` — histórico da SAN

### 2.2 Listas criadas automaticamente pelos endpoints de migração
Após primeiro deploy, chame estes endpoints uma vez (logado como admin):
- `GET /api/CriarListaContratos` — cria `PRONEP-NF-Contratos` + 17 colunas
- `GET /api/MigrarColunaAlinhamentoFinanceiro` — adiciona 2 colunas em NotasFiscais
- `GET /api/MigrarColunaNFAuto` — adiciona coluna NumeroAutogerado em NotasFiscais
- `GET /api/MigrarColunasURL` — adiciona UrlPDFStr / UrlPDFAprovadoStr

### 2.3 Estrutura de pastas no SharePoint Drive
Criar manualmente:
```
Notas Fiscais/
├── Pendentes/
│   ├── SP/
│   │   ├── Diretoria Comercial/
│   │   ├── Diretoria Tecnologia/
│   │   └── ... (uma por diretoria configurada)
│   ├── RJ/
│   └── ES/
├── Rejeitadas/    (subpastas Unidade/Diretoria criadas automaticamente)
└── Notas Aprovadas/
    ├── SP/        (subpastas por DATA criadas automaticamente)
    ├── RJ/
    └── ES/
```

Adapte unidades e diretorias para a nova empresa.

### 2.4 Site SharePoint de Contratos (separado, opcional)
Se a empresa também usar o módulo de Contratos, criar um site separado contendo a estrutura:
```
/CONTRATOS/CONTRATOS E DOCUMENTOS - PRESTADORES/
├── <Diretoria 1>/
│   ├── <Unidade>/
│   │   └── <Fornecedor>/
│   │       └── *.pdf
```

---

## 3. Configuração Entra ID (AAD)

### 3.1 Criar App Registration
Necessário para o backend acessar Graph API:
1. Entra ID → App registrations → New registration
2. Name: `<empresa>-NF-Backend`
3. Supported account types: Single tenant
4. Adicionar API permissions (application, não delegated):
   - `Group.Read.All`
   - `User.Read.All`
   - `Sites.ReadWrite.All`
   - `Sites.Manage.All` (pra criação de listas via Graph)
   - `Mail.Send` (pra emails)
   - `TeamsAppInstallation.ReadWriteAndConsentForUser.All` + `TeamsActivity.Send` (pra notificações Teams)
5. Grant admin consent
6. Criar Client Secret e guardar (vai no .env)

### 3.2 Criar Grupos de Segurança
Criar os seguintes grupos no Entra ID e mapear os OIDs em `api/shared/userRoles.js` constante `GROUP_TO_ROLE`:

| Grupo | Role | Função |
|---|---|---|
| `<EMP>-NF-Submitter` | submitter | Pode lançar NFs |
| `<EMP>-NF-Admin` | administrador | Acesso total + auditoria |
| `<EMP>-NF-TI` | ti | Vê tudo do TI |
| `<EMP>-Financeiro-Gestao` | financeiro_nf | Aprovação financeira + integração Omie |
| `<EMP>-NF-Gestor-Suprimentos` | gestor_suprimentos | Aprovação Suprimentos |
| `<EMP>-NF-Gestor-Financeira` | gestor_financeira | Aprovação dir. Financeira |
| `<EMP>-NF-Gestor-Tecnologia` | gestor_tecnologia | Aprovação TI |
| `<EMP>-NF-Gestor-Qualidade` | gestor_qualidade | Aprovação Qualidade |
| `<EMP>-NF-Gestor-RH-DP` | gestor_rh_dp | Aprovação RH |
| `<EMP>-NF-Gestor-Fiscal-Contabil` | gestor_fiscal_contabil | Aprovação Fiscal/Contábil |
| `<EMP>-NF-Gestor-Juridica` | gestor_juridica | Aprovação Jurídica + acesso total a Contratos |
| `<EMP>-NF-Gestor-Administrativa` | gestor_administrativa | Aprovação Administrativa |
| `<EMP>-NF-Gestor-Tecnica-SP` | gestor_tecnica_sp | Aprovação Técnica SP |
| `<EMP>-NF-Gestor-Tecnica-RJES` | gestor_tecnica_rjes | Aprovação Técnica RJ/ES |

Cada grupo recebe um OID (GUID). Pegar os OIDs e substituir em `api/shared/userRoles.js` linha 18-33.

Replicar a mesma substituição em `api/MeusGrupos/index.js`.

---

## 4. Azure Static Web App

1. Azure Portal → Create resource → Static Web App
2. Plan: **Standard** (necessário para Managed Functions com 4GB+ memória)
3. Conectar com seu repo Git
4. Build presets:
   - App location: `/`
   - Output location: `wwwroot`
   - API location: `api`

### 4.1 Easy Auth (Entra ID)
1. SWA → Authentication → adicionar Microsoft (Entra ID)
2. Configurar com o tenant da empresa
3. Vai criar uma App Registration separada (login do usuário). Anotar o Client ID + Tenant ID.

### 4.2 App Settings (env vars no Azure)
Configurar em SWA → Configuration:

| Variável | Descrição |
|---|---|
| `AAD_TENANT_ID` | Tenant ID do Entra ID |
| `AAD_CLIENT_ID` | Client ID da App Registration (3.1) |
| `AAD_CLIENT_SECRET` | Client Secret (3.1) |
| `SHAREPOINT_SITE_HOSTNAME` | ex: `empresa.sharepoint.com` |
| `SHAREPOINT_SITE_PATH` | ex: `/sites/Aprovacao-NF` |
| `EMAIL_FROM_ADDRESS` | Email institucional remetente |
| `ANTHROPIC_API_KEY` | sk-ant-... |
| `ANTHROPIC_MODEL_HAIKU` | `claude-haiku-4-5-20251001` |
| `ANTHROPIC_MODEL_SONNET` | `claude-sonnet-4-6` |
| `OPENAI_API_KEY` | (opcional, fallback) |
| `LINK_APROVACAO_SECRET` | string aleatória 64 chars |
| `VAPID_PUBLIC_KEY` | pra push notifications (gerar via web-push) |
| `VAPID_PRIVATE_KEY` | pra push notifications |
| `VAPID_SUBJECT` | `mailto:admin@empresa.com.br` |
| `ALERTA_DIARIO_SECRET` | string aleatória pro cron diário de NFs |
| `ALERTA_CONTRATOS_SECRET` | string aleatória pro cron diário de contratos |
| `SOL_ADMIN_EMAILS` | lista de emails admin (csv) |
| `OMIE_APP_KEY_SP` / `OMIE_APP_SECRET_SP` | (se usar integração Omie SP) |
| `OMIE_APP_KEY_RJ` / `OMIE_APP_SECRET_RJ` | (se usar integração Omie RJ) |
| `OMIE_APP_KEY_ES` / `OMIE_APP_SECRET_ES` | (se usar integração Omie ES) |

---

## 5. Hardcodes a substituir no código

Grep antes de subir:

| Arquivo | O que substituir |
|---|---|
| `wwwroot/index.html` | URLs, nome "Pronep", logo, cores |
| `api/AlertaDiario/index.js` | `SISTEMA_URL` |
| `api/AlertaContratosDiario/index.js` | `SISTEMA_URL`, `GRUPO_JURIDICO_OID` (linha ~37) |
| `api/shared/email.js` | `DEFAULT_FROM`, base URL no `gerarLinks` |
| `api/shared/userRoles.js` | TODO o objeto `GROUP_TO_ROLE` (linha 18-33) |
| `api/MeusGrupos/index.js` | Mesmo mapa de OIDs |
| `api/shared/contratos.js` | `ROOT_FOLDER_PATH`, `UNIDADES_VALIDAS`, `MAPA_DIRETORIA` |
| `wwwroot/contratos-bfs.js` | Lista de diretorias |
| `.github/workflows/*.yml` | URL do Azure SWA |

Recomendação: roda um `grep -ri "purple-forest-09588fe10" .` e `grep -ri "pronep" .` no projeto pra mapear todas as referências.

---

## 6. Templates de listas SharePoint

Os arquivos .xlsx no zip são **templates do schema** das listas principais:
- `PRONEP-NF-NotasFiscais.xlsx` — colunas: Title, NumeroNF, CNPJFornecedor, Valor, DataVencimento, Unidade, Diretoria, AprovadorAtual, Status, LancadoPor, etc.
- `PRONEP-NF-Fornecedores.xlsx` — base de fornecedores
- `PRONEP-NF-Diretorias.xlsx` + `.csv` — mapeamento de gestores

Criar as listas no SP com essas colunas (Configurações da Lista → Adicionar coluna), respeitando tipos.

---

## 7. GitHub Actions (cron jobs)

Em `.github/workflows/`:
- `azure-static-web-apps-*.yml` — CI/CD do deploy (configurado automaticamente pelo Azure)
- `alerta-diario.yml` — Cron 9h, 17h, sexta — emails de NFs pendentes
- `alerta-contratos-diario.yml` — Cron 8h seg-sex — emails de contratos vencendo

Editar URLs nos workflows e cadastrar as 2 secrets no GitHub:
- `ALERTA_DIARIO_SECRET` (igual ao Azure App Setting)
- `ALERTA_CONTRATOS_SECRET` (igual ao Azure App Setting)

---

## 8. Teams (opcional)

Se quiser notificações Teams nativas (não só email):
- Ver `SETUP_TEAMS_SSO.md`
- Ver `SETUP_TEAMS_ACTIVITY.md`
- Pasta `teams-app/` contém o manifesto da App

---

## 9. Deploy

Após tudo configurado:

```bash
git remote add origin <novo-repo>
git push -u origin main
```

O Azure SWA detecta o push e deploya automaticamente (~2-3 min).

Após deploy verde, rodar os endpoints de migração da seção 2.2 uma vez.

---

## 10. Validação pós-setup

Checklist mínimo:
- [ ] Login funciona (Easy Auth Microsoft)
- [ ] `/api/Hello` retorna 200 (Function App ativo)
- [ ] `/api/MeusGrupos` retorna os grupos do user logado
- [ ] `/api/ListarFornecedores` retorna dados (mesmo que vazio)
- [ ] `/api/ListarNotas` retorna a lista
- [ ] Tela de Contratos renderiza
- [ ] FAB da SAN aparece e responde
- [ ] Botão "Lançar NF" abre o formulário
- [ ] Push notification subscreve (Settings do browser → permitir notificações)

---

## 11. Arquivos críticos pra entender o código

| Arquivo | Função |
|---|---|
| `wwwroot/index.html` | UI principal (single-page, ~7600 linhas) |
| `wwwroot/contratos-bfs.js` | Orquestrador de sync de contratos |
| `api/PostNota/index.js` | Lançamento de NF |
| `api/AprovarNota/index.js` | Aprovação + compliance financeiro |
| `api/RejeitarNota/index.js` | Rejeição + estorno |
| `api/SolChat/index.js` | Backend da SAN (Claude via tool use) |
| `api/shared/sol.js` | Lógica da SAN: tools, prompt, dispatch |
| `api/shared/email.js` | Engine de email + Teams |
| `api/shared/contratos.js` | Crawler SP + Claude pra extrair vigências |
| `api/SincronizarContratos/index.js` | Sync individual de pasta de contratos |
| `api/AlertaContratosDiario/index.js` | Cron alerta de contratos |

---

## 12. Suporte / Dúvidas

Esse código foi desenvolvido com auxílio de IA (Claude) ao longo de vários sprints. Pra suporte:
- Issues técnicos: ler o código (bem comentado em pt-br)
- Adaptação pra nova empresa: este documento + grep dos hardcodes
- Dúvidas sobre arquitetura: o JSDoc dos arquivos `api/shared/*.js` e `api/*/index.js` explica cada decisão

Boa sorte com a implementação!
