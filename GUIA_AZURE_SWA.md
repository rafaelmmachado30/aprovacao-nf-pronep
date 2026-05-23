# Guia — Provisionar Azure Static Web App e fazer o primeiro deploy

Pré-requisito: ter completado `GUIA_ENTRA_ID.md` (App Registration + grupos criados).

## Arquitetura final

```
[Colaborador Pronep]
       ↓
[https://aprovacao-nf-pronep.azurestaticapps.net]
       ↓ (Easy Auth — bloqueia quem não tem grupo PRONEP-NF-*)
[Entra ID] → valida login + retorna grupos no token
       ↓
[/api/GetRoles] → mapeia grupos para roles
       ↓
[Frontend HTML (wwwroot/index.html) + APIs (api/*)]
       ↓
[SharePoint Online via Microsoft Graph] (Notas + Listas)
```

## Etapa 1 — Criar Static Web App (5 min)

1. https://portal.azure.com → busca por **Aplicativos Web Estáticos** → **+ Criar**
2. Preenche:
   - **Assinatura**: a da Pronep (mesma do Analytics)
   - **Grupo de recursos**: criar novo `rg-pronep-aprovacao-nf` (ou reaproveitar `rg-pronep-analytics`)
   - **Nome**: `aprovacao-nf-pronep`
   - **Tipo de plano**: **Gratuito** (Free)
   - **Região do plano**: **East US 2** (Free tier não tem Brasil — sem impacto, é só para os Functions; o conteúdo estático é CDN global)
   - **Origem de implantação**: **Outro** (não usa GitHub — vamos usar SWA CLI manual)
3. **Revisar + criar** → **Criar**. Aguarda ~2 minutos.

## Etapa 2 — Pegar deployment token (1 min)

1. No painel do SWA → **Visão geral** → topo → **Gerenciar token de implantação**
2. **Copia** o token (string longa). Guarda apenas no `.env` local — NÃO cole o valor aqui no arquivo .md.

## Etapa 3 — Configurar redirect URI na App Registration (2 min)

Agora que o SWA existe, sabemos a URL. Volta na App Registration `Pronep Aprovacao NF SWA`:

1. **Autenticação** → **+ Adicionar plataforma** → **Web**
2. Preenche:
   - **Redirect URI**: `https://aprovacao-nf-pronep.azurestaticapps.net/.auth/login/aad/callback`
3. **Configurar**

## Etapa 4 — App Settings do SWA (3 min)

No SWA no portal Azure:

1. Painel esquerdo → **Configuração** (ou "Variáveis de ambiente" em UIs novas)
2. **+ Adicionar** duas configurações:
   - `AAD_CLIENT_ID` = (Application ID da App Reg — ver Azure Portal → App Registrations → Pronep Aprovacao NF SWA → Overview)
   - `AAD_CLIENT_SECRET` = (Valor do segredo criado — guardado no `.env` local, NÃO neste arquivo)
3. **Salvar**

> ⚠️ **Segurança**: nunca cole o valor real do client secret aqui. O arquivo `.env` (que está no `.gitignore`) é o único lugar seguro pra esse valor no projeto local. Em produção, os valores ficam apenas no App Settings do SWA.

Depois das implementações reais, vão entrar mais:
- `GRAPH_TENANT_ID`, `GRAPH_CLIENT_ID`, `GRAPH_CLIENT_SECRET` (acesso ao SharePoint via Graph)
- `SHAREPOINT_SITE_ID`
- `NOTIF_MAIL_FROM`
- `TEAMS_WEBHOOK_URL`

(Por enquanto, só `AAD_CLIENT_ID` e `AAD_CLIENT_SECRET` são suficientes para o primeiro deploy.)

## Etapa 5 — Permissões Graph na App Registration (5 min)

A mesma App Reg vai precisar acessar o SharePoint. No `Pronep Aprovacao NF SWA`:

1. **Permissões de API** → **+ Adicionar permissão** → **Microsoft Graph** → **Permissões de aplicativo**
2. Adicionar:
   - `Sites.Selected` (vamos restringir a sites específicos)
   - `Mail.Send` (envio de e-mail transacional)
3. **Conceder consentimento de administrador**

Em seguida, autorizar o acesso ao site específico (Sites.Selected exige isso):

4. No SharePoint Admin Center, ou via Graph Explorer rodando:
   ```
   POST https://graph.microsoft.com/v1.0/sites/{site-id}/permissions
   Body:
   {
     "roles": ["write"],
     "grantedToIdentities": [{
       "application": { "id": "<CLIENT_ID_DA_APP_REG>", "displayName": "Pronep Aprovacao NF SWA" }
     }]
   }
   ```

(Mais detalhes em `GUIA_SHAREPOINT.md`.)

## Etapa 6 — Instalar SWA CLI (já feito pelo instalar.bat)

Se ainda não rodou, execute na raiz do projeto:

```
.\instalar.bat
```

## Etapa 7 — Primeiro deploy (1 min)

No PowerShell, na pasta do projeto:

```powershell
# Setar o token (substitua pelo seu — pegue em Azure Portal → SWA → Gerenciar token de implantação)
$env:DEPLOYMENT_TOKEN = "SEU_DEPLOYMENT_TOKEN_AQUI"

# Rodar o deploy
.\04_deploy_azure.bat
```

Vai subir `wwwroot/` + `api/`.

## Etapa 8 — Testar

Abre no browser: `https://aprovacao-nf-pronep.azurestaticapps.net`

Sequência esperada:
1. Redireciona pra tela de login Microsoft
2. Você loga com `@pronep.com.br`
3. Volta pra URL e o protótipo HTML renderiza

Se você está no grupo `PRONEP-NF-Admin`, todos os menus aparecem. Se está só no `PRONEP-NF-Submitter`, vê só Nova NF / Minhas NFs / etc.

## Atualizando depois

Toda vez que mexer no `wwwroot/index.html` ou em qualquer função do `api/`, basta:

```powershell
$env:DEPLOYMENT_TOKEN = "..."   # só uma vez por sessão
.\04_deploy_azure.bat
```

Em ~30s o conteúdo atualizado está no ar.

## Diagnóstico de erros

| Erro | Solução |
|---|---|
| `AADSTS50011` na hora do login | Redirect URI da App Reg não confere — confere se está `https://aprovacao-nf-pronep.azurestaticapps.net/.auth/login/aad/callback` |
| `404 sign-in failed` | `AAD_CLIENT_SECRET` errado ou expirado no Config do SWA |
| `403 sem-acesso.html` | Usuário não pertence a nenhum grupo PRONEP-NF-* |
| `swa: comando não reconhecido` | Falta o `npm install -g @azure/static-web-apps-cli` (rodar `instalar.bat`) |
| Deploy reclama de token | Variável `DEPLOYMENT_TOKEN` não setada — voltar à Etapa 2 |
| Função retorna 501 | Esperado nesta fase — apenas `GetRoles` está implementada. As demais ainda são esqueletos. |

## Custos

Plano Free do Azure Static Web Apps cobre:
- 100 GB de banda/mês
- 0.5 GB de storage
- Functions integradas (até ~1M invocações/mês)
- Custom domain incluso
- SSL automático

Volume previsto (~1.000 NF/mês × ~10 requests/NF = 10k requests/mês) está confortável no free tier.
