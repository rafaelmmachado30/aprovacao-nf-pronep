# Setup SSO Teams - passo a passo

**Objetivo:** permitir que a Personal Tab "Aprovacao NF" funcione dentro do Teams sem o problema de iframe bloqueado.

Como funciona: dentro do Teams, o usuario ja esta autenticado no Microsoft 365. A gente pede pro Teams um token MSAL desse usuario via `getAuthToken()`, manda esse token pro nosso backend (`/api/AuthTeamsSSO`), o backend valida e estabelece a sessao. A partir dai todas as Functions funcionam normalmente dentro do iframe.

---

## Acao 1 - Expose an API no App Reg

1. Azure Portal -> **Microsoft Entra ID** -> **App registrations**
2. Procura o App Reg com Client ID `85f7fa68-a241-4caf-803f-9991bd1f0eee`
3. Lateral esquerda -> **Expose an API**
4. Clica em **Set** ao lado de "Application ID URI"
5. O Azure vai sugerir algo tipo `api://85f7fa68-a241-4caf-803f-9991bd1f0eee`. Substitui por:

   ```
   api://purple-forest-09588fe10.7.azurestaticapps.net/85f7fa68-a241-4caf-803f-9991bd1f0eee
   ```

6. **Save**

> **Importante:** copia esse Application ID URI completo (com a URL e o GUID no final). Voce vai precisar dele depois.

---

## Acao 2 - Adicionar o scope `access_as_user`

Ainda em **Expose an API**:

1. Clica em **Add a scope**
2. Preenche:
   - **Scope name:** `access_as_user`
   - **Who can consent:** Admins and users
   - **Admin consent display name:** `Acessar Sistema NF como o usuario`
   - **Admin consent description:** `Permite que o Sistema de Aprovacao de NF acesse dados do usuario via Teams SSO.`
   - **User consent display name:** `Acessar o Sistema de Aprovacao de NF`
   - **User consent description:** `Permite que o sistema acesse seus dados ao abrir a tab no Teams.`
   - **State:** Enabled
3. **Add scope**

A scope completa vai ficar tipo:
```
api://purple-forest-09588fe10.7.azurestaticapps.net/85f7fa68-a241-4caf-803f-9991bd1f0eee/access_as_user
```

---

## Acao 3 - Authorized client applications (Teams)

Ainda em **Expose an API**, role pra baixo ate **Authorized client applications**:

1. Clica em **Add a client application**
2. Adiciona o **primeiro** cliente:
   - **Client ID:** `1fec8e78-bce4-4aaf-ab1b-5451cc387264` (Microsoft Teams desktop / mobile)
   - **Authorized scopes:** marca o `api://.../access_as_user` que voce acabou de criar
   - **Add application**
3. **Add a client application** de novo
4. Adiciona o **segundo** cliente:
   - **Client ID:** `5e3ce6c0-2b1f-4285-8d4b-75ee78787346` (Microsoft Teams web)
   - **Authorized scopes:** marca o `access_as_user`
   - **Add application**
5. **Add a client application** de novo
6. Adiciona o **terceiro** cliente:
   - **Client ID:** `4765445b-32c6-49b0-83e6-1d93765276ca` (Microsoft Teams Office app web)
   - **Authorized scopes:** marca o `access_as_user`
   - **Add application**

> **Esses 3 GUIDs sao fixos da Microsoft** - representam os clientes Teams. Voce nao precisa inventar nada.

---

## Acao 4 - Atualizar manifest da Teams App

Vou atualizar o `teams-app/manifest.json` com:
- `version: "1.0.0"` -> `"1.0.1"` (forca o Teams a reconhecer atualizacao)
- `webApplicationInfo.resource` apontando pro Application ID URI

**Voce nao precisa editar nada manualmente** - eu ja vou commitar o manifest atualizado. Mas voce vai precisar:

1. Re-zipar o conteudo de `teams-app/`:
   ```powershell
   cd C:\Pronep\Aprovacao_NF\teams-app
   Compress-Archive -Path .\manifest.json,.\color.png,.\outline.png -DestinationPath aprovacao-nf-teams-v1.0.1.zip -Force
   ```

2. Teams Admin Center -> Manage apps -> procura "Aprovacao NF" -> clica nela
3. Aba **Manifest** ou **Detalhes** -> botao **Update** ou **Atualizar**
4. Faz upload do novo `aprovacao-nf-teams-v1.0.1.zip`

---

## Acao 5 - me retornar o Application ID URI

Cola aqui no chat o **Application ID URI** que ficou na Acao 1 (algo tipo `api://purple-forest-09588fe10.7.azurestaticapps.net/85f7fa68-...`).

Eu vou usar esse valor exato no codigo do frontend (na chamada `getAuthToken({ scopes: [...] })`).

Tambem preciso confirmar:
- [ ] Application ID URI definido
- [ ] Scope `access_as_user` criado e habilitado
- [ ] 3 client apps autorizados (Teams desktop, web, Office)

---

## IMPORTANTE - Limitacao do SWA com Authorization header

**Bug nao documentado da Microsoft:** o Easy Auth do Azure Static Web Apps SUBSTITUI o header `Authorization: Bearer` que o frontend envia, trocando pelo proprio token interno do Azure Functions runtime (aud=azurewebsites.net/azurefunctions).

**Sintoma:** o backend recebe um token diferente do que o frontend mandou. O `validateTeamsToken` falha porque o aud/iss nao bate.

**Workaround usado neste projeto:** usar header custom `X-Teams-Token` em vez de `Authorization: Bearer`. SWA nao interfere em headers nao-padrao.

```js
// Frontend (interceptor fetch):
init.headers.set('X-Teams-Token', window._teamsAuthToken);

// Backend (shared/auth.js):
const token = req.headers['x-teams-token'];
```

---

## IMPORTANTE - Configuracao adicional pra Standard SKU

Se o SWA esta no plano **Standard** (custom authentication ativada via bloco `auth` no `staticwebapp.config.json`), o App Reg precisa:

1. App Registration -> **Autenticacao** -> bloco "Concessao implicita e fluxos hibridos"
2. Marca **"Tokens de ID (usados para fluxos implicitos e hibridos)"**
3. Deixa "Tokens de acesso" DESMARCADO (nao precisa, mais seguro)
4. Salva

> **Por que:** o Easy Auth do SWA Standard exige ID token no flow de autenticacao OAuth. O SKU Free ignora essa exigencia, o Standard valida e da 401 no callback se estiver desmarcado. **Sintoma:** `https://<seu-app>.azurestaticapps.net/.auth/login/aad/callback` retorna `401: Unauthorized`.

---

## Como funcionara depois

1. Voce lanca NF -> notif Teams chega (ja funciona)
2. Voce clica na notif
3. Personal Tab abre dentro do Teams
4. Frontend detecta Teams + chama `getAuthToken()` -> recebe token MSAL
5. Frontend manda token pra `/api/AuthTeamsSSO` -> backend valida + seta cookie
6. Frontend recarrega contexto -> agora autenticado
7. Frontend abre direto o modal da NF que veio no `subEntityId`
8. Voce ve dados + botoes Aprovar/Rejeitar
9. Aprova direto, dentro do Teams - missao cumprida

Fora do Teams (no browser), nada muda - continua usando Easy Auth normal.
