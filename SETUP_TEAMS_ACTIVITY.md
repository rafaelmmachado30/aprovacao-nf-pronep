# Setup Notificação Teams 1:1 via Graph API

**Objetivo:** substituir o webhook 1:1 do Power Automate (depreciated, dando 403) pelo caminho oficial Microsoft: `sendActivityNotification`.

Quando funcionar, o aprovador recebe uma notificação no sino do Teams (bell icon) toda vez que houver NF aguardando aprovação, com link direto pro sistema.

---

## 3 ações que VOCÊ precisa fazer

### Ação 1 — Adicionar permissão no App Registration

1. Abre o Azure Portal → **Microsoft Entra ID** → **App registrations**
2. Procura o App Reg que já usa (o do `AAD_CLIENT_ID`)
3. Lateral esquerda → **API permissions** → **Add a permission**
4. Escolhe **Microsoft Graph** → **Application permissions**
5. Procura `TeamsActivity.Send` → marca → **Add permissions**
6. De volta na tela de permissions, clica **Grant admin consent for Pronep** (botão azul no topo)
7. Confirma. A permissão deve ficar com check verde ✓ na coluna "Status"

### Ação 2 — Anotar o Client ID do App Reg

Na home page do App Registration, copia o valor de **Application (client) ID**.
Vai parecer algo tipo: `abc12345-6789-...-...-...`

Esse é o **AAD_CLIENT_ID** que vamos usar no manifest.

### Ação 3 — Subir a Teams App no tenant

O sistema tem um diretório novo: `teams-app/` com 3 arquivos:
- `manifest.json` — descreve a Teams App (precisa editar 1 linha)
- `color.png` — ícone colorido 192x192
- `outline.png` — ícone outline 32x32

**Passo 3a — editar o manifest:**

Abre `teams-app/manifest.json` num editor. Procura a linha:

```json
"id": "REPLACE_WITH_AAD_CLIENT_ID",
```

Dentro do bloco `"webApplicationInfo"`. Substitui pelo Client ID da Ação 2.

Resultado parecido com:
```json
"webApplicationInfo": {
  "id": "abc12345-6789-...-...-...",
  "resource": "https://purple-forest-09588fe10.7.azurestaticapps.net"
}
```

**Passo 3b — zipar a Teams App:**

No PowerShell, dentro da pasta `C:\Pronep\Aprovacao_NF\teams-app`:

```powershell
Compress-Archive -Path .\manifest.json,.\color.png,.\outline.png -DestinationPath aprovacao-nf-teams.zip -Force
```

O arquivo `aprovacao-nf-teams.zip` é o pacote que vai pro Teams Admin.

**Passo 3c — upload no Teams Admin Center:**

1. Abre https://admin.teams.microsoft.com
2. Lateral esquerda → **Teams apps** → **Manage apps**
3. Topo da página → **+ Upload new app** → **Upload**
4. Seleciona o `aprovacao-nf-teams.zip`
5. Espera processar (~30s). Deve aparecer "Aprovacao NF" na lista de apps com status **Allowed**

Se aparecer erro de validação, manda print que a gente corrige.

**Passo 3d — permitir a app pros aprovadores (uma vez):**

1. Ainda no Teams Admin Center → **Teams apps** → **Permission policies**
2. Edita a policy padrão (Global) ou cria uma específica pros aprovadores
3. Garante que "Aprovacao NF" está na lista de Allowed apps
4. Salva

> **Importante:** a Teams App **não precisa estar instalada** pelo aprovador pra `sendActivityNotification` funcionar. Basta estar **registrada e permitida** no tenant. O usuário recebe a notificação direto no sino do Teams.

---

## Como saber se funcionou

Depois de fazer as 3 ações e o deploy do código:

1. Lança uma NF de teste no sistema
2. Olha o response do `/api/PostNota` (F12 → Network)
3. Procura no JSON o campo `result.teamsAtividade`:

```json
{
  "teamsAtividade": {
    "ok": true,
    "sentTo": "sandra@pronep.com.br",
    "userId": "abc...",
    "activityType": "approvalRequired"
  }
}
```

Se `ok: true` → deu certo. O aprovador deve ver o sino vermelho no Teams desktop/mobile.

Se `ok: false` → manda o body do erro pra eu diagnosticar.

---

## Comportamento durante a transição

Enquanto você não fizer as 3 ações acima, o sistema **continua funcionando** assim:

- ✅ Email com botões Aprovar/Rejeitar continua sendo enviado (sempre foi o canal principal)
- ⚠️ `enviarTeamsAtividade` vai falhar com erro tipo "Teams App not found" ou similar (esperado, ainda não tem app registrada)
- ✅ Fallback: se `TEAMS_WEBHOOK_URL` estiver setado, posta no canal central
- ❌ Webhook 1:1 (`TEAMS_WEBHOOKS` por email) **NÃO** é mais usado, porque está dando 403 sistêmico

Depois das 3 ações: passa a usar `sendActivityNotification` (1:1 oficial) com sucesso.

---

## Trade-offs vs webhook antigo

| Aspecto | Webhook 1:1 (antigo) | sendActivityNotification (novo) |
|---|---|---|
| Funciona? | ❌ 403 desde Microsoft depreciation | ✅ Caminho oficial Microsoft |
| Setup | Cada aprovador instala workflow + cola URL | Admin sobe Teams App 1x |
| Manutenção | URL pode expirar/quebrar individualmente | Estável |
| UI no chat | Adaptive Card com botões inline | Notificação no sino + link |
| Card interativo? | Sim (Aprovar/Rejeitar inline) | Não nativo — mas email já cobre isso |

**Nota:** se no futuro quiser cards interativos no chat 1:1, o caminho é criar um Bot Framework (mais pesado). Hoje o email + notificação Teams cobrem o caso bem.
