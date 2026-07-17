# Handoff — Automação NF / n8n / WhatsApp (estado consolidado)

> **Fonte única da verdade, versionada no repo.** Substitui os handoffs soltos
> (sessão 1 e o `..._sessao2_...` que ficou no Downloads). Sempre atualize ESTE
> arquivo — assim os dois notebooks veem o mesmo estado via git.
> Arquitetura detalhada: [`automacao-emails-nf.md`](automacao-emails-nf.md).
> Última atualização: 2026-07-17 (sessão 3, no Mac).

---

## 0. Reconciliação de estado (importante — corrige o handoff da sessão 2)

O handoff da sessão 2 afirmava que `api/VarrerEmailsNF/` **não existia** e que as
Fases 1/2a/2b/2d **não tinham sido construídas**. **Isso estava incorreto** — quase
certo por ter olhado um clone desatualizado. Verificado no `main` (repo
`rafaelmmachado30/aprovacao-nf-pronep`, m duplo):

- ✅ `api/VarrerEmailsNF/` **existe no `main`** com histórico real (Fases 2b/2d + calibração).
- ✅ **Não há nada a reconstruir** — a "seção 4: construir VarrerEmailsNF" do doc antigo é desnecessária.
- ✅ A coluna `TelefoneNotificacao` foi além do doc: **PR #28 mergeado** e **deploy concluído** (a migração já está no ar).
- ✅ Todos os PRs #23–#28 mergeados.

---

## 1. Fluxo ponta a ponta

```
Caixa de e-mail da diretoria
      │  (Graph Mail.Read — permissão já liberada)
      ▼
api/VarrerEmailsNF  ── classifica NF x ruído ── baixa PDFs ── salva no SharePoint
      │                          "Novas NFs - Automacao/{Unidade}/Diretoria {Diretoria}/"
      │  POST no N8N_WEBHOOK_NF (Fase 2d, best-effort)
      ▼
   n8n (Webhook → Montar msg → Chatwoot) ──► WhatsApp do gestor da diretoria
```

---

## 2. Pronto (lado do sistema) — em produção no `main`

- **Fase 1** — `api/VarrerEmailsNF` lê e-mails, classifica (NF vs cobrança/boleto/interno),
  baixa os PDFs e grava no SharePoint. Classificador calibrado com dry-run real.
- **Fase 2a** — campo **e-mail** no cadastro de fornecedor (front).
- **Fase 2b** — corroboração pelo **Fechamento do Mês**.
- **Fase 2d (lado sistema)** — POST no webhook do n8n quando há candidatos (se webhook vazio, pula e reporta).
- **Coluna `TelefoneNotificacao`** — endpoint `api/MigrarColunaTelefoneNotificacao`
  (admin, idempotente) **no ar**. Cria coluna Texto na lista `PRONEP-NF-Diretorias`.

**Payload EXATO enviado ao n8n** (de `api/VarrerEmailsNF/index.js`; contrato travado):

```json
{
  "gestor": "fulano@pronep.com.br",
  "diretoria": "Comercial",
  "unidade": "SP",
  "telefone": "+5511999998888",
  "pasta": "Novas NFs - Automacao/SP/Diretoria Comercial",
  "total": 2,
  "candidatos": [
    { "fornecedor": "ACME LTDA", "assunto": "NF 12345 - ACME", "confianca": "alta", "esperado": "ACME LTDA", "arquivos": ["nf-12345.pdf"] }
  ]
}
```

- `telefone` vem em **E.164** (com `+`). O n8n deriva `numero` (só dígitos) disso.
- **Não** há `numero` nem `semTelefone` no corpo — se a diretoria não tiver telefone,
  vem `telefone: ""` e o n8n barra pelo IF `numero is not empty`.

---

## 3. O que FALTA (checklist real)

1. [ ] **Rodar 1× `GET /api/MigrarColunaTelefoneNotificacao`** (logado admin) → depois
       preencher os telefones **E.164** (`+5511999998888`) na lista `PRONEP-NF-Diretorias`,
       por Unidade+Diretoria. (Endpoint já no ar.)
2. [ ] **n8n**: importar o `PronepNF_Notificacao.json` limpo e **validar o envio Chatwoot** ponta a ponta (seção 5).
3. [ ] Copiar a **Production URL** do Webhook n8n → setar `N8N_WEBHOOK_NF` nas App Settings (seção 6).
4. [ ] Fase 3: cron chamando `VarrerEmailsNF` p/ todas as diretorias + `AUTOMACAO_EMAILS_SECRET`.

> ⚠️ **O arquivo `PronepNF_Notificacao.json` não está no Mac** (`~/Downloads/30_n8n_Hostinger/`
> não existe aqui). Trazer do Windows **ou** re-exportar da instância n8n.

---

## 4. n8n — workflow (montado na sessão 2)

- **Fluxo**: `Webhook NF → Montar mensagem → Config → Responder 200 → Tem telefone? (IF numero not empty) → Criar contato → Criar conversa + msg`
- **Envio: Chatwoot proativo** (NÃO Evolution direto). O node reativo do Byoterapia não serve (depende de conversa existente).
- **Montagem da mensagem: OK** (texto formatado + `numero` só dígitos, sem `+`). **Falta validar o envio Chatwoot.**

**Chatwoot desta instância:**
- URL: `http://82.25.79.134:3000` · `account_id = 1` · `inbox_id = 1` (inbox WhatsApp)
- Credencial n8n: **"ChatWoot account"** (id `XotA5fFCZQjC9Iki`)

**Bodies dos nodes Chatwoot** (backup):

- **Criar contato** — `POST {{url}}/api/v1/accounts/{{account_id}}/contacts`
  ```
  ={{ { "inbox_id": $json.inbox_id, "name": "Diretoria " + $json.diretoria, "phone_number": "+" + $json.numero } }}
  ```
- **Criar conversa + msg** — `POST {{url}}/api/v1/accounts/{{account_id}}/conversations`
  ```
  ={{ {
    "source_id": $('Criar contato').item.json.payload.contact_inboxes[0].source_id,
    "inbox_id": $('Config').item.json.inbox_id,
    "contact_id": $('Criar contato').item.json.payload.contact.id,
    "message": { "content": $('Montar mensagem').item.json.texto }
  } }}
  ```
  (`message.content` no `POST /conversations` cria a conversa **e** manda a 1ª msg.)

**Gotchas:**
- **Contato repetido** → 2ª notificação pro mesmo número: "Criar contato" retorna **422 (já existe)**.
  Fallback a fazer: `GET /api/v1/accounts/1/contacts/search?q=+{numero}` e reusar.
- Confirmar que o inbox aceita **mensagem livre proativa** (gateway não-oficial não tem janela 24h/template).

---

## 5. Validar o envio Chatwoot (passo a passo)

**Webhook** — node Webhook NF → "Listen for test event". Test URL:
`https://n8n-wrxo.srv1526230.hstgr.cloud/webhook-test/pronep-nf`

Disparar (trocar `telefone` pelo seu número):

```bash
curl -X POST "https://n8n-wrxo.srv1526230.hstgr.cloud/webhook-test/pronep-nf" \
  -H "Content-Type: application/json" \
  -d '{
    "gestor": "rafael.machado@pronep.com.br",
    "diretoria": "Comercial",
    "unidade": "SP",
    "telefone": "+5511999998888",
    "pasta": "Novas NFs - Automacao/SP/Diretoria Comercial",
    "total": 2,
    "candidatos": [
      {"fornecedor":"ACME LTDA","assunto":"NF 12345 - ACME","confianca":"alta","esperado":"ACME LTDA","arquivos":["nf-12345.pdf"]},
      {"fornecedor":"BETA SA","assunto":"NF 67890 - BETA","confianca":"media","esperado":"BETA SA","arquivos":["nf-67890.pdf"]}
    ]
  }'
```

**Ordem de validação:**
1. Rode **só o node "Criar contato"** e olhe o OUTPUT. Confirme os caminhos
   `payload.contact_inboxes[0].source_id` e `payload.contact.id` (o wrapper **varia por versão do Chatwoot**).
2. Se o contato criar OK, o "Criar conversa + msg" dispara e a mensagem cai no WhatsApp.

---

## 6. App Settings pendentes (Azure SWA → Configuration)

| App Setting | Para quê | Status |
|---|---|---|
| `N8N_WEBHOOK_NF` | Production URL do Webhook n8n (destino do POST da Fase 2d) | ⬜ setar após validar envio |
| `AUTOMACAO_EMAILS_SECRET` | Segredo do header `X-Automacao-Secret` (cron/Fase 3) | ⬜ definir |
| `EMAIL_DOMINIO_INTERNO` | Domínio interno a ignorar (default `pronep.com.br`) | opcional |

---

## 7. Como testar o VarrerEmailsNF (referência)

Logado admin (ou header `X-Automacao-Secret`):

```
# DRY-RUN — só classifica e reporta, não baixa:
/api/VarrerEmailsNF?gestor=<email-da-caixa>&dias=3&dryRun=1
# Real — baixa PDFs e (se N8N_WEBHOOK_NF setado) dispara o webhook:
/api/VarrerEmailsNF?gestor=<email-da-caixa>&dias=3
```

Params: `gestor` (obrigatório), `dias` (1–30, def 3), `limite` (1–100, def 50),
`dryRun` (`1`/`true`), `unidade`, `diretoria`. Idempotência via **ledger** no SharePoint.

Schema `PRONEP-NF-Diretorias`: chaveada por `Title = "Unidade|Diretoria"` (ex.: `SP|Suprimentos`);
colunas internas renomeadas `field_1..field_5` (Unidade, Diretoria, Email, Nome, GrupoEntraId).
A `TelefoneNotificacao` (criada via Graph) mantém nome interno próprio (não vira `field_6`).

---

## 8. Decisões já tomadas (não reabrir)

- **Estrutura de pastas** mantida (Aprovadas por data; Pendentes/Rejeitadas por `{Unidade}/Diretoria {Diretoria}`). Limpeza do legado **adiada**. "Garantir daqui pra frente" = feito (#26: aprovação/rejeição por identidade exata do PDF).
- Pasta da automação: `Novas NFs - Automacao/{Unidade}/Diretoria {Diretoria}/`.
- WhatsApp via **Chatwoot proativo** em instância dedicada; gateway **não-oficial**.
- `Mail.Read` (Graph) liberada e escopada.

---

**Produção:** `https://purple-forest-09588fe10.7.azurestaticapps.net`
**Repo:** `github.com/rafaelmmachado30/aprovacao-nf-pronep`
**n8n:** `https://n8n-wrxo.srv1526230.hstgr.cloud`
