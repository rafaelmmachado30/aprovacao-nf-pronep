# Handoff — Continuação da automação (n8n / WhatsApp)

> Documento de retomada. Foco: **o que falta**. Para arquitetura e detalhes já
> implementados, ver [`automacao-emails-nf.md`](automacao-emails-nf.md).
> Última atualização: 2026-07-17.

---

## 1. Onde paramos (resumo de 30s)

A automação de **ingestão de NF por e-mail** está pronta **do lado do sistema**
(Azure Functions). O que falta é **o lado do n8n** (montar o fluxo que recebe o
webhook e dispara o WhatsApp) + algumas **configurações** no Azure/SharePoint.

**Fluxo pretendido, ponta a ponta:**

```
Caixa de e-mail da diretoria
      │  (Graph Mail.Read)
      ▼
api/VarrerEmailsNF  ── classifica NF x ruído ── baixa PDFs ── salva no SharePoint
      │                                                         "Novas NFs - Automacao/{Unidade}/Diretoria {Diretoria}/"
      │  POST (webhook, Fase 2d)
      ▼
   n8n  ◄── FALTA MONTAR ──►  gateway WhatsApp (não-oficial) ──► gestor da diretoria
```

---

## 2. Pronto (lado do sistema) — já em produção

- **Fase 1** — `api/VarrerEmailsNF` lê e-mails, classifica (NF vs cobrança/boleto/
  interno), baixa os PDFs e grava no SharePoint. Classificador calibrado com dry-run real.
- **Fase 2a** — campo **e-mail** no cadastro de fornecedor (front).
- **Fase 2b** — corroboração pelo **Fechamento do Mês** (reforça a classificação).
- **Fase 2d (lado sistema)** — `VarrerEmailsNF` faz **POST no webhook do n8n** quando
  há candidatos (best-effort; se o webhook não estiver setado, apenas pula e reporta).

---

## 3. O que FALTA

### 3.1 Config no Azure (App Settings) — **pré-requisito do WhatsApp**

Static Web App → Configuration → Application settings:

| App Setting | Para quê | Status |
|---|---|---|
| `N8N_WEBHOOK_NF` | URL do webhook do n8n que recebe o POST da Fase 2d | ⬜ definir |
| `AUTOMACAO_EMAILS_SECRET` | Segredo do header `X-Automacao-Secret` p/ o cron chamar sem sessão (Fase 3) | ⬜ definir |
| `EMAIL_DOMINIO_INTERNO` | Domínio interno a ignorar (default `pronep.com.br`) | opcional |

> Enquanto `N8N_WEBHOOK_NF` estiver vazio, a varredura funciona normal e só
> reporta `notificado: { enviado:false, motivo:'N8N_WEBHOOK_NF nao configurado' }`.

### 3.2 Coluna no SharePoint (lista **PRONEP-NF-Diretorias**)

| Coluna | Tipo | Para quê |
|---|---|---|
| `TelefoneNotificacao` | Texto | Telefone **E.164** (ex.: `+5511999998888`) que RECEBE o WhatsApp. Casado por Unidade+Diretoria (fallback: só Diretoria). | ⬜ criar + preencher |

### 3.3 n8n — **o fluxo a montar (sua parte principal)**

Montar um workflow **novo, dedicado à Pronep** (decisão já tomada: novo Chatwoot
dedicado, gateway **não-oficial** tipo Evolution/Z-API). Ponto de partida: seu
`Byoterapia.json` (mesmo modelo, duplicado e adaptado).

**O n8n precisa:**
1. **Webhook node** (trigger) — a URL dele vira o `N8N_WEBHOOK_NF`.
2. Ler o **payload** (formato abaixo).
3. Montar a mensagem de WhatsApp (template curto: gestor, diretoria, nº de NFs, nomes de arquivo).
4. Enviar via **gateway** para o número `payload.telefone`.
5. (Opcional) Responder 200 rápido — o sistema trata como best-effort.

**Payload EXATO que o sistema envia** (de `api/VarrerEmailsNF/index.js`):

```json
{
  "gestor": "fulano@pronep.com.br",
  "diretoria": "Comercial",
  "unidade": "SP",
  "telefone": "+5511999998888",
  "pasta": "Novas NFs - Automacao/SP/Diretoria Comercial",
  "total": 2,
  "candidatos": [
    {
      "fornecedor": "ACME LTDA",
      "assunto": "NF 12345 - ACME",
      "confianca": "alta",
      "esperado": "ACME LTDA",
      "arquivos": ["nf-12345.pdf"]
    }
  ]
}
```

Campos úteis pra mensagem: `diretoria`, `unidade`, `total`, e por candidato
`assunto` / `fornecedor` / `arquivos`. O `telefone` é o destino.

> Se `telefone` vier vazio (diretoria sem `TelefoneNotificacao`), o sistema ainda
> faz o POST mas marca `semTelefone:true` — o n8n deve tratar (ex.: não enviar, ou
> mandar pra um número de fallback).

### 3.4 Fase 3 — automação agendada (depois do WhatsApp funcionando)

- **Cron** (GitHub Actions ou Azure) chamando `VarrerEmailsNF` para **todas as
  diretorias**, autenticado pelo header `X-Automacao-Secret: <AUTOMACAO_EMAILS_SECRET>`.
- Hoje o endpoint já aceita esse header (`viaSecret`) OU sessão admin.
- Decidir janela (`?dias=`) e frequência.

### 3.5 Fase 2c (opcional / adiada)

Notificação via SAN (assistente interno) além do WhatsApp. Não bloqueia nada.

---

## 4. Como testar (referência rápida)

Logado como **admin** no app (ou via `X-Automacao-Secret`):

```
# DRY-RUN — só classifica e reporta, não baixa nada:
/api/VarrerEmailsNF?gestor=<email-da-caixa>&dias=3&dryRun=1

# Real — baixa PDFs pra pasta e (se N8N_WEBHOOK_NF setado) dispara o webhook:
/api/VarrerEmailsNF?gestor=<email-da-caixa>&dias=3
```

Query params: `gestor` (obrigatório p/ resolver Unidade/Diretoria), `dias` (1–30,
default 3), `limite` (1–100, default 50), `dryRun` (`1`/`true`), `unidade`,
`diretoria` (override). Resposta traz `candidatos`, `baixados`, `notificado`,
`ignoradosResumo`.

Idempotência: um **ledger** no SharePoint evita reprocessar o mesmo e-mail.

---

## 5. Decisões já tomadas (não reabrir)

- **Estrutura de pastas** (geral do sistema): **mantida** — Aprovadas por data;
  Pendentes/Rejeitadas por `{Unidade}/Diretoria {Diretoria}`. Limpeza do legado
  **adiada** (não é prioridade). "Garantir daqui pra frente" = feito (ver PR #26).
- Pasta da automação: `Novas NFs - Automacao/{Unidade}/Diretoria {Diretoria}/`.
- WhatsApp: gateway **não-oficial**; Chatwoot **novo, dedicado à Pronep**.
- Permissão `Mail.Read` (Graph) já liberada e escopada.

---

## 6. Estado dos PRs abertos (verificar no GitHub ao retomar)

| PR | O quê | Ação ao retomar |
|---|---|---|
| #25 | Filtro de Motivo (dropdown) na tela de Rejeitadas | mergear se ainda aberto |
| #26 | `AprovarNota` resolve PDF por identidade exata (blindagem NumeroNF duplicado) + `shared/pdfNota.js` | mergear (correção crítica) |

Após mergear #23 (já feito): rodar 1× `/api/MigrarColunaRejeitadoPor` (admin) — **já executado** (colunas `RejeitadoPor`/`RejeitadoEm` criadas).

---

## 7. Primeiros passos na próxima sessão (checklist)

1. [ ] Conferir/mergear PRs #25 e #26; acompanhar deploy.
2. [ ] Criar coluna `TelefoneNotificacao` na lista Diretorias e preencher os números.
3. [ ] Montar o workflow n8n (webhook → mensagem → gateway) a partir do `Byoterapia.json`.
4. [ ] Setar `N8N_WEBHOOK_NF` (URL do webhook) nas App Settings.
5. [ ] Teste real de 1 diretoria: `/api/VarrerEmailsNF?gestor=...&dias=3` e conferir o WhatsApp.
6. [ ] Ajustar template da mensagem conforme o resultado.
7. [ ] Fase 3: cron + `AUTOMACAO_EMAILS_SECRET` + todas as diretorias.

**Produção:** `https://purple-forest-09588fe10.7.azurestaticapps.net`
