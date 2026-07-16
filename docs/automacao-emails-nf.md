# Automação: ingestão de NF por e-mail

Varre a caixa dos gestores por e-mails de NF/Fatura, baixa os PDFs para o SharePoint
e (nas próximas fases) notifica WhatsApp (via n8n) + SAN.

## Arquitetura
```
GitHub Actions (cron)  ->  /api/VarrerEmailsNF (secret)
  para cada gestor:
    Graph (Mail.Read) le a inbox  ->  classifica NF (assunto/remetente)
      ->  baixa PDF  ->  SharePoint: Novas NFs - Automacao/{Unidade}/Diretoria {Diretoria}/
      ->  ledger (nao repetir)  ->  [Fase 2] SAN + POST webhook n8n (WhatsApp)
```
Divisão: o **sistema** faz Graph/classificação/download/dedup/SAN; o **n8n** só recebe
um `POST` e dispara o WhatsApp (canal desacoplado — troca p/ Telegram sem mexer no app).

## Pré-requisitos (admin)
1. **Graph `Mail.Read` (Application)** com consentimento de admin, **ESCOPADO** por
   *Application Access Policy* (Exchange Online PowerShell) apenas às caixas dos gestores
   (um grupo de segurança). Sem isso o app poderia ler a organização inteira.
2. **App Setting** `AUTOMACAO_EMAILS_SECRET` (para o cron; Fase 3).
3. **App Setting** `N8N_WEBHOOK_NF` = URL do webhook do n8n (Fase 2).
4. **Lista Diretorias**: coluna `TelefoneNotificacao` (E.164, ex.: `+5511987654321`)
   por diretoria — número que RECEBE o WhatsApp, editável no Mapa de Aprovadores (Fase 2).

## Destino dos arquivos
SharePoint: `Novas NFs - Automacao/{Unidade}/Diretoria {Diretoria}/` (pastas criadas sob
demanda). Para ter localmente, o gestor sincroniza essa biblioteca via **OneDrive** — a
automação nunca escreve no disco da máquina.

## Fases
- **Fase 1 (feita): ler + classificar + baixar.** Endpoint `VarrerEmailsNF` varre 1 caixa,
  identifica NF (assunto: nota fiscal/NF/fatura/boleto/DANFE; ou remetente ∈ Fornecedores),
  baixa o PDF e grava no ledger. Sem notificação.
- **Fase 2:** notificação SAN + `POST` no webhook n8n (WhatsApp), com telefone da lista.
- **Fase 3:** cron (GitHub Actions) varrendo todas as diretorias + rota anonymous p/ o secret.

## Como testar a Fase 1 (após o `Mail.Read` estar ativo)
Logado como **admin**, no navegador:
```
# 1) DRY-RUN — so classifica e reporta, nao baixa nada:
/api/VarrerEmailsNF?gestor=fulano@pronep.com.br&dias=7&dryRun=1

# 2) Real — baixa os PDFs para a pasta:
/api/VarrerEmailsNF?gestor=fulano@pronep.com.br&dias=7&unidade=SP&diretoria=Tecnologia
```
- `dryRun=1` primeiro para conferir a classificação sem efeito colateral.
- `unidade`/`diretoria` são opcionais (tenta derivar da lista Diretorias pelo e-mail do
  gestor); passe explícito se a derivação não achar.
- A resposta traz `candidatos`, `baixados`, `ignoradosResumo` e `avaliados`.

## Idempotência
Sem `Mail.ReadWrite` (não alteramos a caixa): os `messageId` processados ficam num ledger
JSON no drive (`_automacao/emails_ledger.json`) + janela de dias. Reexecuções não repetem.

## Classificação (calibrada com dry-run real de 16/07)
Anexo **PDF** é obrigatório. A decisão usa assunto + **nome do PDF** + remetente, e
devolve um nível de **confiança** pra o gestor priorizar:
- **FORTE** (`confianca: alta`): "nota fiscal" / NF / NFe / NFS-e / DANFE no assunto
  **ou** no nome do arquivo (ex.: `NF 10832 - PRONEP.pdf`). Vira candidato sozinho.
- **FRACO** (`fatura` / `boleto`): só entra se **corroborado** —
  - `media`: remetente ∈ Fornecedores conhecidos;
  - `baixa`: e-mail **encaminhado** (ENC:/FW:) — padrão real (colega repassa a NF).
- **NEGATIVO**: aviso de cobrança/débito (`não identificado`, `débito`, `em aberto`,
  `inadimplência`…) → **excluído**, salvo se houver sinal FORTE. Mata o falso positivo
  típico de "fatura não identificada / débitos abertos".
- **Domínio interno** (`pronep.com.br`, configurável por `EMAIL_DOMINIO_INTERNO`) é
  **excluído** do allowlist de remetente — senão todo colega vira "fornecedor".

> Aprendizado do dry-run: NFs chegam muito por **encaminhamento interno** (o "De" é um
> colega, não o fornecedor). Por isso casamos pelo **nome do fornecedor recorrente**
> (do Fechamento) no assunto/arquivo, não só pelo remetente.

## Corroboração pelo Fechamento (Fase 2b — feita)
Reusa o `computar()` do Fechamento do Mês: pega os fornecedores **recorrentes ainda
pendentes** no mês (status atrasada/risco/aguardando) da diretoria, e cruza o **nome**
deles com o assunto/nome do PDF do e-mail. Funciona mesmo em encaminhamentos (casa pelo
conteúdo, não pelo remetente). Efeito:
- Candidato que casa com um recorrente esperado → sobe pra `confianca: alta`
  (`motivo` ganha `+esperado_fechamento`).
- Sinal **FRACO** não-negativo (ex.: boleto de um recorrente) que sozinho seria
  descartado → **resgatado** como candidato `alta` (`fraco+esperado_fechamento`).
- Aviso de cobrança/débito (NEGATIVO) **não** é resgatado, mesmo citando um esperado.

Cada candidato traz `esperado: "<nome do fornecedor>"` (ou `null`), e a resposta lista
`esperadosFechamento` (o que o Fechamento esperava naquele mês/diretoria).
