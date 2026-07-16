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

## Classificação (ajustável)
- **Assunto**: regex `nota fiscal | NF | NFS-e | fatura | boleto | DANFE | cobrança`.
- **Remetente**: e-mail/domínio ∈ Fornecedores (campo email). Best-effort (se a leitura
  falhar, cai só no assunto). Anexo **PDF** é obrigatório para baixar.
