# Sistema de Aprovacao de NF — Roadmap Sprint 4 / Fechamento

**Pronep Life Care · Sistema de Aprovacao de Notas Fiscais**

Documento consolidado de fechamento. Cobre o que ficou pronto, o que ainda esta pendente, riscos conhecidos e recomendacoes pos go-live.

---

## 1. Estado atual

### Arquitetura

- **Front-end:** Azure Static Web Apps (Free tier), HTML/CSS/JS vanilla
- **Back-end:** Azure Functions (Node.js v22) integradas ao SWA
- **Identidade:** Microsoft Entra ID via Easy Auth (browser) + getAuthToken SSO (Teams)
- **Persistencia:** Microsoft SharePoint Lists (3 listas + Drive)
- **Notificacoes:** Microsoft Graph API (email Mail.Send + Teams sendActivityNotification)
- **Deploy:** GitHub Actions

### Funcionalidades implementadas

**Operacao basica (Sprint 1-2):**
- Lancamento de NF com upload de PDF
- Validacao automatica de CNPJ (Brasilapi)
- Detecao de duplicidade (hash SHA-256 + CNPJ+Numero+Serie)
- Roteamento automatico de aprovador pela matriz Unidade x Diretoria
- Bloqueio de envio quando fornecedor incompleto ou vencimento fora do prazo
- Cadastro de Fornecedores via CRUD completo (criar, editar, importar em massa)
- Importacao XLSX/CSV com validacao linha-a-linha

**Aprovacao (Sprint 3):**
- Fila de Aprovacao filtrada por RBAC (admin/financeiro/gestor/submitter)
- Aprovar com watermark APROVADA + data + aprovador no PDF
- Rejeitar com motivo + watermark REJEITADA
- Arquivamento automatico em pastas SharePoint (Pendentes/Aprovadas/Rejeitadas)
- Aprovacao via link assinado no email (sem precisar logar no sistema)

**Notificacao (Sprint 3-4):**
- Email automatico em todos os eventos (lancada/aprovada/rejeitada)
- Botoes Aprovar/Rejeitar inline no email com JWT assinado (validade 7 dias)
- Notificacao Teams 1:1 via Graph API sendActivityNotification (substituiu webhook depreciated)
- Instalacao automatica da Teams App pro aprovador antes de notificar
- Personal Tab Teams: clicar na notif abre o sistema dentro do Teams

**Dashboard / Auditoria:**
- KPIs do mes: valor aprovado, NFs aprovadas, pendentes, duplicidades evitadas
- Graficos: valor por diretoria, distribuicao por unidade, evolucao mensal
- Mapa de Aprovadores (matriz visual)
- Tela Auditoria com historico completo

---

## 2. O que sobra (backlog)

### Confirmados pendentes

| Item | Esforco | Bloqueio |
|---|---|---|
| #79 PWA (app instalavel mobile) | ~1h-1h30 | Nao bloqueante |
| #127 Validar JWT Teams (silent SSO) | ~15min | Precisa print do authError |
| #128 Testar aprovacao end-to-end no Teams | ~10min | Apos #127 |

### Setups manuais no SharePoint (voce executa quando puder)

- Mudar **Valor (R$)** pra coluna do tipo Currency
- Garantir **NumeroNF** como Single line of text
- Mudar **UrlPDF** e **UrlPDFAprovado** pra Hyperlink

Guia detalhado em [SETUP_TIPOS_COLUNA_SHAREPOINT.md](SETUP_TIPOS_COLUNA_SHAREPOINT.md). Backend ja suporta os 3 tipos automaticamente via `formatByType`.

### Melhorias futuras (nao prioritarias)

- Bot Framework do Teams pra Adaptive Card com botoes Aprovar/Rejeitar inline no chat (1 dia de trabalho)
- Workflow de re-aprovacao quando NF rejeitada eh re-lancada
- Dashboard com filtros por mes/diretoria/aprovador
- Export do Dashboard pra Excel
- Cache mais agressivo na lista de Fornecedores (3300 items)
- Multi-aprovador (cadeia de aprovacao N1 -> N2)

---

## 3. Riscos conhecidos

### Tecnicos

**Cookie partitioning em iframe Teams:** Easy Auth cookies nao passam entre o iframe Teams e o popup de auth. **Mitigado** com SSO via getAuthToken (token MSAL passa em Authorization Bearer, dispensa cookie). Necessario validar audience/issuer no JWT (task #127).

**Limites do Graph API:** sendActivityNotification limitado a ~50 notif/usuario/hora. Em volume normal de NFs (<10/dia/aprovador) nao tem problema.

**SharePoint List 5000 items:** lista de Fornecedores pode crescer alem disso. Mitigado: usa paginacao via @odata.nextLink no ListarFornecedores.

**Free tier do SWA:** limite de 500.000 requests/mes + 100 GB bandwidth. Ate ~50 aprovadores ativos cabe folgado.

### Operacionais

**Dependencia de admin Microsoft 365:** mudancas nas permissoes do App Reg ou Teams App requerem admin consent. Documentado em SETUP_TEAMS_ACTIVITY.md e SETUP_TEAMS_SSO.md.

**Token de email assinado (7 dias):** se a NF ficar pendente alem disso, o link Aprovar/Rejeitar do email expira. Usuario tem que abrir o sistema. Aceitavel pra fluxo padrao (aprovacao em 1-2 dias).

**1 unica conta enviando emails:** `datanalytics@pronep.com.br` envia todos os emails do sistema. Se essa conta tiver problema (suspensa/MFA), email para. Recomendado: criar uma service account dedicada (`sistema-aprovacao-nf@pronep.com.br`) sem licenca interativa.

---

## 4. Recomendacoes pos go-live

### Imediato (primeira semana)

1. **Monitoramento basico**: Application Insights ja deveria estar plugado no SWA. Confirmar dashboards de erros 5xx.
2. **Comunicacao com aprovadores**: mensagem-padrao explicando o novo fluxo (Atividade Teams + email com botoes + sistema web).
3. **Backup da lista Fornecedores**: exportar XLSX semanal (manual ou via Power Automate).

### Curto prazo (primeiros 30 dias)

1. Coletar feedback dos aprovadores. As reclamacoes mais comuns viraram melhorias prioritarias.
2. Auditoria de log: alguma NF foi rejeitada por erro de roteamento (aprovador errado)? Atualizar matriz Diretorias.
3. Avaliar se o volume justifica criar bot Framework com Adaptive Card.

### Medio prazo (3-6 meses)

1. Re-avaliar permissoes do App Reg — remover as que nao estao em uso.
2. Considerar migracao do SharePoint List pra Azure SQL ou Cosmos se a volumetria justificar (>10k NFs/mes).
3. Implementar relatorio mensal automatico pra Diretoria via email/Teams (Azure Functions + scheduler).

---

## 5. Documentacao do projeto

**Setup / operacao:**
- `LEIA-ME.md` — overview do projeto
- `GUIA_AZURE_SWA.md` — provisionamento do SWA
- `GUIA_ENTRA_ID.md` — App Registration + permissoes Graph
- `GUIA_SHAREPOINT.md` — estrutura das 3 listas
- `SETUP_TEAMS_ACTIVITY.md` — Teams App + sendActivityNotification
- `SETUP_TEAMS_SSO.md` — Expose API + scope access_as_user + Teams clients autorizados
- `SETUP_TIPOS_COLUNA_SHAREPOINT.md` — ajustes pendentes nos tipos de coluna

**Codigo:**
- `wwwroot/index.html` — SPA inteira em HTML/CSS/JS vanilla
- `api/` — 14 Azure Functions (ver lista abaixo)
- `teams-app/` — manifest da Teams App + icones

**Functions implementadas:**

| Function | Metodo | Funcao |
|---|---|---|
| `MeusGrupos` | GET | Resolve user via Entra ID + retorna roles |
| `ListarFornecedores` | GET | Lista da SharePoint List |
| `AdicionarFornecedor` | POST | Cria fornecedor (com check de duplicata) |
| `EditarFornecedor` | PATCH | Edita parcial de fornecedor |
| `ListarDiretorias` | GET | Matriz Unidade x Diretoria x Aprovador |
| `ListarGestoresFinanceiro` | GET | Membros do grupo Entra ID gestores |
| `ConsultarCNPJ` | GET | Brasilapi (fallback ReceitaWS) |
| `ListarNotas` | GET | Lista NFs filtrada por RBAC |
| `PostNota` | POST | Lanca NF (upload PDF + SP item + notificar) |
| `AprovarNota` | POST | Aprova + watermark + move PDF + notificar |
| `RejeitarNota` | POST | Rejeita + watermark + move PDF + notificar |
| `AprovacaoViaLink` | GET | Aprovacao via token JWT do email |
| `AbrirPdfDaNota` | GET | Redirect 302 pro PDF no SharePoint |
| `EnviarNotificacao` | POST | Endpoint isolado pra notificar manualmente |

---

## 6. Linha do tempo

**Sprint 0** — Prototipo + arquitetura
**Sprint 1** — Deploy + Easy Auth + Graph API setup
**Sprint 2** — SharePoint Lists + ListarNotas/Fornecedores + lancamento real
**Sprint 3** — Aprovar/Rejeitar com watermark + email com links assinados + Teams webhook
**Sprint 4** — Personal Tab Teams + SSO real + Editar/Adicionar Fornecedor + Importacao em massa + Roadmap

**Linhas de codigo:** ~3300 linhas no `index.html` + ~14 Functions Azure (~1500 linhas Node.js) + ~200 linhas de shared modules

**Bugs resolvidos no caminho:** 50+ (lista completa nas tasks do projeto, da #66 a #128)

---

## 7. Contatos / responsabilidades

| Funcao | Responsavel | Acao |
|---|---|---|
| Owner do produto | Rafael Machado | Decisao sobre roadmap e prioridade |
| Aprovadores | Sandra, Vitor, Monica, Bruno, etc | Aprovar/rejeitar NFs |
| Admin Microsoft 365 | TI Pronep | Permissoes Graph + Teams Admin |
| Owner SharePoint | TI Pronep | Listas + estrutura de pastas |
| Owner email | datanalytics@pronep.com.br | Conta que envia emails do sistema |

---

## 8. Decisao de fechamento

O sistema esta **production-ready** com excecao dos pendentes #127 (validacao JWT no SSO Teams) e #128 (teste end-to-end Teams). Mesmo sem #127/#128 fechados, o sistema funciona via:
- Browser (Easy Auth normal) — caminho principal
- Email com botoes Aprovar/Rejeitar — caminho rapido
- Teams personal tab com login interativo (fallback) — quando SSO silent falhar

**Go/no-go pra producao:** GO. Pendencias #127/#128 sao de polish (UX dentro do Teams), nao funcionais.

---

*Documento gerado em fechamento do Sprint 4. Versao 1.0.*
