# Contexto do Projeto — Sistema de Aprovação de NF (Pronep Life Care)

> Arquivo de handoff para continuidade entre máquinas/sessões do Cowork.
> Se você (Claude) está lendo isto numa máquina nova, este é o resumo do projeto
> e das convenções. O histórico de conversa do Cowork é **local por máquina** e não
> sincroniza — este documento existe para reconstituir o contexto rapidamente.

## Visão geral

Sistema interno de aprovação de Notas Fiscais de Serviço da Pronep. Dono: Rafael
Machado (rafael.machado@pronep.com.br), admin e gestor da diretoria de **Tecnologia**.

- **Frontend:** SPA única em `wwwroot/index.html` (~8k linhas, vanilla JS + CSS inline).
- **Backend:** Azure Functions (Node, modelo v3 com `function.json`), pasta `api/`.
  authLevel `anonymous` + Easy Auth. Helpers em `api/shared/`.
- **Dados:** listas SharePoint via Microsoft Graph (PRONEP-NF-NotasFiscais,
  -Fornecedores, -Contratos, -Diretorias, -Config, -Recorrentes).
- **Auth/RBAC:** Entra ID; grupos AAD → roles em `shared/userRoles.js` (GROUP_TO_ROLE,
  ROLE_LABELS). Admin centralizado em `shared/authz.js` (isAdminEmail/ADMIN_EMAILS).
- **IA (SAN):** assistente com tool-calling em `shared/sol.js` (endpoint SolChat).
- **Integração Omie (ERP):** `shared/omie.js` + endpoint IntegrarOmie (anexa PDF
  aprovado na conta a pagar do Omie da unidade — SP/RJ/ES).

## Deploy

- Git → GitHub → **GitHub Actions** → Azure Static Web Apps.
- Fluxo do Rafa (no terminal, na pasta do projeto):
  ```
  git add <arquivos>
  git commit -m "..."
  git push origin main
  ```
- Alertas diários NÃO usam TimerTrigger (SWA não suporta) — são endpoints HTTP
  disparados por cron do GitHub Actions (`.github/workflows/*.yml`) com header
  `X-Alerta-Secret: <ALERTA_DIARIO_SECRET>`.

## Convenções importantes

- Datas de pasta no SharePoint usam **BRT (UTC-3)**; `AprovadoEm` é gravado em UTC.
  (Bug corrigido no IntegrarOmie: procurava o PDF pela data UTC e errava a pasta
  após 21h BRT — agora calcula BRT + varre subpastas como fallback.)
- Vencimento "fora do prazo" = menos de **D+5 dias úteis**.
- Escape XSS no front: usar sempre `escHtml` / `escAttr` (globais).
- Validar sintaxe antes de commitar: `node --check` nos arquivos `.js`.
  (O mount do shell às vezes desincroniza/trunca; quando isso ocorrer, confirmar o
  conteúdo real pela ferramenta Read e reconstruir o arquivo no /tmp pra checar.)

## O que já foi construído nesta jornada

1. **Auditoria de segurança + correções (C1–C7):** fechamento de rotas, gates de
   auth/role, XSS (escHtml global), hardening do link de aprovação por e-mail.
2. **Bugs ALTOS** corrigidos (A1, A2, A5, A6, A8, A9, A10).
3. **Controle de Acessos** a contratos (granular: pasta → grupo → pessoas), helper
   `shared/acessoContratos.js` + endpoints GetControleAcessos/SalvarControleAcessos/
   ListarMembrosGrupo.
4. **Contratos:** árvore espelhando o SharePoint, PDF servido pelo app (streaming),
   limpeza de duplicados (LimparContratosDuplicados).
5. **Omie:** correção do "PDF não encontrado" (data BRT) e UI de progresso honesta.
6. **Compliance fora do prazo:** layout do bloco de alinhamento financeiro corrigido.
7. **Contas Recorrentes (Estágio 1):** detecção + confirmação do gestor.
   - `shared/recorrentes.js`, endpoints CriarListaRecorrentes / ListarRecorrentes /
     SalvarRecorrente. Lista `PRONEP-NF-Recorrentes`. Aba "Contas Recorrentes".
8. **Fechamento do Mês (Estágio 2):** checklist do mês com status anti-D+5
   (aguardando / risco D+5 / atrasada / lançada / aprovada). Endpoint
   ChecklistRecorrentes + aba "Fechamento do Mês".

## Pendências / próximos passos

- **Estágio 3 (Contas Recorrentes):** alerta automático (e-mail + in-app) das contas
  em risco D+5 / atrasadas, via cron do GitHub Actions (reusar `shared/notificar.js`).
- **Caixa de Entrada de PDFs:** subir nota e boleto separados, combinar num PDF e
  lançar (a partir de PDFs avulsos no staging).
- Feriados no cálculo de D+5 (hoje só conta dias úteis seg–sex).
- Rollout das abas de recorrentes para os demais gestores (hoje só admin).

## Identidade visual / skills

Há skills Pronep instaladas (papel timbrado, wrapper de PPT 4K, auditoria financeira
de pacientes, repactuação de diária). Usar quando gerar documentos formais Pronep.
