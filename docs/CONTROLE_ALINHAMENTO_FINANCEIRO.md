# Controle: Confirmação de Alinhamento com o Financeiro (NFs < D+5)

> Gate que impede o aprovador de "burlar" o alinhamento financeiro. Antes, ao aprovar
> uma NF vencendo em **menos de D+5**, bastava marcar "alinhei com o financeiro" e a NF
> era aprovada na hora — sem o financeiro confirmar nada. Agora é um **gate real**.
> Entregue na sessão de 2026-07-24.

---

## O problema
NF vencendo em < D+5 exige, na aprovação, que o gestor declare se **alinhou** o pagamento
com o financeiro. O sistema **confiava** na declaração → aprovadores marcavam "alinhei"
sem alinhar de verdade.

## A solução (gate)
Quando o aprovador declara **"alinhei" + escolhe uma pessoa do Financeiro-Gestão**, a NF
**não é aprovada**: fica num estado intermediário até o **Financeiro-Gestão confirmar**.

### Fluxo
```
Aprovador aprova NF <D+5 + "alinhei" + escolhe financeiro
      │
      ▼
Status = AguardandoAlinhamento   (PDF continua em Pendentes; NADA é carimbado/movido)
      │  alerta "OK" pro aprovador  +  e-mail/Teams pro Financeiro-Gestão
      ▼
Tela "Confirma Alinhamento"  (Financeiro-Gestão age; Gestor acompanha)
      ├── Confirmar → executa a aprovação REAL (carimbo do GESTOR, move p/ Aprovadas, Status=Aprovada)
      └── Rejeitar  → Status=Rejeitada (motivo "alinhamento não confirmado pelo financeiro")
```

### Decisões de negócio (travadas)
- **Rejeição do financeiro → NF Rejeitada** (o solicitante relança).
- **Qualquer** membro do Financeiro-Gestão pode confirmar/rejeitar (não trava se a pessoa
  escolhida estiver ausente).
- Financeiro é avisado por **e-mail + Teams** (além da tela).
- **O carimbo continua sendo o do GESTOR** (aprovador original), não do financeiro. A
  confirmação do financeiro só *destrava*; a responsabilidade da aprovação segue do gestor.
  O papel do financeiro fica na **auditoria** (`alinhamento_confirmado`/`alinhamento_rejeitado`).

---

## Como funciona por dentro

**Estado novo:** `Status = AguardandoAlinhamento` (a coluna Status é **texto** — aceita
qualquer valor; não precisou migração).

**Backend:**
- `AprovarNota` — após o roteamento multi-nível, um **gate**: se `alinhouFinanceiro=true`
  (e há `gestorFinanceiroAlinhado`), grava `AguardandoAlinhamento` + notifica o grupo e
  **retorna sem mover o PDF**. Flag interno `bypassAlinhamento` pula o gate **e o N2**
  (usado pela confirmação).
- `ConfirmarAlinhamento` (POST) — RBAC: `financeiro_nf` ou admin. Valida
  `Status=AguardandoAlinhamento`. `confirmar` → chama `AprovarNota` em-processo com
  `bypassAlinhamento` e **principal do aprovador original** (RBAC valida por ele) →
  aprovação real. `rejeitar` → chama `RejeitarNota` em-processo (Status=Rejeitada).
- `shared/financeiroGestao.js` — e-mails do grupo `PRONEP-NF-Financeiro-Gestao`.
- `email.js` — evento `alinhamento_pendente` (e-mail + Teams).
- `MigrarStatusAlinhamento` — diagnóstico/idempotente (Status é texto → nada a migrar;
  se fosse choice, adicionaria a opção).

**Front (`wwwroot/index.html`):**
- Menu **"Confirma Alinhamento"** (abaixo de Fila de Aprovação), visível a
  Financeiro/Gestor/Admin (`canSee`), com badge.
- View `confirma-alinhamento`: lista `AguardandoAlinhamento` — **efêmera** (só o pendente;
  some após a ação). Financeiro/Admin: botões Confirmar/Rejeitar. Gestor: leitura.
- `doApprove`: ao receber `aguardandoAlinhamento:true`, mostra o modal "⏳ Enviado para
  confirmação do Financeiro" com **OK**.
- Funções `confirmarAlinhamento`/`rejeitarAlinhamento` → `POST /api/ConfirmarAlinhamento`.
- `statusPill`: selo âmbar "⏳ AGUARDANDO FINANCEIRO".

**RBAC:** o grupo Entra **`PRONEP-NF-Financeiro-Gestao`** (OID `c2a73d16-…`, role
`financeiro_nf`, App Setting opcional `GESTOR_FINANCEIRO_GROUP_ID`) define quem confirma
e quem recebe o aviso.

---

## Operação / pontos de atenção
- Manter o grupo `PRONEP-NF-Financeiro-Gestao` com os membros certos (recebem o aviso +
  confirmam).
- Pra testar, gerar uma NF com **vencimento dentro de ~5 dias úteis** (aciona o bloco de
  alinhamento na aprovação). Admin também consegue confirmar (teste solo).
- **Brecha conhecida (follow-up):** o backend não *exige* o alinhamento em NFs <D+5 — a
  regra é imposta pelo modal do front. Uma aprovação por caminho alternativo (ex.: ação via
  SAN) que não envie `alinhouFinanceiro` aprovaria direto. Se virar problema, dá pra
  reforçar no `AprovarNota` (calcular D+5 no backend e exigir o alinhamento).

## PRs
#48 (feature Fases 1+2), #49 (migração tolerante a Status texto).
