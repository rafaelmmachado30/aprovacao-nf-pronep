# Ajuste de tipos de colunas no SharePoint (NotasFiscais)

**Objetivo:** corrigir 4 colunas da lista `PRONEP-NF-NotasFiscais` para os tipos corretos. Hoje funciona porque tudo está como `text`, mas isso é dívida técnica:
- **Valor** deveria ser `Currency (R$)` — pra somas/relatórios funcionarem
- **NumeroNF** deveria ser `Single line of text` — hoje pode ter sido criada como número, o que limita formatos
- **UrlPDF** e **UrlPDFAprovado** deveriam ser `Hyperlink or Picture` — hoje guardam URL como string

O backend (PostNota, AprovarNota) **já detecta** o tipo da coluna e formata o valor adequadamente. Então depois que você muda no SharePoint, **não precisa redeploy**.

---

## Como chegar na configuração de colunas

1. Abre o site SharePoint: https://pronepadmin.sharepoint.com/sites/Aprovacao-NotasFiscaisServicos
2. Lateral esquerda → **PRONEP-NF-NotasFiscais** (ou via "Conteúdo do site" se não estiver na lateral)
3. Engrenagem no topo direito ⚙ → **Configurações da lista** (List settings)
4. Role pra baixo até a seção **Colunas**

Cada coluna lá tem o tipo entre parênteses, tipo "(Linha única de texto)", "(Número)", etc.

---

## Coluna 1 — Valor (R$): mudar pra Currency

**Verificação primeiro:** clica em "Valor" na lista de colunas. Se já estiver como **Moeda (R$)** ou **Currency**, pula essa parte.

**Se estiver como Número ou Linha de texto:**

1. Clica em **Valor** → abre a página de edição
2. No bloco "O tipo de informação nesta coluna é", muda para **Moeda (Currency)**
3. Configurações de moeda:
   - **Casas decimais:** 2
   - **Formato de moeda:** R$ (Brasil)
   - **Valor mínimo:** 0
4. **OK** no final

Se o SharePoint perguntar se quer converter dados existentes, **diga sim** — ele converte os números armazenados como string pra float.

---

## Coluna 2 — NumeroNF: garantir Single line of text

**Verificação:** clica em **NumeroNF**. Se já estiver como **Linha única de texto (Single line of text)**, pula.

**Se estiver como Número:**

1. Clica em **NumeroNF**
2. Muda tipo pra **Linha única de texto**
3. Tamanho máximo: 50
4. **OK**

**Por quê texto:** notas fiscais às vezes têm zeros à esquerda (`000123`) ou letras (`A-12345`). Como número, o SharePoint corta os zeros e rejeita letras.

---

## Coluna 3 — UrlPDF: mudar pra Hyperlink or Picture

1. Clica em **UrlPDF**
2. Muda tipo pra **Hiperlink ou imagem (Hyperlink or Picture)**
3. Formato: **Hiperlink** (não Imagem)
4. **OK**

Se o SharePoint avisar que tem dados existentes que não cabem no novo tipo, ele vai oferecer 2 opções:
- **Manter dados existentes** (deixa o que tem, valores antigos vão aparecer como link clicável)
- **Apagar dados existentes** (limpa tudo)

Escolha **Manter** — as NFs já lançadas vão funcionar como link.

---

## Coluna 4 — UrlPDFAprovado: mudar pra Hyperlink or Picture

Mesmo processo da Coluna 3, mas pra **UrlPDFAprovado**.

---

## Validação depois das mudanças

1. Lança uma NF nova de teste pelo sistema
2. Abre a lista no SharePoint
3. Confere que:
   - **Valor** aparece formatado tipo `R$ 1.000,00` (não `1000`)
   - **NumeroNF** mantém zeros à esquerda se tiver
   - **UrlPDF** aparece como link clicável (não texto)
   - Clicar no link abre o PDF
4. Aprova essa NF
5. Confere **UrlPDFAprovado** também como link clicável

Tudo OK → encerra esse débito técnico.

---

## Troubleshooting

**"A coluna não pode ser convertida porque tem dados incompatíveis":**
Faz a conversão em 2 etapas:
1. Adiciona nova coluna (`ValorNum`, `UrlPDFLink`) com o tipo certo
2. Em "Configurações da lista", roda uma View que copia dados via Power Automate ou export/import CSV
3. Apaga a coluna antiga e renomeia a nova

**Sistema parou de funcionar depois da mudança:**
Lança NF nova de teste. Olha no F12 → Network → `/api/PostNota` → Response → bloco `diag.formatLog`. Cada coluna lista o tipo detectado pelo backend (`text`, `currency`, `hyperlink`, etc) e o valor formatado. Se algo virou string num lugar que precisava ser número, me manda esse JSON e ajusto o `formatByType` no backend.

**As NFs antigas (importadas) perderam Valor / Url:**
Isso pode acontecer se você escolheu "Apagar dados existentes" na conversão. O sistema continua funcionando pras novas; pras antigas, edita item por item no SharePoint pra preencher de novo (raramente compensa).
