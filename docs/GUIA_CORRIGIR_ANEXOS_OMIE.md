# Guia rápido — Corrigir anexos de NF no Omie

**Para:** Financeiro (Izabel)
**Quando usar:** quando a tela **Auditoria Omie (PDFs)** apontar NFs suspeitas de terem o PDF errado anexado no Omie.

---

## 1. Abrir a auditoria

1. Entre no sistema de Aprovação de NF com seu login (perfil Financeiro ou Admin).
2. No menu lateral, clique em **Auditoria Omie (PDFs)** (logo abaixo de "Auditoria").
3. A tela lista, no topo, três números:
   - **Integradas** — total de NFs já enviadas ao Omie.
   - **Suspeitas** — as que precisam de conferência (é a lista da tabela).
   - **OK** — as que estão certas (não precisa mexer).
4. Se aparecer **"✅ Nenhuma integração suspeita"**, não há nada a corrigir. Fim.

> Dica: o botão **↗ Abrir em nova aba** mostra a mesma lista em página limpa, boa pra imprimir ou ir riscando conforme corrige.

---

## 2. Entender cada linha da tabela

Cada linha é uma NF que pode estar com o PDF errado no Omie:

| Coluna | O que significa |
|---|---|
| **NF** | Número da nota. |
| **Unidade / Diretoria** | Onde a NF foi lançada. |
| **Fornecedor / Valor** | Pra você confirmar de qual conta se trata. |
| **PDF correto** (verde) | O nome do arquivo que **deveria** estar anexado. |
| **Provável anexado (antigo)** (âmbar) | O arquivo que o sistema antigo **provavelmente** anexou por engano. |
| **Motivo** | Por que entrou na lista de suspeitas. |

**Motivos possíveis:**
- **Regra antiga anexaria OUTRO arquivo** → alto risco: compare "PDF correto" x "provável anexado". Se forem diferentes, o anexo no Omie está errado.
- **Número casa com vários arquivos (ambíguo)** → confira manualmente qual PDF está no Omie.
- **PDF correto não identificável com segurança** → não deu pra determinar o certo automaticamente; abra a NF no sistema pra ver o PDF aprovado.

---

## 3. Conferir e corrigir no Omie

Para **cada linha suspeita**:

1. No Omie, localize a **conta a pagar** dessa NF (pelo número da NF + fornecedor + valor da tabela).
2. Abra os **anexos** da conta e veja qual PDF está lá.
3. **Abra o PDF anexado** e confira o conteúdo:
   - Confere com a NF (número, fornecedor, valor)? → está **certo**, não mexa.
   - É de **outra nota**? → siga pra correção:
4. **Remova** o anexo errado da conta no Omie.
5. **Anexe o PDF correto**:
   - No sistema de Aprovação de NF, vá em **Notas Aprovadas**, localize a NF e baixe o PDF aprovado dela (é o arquivo com o nome que aparece em **"PDF correto"** na auditoria).
   - Anexe esse arquivo na conta do Omie.
6. Risque a linha da sua lista e siga pra próxima.

---

## 4. Confirmar que ficou tudo certo

- Depois de corrigir, volte à tela **Auditoria Omie (PDFs)** e clique em **🔄 Atualizar**.
- As NFs já corrigidas continuarão aparecendo (a auditoria não sabe o que você mudou no Omie — ela olha o histórico do sistema, não o Omie em tempo real). **Use sua lista impressa como controle do que já foi feito.**

---

## Observações importantes

- **A auditoria é conservadora:** ela sinaliza tudo que estava *em risco*. "Suspeita" quer dizer **"confira este anexo"**, não necessariamente "está errado". Sempre abra o PDF no Omie pra confirmar antes de trocar.
- **Novas NFs já sobem corretas.** O erro que causava a troca de PDF foi corrigido no sistema; esta auditoria serve só pra limpar o que ficou errado do período anterior.
- **Dúvida em alguma linha?** Fale com o Rafael antes de remover o anexo, pra não apagar um PDF que estava certo.
