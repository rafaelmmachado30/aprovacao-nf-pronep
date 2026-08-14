# Handoff — trazer XML da NF-e e DANFE para o sistema Pronep

Escrito em 14/08/2026, no fim de uma sessão que resolveu isso no **sistema autônomo**
(`~/Pronep/aprovacao-nf-supabase`). Este documento existe para a próxima sessão começar
sabendo o que já foi decidido, o que foi **provado**, e — principalmente — **o que eu não
verifiquei**, para ninguém repetir o erro de opinar sem olhar.

---

## 1. O pedido, na frase do Rafael

> "Na caixinha do kanban, não dá pra trazer o link pra acessar a DANFE? Para baixar a NF.
> Ou alguma informação a mais"

E depois: **"conseguimos replicar esta mesma ação do xml no sistema Pronep?"**

Ele aprovou a ordem em três etapas, e as duas primeiras **já estão prontas no autônomo**:

1. Baixar XML — feito
2. Ler o XML e mostrar itens, impostos, informação complementar — feito
3. **DANFE em PDF gerada a partir do XML** — não começou

---

## 2. A decisão de arquitetura, já tomada e com motivo

**Link para DANFE não existe.** O portal da SEFAZ exige CAPTCHA (e contornar CAPTCHA está
fora de questão). Sites terceiros que "abrem DANFE pela chave" significam mandar dado fiscal
da empresa para fora — não fazer sem decisão explícita do Rafael.

**O XML é a resposta**, porque o XML *é* o documento fiscal; a DANFE é só a representação
impressa dele. Com o XML em mãos, dá para baixar, mostrar detalhes e gerar o PDF.

### E aqui está o ponto que decide tudo na Pronep

No autônomo o XML **passa pela nossa mão**: a fonte é a SEFAZ, e o único problema era que o
código lia o XML, extraía número/valor/emitente e **jogava fora**.

Na Pronep a fonte do quadro NFs a Pagar é o **Omie**, que entrega dados de conta a pagar —
**não o XML autorizado**. Então lá não há XML para guardar.

Dois caminhos, e a recomendação já está dada ao Rafael:

| | Caminho A — religar SEFAZ na Pronep | Caminho B — anexo do Omie |
|---|---|---|
| O que dá | XML autorizado, original, com validade jurídica | provavelmente PDF anexado à mão |
| Serve para DANFE? | **sim** | não |
| Custo | ligar o que já existe e foi provado | rápido de tentar, mas incerto |
| Risco | — | depende de alguém ter anexado; não é o documento original |

**Recomendação: caminho A.** O XML da SEFAZ tem prazo de captura de **90 dias** — o que não
for buscado se perde. Anexo de ERP é cópia.

> ⚠️ **Não verifiquei o repositório da Pronep nesta sessão.** Tudo acima sobre a Pronep vem do
> que construímos em sessões anteriores, não de leitura de código agora. O Rafael foi avisado
> disso explicitamente.

---

## 3. PRIMEIRO PASSO da próxima sessão: ler antes de falar

Nesta ordem, e sem prometer nada antes de confirmar:

1. **Estado do conector SEFAZ na Pronep.** Memória diz "SEFAZ provada mas desativada, fonte
   passou a ser o Omie". Confirmar: o código existe? O que exatamente o desativou?
2. **Tela de certificado A1** — `api/ConfigCertificadoSefaz/`, `api/shared/certA1.js`,
   `wwwroot/index.html`. Foi construída e o Rafael confirmou que a importação de certificado
   funcionou. Os certificados A1 da Pronep ele já tem em mão.
3. **`DocumentosFiscais`** no SharePoint — existe, e ganhou coluna `ChaveAcesso`. Confirmar se
   há onde guardar o XML (coluna, anexo, ou blob).
4. **Onde o quadro lê hoje** — `api/shared/documentosFiscais.js` e o sync do Omie.

---

## 4. O que vai pronto do autônomo (e pode ser reaproveitado)

### `web/app/nfe-xml.js` — leitor de XML da NF-e
Arquivo isolado de propósito, para o gerador de DANFE usar os mesmos campos que a tela.
**Duas armadilhas já resolvidas e testadas:**

- **`vBC` aparece três vezes** no XML (base do ICMS em `ICMSTot`, base do item em `ICMS00`,
  base do PIS em `PISAliq`). `getElementsByTagName` desce a árvore e devolve o primeiro —
  traria R$ 111,11 ou R$ 999,99 em vez dos R$ 14.400 reais. Solução: busca em **filho direto**.
- **Data com fuso.** Nota emitida 13/08 às 22:40 −03:00 vira 14/08 se passar por `new Date()`.
  Solução: cortar a string, nunca converter.
- XML inválido devolve `null`. `parseFromString` **não lança** — devolve documento com
  `<parsererror>` dentro, e sem checar isso um XML truncado viraria uma nota com todos os
  campos nulos, indistinguível de nota sem dados.

Dois avisos que ele calcula e que existem para evitar pagamento errado:
`cStat ≠ 100` (cancelada/denegada — o XML continua válido e bonito) e **todos os itens com
CFOP de devolução** (devolução é crédito, não conta a pagar).

### Modelo de dados
`documento_xml` em **tabela separada**, não coluna: o quadro lista muitas linhas, e 5–50 KB por
linha viraria megabytes trafegando a cada abertura de tela para um dado que quase nunca é
aberto. A view expõe só `tem_xml` (booleano) e o XML é buscado **no clique**.

A RPC de gravação **recusa o resumo** (`resNFe`): ele também chega pela SEFAZ, também tem a
chave, e aceitar sobrescreveria o completo — perda silenciosa, porque resumo não tem itens nem
impostos, que é exatamente o que a DANFE precisa.

---

## 5. Armadilhas que custaram tempo nesta sessão — não repetir

1. **`active: false` no JSON do n8n.** Guarda que eu criei "para não ligar antes do
   certificado". A razão venceu e virou o defeito: importou parado, o Rafael executou à mão,
   viu sucesso, e ficou **dois dias sem baixar nada**. Execução manual bem-sucedida é
   indistinguível de integração funcionando.

2. **`SEFAZ_AMBIENTE` em variável de container.** Ficou apontando para homologação (que
   devolve 404 — o Ambiente Nacional não publica o `NFeDistribuicaoDFe`) e a correção exigia
   SSH. Agora a fonte de verdade é o **banco**, e o banco **prevalece** sobre a variável.
   *Se a Pronep tiver decisão equivalente em App Setting do SWA, mover para tela.*

3. **Contador somando chamadas, não notas.** A SEFAZ manda o **resumo e o completo da mesma
   nota**, com a mesma chave. Contar chamadas dizia "2 gravados" com 1 linha na tabela.

4. **Interface e pipeline com contratos diferentes.** O botão "Baixar XML" já existia e
   **nunca apareceu**: testava `dados.xml`, e o pipeline gravava `dados.completo` (um
   booleano). Sem erro, sem log — só um botão ausente, que se confunde com "não tem XML".

5. **Não clicar "Execute" duas vezes seguidas** no n8n. Duas consultas em segundos = cStat
   **656 consumo indevido**, e a SEFAZ bloqueia a raiz de CNPJ por ~1h.

---

## 6. Regra de trabalho que vale mais que o código acima

Quatro defeitos desta sessão foram meus, e **três falhavam em silêncio**. Em todos eu validei
pelo lado errado: leitura em vez de escrita, código em vez da tela, "aplicou sem erro" como
prova de efeito.

- Exercitar **a porta que o usuário usa**: se é cadastro, gravar; se é botão, clicar.
- Conferência de migration prova o **efeito**, não que a função existe.
- **Ausência de erro nunca é prova.** Medir o estado depois.
- Sucesso só se afirma quando algo **voltou** confirmando.
- Cache do navegador me fez ler resultado de arquivo antigo duas vezes — recarregar o script
  com URL nova antes de medir.

---

## 7. Estado do autônomo, para referência

No ar e provado: isolamento por unidade na RLS, totalizadores em valor a pagar, parcelamento
com alerta no painel e por e-mail (enviando, 201), SEFAZ em produção com ciclo de 3h, XML sendo
guardado, download e detalhes no cartão. 23 commits no GitHub.

Aberto: **DANFE em PDF** (etapa 3), definir parcelas pela tela de NFs a Pagar, e publicar no
Vercel — que continua sem nenhum projeto.
