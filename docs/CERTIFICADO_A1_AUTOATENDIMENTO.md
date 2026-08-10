# Trocar o certificado A1 da SEFAZ pela tela

O A1 vence a cada 12 meses. Antes, renovar exigia subir o blob no portal do Azure e
reescrever a App Setting da senha — dependia de quem tem acesso ao Azure. Agora está
em **Configurações › 🔑 Certificado da SEFAZ (A1)** (só administrador).

## O único passo de configuração (uma vez)

Falta uma App Setting nova. Gere a chave **na sua máquina**:

```bash
node -e "console.log(require('crypto').randomBytes(32).toString('base64'))"
```

Cole o valor direto no portal do Azure, em **Configuration › Application settings**,
com o nome `CONFIG_CRYPTO_KEY`. Não passe esse valor por chat, e-mail ou anotação: é
a chave que protege a senha do certificado.

Dois cuidados que já nos custaram tempo neste projeto:

- **Recarregue a página do portal antes de editar.** Salvar pelo portal regrava o
  conjunto inteiro de App Settings e reverte o que foi feito por CLI.
- Se preferir a CLI, use `-o none` para o valor não ficar no histórico do terminal.

Enquanto essa App Setting não existir, a tela aparece com o botão Enviar desativado e
um aviso dizendo exatamente isso. Nada quebra — o caminho antigo continua valendo.

## Como usar

Um card por CNPJ de `SEFAZ_CNPJS`, com a validade em destaque (verde acima de 60 dias,
laranja até 60, vermelho até 30, e "VENCIDO" com a contagem). Escolha o `.pfx`, digite
a senha, Enviar.

- **Trocar só a senha:** deixe o arquivo vazio. A senha é conferida contra o
  certificado já gravado antes de substituir.
- **Filiais da mesma raiz:** a SEFAZ autentica pela raiz do CNPJ (8 primeiros
  dígitos), então o mesmo arquivo serve para todas. Quando há mais de uma filial na
  raiz, aparece marcada a opção de aplicar a todas de uma vez.

## O que é conferido antes de gravar

O par (arquivo, senha) é aberto de verdade no servidor antes de qualquer gravação.
Então o caso comum de erro nunca chega a escrever nada:

| Situação | O que a tela diz |
|---|---|
| Senha errada | "Senha do certificado incorreta." |
| Arquivo que não é `.pfx` | Barrado no navegador, sem subir os 9 KB |
| `.pfx` no formato antigo (RC2/40 bits) | Explica que é o formato e **dá o comando de reexportação** |
| Arquivo grande demais | Diz que um A1 tem ~9 KB |

O formato antigo é o caso que mais engana: o OpenSSL 3 só diz
`Unsupported PKCS12 PFX data`, que não é senha errada nem arquivo corrompido. Se
acontecer, reexporte:

```bash
openssl pkcs12 -in antigo.pfx -out novo.pfx -export -keypbe AES-256-CBC -certpbe AES-256-CBC -macalg sha256
```

## Onde as coisas ficam

Três blobs por CNPJ, no container privado de `SEFAZ_CERT_STORAGE`:

| Blob | Conteúdo |
|---|---|
| `<cnpj>.pfx` | o certificado |
| `<cnpj>.senha` | a senha cifrada em AES-256-GCM com `CONFIG_CRYPTO_KEY` |
| `<cnpj>.meta.json` | titular, validade, tamanho, quem trocou e quando |

**A senha nunca volta para o navegador** — nem mascarada. A tela mostra apenas *se*
existe senha gravada, e o campo começa sempre vazio. A auditoria registra a troca sem
a senha. Sendo honesto sobre o ganho da cifra: a chave da conta de Storage está numa
App Setting, então quem tem o portal alcança o blob; cifrar com uma chave separada faz
com que ler o blob não baste, e evita senha em texto puro em repouso e em backup.

## Convivência com o jeito antigo

A senha em `SEFAZ_CERT_<cnpj>_SENHA` continua funcionando. Quando existe senha no blob
**e** o certificado também vem do blob, a do blob vence — senão, depois de uma
renovação pela tela, a App Setting antiga seria usada contra o arquivo novo e daria
"mac verify failure" num certificado perfeitamente válido. Quando o `.pfx` vem da App
Setting (caminho legado), a senha da App Setting continua valendo: certificado e senha
ficam sempre da mesma geração.

## Verificadores

```bash
node scripts/verificar-cert-a1.js && node scripts/verificar-segredos.js && node scripts/verificar-config-cert.js && node scripts/verificar-front.js
```

Os dois primeiros geram certificados descartáveis com openssl e apagam no fim; nenhum
certificado da Pronep é lido. O de endpoint dubla só o que sai da máquina (Blob, Graph)
e roda a cifra de verdade — se a senha aparecesse em texto puro no que vai para o
Storage, ele acusaria.

## Como a validade é lida sem parser de PKCS#12

O Node não abre PKCS#12 (`new X509Certificate(pfx)` não existe), mas sabe **usar** um
pfx num handshake TLS. O servidor abre um TLS em `127.0.0.1` e conecta nele mesmo; o
`socket.getCertificate()` devolve o certificado local, com as datas. Nada sai da
máquina e não entra dependência nova. Se esse handshake falhar por qualquer motivo, o
upload passa mesmo assim, com o aviso de que a validade não foi lida — bloquear a troca
de um certificado por causa de uma data seria trocar um problema pequeno por um grande.
