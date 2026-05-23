# Guia — Configuração do Entra ID

## Etapa 1 — Criar App Registration `Pronep Aprovacao NF SWA` (5 min)

Vamos criar uma App Registration **separada** da do Analytics. Mantemos o tenant `4b30645b-0888-45c0-9481-712bde435ffd`.

1. Acessa https://entra.microsoft.com → **Aplicações** → **Registros de aplicativos** → **+ Novo registro**
2. Preenche:
   - **Name**: `Pronep Aprovacao NF SWA`
   - **Supported account types**: "Apenas contas neste diretório organizacional"
   - **Redirect URI**: deixa em branco (preencher depois que o SWA existir)
3. **Register**
4. Na tela de Visão geral, anota:
   - **Application (client) ID** → vai virar `AAD_CLIENT_ID` no SWA
   - **Directory (tenant) ID** → confirma que é `4b30645b-0888-45c0-9481-712bde435ffd`

### Criar client secret

5. Menu esquerdo → **Certificados e segredos**
6. **+ Novo segredo do cliente**
7. Descrição: `SWA Auth`, expiração: `730 dias (2 anos)`
8. **Adicionar**
9. **COPIE O VALOR IMEDIATAMENTE** (depois de sair, não dá mais) — salva como `AAD_CLIENT_SECRET` no arquivo `.env` local (que está no `.gitignore`, então não vaza pro git).

> ⚠️ **Segurança**: nunca cole o valor real do client secret neste arquivo .md. O `.env` é o único lugar local seguro. Em produção, ele fica no App Settings do SWA no Azure Portal.

### Configurar claim de grupos

Para o `GetRoles` receber os grupos do usuário, configurar a App Registration para emitir o claim "groups":

10. Menu esquerdo → **Configuração de Token (Token configuration)**
11. **+ Adicionar declaração de grupos** → **Grupos de segurança**
12. Para ID, Access e SAML, marca: **ID do grupo**
13. **Salvar**

## Etapa 2 — Criar os 12 grupos no Entra ID

No portal Entra ID → **Grupos** → **+ Novo grupo**. Para cada grupo:

- **Tipo**: Segurança
- **Tipo de associação**: Atribuído

| # | Nome do grupo | Descrição | Membros iniciais |
|---|---|---|---|
| 1 | `PRONEP-NF-Submitter` | Colaboradores que podem lançar NF | Todos do operacional (RH-DP pode usar regra dinâmica via "Tipo de associação: Dinâmico Usuário") |
| 2 | `PRONEP-NF-Admin` | Administradores do sistema | Rafael Machado |
| 3 | `PRONEP-Financeiro-Gestao` | Gestores do Financeiro (negociação D+5) | Sandra Ferreira, Vitor Costa, Monica Pires, Bruno Mendes |
| 4 | `PRONEP-NF-Gestor-Suprimentos` | Aprovador Suprimentos (SP/RJ/ES) | Bruno Hioka |
| 5 | `PRONEP-NF-Gestor-Financeira` | Aprovador Financeira (SP/RJ/ES) | Henrico Molina |
| 6 | `PRONEP-NF-Gestor-Tecnologia` | Aprovador Tecnologia (SP/RJ/ES) | Rafael Machado |
| 7 | `PRONEP-NF-Gestor-Qualidade` | Aprovador Qualidade (SP/RJ/ES) | Sabrina Fernandes |
| 8 | `PRONEP-NF-Gestor-RH-DP` | Aprovador RH-DP (SP/RJ/ES) | Janilene Santos |
| 9 | `PRONEP-NF-Gestor-Fiscal-Contabil` | Aprovador Fiscal-Contábil (SP/RJ/ES) | Janilene Santos |
| 10 | `PRONEP-NF-Gestor-Juridica` | Aprovador Jurídica (SP/RJ/ES) | Rafaella Santos |
| 11 | `PRONEP-NF-Gestor-Administrativa` | Aprovador Administrativa (SP/RJ/ES) | Rafaella Santos |
| 12 | `PRONEP-NF-Gestor-Tecnica-SP` | Aprovador Técnica apenas SP | Nicolle Boas |
| 13 | `PRONEP-NF-Gestor-Tecnica-RJES` | Aprovador Técnica RJ + ES | Vitor Amaral |

Para cada grupo, **copie o "Object Id" (GUID)** logo após criar — você vai colar no `api/GetRoles/index.js`.

## Etapa 3 — Colar os GUIDs no GetRoles

Abre `api/GetRoles/index.js` e substitui os placeholders no objeto `GROUP_TO_ROLE`:

```javascript
const GROUP_TO_ROLE = {
  // Cole aqui os Object Ids dos grupos criados acima
  '01d540d1-8596-42d0-9a20-de5c361c7c96':       'submitter',
  '480a1595-bdc3-492a-9ef2-317f148a237e':           'administrador',
  'c2a73d16-4659-4b3c-93a1-0c0fbfaaaa96':  'financeiro_gestao',
  '2d9f5bcf-2ae0-494e-957b-a1c69016664d': 'suprimentos_nf',
  '4f28f31b-b704-4615-961b-a9ca0898cea8': 'administrativa_nf',
  '5aa9fc6b-900d-40eb-861d-8bbf72499da1': 'juridica_nf',
  '6b77405b-ba89-47ee-af21-58ec19bb3ff7': 'financeira_nf',
  'a6711877-8746-4ca5-a955-c15980c7e90d': 'qualidade_nf',
  'a7826b5c-7c29-4a24-836b-a7432aa941ec': 'tecnologia_nf',
  '13a544d8-3dde-4820-9695-c492e58a2782': 'rh-dp_nf',
  '334eb19b-c138-4551-8e45-a36ca4e32e48': 'tecnica-rjes_nf',
  'b9272d98-3e26-4e2d-aae6-ff9057f57e5c': 'fiscal-contabil_nf',
  'fc3a375b-329c-4d9c-81be-06180f0598af': 'tecnica-sp_nf'
  // ... e os 10 gestores
};
```

A ordem do objeto não importa — o que importa é o GUID correto à esquerda.

## Etapa 4 — Confirmar tenant ID

Confirma que o `wwwroot/staticwebapp.config.json` está apontando pro tenant correto:

```json
"openIdIssuer": "https://login.microsoftonline.com/4b30645b-0888-45c0-9481-712bde435ffd/v2.0"
```

(É o mesmo do Analytics — não precisa mexer se mantivermos o tenant.)

## Etapa 5 — Próximo passo

Depois de criados os 13 grupos + App Registration, segue para `GUIA_AZURE_SWA.md` para provisionar o Static Web App e fazer o primeiro deploy.

## Diagnóstico

| Sintoma | Causa | Solução |
|---|---|---|
| Login OK mas `GetRoles` devolve `roles: []` | Grupos não estão sendo emitidos no token | Conferir Etapa 1.10-13 (Configuração de Token) |
| Login OK e GetRoles devolve só 1-2 roles | Membros não foram adicionados ao grupo | Conferir Etapa 2 (membros iniciais) |
| `AADSTS50011` | Redirect URI da App Reg não bate | Confere se está `https://aprovacao-nf-pronep.azurestaticapps.net/.auth/login/aad/callback` |
| Tela 403 sem-acesso | Usuário não está em nenhum grupo | Adicionar ao `PRONEP-NF-Submitter` para acesso básico |
