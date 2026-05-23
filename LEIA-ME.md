# Sistema de Aprovação de NF — Pronep Life Care

Substitui o Power Automate atual `Aprovação de Notas Fiscais de Serviço - Novo` por uma aplicação web hospedada em **Azure Static Web Apps** com autenticação **Entra ID** e armazenamento no **SharePoint** (Graph API), seguindo o mesmo padrão da plataforma Pronep Analytics.

## Estrutura

```
C:\Pronep\Aprovacao_NF\
├── LEIA-ME.md                      Este arquivo
├── GUIA_ENTRA_ID.md                Criar App Registration + 12 grupos
├── GUIA_AZURE_SWA.md               Provisionar Static Web App + deploy
├── GUIA_SHAREPOINT.md              Criar listas Fornecedores/NotasFiscais/Diretorias
├── .env.exemplo                    Template de variáveis (cópia → .env)
├── .gitignore
├── instalar.bat                    Instala Node + SWA CLI + deps
├── 04_deploy_azure.bat             Deploy de wwwroot + api
├── wwwroot/                        Frontend (SPA HTML self-contained)
│   ├── index.html                  Protótipo atual (refatorado para chamar /api)
│   ├── pronep-logo.png
│   ├── staticwebapp.config.json    Rotas, roles, Easy Auth
│   ├── sem-acesso.html
│   └── favicon.*
├── api/                            Azure Functions (Node.js)
│   ├── host.json
│   ├── package.json
│   ├── GetRoles/                   Mapeia grupos AD -> roles
│   ├── ListarNotas/                GET lista paginada (filtros por escopo)
│   ├── PostNota/                   Lança NF (validação + anti-duplicidade + watermark)
│   ├── AprovarNota/                Aprova + watermark + move pasta
│   ├── RejeitarNota/               Rejeita + move pasta + notifica
│   ├── ListarFornecedores/         CRUD de fornecedores
│   └── EnviarNotificacao/          E-mail + Teams card
├── lib/                            Helpers compartilhados (futuro)
├── logs/
└── backups/
```

## Status atual

- [x] **Sprint 0**: Estrutura do projeto criada, padrão Analytics replicado
- [x] **Sprint 0**: Protótipo HTML clicável (validado pelo Rafa)
- [x] **Sprint 0**: Arquitetura técnica documentada (v1.1)
- [x] **Sprint 0**: `GetRoles` funcional (esqueleto pronto, precisa dos GUIDs reais)
- [ ] **Sprint 0**: Criar App Registration "Pronep Aprovacao NF SWA" → preencher MS_CLIENT_ID
- [ ] **Sprint 0**: Criar 12 grupos no Entra ID → atualizar `api/GetRoles/index.js`
- [ ] **Sprint 0**: Criar Static Web App `aprovacao-nf-pronep` → pegar token
- [ ] **Sprint 0**: Primeiro deploy (só do HTML — funções esqueleto)
- [ ] **Sprint 1**: Criar listas SharePoint (Fornecedores, NotasFiscais, Diretorias) e popular com matriz do Anexo A
- [ ] **Sprint 1**: Implementar `ListarFornecedores` e `PostFornecedor` (CRUD real)
- [ ] **Sprint 2**: Implementar `PostNota` com anti-duplicidade
- [ ] **Sprint 3**: Implementar `AprovarNota` / `RejeitarNota` com watermark
- [ ] **Sprint 4**: Implementar `EnviarNotificacao` + dashboards reais
- [ ] **Sprint 5**: UAT com Bruno (Suprimentos) e Vitor (Técnica RJ)
- [ ] **Go-Live**: Cutover do Power Automate

Veja o documento de arquitetura completo em `outputs/Pronep_Aprovacao_NF_Arquitetura_v11.pdf`.

## Primeiro deploy

Antes de qualquer linha de código real, validar a infra com um deploy do **HTML puro + GetRoles esqueleto**. Ordem:

1. **`instalar.bat`** — instala Node, npm e SWA CLI
2. **Seguir GUIA_ENTRA_ID.md** — criar App Registration + 12 grupos
3. **Seguir GUIA_AZURE_SWA.md** — criar SWA, pegar token de deploy
4. **Editar `api/GetRoles/index.js`** — colar os 12 GUIDs reais dos grupos
5. **Editar `.env`** — preencher tokens e IDs
6. **`04_deploy_azure.bat`** — deploy

Resultado esperado: ao acessar `https://aprovacao-nf-pronep.azurestaticapps.net`, o sistema pede login Microsoft. Após login, o protótipo renderiza (mesmo sem dados reais ainda — Sprint 1+ resolve isso).

## Convenção de nomes (Entra ID)

Mantemos o mesmo prefixo do Analytics:

| Recurso | Nome |
|---|---|
| App Registration | `Pronep Aprovacao NF SWA` |
| Resource Group Azure | `rg-pronep-aprovacao-nf` (ou reaproveitar `rg-pronep-analytics`) |
| Static Web App | `aprovacao-nf-pronep` |
| Site SharePoint | `Aprovacao-NotasFiscaisServicos` (já existe) |
| Grupos de segurança | `PRONEP-NF-*` (12 grupos — ver GUIA_ENTRA_ID.md) |
| Caixa institucional | `nf-aprovacoes@pronep.com.br` |

## Quem é quem (matriz de aprovadores extraída do PA atual)

| Diretoria | SP | RJ | ES |
|---|---|---|---|
| Suprimentos | Bruno Hioka | Bruno Hioka | Bruno Hioka |
| Técnica | **Nicolle Boas** | Vitor Amaral | Vitor Amaral |
| Financeira | Henrico Molina | Henrico Molina | Henrico Molina |
| RH-DP / Fiscal-Contábil | Janilene Santos | Janilene Santos | Janilene Santos |
| Tecnologia | Rafael Machado | Rafael Machado | Rafael Machado |
| Jurídica / Administrativa | Rafaella Santos | Rafaella Santos | Rafaella Santos |
| Qualidade | Sabrina Fernandes | Sabrina Fernandes | Sabrina Fernandes |

## Contato

- Sponsor: Rafael Machado (Tecnologia)
- Time financeiro: Sandra Ferreira, Vitor Costa, Monica Pires, Bruno Mendes
