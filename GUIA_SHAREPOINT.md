# Guia — SharePoint Online

Pré-requisito: ter completado `GUIA_ENTRA_ID.md` + `GUIA_AZURE_SWA.md`.

## Site de trabalho

Reaproveitamos o site existente: `https://pronepadmin.sharepoint.com/sites/Aprovacao-NotasFiscaisServicos`

## Estrutura de pastas (Biblioteca `Documentos`)

Já existe a estrutura criada pelo Power Automate atual:

```
Notas Fiscais/
├── Notas Pendentes/
│   ├── SP/
│   │   ├── Diretoria Suprimentos/
│   │   ├── Diretoria Técnica/
│   │   ├── Diretoria Financeira/
│   │   ├── Diretoria RH-DP/
│   │   ├── Diretoria Tecnologia/
│   │   ├── Diretoria Fiscal-Contábil/
│   │   ├── Diretoria Jurídica/
│   │   ├── Diretoria Administrativa/
│   │   └── Diretoria Qualidade/
│   ├── RJ/  (mesmas 9 diretorias)
│   └── ES/  (mesmas 9 diretorias)
├── Notas Aprovadas/
│   ├── SP/{AAAA-MM-DD}/   (subpastas diárias criadas dinamicamente)
│   ├── RJ/{AAAA-MM-DD}/
│   └── ES/{AAAA-MM-DD}/
└── Notas Rejeitadas/
    ├── SP/  (9 diretorias)
    ├── RJ/
    └── ES/
```

Não precisa criar — já está pronta. O novo sistema apenas reaproveita.

## Listas SharePoint (a criar)

3 listas precisam ser criadas no mesmo site, na primeira sprint:

### Lista 1: `Fornecedores`

| Coluna | Tipo | Obrigatório | Observação |
|---|---|---|---|
| CNPJ | Linha única de texto | Sim | **Chave única** (validar índice) |
| RazaoSocial | Linha única de texto | Sim | |
| NomeFantasia | Linha única de texto | Não | |
| UnidadeAtendimento | Escolha (SP, RJ, ES) | Sim | |
| DiretoriaPadrao | Linha única de texto | Sim | Validar contra lista de 9 diretorias |
| Categoria | Escolha | Sim | Serviços, Materiais, Locação, Reembolso, Medicamentos, Outros |
| Ativo | Sim/Não | Sim | Default: Sim |

### Lista 2: `NotasFiscais`

Ver modelo completo no doc de arquitetura, seção 4.2. Colunas críticas:
- ChaveAcesso (44 chars) — indexada
- HashArquivo (64 chars) — indexada
- Numero, Serie, Fornecedor (lookup), Unidade, Diretoria, Valor, Vencimento
- Status (PENDENTE, APROVADA, REJEITADA, AUTO_REJEITADA)
- MotivoRejeicao
- LancadoPor, LancadoEm, AprovadoPor, AprovadoEm
- ForaPrazo (Sim/Não), NegociouCom, NegociouComEmail
- CaminhoSharePoint

### Lista 3: `Diretorias`

Matriz 27 linhas (3 unidades × 9 diretorias) — popular com o Anexo A do documento de arquitetura.

| Unidade | Diretoria | GrupoEntraID | Email aprovador atual |
|---|---|---|---|
| SP | Suprimentos | <obj-id PRONEP-NF-Gestor-Suprimentos> | bruno.hioka@pronep.com.br |
| SP | Técnica | <obj-id PRONEP-NF-Gestor-Tecnica-SP> | nicolle.boas@pronep.com.br |
| ... (mais 25 linhas) | | | |

## Permissão Sites.Selected (Etapa adicional)

Como usamos `Sites.Selected` (mais restritivo que `Sites.ReadWrite.All`), precisamos autorizar explicitamente o acesso ao site.

### Via Graph Explorer

1. Abre https://developer.microsoft.com/en-us/graph/graph-explorer
2. Login com conta admin Pronep
3. Conceda permissão **Sites.FullControl.All** (apenas temporariamente, pra conseguir gerenciar)
4. Executa:
   ```
   GET https://graph.microsoft.com/v1.0/sites/pronepadmin.sharepoint.com:/sites/Aprovacao-NotasFiscaisServicos
   ```
   Anota o `id` do site (formato `pronepadmin.sharepoint.com,<guid>,<guid>`)
    312140d5-0fdc-4284-80d5-2a2dfe110f1c

5. Concede acesso à App Reg:
   ```
   POST https://graph.microsoft.com/v1.0/sites/<site-id>/permissions
   {
     "roles": ["write"],
     "grantedToIdentities": [{
       "application": {
         "id": "<85f7fa68-a241-4caf-803f-9991bd1f0eee>",
         "displayName": "Pronep Aprovacao NF SWA"
       }
     }]
   }
   ```

### Via PowerShell (alternativa)

```powershell
Connect-PnPOnline -Url "https://pronepadmin.sharepoint.com/sites/Aprovacao-NotasFiscaisServicos" -Interactive

Grant-PnPAzureADAppSitePermission `
  -AppId "<CLIENT_ID>" `
  -DisplayName "Pronep Aprovacao NF SWA" `
  -Site "https://pronepadmin.sharepoint.com/sites/Aprovacao-NotasFiscaisServicos" `
  -Permissions Write
```

## Migração dos dados atuais

Quando ligar o sistema novo:
- **Notas Aprovadas históricas**: ficam onde estão. O sistema novo lê delas via Graph.
- **Notas Pendentes existentes**: 2 opções:
  1. Esvaziar a pasta antes do cutover (pedir pros gestores aprovarem ou rejeitarem tudo)
  2. Migrar via script (`lib/migrar_pendentes.js` — TODO Sprint 5)
- **Fornecedores**: importar do controle atual (Excel, Power BI, etc.) usando a feature de import da tela de Fornecedores

## Próximas etapas

- [ ] Criar as 3 listas SharePoint
- [ ] Popular `Diretorias` com a matriz do Anexo A
- [ ] Conceder permissão Sites.Selected
- [ ] Implementar `ListarFornecedores` consumindo Graph (Sprint 1)
