/**
 * shared/contratos.js — Modulo de Controle de Contratos da Pronep.
 *
 * Crawler recursivo do SharePoint "CONTRATOS-SERVICOS-CONTRATOS" + extracao
 * de vigencia (DataInicio/DataFim) via Claude Haiku com escalacao pra Sonnet
 * quando a confianca for baixa. Persiste em PRONEP-NF-Contratos.
 *
 * Estrutura do SP:
 *   /Shared Documents/CONTRATOS/CONTRATOS E DOCUMENTOS - PRESTADORES/
 *     <Diretoria>/
 *       <Unidade ou subpasta>/
 *         <Prestador>/
 *           contrato.pdf, aditivos, etc.
 *
 * Mapeamento de diretorias do SP -> diretorias do sistema NF:
 *   - DIRETORIA COMERCIAL       -> Comercial
 *   - DIRETORIA DE OPERACOES    -> Operacoes
 *   - DIRETORIA DE SUPRIMENTOS  -> Suprimentos
 *   - DIRETORIA FINANCEIRA      -> Financeira
 *   - GERENCIA DE PROJETOS E TI -> Tecnica  (TI esta sob Tecnica no NF)
 *   - GERENCIA DE RH            -> RH
 *   - JURIDICO                  -> Juridica
 *   - OUVIDORIA                 -> Ouvidoria
 *   - PACIENTES PARTICULARES    -> Particulares
 *   - QUALIDADE                 -> Qualidade
 *
 * Env vars necessarias:
 *   SHAREPOINT_CONTRATOS_HOSTNAME (default: pronepadmin.sharepoint.com)
 *   SHAREPOINT_CONTRATOS_PATH     (default: /sites/CONTRATOS-SERVICOS-CONTRATOS)
 *   ANTHROPIC_API_KEY             (ja usado pelo SAN)
 */

require('isomorphic-fetch');
const { ClientSecretCredential } = require('@azure/identity');
const { Client } = require('@microsoft/microsoft-graph-client');
const { TokenCredentialAuthenticationProvider } =
  require('@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials');
const Anthropic = require('@anthropic-ai/sdk');
const pdfParse = require('pdf-parse');

const LIST_CONTRATOS = 'PRONEP-NF-Contratos';
const ROOT_FOLDER_PATH = '/CONTRATOS/CONTRATOS E DOCUMENTOS - PRESTADORES';

// Cache de site/list pra economizar chamadas
const cache = {
  contratoSite: null,
  contratoListId: null,
  contratoColMap: null,
  driveId: null
};
