/**
 * shared/financeiroGestao.js
 *
 * E-mails do grupo Entra ID PRONEP-NF-Financeiro-Gestao — os gestores do financeiro
 * autorizados a CONFIRMAR o alinhamento de NFs vencendo em < D+5 (tela Confirma
 * Alinhamento). Best-effort: se a leitura do grupo falhar, devolve lista vazia.
 *
 * App Setting opcional: GESTOR_FINANCEIRO_GROUP_ID (OID do grupo).
 */

require('isomorphic-fetch');

const DEFAULT_GROUP_ID = 'c2a73d16-4659-4b3c-93a1-0c0fbfaaaa96'; // PRONEP-NF-Financeiro-Gestao

async function emailsFinanceiroGestao(client) {
  try {
    const groupId = process.env.GESTOR_FINANCEIRO_GROUP_ID || DEFAULT_GROUP_ID;
    const resp = await client.api('/groups/' + groupId + '/members')
      .select('id,displayName,mail,userPrincipalName').top(50).get();
    return (resp.value || [])
      .map(function (u) { return String(u.mail || u.userPrincipalName || '').toLowerCase(); })
      .filter(Boolean);
  } catch (e) {
    return [];
  }
}

module.exports = { emailsFinanceiroGestao, GRUPO_FINANCEIRO_GESTAO_ID: DEFAULT_GROUP_ID };
