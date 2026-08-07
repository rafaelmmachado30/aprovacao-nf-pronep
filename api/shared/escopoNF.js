/**
 * Escopo de visao do quadro "NFs a Pagar" — quem enxerga qual card.
 *
 * ISTO E CONTROLE DE ACESSO E POR ISSO VIVE NO SERVIDOR. Filtrar no front seria
 * so esconder: bastaria chamar a API direto para ver as notas de todas as
 * diretorias. O que nao e devolvido nao existe para quem chamou.
 *
 * MORA AQUI, E NAO DENTRO DE UM ENDPOINT, porque agora sao dois os consumidores
 * (a lista do quadro e o detalhe da NF) e vao ser mais. Regra de visibilidade
 * duplicada e regra que diverge: um dia alguem corrige um lado, esquece o outro,
 * e o buraco fica aberto justamente no endpoint que ninguem lembra que existe.
 *
 * Regras:
 *   admin e financeiro       veem tudo (o financeiro paga; precisa da visao completa)
 *   gestor                   ve as diretorias que ELE aprova, conforme a lista
 *                            PRONEP-NF-Diretorias (Unidade x Diretoria -> e-mail)
 *   cadastro incompleto      visivel para TODOS, de proposito: e um card que ainda
 *                            nao tem diretoria, entao nao ha a quem pertencer, e
 *                            alguem precisa poder corrigir o cadastro. Esconder
 *                            deixaria a nota travada sem ninguem responsavel.
 */

const { resolveAuthz } = require('./authz');
const { resolveListId } = require('./documentosFiscais');

async function todosItens(client, siteId, listId, maxPaginas) {
  const out = [];
  let url = '/sites/' + siteId + '/lists/' + listId + '/items?expand=fields&$top=999';
  let p = 0;
  while (url && p < (maxPaginas || 20)) {
    const r = await client.api(url).get();
    out.push.apply(out, r.value || []);
    p++;
    const nl = r['@odata.nextLink'];
    url = nl ? nl.replace('https://graph.microsoft.com/v1.0', '') : null;
  }
  return out;
}

async function montarEscopo(client, siteId, req, verComo) {
  let authz = null;
  try { authz = await resolveAuthz(req); } catch (e) { /* segue restrito */ }

  if (!authz) return { autenticado: false, verTudo: false, pares: {}, email: '' };
  if (authz.isAdmin || authz.isFinanceiro) {
    /* LENTE "Visualizando como". Honrada SO para quem ja ve tudo, e so para
       RESTRINGIR — nunca para ampliar. Estreitar a propria visao nao e escalada
       de privilegio, entao aceitar o parametro aqui e seguro; aceitar de um
       gestor para ALARGAR seria o oposto, e por isso nem e considerado.
       Sem isto o seletor da tela nao teria efeito nenhum: o servidor olha a
       identidade real e devolveria tudo, dando a impressao de que o escopo esta
       quebrado quando na verdade ele nunca foi consultado. */
    if (verComo) {
      return { autenticado: true, verTudo: false, motivo: 'lente',
               lente: verComo, diretorias: [verComo],
               pares: { ['TODAS|' + verComo]: true }, email: authz.email };
    }
    return { autenticado: true, verTudo: true, motivo: authz.isAdmin ? 'admin' : 'financeiro',
             pares: {}, email: authz.email };
  }

  const pares = {};
  const diretorias = [];
  try {
    const idDir = await resolveListId(client, siteId, 'PRONEP-NF-Diretorias');
    if (idDir) {
      for (const it of await todosItens(client, siteId, idDir, 5)) {
        const f = it.fields || {};
        /* field_1=Unidade, field_2=Diretoria, field_3=Email do aprovador
           (mesmo mapa de ListarDiretorias — se um mudar, os dois mudam). */
        const email = String(f.field_3 || '').toLowerCase().trim();
        if (!email || email !== authz.email) continue;
        const un = String(f.field_1 || '').trim();
        const dir = String(f.field_2 || '').trim();
        if (!dir) continue;
        pares[un + '|' + dir] = true;
        if (diretorias.indexOf(dir) < 0) diretorias.push(dir);
      }
    }
  } catch (e) { /* sem mapa, o gestor fica sem par nenhum — restritivo por padrao */ }

  return { autenticado: true, verTudo: false, pares: pares, diretorias: diretorias,
           email: authz.email };
}

/* Decide se o card entra na visao deste usuario. */
function podeVer(card, escopo) {
  if (escopo.verTudo) return true;
  /* Sem diretoria definida nao ha dono: fica visivel para quem puder consertar. */
  if (!card.fornecedorCadastrado || card.cadastroIncompleto || !card.diretoria) return true;
  /* Aprovador de "TODAS" as unidades cobre qualquer unidade daquela diretoria. */
  return !!(escopo.pares[(card.unidade || '') + '|' + card.diretoria] ||
            escopo.pares['TODAS|' + card.diretoria] ||
            escopo.pares['|' + card.diretoria]);
}

module.exports = { montarEscopo, podeVer, todosItens };
