/**
 * /api/GetControleAcessosTelas (GET) — ADMIN ONLY
 *
 * Dados da aba "Telas do sistema" da Central de Controle de Acessos:
 *   {
 *     telas:   [{ id:'aprovadas', label:'Notas Aprovadas' }, ...]  // telas configuraveis
 *     grupos:  [{ role:'gestor_fiscal_contabil', label:'Fiscal-Contábil' }, ...] // grupos Entra
 *     acessos: { 'aprovadas': ['gestor_fiscal_contabil'], ... }    // mapa atual (tela -> tokens)
 *   }
 *
 * Espelha GetControleAcessos (contratos), mas o alvo e uma TELA em vez de uma pasta.
 */

require('isomorphic-fetch');
const { requireAdmin } = require('../shared/authz');
const { lerMapaTelas, TELAS } = require('../shared/acessoTelas');
const { gruposContrato } = require('../shared/userRoles');
const { getGraphClient, resolveSiteId } = require('../shared/graph');

module.exports = async function (context, req) {
  try {
    const authz = await requireAdmin(context, req);
    if (!authz) return;

    const client = getGraphClient();
    const siteId = await resolveSiteId(client);
    const acessos = await lerMapaTelas(client, siteId, null);
    // Grupos liberaveis nesta aba: alem dos grupos de diretoria (gruposContrato),
    // tambem Submetedores e TI — fazem sentido pra liberar TELAS (nao pra contratos).
    // roleParaGrupoId resolve esses roles (estao no GROUP_TO_ROLE), entao a listagem
    // de membros (ListarMembrosGrupo) funciona normalmente.
    const extras = [
      { role: 'submitter', label: 'Submetedores' },
      { role: 'ti',        label: 'TI / Suporte' }
    ];
    const grupos = extras.concat(gruposContrato());

    context.res = {
      status: 200,
      headers: { 'Content-Type': 'application/json', 'Cache-Control': 'no-store' },
      body: { ok: true, telas: TELAS, grupos: grupos, acessos: acessos }
    };
  } catch (err) {
    context.log && context.log.error && context.log.error('GetControleAcessosTelas:', err);
    context.res = { status: 500, headers: { 'Content-Type': 'application/json' },
      body: { error: (err && err.message) || String(err) } };
  }
};
