/**
 * /api/ListarMembrosGrupo?role=gestor_financeira  — ADMIN ONLY
 *
 * Retorna os integrantes do grupo do Entra correspondente ao role:
 *   { ok, role, label, total, membros: [{ nome, email }] }
 *
 * Usado pela tela Controle de Acessos (modelo granular): ao escolher um grupo,
 * lista as pessoas pra liberar acesso individual.
 */

require('isomorphic-fetch');
const { requireAdmin } = require('../shared/authz');
const { roleParaGrupoId, ROLE_LABELS } = require('../shared/userRoles');
const { getGraphClient } = require('../shared/graph');

module.exports = async function (context, req) {
  try {
    const authz = await requireAdmin(context, req);
    if (!authz) return;

    const role = (req.query && req.query.role) || '';
    const groupId = roleParaGrupoId(role);
    if (!groupId) {
      context.res = { status: 400, body: { error: 'role invalido ou sem grupo correspondente: ' + role } };
      return;
    }

    const client = getGraphClient();
    const membros = [];
    let url = '/groups/' + groupId + '/members?$select=id,displayName,mail,userPrincipalName&$top=100';
    let pages = 0;
    while (url && pages < 20) {
      const resp = await client.api(url).get();
      for (const u of (resp.value || [])) {
        const email = String(u.mail || u.userPrincipalName || '').toLowerCase().trim();
        if (email) membros.push({ nome: u.displayName || email, email: email });
      }
      pages++;
      url = resp['@odata.nextLink']
        ? resp['@odata.nextLink'].replace('https://graph.microsoft.com/v1.0', '')
        : null;
    }
    membros.sort(function (a, b) { return String(a.nome).localeCompare(String(b.nome)); });

    context.res = {
      status: 200,
      headers: { 'Content-Type': 'application/json', 'Cache-Control': 'no-store' },
      body: { ok: true, role: role, label: ROLE_LABELS[role] || role, total: membros.length, membros: membros }
    };
  } catch (err) {
    context.log && context.log.error && context.log.error('ListarMembrosGrupo:', err);
    context.res = { status: 500, headers: { 'Content-Type': 'application/json' },
      body: { error: (err && err.message) || String(err) } };
  }
};
