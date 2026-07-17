/**
 * /api/SalvarDiretoria (POST) — ADMIN ONLY.
 *
 * Atualiza o APROVADOR de uma linha da lista PRONEP-NF-Diretorias (Unidade x Diretoria).
 * Grava field_3 (email) e field_4 (nome). Fonte de verdade do roteamento de NFs.
 *
 * OBS: afeta apenas NFs FUTURAS. As NFs ja pendentes guardam o AprovadorAtual do momento
 * da criacao (nao sao reescritas aqui) — comportamento definido com o negocio.
 *
 * Body: { id, email: "nome@pronep.com.br", nome: "Nome do aprovador", telefone?: "+5521999998888" }
 *   - telefone (opcional) grava em TelefoneNotificacao (WhatsApp da automacao de NF por e-mail);
 *     string vazia limpa o campo. Formato E.164 (+ e 10-15 digitos).
 */

require('isomorphic-fetch');
const { resolveAuthz } = require('../shared/authz');
const { getGraphClient, resolveSiteAndList } = require('../shared/graph');

const LIST_NAME = 'PRONEP-NF-Diretorias';

module.exports = async function (context, req) {
  try {
    // Admin OU TI podem alterar aprovadores.
    const authz = await resolveAuthz(req);
    if (!authz) { context.res = { status: 401, headers: { 'Content-Type': 'application/json' }, body: { error: 'Nao autenticado' } }; return; }
    const podeEditar = authz.isAdmin || (authz.roles || []).indexOf('ti') >= 0;
    if (!podeEditar) { context.res = { status: 403, headers: { 'Content-Type': 'application/json' }, body: { error: 'Acesso restrito a Admin ou TI' } }; return; }

    const body = req.body || {};
    const id = String(body.id || '').trim();
    const email = String(body.email || '').trim().toLowerCase();
    const nome = String(body.nome || '').trim();
    const telefone = String(body.telefone || '').trim();

    if (!id) { context.res = { status: 400, headers: { 'Content-Type': 'application/json' }, body: { error: 'id obrigatorio' } }; return; }
    if (!email || !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) {
      context.res = { status: 400, headers: { 'Content-Type': 'application/json' }, body: { error: 'email invalido' } }; return;
    }
    if (!/@pronep\.com\.br$/i.test(email)) {
      context.res = { status: 400, headers: { 'Content-Type': 'application/json' }, body: { error: 'O e-mail do aprovador deve ser @pronep.com.br' } }; return;
    }
    // telefone opcional; se informado, precisa ser E.164 (+ e 10-15 digitos). Vazio limpa o campo.
    if (telefone && !/^\+\d{10,15}$/.test(telefone)) {
      context.res = { status: 400, headers: { 'Content-Type': 'application/json' }, body: { error: 'Telefone deve estar em E.164, ex.: +5521999998888' } }; return;
    }

    const client = getGraphClient();
    const { siteId, listId } = await resolveSiteAndList(client, LIST_NAME);

    // field_3 = Email do aprovador · field_4 = Nome do aprovador · TelefoneNotificacao = WhatsApp
    await client.api('/sites/' + siteId + '/lists/' + listId + '/items/' + id + '/fields')
      .patch({ field_3: email, field_4: nome, TelefoneNotificacao: telefone });

    context.res = { status: 200, headers: { 'Content-Type': 'application/json', 'Cache-Control': 'no-store' },
      body: { ok: true, id: id, email: email, nome: nome, telefone: telefone, atualizadoPor: authz.email } };
  } catch (err) {
    context.log && context.log.error && context.log.error('SalvarDiretoria:', err);
    context.res = { status: 500, headers: { 'Content-Type': 'application/json' },
      body: { error: (err && err.message) || String(err) } };
  }
};
