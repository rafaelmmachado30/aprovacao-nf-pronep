/**
 * /api/SalvarControleAcessosTelas (POST) — ADMIN ONLY
 *
 * Body (por tela):     { tela: 'aprovadas', tokens: ['gestor_fiscal_contabil','a@x'] }
 * Body (mapa inteiro): { mapa: { 'aprovadas': ['gestor_fiscal_contabil'], ... } }
 *
 * Salva no item Config 'acessoTelas'. Modelo ADITIVO: definir tokens numa tela
 * concede acesso ALEM do que o papel ja da. Tela com tokens=[] (lista vazia) = remove
 * a config daquela tela (volta a so o canSee do papel).
 *
 * So aceita telas da lista canonica (TELAS_IDS) — telas administrativas nao entram.
 */

require('isomorphic-fetch');
const { requireAdmin } = require('../shared/authz');
const { lerMapaTelas, salvarMapaTelas, resolveConfigListId, TELAS_IDS } = require('../shared/acessoTelas');
const { registrar: auditRegistrar } = require('../shared/auditLog');
const { getGraphClient, resolveSiteId } = require('../shared/graph');

// Normaliza tokens (role ou email): trim, sem vazios, sem duplicatas (case-insensitive).
function normTokens(arr) {
  const out = [];
  const seen = new Set();
  for (const t of (Array.isArray(arr) ? arr : [])) {
    const v = String(t || '').trim();
    const k = v.toLowerCase();
    if (v && !seen.has(k)) { seen.add(k); out.push(v); }
  }
  return out;
}

module.exports = async function (context, req) {
  try {
    const authz = await requireAdmin(context, req);
    if (!authz) return;

    const body = req.body || {};
    const client = getGraphClient();
    const siteId = await resolveSiteId(client);
    const cfgId = await resolveConfigListId(client, siteId);
    if (!cfgId) {
      context.res = { status: 500, body: { error: "Lista 'PRONEP-NF-Config' nao encontrada" } };
      return;
    }

    const mapaAtual = await lerMapaTelas(client, siteId, cfgId);
    let novoMapa;

    if (body.mapa && typeof body.mapa === 'object') {
      // Substitui o mapa inteiro (so telas validas, so com tokens)
      novoMapa = {};
      for (const t of Object.keys(body.mapa)) {
        const tela = String(t).trim();
        if (TELAS_IDS.indexOf(tela) < 0) continue; // ignora telas fora do escopo
        const toks = normTokens(body.mapa[t]);
        if (toks.length) novoMapa[tela] = toks;
      }
    } else if (body.tela) {
      // Atualiza apenas uma tela.
      const tela = String(body.tela).trim();
      if (TELAS_IDS.indexOf(tela) < 0) {
        context.res = { status: 400, body: { error: 'Tela invalida ou nao configuravel: ' + tela } };
        return;
      }
      novoMapa = Object.assign({}, mapaAtual);
      const toks = normTokens(body.tokens);
      if (toks.length) novoMapa[tela] = toks;
      else delete novoMapa[tela]; // lista vazia = remove config (volta ao padrao do papel)
    } else {
      context.res = { status: 400, body: { error: 'Envie { tela, tokens } ou { mapa }' } };
      return;
    }

    const r = await salvarMapaTelas(client, siteId, cfgId, novoMapa);

    auditRegistrar(authz.user, 'config_update',
      { tipo: 'acesso_telas', id: r.itemId },
      'sucesso',
      { action: r.action, tela: body.tela || '(mapa completo)', acessos: novoMapa }
    ).catch(function () {});

    context.res = {
      status: 200,
      headers: { 'Content-Type': 'application/json' },
      body: { ok: true, itemId: r.itemId, action: r.action, acessos: novoMapa }
    };
  } catch (err) {
    context.log && context.log.error && context.log.error('SalvarControleAcessosTelas:', err);
    context.res = { status: 500, headers: { 'Content-Type': 'application/json' },
      body: { error: (err && err.message) || String(err) } };
  }
};
