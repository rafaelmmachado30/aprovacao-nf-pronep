/**
 * /api/CasarDocumentosPendentes  (GET) — ADMIN.
 *
 * Liga os cards do quadro as NFs que ja foram lancadas no sistema.
 *
 * O QUE ESTAVA ACONTECENDO: uma NF lancada e aprovada continuava aparecendo em
 * "Novas". O casamento automatico roda quando a NOTA e criada; quando a conta
 * chega do Omie depois disso, ninguem mais tenta. E o quadro so reconhece o
 * vinculo por NotaItemId ou por chave de acesso — entao NFS-e de servico, que nao
 * tem chave, nunca fecha o ciclo. Nao e um card perdido: sao todas as NFS-e
 * lancadas antes da sincronizacao.
 *
 * Casa por chave (prova) ou por CNPJ + numero, e SO quando ha um unico candidato.
 * Ambiguidade nao vira palpite: vincular a nota errada faria um card mostrar o
 * status de outra, e uma conta em aberto apareceria como quitada.
 *
 * ?dryRun=1 mostra o que casaria sem gravar nada.
 */

require('isomorphic-fetch');
const { getGraphClient, resolveSiteId } = require('../shared/graph');
const { casarDocumentosPendentes } = require('../shared/documentosFiscais');

const ORCAMENTO_MS = 38000;

function readClientPrincipal(req) {
  const h = req.headers && req.headers['x-ms-client-principal'];
  if (!h) return null;
  try { return JSON.parse(Buffer.from(h, 'base64').toString('utf-8')); } catch (e) { return null; }
}

async function autorizado(req) {
  const segredo = process.env.SEFAZ_SECRET;
  const enviado = (req.headers &&
    (req.headers['x-automacao-secret'] || req.headers['X-Automacao-Secret'])) || '';
  if (segredo && enviado && enviado === segredo) return true;
  const p = readClientPrincipal(req);
  const roles = (p && p.userRoles) || [];
  if (roles.includes('administrador') || roles.includes('admin')) return true;
  try {
    const { getUser } = require('../shared/auth');
    const user = await getUser(req);
    if (!user) return false;
    const { isAdminEmail } = require('../shared/authz');
    if (isAdminEmail((user.email || '').toLowerCase())) return true;
    const { getUserRoles } = require('../shared/userRoles');
    return ((await getUserRoles(user)) || []).includes('administrador');
  } catch (e) { return false; }
}

module.exports = async function (context, req) {
  const t0 = Date.now();
  const q = req.query || {};
  const dryRun = q.dryRun === '1' || q.dryRun === 'true';

  try {
    if (!(await autorizado(req))) {
      context.res = { status: 403, headers: { 'Content-Type': 'application/json' },
        body: { error: 'Nao autorizado' } };
      return;
    }

    const client = await getGraphClient();
    const siteId = await resolveSiteId(client);

    /* dryRun com prazo ja vencido: monta tudo e nao grava nada. Reaproveita a
       trava de tempo do gravarEmLote em vez de duplicar um caminho "so simular",
       que seria um segundo codigo capaz de divergir do que grava de verdade. */
    const r = await casarDocumentosPendentes(client, siteId, {
      prazoFinal: dryRun ? Date.now() - 1 : t0 + ORCAMENTO_MS
    });

    const avisos = [];
    if (r.mapaDeColunasEIdentidade === false) {
      /* Se os nomes internos diferem dos de exibicao, quem le fields.Status cru
         recebe undefined e o card cai na coluna errada em silencio. */
      avisos.push('A lista de Notas tem colunas com nome interno diferente do nome ' +
        'de exibicao. Este endpoint traduz, mas outros pontos que leem o campo cru ' +
        'podem estar lendo undefined — vale revisar.');
    }
    if (r.ambiguos && r.ambiguos.length) {
      avisos.push(r.ambiguos.length + ' documento(s) com mais de uma nota candidata. ' +
        'Nao foram vinculados de proposito: escolher errado faria o card exibir o ' +
        'status de outra NF.');
    }
    if (r.linhasRestantes) {
      avisos.push(r.linhasRestantes + ' linha(s) ficaram sem gravar por tempo. ' +
        'Rode de novo — o que ja foi vinculado nao se repete.');
    }

    context.res = { status: 200, headers: { 'Content-Type': 'application/json' },
      body: Object.assign({
        ok: true,
        mensagem: (dryRun ? '[SIMULACAO] ' : '') +
          (r.gruposCasados || 0) + ' NF(s) reconhecida(s)' +
          (dryRun ? ' — nada gravado' : ' · ' + (r.linhasVinculadas || 0) + ' linha(s) vinculada(s)'),
        avisos: avisos, timeMs: Date.now() - t0
      }, r) };
  } catch (err) {
    context.res = { status: 500, headers: { 'Content-Type': 'application/json' },
      body: { error: (err && err.message) || String(err), timeMs: Date.now() - t0 } };
  }
};
