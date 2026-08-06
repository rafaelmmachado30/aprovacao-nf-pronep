/**
 * /api/LimparDocumentosAntigos  (GET|POST) — ADMIN.
 *
 * Remove da lista as contas com vencimento ANTERIOR a data de corte. Existe
 * porque a primeira sincronizacao entrou sem o corte e trouxe historico vencido
 * do Omie — o filtro novo impede que entre mais, mas nao apaga o que ja entrou.
 *
 * SEMPRE COMECE COM ?dryRun=1. Ele conta e mostra amostra sem apagar nada.
 *
 * DUAS PROTECOES QUE NAO PODEM SAIR:
 *   1. Nunca apaga linha VINCULADA a uma nota (NotaItemId preenchido). Esse
 *      vinculo e trabalho humano — alguem lancou a NF e o sistema casou as duas.
 *      Apagar destruiria o merge, e o Omie nao sabe recria-lo.
 *   2. Nunca apaga linha que NAO veio do Omie. Documento da SEFAZ ou lancamento
 *      manual nao esta sob a regra do corte; a limpeza e da carga automatica.
 *
 * Query:
 *   ?dryRun=1     conta e mostra amostra, nao apaga  (COMECE POR AQUI)
 *   ?unidade=RJ   limita a uma unidade
 *   ?corte=2026-08-01  sobrepoe a data de corte da configuracao
 */

require('isomorphic-fetch');
const { getGraphClient, resolveSiteId } = require('../shared/graph');
const { resolveListId, LIST_DOCFIS, lerCorteVencimento } = require('../shared/documentosFiscais');

const ORCAMENTO_MS = 30000;

function readClientPrincipal(req) {
  const h = req.headers && req.headers['x-ms-client-principal'];
  if (!h) return null;
  try { return JSON.parse(Buffer.from(h, 'base64').toString('utf-8')); } catch (e) { return null; }
}

async function isAdmin(req) {
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
  const unidade = q.unidade ? String(q.unidade).toUpperCase() : null;
  const diag = { step: 'init', dryRun: dryRun, unidade: unidade, timeMs: 0 };

  try {
    if (!(await isAdmin(req))) {
      context.res = { status: 403, headers: { 'Content-Type': 'application/json' },
        body: { error: 'Apenas admin' } };
      return;
    }

    const client = await getGraphClient();
    const siteId = await resolveSiteId(client);
    const listId = await resolveListId(client, siteId, LIST_DOCFIS);
    if (!listId) throw new Error('Lista ' + LIST_DOCFIS + ' nao existe');

    let corteData = String(q.corte || '');
    if (!/^\d{4}-\d{2}-\d{2}$/.test(corteData)) {
      const c = await lerCorteVencimento(client, siteId);
      corteData = c.data;
      diag.corteOrigem = c.origem;
    } else {
      diag.corteOrigem = 'query';
    }
    diag.corte = corteData;

    diag.step = 'ler';
    const itens = [];
    let url = '/sites/' + siteId + '/lists/' + listId + '/items?expand=fields&$top=999';
    let p = 0;
    while (url && p < 20) {
      const r = await client.api(url).get();
      itens.push.apply(itens, r.value || []);
      p++;
      const nl = r['@odata.nextLink'];
      url = nl ? nl.replace('https://graph.microsoft.com/v1.0', '') : null;
    }
    diag.totalNaLista = itens.length;

    const alvos = [];
    let protegidasPorVinculo = 0;
    let foraDoOmie = 0;
    let semVencimento = 0;

    for (const it of itens) {
      const f = it.fields || {};
      if (!f.CodigoLancamentoOmie) { foraDoOmie++; continue; }
      if (unidade && f.UnidadeOmie !== unidade) continue;
      if (f.NotaItemId) { protegidasPorVinculo++; continue; }

      const venc = f.DataVencimento ? String(f.DataVencimento).substring(0, 10) : '';
      if (!venc) { semVencimento++; continue; }   /* sem data nao da para julgar: fica */
      if (venc >= corteData) continue;

      alvos.push({
        id: it.id, venc: venc, unidade: f.UnidadeOmie || '',
        nf: f.NumeroNF || '', valor: f.Valor, status: f.StatusOmie || '',
        emitente: f.EmitenteNome || ''
      });
    }

    diag.aRemover = alvos.length;
    diag.protegidasPorVinculo = protegidasPorVinculo;
    diag.foraDoOmie = foraDoOmie;
    diag.semVencimento = semVencimento;
    diag.amostra = alvos.slice(0, 5);

    if (dryRun) {
      diag.step = 'done';
      diag.timeMs = Date.now() - t0;
      context.res = { status: 200, headers: { 'Content-Type': 'application/json' },
        body: Object.assign({ ok: true,
          mensagem: '[SIMULACAO] ' + alvos.length + ' linha(s) seriam removidas (vencimento < ' +
                    corteData + '). Nada foi apagado.' }, diag) };
      return;
    }

    /* Apaga em lotes de 20 pelo $batch, respeitando o prazo. Interromper no meio
       nao corrompe nada: cada linha e independente e a proxima execucao continua. */
    diag.step = 'apagar';
    let removidas = 0;
    const falhas = [];
    for (let i = 0; i < alvos.length; i += 20) {
      if (Date.now() - t0 > ORCAMENTO_MS) break;
      const fatia = alvos.slice(i, i + 20);
      const requests = fatia.map(function (a, j) {
        return { id: String(j + 1), method: 'DELETE',
                 url: '/sites/' + siteId + '/lists/' + listId + '/items/' + a.id };
      });
      let resp;
      try { resp = await client.api('/$batch').post({ requests: requests }); }
      catch (e) { falhas.push({ lote: i, erro: e.message }); continue; }
      for (const r of ((resp && resp.responses) || [])) {
        if (r.status >= 200 && r.status < 300) removidas++;
        else falhas.push({ status: r.status, erro: (r.body && r.body.error && r.body.error.message) || '?' });
      }
    }

    diag.removidas = removidas;
    diag.restantes = alvos.length - removidas;
    if (falhas.length) diag.falhas = falhas.slice(0, 5);

    diag.step = 'done';
    diag.timeMs = Date.now() - t0;
    context.res = { status: 200, headers: { 'Content-Type': 'application/json' },
      body: Object.assign({ ok: falhas.length === 0,
        mensagem: removidas + ' linha(s) removida(s)' +
                  (diag.restantes ? ' · faltam ' + diag.restantes + ', rode de novo' : '') }, diag) };
  } catch (err) {
    diag.timeMs = Date.now() - t0;
    context.res = { status: 500, headers: { 'Content-Type': 'application/json' },
      body: Object.assign({ error: (err && err.message) || String(err) }, diag) };
  }
};
