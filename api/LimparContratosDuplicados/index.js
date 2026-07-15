/**
 * /api/LimparContratosDuplicados (GET) — ADMIN ONLY
 *
 * Remove itens DUPLICADOS na lista PRONEP-NF-Contratos. Duplicata = mesmo DriveItemId
 * (mesmo arquivo no SharePoint) aparecendo em mais de um item de lista — causado pelo bug
 * antigo de dedup (que lia so a 1a pagina). Mantem 1 item por DriveItemId (o mais
 * recentemente modificado, pra preservar edicoes dos gestores) e remove os demais.
 *
 * SEGURANCA:
 *   - Admin only.
 *   - DRY-RUN por padrao: so REPORTA o que removeria. Para aplicar de fato: ?aplicar=true
 *   - Nunca remove item unico (sem duplicata). Itens sem DriveItemId sao ignorados.
 */

require('isomorphic-fetch');
const { requireAdmin } = require('../shared/authz');
const { getGraphClient, resolveSiteId } = require('../shared/graph');

const LIST_CONTRATOS = 'PRONEP-NF-Contratos';


module.exports = async function (context, req) {
  try {
    const authz = await requireAdmin(context, req);
    if (!authz) return;

    const aplicar = String((req.query && req.query.aplicar) || '') === 'true';
    const client = getGraphClient();
    const siteId = await resolveSiteId(client);

    const lists = await client.api('/sites/' + siteId + '/lists')
      .filter("displayName eq '" + LIST_CONTRATOS + "'").get();
    if (!lists.value || !lists.value.length) {
      context.res = { status: 404, body: { error: "Lista '" + LIST_CONTRATOS + "' nao encontrada" } };
      return;
    }
    const listId = lists.value[0].id;
    const cols = await client.api('/sites/' + siteId + '/lists/' + listId + '/columns').get();
    let cDrive = 'DriveItemId';
    for (const c of (cols.value || [])) { if (c.displayName === 'DriveItemId' && c.name) cDrive = c.name; }

    // Pagina TODOS os itens
    const todos = [];
    let url = '/sites/' + siteId + '/lists/' + listId + '/items?expand=fields&$top=999';
    let pgs = 0;
    while (url && pgs < 60) {
      const r = await client.api(url).get();
      for (const it of (r.value || [])) {
        const f = it.fields || {};
        todos.push({
          id: it.id,
          driveId: String(f[cDrive] || f.DriveItemId || ''),
          modificado: it.lastModifiedDateTime || (f && f.Modified) || ''
        });
      }
      pgs++;
      url = r['@odata.nextLink'] ? r['@odata.nextLink'].replace('https://graph.microsoft.com/v1.0', '') : null;
    }

    // Agrupa por DriveItemId; em cada grupo com >1, mantem o mais recente.
    const grupos = {};
    for (const it of todos) {
      if (!it.driveId) continue; // sem DriveItemId -> ignora (nunca remove)
      (grupos[it.driveId] = grupos[it.driveId] || []).push(it);
    }
    const aRemover = [];
    let gruposComDup = 0;
    for (const did of Object.keys(grupos)) {
      const g = grupos[did];
      if (g.length <= 1) continue;
      gruposComDup++;
      g.sort(function (a, b) { return String(b.modificado).localeCompare(String(a.modificado)); });
      // mantem g[0] (mais recente), remove o resto
      for (let i = 1; i < g.length; i++) aRemover.push(g[i].id);
    }

    let removidos = 0;
    const erros = [];
    if (aplicar) {
      // Apaga em LOTE via Graph $batch (20 por request) com orcamento de tempo (~30s),
      // pra nao estourar o timeout do gateway. Retorna quantos ainda faltam; rode de novo
      // ate restantes = 0. Sao milhares — leva algumas execucoes.
      const inicio = Date.now();
      const BUDGET_MS = 30000;
      let i = 0;
      while (i < aRemover.length && (Date.now() - inicio) < BUDGET_MS) {
        const lote = aRemover.slice(i, i + 20);
        const requests = lote.map(function (itemId, idx) {
          return { id: String(idx + 1), method: 'DELETE', url: '/sites/' + siteId + '/lists/' + listId + '/items/' + itemId };
        });
        try {
          const resp = await client.api('/$batch').post({ requests: requests });
          for (const r of (resp.responses || [])) {
            if (r.status >= 200 && r.status < 300) removidos++;
            else if (r.status !== 404) erros.push({ id: r.id, status: r.status });
            else removidos++; // 404 = ja nao existe, conta como removido
          }
        } catch (e) {
          erros.push({ erro: e.message });
        }
        i += lote.length;
      }
    }
    const restantes = aplicar ? (aRemover.length - removidos) : aRemover.length;

    context.res = {
      status: 200,
      headers: { 'Content-Type': 'application/json', 'Cache-Control': 'no-store' },
      body: {
        ok: true,
        dryRun: !aplicar,
        totalItens: todos.length,
        gruposComDuplicata: gruposComDup,
        duplicatasIdentificadas: aRemover.length,
        removidos: aplicar ? removidos : 0,
        restantes: restantes,
        erros: erros.slice(0, 20),
        dica: !aplicar
          ? 'DRY-RUN. Para remover de fato, chame com ?aplicar=true (em lotes; rode de novo ate restantes=0).'
          : (restantes > 0
              ? 'Removidos ' + removidos + ' neste lote. Ainda faltam ' + restantes + ' — chame ?aplicar=true de novo.'
              : 'Limpeza concluida! Nao ha mais duplicatas.')
      }
    };
  } catch (err) {
    context.log && context.log.error && context.log.error('LimparContratosDuplicados:', err);
    context.res = { status: 500, headers: { 'Content-Type': 'application/json' },
      body: { error: (err && err.message) || String(err) } };
  }
};
