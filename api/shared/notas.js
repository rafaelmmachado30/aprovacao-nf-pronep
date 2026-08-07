/**
 * shared/notas.js — leitura da lista de Notas com os nomes de coluna certos.
 *
 * POR QUE ISTO EXISTE: o SharePoint guarda cada coluna com um `name` interno que
 * nem sempre e igual ao `displayName`. Colunas criadas por importacao viram
 * field_1, field_2... Alguns endpoints ja traduzem (IntegrarOmie, ListarRecorrentes)
 * e outros leem `fields.Status` cru — o que funciona enquanto o nome bate e
 * silenciosamente devolve undefined quando nao bate. Undefined em campo de status
 * nao explode: so faz a nota parecer sem status, e o card vai parar na coluna
 * errada sem ninguem perceber.
 *
 * Aqui a traducao e obrigatoria e o resultado diz se o mapa e identidade, para o
 * chamador poder denunciar quando alguem estiver lendo cru em outro lugar.
 */

const LIST_NOTAS = 'PRONEP-NF-NotasFiscais';

async function mapaDeColunas(client, siteId, listId) {
  const resp = await client.api('/sites/' + siteId + '/lists/' + listId + '/columns').get();
  const inv = {};
  let identidade = true;
  for (const c of (resp.value || [])) {
    if (!c.displayName || !c.name) continue;
    inv[c.name] = c.displayName;
    if (c.name !== c.displayName) identidade = false;
  }
  return { inv: inv, identidade: identidade };
}

/**
 * Le TODAS as notas com os campos ja traduzidos para displayName.
 * @returns {{notas:[{id,f}], identidade:boolean, paginas:number}}
 */
async function carregarNotas(client, siteId, listId, maxPaginas) {
  const { inv, identidade } = await mapaDeColunas(client, siteId, listId);
  const notas = [];
  let url = '/sites/' + siteId + '/lists/' + listId + '/items?expand=fields&$top=999';
  let p = 0;
  while (url && p < (maxPaginas || 20)) {
    const r = await client.api(url).get();
    for (const it of (r.value || [])) {
      const f = {};
      for (const k of Object.keys(it.fields || {})) f[inv[k] || k] = it.fields[k];
      notas.push({ id: String(it.id), f: f });
    }
    p++;
    const nl = r['@odata.nextLink'];
    url = nl ? nl.replace('https://graph.microsoft.com/v1.0', '') : null;
  }
  return { notas: notas, identidade: identidade, paginas: p };
}

module.exports = { LIST_NOTAS, carregarNotas, mapaDeColunas };
