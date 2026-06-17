/**
 * shared/recorrentes.js — Contas recorrentes (Estagio 1).
 *
 * Guarda as DECISOES do gestor sobre quais fornecedores sao contas recorrentes,
 * na lista SharePoint PRONEP-NF-Recorrentes. A DETECCAO (varredura do historico de
 * NFs) fica no endpoint ListarRecorrentes; aqui ficam so os helpers de lista.
 *
 * Chave de uma conta recorrente = CNPJ(so digitos) | Diretoria(norm) | Unidade(norm).
 * Cada item guarda: EhRecorrente (Sim/Nao), DiaVencimento, ValorEstimado, DataFim,
 * Ativo (Sim/Nao), ConfirmadoPor, ConfirmadoEm.
 */

require('isomorphic-fetch');

const LIST_RECORRENTES = 'PRONEP-NF-Recorrentes';

// Colunas esperadas (name == internal, sem espacos pra evitar rename do SP).
const COLUNAS_RECORRENTES = [
  { name: 'CNPJFornecedor', def: { text: {} } },
  { name: 'Fornecedor',     def: { text: {} } },
  { name: 'Diretoria',      def: { text: {} } },
  { name: 'Unidade',        def: { text: {} } },
  { name: 'EhRecorrente',   def: { text: {} } },
  { name: 'DiaVencimento',  def: { number: {} } },
  { name: 'ValorEstimado',  def: { number: {} } },
  { name: 'DataFim',        def: { dateTime: { displayAs: 'standard', format: 'dateOnly' } } },
  { name: 'Ativo',          def: { text: {} } },
  { name: 'ConfirmadoPor',  def: { text: {} } },
  { name: 'ConfirmadoEm',   def: { dateTime: { displayAs: 'standard', format: 'dateTime' } } },
  { name: 'Observacoes',    def: { text: { allowMultipleLines: true } } }
];

const _cache = {}; // siteId -> listId

function _norm(s) {
  return String(s || '').toLowerCase()
    .normalize('NFD').replace(/\p{Diacritic}/gu, '')
    .replace(/[^a-z0-9]/g, '');
}

// Chave canonica de uma conta recorrente.
function chaveRecorrente(cnpj, diretoria, unidade) {
  const doc = String(cnpj || '').replace(/\D/g, '');
  return doc + '|' + _norm(diretoria) + '|' + _norm(unidade);
}

async function resolveListId(client, siteId) {
  if (_cache[siteId]) return _cache[siteId];
  let id = null;
  try {
    const lists = await client.api('/sites/' + siteId + '/lists')
      .filter("displayName eq '" + LIST_RECORRENTES + "'").get();
    if (lists.value && lists.value.length) id = lists.value[0].id;
  } catch (e) {
    // fallback: lista tudo e filtra
    try {
      const all = await client.api('/sites/' + siteId + '/lists').get();
      const found = (all.value || []).find(function (l) { return l.displayName === LIST_RECORRENTES; });
      if (found) id = found.id;
    } catch (e2) { /* ignore */ }
  }
  if (id) _cache[siteId] = id;
  return id;
}

// Le todas as decisoes salvas. Retorna { porChave: {chave: obj}, itens: [obj] }.
// obj: { itemId, chave, cnpj, fornecedor, diretoria, unidade, ehRecorrente(bool),
//        diaVencimento, valorEstimado, dataFim, ativo(bool), confirmadoPor, confirmadoEm }
async function lerDecisoes(client, siteId, listId) {
  const lid = listId || await resolveListId(client, siteId);
  const out = { porChave: {}, itens: [], listId: lid };
  if (!lid) return out;
  let url = '/sites/' + siteId + '/lists/' + lid + '/items?expand=fields&$top=500';
  let pages = 0;
  while (url && pages < 20) {
    const resp = await client.api(url).get();
    for (const it of (resp.value || [])) {
      const f = it.fields || {};
      const cnpj = String(f.CNPJFornecedor || '').replace(/\D/g, '');
      const diretoria = f.Diretoria || '';
      const unidade = f.Unidade || '';
      const chave = (f.Title && String(f.Title).indexOf('|') >= 0)
        ? f.Title
        : chaveRecorrente(cnpj, diretoria, unidade);
      const obj = {
        itemId: it.id,
        chave: chave,
        cnpj: cnpj,
        fornecedor: f.Fornecedor || '',
        diretoria: diretoria,
        unidade: unidade,
        ehRecorrente: (f.EhRecorrente === 'Sim'),
        diaVencimento: (f.DiaVencimento != null ? Number(f.DiaVencimento) : null),
        valorEstimado: (f.ValorEstimado != null ? Number(f.ValorEstimado) : null),
        dataFim: f.DataFim || '',
        ativo: (f.Ativo !== 'Nao'),
        confirmadoPor: f.ConfirmadoPor || '',
        confirmadoEm: f.ConfirmadoEm || ''
      };
      out.itens.push(obj);
      out.porChave[chave] = obj;
    }
    pages++;
    url = resp['@odata.nextLink']
      ? resp['@odata.nextLink'].replace('https://graph.microsoft.com/v1.0', '')
      : null;
  }
  return out;
}

// Upsert de uma decisao. dados: { cnpj, fornecedor, diretoria, unidade, ehRecorrente(bool),
//   diaVencimento, valorEstimado, dataFim, ativo(bool), confirmadoPor }
async function salvarDecisao(client, siteId, listId, dados) {
  const lid = listId || await resolveListId(client, siteId);
  if (!lid) throw new Error("Lista '" + LIST_RECORRENTES + "' nao encontrada. Rode CriarListaRecorrentes.");
  const cnpj = String(dados.cnpj || '').replace(/\D/g, '');
  const chave = chaveRecorrente(cnpj, dados.diretoria, dados.unidade);

  const fields = {
    Title: chave,
    CNPJFornecedor: cnpj,
    Fornecedor: String(dados.fornecedor || ''),
    Diretoria: String(dados.diretoria || ''),
    Unidade: String(dados.unidade || ''),
    EhRecorrente: dados.ehRecorrente ? 'Sim' : 'Nao',
    Ativo: (dados.ativo === false) ? 'Nao' : 'Sim',
    ConfirmadoPor: String(dados.confirmadoPor || ''),
    ConfirmadoEm: new Date().toISOString()
  };
  if (dados.diaVencimento != null && !isNaN(Number(dados.diaVencimento))) fields.DiaVencimento = Number(dados.diaVencimento);
  if (dados.valorEstimado != null && !isNaN(Number(dados.valorEstimado))) fields.ValorEstimado = Number(dados.valorEstimado);
  // DataFim: aceita 'AAAA-MM-DD' ou vazio (limpa)
  if (dados.dataFim) {
    const d = String(dados.dataFim).substring(0, 10);
    fields.DataFim = d + 'T00:00:00Z';
  } else {
    fields.DataFim = null; // limpa eventual data anterior
  }

  // Procura item existente pela chave (lista pequena: le tudo e acha)
  const atuais = await lerDecisoes(client, siteId, lid);
  const existente = atuais.porChave[chave];
  if (existente && existente.itemId) {
    await client.api('/sites/' + siteId + '/lists/' + lid + '/items/' + existente.itemId + '/fields').patch(fields);
    return { ok: true, action: 'updated', itemId: existente.itemId, chave: chave };
  }
  const created = await client.api('/sites/' + siteId + '/lists/' + lid + '/items').post({ fields: fields });
  return { ok: true, action: 'created', itemId: created.id, chave: chave };
}

module.exports = {
  LIST_RECORRENTES,
  COLUNAS_RECORRENTES,
  chaveRecorrente,
  resolveListId,
  lerDecisoes,
  salvarDecisao,
  _norm
};
