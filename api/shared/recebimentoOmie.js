/**
 * shared/recebimentoOmie.js — o caminho unico ate o conteudo de uma NF-e no Omie.
 *
 * POR QUE ISTO EXISTE
 * O DetalheNFOmie ja fazia tudo isto — resolver o docId, aplicar escopo, chamar o
 * ConsultarRecebimento, conferir a identidade da chave e buscar o fornecedor. Ao
 * escrever o EspelhoDaNota eu reescrevi o mesmo caminho do zero, com um indice
 * proprio e ate uma chave de entrada diferente (chave em vez de docId). Duas
 * copias da mesma regra de acesso a documento fiscal nao ficam iguais por muito
 * tempo, e a que diverge e sempre a que ninguem lembra de olhar — no caso, a
 * trava de escopo. Entao o caminho passa a ser um so.
 *
 * O QUE FICA AQUI: resolver e buscar.
 * O QUE NAO FICA: como cada tela FORMATA. O modal achata para uma tabela curta e
 * converte ausente em 0; o espelho precisa distinguir ausente de zero para nao
 * afirmar imposto zerado. Sao decisoes de tela opostas, e juntar as duas num
 * "mapeador comum" obrigaria uma das duas a mentir.
 */

require('isomorphic-fetch');
const { resolveListId, LIST_DOCFIS, soDigitos } = require('./documentosFiscais');
const { todosItens } = require('./escopoNF');
const { getCredentials } = require('./omie');

const OMIE_BASE = 'https://app.omie.com.br/api/v1';

/* O cadastro de fornecedores muda raramente e a lista inteira custa segundos.
   Sem cache, abrir cinco cards seguidos releria tudo cinco vezes. TTL curto para
   um cadastro corrigido agora aparecer no proximo card, nao so amanha. */
const TTL_FORN_MS = 5 * 60 * 1000;
let _fornCache = { em: 0, mapa: null };

async function fornecedorPorCnpj(client, siteId) {
  if (_fornCache.mapa && (Date.now() - _fornCache.em) < TTL_FORN_MS) return _fornCache.mapa;
  const mapa = {};
  const idForn = await resolveListId(client, siteId, 'PRONEP-NF-Fornecedores');
  if (idForn) {
    for (const it of await todosItens(client, siteId, idForn)) {
      const f = it.fields || {};
      /* Title=razao, field_1=tipoDoc, field_2=documento, field_3=fantasia,
         field_4=unidade, field_5=diretoria, field_6=uf, field_7=ativo,
         field_8=telefone, field_9=email, field_10=cidade, field_11=cep
         (mesmo mapa de ListarFornecedores — se um mudar, os dois mudam). */
      const doc = soDigitos(f.field_2);
      /* 11 digitos tambem entram: fornecedor pessoa fisica existe. */
      if ((doc.length !== 14 && doc.length !== 11) || mapa[doc]) continue;
      mapa[doc] = {
        razao: f.Title || '', fantasia: f.field_3 || '',
        tipoDocumento: f.field_1 || '', documento: doc,
        unidade: f.field_4 || '', diretoria: f.field_5 || '',
        uf: f.field_6 || '', cidade: f.field_10 || '', cep: f.field_11 || '',
        telefone: f.field_8 || '', email: f.field_9 || '',
        ativo: String(f.field_7 || '').toLowerCase() === 'sim'
      };
    }
  }
  _fornCache = { em: Date.now(), mapa: mapa };
  return mapa;
}

/* O cadastro local vence o Omie: unidade e diretoria so existem aqui, e consultar
   o Omie gasta chamada numa API de ~60/min. O Omie entra onde o local falha —
   fornecedor ainda nao cadastrado, que e tambem quando esses dados mais servem,
   porque e deles que sai o cadastro. */
async function fornecedorNoOmie(codigoClienteOmie, unidade) {
  if (!codigoClienteOmie) return null;
  const creds = getCredentials(unidade);
  const resp = await fetch(OMIE_BASE + '/geral/clientes/', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', 'Accept': 'application/json',
               'User-Agent': 'PronepNF/1.0 (Azure SWA Functions)' },
    body: JSON.stringify({ call: 'ConsultarCliente', app_key: creds.appKey,
                           app_secret: creds.appSecret,
                           param: [{ codigo_cliente_omie: Number(codigoClienteOmie) }] })
  });
  const texto = await resp.text();
  let d;
  try { d = JSON.parse(texto); } catch (e) { return null; }
  if (!d || d.faultstring) return null;
  const partes = [d.endereco, d.endereco_numero, d.complemento].filter(Boolean);
  return {
    razao: d.razao_social || '', fantasia: d.nome_fantasia || '',
    documento: soDigitos(d.cnpj_cpf),
    tipoDocumento: soDigitos(d.cnpj_cpf).length === 11 ? 'CPF' : 'CNPJ',
    logradouro: partes.join(', '), bairro: d.bairro || '',
    cidade: d.cidade || '', uf: d.estado || '', cep: d.cep || '',
    telefone: [d.telefone1_ddd, d.telefone1_numero].filter(Boolean).join(' '),
    email: d.email || '', inscricaoEstadual: d.inscricao_estadual || ''
  };
}

async function consultarRecebimento(chave, creds) {
  const resp = await fetch(OMIE_BASE + '/produtos/recebimentonfe/', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', 'Accept': 'application/json',
               'User-Agent': 'PronepNF/1.0 (Azure SWA Functions)' },
    body: JSON.stringify({ call: 'ConsultarRecebimento', app_key: creds.appKey,
                           app_secret: creds.appSecret, param: [{ cChaveNFe: chave }] })
  });
  const texto = await resp.text();
  let data;
  try { data = JSON.parse(texto); }
  catch (e) { throw new Error('Omie respondeu em formato inesperado (HTTP ' + resp.status + ')'); }
  if (data && data.faultstring) {
    const err = new Error(String(data.faultstring).slice(0, 200));
    /* "Nao encontrado" e RESPOSTA, nao falha: a nota existe no quadro mas nunca
       passou pelo modulo de recebimento. Quem chama precisa poder dizer isso em
       vez de "erro ao abrir", que faz a pessoa tentar de novo para sempre. */
    err.naoEncontrado = /n[aã]o (foi )?(encontrad|localizad)|inexistente|nenhum registro/i
      .test(err.message);
    throw err;
  }
  if (!resp.ok) throw new Error('Omie HTTP ' + resp.status);
  return data;
}

/* Le a linha do documento e devolve o que TODA tela precisa antes de decidir
   qualquer coisa: os campos crus, a chave, a unidade Omie e o card no formato
   exato que podeVer espera — nao um parecido. */
async function resolverDocumento(client, siteId, docId) {
  const listId = await resolveListId(client, siteId, LIST_DOCFIS);
  if (!listId) throw new Error('Lista de documentos nao existe');

  const item = await client.api('/sites/' + siteId + '/lists/' + listId + '/items/' + docId)
    .expand('fields').get();
  const f = (item && item.fields) || {};
  const cnpj = soDigitos(f.EmitenteCNPJ);
  const forn = cnpj ? (await fornecedorPorCnpj(client, siteId))[cnpj] : null;

  return {
    f: f,
    cnpj: cnpj,
    forn: forn,
    chave: soDigitos(f.ChaveAcesso),
    unidade: String(f.UnidadeOmie || 'RJ').toUpperCase(),
    card: {
      unidade: (forn && forn.unidade) || f.UnidadeOmie || '',
      diretoria: (forn && forn.diretoria) || '',
      fornecedorCadastrado: !!forn,
      cadastroIncompleto: !!forn && !(forn.unidade && forn.diretoria)
    }
  };
}

/* Monta a ficha do emitente na ordem de confianca: cadastro da Pronep, depois o
   Omie, e por ultimo o que o proprio documento guarda. Sem o ultimo degrau, um
   fornecedor sem cadastro e sem codigo Omie abriria a tela sem nome nenhum. */
async function fichaFornecedor(res, unidade) {
  if (res.forn) return Object.assign({ origem: 'cadastro' }, res.forn);
  if (res.f.CodigoClienteOmie) {
    try {
      const noOmie = await fornecedorNoOmie(res.f.CodigoClienteOmie, unidade);
      if (noOmie) return Object.assign({ origem: 'omie' }, noOmie);
    } catch (e) { /* ficha incompleta nao invalida o resto da tela */ }
  }
  if (res.f.EmitenteNome || res.cnpj) {
    return { origem: 'documento', razao: res.f.EmitenteNome || '', documento: res.cnpj,
             tipoDocumento: res.cnpj.length === 11 ? 'CPF' : 'CNPJ' };
  }
  return null;
}

/* A trava que nao pode cair: HTTP 200 nao prova que veio a nota certa. Se a API
   passar a devolver o ultimo recebimento ou um registro vizinho, a tela mostraria
   os itens de OUTRA nota para quem autoriza o pagamento. */
function conferirIdentidade(d, chavePedida) {
  const voltou = soDigitos(((d || {}).cabec || {}).cChaveNFe);
  return { ok: voltou === chavePedida, chaveDevolvida: voltou || '(vazia)' };
}

module.exports = {
  OMIE_BASE, fornecedorPorCnpj, fornecedorNoOmie, consultarRecebimento,
  resolverDocumento, fichaFornecedor, conferirIdentidade
};
