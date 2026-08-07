/**
 * shared/documentosFiscais.js — NFs a Pagar (quadro do ciclo completo).
 *
 * Uma NF-e emitida CONTRA o nosso CNPJ existe antes de alguem lanca-la. Aqui mora
 * esse fato: o documento e conhecido, e o lancamento e um passo posterior que pode
 * ou nao ter acontecido.
 *
 * Chave natural = CHAVE DE ACESSO (44 digitos), unica por documento no Brasil todo.
 *
 * DIFERENCA IMPORTANTE PARA O SISTEMA NOVO (Supabase):
 * la a unicidade da chave e garantida pelo BANCO e o merge e um TRIGGER, entao vale
 * para toda porta de entrada. Aqui nao ha constraint nem trigger no SharePoint:
 *   - a unicidade e consulta-antes-de-gravar, com janela de corrida real
 *   - o merge tem de ser CHAMADO pelo codigo de lancamento
 * Duas execucoes simultaneas do cron podem duplicar. Mitigacao: o cron e sequencial
 * por CNPJ e o ponteiro de NSU so avanca depois da gravacao, entao uma reexecucao
 * relê o mesmo lote e a consulta-antes-de-gravar pega o duplicado na segunda volta.
 * Nao e garantia como no Postgres — e o que a plataforma permite.
 */

require('isomorphic-fetch');

const LIST_DOCFIS = 'PRONEP-NF-DocumentosFiscais';
const LIST_SEFAZ  = 'PRONEP-NF-SefazControle';
const LIST_NOTAS  = 'PRONEP-NF-NotasFiscais';

/* name == internal, sem espacos, pra evitar o rename automatico do SharePoint. */
const COLUNAS_DOCFIS = [
  { name: 'ChaveAcesso',    def: { text: {} } },
  { name: 'Origem',         def: { text: {} } },          // sefaz | xml | email | manual
  { name: 'NSU',            def: { number: {} } },
  { name: 'NumeroNF',       def: { text: {} } },
  { name: 'Serie',          def: { text: {} } },
  { name: 'EmitenteCNPJ',   def: { text: {} } },
  { name: 'EmitenteNome',   def: { text: {} } },
  { name: 'Valor',          def: { number: {} } },
  { name: 'DataEmissao',    def: { dateTime: { displayAs: 'standard', format: 'dateOnly' } } },
  { name: 'DataVencimento', def: { dateTime: { displayAs: 'standard', format: 'dateOnly' } } },
  { name: 'CNPJDestino',    def: { text: {} } },          // qual filial recebeu
  { name: 'XmlPath',        def: { text: {} } },
  { name: 'NotaItemId',     def: { text: {} } },          // id do item na lista de Notas
  { name: 'VinculadoEm',    def: { dateTime: { displayAs: 'standard', format: 'dateTime' } } },
  { name: 'VinculadoAuto',  def: { text: {} } },          // Sim | Nao
  { name: 'Descartado',     def: { text: {} } },          // Sim | Nao
  { name: 'MotivoDescarte', def: { text: { allowMultipleLines: true } } },

  /* --- origem OMIE ---------------------------------------------------------
     O Omie e a fonte do quadro (a SEFAZ ficou como fallback: os dois consomem o
     mesmo DFe e a cota e por CNPJ). Uma conta a pagar do Omie NAO e a mesma coisa
     que um documento fiscal: uma NF parcelada vira N contas, cada uma com seu
     vencimento e seu status. Por isso a chave natural aqui e o
     CodigoLancamentoOmie — uma linha por PARCELA — e o agrupamento em um unico
     card acontece na leitura, por ChaveAcesso. */
  { name: 'CodigoLancamentoOmie', def: { text: {} } },    // chave natural da parcela
  { name: 'CodigoClienteOmie',    def: { text: {} } },    // guarda o codigo p/ reconsultar o CNPJ que ficou pendente
  { name: 'NumeroParcela',        def: { text: {} } },    // "001/003"
  { name: 'StatusOmie',           def: { text: {} } },    // PAGO | A VENCER | ATRASADO | VENCE HOJE | CANCELADO
  { name: 'UnidadeOmie',          def: { text: {} } },    // RJ | SP | ES (qual empresa Omie)
  { name: 'CodigoBarras',         def: { text: {} } },    // ficha de compensacao — casa com a Conciliacao Bancaria
  { name: 'SincronizadoEm',       def: { dateTime: { displayAs: 'standard', format: 'dateTime' } } }
];

const COLUNAS_SEFAZ = [
  { name: 'CNPJ',            def: { text: {} } },
  { name: 'Apelido',         def: { text: {} } },
  { name: 'UltimoNSU',       def: { number: {} } },
  { name: 'MaxNSU',          def: { number: {} } },
  { name: 'UltimaConsulta',  def: { dateTime: { displayAs: 'standard', format: 'dateTime' } } },
  { name: 'UltimoCStat',     def: { text: {} } },
  { name: 'UltimoMotivo',    def: { text: { allowMultipleLines: true } } },
  { name: 'Baixados',        def: { number: {} } }
];

/* A lista de Notas NAO tinha a chave de acesso. Sem ela o merge cai no caminho
   fraco (CNPJ + numero + serie), que erra quando o mesmo fornecedor emite duas
   notas de mesmo numero em series diferentes ou o numero se repete no ano. */
const COLUNAS_NOTAS_EXTRA = [
  { name: 'ChaveAcesso', def: { text: {} } }
];

const _cacheList = {};   // siteId|listName -> listId

async function resolveListId(client, siteId, listName) {
  const ck = siteId + '|' + listName;
  if (_cacheList[ck]) return _cacheList[ck];
  let id = null;
  try {
    const r = await client.api('/sites/' + siteId + '/lists')
      .filter("displayName eq '" + listName + "'").get();
    if (r.value && r.value.length) id = r.value[0].id;
  } catch (e) {
    /* O filtro do Graph falha em alguns tenants; cai pra listagem completa. */
    try {
      const all = await client.api('/sites/' + siteId + '/lists').get();
      const f = (all.value || []).find(function (l) { return l.displayName === listName; });
      if (f) id = f.id;
    } catch (e2) { /* segue null */ }
  }
  if (id) _cacheList[ck] = id;
  return id;
}

function soDigitos(s) { return String(s || '').replace(/\D/g, ''); }

function chaveValida(ch) { return /^[0-9]{44}$/.test(soDigitos(ch)); }

/* Chave fraca, usada quando o documento nao traz a chave de acesso. */
function chaveFraca(cnpj, numero, serie) {
  return soDigitos(cnpj) + '|' + String(numero || '').trim() +
         '|' + String(serie == null ? '' : serie).trim();
}

async function buscarPorChave(client, siteId, chave) {
  const listId = await resolveListId(client, siteId, LIST_DOCFIS);
  if (!listId) return null;
  const ch = soDigitos(chave);
  const r = await client
    .api('/sites/' + siteId + '/lists/' + listId + '/items')
    .expand('fields')
    .filter("fields/ChaveAcesso eq '" + ch + "'")
    .header('Prefer', 'HonorNonIndexedQueriesWarningMayFailRandomly')
    .top(2)
    .get();
  const v = (r && r.value) || [];
  return v.length ? v[0] : null;
}

/**
 * Grava o documento se ele ainda nao existir. Idempotente por chave de acesso
 * DENTRO DO POSSIVEL: sem constraint no SharePoint, duas gravacoes simultaneas da
 * mesma chave passam as duas. Quem chama deve ser sequencial.
 * @returns {{ novo: boolean, itemId: string }}
 */
async function gravarDocumento(client, siteId, doc) {
  const ch = soDigitos(doc.chaveAcesso);
  if (!chaveValida(ch)) throw new Error('Chave de acesso invalida: ' + doc.chaveAcesso);

  const jaTem = await buscarPorChave(client, siteId, ch);
  if (jaTem) {
    /* Ja conhecido: COMPLETA o que falta sem sobrescrever. A SEFAZ entrega primeiro
       o RESUMO e so depois o XML completo; a segunda visita nao pode apagar o que a
       primeira trouxe. */
    const f = jaTem.fields || {};
    const patch = {};
    if (!f.EmitenteNome && doc.emitenteNome) patch.EmitenteNome = doc.emitenteNome;
    if (f.Valor == null && doc.valor != null) patch.Valor = Number(doc.valor);
    if (!f.DataEmissao && doc.dataEmissao) patch.DataEmissao = doc.dataEmissao;
    if (!f.DataVencimento && doc.dataVencimento) patch.DataVencimento = doc.dataVencimento;
    if (f.NSU == null && doc.nsu != null) patch.NSU = Number(doc.nsu);
    if (!f.XmlPath && doc.xmlPath) patch.XmlPath = doc.xmlPath;
    if (Object.keys(patch).length) {
      const listId = await resolveListId(client, siteId, LIST_DOCFIS);
      await client.api('/sites/' + siteId + '/lists/' + listId + '/items/' + jaTem.id + '/fields')
        .patch(patch);
    }
    return { novo: false, itemId: jaTem.id };
  }

  const listId = await resolveListId(client, siteId, LIST_DOCFIS);
  const criado = await client.api('/sites/' + siteId + '/lists/' + listId + '/items').post({
    fields: {
      Title: 'NF ' + (doc.numeroNF || '') + ' - ' + (doc.emitenteNome || ''),
      ChaveAcesso: ch,
      Origem: doc.origem || 'sefaz',
      NSU: doc.nsu != null ? Number(doc.nsu) : null,
      NumeroNF: String(doc.numeroNF || ''),
      Serie: doc.serie != null ? String(doc.serie) : '',
      EmitenteCNPJ: soDigitos(doc.emitenteCNPJ),
      EmitenteNome: doc.emitenteNome || '',
      Valor: doc.valor != null ? Number(doc.valor) : null,
      DataEmissao: doc.dataEmissao || null,
      DataVencimento: doc.dataVencimento || null,
      CNPJDestino: soDigitos(doc.cnpjDestino),
      XmlPath: doc.xmlPath || '',
      Descartado: 'Nao',
      VinculadoAuto: 'Nao'
    }
  });
  return { novo: true, itemId: criado.id };
}

/* Indexa as contas do Omie ja gravadas, por CodigoLancamentoOmie.
   Uma leitura so, em vez de uma consulta por conta: sincronizar 548 contas com
   uma busca cada seriam 548 idas ao Graph e o estouro certo dos 45s da Function. */
async function indexarPorCodigoOmie(client, siteId, unidade) {
  const listId = await resolveListId(client, siteId, LIST_DOCFIS);
  if (!listId) return {};
  const idx = {};
  let url = '/sites/' + siteId + '/lists/' + listId + '/items?expand=fields&$top=999';
  let p = 0;
  while (url && p < 20) {
    const r = await client.api(url).get();
    for (const it of (r.value || [])) {
      const f = it.fields || {};
      if (!f.CodigoLancamentoOmie) continue;
      if (unidade && f.UnidadeOmie && f.UnidadeOmie !== unidade) continue;
      idx[String(f.CodigoLancamentoOmie)] = { id: it.id, fields: f };
    }
    p++;
    const nl = r['@odata.nextLink'];
    url = nl ? nl.replace('https://graph.microsoft.com/v1.0', '') : null;
  }
  return idx;
}

/**
 * Grava (ou atualiza) uma conta a pagar do Omie.
 *
 * DIFERENCA IMPORTANTE PARA O DOCUMENTO DA SEFAZ: la o documento e imutavel e a
 * regra e "completar sem sobrescrever". Aqui o registro MUDA — o status vira PAGO,
 * o vencimento pode ser renegociado — e e essa mudanca que move o card de coluna.
 * Entao aqui a regra e o oposto: o Omie manda, e sobrescreve.
 *
 * O que NAO se sobrescreve e o vinculo com a nota lancada aqui (NotaItemId): esse
 * dado e nosso, o Omie nao sabe dele, e apaga-lo desfaria o merge em silencio.
 *
 * @returns {{ novo: boolean, mudou: boolean, itemId: string }}
 */
async function gravarContaOmie(client, siteId, conta, indice) {
  const listId = await resolveListId(client, siteId, LIST_DOCFIS);
  if (!listId) throw new Error('Lista ' + LIST_DOCFIS + ' nao existe');

  const cod = String(conta.codigoLancamentoOmie || '');
  if (!cod) throw new Error('Conta do Omie sem codigo_lancamento_omie');

  const campos = {
    CodigoLancamentoOmie: cod,
    CodigoClienteOmie: String(conta.codigoClienteOmie || ''),
    Origem: 'omie',
    NumeroNF: String(conta.numeroNF || ''),
    NumeroParcela: String(conta.numeroParcela || ''),
    EmitenteNome: conta.emitenteNome || '',
    Valor: conta.valor != null ? Number(conta.valor) : null,
    DataEmissao: conta.dataEmissao || null,
    DataVencimento: conta.dataVencimento || null,
    StatusOmie: conta.statusOmie || '',
    UnidadeOmie: conta.unidade || '',
    CodigoBarras: conta.codigoBarras || '',
    SincronizadoEm: new Date().toISOString()
  };
  const ch = soDigitos(conta.chaveAcesso);
  if (chaveValida(ch)) campos.ChaveAcesso = ch;

  /* CNPJ so entra quando FOI resolvido. Gravar vazio apagaria o que uma execucao
     anterior ja tinha descoberto — e sem CNPJ o card perde unidade e diretoria,
     ou seja, deixa de ter aprovador. Ausencia aqui significa "ainda nao sei",
     nunca "e vazio". */
  const cnpjResolvido = soDigitos(conta.emitenteCNPJ);
  if (cnpjResolvido.length === 14) campos.EmitenteCNPJ = cnpjResolvido;
  if (!conta.emitenteNome) delete campos.EmitenteNome;

  const existente = indice ? indice[cod] : null;
  if (!existente) {
    campos.Title = 'NF ' + (conta.numeroNF || '') + ' - ' + (conta.emitenteNome || '');
    campos.Descartado = 'Nao';
    campos.VinculadoAuto = 'Nao';
    const criado = await client.api('/sites/' + siteId + '/lists/' + listId + '/items')
      .post({ fields: campos });
    return { novo: true, mudou: true, itemId: criado.id };
  }

  /* So grava se algo REALMENTE mudou. Sem essa comparacao, cada sincronizacao
     faria 548 PATCHes identicos — gasto de cota do Graph e ruido no historico da
     lista, que e o que alguem vai consultar quando quiser entender um card. */
  const antes = existente.fields || {};
  const mudou = Object.keys(campos).some(function (k) {
    if (k === 'SincronizadoEm') return false;   /* muda sempre; nao conta como mudanca */
    const a = antes[k], b = campos[k];
    if (a == null && (b == null || b === '')) return false;
    if (k === 'Valor') return Number(a || 0) !== Number(b || 0);
    if (k === 'DataEmissao' || k === 'DataVencimento') {
      return String(a || '').substring(0, 10) !== String(b || '').substring(0, 10);
    }
    return String(a || '') !== String(b || '');
  });
  if (!mudou) return { novo: false, mudou: false, itemId: existente.id };

  await client.api('/sites/' + siteId + '/lists/' + listId + '/items/' + existente.id + '/fields')
    .patch(campos);
  return { novo: false, mudou: true, itemId: existente.id };
}

/**
 * Grava varias contas do Omie de uma vez, usando o $batch do Graph.
 *
 * POR QUE ISTO EXISTE: medido em producao, ler as contas do RJ leva ~20s e a
 * Function corta em ~45s. Gravando uma por uma (~0,3s cada), 541 linhas levariam
 * ~160s — seriam 6 execucoes so para a primeira carga de UMA unidade, 18 nas
 * tres. Com lotes de 20 por requisicao isso vira ~28 chamadas, e cabe numa
 * execucao so.
 *
 * O $batch NAO e transacional: cada requisicao do lote pode falhar sozinha. Por
 * isso o retorno diz o que passou e o que falhou, item a item — em vez de tratar
 * o lote como tudo-ou-nada, que esconderia falhas parciais.
 *
 * @param operacoes [{ tipo:'post'|'patch', itemId?, fields }]
 * @returns {{ ok:number, falhas:[{indice,status,erro}] }}
 */
async function gravarEmLote(client, siteId, listId, operacoes, prazoFinal) {
  const LIMITE = 20;                 /* teto do $batch do Graph */
  let ok = 0;
  let processadas = 0;
  const falhas = [];

  for (let i = 0; i < operacoes.length; i += LIMITE) {
    /* PRAZO. Sem esta parada a funcao seguia gravando ate a plataforma matar a
       execucao — e ai o cliente recebe "Backend call failure", sem resposta e sem
       saber o que entrou. Parar por conta propria devolve o diagnostico e diz
       quantas ficaram; a proxima execucao continua de onde parou. */
    if (prazoFinal && Date.now() > prazoFinal) break;
    const fatia = operacoes.slice(i, i + LIMITE);
    const requests = fatia.map(function (op, j) {
      const base = '/sites/' + siteId + '/lists/' + listId + '/items';
      return op.tipo === 'patch'
        ? { id: String(j + 1), method: 'PATCH', url: base + '/' + op.itemId + '/fields',
            headers: { 'Content-Type': 'application/json' }, body: op.fields }
        : { id: String(j + 1), method: 'POST', url: base,
            headers: { 'Content-Type': 'application/json' }, body: { fields: op.fields } };
    });

    let resp;
    try {
      resp = await client.api('/$batch').post({ requests: requests });
    } catch (e) {
      /* Lote inteiro nao saiu: registra todos como falha e segue para o proximo.
         Abortar aqui perderia o que ainda daria certo depois. */
      for (let j = 0; j < fatia.length; j++) {
        falhas.push({ indice: i + j, status: 0, erro: e.message });
      }
      processadas += fatia.length;
      continue;
    }

    const porId = {};
    for (const r of ((resp && resp.responses) || [])) porId[String(r.id)] = r;
    for (let j = 0; j < fatia.length; j++) {
      const r = porId[String(j + 1)];
      if (r && r.status >= 200 && r.status < 300) { ok++; continue; }
      falhas.push({
        indice: i + j,
        status: r ? r.status : -1,
        erro: (r && r.body && r.body.error && r.body.error.message) || 'sem resposta no lote'
      });
    }
    processadas += fatia.length;
  }
  return { ok: ok, falhas: falhas, processadas: processadas,
           restantes: operacoes.length - processadas };
}

/* Monta a operacao de gravacao SEM executar — para o chamador juntar tudo e
   mandar em lote. Mesma regra do gravarContaOmie: o Omie manda, mas nunca apaga
   o CNPJ ja resolvido nem o vinculo com a nota. */
function prepararContaOmie(conta, indice) {
  const cod = String(conta.codigoLancamentoOmie || '');
  if (!cod) return null;

  const campos = {
    CodigoLancamentoOmie: cod,
    CodigoClienteOmie: String(conta.codigoClienteOmie || ''),
    Origem: 'omie',
    NumeroNF: String(conta.numeroNF || ''),
    NumeroParcela: String(conta.numeroParcela || ''),
    Valor: conta.valor != null ? Number(conta.valor) : null,
    DataEmissao: conta.dataEmissao || null,
    DataVencimento: conta.dataVencimento || null,
    StatusOmie: conta.statusOmie || '',
    UnidadeOmie: conta.unidade || '',
    CodigoBarras: conta.codigoBarras || '',
    SincronizadoEm: new Date().toISOString()
  };
  const ch = soDigitos(conta.chaveAcesso);
  if (chaveValida(ch)) campos.ChaveAcesso = ch;
  const cnpj = soDigitos(conta.emitenteCNPJ);
  if (cnpj.length === 14) campos.EmitenteCNPJ = cnpj;
  if (conta.emitenteNome) campos.EmitenteNome = conta.emitenteNome;

  const existente = indice ? indice[cod] : null;
  if (!existente) {
    campos.Title = 'NF ' + (conta.numeroNF || '') + ' - ' + (conta.emitenteNome || '');
    campos.Descartado = 'Nao';
    campos.VinculadoAuto = 'Nao';
    return { tipo: 'post', fields: campos, novo: true };
  }

  const antes = existente.fields || {};
  const mudou = Object.keys(campos).some(function (k) {
    if (k === 'SincronizadoEm') return false;
    const a = antes[k], b = campos[k];
    if (a == null && (b == null || b === '')) return false;
    if (k === 'Valor') return Number(a || 0) !== Number(b || 0);
    if (k === 'DataEmissao' || k === 'DataVencimento') {
      return String(a || '').substring(0, 10) !== String(b || '').substring(0, 10);
    }
    return String(a || '') !== String(b || '');
  });
  if (!mudou) return { tipo: 'nenhum' };
  return { tipo: 'patch', itemId: existente.id, fields: campos, novo: false };
}

async function vincularNota(client, siteId, docItemId, notaItemId, auto) {
  const listId = await resolveListId(client, siteId, LIST_DOCFIS);
  await client.api('/sites/' + siteId + '/lists/' + listId + '/items/' + docItemId + '/fields')
    .patch({
      NotaItemId: notaItemId == null ? '' : String(notaItemId),
      VinculadoEm: notaItemId == null ? null : new Date().toISOString(),
      VinculadoAuto: auto ? 'Sim' : 'Nao'
    });
}

/**
 * MERGE — procura um documento que corresponda a NF recem-lancada e vincula.
 *
 * ESTA E A UNICA GUARDA CONTRA DUPLICIDADE NO QUADRO. No sistema novo isso e um
 * trigger do banco e pega toda porta de entrada; aqui e uma CHAMADA. Se algum dia
 * surgir outro caminho que crie NF sem passar por aqui, o quadro passa a duplicar
 * em silencio. Quem criar esse caminho precisa chamar esta funcao.
 *
 * @returns {{ vinculou: boolean, docItemId?: string, por?: 'chave'|'fraca' }}
 */
async function casarNotaComDocumento(client, siteId, nota) {
  const listId = await resolveListId(client, siteId, LIST_DOCFIS);
  if (!listId) return { vinculou: false };

  const ch = soDigitos(nota.chaveAcesso);

  /* 1) Chave de acesso e prova. */
  if (chaveValida(ch)) {
    const achado = await buscarPorChave(client, siteId, ch);
    if (achado && !(achado.fields || {}).NotaItemId) {
      await vincularNota(client, siteId, achado.id, nota.itemId, true);
      return { vinculou: true, docItemId: achado.id, por: 'chave' };
    }
    if (achado) return { vinculou: false, docItemId: achado.id, por: 'ja_vinculado' };
  }

  /* 2) Sem chave: emitente + numero + serie. Caminho FRACO — pode casar a NF
        errada se o mesmo fornecedor repetir numero em series diferentes. Por isso
        so casa quando ha UM candidato; havendo mais, deixa para decisao humana. */
  const cnpj = soDigitos(nota.cnpjFornecedor);
  const num = String(nota.numeroNF || '').trim();
  if (!cnpj || !num) return { vinculou: false };

  const r = await client
    .api('/sites/' + siteId + '/lists/' + listId + '/items')
    .expand('fields')
    .filter("fields/EmitenteCNPJ eq '" + cnpj + "' and fields/NumeroNF eq '" + num + "'")
    .header('Prefer', 'HonorNonIndexedQueriesWarningMayFailRandomly')
    .top(20)
    .get();

  const serieNota = String(nota.serie == null ? '' : nota.serie).trim();
  const candidatos = ((r && r.value) || []).filter(function (it) {
    const f = it.fields || {};
    if (f.NotaItemId) return false;
    if (String(f.Descartado || '') === 'Sim') return false;
    return String(f.Serie == null ? '' : f.Serie).trim() === serieNota;
  });

  if (candidatos.length !== 1) {
    return { vinculou: false, ambiguo: candidatos.length > 1, candidatos: candidatos.length };
  }
  await vincularNota(client, siteId, candidatos[0].id, nota.itemId, true);
  return { vinculou: true, docItemId: candidatos[0].id, por: 'fraca' };
}

/**
 * Acha o codigo_lancamento_omie de uma nota SEM perguntar ao Omie.
 *
 * POR QUE ISTO EXISTE: o Omie nao tem filtro por vencimento nem por fornecedor
 * (os dois foram sondados e recusados pelo nome — ver DiagOmieFiltros). Sobrava
 * varrer o ListarContasPagar pagina a pagina, e o buscarContaPagar fazia isso
 * com uma janela montada em torno do VENCIMENTO usando o unico filtro de data
 * que existe, o de ALTERACAO. Em conta de vencimento longo (IPTU em 10 cotas,
 * parcela 010/013) a alteracao foi hoje e a janela esta meses a frente: a conta
 * nunca voltava e o lancamento falhava com "conta a pagar nao encontrada".
 *
 * A sincronizacao ja gravou o codigo. Procurar de novo no Omie era refazer, com
 * heuristica, um trabalho ja feito com precisao.
 *
 * TRES CAMINHOS, do mais forte ao mais fraco:
 *   NotaItemId  vinculo explicito documento<->nota. Prova, nao indicio.
 *   ChaveAcesso 44 digitos, unica no Brasil.
 *   CNPJ+numero fraco: o mesmo fornecedor repete numero em series diferentes.
 *               So aceita quando ha UM candidato — havendo mais, prefere nao
 *               achar a anexar o PDF na conta errada.
 *
 * @returns {{encontrado:boolean, codigoLancamentoOmie?:string, docItemId?:string,
 *            via?:'notaItemId'|'chave'|'cnpj+numero', motivo?:string}}
 */
async function acharCodigoOmieDaNota(client, siteId, alvo) {
  const listId = await resolveListId(client, siteId, LIST_DOCFIS);
  if (!listId) return { encontrado: false, motivo: 'lista de documentos nao existe' };

  const base = '/sites/' + siteId + '/lists/' + listId + '/items';
  async function consultar(filtro, topo) {
    const r = await client.api(base).expand('fields').filter(filtro)
      .header('Prefer', 'HonorNonIndexedQueriesWarningMayFailRandomly')
      .top(topo || 20).get();
    return (r && r.value) || [];
  }
  /* Linha sem codigo nao serve: existe no quadro mas ainda nao foi sincronizada
     com o Omie. Tratar como "achei" faria o anexo ir para lugar nenhum. */
  function comCodigo(itens) {
    return itens.filter(function (it) {
      return String((it.fields || {}).CodigoLancamentoOmie || '').trim() !== '';
    });
  }
  function devolver(it, via) {
    return { encontrado: true, via: via, docItemId: String(it.id),
             codigoLancamentoOmie: String((it.fields || {}).CodigoLancamentoOmie).trim() };
  }

  try {
    if (alvo.notaItemId) {
      const achados = comCodigo(await consultar(
        "fields/NotaItemId eq '" + String(alvo.notaItemId) + "'"));
      if (achados.length) return devolver(achados[0], 'notaItemId');
    }

    const ch = soDigitos(alvo.chave);
    if (chaveValida(ch)) {
      const achados = comCodigo(await consultar("fields/ChaveAcesso eq '" + ch + "'"));
      if (achados.length) return devolver(achados[0], 'chave');
    }

    const cnpj = soDigitos(alvo.cnpj);
    const num = String(alvo.numeroNF || '').trim();
    if (cnpj.length === 14 && num) {
      const achados = comCodigo(await consultar(
        "fields/EmitenteCNPJ eq '" + cnpj + "' and fields/NumeroNF eq '" + num + "'"));
      if (achados.length === 1) return devolver(achados[0], 'cnpj+numero');
      if (achados.length > 1) {
        return { encontrado: false, motivo: achados.length + ' documentos com o mesmo ' +
                 'CNPJ e numero — ambiguo, nao dá para escolher com seguranca' };
      }
    }
  } catch (e) {
    return { encontrado: false, motivo: 'consulta falhou: ' + ((e && e.message) || String(e)) };
  }

  return { encontrado: false, motivo: 'nenhum documento sincronizado corresponde a esta nota' };
}

/**
 * Casa documentos ORFAOS com notas ja lancadas — do lado do documento.
 *
 * O BURACO QUE ISTO TAPA: casarNotaComDocumento roda quando a NOTA e criada. Se a
 * conta chega do Omie DEPOIS de a nota existir, ninguem mais tenta casar. E o
 * quadro so reconhece o vinculo por NotaItemId ou por chave de acesso — entao
 * NFS-e, que nao tem chave, fica em "Novas" para sempre mesmo ja aprovada e paga.
 * Nao e caso isolado: sao todas as NFS-e lancadas antes da sincronizacao.
 *
 * REGRA, e ela e conservadora de proposito:
 *   chave de acesso   prova, casa direto
 *   CNPJ + numero     casa SO se houver exatamente UM candidato
 * Havendo mais de um, NAO escolhe. Vincular a nota errada faz um card exibir o
 * status de outra: uma conta em aberto apareceria como quitada, e alguem deixaria
 * de pagar. Ambiguidade volta na resposta para decisao humana.
 *
 * Casa TODAS as linhas do mesmo grupo (as parcelas de uma NF sao varias linhas):
 * o quadro escolhe uma delas como principal e e ela que precisa estar vinculada.
 *
 * NAO altera a nota, so grava o vinculo no documento. A coluna do quadro continua
 * derivada do status da nota — nao existe estado duplicado para sair de sincronia.
 */
async function casarDocumentosPendentes(client, siteId, opts) {
  const o = opts || {};
  const prazoFinal = o.prazoFinal || null;
  const listDoc = await resolveListId(client, siteId, LIST_DOCFIS);
  const listNotas = await resolveListId(client, siteId, LIST_NOTAS);
  if (!listDoc || !listNotas) return { erro: 'listas nao encontradas' };

  const { carregarNotas } = require('./notas');
  const { notas, identidade } = await carregarNotas(client, siteId, listNotas);

  /* Numero de NF comparavel: "000084" e "84" sao a mesma nota. */
  function numNorm(v) {
    const s = String(v == null ? '' : v).replace(/[^A-Za-z0-9]/g, '');
    return /^\d+$/.test(s) ? (s.replace(/^0+/, '') || '0') : s.toUpperCase();
  }

  const notaPorChave = {};
  const notasPorCnpjNum = {};
  for (const n of notas) {
    const ch = soDigitos(n.f.ChaveAcesso);
    if (ch.length === 44 && !notaPorChave[ch]) notaPorChave[ch] = n;
    const cnpj = soDigitos(n.f.CNPJFornecedor);
    const num = numNorm(n.f.NumeroNF);
    if (cnpj.length >= 11 && num) {
      const k = cnpj + '|' + num;
      if (!notasPorCnpjNum[k]) notasPorCnpjNum[k] = [];
      notasPorCnpjNum[k].push(n);
    }
  }

  const docs = await (async function () {
    const out = [];
    let url = '/sites/' + siteId + '/lists/' + listDoc + '/items?expand=fields&$top=999';
    let p = 0;
    while (url && p < 20) {
      const r = await client.api(url).get();
      out.push.apply(out, r.value || []);
      p++;
      const nl = r['@odata.nextLink'];
      url = nl ? nl.replace('https://graph.microsoft.com/v1.0', '') : null;
    }
    return out;
  })();

  /* Agrupa as linhas orfas por (CNPJ, numero) — as parcelas da mesma NF. */
  const grupos = {};
  let jaVinculados = 0, semDados = 0;
  for (const d of docs) {
    const f = d.fields || {};
    if (String(f.Descartado || '') === 'Sim') continue;
    if (f.NotaItemId) { jaVinculados++; continue; }
    const cnpj = soDigitos(f.EmitenteCNPJ);
    const num = numNorm(f.NumeroNF);
    const ch = soDigitos(f.ChaveAcesso);
    if (!ch && (!cnpj || !num)) { semDados++; continue; }
    const k = ch.length === 44 ? 'ch:' + ch : 'nf:' + cnpj + '|' + num;
    if (!grupos[k]) grupos[k] = { chave: ch, cnpj: cnpj, num: num, linhas: [] };
    grupos[k].linhas.push(d);
  }

  const ops = [];
  const ambiguos = [];
  const casados = [];
  for (const k of Object.keys(grupos)) {
    const g = grupos[k];
    let nota = null, via = null;

    if (g.chave && g.chave.length === 44 && notaPorChave[g.chave]) {
      nota = notaPorChave[g.chave]; via = 'chave';
    } else if (g.cnpj && g.num) {
      const cand = notasPorCnpjNum[g.cnpj + '|' + g.num] || [];
      if (cand.length === 1) { nota = cand[0]; via = 'cnpj+numero'; }
      else if (cand.length > 1) {
        ambiguos.push({ cnpj: g.cnpj, numeroNF: g.num, candidatos: cand.length,
                        notaIds: cand.map(function (c) { return c.id; }).slice(0, 5) });
        continue;
      }
    }
    if (!nota) continue;

    for (const linha of g.linhas) {
      ops.push({ tipo: 'patch', itemId: linha.id,
                 fields: { NotaItemId: String(nota.id), VinculadoAuto: 'Sim' } });
    }
    casados.push({ numeroNF: g.num, cnpj: g.cnpj, notaId: nota.id, via: via,
                   linhas: g.linhas.length, status: nota.f.Status || '' });
  }

  const r = ops.length
    ? await gravarEmLote(client, siteId, listDoc, ops, prazoFinal)
    : { ok: 0, falhas: [], processadas: 0, restantes: 0 };

  return {
    notasLidas: notas.length,
    mapaDeColunasEIdentidade: identidade,
    documentosOrfaos: Object.keys(grupos).length,
    jaVinculados: jaVinculados,
    semDadosParaCasar: semDados,
    gruposCasados: casados.length,
    linhasVinculadas: r.ok,
    linhasRestantes: r.restantes,
    falhas: r.falhas.length,
    ambiguos: ambiguos,
    casados: casados.slice(0, 50)
  };
}

/* ----------------------------------------------------------- ponteiro de NSU */

/**
 * Um ponteiro POR CNPJ. Com ponteiro unico, a consulta de uma filial sobrescreveria
 * o ponto de parada da outra, e os documentos no meio nunca seriam buscados — e como
 * a SEFAZ so guarda 90 dias, a perda seria definitiva e silenciosa.
 */
async function lerPonteiro(client, siteId, cnpj) {
  const listId = await resolveListId(client, siteId, LIST_SEFAZ);
  if (!listId) return null;
  const doc = soDigitos(cnpj);
  const r = await client
    .api('/sites/' + siteId + '/lists/' + listId + '/items')
    .expand('fields')
    .filter("fields/CNPJ eq '" + doc + "'")
    .header('Prefer', 'HonorNonIndexedQueriesWarningMayFailRandomly')
    .top(2)
    .get();
  const v = (r && r.value) || [];
  if (!v.length) return null;
  const f = v[0].fields || {};
  return {
    itemId: v[0].id,
    cnpj: doc,
    apelido: f.Apelido || '',
    ultimoNSU: Number(f.UltimoNSU || 0),
    maxNSU: f.MaxNSU == null ? null : Number(f.MaxNSU),
    /* Necessarios para a quarentena do cStat 656 em BuscarNFeSefaz: sem eles a
       guarda nunca dispara e uma reexecucao manual renova o bloqueio da SEFAZ. */
    cStat: f.UltimoCStat || '',
    ultimaConsulta: f.UltimaConsulta || null
  };
}

async function garantirPonteiro(client, siteId, cnpj, apelido) {
  const existente = await lerPonteiro(client, siteId, cnpj);
  if (existente) return existente;
  const listId = await resolveListId(client, siteId, LIST_SEFAZ);
  const criado = await client.api('/sites/' + siteId + '/lists/' + listId + '/items').post({
    fields: {
      Title: apelido || soDigitos(cnpj),
      CNPJ: soDigitos(cnpj),
      Apelido: apelido || '',
      UltimoNSU: 0,
      Baixados: 0
    }
  });
  return { itemId: criado.id, cnpj: soDigitos(cnpj), apelido: apelido || '',
           ultimoNSU: 0, maxNSU: null };
}

/**
 * Avanca o ponteiro. CHAMAR SO DEPOIS de gravar os documentos do lote: avancar
 * antes e perder o lote para sempre se a gravacao falhar.
 * O ponteiro NUNCA anda para tras — reexecucao ou lote fora de ordem nao pode
 * rebaixar 90 dias de documentos.
 */
async function gravarPonteiro(client, siteId, ponteiro, dados) {
  const listId = await resolveListId(client, siteId, LIST_SEFAZ);
  const novoNSU = Math.max(Number(ponteiro.ultimoNSU || 0), Number(dados.ultimoNSU || 0));
  await client.api('/sites/' + siteId + '/lists/' + listId + '/items/' + ponteiro.itemId + '/fields')
    .patch({
      UltimoNSU: novoNSU,
      MaxNSU: dados.maxNSU != null ? Number(dados.maxNSU) : null,
      UltimaConsulta: new Date().toISOString(),
      UltimoCStat: String(dados.cStat || ''),
      UltimoMotivo: String(dados.motivo || ''),
      Baixados: Number(dados.baixadosAcumulado || 0)
    });
  return novoNSU;
}

/* -------------------------------------------------- CNPJs e certificados */

/**
 * Le a App Setting SEFAZ_CNPJS. Formato:
 *   [{"cnpj":"00092929000198","apelido":"Pronep RJ","unidade":"RJ"}, ...]
 * O certificado de cada um vem de SEFAZ_CERT_<cnpj>_PFX (base64) e
 * SEFAZ_CERT_<cnpj>_SENHA — uma App Setting por CNPJ para que trocar o
 * certificado de uma filial nao mexa nas outras.
 */
function lerCnpjsConfigurados() {
  const bruto = process.env.SEFAZ_CNPJS;
  if (!bruto) return [];
  let lista;
  try { lista = JSON.parse(bruto); }
  catch (e) { throw new Error('SEFAZ_CNPJS nao e um JSON valido: ' + e.message); }
  if (!Array.isArray(lista)) throw new Error('SEFAZ_CNPJS precisa ser uma lista');
  return lista.map(function (x) {
    return {
      cnpj: soDigitos(x.cnpj),
      apelido: x.apelido || soDigitos(x.cnpj),
      unidade: x.unidade || null,
      diretoria: x.diretoria || null
    };
  }).filter(function (x) { return /^[0-9]{14}$/.test(x.cnpj); });
}

/**
 * Le o base64 do certificado das App Settings.
 *
 * ACEITA DUAS FORMAS:
 *   SEFAZ_CERT_<cnpj>_PFX              inteiro numa setting so
 *   SEFAZ_CERT_<cnpj>_PFX1, _PFX2, ... dividido em partes, concatenadas em ordem
 *
 * A segunda existe porque o painel do Azure truncou o valor na primeira tentativa
 * em producao: o .pfx chegou cortado e o OpenSSL recusou com "not enough data" —
 * erro que nao diz nada sobre a causa. Dividir em pedacos menores contorna o limite
 * sem tirar o certificado das App Settings.
 *
 * A ordem e numerica, nao alfabetica: com 10+ partes, ordenar como texto colocaria
 * _PFX10 antes de _PFX2 e o arquivo sairia embaralhado.
 */
function lerBase64Certificado(doc) {
  const inteiro = process.env['SEFAZ_CERT_' + doc + '_PFX'];
  if (inteiro) return { b64: inteiro.replace(/\s/g, ''), partes: 1 };

  const prefixo = 'SEFAZ_CERT_' + doc + '_PFX';
  const encontradas = [];
  for (const chave of Object.keys(process.env)) {
    if (chave.indexOf(prefixo) !== 0) continue;
    const sufixo = chave.slice(prefixo.length);
    if (!/^[0-9]+$/.test(sufixo)) continue;
    encontradas.push({ n: parseInt(sufixo, 10), v: process.env[chave] || '' });
  }
  if (!encontradas.length) return { b64: '', partes: 0 };
  encontradas.sort(function (a, b) { return a.n - b.n; });

  /* Numeracao com buraco (_PFX1, _PFX3) montaria um arquivo corrompido em silencio.
     Melhor falhar aqui, com o motivo, do que devolver um erro de TLS ilegivel. */
  for (let i = 0; i < encontradas.length; i++) {
    if (encontradas[i].n !== i + 1) {
      throw new Error('Partes do certificado ' + doc + ' fora de sequencia: esperava _PFX' +
                      (i + 1) + ' e achei _PFX' + encontradas[i].n +
                      '. Numere de 1 em diante, sem pular.');
    }
  }
  return {
    b64: encontradas.map(function (p) { return p.v.replace(/\s/g, ''); }).join(''),
    partes: encontradas.length
  };
}

/**
 * Chave de desligamento da busca automatica, gravada na config global (SharePoint)
 * e nao em App Setting: precisa ser desligavel pela TELA, em segundos, quando o
 * volume de NF-e virar gargalo. Mexer em App Setting reinicia as Functions e exige
 * acesso ao Azure — nao serve como freio de mao.
 *
 * AUSENCIA SIGNIFICA LIGADO. Config gravada antes desta funcionalidade nao tem a
 * chave, e o padrao seguro aqui e continuar funcionando, nao parar em silencio.
 * Falha de leitura tambem devolve ligado, pelo mesmo motivo — mas informa o erro,
 * para nao esconder que a consulta nao foi conclusiva.
 */
async function lerConfigSefaz(client, siteId) {
  try {
    const listId = await resolveListId(client, siteId, 'PRONEP-NF-Config');
    if (!listId) return { habilitado: true, origem: 'padrao' };
    const r = await client.api('/sites/' + siteId + '/lists/' + listId + '/items')
      .expand('fields').top(20).get();
    const item = ((r && r.value) || []).find(function (x) {
      return x.fields && x.fields.Title === 'global';
    });
    if (!item || !item.fields || !item.fields.ConfigJson) {
      return { habilitado: true, origem: 'padrao' };
    }
    const cfg = JSON.parse(item.fields.ConfigJson);
    if (!cfg.sefaz || typeof cfg.sefaz.habilitado !== 'boolean') {
      return { habilitado: true, origem: 'padrao' };
    }
    return {
      habilitado: cfg.sefaz.habilitado,
      motivo: cfg.sefaz.motivoDesligamento || '',
      desligadoPor: cfg.sefaz.desligadoPor || '',
      desligadoEm: cfg.sefaz.desligadoEm || '',
      origem: 'config'
    };
  } catch (e) {
    return { habilitado: true, origem: 'erro', erroLeitura: e.message };
  }
}

/**
 * Monta o material TLS do CNPJ.
 *
 * O ARQUIVO vem do Blob Storage privado; a SENHA continua em App Setting.
 * Essa separacao nao e capricho: o .pfx de ~9 KB nao cabe nas App Settings (teto
 * de 10 KB para o total), e a senha e pequena e nao tem por que sair de la.
 *
 * Fallback para App Setting: se SEFAZ_CERT_<cnpj>_PFX existir, ela vence. Serve
 * para um certificado de teste pequeno, e evita que este modulo dependa do Blob
 * Storage para sempre.
 */
/**
 * Data de corte: so entram no quadro contas com vencimento A PARTIR dela.
 *
 * Existe para o sistema nao nascer com anos de historico vencido do Omie. O
 * financeiro so quer acompanhar dali para frente.
 *
 * E FIXA, NAO MOVEL, e isso e a parte importante. "Do primeiro dia do mes
 * corrente para frente" parece equivalente e nao e: em setembro, as contas de
 * agosto ainda em aberto sairiam do quadro — exatamente as vencidas e nao pagas,
 * que sao as que mais precisam ser vistas. Uma vez definida, a data so muda se
 * alguem mudar de proposito.
 *
 * Sem configuracao gravada, assume o primeiro dia do mes ATUAL. Isso so acontece
 * antes da primeira gravacao da config; depois o valor vem de la e para de andar.
 */
function corteVencimentoPadrao() {
  const h = new Date();
  return h.getFullYear() + '-' + String(h.getMonth() + 1).padStart(2, '0') + '-01';
}

async function lerCorteVencimento(client, siteId) {
  try {
    const listId = await resolveListId(client, siteId, 'PRONEP-NF-Config');
    if (!listId) return { data: corteVencimentoPadrao(), origem: 'padrao' };
    const r = await client.api('/sites/' + siteId + '/lists/' + listId + '/items')
      .expand('fields').top(20).get();
    const item = ((r && r.value) || []).find(function (x) {
      return x.fields && x.fields.Title === 'global';
    });
    if (!item || !item.fields || !item.fields.ConfigJson) {
      return { data: corteVencimentoPadrao(), origem: 'padrao' };
    }
    const cfg = JSON.parse(item.fields.ConfigJson);
    const d = cfg.sefaz && cfg.sefaz.corteVencimento;
    if (!/^\d{4}-\d{2}-\d{2}$/.test(String(d || ''))) {
      return { data: corteVencimentoPadrao(), origem: 'padrao' };
    }
    return { data: d, origem: 'config' };
  } catch (e) {
    return { data: corteVencimentoPadrao(), origem: 'erro', erroLeitura: e.message };
  }
}

async function lerCertificado(cnpj) {
  const doc = soDigitos(cnpj);
  const senha = process.env['SEFAZ_CERT_' + doc + '_SENHA'];

  /* O fallback de App Setting so vence o blob se REALMENTE for um PKCS#12
     (SEQUENCE ASN.1: 0x30 0x82). Sem essa checagem, um valor residual — foi
     exatamente o que aconteceu na primeira configuracao, com a contagem de
     caracteres gravada no lugar do arquivo — se sobrepoe ao certificado bom e
     derruba a integracao com um erro que nao aponta para a causa. */
  const { b64 } = lerBase64Certificado(doc);
  if (b64) {
    const buf = Buffer.from(b64, 'base64');
    if (buf.length > 2 && buf[0] === 0x30 && buf[1] === 0x82) {
      return { pfx: buf, passphrase: senha || '', origem: 'appsetting' };
    }
    /* Nao explode: cai para o blob, que e a fonte principal. */
  }

  const { lerPfx } = require('./blobCert');
  let pfx;
  try {
    pfx = await lerPfx(doc);
  } catch (e) {
    throw new Error('Certificado do CNPJ ' + doc + ' indisponivel: ' + e.message);
  }
  return { pfx: pfx, passphrase: senha || '', origem: 'blob' };
}

module.exports = {
  LIST_DOCFIS, LIST_SEFAZ, LIST_NOTAS,
  COLUNAS_DOCFIS, COLUNAS_SEFAZ, COLUNAS_NOTAS_EXTRA,
  resolveListId, soDigitos, chaveValida, chaveFraca,
  buscarPorChave, gravarDocumento, vincularNota, casarNotaComDocumento,
  acharCodigoOmieDaNota, casarDocumentosPendentes,
  gravarContaOmie, indexarPorCodigoOmie, gravarEmLote, prepararContaOmie,
  lerPonteiro, garantirPonteiro, gravarPonteiro,
  lerCnpjsConfigurados, lerCertificado, lerBase64Certificado, lerConfigSefaz,
  lerCorteVencimento, corteVencimentoPadrao
};
