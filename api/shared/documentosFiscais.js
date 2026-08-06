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
  { name: 'MotivoDescarte', def: { text: { allowMultipleLines: true } } }
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
    maxNSU: f.MaxNSU == null ? null : Number(f.MaxNSU)
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
  lerPonteiro, garantirPonteiro, gravarPonteiro,
  lerCnpjsConfigurados, lerCertificado, lerBase64Certificado, lerConfigSefaz
};
