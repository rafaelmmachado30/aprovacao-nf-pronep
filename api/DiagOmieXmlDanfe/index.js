/**
 * /api/DiagOmieXmlDanfe  (GET) — ADMIN. Descoberta, nao producao.
 *
 * PROCURA O XML E A DANFE NA API DO OMIE. Eles EXISTEM — isto nao e hipotese.
 *
 * O Rafael abriu o Recebimento de NF-e na tela do Omie e clicou em "Exibir DANFE
 * do Fornecedor" e "Exibir XML da NF-e". Os dois carregaram:
 *
 *   .../resources/temp/<sessao>/php/<chave>.pdf?SrvCheck=...
 *   .../resources/temp/<sessao>/<chave>-procnfe.xml?SrvCheck=...
 *
 * E o XML e um <nfeProc versao="4.00"> completo, COM protocolo de autorizacao —
 * exatamente o campo que eu tinha declarado impossivel de obter, e que e o unico
 * que separa um espelho de uma DANFE de verdade.
 *
 * ONDE EU ERREI, PARA NAO ERRAR DE NOVO
 * O DiagOmieAnexos testou /produtos/dfedocsfiscais/ e levou 404. O nome real do
 * servico e /produtos/dfedocs/ — "Disponibiliza PDF e XML de documentos fiscais".
 * Uma palavra de diferenca, e eu concluí "o Omie nao tem o XML" a partir de um
 * 404 que so dizia "esse endereco nao existe". 404 em endpoint chutado nao e
 * evidencia sobre o dado; e evidencia sobre o chute.
 *
 * OS QUATRO CANDIDATOS, E POR QUE CADA UM
 *   /produtos/dfedocs/       diz explicitamente "disponibiliza PDF e XML"
 *   /produtos/notafiscalutil/ "recupera URL da NF-e (XML), do Danfe ou do logo"
 *   /contador/xml/           "listagem dos XMLs de documentos fiscais" — e o
 *                            unico do Painel do Contador, que precisa de ENTRADA
 *                            e saida; os outros vivem em modulos de venda e
 *                            podem so conhecer o que a empresa EMITIU. Nossa
 *                            nota e de fornecedor: entrada.
 *   /produtos/recebimentonfe/ o modulo que ja responde, testado com calls de
 *                            arquivo em vez de conteudo, e tambem por nIdReceb
 *                            (que vem no cabec) alem da chave.
 *
 * SO LE. Whitelist de verbos por regex, igual as sondas anteriores: isto roda
 * contra a base de producao da Pronep, e um call de escrita sondado "para ver o
 * que responde" pode alterar documento fiscal de verdade.
 *
 * Query:
 *   ?unidade=RJ    RJ | SP | ES (default RJ)
 *   ?chave=<44>    obrigatorio — a nota a procurar
 *   ?nIdReceb=N    opcional; se vier, testa tambem os calls por id de recebimento
 */

require('isomorphic-fetch');
const { getCredentials } = require('../shared/omie');

const OMIE_BASE = 'https://app.omie.com.br/api/v1';
const PAUSA_MS = 350;
const ORCAMENTO_MS = 38000;

function dorme(ms) { return new Promise(function (r) { setTimeout(r, ms); }); }

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

const VERBOS_LEITURA = /^(Listar|Obter|Consultar|Pesquisar|Exibir|Baixar|Download)/;

/* O achado nunca e o XML inteiro na resposta do diagnostico — seria despejar
   documento fiscal no log. O que interessa e: veio? de que tamanho? em que
   campo? e, se for URL, qual o padrao dela. */
function resumir(v, prof) {
  if (v == null) return null;
  if (typeof v === 'string') {
    const ehUrl = /^https?:\/\//i.test(v);
    const ehXml = /^\s*<\?xml|<nfeProc|<NFe/i.test(v);
    const ehB64 = v.length > 200 && /^[A-Za-z0-9+/=\s]+$/.test(v.slice(0, 200));
    if (ehUrl) return { tipo: 'URL', valor: v.slice(0, 300) };
    if (ehXml) return { tipo: 'XML', bytes: v.length, inicio: v.slice(0, 120) };
    if (ehB64) return { tipo: 'BASE64?', bytes: v.length };
    return v.length > 160 ? { tipo: 'texto', bytes: v.length, inicio: v.slice(0, 120) } : v;
  }
  if (Array.isArray(v)) {
    return (prof || 0) > 2 ? '[array ' + v.length + ']'
      : { array: v.length, primeiro: v.length ? resumir(v[0], (prof || 0) + 1) : null };
  }
  if (typeof v === 'object') {
    if ((prof || 0) > 2) return '{...}';
    const o = {};
    for (const k of Object.keys(v)) o[k] = resumir(v[k], (prof || 0) + 1);
    return o;
  }
  return v;
}

async function sondar(endpoint, call, param, creds) {
  if (!VERBOS_LEITURA.test(call)) {
    return { aceito: false, erro: 'BLOQUEADO: ' + call + ' nao e call de leitura' };
  }
  let resp, texto;
  try {
    resp = await fetch(OMIE_BASE + endpoint, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'Accept': 'application/json',
                 'User-Agent': 'PronepNF/1.0 (Azure SWA Functions)' },
      body: JSON.stringify({ call: call, app_key: creds.appKey,
                             app_secret: creds.appSecret, param: [param] })
    });
    texto = await resp.text();
  } catch (e) {
    return { aceito: false, erro: 'rede: ' + ((e && e.message) || String(e)) };
  }
  let data;
  try { data = JSON.parse(texto); }
  catch (e) {
    return { aceito: false, erro: 'HTTP ' + resp.status + ' nao-JSON: ' + texto.slice(0, 120) };
  }
  if (data && data.faultstring) {
    const fs = String(data.faultstring);
    const m = /\[([A-Z_]+)\]/.exec(fs);
    /* Tres "nao" bem diferentes, e confundi-los foi o que me custou esta rodada:
       - endpoint/call inexistente  => o caminho esta errado, nada se conclui do dado
       - tag recusada               => o caminho esta CERTO, so o campo tem outro nome
       - registro nao encontrado    => caminho e campo certos, a nota e que nao esta la */
      const tipo = /m[eé]todo|method|servi[cç]o|n[aã]o (existe|encontrado|dispon)/i.test(fs) && !m
        ? 'call_ou_endpoint_inexistente'
        : m ? 'tag_recusada' : 'outro';
    return { aceito: false, erro: fs.slice(0, 200), tagCitada: m ? m[1] : null, tipo: tipo };
  }
  /* Medido na rodada 1: o Omie devolve HTTP 500 SEM corpo util quando o call nao
     existe naquele endpoint. Nomear isso evita a leitura que ja me custou uma
     conclusao errada — 500 aqui nao e "o Omie caiu" nem "o dado nao existe", e
     "esse call que eu inventei nao existe". */
  if (resp.status === 500) {
    return { aceito: false, erro: 'HTTP 500', tipo: 'call_ou_endpoint_inexistente' };
  }
  if (!resp.ok) return { aceito: false, erro: 'HTTP ' + resp.status };
  return { aceito: true, chaves: Object.keys(data || {}), resumo: resumir(data),
           bytes: JSON.stringify(data || {}).length };
}

module.exports = async function (context, req) {
  const t0 = Date.now();
  const q = req.query || {};
  const u = String(q.unidade || 'RJ').toUpperCase();
  const chave = String(q.chave || '').replace(/\D/g, '');
  const nIdReceb = String(q.nIdReceb || '').replace(/\D/g, '');
  const out = { unidade: u, chave: chave, sondas: {}, timeMs: 0 };

  try {
    if (!(await isAdmin(req))) {
      context.res = { status: 403, headers: { 'Content-Type': 'application/json' },
        body: { error: 'Apenas admin' } };
      return;
    }
    if (chave.length !== 44) {
      context.res = { status: 400, headers: { 'Content-Type': 'application/json' },
        body: { error: 'Passe ?chave=<44 digitos>. Use a mesma nota que voce abriu ' +
                       'na tela do Omie, para poder comparar o que volta.' } };
      return;
    }
    const creds = getCredentials(u);
    out.empresa = creds.empresa;

    /* Cada entrada: [rotulo, endpoint, call, param].
       O `{}` vazio vem primeiro em cada endpoint de proposito: sem campo nenhum o
       Omie tende a reclamar do que FALTA, e a reclamacao entrega o nome certo do
       parametro sem eu precisar adivinhar — foi assim que o recebimentonfe caiu. */
    /* RODADA 2 — agora com os nomes lidos na documentacao, nao chutados.
       A rodada 1 levou HTTP 500 em quase tudo, e 500 aqui significa CALL
       INEXISTENTE: eu tinha inventado ObterDocumentos, ObterXML, ExibirDANFE.
       Sobreviveu um: /contador/xml/ ListarDocumentos, que recusou a tag citando
       o tipo [xmlListarDocumentosRequest] — ou seja, o call existe e so o campo
       tinha outro nome. Na doc do servico:
         request  nChave (string44), nIdReceb (integer), cOperacao, cModelo
         resposta documentosEncontrados[] com cXml, nIdNF, nIdReceb, nChave
       `cXml` e o XML do documento fiscal. E `nIdNF` alimenta o segundo passo:
       /produtos/dfedocs/ ObterNfe recebe nIdNfe e devolve cPdf — o link da DANFE
       ja renderizada, que e o mesmo arquivo que o botao da tela do Omie abre. */
    const alvos = [
      /* --- passo 1: o XML, por chave --- */
      ['contador__porChave',   '/contador/xml/', 'ListarDocumentos',
        { nPagina: 1, nRegPorPagina: 5, nChave: chave }],
      /* Sem cModelo: filtrar por 55 antes de saber se a nota esta na base
         confundiria "nao e modelo 55" com "nao esta la". */
      ['contador__porChave_55', '/contador/xml/', 'ListarDocumentos',
        { nPagina: 1, nRegPorPagina: 5, nChave: chave, cModelo: '55' }],
      /* cOperacao decide se a base do contador cobre ENTRADA. Se so houver saida,
         nenhuma nota de fornecedor aparece e o caminho morre aqui. */
      ['contador__entrada',    '/contador/xml/', 'ListarDocumentos',
        { nPagina: 1, nRegPorPagina: 3, cOperacao: 'E' }],
      ['contador__saida',      '/contador/xml/', 'ListarDocumentos',
        { nPagina: 1, nRegPorPagina: 3, cOperacao: 'S' }],

      /* --- passo 2: o PDF da DANFE, pelos calls que a doc REALMENTE lista --- */
      ['dfedocs__ObterNfe_id0', '/produtos/dfedocs/', 'ObterNfe', { nIdNfe: 0 }]
    ];

    /* nIdReceb vem no cabec do ConsultarRecebimento que ja usamos. Se a chave nao
       achar, o id do recebimento e o filtro mais direto que existe para entrada. */
    if (nIdReceb) {
      alvos.push(['contador__porIdReceb', '/contador/xml/', 'ListarDocumentos',
                  { nPagina: 1, nRegPorPagina: 5, nIdReceb: Number(nIdReceb) }]);
    }

    const endpointsMortos = {};
    for (const [rotulo, endpoint, call, param] of alvos) {
      if (Date.now() - t0 > ORCAMENTO_MS) { out.sondas[rotulo] = { pulado: 'sem tempo' }; continue; }
      /* Se o endpoint inteiro nao existe, nao gasta as outras sondas dele — o
         orcamento e melhor usado nos que respondem. */
      if (endpointsMortos[endpoint]) {
        out.sondas[rotulo] = { pulado: 'endpoint ja recusado: ' + endpointsMortos[endpoint] };
        continue;
      }
      const r = await sondar(endpoint, call, param, creds);
      out.sondas[rotulo] = Object.assign({ endpoint: endpoint, call: call }, r);
      if (!r.aceito && /HTTP 404|nao-JSON/.test(String(r.erro || ''))) {
        endpointsMortos[endpoint] = r.erro;
      }
      await dorme(PAUSA_MS);
    }

    /* Resumo no topo: o que responde e, dentre esses, quem trouxe arquivo. */
    const vivos = Object.keys(out.sondas).filter(function (k) { return out.sondas[k].aceito; });
    out.responderam = vivos;
    out.comArquivo = vivos.filter(function (k) {
      return /"tipo":"(URL|XML|BASE64\?)"/.test(JSON.stringify(out.sondas[k].resumo || {}));
    });
    out.leitura =
      'comArquivo e a resposta: qualquer entrada ali devolveu URL, XML cru ou base64 — ' +
      'e entao a DANFE sai do documento ORIGINAL, com protocolo, e o espelho vira ' +
      'desnecessario. Se comArquivo vier vazio mas `responderam` nao, o caminho existe ' +
      'e falta o parametro certo: olhe tagCitada em cada sonda recusada, que e o Omie ' +
      'dizendo o nome real do campo. tipo=call_ou_endpoint_inexistente significa que ' +
      'aquele chute morreu e nao diz NADA sobre o dado existir.';

    out.timeMs = Date.now() - t0;
    context.res = { status: 200, headers: { 'Content-Type': 'application/json' },
      body: Object.assign({ ok: true }, out) };
  } catch (err) {
    out.timeMs = Date.now() - t0;
    context.res = { status: 500, headers: { 'Content-Type': 'application/json' },
      body: Object.assign({ error: (err && err.message) || String(err) }, out) };
  }
};
