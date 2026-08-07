/**
 * /api/DiagOmieRecebimento  (GET) — ADMIN. Descoberta, nao producao.
 *
 * Acha o formato da API de RECEBIMENTO DE NF-e do Omie — o modulo da tela que
 * mostra itens, totais e os botoes "Exibir DANFE" / "Exibir XML".
 *
 * POR QUE ISTO E O CAMINHO CERTO:
 * o DiagOmieAnexos ja eliminou os outros. Conta a pagar NAO guarda anexo
 * (qtdAnexos 0 em todas as testadas), /produtos/xml/ e /produtos/dfedocsfiscais/
 * dao 404, /produtos/nfe/ da 500. Sobrou UM candidato vivo:
 *
 *     /produtos/recebimentonfe/  call ListarRecebimentos
 *     -> "Tag [PAGINA] nao faz parte da estrutura do tipo complexo
 *         [rcbtoListarRequest]"
 *
 * Recusar UM parametro e diferente de recusar a chamada: o endpoint existe, o
 * call existe, e so o nome do campo esta errado. E a mensagem de erro do Omie
 * nomeia a tag ofensora — ou seja, ela e um oraculo. Sondar e barato e converge.
 *
 * COMO SONDA: manda {} primeiro. Sem campo nenhum, o Omie costuma responder
 * dizendo qual tag FALTA, o que da o nome certo de graca. Depois testa as
 * convencoes plausiveis do prefixo `n`/`c`/`d` que o resto da API usa.
 *
 * SO LE. A lista de calls e um whitelist de verbos de leitura — Incluir/Alterar/
 * Excluir nao entram aqui nem por engano: este arquivo roda contra a base de
 * producao da Pronep, e um call de escrita sondado "para ver o que responde"
 * pode alterar documento fiscal de verdade.
 *
 * Query:
 *   ?unidade=RJ    RJ | SP | ES  (default RJ)
 *   ?chave=<44>    chave de NF-e pra testar o Obter (default: pega uma do quadro)
 */

require('isomorphic-fetch');
const { getCredentials } = require('../shared/omie');

const OMIE_BASE = 'https://app.omie.com.br/api/v1';
const EP = '/produtos/recebimentonfe/';

/* O Omie derruba chamada identica repetida ("Consumo redundante, aguarde 44s").
   Cada sonda daqui manda param diferente, mas o espacamento evita esbarrar no
   limite por rajada — sondar rapido demais desperdicaria a execucao inteira. */
const PAUSA_MS = 400;
const ORCAMENTO_MS = 30000;

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

/* Trava dura: nenhum call que nao comece por um verbo de leitura sai daqui.
   Nao e paranoia — e a diferenca entre um diagnostico e um incidente fiscal. */
const VERBOS_LEITURA = /^(Listar|Obter|Consultar|Pesquisar)/;

async function sondar(call, param, creds) {
  if (!VERBOS_LEITURA.test(call)) {
    return { aceito: false, erro: 'BLOQUEADO: ' + call + ' nao e call de leitura' };
  }
  let resp, texto;
  try {
    resp = await fetch(OMIE_BASE + EP, {
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
  catch (e) { return { aceito: false, erro: 'HTTP ' + resp.status + ' nao-JSON: ' + texto.slice(0, 140) }; }
  if (data && data.faultstring) {
    const fs = String(data.faultstring);
    /* A tag citada e o achado: e ela que diz como o campo se chama de verdade. */
    const m = /\[([A-Z_]+)\]/.exec(fs);
    return { aceito: false, erro: fs.slice(0, 220), tagCitada: m ? m[1] : null };
  }
  if (!resp.ok) return { aceito: false, erro: 'HTTP ' + resp.status };
  return { aceito: true, data: data };
}

module.exports = async function (context, req) {
  const t0 = Date.now();
  const q = req.query || {};
  const u = String(q.unidade || 'RJ').toUpperCase();
  const chave = String(q.chave || '').replace(/\D/g, '');
  const diag = { unidade: u, endpoint: EP, sondas: {}, timeMs: 0 };

  try {
    if (!(await isAdmin(req))) {
      context.res = { status: 403, headers: { 'Content-Type': 'application/json' },
        body: { error: 'Apenas admin' } };
      return;
    }
    const creds = getCredentials(u);
    diag.empresa = creds.empresa;

    /* ---------- FASE 1: descobrir o formato do Listar ----------
       {} primeiro: sem campo nenhum o Omie tende a reclamar do que FALTA, e a
       reclamacao entrega o nome certo sem eu precisar adivinhar. */
    const formatos = [
      ['vazio',              {}],
      ['nPagina_nRegPorPagina',      { nPagina: 1, nRegPorPagina: 5 }],
      ['nPagina_nRegistrosPorPagina',{ nPagina: 1, nRegistrosPorPagina: 5 }],
      ['nPag_nRegPorPag',            { nPag: 1, nRegPorPag: 5 }],
      ['nPagina_so',                 { nPagina: 1 }]
    ];

    let formatoBom = null;
    for (const [nome, param] of formatos) {
      if (Date.now() - t0 > ORCAMENTO_MS) { diag.sondas[nome] = { pulado: 'sem tempo' }; continue; }
      const r = await sondar('ListarRecebimentos', param, creds);
      diag.sondas['Listar__' + nome] = r.aceito
        ? { aceito: true, chaves: Object.keys(r.data || {}).slice(0, 15),
            tamanho: JSON.stringify(r.data || {}).length }
        : r;
      if (r.aceito) { formatoBom = { nome: nome, param: param, data: r.data }; break; }
      await dorme(PAUSA_MS);
    }

    /* ---------- FASE 2: se o Listar respondeu, o que vem dentro? ----------
       A pergunta real nao e "a API existe", e "ela entrega o XML/DANFE". Um
       Listar que so devolve id e data nao resolve nada — o detalhe da nota e o
       que o Rafael quer ver no card. */
    if (formatoBom) {
      diag.formatoQueFunciona = formatoBom.nome;
      const d = formatoBom.data || {};
      const lista = d.recebimentos || d.listaRecebimentos || d.cadastros ||
                    d.rcbto_lista || d.registros || null;
      diag.total = d.nTotRegistros || d.total_de_registros || null;
      if (Array.isArray(lista) && lista.length) {
        diag.camposDoRegistro = Object.keys(lista[0]).sort();
        /* Nao despeja o XML inteiro: so diz SE veio e o tamanho. */
        const primeiro = lista[0];
        diag.temSinalDeXML = Object.keys(primeiro).filter(function (k) {
          return /xml|danfe|chave|arquivo|pdf|url/i.test(k);
        });
        diag.amostraRegistro = JSON.parse(JSON.stringify(primeiro, function (k, v) {
          return (typeof v === 'string' && v.length > 300) ? '[' + v.length + ' chars]' : v;
        }));
      } else {
        diag.avisoLista = 'Listar aceito, mas nao achei array de registros. ' +
                          'Chaves: ' + Object.keys(d).join(', ');
      }
    }

    /* ---------- FASE 3: existe um Obter por chave? ----------
       E este que interessa de verdade: o card tem a chave, precisa do detalhe.
       Roda mesmo sem o Listar ter fechado — sao perguntas independentes. */
    if (chave.length === 44 && Date.now() - t0 < ORCAMENTO_MS) {
      const porChave = [
        ['ObterRecebimento_cChaveNFe',    'ObterRecebimento',    { cChaveNFe: chave }],
        ['ObterRecebimento_nChaveNFe',    'ObterRecebimento',    { nChaveNFe: chave }],
        ['ConsultarRecebimento_cChaveNFe','ConsultarRecebimento',{ cChaveNFe: chave }]
      ];
      for (const [nome, call, param] of porChave) {
        if (Date.now() - t0 > ORCAMENTO_MS) { diag.sondas[nome] = { pulado: 'sem tempo' }; continue; }
        const r = await sondar(call, param, creds);
        diag.sondas[nome] = r.aceito
          ? { aceito: true, chaves: Object.keys(r.data || {}).slice(0, 20),
              tamanho: JSON.stringify(r.data || {}).length }
          : r;
        if (r.aceito) { diag.obterQueFunciona = nome; break; }
        await dorme(PAUSA_MS);
      }
    } else if (chave.length !== 44) {
      diag.avisoChave = 'Passe ?chave=<44 digitos> pra testar o Obter por chave.';
    }

    diag.timeMs = Date.now() - t0;
    context.res = { status: 200, headers: { 'Content-Type': 'application/json' },
      body: Object.assign({ ok: true,
        leitura: 'Olhe "tagCitada" em cada sonda: e o nome que o Omie espera. ' +
                 'Se algum aceito=true trouxer campo de xml/danfe/chave, o detalhe ' +
                 'da NF vem do Omie e nao precisamos da SEFAZ.' }, diag) };
  } catch (err) {
    diag.timeMs = Date.now() - t0;
    context.res = { status: 500, headers: { 'Content-Type': 'application/json' },
      body: Object.assign({ error: (err && err.message) || String(err) }, diag) };
  }
};
