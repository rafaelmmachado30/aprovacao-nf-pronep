/**
 * /api/DiagOmieFiltros  (GET) — ADMIN. Descoberta, nao producao.
 *
 * Descobre QUAIS filtros o ListarContasPagar realmente aceita, testando um por
 * vez e lendo a resposta do Omie.
 *
 * POR QUE ISTO EXISTE:
 * o DiagOmieContasPagar mostrou que `filtrar_por_data_de/ate` NAO filtra por
 * vencimento — pedimos a janela 22/06 a 20/09 e voltaram contas com vencimento
 * em 20/03, todas com data de ALTERACAO dentro da janela. Ou seja, o filtro e
 * por data de registro/alteracao.
 *
 * Isso tem duas consequencias:
 *   1. shared/omie.js documenta e usa esse filtro COMO SE fosse vencimento, no
 *      buscarContaPagar que o IntegrarOmie chama em producao. Precisa ser
 *      corrigido, mas so depois de saber qual e o parametro certo.
 *   2. O quadro precisa de "em aberto por vencimento", nao de "tudo que mudou".
 *      Com 1598 registros no RJ mas apenas 5 A VENCER, o filtro certo transforma
 *      o volume de problema em nao-problema.
 *
 * O Omie devolve `faultstring` quando o parametro nao existe, entao testar e
 * seguro: parametro invalido vira erro identificado, nao dado errado.
 *
 * NAO grava nada e nao altera nada no Omie.
 */

require('isomorphic-fetch');
const { getCredentials } = require('../shared/omie');

const OMIE_BASE = 'https://app.omie.com.br/api/v1';

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

async function tentar(param, creds) {
  const resp = await fetch(OMIE_BASE + '/financas/contapagar/', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', 'Accept': 'application/json',
               'User-Agent': 'PronepNF/1.0 (Azure SWA Functions)' },
    body: JSON.stringify({
      call: 'ListarContasPagar',
      app_key: creds.appKey, app_secret: creds.appSecret,
      param: [Object.assign({ pagina: 1, registros_por_pagina: 20, apenas_importado_api: 'N' }, param)]
    })
  });
  const texto = await resp.text();
  let data;
  try { data = JSON.parse(texto); }
  catch (e) { return { aceito: false, erro: 'resposta nao-JSON: ' + texto.slice(0, 120) }; }
  if (data && data.faultstring) return { aceito: false, erro: String(data.faultstring).slice(0, 180) };
  if (!resp.ok) return { aceito: false, erro: 'HTTP ' + resp.status };

  const itens = data.conta_pagar_cadastro || data.contas_pagar_cadastro || [];
  const status = {};
  const vencs = [];
  for (const c of itens) {
    status[String(c.status_titulo || '?')] = (status[String(c.status_titulo || '?')] || 0) + 1;
    if (c.data_vencimento) vencs.push(c.data_vencimento);
  }
  return {
    aceito: true,
    totalRegistros: data.total_de_registros,
    amostraStatus: status,
    /* Os vencimentos provam se o filtro realmente mordeu: se pedi uma janela e
       voltou coisa fora dela, o parametro foi ACEITO mas ignorado — que e pior
       do que ser recusado, porque parece que funcionou. */
    vencimentos: vencs.slice(0, 6)
  };
}

module.exports = async function (context, req) {
  const t0 = Date.now();
  const u = String((req.query || {}).unidade || 'RJ').toUpperCase();
  const diag = { unidade: u, testes: {}, timeMs: 0 };

  try {
    if (!(await isAdmin(req))) {
      context.res = { status: 403, headers: { 'Content-Type': 'application/json' },
        body: { error: 'Apenas admin' } };
      return;
    }
    const creds = getCredentials(u);
    diag.empresa = creds.empresa;

    const hoje = new Date();
    const f = function (d) {
      return String(d.getDate()).padStart(2, '0') + '/' +
             String(d.getMonth() + 1).padStart(2, '0') + '/' + d.getFullYear();
    };
    const de = f(new Date(hoje.getTime() - 5 * 86400000));
    const ate = f(new Date(hoje.getTime() + 60 * 86400000));
    diag.janelaTeste = { de: de, ate: ate };

    /* Candidatos, do mais util para o quadro ao menos util. Nomes conforme o
       padrao do lcpListarRequest; o que nao existir volta como faultstring. */
    const candidatos = [
      ['sem_filtro',                {}],
      ['filtrar_por_vencimento',    { filtrar_por_vencimento_de: de, filtrar_por_vencimento_ate: ate }],
      ['filtrar_por_status_ABERTO', { filtrar_por_status: 'ABERTO' }],
      ['filtrar_por_status_AVENCER',{ filtrar_por_status: 'A VENCER' }],
      ['filtrar_por_emissao',       { filtrar_por_emissao_de: de, filtrar_por_emissao_ate: ate }],
      ['filtrar_por_pagamento',     { filtrar_por_pagamento_de: de, filtrar_por_pagamento_ate: ate }],
      ['filtrar_apenas_alteracao',  { filtrar_por_data_de: de, filtrar_por_data_ate: ate, filtrar_apenas_alteracao: 'S' }],
      ['filtrar_apenas_inclusao',   { filtrar_por_data_de: de, filtrar_por_data_ate: ate, filtrar_apenas_inclusao: 'S' }],
      ['ordenar_por_vencimento',    { ordenar_por: 'DATA_VENCIMENTO', ordem_decrescente: 'N' }]
    ];

    for (const [nome, param] of candidatos) {
      try { diag.testes[nome] = await tentar(param, creds); }
      catch (e) { diag.testes[nome] = { aceito: false, erro: e.message }; }
    }

    diag.timeMs = Date.now() - t0;
    context.res = { status: 200, headers: { 'Content-Type': 'application/json' },
      body: Object.assign({ ok: true,
        leitura: 'aceito=true com "vencimentos" DENTRO da janelaTeste = o filtro serve. ' +
                 'aceito=true com vencimentos fora = parametro ignorado pelo Omie.' }, diag) };
  } catch (err) {
    diag.timeMs = Date.now() - t0;
    context.res = { status: 500, headers: { 'Content-Type': 'application/json' },
      body: Object.assign({ error: (err && err.message) || String(err) }, diag) };
  }
};
