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
  const forn = {};
  for (const c of itens) {
    status[String(c.status_titulo || '?')] = (status[String(c.status_titulo || '?')] || 0) + 1;
    if (c.data_vencimento) vencs.push(c.data_vencimento);
    forn[String(c.codigo_cliente_fornecedor || '?')] = true;
  }
  const codigos = Object.keys(forn);
  return {
    aceito: true,
    totalRegistros: data.total_de_registros,
    amostraStatus: status,
    /* Os vencimentos provam se o filtro realmente mordeu: se pedi uma janela e
       voltou coisa fora dela, o parametro foi ACEITO mas ignorado — que e pior
       do que ser recusado, porque parece que funcionou. */
    vencimentos: vencs.slice(0, 6),
    /* Mesma logica para o filtro de fornecedor: um so codigo na amostra e sinal
       de que mordeu; varios significam parametro ignorado. */
    fornecedoresNaAmostra: codigos.slice(0, 6),
    umFornecedorSo: codigos.length === 1,
    primeiroFornecedor: codigos[0] || null
  };
}

/* Fecha antes dos ~45s da plataforma. Sem isto a Function e morta no meio e o
   navegador recebe "Backend call failure", que nao diz nada sobre a sondagem. */
const ORCAMENTO_MS = 25000;

module.exports = async function (context, req) {
  const t0 = Date.now();
  const q = req.query || {};
  const u = String(q.unidade || 'RJ').toUpperCase();
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
    /* Janela de 7 dias para medir o fluxo diario: e ela que define a cadencia do
       cron incremental. */
    const de7 = f(new Date(hoje.getTime() - 7 * 86400000));
    const ate0 = f(hoje);
    diag.janelaTeste = { de: de, ate: ate };
    diag.janela7dias = { de: de7, ate: ate0 };

    /* Candidatos, do mais util para o quadro ao menos util. Nomes conforme o
       padrao do lcpListarRequest; o que nao existir volta como faultstring. */
    /* Rodada 2. A rodada 1 estabeleceu que filtrar_por_vencimento e
       filtrar_por_pagamento NAO existem, que ordem_decrescente nao existe, e —
       pela propria mensagem de recusa do Omie — o vocabulario valido de
       filtrar_por_status. Agora falta a MEDIDA: quantas contas cada status
       devolve. E ela que decide se a sincronizacao do quadro cabe numa execucao
       da Function (~45s, ~60 req/min no Omie) ou precisa de ponteiro.
       'EMABERTO' e o candidato a consulta principal do quadro. */
    const candidatos = [
      ['status_EMABERTO',      { filtrar_por_status: 'EMABERTO' }],
      ['status_AVENCER',       { filtrar_por_status: 'AVENCER' }],
      ['status_ATRASADO',      { filtrar_por_status: 'ATRASADO' }],
      ['status_VENCEHOJE',     { filtrar_por_status: 'VENCEHOJE' }],
      ['status_PAGTO_PARCIAL', { filtrar_por_status: 'PAGTO_PARCIAL' }],
      ['status_PAGO',          { filtrar_por_status: 'PAGO' }],
      ['ordenar_por_sozinho',  { ordenar_por: 'DATA_VENCIMENTO' }],
      ['inclusao_7dias',       { filtrar_por_data_de: de7, filtrar_por_data_ate: ate0, filtrar_apenas_inclusao: 'S' }],
      ['alteracao_7dias',      { filtrar_por_data_de: de7, filtrar_por_data_ate: ate0, filtrar_apenas_alteracao: 'S' }]
    ];

    /* AS RODADAS 1 E 2 JA ESTAO RESPONDIDAS e por isso nao rodam mais por padrao:
       empilhar tudo numa execucao passou dos ~45s da plataforma e voltou
       "Backend call failure" — 14 chamadas sequenciais ao Omie nao cabem.
       Quem quiser remedir pede ?rodada=status. */
    if (String(q.rodada || '') === 'status') {
      for (const [nome, param] of candidatos) {
        if (Date.now() - t0 > ORCAMENTO_MS) { diag.testes[nome] = { pulado: 'sem tempo' }; continue; }
        try { diag.testes[nome] = await tentar(param, creds); }
        catch (e) { diag.testes[nome] = { aceito: false, erro: e.message }; }
      }
    } else {
      diag.rodadasAnteriores = 'puladas (use ?rodada=status para remedir)';
    }

    /* RODADA 3 — FILTRO POR FORNECEDOR. Isto conserta um bug real.
       buscarContaPagar procura a conta de UMA nota montando uma janela de data
       em torno do VENCIMENTO, mas o unico filtro de data que existe e o de
       ALTERACAO. Numa conta de vencimento longo (IPTU em 10 cotas, parcela
       010/013) a alteracao foi hoje e a janela esta meses a frente: a conta nao
       volta, e o lancamento falha com "conta a pagar nao encontrada".
       Filtrar por fornecedor eliminaria a adivinhacao de data por completo — o
       conjunto vira "as contas deste fornecedor", que e pequeno e exato.
       O codigo do fornecedor sai de uma conta real da propria base: inventar um
       numero testaria o parser do Omie, nao o filtro. */
    const semFiltro = await tentar({}, creds);
    const codForn = semFiltro.primeiroFornecedor;
    diag.fornecedorDeTeste = codForn;
    if (!codForn) {
      diag.avisoFornecedor = 'Nao achei nenhuma conta para extrair um codigo de fornecedor.';
    } else {
      const porForn = [
        ['forn_filtrar_por_cliente',        { filtrar_por_cliente: Number(codForn) }],
        ['forn_codigo_cliente_fornecedor',  { codigo_cliente_fornecedor: Number(codForn) }],
        ['forn_filtrar_por_fornecedor',     { filtrar_por_fornecedor: Number(codForn) }],
        ['forn_filtrar_cliente_fornecedor', { filtrar_cliente_fornecedor: Number(codForn) }]
      ];
      for (const [nome, param] of porForn) {
        if (Date.now() - t0 > ORCAMENTO_MS) { diag.testes[nome] = { pulado: 'sem tempo' }; continue; }
        try {
          const r = await tentar(param, creds);
          /* ACEITO NAO BASTA. Se voltou mais de um fornecedor na amostra, o Omie
             engoliu o parametro e devolveu tudo — usar isso como filtro daria a
             falsa sensacao de precisao. */
          if (r.aceito) {
            r.filtroMordeu = r.umFornecedorSo && r.fornecedoresNaAmostra[0] === String(codForn);
            if (!r.filtroMordeu) r.alerta = 'ACEITO PORÉM IGNORADO — nao use como filtro';
          }
          diag.testes[nome] = r;
        } catch (e) { diag.testes[nome] = { aceito: false, erro: e.message }; }
      }
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
