/**
 * /api/DiagOmieAnexos  (GET) — ADMIN. Descoberta, nao producao.
 *
 * Pergunta ao Omie se ele guarda o DANFE e o XML como ANEXO da conta a pagar.
 *
 * POR QUE ISTO DECIDE O DESENHO:
 * o quadro tem a CHAVE da NF-e, nao o documento. Para o card abrir a NF em PDF,
 * existem tres caminhos, em ordem de custo:
 *   1. o Omie ja tem o arquivo anexado  -> so baixar (barato, e o que se testa aqui)
 *   2. buscar o XML na SEFAZ por chave  -> temos certificado e codigo provado,
 *      mas o XML nao e o PDF: gerar a DANFE a partir dele e um projeto a parte
 *      (layout fiscal completo, codigo de barras, regras de impressao)
 *   3. o usuario anexa o PDF a mao      -> sempre funciona, mas e trabalho manual
 * Sem saber se (1) existe, qualquer decisao aqui e chute.
 *
 * O codebase ja usa IncluirAnexo (cTabela 'conta-pagar'). Aqui testamos as
 * chamadas de LEITURA no mesmo endpoint.
 *
 * NAO grava e NAO altera nada no Omie — so lista e consulta.
 */

require('isomorphic-fetch');
const { getCredentials, listarContasPagarPorVencimento } = require('../shared/omie');

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

async function chamar(endpoint, call, param, creds) {
  const resp = await fetch(OMIE_BASE + endpoint, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', 'Accept': 'application/json',
               'User-Agent': 'PronepNF/1.0 (Azure SWA Functions)' },
    body: JSON.stringify({ call: call, app_key: creds.appKey,
                           app_secret: creds.appSecret, param: [param] })
  });
  const texto = await resp.text();
  let data;
  try { data = JSON.parse(texto); }
  catch (e) { return { aceito: false, erro: 'nao-JSON: ' + texto.slice(0, 150) }; }
  if (data && data.faultstring) return { aceito: false, erro: String(data.faultstring).slice(0, 200) };
  if (!resp.ok) return { aceito: false, erro: 'HTTP ' + resp.status };
  return { aceito: true, data: data };
}

module.exports = async function (context, req) {
  const t0 = Date.now();
  const u = String((req.query || {}).unidade || 'RJ').toUpperCase();
  const quantas = Math.max(1, Math.min(10, parseInt((req.query || {}).contas, 10) || 5));
  const diag = { unidade: u, testes: {}, contasTestadas: [], timeMs: 0 };

  try {
    if (!(await isAdmin(req))) {
      context.res = { status: 403, headers: { 'Content-Type': 'application/json' },
        body: { error: 'Apenas admin' } };
      return;
    }
    const creds = getCredentials(u);
    diag.empresa = creds.empresa;

    /* Pega contas EM ABERTO reais para testar contra dados de verdade. */
    const r = await listarContasPagarPorVencimento(
      { filtroExtra: { filtrar_por_status: 'EMABERTO' }, maxPaginas: 1 }, creds);
    const amostra = r.contas.slice(0, quantas);
    if (!amostra.length) throw new Error('Nenhuma conta em aberto para testar');

    /* 1) O ListarAnexo existe? Testa uma vez; a recusa ja diz o nome certo. */
    diag.testes.ListarAnexo = await chamar('/geral/anexo/', 'ListarAnexo',
      { cTabela: 'conta-pagar', nId: Number(amostra[0].codigo_lancamento_omie) }, creds);

    /* 2) Quantas das contas realmente TEM anexo, e de que tipo. Existir a chamada
          nao significa que a Pronep use o recurso — sao perguntas diferentes. */
    for (const c of amostra) {
      const bloco = {
        codigo_lancamento_omie: c.codigo_lancamento_omie,
        nf: c.numero_documento_fiscal || c.numero_documento || '',
        temChaveNFe: !!(c.chave_nfe && String(c.chave_nfe).replace(/\D/g, '').length === 44)
      };
      const rr = await chamar('/geral/anexo/', 'ListarAnexo',
        { cTabela: 'conta-pagar', nId: Number(c.codigo_lancamento_omie) }, creds);
      if (!rr.aceito) { bloco.erro = rr.erro; diag.contasTestadas.push(bloco); continue; }
      const lista = (rr.data && (rr.data.listaAnexos || rr.data.anexos ||
                     rr.data.anexo_cadastro || rr.data.lista_anexo)) || [];
      bloco.qtdAnexos = Array.isArray(lista) ? lista.length : 0;
      bloco.chavesDaResposta = rr.data ? Object.keys(rr.data).slice(0, 8) : [];
      if (Array.isArray(lista) && lista.length) {
        bloco.anexos = lista.slice(0, 4).map(function (a) {
          return { nome: a.cNomeArquivo || a.nome_arquivo || a.cNome || '?',
                   id: a.nIdAnexo || a.id_anexo || null,
                   tipo: a.cTipoArquivo || a.tipo || '' };
        });
      }
      diag.contasTestadas.push(bloco);
    }

    /* 2b) RECEBIMENTO DE NF-e — o modulo que de fato tem o que interessa.
       A tela do Omie "Recebimento NF-e" mostra itens, transporte, totais e
       parcelas, e tem os botoes "Exibir DANFE" e "Exibir XML". Isso NAO vem de
       contas a pagar: e outro modulo, com outra API. E o XML de la que permite
       montar um detalhe igual ao do Omie sem depender de raspagem de tela.
       Os nomes abaixo sao candidatos; o Omie recusa com faultstring o que nao
       existe, entao sondar e barato e conclusivo. */
    const chaveTeste = (amostra.find(function (c) {
      return c.chave_nfe && String(c.chave_nfe).replace(/\D/g, '').length === 44;
    }) || {}).chave_nfe;
    diag.chaveUsadaNoTeste = chaveTeste || '(nenhuma conta da amostra tem chave)';

    const candidatosNFe = [
      ['xml_ListarDocumentos',   '/produtos/xml/',        'ListarDocumentos',
        { pagina: 1, registros_por_pagina: 5 }],
      ['xml_ObterDocumento',     '/produtos/xml/',        'ObterDocumento',
        { nChaveNFe: chaveTeste || '' }],
      ['nfe_ListarNFesRecebidas','/produtos/nfe/',        'ListarNFesRecebidas',
        { pagina: 1, registros_por_pagina: 5 }],
      ['recebimento_Listar',     '/produtos/recebimentonfe/', 'ListarRecebimentos',
        { pagina: 1, registros_por_pagina: 5 }],
      ['dfe_ListarDocumentos',   '/produtos/dfedocsfiscais/', 'ListarDocumentos',
        { pagina: 1, registros_por_pagina: 5 }]
    ];
    diag.recebimentoNFe = {};
    for (const [nome, ep, call, param] of candidatosNFe) {
      if (/Obter/.test(call) && !chaveTeste) {
        diag.recebimentoNFe[nome] = { pulado: 'sem chave na amostra' };
        continue;
      }
      const rr = await chamar(ep, call, param, creds);
      /* Nao despeja o XML inteiro na resposta: so o formato e o tamanho. */
      diag.recebimentoNFe[nome] = rr.aceito
        ? { aceito: true, chaves: Object.keys(rr.data || {}).slice(0, 12),
            total: (rr.data && (rr.data.total_de_registros || rr.data.nTotRegistros)) || null,
            tamanhoResposta: JSON.stringify(rr.data || {}).length }
        : rr;
    }

    /* 3) Se houver anexo, da para BAIXAR? Testa o Obter no primeiro que tiver. */
    const comAnexo = diag.contasTestadas.find(function (b) { return b.anexos && b.anexos.length; });
    if (comAnexo) {
      const idAnexo = comAnexo.anexos[0].id;
      diag.testes.ObterAnexo = await chamar('/geral/anexo/', 'ObterAnexo',
        { cTabela: 'conta-pagar', nId: Number(comAnexo.codigo_lancamento_omie),
          nIdAnexo: Number(idAnexo) }, creds);
      if (diag.testes.ObterAnexo.aceito) {
        const d = diag.testes.ObterAnexo.data || {};
        /* Nao devolve o arquivo inteiro na resposta do diagnostico: so o formato. */
        diag.testes.ObterAnexo = {
          aceito: true,
          chaves: Object.keys(d),
          temUrl: !!(d.cUrl || d.url || d.cLinkDownload),
          temBase64: !!(d.cArquivo || d.arquivo || d.cBase64),
          amostraUrl: String(d.cUrl || d.url || d.cLinkDownload || '').slice(0, 120)
        };
      }
    } else {
      diag.testes.ObterAnexo = { pulado: 'nenhuma conta da amostra tem anexo' };
    }

    diag.timeMs = Date.now() - t0;
    context.res = { status: 200, headers: { 'Content-Type': 'application/json' },
      body: Object.assign({ ok: true,
        leitura: 'ListarAnexo.aceito=false -> a chamada nao existe. aceito=true com ' +
                 'qtdAnexos 0 em todas -> existe mas a Pronep nao usa. Com anexos e ' +
                 'ObterAnexo.temUrl/temBase64 -> da para trazer o DANFE para o card.' }, diag) };
  } catch (err) {
    diag.timeMs = Date.now() - t0;
    context.res = { status: 500, headers: { 'Content-Type': 'application/json' },
      body: Object.assign({ error: (err && err.message) || String(err) }, diag) };
  }
};
