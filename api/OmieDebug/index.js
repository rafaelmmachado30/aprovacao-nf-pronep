/**
 * /api/OmieDebug — diagnostico da integracao Omie.
 *
 * Anonymous, retorna info sobre:
 *   - Env vars OMIE_APP_KEY/SECRET por unidade
 *   - Module load do shared/omie.js
 *   - Test call (opcional via ?test=SP|RJ|ES)
 *   - Test ListarClientes (opcional via ?test_cli=SP|RJ|ES&cnpj=...)
 */

module.exports = async function (context, req) {
  const out = {
    timestamp: new Date().toISOString(),
    env: {
      hasOmieKeyES: !!process.env.OMIE_APP_KEY_ES,
      hasOmieSecretES: !!process.env.OMIE_APP_SECRET_ES,
      hasOmieKeyRJ: !!process.env.OMIE_APP_KEY_RJ,
      hasOmieSecretRJ: !!process.env.OMIE_APP_SECRET_RJ,
      hasOmieKeySP: !!process.env.OMIE_APP_KEY_SP,
      hasOmieSecretSP: !!process.env.OMIE_APP_SECRET_SP,
      omieKeyESPrefix: (process.env.OMIE_APP_KEY_ES || '').slice(0, 4) + '...',
      omieKeyRJPrefix: (process.env.OMIE_APP_KEY_RJ || '').slice(0, 4) + '...',
      omieKeySPPrefix: (process.env.OMIE_APP_KEY_SP || '').slice(0, 4) + '...'
    },
    requires: {},
    testCall: null,
    testListarClientes: null
  };

  try {
    const omie = require('../shared/omie');
    out.requires.omie = { ok: true, funcs: Object.keys(omie) };
  } catch (e) {
    out.requires.omie = { ok: false, error: e.message, stack: (e.stack || '').split('\n').slice(0, 5) };
  }
  try { require('isomorphic-fetch'); out.requires.fetch = { ok: true }; }
  catch (e) { out.requires.fetch = { ok: false, error: e.message }; }
  try { require('@azure/identity'); out.requires.azureIdentity = { ok: true }; }
  catch (e) { out.requires.azureIdentity = { ok: false, error: e.message }; }
  try { require('@microsoft/microsoft-graph-client'); out.requires.graph = { ok: true }; }
  catch (e) { out.requires.graph = { ok: false, error: e.message }; }

  // Test ListarContasPagar: ?test=SP|RJ|ES
  const testParam = (req.query && req.query.test) || '';
  if (['SP','RJ','ES'].includes(testParam)) {
    try {
      const { getCredentials } = require('../shared/omie');
      const creds = getCredentials(testParam);
      const fetch = require('isomorphic-fetch');
      const body = {
        call: 'ListarContasPagar',
        app_key: creds.appKey,
        app_secret: creds.appSecret,
        param: [{ pagina: 1, registros_por_pagina: 1, apenas_importado_api: 'N' }]
      };
      const url = 'https://app.omie.com.br/api/v1/financas/contapagar/';
      const headers = {
        'Content-Type': 'application/json',
        'Accept': 'application/json',
        'User-Agent': 'PronepNF/1.0 (Azure SWA Functions)'
      };
      const resp = await fetch(url, { method: 'POST', headers: headers, body: JSON.stringify(body) });
      const text = await resp.text();
      let data;
      try { data = JSON.parse(text); } catch (e) { data = { rawText: text.slice(0, 500) }; }
      out.testCall = {
        ok: resp.ok && !data.faultstring,
        status: resp.status,
        empresa: creds.empresa,
        totalContas: (data && data.total_de_registros) || 0,
        faultstring: data && data.faultstring,
        faultcode: data && data.faultcode,
        rawBody: text.slice(0, 1200)
      };
    } catch (e) {
      out.testCall = { ok: false, error: e.message, stack: (e.stack || '').split('\n').slice(0, 8) };
    }
  }

  // Test ListarClientes: ?test_cli=SP|RJ|ES&cnpj=...
  const testCliParam = (req.query && req.query.test_cli) || '';
  const cnpjQuery = (req.query && req.query.cnpj) || '23484104000114';
  if (['SP','RJ','ES'].includes(testCliParam)) {
    try {
      const { getCredentials, buscarCliente, normalizaDoc } = require('../shared/omie');
      const creds = getCredentials(testCliParam);
      const fetch = require('isomorphic-fetch');
      const cnpjLimpo = normalizaDoc(cnpjQuery);

      // Raw call pra ver o que Omie devolve
      const body = {
        call: 'ListarClientes',
        app_key: creds.appKey,
        app_secret: creds.appSecret,
        param: [{
          pagina: 1,
          registros_por_pagina: 5,
          apenas_importado_api: 'N',
          clientesFiltro: { cnpj_cpf: cnpjLimpo }
        }]
      };
      const resp = await fetch('https://app.omie.com.br/api/v1/geral/clientes/', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', 'Accept': 'application/json', 'User-Agent': 'PronepNF/1.0' },
        body: JSON.stringify(body)
      });
      const text = await resp.text();
      let data;
      try { data = JSON.parse(text); } catch (e) { data = { rawText: text.slice(0, 600) }; }

      out.testListarClientes = {
        status: resp.status,
        ok: resp.ok && !data.faultstring,
        empresa: creds.empresa,
        cnpjConsultado: cnpjLimpo,
        faultstring: data && data.faultstring,
        faultcode: data && data.faultcode,
        total_de_registros: data && data.total_de_registros,
        lista_resumido_len: (data && data.clientes_cadastro_resumido) ? data.clientes_cadastro_resumido.length : null,
        lista_cadastro_len: (data && data.clientes_cadastro) ? data.clientes_cadastro.length : null,
        primeiroResultado: (data && (data.clientes_cadastro_resumido || data.clientes_cadastro) || [])[0] || null,
        rawBody: text.slice(0, 1500)
      };

      // E tambem testa via buscarCliente do shared
      try {
        const result = await buscarCliente(cnpjLimpo, creds);
        out.testBuscarClienteShared = result;
      } catch (e) {
        out.testBuscarClienteShared = { error: e.message, stack: (e.stack || '').split('\n').slice(0, 5) };
      }
    } catch (e) {
      out.testListarClientes = { ok: false, error: e.message, stack: (e.stack || '').split('\n').slice(0, 8) };
    }
  }


  // Test ListarContasPagar filtrado: ?test_conta=SP|RJ|ES&cli=10503515421&num=6282
  const testContaParam = (req.query && req.query.test_conta) || '';
  const cliParam = (req.query && req.query.cli) || '';
  const numParam = (req.query && req.query.num) || '';
  if (['SP','RJ','ES'].includes(testContaParam)) {
    try {
      const { getCredentials, normalizaNumeroNF } = require('../shared/omie');
      const creds = getCredentials(testContaParam);
      const fetch = require('isomorphic-fetch');
      const numAlvo = normalizaNumeroNF(numParam);

      const body = {
        call: 'ListarContasPagar',
        app_key: creds.appKey,
        app_secret: creds.appSecret,
        param: [{
          pagina: 1,
          registros_por_pagina: 50,
          apenas_importado_api: 'N',
          clientesFiltro: { codigo_cliente_omie: Number(cliParam) }
        }]
      };
      const resp = await fetch('https://app.omie.com.br/api/v1/financas/contapagar/', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', 'Accept': 'application/json', 'User-Agent': 'PronepNF/1.0' },
        body: JSON.stringify(body)
      });
      const text = await resp.text();
      let data;
      try { data = JSON.parse(text); } catch (e) { data = { rawText: text.slice(0, 600) }; }

      const items = (data && (data.conta_pagar_cadastro || data.contas_pagar_cadastro)) || [];
      const matchNum = items.filter(function (it) {
        const a = normalizaNumeroNF(it.numero_documento || '');
        const b = normalizaNumeroNF(it.numero_documento_fiscal || it.nota_fiscal || '');
        return numAlvo && (a === numAlvo || b === numAlvo);
      });

      out.testListarContasPagarFiltrado = {
        status: resp.status,
        ok: resp.ok && !data.faultstring,
        empresa: creds.empresa,
        codigo_cliente_omie: Number(cliParam),
        numAlvo: numAlvo,
        faultstring: data && data.faultstring,
        faultcode: data && data.faultcode,
        total_de_registros: data && data.total_de_registros,
        total_de_paginas: data && data.total_de_paginas,
        itemsNaPaginaUm: items.length,
        matchPorNumero: matchNum.map(function (it) {
          return {
            codigo_lancamento_omie: it.codigo_lancamento_omie,
            numero_documento: it.numero_documento,
            nota_fiscal: it.numero_documento_fiscal || it.nota_fiscal,
            valor: it.valor_documento,
            status: it.status_titulo
          };
        }),
        primeiros3: items.slice(0, 3).map(function (it) {
          return {
            codigo_lancamento_omie: it.codigo_lancamento_omie,
            numero_documento: it.numero_documento,
            nota_fiscal: it.numero_documento_fiscal || it.nota_fiscal,
            valor: it.valor_documento,
            cnpj: it.cnpj_cpf_fornecedor || it.cnpj_cpf
          };
        }),
        rawBody: text.slice(0, 800)
      };
    } catch (e) {
      out.testListarContasPagarFiltrado = { ok: false, error: e.message, stack: (e.stack || '').split('\n').slice(0, 8) };
    }
  }

  context.res = {
    status: 200,
    headers: { 'Content-Type': 'application/json' },
    body: out
  };
};
