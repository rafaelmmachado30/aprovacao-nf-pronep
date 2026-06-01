/**
 * /api/OmieDebug — diagnostico da integracao Omie.
 *
 * Anonymous, retorna info sobre:
 *   - Env vars OMIE_APP_KEY/SECRET por unidade
 *   - Module load do shared/omie.js
 *   - Test call (opcional via ?test=SP|RJ|ES)
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
    testCall: null
  };

  try {
    const omie = require('../shared/omie');
    out.requires.omie = {
      ok: true,
      funcs: Object.keys(omie)
    };
  } catch (e) {
    out.requires.omie = { ok: false, error: e.message, stack: (e.stack || '').split('\n').slice(0, 5) };
  }

  try {
    require('isomorphic-fetch');
    out.requires.fetch = { ok: true };
  } catch (e) {
    out.requires.fetch = { ok: false, error: e.message };
  }

  try {
    require('@azure/identity');
    out.requires.azureIdentity = { ok: true };
  } catch (e) {
    out.requires.azureIdentity = { ok: false, error: e.message };
  }

  try {
    require('@microsoft/microsoft-graph-client');
    out.requires.graph = { ok: true };
  } catch (e) {
    out.requires.graph = { ok: false, error: e.message };
  }

  // Test call: ?test=SP|RJ|ES — chama ListarContasPagar com VARIAS combinacoes pra cravar.
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

      // Variavel: testUrl + testHeaders permitem testar combos via query strings
      const urlMode = (req.query && req.query.url) || 'slash';  // slash | noslash | absolute
      const headerMode = (req.query && req.query.hdr) || 'default'; // default | minimal | curl
      let url;
      if (urlMode === 'noslash') url = 'https://app.omie.com.br/api/v1/financas/contapagar';
      else if (urlMode === 'absolute') url = 'https://app.omie.com.br/api/v1/financas/contapagar/?JSON';
      else url = 'https://app.omie.com.br/api/v1/financas/contapagar/';

      let headers;
      if (headerMode === 'minimal') {
        headers = { 'Content-Type': 'application/json' };
      } else if (headerMode === 'curl') {
        headers = { 'Content-Type': 'application/json', 'Accept': '*/*', 'User-Agent': 'curl/7.68.0' };
      } else {
        headers = {
          'Content-Type': 'application/json',
          'Accept': 'application/json',
          'User-Agent': 'PronepNF/1.0 (Azure SWA Functions)'
        };
      }

      const resp = await fetch(url, {
        method: 'POST',
        headers: headers,
        body: JSON.stringify(body)
      });
      const text = await resp.text();
      let data;
      try { data = JSON.parse(text); } catch (e) { data = { rawText: text.slice(0, 500) }; }
      out.testCall = {
        ok: resp.ok && !data.faultstring,
        status: resp.status,
        empresa: creds.empresa,
        totalContas: (data && data.total_de_registros) || (data && data.conta_pagar_cadastro && data.conta_pagar_cadastro.length) || 0,
        faultstring: data && data.faultstring,
        faultcode: data && data.faultcode,
        // Tudo abaixo eh debug bruto
        responseHeaders: {
          contentType: resp.headers && resp.headers.get && resp.headers.get('content-type'),
          server: resp.headers && resp.headers.get && resp.headers.get('server')
        },
        rawBody: text.slice(0, 1500),  // body COMPLETO (truncado em 1500 chars)
        requestSent: {
          url: url,
          urlMode: urlMode,
          headerMode: headerMode,
          headers: headers,
          body: { call: body.call, app_key: (body.app_key||'').slice(0,4)+'...', app_secret: '***hidden***', param: body.param }
        }
      };
    } catch (e) {
      out.testCall = { ok: false, error: e.message, stack: (e.stack || '').split('\n').slice(0, 8) };
    }
  }

  context.res = {
    status: 200,
    headers: { 'Content-Type': 'application/json' },
    body: out
  };
};
