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

  // Test call: ?test=SP|RJ|ES — chama ListarContasPagar pra verificar credenciais
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
      const resp = await fetch('https://app.omie.com.br/api/v1/financas/contapagar/', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json; charset=utf-8' },
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
          url: 'https://app.omie.com.br/api/v1/financas/contapagar/',
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
