/**
 * /api/ReatribuirPendentes (POST) — ADMIN ou TI.
 *
 * Reatribui o AprovadorAtual das NFs PENDENTES de uma Unidade x Diretoria, do aprovador
 * antigo para o novo. Usado ao trocar o aprovador em Configuração Aprovadores, para levar
 * junto as NFs que ja estavam na fila do aprovador anterior.
 *
 * Body: { unidade, diretoria, deEmail?, paraEmail, dryRun? }
 *   - dryRun=true: só conta/lista (não altera).
 *   - deEmail: se informado, só reatribui NFs cujo AprovadorAtual == deEmail (mais seguro).
 *
 * Match: Status ∈ (Lancada, EmAprovacao, Pendente) AND Unidade==unidade AND Diretoria==diretoria
 *        [AND AprovadorAtual==deEmail, se deEmail].
 */

require('isomorphic-fetch');
const { resolveAuthz } = require('../shared/authz');
const { ClientSecretCredential } = require('@azure/identity');
const { Client } = require('@microsoft/microsoft-graph-client');
const { TokenCredentialAuthenticationProvider } =
  require('@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials');

const LIST_NOTAS = 'PRONEP-NF-NotasFiscais';
const PENDENTES = ['Lancada', 'EmAprovacao', 'Pendente'];
const cache = { siteId: null, listId: null, disp2int: null, int2disp: null };

function getGraphClient() {
  const credential = new ClientSecretCredential(
    process.env.AAD_TENANT_ID, process.env.AAD_CLIENT_ID, process.env.AAD_CLIENT_SECRET
  );
  const authProvider = new TokenCredentialAuthenticationProvider(credential, {
    scopes: ['https://graph.microsoft.com/.default']
  });
  return Client.initWithMiddleware({ authProvider });
}

async function resolve(client) {
  if (cache.siteId && cache.listId && cache.disp2int) return cache;
  const host = process.env.SHAREPOINT_SITE_HOSTNAME;
  const path = process.env.SHAREPOINT_SITE_PATH;
  const siteResp = await client.api('/sites/' + host + ':' + path).get();
  cache.siteId = siteResp.id;
  const listsResp = await client.api('/sites/' + cache.siteId + '/lists').filter("displayName eq '" + LIST_NOTAS + "'").get();
  if (!listsResp.value || !listsResp.value.length) throw new Error('Lista de NFs nao encontrada');
  cache.listId = listsResp.value[0].id;
  const cols = await client.api('/sites/' + cache.siteId + '/lists/' + cache.listId + '/columns').get();
  cache.disp2int = {}; cache.int2disp = {};
  for (const c of (cols.value || [])) {
    if (c.displayName && c.name) { cache.disp2int[c.displayName] = c.name; cache.int2disp[c.name] = c.displayName; }
  }
  return cache;
}

module.exports = async function (context, req) {
  try {
    const authz = await resolveAuthz(req);
    if (!authz) { context.res = { status: 401, headers: { 'Content-Type': 'application/json' }, body: { error: 'Nao autenticado' } }; return; }
    if (!(authz.isAdmin || (authz.roles || []).indexOf('ti') >= 0)) {
      context.res = { status: 403, headers: { 'Content-Type': 'application/json' }, body: { error: 'Acesso restrito a Admin ou TI' } }; return;
    }

    const body = req.body || {};
    const unidade = String(body.unidade || '').trim();
    const diretoria = String(body.diretoria || '').trim();
    const deEmail = String(body.deEmail || '').trim().toLowerCase();
    const paraEmail = String(body.paraEmail || '').trim().toLowerCase();
    const dryRun = !!body.dryRun;

    if (!unidade || !diretoria) { context.res = { status: 400, headers: { 'Content-Type': 'application/json' }, body: { error: 'unidade e diretoria obrigatorios' } }; return; }
    if (!dryRun && (!paraEmail || !/@pronep\.com\.br$/i.test(paraEmail))) {
      context.res = { status: 400, headers: { 'Content-Type': 'application/json' }, body: { error: 'paraEmail @pronep.com.br obrigatorio' } }; return;
    }

    const client = getGraphClient();
    const { siteId, listId, disp2int, int2disp } = await resolve(client);
    const iUn = disp2int['Unidade'] || 'Unidade';
    const iDir = disp2int['Diretoria'] || 'Diretoria';
    const iStatus = disp2int['Status'] || 'Status';
    const iAprov = disp2int['AprovadorAtual'] || 'AprovadorAtual';
    const iNum = disp2int['NumeroNF'] || 'NumeroNF';

    // Pagina todas as NFs.
    const itens = [];
    let url = '/sites/' + siteId + '/lists/' + listId + '/items?expand=fields&$top=999';
    let pages = 0;
    while (url && pages < 40) {
      const resp = await client.api(url).get();
      itens.push(...(resp.value || []));
      pages++;
      url = resp['@odata.nextLink'] ? resp['@odata.nextLink'].replace('https://graph.microsoft.com/v1.0', '') : null;
    }

    const alvo = itens.filter(function (it) {
      const f = it.fields || {};
      if (PENDENTES.indexOf(String(f[iStatus] || '')) < 0) return false;
      if (String(f[iUn] || '') !== unidade) return false;
      if (String(f[iDir] || '') !== diretoria) return false;
      if (deEmail && String(f[iAprov] || '').toLowerCase() !== deEmail) return false;
      return true;
    });

    if (dryRun) {
      context.res = { status: 200, headers: { 'Content-Type': 'application/json', 'Cache-Control': 'no-store' },
        body: { ok: true, dryRun: true, total: alvo.length,
          amostra: alvo.slice(0, 5).map(function (it) { return { numero: (it.fields || {})[iNum] || '?', aprovadorAtual: (it.fields || {})[iAprov] || '' }; }) } };
      return;
    }

    let atualizados = 0;
    const erros = [];
    for (const it of alvo) {
      try {
        const patch = {}; patch[iAprov] = paraEmail;
        await client.api('/sites/' + siteId + '/lists/' + listId + '/items/' + it.id + '/fields').patch(patch);
        atualizados++;
      } catch (e) { erros.push({ id: it.id, erro: (e && e.message) || String(e) }); }
    }

    context.res = { status: 200, headers: { 'Content-Type': 'application/json', 'Cache-Control': 'no-store' },
      body: { ok: true, total: alvo.length, atualizados: atualizados, erros: erros, para: paraEmail, por: authz.email } };
  } catch (err) {
    context.log && context.log.error && context.log.error('ReatribuirPendentes:', err);
    context.res = { status: 500, headers: { 'Content-Type': 'application/json' }, body: { error: (err && err.message) || String(err) } };
  }
};
