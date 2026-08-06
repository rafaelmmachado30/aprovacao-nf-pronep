/**
 * /api/CriarListaDocumentosFiscais  (GET|POST) — ADMIN.
 *
 * Cria (ou valida) a estrutura do quadro "NFs a Pagar":
 *   - lista PRONEP-NF-DocumentosFiscais  (as NF-e emitidas contra nossos CNPJs)
 *   - lista PRONEP-NF-SefazControle      (ponteiro de NSU, UM POR CNPJ)
 *   - coluna ChaveAcesso na lista de Notas ja existente
 *
 * Mesmo padrao de CriarListaContratos/CriarListaRecorrentes: IDEMPOTENTE, cria o
 * que falta uma coisa por vez e devolve diagnostico por etapa. Rodar de novo e
 * seguro — e e o jeito de conferir se esta tudo no lugar.
 *
 * A coluna ChaveAcesso na lista de Notas e ADITIVA: nenhuma nota existente muda,
 * e o codigo que le a lista continua funcionando sem enxerga-la. Sem essa coluna
 * o merge do quadro cai no caminho fraco (CNPJ + numero + serie).
 */

require('isomorphic-fetch');
const { getGraphClient } = require('../shared/graph');
const {
  LIST_DOCFIS, LIST_SEFAZ, LIST_NOTAS,
  COLUNAS_DOCFIS, COLUNAS_SEFAZ, COLUNAS_NOTAS_EXTRA,
  lerCnpjsConfigurados, garantirPonteiro
} = require('../shared/documentosFiscais');

function readClientPrincipal(req) {
  const header = req.headers && req.headers['x-ms-client-principal'];
  if (!header) return null;
  try { return JSON.parse(Buffer.from(header, 'base64').toString('utf-8')); } catch (e) { return null; }
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

async function garantirLista(client, siteId, nome, colunas, diag) {
  const bloco = { lista: nome, jaExistia: null, criada: false, colunasCriadas: [], falhas: [] };

  let listId = null;
  try {
    const r = await client.api('/sites/' + siteId + '/lists')
      .filter("displayName eq '" + nome + "'").get();
    if (r.value && r.value.length) listId = r.value[0].id;
  } catch (e) {
    /* O filtro do Graph falha em alguns tenants; cai pra listagem completa. */
    const all = await client.api('/sites/' + siteId + '/lists').get();
    const f = (all.value || []).find(function (l) { return l.displayName === nome; });
    if (f) listId = f.id;
  }
  bloco.jaExistia = !!listId;

  if (!listId) {
    const nova = await client.api('/sites/' + siteId + '/lists').post({
      displayName: nome, list: { template: 'genericList' }
    });
    listId = nova.id;
    bloco.criada = true;
  }
  bloco.listId = listId;

  const cols = await client.api('/sites/' + siteId + '/lists/' + listId + '/columns').get();
  const existentes = new Set();
  for (const c of (cols.value || [])) {
    if (c.displayName) existentes.add(c.displayName);
    if (c.name) existentes.add(c.name);
  }

  for (const col of colunas) {
    if (existentes.has(col.name)) continue;
    try {
      await client.api('/sites/' + siteId + '/lists/' + listId + '/columns')
        .post(Object.assign({ name: col.name }, col.def));
      bloco.colunasCriadas.push(col.name);
    } catch (eCol) {
      bloco.falhas.push({ coluna: col.name, error: eCol.message });
    }
  }

  diag.listas.push(bloco);
  return listId;
}

module.exports = async function (context, req) {
  const diag = { step: 'init', listas: [], ponteiros: [], avisos: [], timeMs: 0 };
  const t0 = Date.now();

  try {
    if (!(await isAdmin(req))) {
      context.res = { status: 403, headers: { 'Content-Type': 'application/json' },
        body: { error: 'Apenas admin' } };
      return;
    }

    const client = await getGraphClient();

    diag.step = 'resolve_site';
    const host = process.env.SHAREPOINT_SITE_HOSTNAME;
    const path = process.env.SHAREPOINT_SITE_PATH;
    if (!host || !path) throw new Error('SHAREPOINT_SITE_HOSTNAME/PATH nao configurados');
    const site = await client.api('/sites/' + host + ':' + path).get();
    const siteId = site.id;
    diag.site = { host, path, siteId };

    diag.step = 'lista_documentos';
    await garantirLista(client, siteId, LIST_DOCFIS, COLUNAS_DOCFIS, diag);

    diag.step = 'lista_sefaz';
    await garantirLista(client, siteId, LIST_SEFAZ, COLUNAS_SEFAZ, diag);

    /* A lista de Notas JA EXISTE e esta em producao: aqui so ACRESCENTAMOS a
       coluna. garantirLista nao cria nada porque a lista ja existe, e nenhuma nota
       e tocada. */
    diag.step = 'coluna_chave_nas_notas';
    await garantirLista(client, siteId, LIST_NOTAS, COLUNAS_NOTAS_EXTRA, diag);

    /* Ponteiro de NSU: um por CNPJ configurado. Sem isso a primeira execucao do
       cron nao saberia de onde comecar. */
    diag.step = 'ponteiros_sefaz';
    let cnpjs = [];
    try { cnpjs = lerCnpjsConfigurados(); }
    catch (eCfg) { diag.avisos.push('SEFAZ_CNPJS: ' + eCfg.message); }

    if (!cnpjs.length) {
      diag.avisos.push('Nenhum CNPJ em SEFAZ_CNPJS — a busca automatica nao vai rodar ' +
                       'ate essa App Setting ser preenchida.');
    }
    for (const c of cnpjs) {
      try {
        const p = await garantirPonteiro(client, siteId, c.cnpj, c.apelido);
        diag.ponteiros.push({ cnpj: c.cnpj, apelido: c.apelido, ultimoNSU: p.ultimoNSU });
      } catch (eP) {
        diag.avisos.push('Ponteiro de ' + c.cnpj + ': ' + eP.message);
      }
    }

    /* Avisa sobre certificado faltando ANTES de o cron falhar de madrugada.
       Precisa olhar as DUAS fontes: o .pfx vive no Blob Storage (nao cabe em App
       Setting — teto de 10 KB no total) e a App Setting so existe como fallback.
       Checar so uma delas produzia "certificado ausente" com o certificado
       presente e validado — aviso errado manda consertar o que nao esta quebrado,
       e e pior do que nao avisar. */
    diag.step = 'certificados';
    let noBlob = [];
    let erroBlob = null;
    try { noBlob = await require('../shared/blobCert').listar(); }
    catch (eB) { erroBlob = eB.message; }
    diag.certificados = { noBlob: noBlob, erroBlob: erroBlob };

    for (const c of cnpjs) {
      const emAppSetting = !!process.env['SEFAZ_CERT_' + c.cnpj + '_PFX'];
      const temBlob = noBlob.indexOf(c.cnpj + '.pfx') >= 0;
      if (!emAppSetting && !temBlob) {
        diag.avisos.push('Certificado ausente para ' + c.apelido + ': nao esta no Blob ' +
          '(container "certificados", arquivo ' + c.cnpj + '.pfx)' +
          (erroBlob ? ' — e a leitura do Blob falhou: ' + erroBlob : '') +
          ' nem em App Setting.');
      }
      if (!process.env['SEFAZ_CERT_' + c.cnpj + '_SENHA']) {
        diag.avisos.push('Senha ausente para ' + c.apelido +
          ' (App Setting SEFAZ_CERT_' + c.cnpj + '_SENHA). O certificado nao abre sem ela.');
      }
    }
    /* O caminho SEFAZ esta desativado de proposito (o Omie ja consome o mesmo DFe
       e a cota e por CNPJ). Deixar claro aqui evita que a ausencia de ponteiros
       avancando pareca defeito. */
    diag.avisos.push('Lembrete: a busca automatica na SEFAZ esta com o cron desligado. ' +
      'A fonte do quadro e o Omie (/api/SincronizarOmie).');

    diag.step = 'done';
    diag.timeMs = Date.now() - t0;
    context.res = {
      status: 200, headers: { 'Content-Type': 'application/json' },
      body: Object.assign({ ok: true, mensagem: 'Estrutura de NFs a Pagar conferida.' }, diag)
    };
  } catch (err) {
    diag.timeMs = Date.now() - t0;
    context.res = {
      status: 500, headers: { 'Content-Type': 'application/json' },
      body: Object.assign({ error: (err && err.message) || String(err) }, diag)
    };
  }
};
