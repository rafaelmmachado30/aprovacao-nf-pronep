/**
 * /api/CaixaEntrada  (GET = listar | POST = upload/combinar/excluir)
 *
 * Caixa de Entrada de PDFs avulsos (nota + boleto chegam separados). Os arquivos
 * ficam guardados numa pasta POR USUARIO no SharePoint ate o lancamento:
 *   "Notas Fiscais/Caixa de Entrada/{slug-do-email}/"
 *
 * Acoes:
 *   GET                          -> lista os PDFs do usuario logado
 *   POST { action:'upload', fileBase64, fileName }   -> guarda 1 PDF
 *   POST { action:'combinar', ids:[...] }            -> une os PDFs (pdf-lib) e
 *                                                       devolve o combinado em base64
 *   POST { action:'excluir', id|ids }                -> remove PDF(s) do usuario
 *
 * Visibilidade: cada usuario so enxerga/mexe na PROPRIA pasta. Sem acesso cruzado.
 */

require('isomorphic-fetch');
const { getUser } = require('../shared/auth');
const { ClientSecretCredential } = require('@azure/identity');
const { Client, ResponseType } = require('@microsoft/microsoft-graph-client');
const { TokenCredentialAuthenticationProvider } = require('@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials');

const BASE_FOLDER = 'Notas Fiscais/Caixa de Entrada';
const MAX_UPLOAD = 8 * 1024 * 1024;      // 8MB por arquivo
const MAX_COMBINADO = 12 * 1024 * 1024;  // 12MB combinado
const _cache = { siteId: null };

function getGraphClient() {
  const credential = new ClientSecretCredential(
    process.env.AAD_TENANT_ID, process.env.AAD_CLIENT_ID, process.env.AAD_CLIENT_SECRET
  );
  const authProvider = new TokenCredentialAuthenticationProvider(credential, {
    scopes: ['https://graph.microsoft.com/.default']
  });
  return Client.initWithMiddleware({ authProvider });
}

async function resolveSiteId(client) {
  if (_cache.siteId) return _cache.siteId;
  const host = process.env.SHAREPOINT_SITE_HOSTNAME;
  const path = process.env.SHAREPOINT_SITE_PATH;
  if (!host || !path) throw new Error('SHAREPOINT_* incompletas');
  const siteResp = await client.api('/sites/' + host + ':' + path).get();
  _cache.siteId = siteResp.id;
  return _cache.siteId;
}

function slugEmail(email) {
  return String(email || 'anon').toLowerCase().replace(/[^a-z0-9]+/g, '_').replace(/^_|_$/g, '') || 'anon';
}
function pastaDoUsuario(email) {
  return BASE_FOLDER + '/' + slugEmail(email);
}
function sanitizeName(name) {
  let n = String(name || 'arquivo.pdf').replace(/[\\/:*?"<>|]+/g, '_').trim();
  if (!/\.pdf$/i.test(n)) n += '.pdf';
  return n.slice(0, 120);
}

async function listarArquivos(client, siteId, email) {
  const folder = pastaDoUsuario(email);
  try {
    const resp = await client.api('/sites/' + siteId + '/drive/root:/' + folder + ':/children')
      .select('id,name,size,createdDateTime,file,folder,@microsoft.graph.downloadUrl').get();
    return (resp.value || []).filter(function (x) { return x.file && !x.folder; }).map(function (x) {
      // O estado é codificado NO NOME do arquivo (description nao persiste de
      // forma confiavel em bibliotecas SharePoint). Marcadores:
      //   ...__combinada__<n>NFs.pdf                       -> combinado (pronto p/ lancar)
      //   ...__combinada__<n>NFs__LANCADA__<data>__<id>.pdf -> lancada
      var nome = x.name || '';
      var estado = 'avulso', lancadaEm = null, notaId = null;
      var mL = nome.match(/__LANCADA__([^_]+)__([^.]*)\.pdf$/i);
      if (mL) {
        estado = 'lancada';
        lancadaEm = mL[1] || null;
        notaId = (mL[2] && mL[2] !== 'NA') ? mL[2] : null;
      } else if (/__combinada__\d+NFs/i.test(nome)) {
        estado = 'combinado';
      }
      return {
        id: x.id, nome: x.name, tamanho: x.size || 0,
        criadoEm: x.createdDateTime || null,
        estado: estado, lancadaEm: lancadaEm, notaId: notaId,
        urlDownload: x['@microsoft.graph.downloadUrl'] || null
      };
    }).sort(function (a, b) {
      // Lançadas sempre no fim; dentro de cada grupo, mais recentes primeiro
      var la = a.estado === 'lancada' ? 1 : 0, lb = b.estado === 'lancada' ? 1 : 0;
      if (la !== lb) return la - lb;
      return String(b.criadoEm).localeCompare(String(a.criadoEm));
    });
  } catch (e) {
    // Pasta ainda nao existe -> caixa vazia
    if (e.statusCode === 404) return [];
    throw e;
  }
}

// Baixa os bytes de um PDF da caixa. Tenta o downloadUrl pre-autenticado e, no
// fallback, usa o Graph /content como ARRAYBUFFER. Importante: nunca usar
// Buffer.from() sobre o stream cru do .get() (no SDK/Node atual o corpo vem como
// ReadableStream e o Buffer.from quebra com "first argument must be of type...").
async function baixarPdfBuffer(client, siteId, item) {
  if (!item || !item.id) return null;
  if (item.urlDownload) {
    try {
      const r = await fetch(item.urlDownload);
      if (r.ok) {
        const ab = await r.arrayBuffer();
        if (ab && ab.byteLength) return Buffer.from(ab);
      }
    } catch (e) { /* cai pro fallback do Graph */ }
  }
  const ab = await client.api('/sites/' + siteId + '/drive/items/' + item.id + '/content')
    .responseType(ResponseType.ARRAYBUFFER).get();
  return Buffer.from(ab);
}

module.exports = async function (context, req) {
  try {
    const user = await getUser(req);
    if (!user || !user.email) { context.res = { status: 401, body: { error: 'Nao autenticado' } }; return; }
    const email = user.email.toLowerCase();
    const client = getGraphClient();
    const siteId = await resolveSiteId(client);
    const folder = pastaDoUsuario(email);

    // ===== LISTAR =====
    if (req.method === 'GET') {
      const arquivos = await listarArquivos(client, siteId, email);
      context.res = { status: 200, headers: { 'Content-Type': 'application/json', 'Cache-Control': 'no-store' }, body: { ok: true, arquivos: arquivos } };
      return;
    }

    const body = req.body || {};
    const action = body.action || '';

    // ===== UPLOAD =====
    if (action === 'upload') {
      if (!body.fileBase64) { context.res = { status: 400, body: { error: 'fileBase64 obrigatorio' } }; return; }
      const b64 = String(body.fileBase64).includes(',') ? String(body.fileBase64).split(',')[1] : String(body.fileBase64);
      const buf = Buffer.from(b64, 'base64');
      if (!buf.length) { context.res = { status: 400, body: { error: 'PDF vazio' } }; return; }
      if (buf.length > MAX_UPLOAD) { context.res = { status: 400, body: { error: 'PDF excede 8MB' } }; return; }
      // valida cabecalho PDF
      if (buf.slice(0, 5).toString('latin1') !== '%PDF-') { context.res = { status: 400, body: { error: 'Arquivo nao parece um PDF valido' } }; return; }
      const stamp = new Date().toISOString().replace(/[:.]/g, '-');
      const nome = stamp + '__' + sanitizeName(body.fileName);
      const uploadPath = '/sites/' + siteId + '/drive/root:/' + encodeURIComponent(folder).replace(/%2F/g, '/') + '/' + encodeURIComponent(nome) + ':/content';
      const up = await client.api(uploadPath).header('Content-Type', 'application/pdf').put(buf);
      context.res = { status: 200, headers: { 'Content-Type': 'application/json' }, body: { ok: true, id: up.id, nome: up.name } };
      return;
    }

    // ===== COMBINAR =====
    if (action === 'combinar') {
      const ids = Array.isArray(body.ids) ? body.ids : [];
      if (ids.length < 1) { context.res = { status: 400, body: { error: 'Selecione ao menos 1 PDF' } }; return; }
      // valida que os ids pertencem a pasta do usuario (sem acesso cruzado)
      const meus = await listarArquivos(client, siteId, email);
      const meusIds = {}; meus.forEach(function (a) { meusIds[a.id] = a; });
      for (const id of ids) { if (!meusIds[id]) { context.res = { status: 403, body: { error: 'Arquivo fora da sua caixa' } }; return; } }
      // So combina avulsos — combinados/lancados nao podem ser recombinados
      for (const id of ids) { if (meusIds[id].estado && meusIds[id].estado !== 'avulso') { context.res = { status: 400, body: { error: 'Selecione apenas arquivos avulsos para combinar' } }; return; } }

      const { PDFDocument } = require('pdf-lib');
      const merged = await PDFDocument.create();
      let totalBytes = 0;
      for (const id of ids) {
        const buf = await baixarPdfBuffer(client, siteId, meusIds[id]);
        if (!buf || !buf.length) { context.res = { status: 502, body: { error: 'Falha ao baixar PDF da caixa' } }; return; }
        totalBytes += buf.length;
        if (totalBytes > MAX_COMBINADO) { context.res = { status: 400, body: { error: 'PDF combinado excede 12MB' } }; return; }
        const src = await PDFDocument.load(buf, { ignoreEncryption: true });
        const paginas = await merged.copyPages(src, src.getPageIndices());
        paginas.forEach(function (p) { merged.addPage(p); });
      }
      const out = await merged.save();
      const outBuf = Buffer.from(out);

      // Grava o PDF unificado na propria pasta da caixa (vira o item do historico).
      const stamp = new Date().toISOString().replace(/[:.]/g, '-');
      const nomeCombinado = stamp + '__combinada__' + ids.length + 'NFs.pdf';
      const upPath = '/sites/' + siteId + '/drive/root:/' + encodeURIComponent(folder).replace(/%2F/g, '/') + '/' + encodeURIComponent(nomeCombinado) + ':/content';
      // O proprio nome (__combinada__<n>NFs.pdf) ja marca o item como "combinado".
      const novo = await client.api(upPath).header('Content-Type', 'application/pdf').put(outBuf);

      // So depois de salvar o unificado com sucesso, exclui os avulsos de origem.
      const apagados = [];
      for (const id of ids) {
        try { await client.api('/sites/' + siteId + '/drive/items/' + id).delete(); apagados.push(id); } catch (e) { /* ignora */ }
      }

      context.res = {
        status: 200, headers: { 'Content-Type': 'application/json' },
        body: { ok: true, id: novo.id, nome: novo.name, tamanho: outBuf.length, qtdArquivos: ids.length, apagados: apagados }
      };
      return;
    }

    // ===== BAIXAR (base64) — usado pelo "Realizar lancamento" =====
    if (action === 'baixar') {
      const id = body.id;
      if (!id) { context.res = { status: 400, body: { error: 'id obrigatorio' } }; return; }
      const meus = await listarArquivos(client, siteId, email);
      const item = meus.find(function (a) { return a.id === id; });
      if (!item) { context.res = { status: 403, body: { error: 'Arquivo fora da sua caixa' } }; return; }
      const buf = await baixarPdfBuffer(client, siteId, item);
      if (!buf || !buf.length) { context.res = { status: 502, body: { error: 'Falha ao baixar PDF' } }; return; }
      context.res = {
        status: 200, headers: { 'Content-Type': 'application/json' },
        body: { ok: true, fileName: item.nome, fileBase64: 'data:application/pdf;base64,' + buf.toString('base64'), tamanho: buf.length }
      };
      return;
    }

    // ===== MARCAR LANCADA — chamado apos o PostNota concluir com sucesso =====
    // Renomeia o arquivo adicionando o marcador __LANCADA__<data>__<notaId> (o
    // nome e a fonte da verdade do estado — ver listarArquivos).
    if (action === 'marcarLancada') {
      const id = body.id;
      if (!id) { context.res = { status: 400, body: { error: 'id obrigatorio' } }; return; }
      const meus = await listarArquivos(client, siteId, email);
      const item = meus.find(function (a) { return a.id === id; });
      if (!item) { context.res = { status: 403, body: { error: 'Arquivo fora da sua caixa' } }; return; }
      if (item.estado !== 'lancada') {
        const base = String(item.nome).replace(/\.pdf$/i, '');
        const dataStr = new Date().toISOString().slice(0, 10);              // YYYY-MM-DD (sem '_')
        const nid = (String(body.notaId || '').replace(/[^A-Za-z0-9]/g, '') || 'NA').slice(0, 40);
        const novoNome = base + '__LANCADA__' + dataStr + '__' + nid + '.pdf';
        await client.api('/sites/' + siteId + '/drive/items/' + id).update({ name: novoNome });
      }
      context.res = { status: 200, headers: { 'Content-Type': 'application/json' }, body: { ok: true } };
      return;
    }

    // ===== EXCLUIR =====
    if (action === 'excluir') {
      const ids = Array.isArray(body.ids) ? body.ids : (body.id ? [body.id] : []);
      if (!ids.length) { context.res = { status: 400, body: { error: 'id ou ids obrigatorio' } }; return; }
      const meus = await listarArquivos(client, siteId, email);
      const meusIds = {}; meus.forEach(function (a) { meusIds[a.id] = true; });
      const removidos = [];
      for (const id of ids) {
        if (!meusIds[id]) continue; // ignora o que nao for do usuario
        try { await client.api('/sites/' + siteId + '/drive/items/' + id).delete(); removidos.push(id); } catch (e) { /* ignora */ }
      }
      context.res = { status: 200, headers: { 'Content-Type': 'application/json' }, body: { ok: true, removidos: removidos } };
      return;
    }

    context.res = { status: 400, body: { error: 'action invalida (use upload|combinar|baixar|marcarLancada|excluir)' } };
  } catch (err) {
    context.log && context.log.error && context.log.error('CaixaEntrada error:', err);
    context.res = { status: 500, headers: { 'Content-Type': 'application/json' }, body: { error: (err && err.message) || String(err) } };
  }
};
