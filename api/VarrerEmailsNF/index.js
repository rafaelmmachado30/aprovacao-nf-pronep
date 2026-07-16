/**
 * /api/VarrerEmailsNF  (GET|POST) — ADMIN/secret. Automacao de ingestao de NF por e-mail.
 *
 * FASE 1 (esta): varre a INBOX de UM gestor via Graph (Mail.Read), identifica e-mails
 * que parecem NF (assunto/remetente) com anexo PDF, e baixa o(s) PDF(s) para o
 * SharePoint em:  Novas NFs - Automacao/{Unidade}/Diretoria {Diretoria}/
 * Marca cada e-mail processado num "ledger" (JSON no proprio drive) para nao repetir.
 * NAO notifica ainda (WhatsApp/SAN vem na Fase 2).
 *
 * Requer permissao Graph Mail.Read (Application), ESCOPADA por Application Access Policy
 * apenas as caixas dos gestores. So le e-mails que batem no criterio de NF.
 *
 * Query:
 *   ?gestor=email@pronep.com.br   (OBRIGATORIO — caixa a varrer)
 *   ?unidade=SP&diretoria=Tecnologia  (destino; se omitido, tenta derivar da lista Diretorias)
 *   ?dias=3        (janela de recebimento; default 3)
 *   ?limite=50     (max e-mails avaliados; default 50)
 *   ?dryRun=1      (so classifica e reporta; NAO baixa nem grava ledger)
 *
 * Auth: X-Automacao-Secret == App Setting AUTOMACAO_EMAILS_SECRET (cron) OU sessao admin.
 */

require('isomorphic-fetch');
const { requireAdmin } = require('../shared/authz');
const { getGraphClient, resolveSiteId } = require('../shared/graph');

const LIST_FORN = 'PRONEP-NF-Fornecedores';
const LIST_DIR = 'PRONEP-NF-Diretorias';
const PASTA_RAIZ = 'Novas NFs - Automacao';
const LEDGER_PATH = '_automacao/emails_ledger.json';
// Palavras no assunto que indicam NF/fatura (sinal principal de classificacao).
const ASSUNTO_NF = /nota\s*fiscal|\bnfs?e?\b|\bnf\b|fatura|boleto|danfe|cobran[cç]a/i;

let _driveId = null;
async function resolveDriveId(client, siteId) {
  if (_driveId) return _driveId;
  const d = await client.api('/sites/' + siteId + '/drive').get();
  _driveId = d.id;
  return _driveId;
}

// Sinais de fornecedor (e-mails e dominios cadastrados) — reforca a classificacao.
// Best-effort: se a leitura falhar, seguimos so com o assunto.
async function fornecedorSignals(client, siteId) {
  const emails = new Set(), dominios = new Set();
  try {
    const fl = await client.api('/sites/' + siteId + '/lists').filter("displayName eq '" + LIST_FORN + "'").get();
    if (!fl.value || !fl.value.length) return { emails, dominios };
    const flId = fl.value[0].id;
    let url = '/sites/' + siteId + '/lists/' + flId + '/items?expand=fields&$top=500';
    let pages = 0;
    while (url && pages < 30) {
      const r = await client.api(url).get();
      for (const it of (r.value || [])) {
        const f = it.fields || {};
        // Campos candidatos (schema pode variar): Email, email, field_9.
        const cand = [f.Email, f.email, f.field_9].filter(Boolean);
        for (const e of cand) {
          const em = String(e).trim().toLowerCase();
          if (em.indexOf('@') > 0) { emails.add(em); dominios.add(em.split('@')[1]); }
        }
      }
      pages++;
      url = r['@odata.nextLink'] ? r['@odata.nextLink'].replace('https://graph.microsoft.com/v1.0', '') : null;
    }
  } catch (e) { /* best-effort */ }
  return { emails, dominios };
}

// Deriva Unidade/Diretoria da lista Diretorias pelo e-mail do gestor (1o match).
async function derivarUnidadeDiretoria(client, siteId, gestorEmail) {
  try {
    const dl = await client.api('/sites/' + siteId + '/lists').filter("displayName eq '" + LIST_DIR + "'").get();
    if (!dl.value || !dl.value.length) return null;
    const dlId = dl.value[0].id;
    const items = await client.api('/sites/' + siteId + '/lists/' + dlId + '/items?expand=fields&$top=500').get();
    const alvo = gestorEmail.toLowerCase();
    for (const it of (items.value || [])) {
      const f = it.fields || {};
      const email = String(f.Aprovador || f.aprovador || f.Email || f.field_1 || '').trim().toLowerCase();
      if (email && email === alvo) {
        return { unidade: f.Unidade || f.unidade || '', diretoria: f.Diretoria || f.diretoria || f.Title || '' };
      }
    }
  } catch (e) { /* ignore */ }
  return null;
}

async function lerLedger(client, siteId) {
  try {
    const r = await client.api('/sites/' + siteId + '/drive/root:/' + LEDGER_PATH + ':/content').get();
    if (r && typeof r === 'object' && r.processados) return r;
    if (typeof r === 'string') { try { return JSON.parse(r); } catch (e) { return { processados: {} }; } }
    return { processados: {} };
  } catch (e) { return { processados: {} }; } // 404 = ainda nao existe
}
async function salvarLedger(client, siteId, ledger) {
  ledger.atualizadoEm = new Date().toISOString();
  await client.api('/sites/' + siteId + '/drive/root:/' + LEDGER_PATH + ':/content')
    .header('Content-Type', 'application/json')
    .put(Buffer.from(JSON.stringify(ledger), 'utf-8'));
}

function classificar(msg, fornSig) {
  const assunto = String(msg.subject || '');
  const from = String((msg.from && msg.from.emailAddress && msg.from.emailAddress.address) || '').toLowerCase();
  const dominio = from.indexOf('@') > 0 ? from.split('@')[1] : '';
  const porAssunto = ASSUNTO_NF.test(assunto);
  const porRemetente = (from && fornSig.emails.has(from)) || (dominio && fornSig.dominios.has(dominio));
  return { ehNF: porAssunto || porRemetente, porAssunto, porRemetente, from, assunto };
}

async function anexosPdf(client, gestor, msgId) {
  const resp = await client.api('/users/' + encodeURIComponent(gestor) + '/messages/' + msgId + '/attachments').get();
  return (resp.value || []).filter(function (a) {
    const isFile = a['@odata.type'] === '#microsoft.graph.fileAttachment' || a.contentBytes;
    const isPdf = String(a.contentType || '').toLowerCase().indexOf('pdf') >= 0 || /\.pdf$/i.test(String(a.name || ''));
    return isFile && isPdf && a.contentBytes;
  });
}

module.exports = async function (context, req) {
  const diag = { step: 'start' };
  try {
    // Auth: secret (cron) OU admin (navegador).
    const secret = process.env.AUTOMACAO_EMAILS_SECRET;
    const hdr = (req.headers && (req.headers['x-automacao-secret'] || req.headers['X-Automacao-Secret'])) || '';
    const viaSecret = !!(secret && hdr && hdr === secret);
    if (!viaSecret) { const authz = await requireAdmin(context, req); if (!authz) return; }

    const gestor = String((req.query && req.query.gestor) || '').trim().toLowerCase();
    if (!gestor || gestor.indexOf('@') < 0) {
      context.res = { status: 400, headers: { 'Content-Type': 'application/json' }, body: { error: 'Passe ?gestor=email@pronep.com.br' } };
      return;
    }
    const dias = Math.min(30, Math.max(1, parseInt((req.query && req.query.dias) || '3', 10) || 3));
    const limite = Math.min(100, Math.max(1, parseInt((req.query && req.query.limite) || '50', 10) || 50));
    const dryRun = req.query && (req.query.dryRun === '1' || req.query.dryRun === 'true');

    const client = getGraphClient();
    diag.step = 'resolve';
    const siteId = await resolveSiteId(client);
    const driveId = await resolveDriveId(client, siteId);

    // Destino (Unidade/Diretoria): query params ou derivado da lista Diretorias.
    let unidade = String((req.query && req.query.unidade) || '').trim();
    let diretoria = String((req.query && req.query.diretoria) || '').trim();
    if (!unidade || !diretoria) {
      const der = await derivarUnidadeDiretoria(client, siteId, gestor);
      if (der) { unidade = unidade || der.unidade; diretoria = diretoria || der.diretoria; }
    }
    if (!unidade || !diretoria) {
      context.res = { status: 400, headers: { 'Content-Type': 'application/json' },
        body: { error: 'Nao consegui determinar Unidade/Diretoria do gestor. Passe ?unidade=&diretoria=', gestor: gestor } };
      return;
    }

    diag.step = 'sinais';
    const fornSig = await fornecedorSignals(client, siteId);
    const ledger = await lerLedger(client, siteId);
    ledger.processados = ledger.processados || {};

    diag.step = 'listar_mail';
    const desde = new Date(Date.now() - dias * 24 * 60 * 60 * 1000).toISOString();
    const listaUrl = '/users/' + encodeURIComponent(gestor) + '/mailFolders/inbox/messages'
      + '?$select=id,subject,from,receivedDateTime,hasAttachments'
      + '&$filter=' + encodeURIComponent('receivedDateTime ge ' + desde)
      + '&$orderby=' + encodeURIComponent('receivedDateTime desc')
      + '&$top=' + limite;
    const msgsResp = await client.api(listaUrl).get();
    const msgs = msgsResp.value || [];

    const pasta = PASTA_RAIZ + '/' + unidade + '/Diretoria ' + diretoria;
    const baixados = [], candidatos = [], ignorados = [];

    for (const m of msgs) {
      if (!m.hasAttachments) { ignorados.push({ assunto: m.subject, motivo: 'sem_anexo' }); continue; }
      if (ledger.processados[m.id]) { ignorados.push({ assunto: m.subject, motivo: 'ja_processado' }); continue; }
      const cls = classificar(m, fornSig);
      if (!cls.ehNF) { ignorados.push({ assunto: m.subject, motivo: 'nao_parece_nf' }); continue; }

      const pdfs = await anexosPdf(client, gestor, m.id);
      if (!pdfs.length) { ignorados.push({ assunto: m.subject, motivo: 'sem_pdf' }); continue; }

      candidatos.push({ assunto: m.subject, de: cls.from, quando: m.receivedDateTime, pdfs: pdfs.map(p => p.name), porAssunto: cls.porAssunto, porRemetente: cls.porRemetente });
      if (dryRun) continue;

      for (const p of pdfs) {
        const nome = String(p.name || ('anexo-' + m.id + '.pdf')).replace(/[\\/:*?"<>|]/g, '_');
        const uploadPath = '/sites/' + siteId + '/drive/root:/' + encodeURIComponent(pasta) + '/' + encodeURIComponent(nome) + ':/content';
        await client.api(uploadPath).header('Content-Type', 'application/pdf').put(Buffer.from(p.contentBytes, 'base64'));
        baixados.push({ assunto: m.subject, arquivo: nome });
      }
      ledger.processados[m.id] = new Date().toISOString();
    }

    if (!dryRun && baixados.length) await salvarLedger(client, siteId, ledger);

    diag.step = 'done';
    context.res = {
      status: 200, headers: { 'Content-Type': 'application/json', 'Cache-Control': 'no-store' },
      body: {
        ok: true, dryRun: !!dryRun, gestor: gestor, unidade: unidade, diretoria: diretoria,
        pastaDestino: pasta, janelaDias: dias, avaliados: msgs.length,
        candidatos: candidatos, baixados: baixados,
        ignoradosResumo: ignorados.reduce((a, x) => { a[x.motivo] = (a[x.motivo] || 0) + 1; return a; }, {}),
        fornecedoresConhecidos: fornSig.emails.size
      }
    };
  } catch (err) {
    context.log && context.log.error && context.log.error('VarrerEmailsNF:', err);
    context.res = { status: 500, headers: { 'Content-Type': 'application/json' },
      body: { error: (err && err.message) || String(err), statusCode: err && err.statusCode, diag: diag } };
  }
};
