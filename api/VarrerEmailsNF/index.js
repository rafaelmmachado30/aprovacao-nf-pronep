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
const { computar } = require('../shared/checklistRecorrentes');
const { _norm } = require('../shared/recorrentes');

const LIST_NOTAS = 'PRONEP-NF-NotasFiscais';
const LIST_FORN = 'PRONEP-NF-Fornecedores';
const LIST_DIR = 'PRONEP-NF-Diretorias';
const PASTA_RAIZ = 'Novas NFs - Automacao';
const LEDGER_PATH = '_automacao/emails_ledger.json';
// Dominio interno da empresa — NUNCA conta como "remetente fornecedor" (senao todo
// colega que encaminha vira falso match). Configuravel por env.
const DOMINIO_INTERNO = (process.env.EMAIL_DOMINIO_INTERNO || 'pronep.com.br').toLowerCase();
// Classificacao calibrada com dados reais (dry-run 16/07):
//  FORTE   = "nota fiscal"/NF/NFe/NFS-e/DANFE no assunto OU nome do PDF -> candidato sozinho.
//  FRACO   = fatura/boleto -> so entra se corroborado (remetente conhecido ou encaminhado).
//  NEGATIVO= aviso de cobranca/debito (nao e NF a lancar) -> exclui, salvo sinal FORTE.
const RE_FORTE   = /nota\s*fiscal|\bnfe?\b|\bnfs-?e\b|danfe/i;
const RE_FRACO   = /fatura|boleto/i;
const RE_NEG     = /n[ãa]o\s*identificad|d[ée]bito|em\s*aberto|inadimpl|pend[êe]ncia\s*de\s*pagamento|cobran[çc]a\s*autom/i;
const RE_FORTE_ARQ = /\b(nf|nfe|nfse|danfe)\b|nota[\s_-]*fiscal/i;
const RE_FRACO_ARQ = /fatura|boleto/i;
const RE_ENCAMINHADO = /^\s*(enc|fw|fwd|res|encaminhad)/i;

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
          if (em.indexOf('@') <= 0) continue;
          const dom = em.split('@')[1];
          if (dom === DOMINIO_INTERNO) continue; // ignora fornecedor cadastrado c/ e-mail interno
          emails.add(em); dominios.add(dom);
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

// Classifica um e-mail (assunto + nome dos PDFs + remetente). Retorna nivel de
// confianca pra o gestor priorizar: alta (sinal forte de NF), media (fraco +
// remetente conhecido), baixa (fraco + encaminhado). null = nao e NF.
function classificar(msg, pdfs, fornSig) {
  const assunto = String(msg.subject || '');
  const from = String((msg.from && msg.from.emailAddress && msg.from.emailAddress.address) || '').toLowerCase();
  const dominio = from.indexOf('@') > 0 ? from.split('@')[1] : '';
  const nomes = (pdfs || []).map(function (p) { return String(p.name || ''); }).join(' ');

  const forte = RE_FORTE.test(assunto) || RE_FORTE_ARQ.test(nomes);
  const fraco = RE_FRACO.test(assunto) || RE_FRACO_ARQ.test(nomes);
  const negativo = RE_NEG.test(assunto);
  const encaminhado = RE_ENCAMINHADO.test(assunto);
  // Remetente conhecido: dominio interno ja foi excluido do fornSig.
  const porRemetente = (from && fornSig.emails.has(from)) || (dominio && fornSig.dominios.has(dominio));

  let confianca = null, motivo = '';
  if (forte) { confianca = 'alta'; motivo = 'sinal_forte_nf'; }
  else if (negativo) { confianca = null; motivo = 'aviso_cobranca_debito'; }
  else if (fraco && porRemetente) { confianca = 'media'; motivo = 'fraco+remetente_conhecido'; }
  else if (fraco && encaminhado) { confianca = 'baixa'; motivo = 'fraco+encaminhado'; }
  else { confianca = null; motivo = 'nao_parece_nf'; }

  return { ehNF: !!confianca, confianca: confianca, motivo: motivo, from: from, assunto: assunto,
           sinais: { forte: forte, fraco: fraco, negativo: negativo, encaminhado: encaminhado, porRemetente: !!porRemetente } };
}

async function anexosPdf(client, gestor, msgId) {
  const resp = await client.api('/users/' + encodeURIComponent(gestor) + '/messages/' + msgId + '/attachments').get();
  return (resp.value || []).filter(function (a) {
    const isFile = a['@odata.type'] === '#microsoft.graph.fileAttachment' || a.contentBytes;
    const isPdf = String(a.contentType || '').toLowerCase().indexOf('pdf') >= 0 || /\.pdf$/i.test(String(a.name || ''));
    return isFile && isPdf && a.contentBytes;
  });
}

// --- Corroboracao pelo Fechamento do Mes (ideia do Rafa) -------------------
// Normaliza pra comparar nomes de fornecedor com assunto/nome de arquivo.
function _slug(s) {
  return String(s || '').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-z0-9]+/g, ' ').trim();
}

// Resolve a lista NotasFiscais + invColMap (necessarios pro computar()).
async function resolveNotasList(client, siteId) {
  const lr = await client.api('/sites/' + siteId + '/lists').filter("displayName eq '" + LIST_NOTAS + "'").get();
  if (!lr.value || !lr.value.length) throw new Error('Lista ' + LIST_NOTAS + ' nao encontrada');
  const listNotasId = lr.value[0].id;
  const cols = await client.api('/sites/' + siteId + '/lists/' + listNotasId + '/columns').get();
  const invColMap = {};
  for (const c of (cols.value || [])) { if (c.displayName && c.name) invColMap[c.name] = c.displayName; }
  return { listNotasId: listNotasId, invColMap: invColMap };
}

// Fornecedores RECORRENTES ainda PENDENTES no mes (atrasada/risco/aguardando), da
// diretoria — sao os que "deviam ter mandado NF e ainda nao mandaram". Best-effort.
async function esperadosDoFechamento(client, siteId, diretoria) {
  try {
    const { listNotasId, invColMap } = await resolveNotasList(client, siteId);
    const agoraBRT = new Date(Date.now() - 3 * 60 * 60 * 1000);
    const r = await computar(client, siteId, listNotasId, invColMap, {
      scopeNorm: diretoria ? [_norm(diretoria)] : null,
      ano: agoraBRT.getUTCFullYear(), mes: agoraBRT.getUTCMonth()
    });
    const pendentes = (r.contas || []).filter(function (c) {
      return ['atrasada', 'risco', 'aguardando'].indexOf(c.status) >= 0;
    });
    return pendentes.map(function (c) {
      const slug = _slug(c.fornecedor);
      return { nome: c.fornecedor, slug: slug, tokens: slug.split(' ').filter(function (t) { return t.length >= 5; }) };
    }).filter(function (e) { return e.slug; });
  } catch (e) { return []; } // se o Fechamento falhar, seguimos sem corroboracao
}

// Um candidato "casa" com um fornecedor esperado se o nome (ou um token distintivo
// dele) aparece no assunto ou no nome de algum PDF.
function corrobora(esperados, assunto, pdfNames) {
  const hay = _slug(assunto + ' ' + (pdfNames || []).join(' '));
  for (const e of (esperados || [])) {
    if (e.slug && hay.indexOf(e.slug) >= 0) return e.nome;
    for (const t of e.tokens) { if (hay.indexOf(t) >= 0) return e.nome; }
  }
  return null;
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

    // Corroboracao (ideia do Rafa): fornecedores recorrentes ainda PENDENTES no mes,
    // pela diretoria. Se um deles aparece no e-mail, sobe a confianca (ou resgata um fraco).
    diag.step = 'esperados';
    const esperados = await esperadosDoFechamento(client, siteId, diretoria);
    diag.esperados = esperados.length;

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

      // Busca os PDFs ANTES de classificar — o nome do arquivo (NF.../DANFE...) e sinal forte.
      const pdfs = await anexosPdf(client, gestor, m.id);
      if (!pdfs.length) { ignorados.push({ assunto: m.subject, motivo: 'sem_pdf' }); continue; }

      const cls = classificar(m, pdfs, fornSig);
      const pdfNames = pdfs.map(p => p.name);
      const fornEsperado = corrobora(esperados, m.subject, pdfNames);

      // Corroboracao pelo Fechamento: reforca candidato existente (-> alta) e resgata
      // um sinal FRACO nao-negativo (ex.: boleto de um recorrente que ainda nao chegou).
      let ehNF = cls.ehNF, confianca = cls.confianca, motivo = cls.motivo;
      if (fornEsperado) {
        if (ehNF) { confianca = 'alta'; motivo = cls.motivo + '+esperado_fechamento'; }
        else if (cls.sinais.fraco && !cls.sinais.negativo) { ehNF = true; confianca = 'alta'; motivo = 'fraco+esperado_fechamento'; }
      }
      if (!ehNF) { ignorados.push({ assunto: m.subject, motivo: cls.motivo }); continue; }

      candidatos.push({ assunto: m.subject, de: cls.from, quando: m.receivedDateTime,
        pdfs: pdfNames, confianca: confianca, motivo: motivo, esperado: fornEsperado || null, sinais: cls.sinais });
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
        fornecedoresConhecidos: fornSig.emails.size,
        esperadosFechamento: esperados.map(function (e) { return e.nome; })
      }
    };
  } catch (err) {
    context.log && context.log.error && context.log.error('VarrerEmailsNF:', err);
    context.res = { status: 500, headers: { 'Content-Type': 'application/json' },
      body: { error: (err && err.message) || String(err), statusCode: err && err.statusCode, diag: diag } };
  }
};
