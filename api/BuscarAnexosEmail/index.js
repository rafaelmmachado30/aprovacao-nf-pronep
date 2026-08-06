/**
 * /api/BuscarAnexosEmail  — acha, nas caixas de e-mail, os anexos de UMA NF do quadro.
 *
 * Tira do caminho o passo manual de hoje: baixar do Outlook, salvar numa pasta e
 * subir no sistema. Aqui o card pergunta ao e-mail direto.
 *
 *   GET  ?docId=<id>            lista candidatos (NAO baixa nada)
 *   POST { docId, escolhas:[{caixa,msgId,anexoId}] }   baixa para a Caixa de Entrada
 *
 * NUNCA ANEXA SOZINHO. Propoe, o usuario confirma. Anexar o boleto errado numa NF
 * e pior do que nao anexar nada — o financeiro pagaria o documento de outra conta.
 *
 * NAO LE O CONTEUDO DOS PDFs. So nome do arquivo, remetente, assunto e data:
 * abrir e interpretar cada anexo seria caro e lento, e os quatro sinais abaixo ja
 * resolvem o caso comum.
 *
 * SINAIS, do mais forte ao mais fraco:
 *   chave de acesso de 44 digitos no nome  -> prova (unica no Brasil inteiro)
 *   numero da NF como BLOCO INTEIRO        -> a regra que o Rafael ja usa a mao
 *   CNPJ/dominio do remetente == fornecedor
 *   a CAIXA indica a unidade (financeiro.sp -> SP)
 *
 * COMECA RESTRITIVO de proposito. Oito caixas com muitos e-mails geram muito
 * candidato fraco; se a lista vier com vinte arquivos, o usuario perde mais tempo
 * do que baixando a mao. Por padrao so devolve sinal forte — ?amplo=1 relaxa.
 */

require('isomorphic-fetch');
const { getUser } = require('../shared/auth');
const { resolveAuthz } = require('../shared/authz');
const { getGraphClient, resolveSiteId } = require('../shared/graph');
const { resolveListId, LIST_DOCFIS, soDigitos } = require('../shared/documentosFiscais');
/* Reaproveita a Caixa de Entrada em vez de criar um segundo lugar para guardar
   arquivo: dois lugares divergem, e o usuario ja mantem um controle visual nela. */
const caixaEntrada = require('../CaixaEntrada/index.js');

const ORCAMENTO_MS = 28000;
const MAX_MSGS_POR_CAIXA = 60;

/* Caixas por unidade. As de TI recebem de tudo, entao entram sempre.
   A caixa e um sinal por si: boleto que chega em financeiro.sp e quase sempre
   conta da PRONEP SP. */
const CAIXAS = {
  RJ: ['financeiro.rj@pronep.com.br', 'administrativo.rj@pronep.com.br'],
  SP: ['financeiro.sp@pronep.com.br', 'administrativo.sp@pronep.com.br'],
  ES: ['financeiro.vix@pronep.com.br', 'administrativo.vix@pronep.com.br']
};
const CAIXAS_TI = ['gestaoti@pronep.com.br', 'rafael.machado@pronep.com.br'];

function caixasPara(unidade) {
  const u = String(unidade || '').toUpperCase();
  return (CAIXAS[u] || []).concat(CAIXAS_TI);
}

function slug(s) {
  return String(s || '').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-z0-9]+/g, ' ').trim();
}

/**
 * Pontua um anexo contra a NF do card.
 * @returns {{pontos:number, motivos:[string]}}
 */
function pontuar(anexoNome, msg, card, caixa) {
  const motivos = [];
  let pontos = 0;

  const nome = String(anexoNome || '');
  const digitosNome = nome.replace(/\D/g, '');

  /* 44 digitos sao unicos no Brasil: isto e prova, nao indicio. */
  const ch = soDigitos(card.chaveAcesso);
  if (ch.length === 44 && digitosNome.indexOf(ch) >= 0) {
    pontos += 100; motivos.push('chave de acesso no arquivo');
  }

  /* Bloco INTEIRO de digitos, nunca "contem": 1440963 nao pode casar 21440963
     nem 14409631, que sao notas diferentes. */
  const num = String(card.numeroNF || '').replace(/^0+/, '');
  if (num && num.length >= 3) {
    const blocos = nome.match(/\d+/g) || [];
    for (const b of blocos) {
      if (b.replace(/^0+/, '') === num) {
        pontos += 50; motivos.push('número da NF no arquivo'); break;
      }
    }
  }

  /* Remetente. Peso ALTO de proposito, o bastante para o anexo passar sozinho.
     E um sinal ESPECIFICO deste card, nao generico: o e-mail veio da empresa que
     emitiu esta nota. E cobre justamente o caso que o nome de arquivo nao cobre —
     boleto_d516b2e3-ff4a-...pdf da Control iD nao tem numero nenhum, e sem isto
     nunca apareceria.
     A assimetria manda: um candidato a mais custa uma olhada; um candidato a menos
     custa voltar ao Outlook, que e exatamente o que este endpoint existe para
     evitar. */
  const from = String((msg.from && msg.from.emailAddress && msg.from.emailAddress.address) || '').toLowerCase();
  const nomeFrom = String((msg.from && msg.from.emailAddress && msg.from.emailAddress.name) || '');
  const fornSlug = slug(card.emitenteNome);
  const tokens = fornSlug.split(' ').filter(function (t) { return t.length >= 5; });
  if (tokens.length && (tokens.some(function (t) { return from.indexOf(t) >= 0; }) ||
                        tokens.some(function (t) { return slug(nomeFrom).indexOf(t) >= 0; }))) {
    pontos += 50; motivos.push('remetente é o fornecedor');
  }

  /* Numero da NF no assunto — mais fraco que no arquivo, mas ajuda. */
  const assunto = String(msg.subject || '');
  if (num && num.length >= 4 && (assunto.match(/\d+/g) || []).some(function (b) {
    return b.replace(/^0+/, '') === num;
  })) {
    pontos += 25; motivos.push('número da NF no assunto');
  }

  /* A caixa da unidade e corroboracao, nunca sozinha: quase toda conta tem uma. */
  const u = String(card.unidade || '').toUpperCase();
  if (u && (CAIXAS[u] || []).indexOf(caixa) >= 0) {
    pontos += 10; motivos.push('caixa da unidade ' + u);
  }

  return { pontos: pontos, motivos: motivos };
}

module.exports = async function (context, req) {
  const t0 = Date.now();
  const q = req.query || {};
  const amplo = q.amplo === '1' || q.amplo === 'true';
  const dias = Math.max(1, Math.min(90, parseInt(q.dias, 10) || 30));
  const diag = { step: 'init', dias: dias, amplo: amplo, caixas: [], timeMs: 0 };

  try {
    const authz = await resolveAuthz(req);
    if (!authz) {
      context.res = { status: 401, headers: { 'Content-Type': 'application/json' },
        body: { error: 'Nao autenticado' } };
      return;
    }

    const body = req.body || {};
    const docId = String(q.docId || body.docId || '');
    if (!docId) throw new Error('docId obrigatorio');

    const client = await getGraphClient();
    const siteId = await resolveSiteId(client);
    const listId = await resolveListId(client, siteId, LIST_DOCFIS);
    if (!listId) throw new Error('Lista de documentos nao existe');

    const item = await client.api('/sites/' + siteId + '/lists/' + listId + '/items/' + docId)
      .expand('fields').get();
    const f = (item && item.fields) || {};
    const card = {
      numeroNF: f.NumeroNF || '',
      chaveAcesso: f.ChaveAcesso || '',
      emitenteNome: f.EmitenteNome || '',
      emitenteCNPJ: f.EmitenteCNPJ || '',
      unidade: f.UnidadeOmie || ''
    };
    diag.card = card;

    /* ---------- POST: baixa os escolhidos para a Caixa de Entrada ---------- */
    if (String(req.method || '').toUpperCase() === 'POST') {
      const escolhas = Array.isArray(body.escolhas) ? body.escolhas : [];
      if (!escolhas.length) throw new Error('Nenhum anexo escolhido');

      const user = await getUser(req);
      const baixados = [];
      const falhas = [];
      for (const e of escolhas) {
        if (Date.now() - t0 > ORCAMENTO_MS) { falhas.push({ msgId: e.msgId, erro: 'tempo' }); continue; }
        try {
          const at = await client.api('/users/' + encodeURIComponent(e.caixa) +
            '/messages/' + e.msgId + '/attachments/' + e.anexoId).get();
          if (!at || !at.contentBytes) throw new Error('anexo sem conteudo');
          const ctxFake = {};
          await caixaEntrada(ctxFake, {
            method: 'POST', headers: req.headers,
            body: { action: 'upload', fileBase64: at.contentBytes, fileName: at.name }
          });
          const corpo = (ctxFake.res && ctxFake.res.body) || {};
          if (!corpo.ok) throw new Error(corpo.error || 'falha ao guardar');
          baixados.push({ nome: at.name, id: corpo.id });
        } catch (eb) {
          falhas.push({ msgId: e.msgId, erro: eb.message });
        }
      }
      diag.timeMs = Date.now() - t0;
      context.res = { status: 200, headers: { 'Content-Type': 'application/json' },
        body: Object.assign({ ok: falhas.length === 0, baixados: baixados,
          falhas: falhas, usuario: (user && user.email) || '' }, diag) };
      return;
    }

    /* ---------- GET: procura candidatos ---------- */
    diag.step = 'buscar';
    const desde = new Date(Date.now() - dias * 86400000).toISOString();
    const caixas = caixasPara(card.unidade);
    const candidatos = [];

    for (const caixa of caixas) {
      const bloco = { caixa: caixa, mensagens: 0, comAnexo: 0, erro: null };
      diag.caixas.push(bloco);
      if (Date.now() - t0 > ORCAMENTO_MS) { bloco.erro = 'sem tempo'; continue; }

      let msgs = [];
      try {
        const url = '/users/' + encodeURIComponent(caixa) + '/mailFolders/inbox/messages'
          + '?$select=id,subject,from,receivedDateTime,hasAttachments'
          + '&$filter=' + encodeURIComponent('receivedDateTime ge ' + desde + ' and hasAttachments eq true')
          + '&$top=' + MAX_MSGS_POR_CAIXA + '&$orderby=receivedDateTime desc';
        const r = await client.api(url).get();
        msgs = r.value || [];
      } catch (em) {
        /* Caixa sem permissao nao derruba a busca nas outras — e informa, porque
           "nao achei" por falta de acesso e diferente de "nao existe". */
        bloco.erro = em.message;
        continue;
      }
      bloco.mensagens = msgs.length;

      for (const m of msgs) {
        if (Date.now() - t0 > ORCAMENTO_MS) break;
        /* Pre-filtro barato: so abre os anexos de mensagens que ja tem algum
           indicio no assunto ou no remetente. Sem isto seriam 60 chamadas de
           anexo por caixa, 480 no total. */
        const pre = pontuar('', m, card, caixa);
        const temIndicio = pre.pontos >= 25;
        if (!temIndicio && !amplo) continue;

        let anexos = [];
        try {
          const ra = await client.api('/users/' + encodeURIComponent(caixa) +
            '/messages/' + m.id + '/attachments?$select=id,name,contentType,size').get();
          anexos = (ra.value || []).filter(function (a) {
            return /pdf/i.test(String(a.contentType || '')) || /\.pdf$/i.test(String(a.name || ''));
          });
        } catch (ea) { continue; }
        if (anexos.length) bloco.comAnexo++;

        for (const a of anexos) {
          const p = pontuar(a.name, m, card, caixa);
          /* 50 = pelo menos numero da NF no arquivo, ou chave. Abaixo disso e
             palpite, e palpite em lista longa e ruido. */
          if (p.pontos < (amplo ? 25 : 50)) continue;
          candidatos.push({
            caixa: caixa, msgId: m.id, anexoId: a.id,
            nome: a.name, tamanho: a.size || 0,
            de: (m.from && m.from.emailAddress && m.from.emailAddress.address) || '',
            deNome: (m.from && m.from.emailAddress && m.from.emailAddress.name) || '',
            assunto: m.subject || '',
            recebidoEm: m.receivedDateTime || '',
            pontos: p.pontos, motivos: p.motivos,
            confianca: p.pontos >= 100 ? 'alta' : (p.pontos >= 50 ? 'media' : 'baixa')
          });
        }
      }
    }

    candidatos.sort(function (a, b) { return b.pontos - a.pontos; });

    diag.step = 'done';
    diag.timeMs = Date.now() - t0;
    context.res = { status: 200, headers: { 'Content-Type': 'application/json' },
      body: Object.assign({
        ok: true,
        candidatos: candidatos.slice(0, 20),
        total: candidatos.length,
        /* Diz que existe um modo mais amplo: sem isso, "nao achei nada" parece
           veredito quando e so o filtro restritivo fazendo o trabalho dele. */
        dica: candidatos.length ? '' :
          'Nada com sinal forte. Tente ?amplo=1 para incluir candidatos fracos, ' +
          'ou aumente ?dias=' + dias + '.'
      }, diag) };
  } catch (err) {
    diag.timeMs = Date.now() - t0;
    context.res = { status: 500, headers: { 'Content-Type': 'application/json' },
      body: Object.assign({ error: (err && err.message) || String(err) }, diag) };
  }
};
