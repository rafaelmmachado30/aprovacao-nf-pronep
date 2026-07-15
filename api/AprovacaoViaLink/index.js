/**
 * /api/AprovacaoViaLink?token=XYZ
 *
 * Endpoint que valida um token JWT e executa Aprovar ou Rejeitar diretamente.
 * Usado pelos botoes nos e-mails de notificacao.
 *
 * Token JWT contem: { itemId, aprovador (email), action: 'aprovar'|'rejeitar', exp }
 * Token expira em 7 dias.
 * Uso unico: validacao adicional via Status atual da NF (se ja Aprovada/Rejeitada, recusa).
 *
 * App Setting necessaria:
 *   LINK_APROVACAO_SECRET = string aleatoria (ex: gerar com `openssl rand -base64 32`)
 */

const jwt = require('jsonwebtoken');
require('isomorphic-fetch');
const { getGraphClient } = require('../shared/graph');

const LIST_NOTAS = 'PRONEP-NF-NotasFiscais';

async function resolveListaNotas(client) {
  const host = process.env.SHAREPOINT_SITE_HOSTNAME;
  const path = process.env.SHAREPOINT_SITE_PATH;
  const siteResp = await client.api(`/sites/${host}:${path}`).get();
  const siteId = siteResp.id;
  const lists = await client.api(`/sites/${siteId}/lists`).filter(`displayName eq '${LIST_NOTAS}'`).get();
  return { siteId, listId: lists.value[0].id };
}

// C7: escape de HTML pra valores dinamicos interpolados nas paginas de resposta.
function esc(s) {
  return String(s == null ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

function htmlPagina(titulo, mensagem, cor, detalhe) {
  const corBg = cor || '#1F4E79';
  return `<!DOCTYPE html><html lang="pt-BR"><head><meta charset="UTF-8"><title>${titulo}</title>
<style>
  body { font-family: 'Segoe UI', Roboto, Arial, sans-serif; background: #F4F8FB; margin: 0; padding: 40px 16px; display: flex; justify-content: center; }
  .card { background: #fff; max-width: 540px; width: 100%; border-radius: 12px; box-shadow: 0 4px 24px rgba(31,78,121,0.12); overflow: hidden; }
  .header { background: ${corBg}; color: #fff; padding: 28px; text-align: center; }
  .header .icon { font-size: 48px; line-height: 1; margin-bottom: 8px; }
  .header h1 { margin: 0; font-size: 22px; font-weight: 600; }
  .body { padding: 28px; color: #2C3E50; font-size: 15px; line-height: 1.6; text-align: center; }
  .body .detalhe { background: #F4F8FB; border-radius: 8px; padding: 14px; margin-top: 16px; font-size: 13px; color: #647883; }
  .btn { display: inline-block; margin-top: 20px; background: ${corBg}; color: #fff; padding: 12px 28px; border-radius: 6px; text-decoration: none; font-weight: 600; }
  .footer { background: #F4F8FB; color: #647883; padding: 14px 28px; font-size: 11px; text-align: center; border-top: 1px solid #DCE3E9; }
</style></head>
<body><div class="card">
  <div class="header"><div class="icon">${cor === '#2E7D32' ? '✓' : cor === '#C62828' ? '✕' : '⚠'}</div><h1>${titulo}</h1></div>
  <div class="body"><p>${mensagem}</p>${detalhe ? `<div class="detalhe">${detalhe}</div>` : ''}
    <a class="btn" href="https://purple-forest-09588fe10.7.azurestaticapps.net/">Abrir Sistema</a>
  </div>
  <div class="footer">Sistema de Aprovacao de NF · Pronep Life Care</div>
</div></body></html>`;
}

module.exports = async function (context, req) {
  try {
    const token = req.query.token;
    if (!token) {
      context.res = { status: 400, headers: { 'Content-Type': 'text/html; charset=utf-8' },
        body: htmlPagina('Link invalido', 'Esse link nao contem o token necessario.', '#C62828') };
      return;
    }

    const secret = process.env.LINK_APROVACAO_SECRET;
    if (!secret) {
      context.res = { status: 500, headers: { 'Content-Type': 'text/html; charset=utf-8' },
        body: htmlPagina('Sistema mal configurado', 'LINK_APROVACAO_SECRET nao definido nas App Settings.', '#C62828') };
      return;
    }

    let payload;
    try {
      // C7: trava o algoritmo em HS256 (evita ataque de confusao de algoritmo).
      payload = jwt.verify(token, secret, { algorithms: ['HS256'] });
    } catch (e) {
      context.res = { status: 401, headers: { 'Content-Type': 'text/html; charset=utf-8' },
        body: htmlPagina('Link expirado ou invalido', `Esse link nao eh mais valido (${esc(e.message)}). Pode ter sido usado antes ou expirado.`, '#C62828') };
      return;
    }

    const { itemId, aprovador, action } = payload;
    if (!itemId || !aprovador || !['aprovar','rejeitar'].includes(action)) {
      context.res = { status: 400, headers: { 'Content-Type': 'text/html; charset=utf-8' },
        body: htmlPagina('Link com dados invalidos', 'Token nao tem os campos esperados.', '#C62828') };
      return;
    }

    // Le NF atual pra validar Status (uso unico = se ja foi processada, recusa)
    const client = await getGraphClient();
    const { siteId, listId } = await resolveListaNotas(client);
    const item = await client.api(`/sites/${siteId}/lists/${listId}/items/${itemId}?expand=fields`).get();
    const fields = item.fields || {};
    // Le Status pelos possiveis nomes (Status display name ou field_9)
    const statusAtual = fields.Status || fields.field_9 || '';
    const aprovadorAtual = (fields.AprovadorAtual || fields.field_8 || '').toLowerCase();

    if (statusAtual === 'Aprovada') {
      context.res = { status: 200, headers: { 'Content-Type': 'text/html; charset=utf-8' },
        body: htmlPagina('NF ja foi aprovada', 'Esta nota fiscal ja foi aprovada anteriormente. Nada a fazer.', '#647883',
          `Status: <b>Aprovada</b>`) };
      return;
    }
    if (statusAtual === 'Rejeitada') {
      context.res = { status: 200, headers: { 'Content-Type': 'text/html; charset=utf-8' },
        body: htmlPagina('NF ja foi rejeitada', 'Esta nota fiscal ja foi rejeitada anteriormente.', '#647883',
          `Status: <b>Rejeitada</b>`) };
      return;
    }
    if (aprovadorAtual && aprovadorAtual !== aprovador.toLowerCase()) {
      context.res = { status: 403, headers: { 'Content-Type': 'text/html; charset=utf-8' },
        body: htmlPagina('Voce nao eh o aprovador atual', `Esta NF foi reatribuida. Aprovador atual: ${esc(aprovadorAtual)}`, '#C62828') };
      return;
    }

    // Executa a acao chamando a Function correspondente em-processo
    if (action === 'aprovar') {
      const aprovarHandler = require('../AprovarNota/index.js');
      const fakeReq = {
        body: { id: itemId },
        headers: {
          'x-ms-client-principal': Buffer.from(JSON.stringify({
            userDetails: aprovador, userId: aprovador, userRoles: ['authenticated']
          })).toString('base64')
        },
        query: {}
      };
      const fakeCtx = { res: null, log: { error: () => {}, info: () => {} } };
      fakeCtx.log.error = () => {}; fakeCtx.log.info = () => {};
      await aprovarHandler(fakeCtx, fakeReq);
      if (fakeCtx.res && fakeCtx.res.status === 200) {
        context.res = { status: 200, headers: { 'Content-Type': 'text/html; charset=utf-8' },
          body: htmlPagina('NF aprovada com sucesso!', `A NF #${esc(itemId)} foi aprovada por <b>${esc(aprovador)}</b>. O PDF com watermark foi arquivado em Notas Aprovadas e o submitter foi notificado.`, '#2E7D32') };
      } else {
        const errBody = fakeCtx.res ? fakeCtx.res.body : { error: 'erro desconhecido' };
        context.res = { status: 500, headers: { 'Content-Type': 'text/html; charset=utf-8' },
          body: htmlPagina('Erro ao aprovar', errBody.error || 'Erro desconhecido', '#C62828') };
      }
    } else {
      // rejeitar
      const rejeitarHandler = require('../RejeitarNota/index.js');
      const fakeReq = {
        body: { id: itemId, motivo: 'Rejeitada via link do e-mail (sem motivo detalhado informado)', observacao: '' },
        headers: {
          'x-ms-client-principal': Buffer.from(JSON.stringify({
            userDetails: aprovador, userId: aprovador, userRoles: ['authenticated']
          })).toString('base64')
        },
        query: {}
      };
      const fakeCtx = { res: null, log: { error: () => {}, info: () => {} } };
      await rejeitarHandler(fakeCtx, fakeReq);
      if (fakeCtx.res && fakeCtx.res.status === 200) {
        context.res = { status: 200, headers: { 'Content-Type': 'text/html; charset=utf-8' },
          body: htmlPagina('NF rejeitada', `A NF #${esc(itemId)} foi rejeitada por <b>${esc(aprovador)}</b>. O PDF com watermark vermelho foi arquivado em Rejeitadas e o submitter foi notificado.`, '#C62828') };
      } else {
        const errBody = fakeCtx.res ? fakeCtx.res.body : { error: 'erro desconhecido' };
        context.res = { status: 500, headers: { 'Content-Type': 'text/html; charset=utf-8' },
          body: htmlPagina('Erro ao rejeitar', errBody.error || 'Erro desconhecido', '#C62828') };
      }
    }
  } catch (err) {
    context.log && context.log.error && context.log.error('AprovacaoViaLink error:', err);
    context.res = { status: 500, headers: { 'Content-Type': 'text/html; charset=utf-8' },
      body: htmlPagina('Erro interno', (err && err.message) || String(err), '#C62828') };
  }
};
