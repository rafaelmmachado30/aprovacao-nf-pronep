/**
 * /api/AprovacaoPagina?token=XYZ
 *
 * Pagina de aprovacao (opcao "botao" no WhatsApp): mostra o resumo da NF + botoes
 * grandes Aprovar / Rejeitar / Ver PDF. Cada botao aponta pro /api/AprovacaoViaLink
 * (que ja valida token + status + aprovador e executa). Assim os "botoes" ficam numa
 * pagina nossa — funciona em qualquer gateway e nao arrisca o numero do WhatsApp.
 *
 * Token 'ver' (action:'ver') carrega o resumo da NF (campo nf) pra exibir sem reler o
 * SharePoint. O Status atual e lido ao vivo pra recusar NF ja processada (uso unico).
 *
 * App Setting: LINK_APROVACAO_SECRET (o mesmo dos links de e-mail).
 */

const jwt = require('jsonwebtoken');
require('isomorphic-fetch');
const { getGraphClient } = require('../shared/graph');
const { gerarLinks } = require('../shared/email');

const LIST_NOTAS = 'PRONEP-NF-NotasFiscais';

async function resolveListaNotas(client) {
  const host = process.env.SHAREPOINT_SITE_HOSTNAME;
  const path = process.env.SHAREPOINT_SITE_PATH;
  const siteResp = await client.api(`/sites/${host}:${path}`).get();
  const siteId = siteResp.id;
  const lists = await client.api(`/sites/${siteId}/lists`).filter(`displayName eq '${LIST_NOTAS}'`).get();
  return { siteId, listId: lists.value[0].id };
}

function esc(s) {
  return String(s == null ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}
function fmtBRL(v) {
  try { return Number(v || 0).toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' }); }
  catch (e) { return 'R$ ' + v; }
}
function fmtData(s) {
  const d = String(s || '').substring(0, 10).split('-');
  return d.length === 3 ? `${d[2]}/${d[1]}/${d[0]}` : esc(s);
}

// Pagina simples (estados: link invalido / ja processada / erro).
function paginaSimples(titulo, mensagem, cor, detalhe) {
  const corBg = cor || '#1F4E79';
  const icon = cor === '#2E7D32' ? '✓' : cor === '#C62828' ? '✕' : '⚠';
  return `<!DOCTYPE html><html lang="pt-BR"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>${esc(titulo)}</title><style>
  body{font-family:'Segoe UI',Roboto,Arial,sans-serif;background:#F4F8FB;margin:0;padding:32px 16px;display:flex;justify-content:center}
  .card{background:#fff;max-width:520px;width:100%;border-radius:14px;box-shadow:0 4px 24px rgba(31,78,121,.12);overflow:hidden}
  .header{background:${corBg};color:#fff;padding:26px;text-align:center}
  .header .icon{font-size:44px;line-height:1;margin-bottom:6px}.header h1{margin:0;font-size:20px}
  .body{padding:26px;color:#2C3E50;font-size:15px;line-height:1.6;text-align:center}
  .detalhe{background:#F4F8FB;border-radius:8px;padding:12px;margin-top:14px;font-size:13px;color:#647883}
  .footer{background:#F4F8FB;color:#647883;padding:14px;font-size:11px;text-align:center;border-top:1px solid #DCE3E9}
</style></head><body><div class="card">
  <div class="header"><div class="icon">${icon}</div><h1>${esc(titulo)}</h1></div>
  <div class="body"><p>${mensagem}</p>${detalhe ? `<div class="detalhe">${detalhe}</div>` : ''}</div>
  <div class="footer">Sistema de Aprovacao de NF · Pronep Life Care</div>
</div></body></html>`;
}

// Pagina de botoes: card da NF + Aprovar / Rejeitar / Ver PDF.
function paginaBotoes(nf, links, aprovador) {
  const n = nf || {};
  const verPdf = n.urlPDF
    ? `<a class="link-pdf" href="${esc(n.urlPDF)}" target="_blank" rel="noopener">📎 Ver PDF da NF</a>` : '';
  return `<!DOCTYPE html><html lang="pt-BR"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>Aprovar NF</title><style>
  body{font-family:'Segoe UI',Roboto,Arial,sans-serif;background:#F4F8FB;margin:0;padding:24px 14px;display:flex;justify-content:center}
  .card{background:#fff;max-width:520px;width:100%;border-radius:14px;box-shadow:0 4px 24px rgba(31,78,121,.12);overflow:hidden}
  .header{background:#1F4E79;color:#fff;padding:22px 24px;text-align:center}
  .header .icon{font-size:38px;line-height:1;margin-bottom:4px}.header h1{margin:0;font-size:19px}
  .body{padding:22px 24px}
  .valor{font-size:30px;font-weight:700;color:#1F4E79;text-align:center;margin:2px 0 16px}
  .grid{display:grid;grid-template-columns:auto 1fr;gap:6px 12px;font-size:14px;color:#2C3E50}
  .grid .k{color:#647883}
  .btns{margin-top:22px;display:flex;flex-direction:column;gap:12px}
  .btn{display:block;text-align:center;padding:15px;border-radius:10px;text-decoration:none;font-weight:700;font-size:16px}
  .btn-ok{background:#2E7D32;color:#fff}.btn-no{background:#C62828;color:#fff}
  .link-pdf{display:block;text-align:center;margin-top:16px;color:#1F4E79;text-decoration:none;font-size:14px}
  .nota{margin-top:16px;font-size:12px;color:#647883;text-align:center;line-height:1.5}
  .footer{background:#F4F8FB;color:#647883;padding:14px;font-size:11px;text-align:center;border-top:1px solid #DCE3E9}
</style></head><body><div class="card">
  <div class="header"><div class="icon">📄</div><h1>Nova NF para sua aprovação</h1></div>
  <div class="body">
    <div class="valor">${fmtBRL(n.valor)}</div>
    <div class="grid">
      <div class="k">Fornecedor</div><div><b>${esc(n.fornecedor || '—')}</b></div>
      <div class="k">NF</div><div>${esc(n.numero || '—')}</div>
      <div class="k">Unidade / Diretoria</div><div>${esc(n.unidade || '—')} · ${esc(n.diretoria || '—')}</div>
      <div class="k">Vencimento</div><div>${fmtData(n.vencimento)}</div>
    </div>
    ${verPdf}
    <div class="btns">
      <a class="btn btn-ok" href="${esc(links.aprovar || '#')}" onclick="return confirm('Confirmar APROVAÇÃO desta NF?')">✓ Aprovar</a>
      <a class="btn btn-no" href="${esc(links.rejeitar || '#')}" onclick="return confirm('Confirmar REJEIÇÃO desta NF?')">✕ Rejeitar</a>
    </div>
    <div class="nota">Aprovando como <b>${esc(aprovador || '')}</b>. A ação é registrada e o solicitante é notificado.</div>
  </div>
  <div class="footer">Sistema de Aprovacao de NF · Pronep Life Care</div>
</div></body></html>`;
}

module.exports = async function (context, req) {
  const H = { 'Content-Type': 'text/html; charset=utf-8', 'Cache-Control': 'no-store' };
  try {
    const token = req.query && req.query.token;
    if (!token) { context.res = { status: 400, headers: H, body: paginaSimples('Link invalido', 'Esse link nao contem o token necessario.', '#C62828') }; return; }

    const secret = process.env.LINK_APROVACAO_SECRET;
    if (!secret) { context.res = { status: 500, headers: H, body: paginaSimples('Sistema mal configurado', 'LINK_APROVACAO_SECRET nao definido nas App Settings.', '#C62828') }; return; }

    let payload;
    try { payload = jwt.verify(token, secret, { algorithms: ['HS256'] }); }
    catch (e) { context.res = { status: 401, headers: H, body: paginaSimples('Link expirado ou invalido', `Esse link nao eh mais valido (${esc(e.message)}).`, '#C62828') }; return; }

    const { itemId, aprovador, action, nf } = payload;
    if (!itemId || !aprovador || action !== 'ver') { context.res = { status: 400, headers: H, body: paginaSimples('Link com dados invalidos', 'Token nao tem os campos esperados.', '#C62828') }; return; }

    // Le Status atual (uso unico) e valida o aprovador.
    const client = await getGraphClient();
    const { siteId, listId } = await resolveListaNotas(client);
    const item = await client.api(`/sites/${siteId}/lists/${listId}/items/${itemId}?expand=fields`).get();
    const fields = item.fields || {};
    const statusAtual = fields.Status || fields.field_9 || '';
    const aprovadorAtual = (fields.AprovadorAtual || fields.field_8 || '').toLowerCase();

    if (statusAtual === 'Aprovada') { context.res = { status: 200, headers: H, body: paginaSimples('NF ja foi aprovada', 'Esta nota fiscal ja foi aprovada anteriormente. Nada a fazer.', '#647883', 'Status: <b>Aprovada</b>') }; return; }
    if (statusAtual === 'Rejeitada') { context.res = { status: 200, headers: H, body: paginaSimples('NF ja foi rejeitada', 'Esta nota fiscal ja foi rejeitada anteriormente.', '#647883', 'Status: <b>Rejeitada</b>') }; return; }
    if (aprovadorAtual && aprovadorAtual !== String(aprovador).toLowerCase()) { context.res = { status: 403, headers: H, body: paginaSimples('Voce nao eh o aprovador atual', `Esta NF foi reatribuida. Aprovador atual: ${esc(aprovadorAtual)}`, '#C62828') }; return; }

    // Gera os links de acao (aprovar/rejeitar) que apontam pro AprovacaoViaLink executor.
    const links = gerarLinks(itemId, aprovador) || {};
    context.res = { status: 200, headers: H, body: paginaBotoes(nf, links, aprovador) };
  } catch (err) {
    context.log && context.log.error && context.log.error('AprovacaoPagina error:', err);
    context.res = { status: 500, headers: H, body: paginaSimples('Erro interno', esc((err && err.message) || String(err)), '#C62828') };
  }
};
