/**
 * /api/AbrirPdfDaNota?id=<spListItemId>
 *
 * Acha o PDF da NF na pasta correta do SharePoint (Pendentes, Notas Aprovadas
 * ou Rejeitadas, baseado no Status) e faz redirect 302 pra webUrl do arquivo.
 *
 * Uso: link direto em modais do front. Abre em nova aba.
 */

require('isomorphic-fetch');
const { ClientSecretCredential } = require('@azure/identity');
const { Client } = require('@microsoft/microsoft-graph-client');
const { TokenCredentialAuthenticationProvider } =
  require('@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials');

const LIST_NOTAS = 'PRONEP-NF-NotasFiscais';
const cache = { siteId: null, listNotasId: null, invColMap: null };

async function getGraphClient() {
  const tenantId = process.env.AAD_TENANT_ID;
  const clientId = process.env.AAD_CLIENT_ID;
  const clientSecret = process.env.AAD_CLIENT_SECRET;
  if (!tenantId || !clientId || !clientSecret) throw new Error('AAD_* incompletas');
  const credential = new ClientSecretCredential(tenantId, clientId, clientSecret);
  const authProvider = new TokenCredentialAuthenticationProvider(credential, {
    scopes: ['https://graph.microsoft.com/.default']
  });
  return Client.initWithMiddleware({ authProvider });
}

async function resolveSiteAndList(client) {
  if (cache.siteId && cache.listNotasId) return cache;
  const host = process.env.SHAREPOINT_SITE_HOSTNAME;
  const path = process.env.SHAREPOINT_SITE_PATH;
  if (!host || !path) throw new Error('SHAREPOINT_* incompletas');
  const siteResp = await client.api(`/sites/${host}:${path}`).get();
  cache.siteId = siteResp.id;
  const lists = await client.api(`/sites/${cache.siteId}/lists`).filter(`displayName eq '${LIST_NOTAS}'`).get();
  cache.listNotasId = lists.value[0].id;
  // Tambem carrega o invColMap pra ler fields normalizados
  const cols = await client.api(`/sites/${cache.siteId}/lists/${cache.listNotasId}/columns`).get();
  cache.invColMap = {};
  for (const col of (cols.value || [])) {
    if (col.displayName && col.name) cache.invColMap[col.name] = col.displayName;
  }
  return cache;
}

function htmlErro(titulo, msg, extra) {
  return `<!DOCTYPE html><html><head><meta charset="UTF-8"><title>${titulo}</title>
<style>body{font-family:'Segoe UI',sans-serif;background:#F4F8FB;margin:0;padding:60px 16px;text-align:center;color:#2C3E50}
.card{max-width:520px;margin:0 auto;background:#fff;padding:32px;border-radius:12px;box-shadow:0 4px 16px rgba(31,78,121,.1)}
h1{color:#C62828;margin:0 0 12px}.extra{background:#F4F8FB;padding:12px;border-radius:6px;margin-top:14px;font-size:13px;color:#647883}</style>
</head><body><div class="card"><h1>${titulo}</h1><p>${msg}</p>${extra ? `<div class="extra">${extra}</div>` : ''}</div></body></html>`;
}

module.exports = async function (context, req) {
  try {
    const itemId = req.query.id;
    if (!itemId) {
      context.res = { status: 400, headers: { 'Content-Type': 'text/html; charset=utf-8' },
        body: htmlErro('Parametro faltando', 'O parametro id eh obrigatorio. Uso: ?id=&lt;sp-item-id&gt;') };
      return;
    }

    const client = await getGraphClient();
    const { siteId, listNotasId, invColMap } = await resolveSiteAndList(client);

    // Le item
    const item = await client.api(`/sites/${siteId}/lists/${listNotasId}/items/${itemId}?expand=fields`).get();
    const raw = item.fields || {};
    // Normaliza
    const fields = {};
    for (const [k, v] of Object.entries(raw)) {
      if (invColMap[k]) fields[invColMap[k]] = v;
    }
    const status = fields.Status || '';
    const unidade = fields.Unidade || '';
    const diretoria = fields.Diretoria || '';
    const numero = String(fields.NumeroNF || '');

    if (!unidade || !diretoria) {
      context.res = { status: 400, headers: { 'Content-Type': 'text/html; charset=utf-8' },
        body: htmlErro('Dados insuficientes', 'Esta NF nao tem Unidade ou Diretoria definidas.', 'Procure manualmente no SharePoint.') };
      return;
    }

    // Determina pasta baseado no Status
    const aprovadoEm = fields.AprovadoEm || '';
    const dataAprovada = aprovadoEm ? String(aprovadoEm).substring(0, 10) : '';
    let folder;
    if (status === 'Aprovada' && dataAprovada) {
      // Vai DIRETO na pasta da data de aprovacao (otimizacao - evita listar todas as subpastas)
      folder = `Notas Fiscais/Notas Aprovadas/${unidade}/${dataAprovada}`;
    } else if (status === 'Aprovada') {
      // Sem data de aprovacao - fallback: lista todas as subpastas (mais lento)
      folder = `Notas Fiscais/Notas Aprovadas/${unidade}`;
    } else if (status === 'Rejeitada') {
      folder = `Notas Fiscais/Rejeitadas`;
    } else {
      folder = `Notas Fiscais/Pendentes/${unidade}/Diretoria ${diretoria}`;
    }

    // Lista arquivos da pasta
    let arquivos = [];
    try {
      if (status === 'Aprovada' && !dataAprovada) {
        // Fallback: itera subpastas de data (lento, evitamos quando possivel)
        const subRespUnidade = await client.api(`/sites/${siteId}/drive/root:/${folder}:/children`).get();
        for (const subfolder of (subRespUnidade.value || []).filter(x => x.folder)) {
          try {
            const filesResp = await client.api(`/sites/${siteId}/drive/items/${subfolder.id}/children`).get();
            for (const f of (filesResp.value || [])) { if (f.file) arquivos.push(f); }
          } catch (e) { /* ignora pastas vazias */ }
        }
      } else {
        // Caminho rapido: direto na pasta especifica
        const resp = await client.api(`/sites/${siteId}/drive/root:/${folder}:/children`).get();
        arquivos = (resp.value || []).filter(x => x.file);
      }
    } catch (e) {
      // Pasta nao encontrada - se Aprovada com data, tenta sem data (NF pode ter sido aprovada em outra data)
      if (status === 'Aprovada' && dataAprovada) {
        const fallbackFolder = `Notas Fiscais/Notas Aprovadas/${unidade}`;
        try {
          const subRespUnidade = await client.api(`/sites/${siteId}/drive/root:/${fallbackFolder}:/children`).get();
          for (const subfolder of (subRespUnidade.value || []).filter(x => x.folder)) {
            try {
              const filesResp = await client.api(`/sites/${siteId}/drive/items/${subfolder.id}/children`).get();
              for (const f of (filesResp.value || [])) { if (f.file) arquivos.push(f); }
            } catch (e2) {}
          }
          folder = fallbackFolder;
        } catch (e2) {
          context.res = { status: 404, headers: { 'Content-Type': 'text/html; charset=utf-8' },
            body: htmlErro('Pasta nao encontrada', `Nao encontrei a pasta: <span style="font-family:monospace">${folder}</span>`, (e.message || '') + ' | fallback: ' + (e2.message || '')) };
          return;
        }
      } else {
        context.res = { status: 404, headers: { 'Content-Type': 'text/html; charset=utf-8' },
          body: htmlErro('Pasta nao encontrada', `Nao encontrei a pasta: <span style="font-family:monospace">${folder}</span>`, e.message) };
        return;
      }
    }

    // Acha o arquivo pelo NumeroNF como prefixo (formato PostNota: {numero}_{razao}_{timestamp}_*.pdf)
    // SEGURANCA: SEM fallback "mais recente" - retorna erro se nao achar match exato pelo NumeroNF.
    let target = null;
    if (numero) {
      // 1. Match exato: comeca com "{numero}_"
      target = arquivos.find(a => a.name && a.name.startsWith(numero + '_'));
      // 2. Reforco: contem "_{numero}_" em qualquer parte (caso o numero seja parte do nome)
      if (!target) {
        target = arquivos.find(a => a.name && a.name.indexOf('_' + numero + '_') >= 0);
      }
      // 3. Match no fileSize+hash poderia ser feito aqui se persistido — TODO
    }
    if (!target) {
      const filenames = arquivos.map(a => a.name).slice(0, 10);
      context.res = { status: 404, headers: { 'Content-Type': 'text/html; charset=utf-8' },
        body: htmlErro(
          'PDF nao encontrado',
          `Nao encontrei o PDF da NF <b>${numero}</b> nesta pasta.<br>` +
          `Pode ser que o arquivo foi renomeado ou apagado fora do sistema. ` +
          `Procure manualmente no SharePoint.`,
          `Pasta procurada: <span style="font-family:monospace">${folder}</span><br>` +
          `Arquivos disponiveis (${arquivos.length}): ${filenames.length ? filenames.join('<br>') : '(nenhum)'}`
        )
      };
      return;
    }

    // Faz redirect pro webUrl do arquivo (SP abre o PDF no preview do Office Online)
    context.res = {
      status: 302,
      headers: { 'Location': target.webUrl }
    };
  } catch (err) {
    context.log && context.log.error && context.log.error('AbrirPdfDaNota:', err);
    context.res = {
      status: 500,
      headers: { 'Content-Type': 'text/html; charset=utf-8' },
      body: htmlErro('Erro ao localizar PDF', (err && err.message) || String(err), '')
    };
  }
};
