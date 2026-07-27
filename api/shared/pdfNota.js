/**
 * shared/pdfNota.js
 *
 * Resolucao do PDF de uma NF por IDENTIDADE EXATA, imune a NumeroNF duplicado.
 *
 * Contexto do bug (corrigido no RejeitarNota #19 e agora no AprovarNota): o nome do
 * arquivo segue {data}_{SEQUENCIAL}_{nome}_{uf}_{valor}.pdf. O 2o campo e um SEQUENCIAL,
 * NAO o NumeroNF. Casar o PDF por NumeroNF frouxo (indexOf('_'+numero+'_')) colide com
 * esse sequencial -> abre/move/deleta o PDF de OUTRA nota. Aqui casamos pelo nome EXATO
 * do arquivo (extraido da URL guardada na propria nota, que inclui o valor -> unico) e
 * so caimos num fallback ESTRITO numero+valor quando aceito UNICO.
 */

// Extrai a URL de um campo hyperlink (pode vir string ou { Url, Description }).
function urlDeCampo(v) {
  if (!v) return '';
  if (typeof v === 'string' && v.indexOf('http') === 0) return v;
  if (typeof v === 'object' && v.Url && String(v.Url).indexOf('http') === 0) return v.Url;
  return '';
}

// Nome exato do arquivo a partir da URL armazenada na nota (unico: inclui o valor).
function nomeArquivoDeUrl(url) {
  if (!url) return '';
  try { return decodeURIComponent(String(url).split('?')[0].split('/').pop() || ''); }
  catch (e) { return String(url).split('?')[0].split('/').pop() || ''; }
}

// Valor no formato do nome do arquivo (ex.: 15000,00).
function valorStrDe(v) {
  const n = (typeof v === 'number' ? v : Number(v)) || 0;
  return n > 0 ? n.toFixed(2).replace('.', ',') : '';
}
// Variantes do valor como aparecem nos nomes: sem milhar ("9915,94") e COM milhar
// ("9.915,94"). Nomes antigos usam o separador de milhar — sem essa variante o
// fallback estrito nao acha o arquivo e a operacao falha por seguranca (404).
function valorStrsDe(v) {
  const semMilhar = valorStrDe(v);
  if (!semMilhar) return [];
  const partes = semMilhar.split(',');
  const comMilhar = partes[0].replace(/\B(?=(\d{3})+(?!\d))/g, '.') + ',' + partes[1];
  return comMilhar === semMilhar ? [semMilhar] : [semMilhar, comMilhar];
}

/**
 * Acha o PDF alvo numa lista de arquivos do drive.
 * @param {Array} files  - itens do drive (com .name) ja filtrados por .file
 * @param {Object} opts  - { url, numero, valor }
 *   url    : URL guardada na nota (UrlPDFStr/UrlPDF ou UrlPDFAprovadoStr no estorno)
 *   numero : NumeroNF
 *   valor  : Valor (number)
 * @returns {{ target: Object|null, matchPor: string|null, ambiguo: Object|null }}
 *   matchPor: 'nome_exato' | 'numero+valor' | null
 *   Sem identificacao CONFIAVEL retorna target=null (o chamador NAO deve mover/deletar).
 */
function acharPdfAlvo(files, opts) {
  const lista = files || [];
  const out = { target: null, matchPor: null, ambiguo: null };

  // (1) FONTE DA VERDADE: nome EXATO do arquivo da propria nota.
  const nomeExato = nomeArquivoDeUrl(urlDeCampo(opts && opts.url) || (opts && opts.url) || '');
  if (nomeExato) {
    const t = lista.find(x => x.name === nomeExato);
    if (t) { out.target = t; out.matchPor = 'nome_exato'; return out; }
  }

  // (2) FALLBACK ESTRITO: exige numero E valor no nome, e so aceita se UNICO.
  // O nome do arquivo usa o numero SEM zeros a esquerda (ex.: NumeroNF "00190" -> "_190_").
  // Por isso testamos variantes: cru, limpo, sem-zeros e com-zeros (6 digitos).
  const numeroRaw = String((opts && opts.numero) || '').trim();
  const numClean = numeroRaw.replace(/[^A-Za-z0-9]/g, '');
  const numUnpadded = /^\d+$/.test(numClean) ? (numClean.replace(/^0+/, '') || '0') : numClean;
  const numPadded = /^\d+$/.test(numClean) ? numClean.padStart(6, '0') : numClean;
  const numeros = Array.from(new Set([numeroRaw, numClean, numUnpadded, numPadded].filter(Boolean)));
  const valorStrs = valorStrsDe(opts && opts.valor);
  if (numeros.length && valorStrs.length) {
    const cand = lista.filter(x => x.name
      && numeros.some(nm => x.name.startsWith(nm + '_') || x.name.includes('_' + nm + '_'))
      && valorStrs.some(vs => x.name.includes('_' + vs + '_') || x.name.includes('_' + vs + '.')));
    if (cand.length === 1) { out.target = cand[0]; out.matchPor = 'numero+valor'; }
    else out.ambiguo = { numeros: numeros, valorStrs: valorStrs, encontrados: cand.length };
  }
  return out;
}

/**
 * Caminho RELATIVO AO DRIVE a partir da URL guardada na nota.
 * Ex.: https://t.sharepoint.com/sites/X/Documentos%20Compartilhados/Notas%20Fiscais/Notas%20Aprovadas/RJ/2026-07-24/a.pdf
 *   -> "Notas Fiscais/Notas Aprovadas/RJ/2026-07-24/a.pdf"
 * Permite buscar o item por caminho direto (1 chamada), sem varrer pastas.
 */
function caminhoDriveDeUrl(url) {
  const u = urlDeCampo(url) || url || '';
  if (!u) return '';
  let p;
  try { p = decodeURIComponent(String(u).split('?')[0]); } catch (e) { p = String(u).split('?')[0]; }
  p = p.replace(/^https?:\/\/[^/]+/i, '');
  // Ancora nas pastas conhecidas do sistema (cobre tambem o legado plano).
  const marcadores = ['/Notas Fiscais/', '/Notas Aprovadas/', '/Pendentes/', '/Rejeitadas/'];
  for (const m of marcadores) {
    const i = p.indexOf(m);
    if (i >= 0) return p.substring(i + 1);
  }
  // Fallback: /sites/{site}/{biblioteca}/{resto}
  const mm = /^\/sites\/[^/]+\/[^/]+\/(.+)$/.exec(p);
  return mm ? mm[1] : '';
}

/**
 * Busca UM item do drive pelo caminho (1 chamada). Retorna null se nao existir.
 * PERF/ESTABILIDADE: usar isto no lugar de varrer pastas — varrer
 * "Notas Aprovadas/{Unidade}" faz 1 chamada por subpasta de data (centenas)
 * e estoura o timeout da Function (HTTP 500 / Backend call failure).
 */
async function buscarItemPorCaminho(client, siteId, caminho) {
  if (!caminho) return null;
  try {
    const it = await client.api('/sites/' + siteId + '/drive/root:/' + caminho).get();
    return (it && it.file) ? it : null;
  } catch (e) { return null; }
}

/**
 * Resolve o PDF de uma nota SEM varrer o drive inteiro. Ordem:
 *   1) caminho completo derivado da URL da propria nota  -> 1 chamada, identidade exata
 *   2) nome do arquivo dentro de cada pasta candidata    -> 1 chamada por pasta
 *   3) lista APENAS as pastas candidatas + acharPdfAlvo  -> 1 chamada por pasta
 * Limitado a poucas chamadas. Sem identificacao confiavel retorna target=null.
 * @param {Array<string>} pastas - caminhos de pasta candidatos (ex.: [".../RJ/2026-07-24"])
 */
async function resolverPdfNota(client, siteId, opts, pastas) {
  const out = { target: null, matchPor: null, ambiguo: null, tentativas: [] };
  const url = (opts && opts.url) || '';
  const listaPastas = (pastas || []).filter(Boolean);

  // (1) Caminho direto da URL — imune a pasta de data errada (bug BRT/UTC).
  const caminho = caminhoDriveDeUrl(url);
  if (caminho) {
    out.tentativas.push('url_direta:' + caminho);
    const it = await buscarItemPorCaminho(client, siteId, caminho);
    if (it) { out.target = it; out.matchPor = 'url_direta'; return out; }
  }

  // (2) Nome exato do arquivo dentro das pastas candidatas.
  const nome = nomeArquivoDeUrl(url);
  if (nome) {
    for (const pasta of listaPastas) {
      out.tentativas.push('nome_em_pasta:' + pasta + '/' + nome);
      const it = await buscarItemPorCaminho(client, siteId, pasta + '/' + nome);
      if (it) { out.target = it; out.matchPor = 'nome_exato'; return out; }
    }
  }

  // (3) Lista SO as pastas candidatas e aplica o match estrito.
  const arquivos = [];
  for (const pasta of listaPastas) {
    out.tentativas.push('listar:' + pasta);
    try {
      const r = await client.api('/sites/' + siteId + '/drive/root:/' + pasta + ':/children').get();
      for (const x of (r.value || [])) if (x.file) arquivos.push(x);
    } catch (e) { /* pasta inexistente/inacessivel */ }
  }
  const achado = acharPdfAlvo(arquivos, opts);
  out.target = achado.target;
  out.matchPor = achado.matchPor;
  out.ambiguo = achado.ambiguo;
  out.arquivosVistos = arquivos.length;
  out.nomesVistos = arquivos.slice(0, 8).map(function (x) { return x.name; });
  return out;
}

module.exports = {
  urlDeCampo, nomeArquivoDeUrl, valorStrDe, valorStrsDe, acharPdfAlvo,
  caminhoDriveDeUrl, buscarItemPorCaminho, resolverPdfNota
};
