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

// Valor no formato do nome do arquivo (ex.: 1.234,56).
function valorStrDe(v) {
  const n = (typeof v === 'number' ? v : Number(v)) || 0;
  return n > 0 ? n.toFixed(2).replace('.', ',') : '';
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
  const numero = String((opts && opts.numero) || '').trim();
  const valorStr = valorStrDe(opts && opts.valor);
  if (numero && valorStr) {
    const cand = lista.filter(x => x.name
      && (x.name.startsWith(numero + '_') || x.name.includes('_' + numero + '_'))
      && (x.name.includes('_' + valorStr + '_') || x.name.includes('_' + valorStr + '.')));
    if (cand.length === 1) { out.target = cand[0]; out.matchPor = 'numero+valor'; }
    else out.ambiguo = { numero, valorStr, encontrados: cand.length };
  }
  return out;
}

module.exports = { urlDeCampo, nomeArquivoDeUrl, valorStrDe, acharPdfAlvo };
