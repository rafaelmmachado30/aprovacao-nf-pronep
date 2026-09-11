/**
 * scripts/verificar-abrirpdf-fuso.js
 *
 * Regressao do bug "PDF nao encontrado" em Notas Aprovadas (11/09/2026).
 *
 * O DEFEITO: a pasta do PDF aprovado e nomeada pela data BRT (AprovarNota faz `agora - 3h`),
 * mas `AprovadoEm` e gravado em UTC. AbrirPdfDaNota tomava `substring(0,10)` desse campo, ou
 * seja a data UTC. A partir das 21h BRT as duas divergem e a busca ia para a pasta do dia
 * seguinte — que EXISTE e tem arquivos, entao o fallback de "pasta nao encontrada" nunca
 * disparava e a tela dizia "PDF nao encontrado" com o arquivo intacto no SharePoint.
 *
 * Caso real: NF 6, RJ, R$ 900,00, aprovada 08/09/2026 21:25 BRT (= 09/09 00:25 UTC).
 * Arquivo em .../RJ/2026-09-08/, busca feita em .../RJ/2026-09-09/.
 *
 * Roda com o Graph SIMULADO — nao toca SharePoint, nao precisa de credencial:
 *     node scripts/verificar-abrirpdf-fuso.js
 *
 * Os nomes de arquivo abaixo sao REAIS, copiados das telas do incidente.
 */
'use strict';

const path = require('path');
const Module = require('module');
const RAIZ = path.join(__dirname, '..', 'api');

/* isomorphic-fetch nao e usado no caminho testado e pode nao estar instalado localmente. */
const origResolve = Module._resolveFilename;
Module._resolveFilename = function (req, ...rest) {
  if (req === 'isomorphic-fetch') return require.resolve('path');
  return origResolve.call(this, req, ...rest);
};

function stub(rel, exports) {
  const p = require.resolve(path.join(RAIZ, rel));
  require.cache[p] = { id: p, filename: p, loaded: true, exports };
}

stub('shared/authz.js', {
  resolveAuthz: async () => ({
    isAdmin: true, isFinanceiro: true, isGestor: false,
    email: 'teste@pronep.com.br', roles: ['admin']
  })
});
stub('shared/acessoTelas.js', { lerMapaTelas: async () => ({}), telasLiberadasPara: () => [] });

/* ------------------------- SharePoint simulado ------------------------- */
const ARVORE = {
  'Notas Fiscais/Notas Aprovadas/RJ/2026-09-08': [
    '2026-09-15_409_CITTA-TELECOM-LTDA_RJ_1134,52_APROVADA_2026-09-08.pdf',
    '2026-09-25_1469567_CONTROLID-INDUSTRIA-COMERCIO-D_RJ_3222,20_APROVADA_2026-09-08.pdf',
    '2026-09-15_6_64-752-349-PIETRO-VINICIUS-DA_RJ_900,00_APROVADA_2026-09-08.pdf',
    '2026-09-17_411_TELEFONICA-BRASIL-S-A_RJ_437,63_APROVADA_2026-09-08.pdf',
    '2026-09-15_148_BMS-SERVICE-LTDA_RJ_3314,37_APROVADA_2026-09-08.pdf'
  ],
  'Notas Fiscais/Notas Aprovadas/RJ/2026-09-09': [
    '2026-09-10_6776_COMERCIAL-YAZBEK-LTDA_RJ_10950,00_APROVADA_2026-09-09.pdf',
    '2026-09-21_1134_EXPRESS-REMOCOES-LTDA_RJ_3760,00_APROVADA_2026-09-09.pdf',
    '2026-09-23_7116_IRON-MOUNTAIN-DO-BRASIL-LTDA_RJ_13,62_APROVADA_2026-09-09.pdf',
    '2026-09-25_262_LOA-SERVICOS-MEDICOS-LTDA_RJ_4250,00_APROVADA_2026-09-09.pdf'
  ],
  /* Pasta que nenhuma regra de fuso preveria: so a varredura acha. */
  'Notas Fiscais/Notas Aprovadas/RJ/2026-09-01': [
    '2026-09-20_999_FORNECEDOR-MOVIDO-A-MAO_RJ_1500,00_APROVADA_2026-09-01.pdf'
  ]
};
const SUBPASTAS = Object.keys(ARVORE).map((p, i) => ({ id: 'F' + i, folder: {}, _path: p }));

let ITEM_FIELDS = {};

function arquivosDe(pasta) {
  return ARVORE[pasta].map((nome, i) => ({
    id: pasta + '#' + i, name: nome, file: {},
    webUrl: 'https://sp/' + encodeURIComponent(pasta) + '/' + encodeURIComponent(nome)
  }));
}

function fakeApi(p) {
  const api = {
    filter: () => api,
    header: () => api,
    get: async () => {
      if (p.startsWith('/sites/') && p.indexOf('/drive/') < 0 &&
          p.indexOf('/lists') < 0 && p.indexOf('/items/') < 0) return { id: 'SITE' };
      if (p === '/sites/SITE/lists') return { value: [{ id: 'LIST' }] };
      if (p === '/sites/SITE/lists/LIST/columns') {
        return { value: Object.keys(ITEM_FIELDS).map(k => ({ name: k, displayName: k })) };
      }
      if (p.indexOf('/items/') >= 0 && p.indexOf('expand=fields') >= 0) return { fields: ITEM_FIELDS };

      const mPasta = /^\/sites\/SITE\/drive\/root:\/(.+):\/children$/.exec(p);
      if (mPasta) {
        const pasta = decodeURIComponent(mPasta[1]);
        if (pasta === 'Notas Fiscais/Notas Aprovadas/RJ') return { value: SUBPASTAS };
        if (ARVORE[pasta]) return { value: arquivosDe(pasta) };
        const err = new Error('itemNotFound: ' + pasta); err.statusCode = 404; throw err;
      }
      const mSub = /^\/sites\/SITE\/drive\/items\/(F\d+)\/children$/.exec(p);
      if (mSub) return { value: arquivosDe(SUBPASTAS.find(s => s.id === mSub[1])._path) };

      throw new Error('rota nao simulada: ' + p);
    }
  };
  return api;
}
stub('shared/graph.js', { getGraphClient: async () => ({ api: fakeApi }) });

process.env.SHAREPOINT_SITE_HOSTNAME = process.env.SHAREPOINT_SITE_HOSTNAME || 'exemplo.sharepoint.com';
process.env.SHAREPOINT_SITE_PATH = process.env.SHAREPOINT_SITE_PATH || '/sites/NF';

const abrirPdf = require(path.join(RAIZ, 'AbrirPdfDaNota/index.js'));

const BASE = {
  Status: 'Aprovada', Unidade: 'RJ', Diretoria: 'Tecnologia',
  LancadoPor: 'teste@pronep.com.br'
};

async function caso(nome, campos, esperado) {
  ITEM_FIELDS = Object.assign({}, BASE, campos);
  const ctx = {};
  await abrirPdf(ctx, { query: { id: '1' }, headers: {} });
  const st = ctx.res.status;
  const loc = (ctx.res.headers || {}).Location || '';
  const arquivo = loc ? decodeURIComponent(loc.split('/').pop()) : '(nenhum)';
  const ok = esperado === 404 ? st === 404 : (st === 302 && arquivo === esperado);
  console.log((ok ? '  OK   ' : '  FALHA') + '  ' + nome);
  console.log('         status ' + st + ' -> ' + arquivo);
  if (!ok) console.log('         esperado: ' + esperado);
  return ok;
}

(async () => {
  console.log('\nAbrirPdfDaNota — regressao de fuso na pasta de Notas Aprovadas\n');
  let tudo = true;

  tudo = await caso(
    'CASO DO INCIDENTE: aprovada 08/09 21:25 BRT (= 09/09 00:25 UTC)',
    { NumeroNF: '6', Valor: 900, AprovadoEm: '2026-09-09T00:25:00Z' },
    '2026-09-15_6_64-752-349-PIETRO-VINICIUS-DA_RJ_900,00_APROVADA_2026-09-08.pdf') && tudo;

  tudo = await caso(
    'aprovada de dia (14h BRT) — BRT e UTC caem no mesmo dia',
    { NumeroNF: '6776', Valor: 10950, AprovadoEm: '2026-09-09T17:00:00Z' },
    '2026-09-10_6776_COMERCIAL-YAZBEK-LTDA_RJ_10950,00_APROVADA_2026-09-09.pdf') && tudo;

  tudo = await caso(
    'legado: arquivo na pasta da data UTC (segunda candidata)',
    { NumeroNF: '262', Valor: 4250, AprovadoEm: '2026-09-09T02:00:00Z' },
    '2026-09-25_262_LOA-SERVICOS-MEDICOS-LTDA_RJ_4250,00_APROVADA_2026-09-09.pdf') && tudo;

  tudo = await caso(
    'pasta imprevisivel (movido a mao) — achado pela varredura',
    { NumeroNF: '999', Valor: 1500, AprovadoEm: '2026-09-20T12:00:00Z' },
    '2026-09-20_999_FORNECEDOR-MOVIDO-A-MAO_RJ_1500,00_APROVADA_2026-09-01.pdf') && tudo;

  /* Este caso vale tanto quanto os outros: a varredura NAO pode afrouxar o criterio e
     comecar a abrir PDF de outra NF. Ela so amplia o conjunto onde o mesmo criterio
     estrito (numero + valor unico) e aplicado. */
  tudo = await caso(
    'NF inexistente — continua 404, sem abrir PDF de outra nota',
    { NumeroNF: '31337', Valor: 42, AprovadoEm: '2026-09-09T00:25:00Z' },
    404) && tudo;

  console.log('\n' + (tudo ? '  >>> todos os casos passam\n' : '  >>> HA FALHA\n'));
  process.exit(tudo ? 0 : 1);
})();
