/**
 * /api/DiagOmieContasPagar  (GET) — ADMIN. Descoberta, nao producao.
 *
 * Mostra COMO o Omie devolve uma conta a pagar, para o quadro "NFs a Pagar" ser
 * construido contra a realidade em vez de contra um palpite.
 *
 * Existe pelo mesmo motivo do ?diagCert=1: o que decide o desenho aqui e o
 * conjunto real de valores de `status_titulo` (o que separa "Aprovadas" de
 * "Quitadas") e quais campos vem preenchidos de fato. A documentacao do Omie
 * lista os campos possiveis, nao os que a Pronep usa.
 *
 * Query:
 *   ?unidade=RJ     RJ | SP | ES | TODAS   (default TODAS)
 *   ?dias=45        janela de vencimento: hoje-dias ate hoje+dias (default 45)
 *   ?amostras=3     quantas contas completas devolver por unidade (default 3, max 10)
 *
 * NAO grava nada e nao altera nada no Omie — so lista.
 */

require('isomorphic-fetch');
const { getCredentials, listarContasPagarPorVencimento } = require('../shared/omie');

function readClientPrincipal(req) {
  const h = req.headers && req.headers['x-ms-client-principal'];
  if (!h) return null;
  try { return JSON.parse(Buffer.from(h, 'base64').toString('utf-8')); } catch (e) { return null; }
}

async function isAdmin(req) {
  const p = readClientPrincipal(req);
  const roles = (p && p.userRoles) || [];
  if (roles.includes('administrador') || roles.includes('admin')) return true;
  try {
    const { getUser } = require('../shared/auth');
    const user = await getUser(req);
    if (!user) return false;
    const { isAdminEmail } = require('../shared/authz');
    if (isAdminEmail((user.email || '').toLowerCase())) return true;
    const { getUserRoles } = require('../shared/userRoles');
    return ((await getUserRoles(user)) || []).includes('administrador');
  } catch (e) { return false; }
}

module.exports = async function (context, req) {
  const t0 = Date.now();
  const q = req.query || {};
  const dias = Math.max(1, Math.min(180, parseInt(q.dias, 10) || 45));
  const amostras = Math.max(1, Math.min(10, parseInt(q.amostras, 10) || 3));
  const pedido = String(q.unidade || 'TODAS').toUpperCase();
  const unidades = pedido === 'TODAS' ? ['RJ', 'SP', 'ES'] : [pedido];

  const diag = { step: 'init', dias: dias, unidades: [], timeMs: 0 };

  try {
    if (!(await isAdmin(req))) {
      context.res = { status: 403, headers: { 'Content-Type': 'application/json' },
        body: { error: 'Apenas admin' } };
      return;
    }

    const hoje = new Date();
    const de = new Date(hoje.getTime() - dias * 86400000);
    const ate = new Date(hoje.getTime() + dias * 86400000);
    diag.janela = { de: de.toISOString().slice(0, 10), ate: ate.toISOString().slice(0, 10) };

    for (const u of unidades) {
      const bloco = { unidade: u };
      diag.unidades.push(bloco);

      let creds;
      try { creds = getCredentials(u); }
      catch (e) { bloco.erro = e.message; continue; }
      bloco.empresa = creds.empresa;

      let r;
      try {
        /* 6 paginas = ate 300 contas. Suficiente para descobrir o formato sem
           gastar a cota do Omie (~60 req/min por app_key). */
        r = await listarContasPagarPorVencimento({ de, ate, maxPaginas: 6 }, creds);
      } catch (e) { bloco.erro = e.message; continue; }

      bloco.totalRegistrosNaJanela = r.totalRegistros;
      bloco.lidos = r.contas.length;
      bloco.paginas = r.paginas;
      bloco.truncado = r.truncado;

      if (!r.contas.length) { bloco.aviso = 'Nenhuma conta a pagar nesta janela.'; continue; }

      /* 1) Quais campos existem de fato. */
      bloco.camposDisponiveis = Object.keys(r.contas[0]).sort();

      /* 2) Quais campos vem PREENCHIDOS em pelo menos uma conta. Campo que existe
            mas vem sempre vazio nao serve para montar coluna nenhuma. */
      const preenchidos = {};
      for (const c of r.contas) {
        for (const k of Object.keys(c)) {
          const v = c[k];
          if (v !== null && v !== undefined && v !== '' && v !== 0) {
            preenchidos[k] = (preenchidos[k] || 0) + 1;
          }
        }
      }
      bloco.camposPreenchidos = Object.keys(preenchidos).sort().map(function (k) {
        return k + ': ' + preenchidos[k] + '/' + r.contas.length;
      });

      /* 3) O DADO MAIS IMPORTANTE: os valores reais de status_titulo. E ele que
            vai separar "Aprovadas" de "Quitadas" no quadro. */
      const status = {};
      const semNF = { com: 0, sem: 0 };
      for (const c of r.contas) {
        const st = String(c.status_titulo || '(vazio)');
        status[st] = (status[st] || 0) + 1;
        const nf = c.numero_documento_fiscal || c.nota_fiscal || '';
        if (String(nf).trim()) semNF.com++; else semNF.sem++;
      }
      bloco.statusTitulo = status;
      /* Conta sem numero de NF nao casa com documento fiscal nenhum — se for
         maioria, o merge por numero nao serve e o desenho muda. */
      bloco.temNumeroNF = semNF;

      /* 4) A PERGUNTA QUE DECIDE DUAS COISAS AO MESMO TEMPO.
            Cruza "tem chave de NF-e" com id_origem (como a conta ENTROU no Omie).
            - a coluna "com chave" mede o alcance de buscar XML na SEFAZ: chave e
              o unico jeito de consultar, e sem ela nao ha o que pedir
            - a coluna "sem chave" diz se existe integracao a reaproveitar. Se as
              NFS-e entraram por API, alguem ja integrou com as prefeituras e vale
              investigar; se entraram manualmente, ninguem integrou — alguem
              digitou, e nao ha o que herdar. */
      const cruzamento = {};
      for (const c of r.contas) {
        const temCh = String(c.chave_nfe || '').replace(/\D/g, '').length === 44;
        const org = String(c.id_origem || '(vazio)');
        const k = org + ' | ' + (temCh ? 'com chave (NF-e)' : 'sem chave (NFS-e/outro)');
        cruzamento[k] = (cruzamento[k] || 0) + 1;
      }
      bloco.origemVsChave = cruzamento;

      /* So as EM ABERTO — que sao as que viram card. Misturar as pagas inflaria o
         numero com historico que nao vai para o quadro. */
      const abertas = r.contas.filter(function (c) {
        return String(c.status_titulo || '').toUpperCase().indexOf('PAGO') < 0;
      });
      const abertasComChave = abertas.filter(function (c) {
        return String(c.chave_nfe || '').replace(/\D/g, '').length === 44;
      }).length;
      bloco.emAberto = {
        total: abertas.length,
        comChave: abertasComChave,
        semChave: abertas.length - abertasComChave,
        percentualComChave: abertas.length
          ? Math.round((abertasComChave / abertas.length) * 100) + '%' : '—'
      };

      /* 5) Amostras completas para eu ler os formatos (data, valor, chave). */
      bloco.amostras = r.contas.slice(0, amostras);
    }

    diag.step = 'done';
    diag.timeMs = Date.now() - t0;
    context.res = { status: 200, headers: { 'Content-Type': 'application/json' },
      body: Object.assign({ ok: true }, diag) };
  } catch (err) {
    diag.timeMs = Date.now() - t0;
    context.res = { status: 500, headers: { 'Content-Type': 'application/json' },
      body: Object.assign({ error: (err && err.message) || String(err) }, diag) };
  }
};
