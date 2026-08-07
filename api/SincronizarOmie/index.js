/**
 * /api/SincronizarOmie  (GET|POST) — cron/admin. Alimenta o quadro "NFs a Pagar".
 *
 * Traz as contas a pagar EM ABERTO do Omie e as guarda no SharePoint, para a tela
 * ler de um lugar so e abrir instantaneamente.
 *
 * POR QUE O OMIE E NAO A SEFAZ: o Omie ja consome o DFe dos mesmos CNPJs e a SEFAZ
 * conta as consultas POR CNPJ — com os dois puxando veio cStat 656 nas tres
 * filiais. Alem disso o Omie entrega o que a SEFAZ nao tem: data de vencimento
 * tratada e o status de pagamento do financeiro.
 *
 * Query:
 *   ?unidade=RJ     RJ | SP | ES | TODAS  (default TODAS)
 *   ?dias=10        alem das em aberto, traz o que foi ALTERADO nos ultimos N dias
 *   ?dryRun=1       consulta o Omie mas nao grava no SharePoint
 *   ?apenasFornecedores=1  so resolve os CNPJs pendentes (nao le contas)
 *
 * DUAS CONSULTAS, POR MOTIVOS DIFERENTES:
 *   1. filtrar_por_status=EMABERTO  -> tudo que esta em aberto agora (RJ: ~548)
 *   2. filtrar_apenas_alteracao     -> o que MUDOU, incluindo o que virou PAGO
 * Sem a segunda, uma conta paga simplesmente sumiria do resultado 1 e o card
 * ficaria eternamente parado em "Aprovadas" — nunca chegaria em "Quitadas".
 *
 * REGRA QUE NAO PODE SER RELAXADA: filtrar para estreitar, CLASSIFICAR pelo
 * status_titulo do proprio registro. O filtro EMABERTO/ATRASADO do Omie devolve
 * tambem contas pagas com atraso (medido: 10 ATRASADO + 10 PAGO numa amostra de
 * 20). Confiar no filtro para decidir "esta em aberto" faria conta ja paga voltar
 * a aparecer como pendente — e alguem pagaria duas vezes.
 */

require('isomorphic-fetch');
const { getGraphClient, resolveSiteId } = require('../shared/graph');
const { getCredentials, listarContasPagarPorVencimento,
        resolverFornecedoresPorCodigo } = require('../shared/omie');
const { indexarPorCodigoOmie, gravarEmLote, prepararContaOmie, resolveListId,
        LIST_DOCFIS, lerConfigSefaz, lerCorteVencimento, soDigitos } =
        require('../shared/documentosFiscais');

const ORCAMENTO_MS = 38000;   // margem para fechar antes dos ~45s da plataforma

/* PRIORIDADE: GRAVAR. Resolver CNPJ de fornecedor e leitura no Omie e disputava o
   mesmo orcamento da escrita no SharePoint — sem prazo proprio, 40 consultas
   comeram tudo e uma execucao leu 875 contas para gravar ZERO. Escrita que nao
   acontece nao volta sozinha; CNPJ pendente volta na proxima e ainda tem modo
   dedicado (?apenasFornecedores=1). Por isso os fornecedores param aos 20s e os
   18s restantes ficam reservados para a gravacao. */
const PRAZO_FORNECEDORES_MS = 20000;

function readClientPrincipal(req) {
  const h = req.headers && req.headers['x-ms-client-principal'];
  if (!h) return null;
  try { return JSON.parse(Buffer.from(h, 'base64').toString('utf-8')); } catch (e) { return null; }
}

async function autorizado(req) {
  const segredo = process.env.SEFAZ_SECRET;
  const enviado = (req.headers &&
    (req.headers['x-automacao-secret'] || req.headers['X-Automacao-Secret'])) || '';
  if (segredo && enviado && enviado === segredo) return true;
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

/* "20/03/2026" -> "2026-03-20". O Omie usa DD/MM/AAAA; o SharePoint quer ISO. */
function dataOmieParaISO(d) {
  const m = /^(\d{2})\/(\d{2})\/(\d{4})$/.exec(String(d || '').trim());
  return m ? (m[3] + '-' + m[2] + '-' + m[1]) : null;
}

/* CANCELADO nao vira card: e uma conta que deixou de existir para o financeiro.
   Mantida na lista (o historico importa) porem marcada, e a leitura a esconde. */
function ehCancelada(status) {
  return String(status || '').toUpperCase().indexOf('CANCEL') >= 0;
}

module.exports = async function (context, req) {
  const t0 = Date.now();
  const q = req.query || {};
  const dryRun = q.dryRun === '1' || q.dryRun === 'true';
  /* 3 dias por padrao, nao 10: o cron roda 3x ao dia, entao 3 dias ja cobrem
     com folga qualquer execucao perdida. Cada dia extra e uma pagina a mais de
     leitura, e a leitura ja custa ~20s por unidade. */
  const dias = Math.max(1, Math.min(90, parseInt(q.dias, 10) || 3));
  /* So resolve CNPJ de fornecedor, sem ler contas do Omie. Ver o bloco no laco. */
  const apenasFornecedores = q.apenasFornecedores === '1' || q.apenasFornecedores === 'true';
  const pedido = String(q.unidade || 'TODAS').toUpperCase();
  /* Retomada da leitura. Uma unidade com mais de 1.000 contas em aberto (SP tem
     2.053) nao cabe numa execucao; sem isto, cada rodada releria a pagina 1 e a
     segunda metade nunca entraria. Faixas sobrepostas sao seguras — a gravacao
     casa por codigo_lancamento_omie e atualiza em vez de duplicar. */
  const paginaInicial = Math.max(1, parseInt(q.paginaInicial, 10) || 1);
  const unidades = pedido === 'TODAS' ? ['RJ', 'SP', 'ES'] : [pedido];

  const diag = { step: 'init', dryRun: dryRun, dias: dias, unidades: [], avisos: [], timeMs: 0 };

  try {
    if (!(await autorizado(req))) {
      context.res = { status: 403, headers: { 'Content-Type': 'application/json' },
        body: { error: 'Nao autorizado' } };
      return;
    }

    diag.step = 'graph';
    const client = await getGraphClient();
    const siteId = await resolveSiteId(client);

    /* Mesmo freio de mao do quadro: se a integracao esta desligada em
       Configuracoes, nao consulta nada. dryRun ignora, por ser diagnostico. */
    const cfg = await lerConfigSefaz(client, siteId);
    diag.integracao = cfg;
    if (!cfg.habilitado && !dryRun) {
      diag.timeMs = Date.now() - t0;
      context.res = { status: 200, headers: { 'Content-Type': 'application/json' },
        body: Object.assign({ ok: true, desligada: true,
          mensagem: 'Integracao DESLIGADA em Configuracoes' +
                    (cfg.motivo ? ' — ' + cfg.motivo : '') + '. Nada foi consultado.' }, diag) };
      return;
    }

    const hoje = new Date();
    const desde = new Date(hoje.getTime() - dias * 86400000);

    /* Contas vencidas antes do corte nao entram: o sistema nao deve nascer com
       anos de historico vencido do Omie. A data e FIXA (ver lerCorteVencimento) —
       movel faria as contas de agosto sumirem em setembro. */
    const corte = await lerCorteVencimento(client, siteId);
    diag.corteVencimento = corte;

    for (const u of unidades) {
      const bloco = {
        unidade: u, emAberto: 0, alteradas: 0, contas: 0,
        planejadosNovos: 0, planejadosAtualizados: 0, gravados: 0,
        semMudanca: 0, canceladas: 0,
        semFornecedor: 0, erro: null, truncado: false
      };
      diag.unidades.push(bloco);

      if (Date.now() - t0 > ORCAMENTO_MS) {
        bloco.erro = 'sem tempo nesta execucao — rode com ?unidade=' + u;
        continue;
      }

      let creds;
      try { creds = getCredentials(u); }
      catch (e) { bloco.erro = e.message; continue; }
      bloco.empresa = creds.empresa;

      /* MODO SO-FORNECEDORES. Resolver CNPJ competia por tempo com a leitura das
         contas, que sozinha come 20s: sobravam ~12 consultas por execucao, e 58
         pendentes levariam 5 dias. Sem ler contas, cabem ~50 por execucao e o
         cadastro fecha numa ou duas rodadas.
         Nao e atalho: e reconhecer que sao dois trabalhos com custos diferentes
         disputando o mesmo orcamento. */
      if (apenasFornecedores) {
        let idx;
        try { idx = await indexarPorCodigoOmie(client, siteId, u); }
        catch (e) { bloco.erro = 'indice: ' + e.message; continue; }

        const pendentes = [];
        const linhasPorCodigo = {};
        for (const k of Object.keys(idx)) {
          const f = idx[k].fields || {};
          const cod = String(f.CodigoClienteOmie || '');
          if (!cod) continue;
          if (soDigitos(f.EmitenteCNPJ).length === 14) continue;   /* ja resolvido */
          if (!linhasPorCodigo[cod]) { linhasPorCodigo[cod] = []; pendentes.push(cod); }
          linhasPorCodigo[cod].push(idx[k]);
        }
        bloco.fornecedoresPendentes = pendentes.length;
        if (!pendentes.length) { bloco.mensagem = 'Nada pendente.'; continue; }

        /* Mesmo prazo por relogio do outro caminho: reserva 8s para a escrita,
           senao as consultas comem tudo e as linhas resolvidas nao chegam a ser
           gravadas — o trabalho seria refeito do zero na rodada seguinte. */
        const prazoForn = t0 + ORCAMENTO_MS - 8000;
        const teto = Math.max(0, Math.min(80, Math.floor((prazoForn - Date.now()) / 600)));
        const rf = await resolverFornecedoresPorCodigo(pendentes, creds, teto, prazoForn);
        bloco.fornecedoresConsultados = rf.consultados;

        /* Uma linha pode repetir o mesmo fornecedor (parcelas): atualiza TODAS. */
        const listaId = await resolveListId(client, siteId, LIST_DOCFIS);
        const ops = [];
        for (const cod of Object.keys(rf.mapa)) {
          const info = rf.mapa[cod];
          if (!info.cnpj) continue;
          for (const linha of (linhasPorCodigo[cod] || [])) {
            ops.push({ tipo: 'patch', itemId: linha.id,
              fields: { EmitenteCNPJ: info.cnpj, EmitenteNome: info.razao || '' } });
          }
        }
        const rr = await gravarEmLote(client, siteId, listaId, ops, t0 + ORCAMENTO_MS);
        bloco.linhasAtualizadas = rr.ok;
        bloco.restantes = rr.restantes;

        /* PENDENTE = SEM CNPJ, nao "sem consulta". A conta antiga era
           pendentes - consultados, e o cache em memoria resolve fornecedor SEM
           consultar: uma execucao preencheu 7 linhas com 0 consultas e mesmo assim
           anunciou "ainda faltam 2". Contar o trabalho pela ferramenta usada, e nao
           pelo resultado, faz o relatorio divergir da realidade — de novo. */
        const resolvidos = Object.keys(rf.mapa).filter(function (c) {
          return rf.mapa[c] && rf.mapa[c].cnpj;
        }).length;
        bloco.fornecedoresResolvidos = resolvidos;
        bloco.fornecedoresAindaPendentes = Math.max(0, pendentes.length - resolvidos);

        /* Codigo consultado que voltou SEM CNPJ nao resolve nunca — e fornecedor
           pessoa fisica ou cadastro incompleto no proprio Omie. Sem separar isso,
           ele volta como "pendente" a cada execucao e o aviso manda rodar de novo
           para sempre. */
        const semCnpjNoOmie = Object.keys(rf.mapa).filter(function (c) {
          return rf.mapa[c] && !rf.mapa[c].cnpj;
        }).length;
        if (semCnpjNoOmie) {
          bloco.semCnpjNoOmie = semCnpjNoOmie;
          diag.avisos.push(u + ': ' + semCnpjNoOmie + ' fornecedor(es) sem CNPJ no ' +
            'cadastro do proprio Omie (pessoa fisica ou cadastro incompleto). ' +
            'Rodar de novo nao resolve — precisa arrumar no Omie.');
        }
        if (bloco.fornecedoresAindaPendentes > 0) {
          diag.avisos.push(u + ': ainda faltam ' + bloco.fornecedoresAindaPendentes +
            ' fornecedor(es). Rode de novo com ?apenasFornecedores=1.');
        }
        continue;
      }

      /* O INDICE VEM PRIMEIRO, de proposito. Ele decide se a consulta de
         alteracoes vale a pena: na primeira carga a lista esta vazia, entao nao
         existe card para mover para "Quitadas" e as ~584 alteracoes seriam ~10s
         jogados fora de um orcamento de 45s. */
      let indice;
      try { indice = await indexarPorCodigoOmie(client, siteId, u); }
      catch (e) { bloco.erro = 'indice: ' + e.message; continue; }
      const primeiraCarga = Object.keys(indice).length === 0;
      bloco.jaNaLista = Object.keys(indice).length;

      /* 1) tudo que esta em aberto */
      let abertas = [];
      try {
        const r = await listarContasPagarPorVencimento(
          { filtroExtra: { filtrar_por_status: 'EMABERTO' },
            maxPaginas: 20, paginaInicial: paginaInicial }, creds);
        abertas = r.contas;
        bloco.emAberto = abertas.length;
        bloco.truncado = r.truncado;
        bloco.paginasLidas = { de: paginaInicial, ate: r.ultimaPaginaLida, total: r.totalPaginas };
        if (r.truncado) {
          /* O aviso antigo mandava "aumentar maxPaginas ou sincronizar a unidade
             sozinha" — nenhum dos dois resolvia, porque a leitura sempre recomecava
             da pagina 1. Agora diz exatamente qual URL rodar em seguida. */
          bloco.proximaPagina = r.proximaPagina;
          diag.avisos.push(u + ': faltam paginas (' + r.ultimaPaginaLida + ' de ' +
            r.totalPaginas + ' lidas). Rode em seguida: ?unidade=' + u +
            '&paginaInicial=' + r.proximaPagina);
        }
      } catch (e) { bloco.erro = 'em aberto: ' + e.message; continue; }

      /* 2) o que mudou — e como o card sai de "Aprovadas" para "Quitadas" */
      let alteradas = [];
      if (primeiraCarga) {
        diag.avisos.push(u + ': primeira carga — pulei a consulta de alteracoes, ' +
          'que so serve para mover cards que ainda nao existem.');
      } else try {
        const r = await listarContasPagarPorVencimento({
          de: desde, ate: hoje,
          filtroExtra: { filtrar_apenas_alteracao: 'S' },
          maxPaginas: 12
        }, creds);
        alteradas = r.contas;
        bloco.alteradas = alteradas.length;
      } catch (e) {
        /* Nao derruba a unidade: sem este passo o quadro fica desatualizado nas
           quitacoes, mas as contas em aberto ja entraram. */
        diag.avisos.push(u + ': nao consegui ler as alteracoes (' + e.message + ').');
      }

      /* Uniao pelo codigo do lancamento. Uma conta que esta em aberto E foi
         alterada aparece nas duas listas; a versao alterada e a mais recente. */
      const porCodigo = {};
      for (const c of abertas) porCodigo[String(c.codigo_lancamento_omie)] = c;
      for (const c of alteradas) porCodigo[String(c.codigo_lancamento_omie)] = c;
      let contas = Object.keys(porCodigo).map(function (k) { return porCodigo[k]; });
      bloco.contas = contas.length;

      if (dryRun) {
        const st = {};
        for (const c of contas) st[String(c.status_titulo || '?')] = (st[String(c.status_titulo || '?')] || 0) + 1;
        bloco.statusTitulo = st;
        bloco.comChaveNFe = contas.filter(function (c) { return soDigitos(c.chave_nfe).length === 44; }).length;
        continue;
      }

      /* CONTA PAGA QUE NUNCA CONHECEMOS NAO ENTRA. A consulta de alteracoes traz
         centenas de titulos ja quitados (medido no RJ: 300 de 830) que nunca
         passaram pelo quadro. Guardar isso so incharia a lista sem informar nada:
         o quadro existe para acompanhar o que ainda vai ser pago. Uma conta que
         JA esta na lista e virou PAGO continua entrando — e assim que o card
         chega em "Quitadas". */
      const antesDoCorte = contas.length;
      contas = contas.filter(function (c) {
        const paga = String(c.status_titulo || '').toUpperCase().indexOf('PAGO') >= 0;
        return !paga || !!indice[String(c.codigo_lancamento_omie)];
      });
      bloco.pagasIgnoradas = antesDoCorte - contas.length;

      /* Vencimento anterior ao corte: fora. Conta SEM vencimento tambem fica de
         fora — nao da para dizer se e recente, e deixar entrar traria de volta
         justamente a sujeira que o corte existe para evitar. */
      const antesDoVenc = contas.length;
      contas = contas.filter(function (c) {
        const v = dataOmieParaISO(c.data_vencimento);
        return v && v >= corte.data;
      });
      bloco.antigasIgnoradas = antesDoVenc - contas.length;

      /* A conta a pagar traz so o CODIGO interno do fornecedor, e quem decide
         unidade e diretoria e o CNPJ. Resolve apenas os codigos que ainda nao
         temos gravados, com teto por execucao: o que sobrar resolve na proxima,
         e o CNPJ ja descoberto fica na linha para sempre. */
      const jaResolvido = {};
      for (const k of Object.keys(indice)) {
        const f = indice[k].fields || {};
        if (f.CodigoClienteOmie && soDigitos(f.EmitenteCNPJ).length === 14) {
          jaResolvido[String(f.CodigoClienteOmie)] = {
            cnpj: soDigitos(f.EmitenteCNPJ), razao: f.EmitenteNome || ''
          };
        }
      }
      const codigosFaltando = [];
      const vistos = {};
      for (const c of contas) {
        const cod = String(c.codigo_cliente_fornecedor || '');
        if (!cod || vistos[cod] || jaResolvido[cod]) continue;
        vistos[cod] = true;
        codigosFaltando.push(cod);
      }
      let mapaForn = jaResolvido;
      if (codigosFaltando.length) {
        /* Prazo ABSOLUTO, verificado a cada consulta dentro da funcao. A versao
           antiga estimava o custo (~600ms por consulta) e parava pela contagem;
           quando o Omie respondia mais devagar, a estimativa furava e a gravacao
           herdava um prazo ja vencido. */
        const prazoForn = t0 + PRAZO_FORNECEDORES_MS;
        const teto = Math.max(0, Math.min(60, Math.floor((prazoForn - Date.now()) / 600)));
        try {
          const rf = await resolverFornecedoresPorCodigo(codigosFaltando, creds, teto, prazoForn);
          mapaForn = Object.assign({}, jaResolvido, rf.mapa);
          bloco.fornecedoresConsultados = rf.consultados;
          bloco.fornecedoresPendentes = codigosFaltando.length - rf.consultados;
          if (bloco.fornecedoresPendentes > 0) {
            diag.avisos.push(u + ': ' + bloco.fornecedoresPendentes + ' fornecedor(es) ainda ' +
              'sem CNPJ resolvido — a proxima sincronizacao continua de onde parou. ' +
              'Ate la esses cards aparecem como "fornecedor sem cadastro".');
          }
        } catch (e) {
          diag.avisos.push(u + ': falha ao resolver fornecedores (' + e.message + ')');
        }
      }

      /* Monta TUDO primeiro, grava em lotes de 20 pelo $batch do Graph. Uma a uma
         seriam ~0,3s cada: 541 linhas dariam ~160s e a Function corta em 45. */
      const listId = await resolveListId(client, siteId, LIST_DOCFIS);
      const ops = [];
      for (const c of contas) {
        if (ehCancelada(c.status_titulo)) { bloco.canceladas++; continue; }
        const cod = String(c.codigo_cliente_fornecedor || '');
        const forn = mapaForn[cod] || {};
        if (!forn.cnpj) bloco.semFornecedor++;

        const op = prepararContaOmie({
          codigoLancamentoOmie: c.codigo_lancamento_omie,
          codigoClienteOmie: cod,
          chaveAcesso: c.chave_nfe,
          numeroNF: c.numero_documento_fiscal || c.nota_fiscal || c.numero_documento || '',
          numeroParcela: c.numero_parcela,
          emitenteCNPJ: forn.cnpj || '',
          emitenteNome: forn.razao || '',
          valor: c.valor_documento,
          dataEmissao: dataOmieParaISO(c.data_emissao),
          dataVencimento: dataOmieParaISO(c.data_vencimento),
          statusOmie: c.status_titulo || '',
          unidade: u,
          codigoBarras: c.codigo_barras_ficha_compensacao || ''
        }, indice);

        if (!op || op.tipo === 'nenhum') { bloco.semMudanca++; continue; }
        ops.push(op);
      }

      const r = await gravarEmLote(client, siteId, listId, ops, t0 + ORCAMENTO_MS);
      /* PLANEJADO nao e FEITO. Estes dois contam o que foi MONTADO; o que entrou
         no SharePoint e r.ok. Enquanto a mensagem do topo somava os planejados,
         uma execucao que gravou zero anunciava "355 nova(s), 44 atualizada(s)" —
         e quem lesse so a primeira linha ia embora achando que estava sincronizado. */
      bloco.planejadosNovos = ops.filter(function (o) { return o.tipo === 'post'; }).length;
      bloco.planejadosAtualizados = ops.filter(function (o) { return o.tipo === 'patch'; }).length;
      bloco.gravados = r.ok;
      if (r.restantes > 0) {
        bloco.restantes = r.restantes;
        diag.avisos.push(u + ': parou no prazo com ' + r.restantes + ' de ' + ops.length +
          ' ainda por gravar. Rode /api/SincronizarOmie?unidade=' + u + ' de novo — ' +
          'ele continua de onde parou e nada se perde.');
      }
      if (r.falhas.length) {
        /* Falha parcial nao pode passar por sucesso: a proxima execucao reprocessa
           o que nao entrou, mas quem le o resultado precisa saber. */
        bloco.falhas = r.falhas.length;
        bloco.primeirasFalhas = r.falhas.slice(0, 3);
        diag.avisos.push(u + ': ' + r.falhas.length + ' de ' + ops.length +
          ' gravacoes falharam — a proxima sincronizacao tenta de novo.');
      }
    }

    diag.step = 'done';
    diag.timeMs = Date.now() - t0;
    /* A mensagem conta o que ENTROU no SharePoint, nao o que foi planejado — e
       precisa cobrir OS DOIS modos. A primeira versao disto so somava `gravados`,
       entao uma execucao de ?apenasFornecedores=1 que preencheu 69 CNPJs anunciava
       "nada mudou". Trocar uma mensagem mentirosa por outra e facil quando o
       resumo conhece so metade do trabalho; por isso agora ele monta as partes a
       partir de cada contador real, e "nada mudou" so aparece se nenhum deles
       tiver acontecido. */
    function somar(campo) {
      return diag.unidades.reduce(function (s, u) { return s + (u[campo] || 0); }, 0);
    }
    const gravados = somar('gravados');
    const cnpjPreenchidos = somar('linhasAtualizadas');
    const restantes = somar('restantes');
    const fornPendentes = somar('fornecedoresAindaPendentes');
    const comErro = diag.unidades.filter(function (u) { return u.erro; });

    const partes = [];
    if (gravados) partes.push(gravados + ' conta(s) gravada(s)');
    if (cnpjPreenchidos) partes.push(cnpjPreenchidos + ' linha(s) com CNPJ preenchido');
    if (restantes) partes.push(restantes + ' ainda por gravar, rode de novo');
    if (fornPendentes) partes.push(fornPendentes + ' fornecedor(es) pendentes');
    if (comErro.length) partes.push(comErro.length + ' unidade(s) com erro');
    if (!partes.length) partes.push('nada mudou');

    context.res = { status: 200, headers: { 'Content-Type': 'application/json' },
      body: Object.assign({
        ok: comErro.length === 0,
        mensagem: (dryRun ? '[SIMULACAO] ' : '') + partes.join(' · ')
      }, diag) };
  } catch (err) {
    diag.timeMs = Date.now() - t0;
    context.res = { status: 500, headers: { 'Content-Type': 'application/json' },
      body: Object.assign({ error: (err && err.message) || String(err) }, diag) };
  }
};
