/**
 * /api/BatimentoAgentes  (POST grava, GET le)
 *
 * O VIGIA DO VIGIA.
 *
 * Desde 15/09/2026 as automacoes estao saindo do GitHub Actions e passando a
 * rodar como agentes no servidor PRONEP-ANALYTICS. A troca e boa — diario passo
 * a passo, auditor em codigo, tela onde se ve o agente trabalhando — mas ela
 * cria um risco que nao existia: o GitHub e altamente disponivel; o nosso
 * servidor nao. Se ele desligar, ou o agendador parar, os agentes param EM
 * SILENCIO. E a tela que mostraria isso roda no mesmo servidor que caiu.
 *
 * Entao este endpoint inverte o papel do GitHub: ele deixa de MANDAR trabalho e
 * passa a CONFERIR que o trabalho aconteceu. O servidor deixa um batimento aqui
 * depois de cada rodada; um workflow agendado le e falha alto se alguem emudeceu.
 *
 * A TOLERANCIA VEM DO DADO, NAO DO WORKFLOW: cada agente declara de quantas em
 * quantas horas ele deveria dar sinal (`tolerancia_horas`, que sai do catalogo
 * de tarefas do servidor). Assim criar um agente novo nao exige editar o vigia —
 * que e exatamente o tipo de passo que se esquece.
 *
 * Auth: X-Automacao-Secret == App Setting AUTOMACAO_EMAILS_SECRET, nos DOIS
 * verbos. Reusa um segredo que ja existe no Azure, no GitHub e no servidor: um
 * vigia que exige credencial nova e um vigia que alguem deixa de instalar.
 */

require('isomorphic-fetch');
const { getGraphClient, resolveSiteId } = require('../shared/graph');

const CAMINHO = '_automacao/batimento_agentes.json';
// Teto de seguranca: um agente nao pode declarar que some por um mes. Se algum
// dia existir uma tarefa assim de verdade, mude AQUI, conscientemente.
const TOLERANCIA_MAX_H = 840;   // 35 dias — cabe a cadeia mensal do BRAZIL

async function ler(client, siteId) {
  try {
    const r = await client.api('/sites/' + siteId + '/drive/root:/' + CAMINHO + ':/content').get();
    if (r && typeof r === 'object') return r;
    if (typeof r === 'string') { try { return JSON.parse(r); } catch (e) { return {}; } }
    return {};
  } catch (e) { return {}; }   // 404 = ainda nao existe, e isso nao e erro
}

async function gravar(client, siteId, dados) {
  await client.api('/sites/' + siteId + '/drive/root:/' + CAMINHO + ':/content')
    .header('Content-Type', 'application/json')
    .put(Buffer.from(JSON.stringify(dados), 'utf-8'));
}

module.exports = async function (context, req) {
  try {
    const segredo = process.env.AUTOMACAO_EMAILS_SECRET;
    if (!segredo) {
      context.res = { status: 500, headers: { 'Content-Type': 'application/json' },
        body: { error: 'AUTOMACAO_EMAILS_SECRET nao configurado' } };
      return;
    }
    const enviado = (req.headers && (req.headers['x-automacao-secret'] || req.headers['X-Automacao-Secret'])) || '';
    if (enviado !== segredo) {
      context.res = { status: 403, headers: { 'Content-Type': 'application/json' },
        body: { error: 'segredo invalido' } };
      return;
    }

    const client = getGraphClient();
    const siteId = await resolveSiteId(client);
    const agora = Date.now();

    if ((req.method || '').toUpperCase() === 'POST') {
      const c = (req.body && typeof req.body === 'object') ? req.body : {};
      const agente = String(c.agente || '').trim();
      if (!agente) {
        context.res = { status: 400, headers: { 'Content-Type': 'application/json' },
          body: { error: 'passe {agente}' } };
        return;
      }
      const dados = await ler(client, siteId);
      dados.agentes = dados.agentes || {};
      dados.agentes[agente] = {
        // `quando` e SEMPRE a hora do servidor que recebeu, nunca a que o
        // cliente mandou: relogio atrasado no servidor de origem faria o
        // batimento parecer velho — ou, pior, sempre novo.
        quando: new Date(agora).toISOString(),
        status: String(c.status || 'ok').slice(0, 30),
        tolerancia_horas: Math.min(TOLERANCIA_MAX_H,
                                   Math.max(1, Number(c.tolerancia_horas) || 3)),
        maquina: String(c.maquina || '').slice(0, 60),
        detalhe: String(c.detalhe || '').slice(0, 200)
      };
      dados.atualizadoEm = new Date(agora).toISOString();
      await gravar(client, siteId, dados);
      context.res = { status: 200, headers: { 'Content-Type': 'application/json' },
        body: { ok: true, agente: agente, quando: dados.agentes[agente].quando } };
      return;
    }

    // ---- GET: devolve a idade de cada batimento e quem passou do prazo
    const dados = await ler(client, siteId);
    const ags = dados.agentes || {};
    const lista = Object.keys(ags).map(function (nome) {
      const a = ags[nome];
      const idade = (agora - Date.parse(a.quando)) / 3600000;
      return {
        agente: nome, quando: a.quando, status: a.status,
        maquina: a.maquina || null, detalhe: a.detalhe || null,
        idade_horas: Math.round(idade * 10) / 10,
        tolerancia_horas: a.tolerancia_horas,
        mudo: idade > a.tolerancia_horas
      };
    }).sort(function (x, y) { return y.idade_horas - x.idade_horas; });

    const mudos = lista.filter(function (x) { return x.mudo; });
    context.res = {
      status: 200,
      headers: { 'Content-Type': 'application/json', 'Cache-Control': 'no-store' },
      body: {
        ok: mudos.length === 0,
        // ⚠ `nunca_reportou` e um estado DIFERENTE de "mudo", e nao da para
        // detecta-lo aqui: um agente que nunca postou simplesmente nao esta
        // nesta lista. Quem sabe quais agentes DEVERIAM existir e o servidor —
        // e se o servidor esta morto, ninguem sabe. Por isso o workflow tambem
        // cobra um numero minimo de agentes: ver vigia-servidor.yml.
        agentes: lista, quantos: lista.length,
        mudos: mudos.map(function (x) { return x.agente; }),
        atualizadoEm: dados.atualizadoEm || null
      }
    };
  } catch (err) {
    context.log && context.log.error && context.log.error('BatimentoAgentes:', err);
    context.res = { status: 500, headers: { 'Content-Type': 'application/json' },
      body: { error: (err && err.message) || String(err) } };
  }
};
