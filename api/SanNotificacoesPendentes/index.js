/**
 * /api/SanNotificacoesPendentes
 *
 * Detecta pendencias relevantes pro user logado e retorna:
 *   - count total
 *   - breakdown por categoria
 *   - saudacao proativa montada por template
 *   - sugestoes de acao clicaveis
 *
 * Categorias detectadas:
 *   - contratos vencendo (gestor da diretoria)
 *   - contratos ja vencidos sem renegociacao (gestor)
 *   - NFs pendentes na fila do user (AprovadorAtual = email do user)
 *   - NFs rejeitadas re-submetidas (LancadoPor do user, ainda em status Rejeitada)
 *
 * Custo Claude: ZERO. So consulta SP.
 * Pode ser chamado em polling a cada 5 minutos pelo front (light).
 *
 * Auth: getUser (Easy Auth ou Teams JWT).
 */

require('isomorphic-fetch');
const { getGraphClient } = require('../shared/graph');
const { getUser } = require('../shared/auth');
const { getUserRoles } = require('../shared/userRoles');

const LIST_NOTAS = 'PRONEP-NF-NotasFiscais';
const LIST_CONTRATOS = 'PRONEP-NF-Contratos';
const LIST_DIRETORIAS = 'PRONEP-NF-Diretorias';

const _cache = { siteId: null, listNotasId: null, listContratosId: null, listDirId: null, invColNotas: null, invColContratos: null };

async function resolveSiteEListas(client) {
  if (_cache.siteId && _cache.listNotasId) return _cache;
  const host = process.env.SHAREPOINT_SITE_HOSTNAME;
  const path = process.env.SHAREPOINT_SITE_PATH;
  const siteResp = await client.api('/sites/' + host + ':' + path).get();
  _cache.siteId = siteResp.id;
  const lists = await client.api('/sites/' + _cache.siteId + '/lists').get();
  for (const l of (lists.value || [])) {
    if (l.displayName === LIST_NOTAS) _cache.listNotasId = l.id;
    if (l.displayName === LIST_CONTRATOS) _cache.listContratosId = l.id;
    if (l.displayName === LIST_DIRETORIAS) _cache.listDirId = l.id;
  }
  // Cols
  const cn = await client.api('/sites/' + _cache.siteId + '/lists/' + _cache.listNotasId + '/columns').get();
  _cache.invColNotas = {};
  for (const c of (cn.value || [])) if (c.displayName && c.name) _cache.invColNotas[c.name] = c.displayName;
  if (_cache.listContratosId) {
    const cc = await client.api('/sites/' + _cache.siteId + '/lists/' + _cache.listContratosId + '/columns').get();
    _cache.invColContratos = {};
    for (const c of (cc.value || [])) if (c.displayName && c.name) _cache.invColContratos[c.name] = c.displayName;
  }
  return _cache;
}

// MESMA logica do AlertaContratosDiario - pra que a SAN proativa nao alerte
// sobre contratos que ja foram notificados nessa janela (via email ou via SAN antes).
// Retorna a janela aplicavel (30|60|90) OU null se ja foi vista OU fora do range.
const JANELAS = [30, 60, 90];
function escolherJanelaAplicavelContrato(diasFalta, observacoes) {
  if (diasFalta == null || diasFalta < 0) return null;
  if (diasFalta > 90) return null;
  const obs = String(observacoes || '');
  const ja = {
    30: /_alerta_30=/.test(obs),
    60: /_alerta_60=/.test(obs),
    90: /_alerta_90=/.test(obs)
  };
  for (const j of JANELAS) {  // mais urgente primeiro
    if (diasFalta <= j && !ja[j]) return j;
  }
  return null;
}
// Pra contratos vencidos, alerta a primeira vez e depois espera 30 dias
function aplicavelParaContratoVencido(diasFalta, observacoes) {
  if (diasFalta == null || diasFalta >= 0) return false;
  const obs = String(observacoes || '');
  const m = obs.match(/_alerta_vencido=(\d{4}-\d{2}-\d{2})/);
  if (!m) return true;  // nunca alertou
  // Re-alerta a cada 30 dias se nao foi resolvido
  const ultimo = new Date(m[1] + 'T00:00:00Z');
  const hoje = new Date(new Date().getTime() - 3*60*60*1000);
  const diasDesde = Math.round((hoje.getTime() - ultimo.getTime()) / (24*60*60*1000));
  return diasDesde >= 30;
}

function diasParaVencer(dataFim) {
  if (!dataFim) return null;
  const hoje = new Date(new Date().getTime() - 3*60*60*1000);
  const hojeStr = hoje.toISOString().substring(0,10);
  const hj = new Date(hojeStr + 'T00:00:00Z');
  const fim = new Date(String(dataFim).substring(0,10) + 'T00:00:00Z');
  return Math.round((fim.getTime() - hj.getTime()) / (24*60*60*1000));
}

function diasUteisEntre(d1, d2) {
  if (!d1 || !d2) return null;
  let count = 0;
  const inicio = new Date(String(d1).substring(0,10) + 'T00:00:00Z');
  const fim = new Date(String(d2).substring(0,10) + 'T00:00:00Z');
  while (inicio <= fim) {
    const dow = inicio.getUTCDay();
    if (dow !== 0 && dow !== 6) count++;
    inicio.setUTCDate(inicio.getUTCDate() + 1);
  }
  return count;
}

async function paginar(client, baseUrl, maxPages) {
  const todos = [];
  let nextUrl = baseUrl;
  let pages = 0;
  while (nextUrl && pages < (maxPages || 50)) {
    pages++;
    const r = await client.api(nextUrl).get();
    for (const it of (r.value || [])) todos.push(it);
    const next = r['@odata.nextLink'];
    if (next) {
      const idx = next.indexOf('/v1.0/');
      nextUrl = idx >= 0 ? next.substring(idx + 5) : null;
    } else nextUrl = null;
  }
  return todos;
}

function normalize(item, invColMap) {
  const raw = item.fields || {};
  const out = { _id: item.id };
  for (const [k, v] of Object.entries(raw)) {
    const display = invColMap[k] || k;
    out[display] = v;
  }
  return out;
}

async function diretoriasDoGestor(client, siteId, listDirId, userEmail) {
  if (!listDirId || !userEmail) return [];
  const r = await client.api('/sites/' + siteId + '/lists/' + listDirId + '/items?expand=fields&$top=200').get();
  const set = new Set();
  for (const it of (r.value || [])) {
    const f = it.fields || {};
    const emailDir = String(f.field_3 || '').toLowerCase().trim();
    if (emailDir === userEmail) {
      const dir = String(f.Title || '').split('|')[1] || '';
      if (dir) set.add(dir.trim());
    }
  }
  return Array.from(set);
}

function primeiroNome(user) {
  return (user.name || user.email || '').split(/[\s.@]/)[0]
    .replace(/[^A-Za-zÀ-ÿ]/g, '')
    .replace(/^./, c => c.toUpperCase());
}

function montarSaudacaoTemplate(nome, breakdown) {
  const partes = [];
  if (breakdown.nfs_pendentes_total > 0) {
    const d5 = breakdown.nfs_d5 > 0 ? ' (**' + breakdown.nfs_d5 + '** com vencimento em ≤5 dias)' : '';
    partes.push('📬 **' + breakdown.nfs_pendentes_total + ' NF' + (breakdown.nfs_pendentes_total === 1 ? '' : 's') + '** aguardando sua aprovação' + d5);
  }
  if (breakdown.contratos_30 > 0) {
    partes.push('🚨 **' + breakdown.contratos_30 + ' contrato' + (breakdown.contratos_30 === 1 ? '' : 's') + '** vencendo em até 30 dias');
  }
  if (breakdown.contratos_60 > 0) {
    partes.push('⚠ **' + breakdown.contratos_60 + ' contrato' + (breakdown.contratos_60 === 1 ? '' : 's') + '** vencendo entre 31 e 60 dias');
  }
  if (breakdown.contratos_90 > 0) {
    partes.push('⏰ **' + breakdown.contratos_90 + ' contrato' + (breakdown.contratos_90 === 1 ? '' : 's') + '** vencendo entre 61 e 90 dias');
  }
  if (breakdown.contratos_vencidos > 0) {
    partes.push('❗ **' + breakdown.contratos_vencidos + ' contrato' + (breakdown.contratos_vencidos === 1 ? '' : 's') + '** já vencido' + (breakdown.contratos_vencidos === 1 ? '' : 's') + ' sem renegociação registrada');
  }
  if (breakdown.nfs_rejeitadas > 0) {
    partes.push('↩ **' + breakdown.nfs_rejeitadas + ' NF' + (breakdown.nfs_rejeitadas === 1 ? '' : 's') + '** sua' + (breakdown.nfs_rejeitadas === 1 ? '' : 's') + ' foi/foram rejeitada' + (breakdown.nfs_rejeitadas === 1 ? '' : 's') + ' — aguarda correção e reenvio');
  }

  if (partes.length === 0) {
    return 'Oi ' + nome + '! 👋 Tudo em dia por aqui. Sem pendências pra você no momento. Posso ajudar com mais alguma coisa?';
  }
  return 'Oi ' + nome + '! 👋 Dá uma olhada no que tá aguardando você:\n\n• ' + partes.join('\n• ') + '\n\nQuer que eu liste com prioridade?';
}

function montarSugestoes(breakdown) {
  const sug = [];
  if (breakdown.nfs_d5 > 0) sug.push('Mostre minhas NFs urgentes (D+5)');
  else if (breakdown.nfs_pendentes_total > 0) sug.push('Liste minha fila de aprovação');
  if (breakdown.contratos_30 > 0) sug.push('Quais contratos vencem em 30 dias?');
  else if (breakdown.contratos_60 > 0 || breakdown.contratos_90 > 0) sug.push('Liste contratos vencendo em 90 dias');
  if (breakdown.contratos_vencidos > 0) sug.push('Contratos vencidos sem renegociação');
  if (breakdown.nfs_rejeitadas > 0) sug.push('Quais NFs minhas foram rejeitadas?');
  return sug.slice(0, 3);  // max 3 sugestoes
}

module.exports = async function (context, req) {
  try {
    const user = await getUser(req);
    if (!user || !user.email) {
      context.res = { status: 401, body: { error: 'Nao autenticado' } };
      return;
    }
    const emailLow = String(user.email).toLowerCase();

    const client = getGraphClient();
    const cache = await resolveSiteEListas(client);
    const { siteId, listNotasId, listContratosId, listDirId, invColNotas, invColContratos } = cache;

    // Perfil
    let roles = [];
    try { roles = await getUserRoles(user); } catch (e) {}
    const isAdmin = roles.includes('administrador') || roles.includes('admin');
    // Juridico tem acesso TOTAL ao acervo de contratos (mesmo tratamento de admin nessa view)
    const isJuridicoFullAccess = roles.includes('gestor_juridica');
    const veTodosContratos = isAdmin || isJuridicoFullAccess;
    const isGestor = roles.includes('gestor') || roles.some(function(r){ return /^gestor_/.test(r); });

    // Diretorias que o user gerencia
    const minhasDiretorias = await diretoriasDoGestor(client, siteId, listDirId, emailLow);
    const filtraPorDiretoria = function(diretoriaItem) {
      if (veTodosContratos) return true;  // admin OU juridico ve tudo
      if (!minhasDiretorias.length) return false;
      return minhasDiretorias.includes(diretoriaItem);
    };

    // Carrega contratos (pode ser grande - paginar)
    const breakdown = {
      contratos_30: 0, contratos_60: 0, contratos_90: 0, contratos_vencidos: 0,
      nfs_pendentes_total: 0, nfs_d5: 0, nfs_rejeitadas: 0
    };
    const itensContratos = [];
    // Pra cada contrato com pendencia ainda nao notificada nessa janela, registra
    // o "marcador" que o front vai gravar quando o gestor abrir a SAN.
    // Cada item: { id, janela: 30|60|90|'vencido' }
    const marcadores = [];
    if (listContratosId && (veTodosContratos || minhasDiretorias.length)) {
      const all = await paginar(client, '/sites/' + siteId + '/lists/' + listContratosId + '/items?expand=fields&$top=999');
      for (const it of all) {
        const c = normalize(it, invColContratos);
        if (c.Status === 'Cancelado') continue;
        if (!filtraPorDiretoria(c.Diretoria)) continue;
        if (!c.DataFim) continue;
        const dias = diasParaVencer(c.DataFim);
        if (dias == null) continue;
        if (dias < 0) {
          // Vencido: alerta primeira vez OU re-alerta a cada 30 dias
          if (!aplicavelParaContratoVencido(dias, c.Observacoes)) continue;
          breakdown.contratos_vencidos++;
          itensContratos.push({ id: c._id, fornecedor: c.Fornecedor, dias: dias, diretoria: c.Diretoria, tipo: 'vencido' });
          marcadores.push({ id: c._id, janela: 'vencido' });
        } else {
          // Dedup por janela: se a janela aplicavel ja foi vista (em email ou SAN), pula
          const janela = escolherJanelaAplicavelContrato(dias, c.Observacoes);
          if (!janela) continue;  // ja viu essa janela OU >90d
          if (janela === 30) breakdown.contratos_30++;
          else if (janela === 60) breakdown.contratos_60++;
          else if (janela === 90) breakdown.contratos_90++;
          itensContratos.push({ id: c._id, fornecedor: c.Fornecedor, dias: dias, diretoria: c.Diretoria, tipo: 'venc_' + janela });
          marcadores.push({ id: c._id, janela: janela });
        }
      }
    }

    // Carrega NFs (pendentes onde sou aprovador + rejeitadas que eu submeti)
    // NFs rejeitadas: dedup persistente via tag em Observacao - se ja vi essa rejeicao, nao mostra mais
    const itensNFs = [];
    const marcadoresNFs = [];  // ids de NFs rejeitadas que vou marcar como vistas quando user abrir SAN
    const tagVistoRejeicao = '_visto_rejeicao_' + emailLow + '=';
    if (listNotasId) {
      const all = await paginar(client, '/sites/' + siteId + '/lists/' + listNotasId + '/items?expand=fields&$top=999');
      const hojeStr = new Date(new Date().getTime() - 3*60*60*1000).toISOString().substring(0,10);
      for (const it of all) {
        const n = normalize(it, invColNotas);
        const status = String(n.Status || '');
        const aprovador = String(n.AprovadorAtual || '').toLowerCase();
        const lancadoPor = String(n.LancadoPor || '').toLowerCase();
        // NFs pendentes onde EU sou o aprovador atual (SEM dedup - urgencia imediata)
        if (aprovador === emailLow && ['Lancada','EmAprovacao','Pendente'].includes(status)) {
          breakdown.nfs_pendentes_total++;
          const venc = String(n.DataVencimento || n.Vencimento || '').substring(0,10);
          const diasUteis = diasUteisEntre(hojeStr, venc);
          const ehD5 = diasUteis != null && diasUteis <= 5;
          if (ehD5) breakdown.nfs_d5++;
          itensNFs.push({ id: n._id, numero: n.NumeroNF, fornecedor: n.Fornecedor, vencimento: venc, ehD5: ehD5, tipo: 'pendente' });
        }
        // NFs rejeitadas que EU submeti — com dedup por user em Observacao
        if (lancadoPor === emailLow && status === 'Rejeitada') {
          const obs = String(n.Observacao || '');
          if (obs.indexOf(tagVistoRejeicao) >= 0) continue;  // ja visto, pula
          breakdown.nfs_rejeitadas++;
          itensNFs.push({ id: n._id, numero: n.NumeroNF, fornecedor: n.Fornecedor, motivo: n.MotivoRejeicao, tipo: 'rejeitada' });
          marcadoresNFs.push({ id: n._id, tipo: 'rejeitada' });
        }
      }
    }

    const count = breakdown.contratos_30 + breakdown.contratos_60 + breakdown.contratos_90 +
                  breakdown.contratos_vencidos + breakdown.nfs_pendentes_total + breakdown.nfs_rejeitadas;
    const nome = primeiroNome(user);
    const saudacao = montarSaudacaoTemplate(nome, breakdown);
    const sugestoes = montarSugestoes(breakdown);

    context.res = {
      status: 200,
      headers: { 'Content-Type': 'application/json', 'Cache-Control': 'no-store' },
      body: {
        ok: true,
        count,
        breakdown,
        saudacao,
        sugestoes,
        // Top 10 itens mais urgentes pra debug/inspecao
        itens: [].concat(
          itensContratos.sort((a, b) => a.dias - b.dias).slice(0, 5),
          itensNFs.slice(0, 5)
        ),
        // Marcadores que o front vai chamar POST pra marcar como visto apos exibir
        marcadoresContratos: marcadores,
        marcadoresNFsRejeitadas: marcadoresNFs,
        user: { nome, email: emailLow, isAdmin, isJuridicoFullAccess, isGestor, diretorias: minhasDiretorias }
      }
    };
  } catch (err) {
    context.log && context.log.error && context.log.error('SanNotificacoesPendentes:', err);
    context.res = { status: 500, body: { error: err.message, stack: (err.stack || '').split('\n').slice(0, 8) } };
  }
};
