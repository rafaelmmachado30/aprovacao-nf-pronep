/**
 * /api/ListarContratos — lista contratos da PRONEP-NF-Contratos com RBAC.
 *
 * RBAC:
 *   - admin: ve TUDO
 *   - gestor: ve so contratos das diretorias que ele gerencia
 *             (cruzamento com PRONEP-NF-Diretorias)
 *   - financeiro / submitter: 403
 *
 * Query params:
 *   - diretoria   (opcional): filtra
 *   - unidade     (opcional): filtra (CORPORATIVO|SP|RJ|ES)
 *   - status      (opcional): Ativo|Vencendo30|Vencendo60|Vencendo90|Vencido|SemVigencia
 *   - busca       (opcional): substring em fornecedor/nome
 *   - formato     (opcional): "arvore" (default) ou "plano"
 *
 * Resposta:
 *   {
 *     ok: true,
 *     userScope: { roles, diretoriasGestorOf },
 *     stats: { total, ativos, vencendo30, vencendo60, vencendo90, vencidos, semVigencia },
 *     arvore: [
 *       { diretoria, total, unidades: [
 *         { unidade, total, fornecedores: [
 *           { fornecedor, contratos: [...] }
 *         ]}
 *       ]}
 *     ]
 *   }
 */

const { ClientSecretCredential } = require('@azure/identity');
const { Client } = require('@microsoft/microsoft-graph-client');
const { TokenCredentialAuthenticationProvider } = require('@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials');
const { getUser } = require('../shared/auth');
const { getUserRoles } = require('../shared/userRoles');
const contratosShared = require('../shared/contratos');

const cache = { siteId: null, listId: null, listDirId: null, colMap: null };

function getGraphClient() {
  const tenantId = process.env.AAD_TENANT_ID;
  const clientId = process.env.AAD_CLIENT_ID;
  const clientSecret = process.env.AAD_CLIENT_SECRET;
  const credential = new ClientSecretCredential(tenantId, clientId, clientSecret);
  const authProvider = new TokenCredentialAuthenticationProvider(credential, {
    scopes: ['https://graph.microsoft.com/.default']
  });
  return Client.initWithMiddleware({ authProvider });
}

async function resolveSiteELists(client) {
  if (cache.siteId && cache.listId && cache.listDirId) return cache;
  const host = process.env.SHAREPOINT_SITE_HOSTNAME;
  const path = process.env.SHAREPOINT_SITE_PATH;
  const siteResp = await client.api('/sites/' + host + ':' + path).get();
  cache.siteId = siteResp.id;
  const lists = await client.api('/sites/' + cache.siteId + '/lists').get();
  for (const l of (lists.value || [])) {
    if (l.displayName === 'PRONEP-NF-Contratos') cache.listId = l.id;
    if (l.displayName === 'PRONEP-NF-Diretorias') cache.listDirId = l.id;
  }
  return cache;
}

async function getColMap(client, siteId, listId) {
  if (cache.colMap) return cache.colMap;
  const resp = await client.api('/sites/' + siteId + '/lists/' + listId + '/columns').get();
  const map = {};
  for (const c of (resp.value || [])) {
    if (c.displayName && c.name) map[c.displayName] = c.name;
  }
  cache.colMap = map;
  return map;
}

function readClientPrincipal(req) {
  const header = req.headers && req.headers['x-ms-client-principal'];
  if (!header) return null;
  try { return JSON.parse(Buffer.from(header, 'base64').toString('utf-8')); } catch (e) { return null; }
}
function readClientPrincipalRoles(req) {
  const p = readClientPrincipal(req);
  return (p && p.userRoles) || [];
}

// Resolve quais diretorias o gestor tem responsabilidade
async function diretoriasDoGestor(client, siteId, listDirId, userEmail) {
  if (!listDirId || !userEmail) return [];
  const resp = await client.api('/sites/' + siteId + '/lists/' + listDirId + '/items?expand=fields&$top=200').get();
  const set = new Set();
  for (const it of (resp.value || [])) {
    const f = it.fields || {};
    // field_3 = email do aprovador. Title formato "Unidade|Diretoria"
    const emailDir = String(f.field_3 || '').toLowerCase().trim();
    if (emailDir === userEmail) {
      const title = String(f.Title || '');
      const dir = title.split('|')[1] || '';
      if (dir) set.add(dir.trim());
    }
  }
  return Array.from(set);
}

module.exports = async function (context, req) {
  try {
    // 1. Identifica user
    const user = await getUser(req);
    if (!user || !user.email) {
      context.res = { status: 401, body: { error: 'Nao autenticado' } };
      return;
    }

    // 2. RBAC
    const claimsRoles = readClientPrincipalRoles(req) || [];
    const userRoles = await getUserRoles(user);
    const allRoles = Array.from(new Set([].concat(claimsRoles, userRoles || [])));
    const isAdmin = allRoles.includes('administrador') || allRoles.includes('admin');
    // Juridico (gestor_juridica) ve TUDO em Contratos - mas no resto do sistema segue o RBAC dele.
    // Decisao Rafa 07/06/2026: Admin + Juridico tem acesso total ao acervo de contratos.
    const isJuridicoFullAccess = allRoles.includes('gestor_juridica');
    const veTodosContratos = isAdmin || isJuridicoFullAccess;
    const isGestor = allRoles.includes('gestor') || allRoles.some(function(r){ return /^gestor_/.test(r); });
    if (!veTodosContratos && !isGestor) {
      context.res = { status: 403, body: { error: 'Acesso negado. Restrito a Gestores, Juridico e Admins.' } };
      return;
    }

    // 3. Resolve site/lists
    const client = getGraphClient();
    const { siteId, listId, listDirId } = await resolveSiteELists(client);
    if (!listId) {
      context.res = {
        status: 200,
        headers: { 'Content-Type': 'application/json' },
        body: {
          ok: true,
          mensagem: 'Lista PRONEP-NF-Contratos ainda nao foi criada. Use /api/SincronizarContratos pra iniciar.',
          stats: zeroStats(),
          arvore: []
        }
      };
      return;
    }
    const colMap = await getColMap(client, siteId, listId);

    // 4. Define escopo de diretorias
    // null = sem filtro (admin OU juridico - ambos veem tudo no acervo de contratos)
    let scopeDiretorias = null;
    if (!veTodosContratos && isGestor) {
      scopeDiretorias = await diretoriasDoGestor(client, siteId, listDirId, user.email);
      if (!scopeDiretorias.length) {
        // Gestor sem mapeamento — retorna vazio
        context.res = {
          status: 200,
          headers: { 'Content-Type': 'application/json' },
          body: {
            ok: true,
            userScope: { roles: allRoles, diretoriasGestorOf: [] },
            mensagem: 'Voce eh gestor mas nao foi encontrado nenhum mapeamento na lista PRONEP-NF-Diretorias.',
            stats: zeroStats(),
            arvore: []
          }
        };
        return;
      }
    }

    // 5. Carrega TODOS os contratos com PAGINACAO (Graph limita 999 por pagina).
    // Lista pode ter milhares de itens — sem paginar so vem os primeiros 999.
    const todosItens = [];
    let nextUrl = '/sites/' + siteId + '/lists/' + listId + '/items?expand=fields&$top=999';
    let paginas = 0;
    while (nextUrl && paginas < 50) {  // hard stop em 50 paginas (~50k itens) por seguranca
      paginas++;
      const pageResp = await client.api(nextUrl).get();
      const itens = pageResp.value || [];
      for (const it of itens) todosItens.push(it);
      // O Graph retorna @odata.nextLink completo (URL absoluta). Extrai pra relativo.
      const next = pageResp['@odata.nextLink'];
      if (next) {
        const idx = next.indexOf('/v1.0/');
        nextUrl = idx >= 0 ? next.substring(idx + 5) : null;  // ex: /sites/.../items?...&$skiptoken=...
      } else {
        nextUrl = null;
      }
    }
    const resp = { value: todosItens };
    const f = req.query || {};
    const filtroBusca = String(f.busca || '').toLowerCase().trim();

    const cD = colMap['Diretoria'] || 'Diretoria';
    const cU = colMap['Unidade'] || 'Unidade';
    const cF = colMap['Fornecedor'] || 'Fornecedor';
    const cDI = colMap['DataInicio'] || 'DataInicio';
    const cDF = colMap['DataFim'] || 'DataFim';
    const cSt = colMap['Status'] || 'Status';
    const cSP = colMap['CaminhoSharepoint'] || 'CaminhoSharepoint';
    const cDriveId = colMap['DriveItemId'] || 'DriveItemId';
    const cLei = colMap['LeituraIAStatus'] || 'LeituraIAStatus';
    const cLeiTxt = colMap['LeituraIATexto'] || 'LeituraIATexto';
    const cVal = colMap['ValorContrato'] || 'ValorContrato';
    const cObs = colMap['Observacoes'] || 'Observacoes';
    const cNomeArq = colMap['NomeArquivo'] || 'NomeArquivo';
    const cPathSP = colMap['PathRelativoSP'] || 'PathRelativoSP';
    const cUltLei = colMap['UltimaLeitura'] || 'UltimaLeitura';
    const cCnpj = colMap['CNPJFornecedor'] || 'CNPJFornecedor';

    const todos = (resp.value || []).map(function(it) {
      const fl = it.fields || {};
      const dataFim = fl[cDF] || null;
      const dataIni = fl[cDI] || null;
      const dataFimStr = dataFim ? String(dataFim).substring(0, 10) : null;
      const status = fl[cSt] || contratosShared.calcularStatus(dataFimStr, false);
      const diasParaVencer = contratosShared.calcularDiasParaVencer(dataFimStr);
      return {
        id: it.id,
        title: fl.Title || '',
        diretoria: fl[cD] || '',
        unidade: fl[cU] || 'CORPORATIVO',
        fornecedor: fl[cF] || '',
        cnpj: fl[cCnpj] || '',
        dataInicio: dataIni ? String(dataIni).substring(0, 10) : null,
        dataFim: dataFimStr,
        diasParaVencer: diasParaVencer,
        status: status,
        valor: fl[cVal] || null,
        observacoes: fl[cObs] || '',
        nomeArquivo: fl[cNomeArq] || fl.Title || '',
        pathSP: fl[cPathSP] || '',
        driveItemId: fl[cDriveId] || '',
        urlSharePoint: fl[cSP] || '',
        leituraStatus: fl[cLei] || '',
        leituraTexto: fl[cLeiTxt] || '',
        ultimaLeitura: fl[cUltLei] || ''
      };
    });

    // 6. Aplica filtros
    let filtrados = todos;
    if (scopeDiretorias) {
      const set = new Set(scopeDiretorias);
      filtrados = filtrados.filter(function(c){ return set.has(c.diretoria); });
    }
    if (f.diretoria) filtrados = filtrados.filter(function(c){ return c.diretoria === f.diretoria; });
    if (f.unidade) filtrados = filtrados.filter(function(c){ return c.unidade === f.unidade; });
    if (f.status) filtrados = filtrados.filter(function(c){ return c.status === f.status; });
    if (filtroBusca) {
      filtrados = filtrados.filter(function(c) {
        const hay = ((c.fornecedor || '') + ' ' + (c.title || '') + ' ' + (c.nomeArquivo || '')).toLowerCase();
        return hay.includes(filtroBusca);
      });
    }

    // 7. Stats
    const stats = computarStats(filtrados);

    // 8. Estrutura em arvore
    const formato = String(f.formato || 'arvore');
    let arvore = [];
    if (formato === 'plano') {
      arvore = filtrados;
    } else {
      arvore = montarArvore(filtrados);
    }

    context.res = {
      status: 200,
      headers: { 'Content-Type': 'application/json' },
      body: {
        ok: true,
        userScope: { roles: allRoles, isAdmin, isJuridicoFullAccess, isGestor, diretoriasGestorOf: scopeDiretorias || 'TODAS' },
        stats,
        total: filtrados.length,
        formato,
        arvore,
        ...(formato === 'plano' ? { contratos: filtrados } : {})
      }
    };
  } catch (err) {
    context.log && context.log.error && context.log.error('ListarContratos error:', err);
    context.res = {
      status: 500,
      headers: { 'Content-Type': 'application/json' },
      body: { error: err.message, stack: (err.stack || '').split('\n').slice(0, 6) }
    };
  }
};

function zeroStats() {
  return { total: 0, ativos: 0, vencendo30: 0, vencendo60: 0, vencendo90: 0, vencidos: 0, semVigencia: 0 };
}

function computarStats(lista) {
  const s = zeroStats();
  for (const c of lista) {
    s.total++;
    switch (c.status) {
      case 'Ativo': s.ativos++; break;
      case 'Vencendo30': s.vencendo30++; s.ativos++; break;
      case 'Vencendo60': s.vencendo60++; s.ativos++; break;
      case 'Vencendo90': s.vencendo90++; s.ativos++; break;
      case 'Vencido': s.vencidos++; break;
      default: s.semVigencia++; break;
    }
  }
  return s;
}

function montarArvore(lista) {
  const byDir = {};
  for (const c of lista) {
    const d = c.diretoria || '(sem diretoria)';
    const u = c.unidade || 'CORPORATIVO';
    const f = c.fornecedor || c.title || '(sem fornecedor)';
    if (!byDir[d]) byDir[d] = { diretoria: d, total: 0, unidades: {} };
    if (!byDir[d].unidades[u]) byDir[d].unidades[u] = { unidade: u, total: 0, fornecedores: {} };
    if (!byDir[d].unidades[u].fornecedores[f]) byDir[d].unidades[u].fornecedores[f] = { fornecedor: f, contratos: [] };
    byDir[d].unidades[u].fornecedores[f].contratos.push(c);
    byDir[d].total++;
    byDir[d].unidades[u].total++;
  }
  // Converte pra array ordenado
  const out = Object.values(byDir).sort(function(a,b){ return String(a.diretoria).localeCompare(b.diretoria); });
  for (const dir of out) {
    dir.unidades = Object.values(dir.unidades).sort(function(a,b){ return String(a.unidade).localeCompare(b.unidade); });
    for (const un of dir.unidades) {
      un.fornecedores = Object.values(un.fornecedores).sort(function(a,b){ return String(a.fornecedor).localeCompare(b.fornecedor); });
    }
  }
  return out;
}
