/**
 * Sistema de Aprovacao de NF - ConsultarCNPJ
 *
 * Consulta dados publicos de um CNPJ via BrasilAPI (gratuito, sem token).
 * Fallback: ReceitaWS (caso BrasilAPI falhe).
 *
 * Uso: GET /api/ConsultarCNPJ?cnpj=00000000000000
 *
 * Retorna JSON normalizado:
 *   {
 *     cnpj: '00.000.000/0001-00',
 *     razaoSocial: 'EMPRESA EXEMPLO LTDA',
 *     nomeFantasia: 'EXEMPLO',
 *     situacao: 'ATIVA',
 *     uf: 'SP',
 *     municipio: 'SAO PAULO',
 *     endereco: 'RUA EXEMPLO, 123',
 *     bairro: 'CENTRO',
 *     cep: '01000-000',
 *     atividadePrincipal: 'Comercio varejista',
 *     telefone: '...',
 *     email: '...',
 *     source: 'brasilapi' | 'receitaws'
 *   }
 */

require('isomorphic-fetch');

function onlyDigits(s) {
  return String(s || '').replace(/\D/g, '');
}

function formatCNPJ(cnpj) {
  const d = onlyDigits(cnpj);
  if (d.length !== 14) return cnpj;
  return `${d.slice(0,2)}.${d.slice(2,5)}.${d.slice(5,8)}/${d.slice(8,12)}-${d.slice(12,14)}`;
}

function isValidCNPJ(cnpj) {
  const d = onlyDigits(cnpj);
  if (d.length !== 14) return false;
  if (/^(\d)\1{13}$/.test(d)) return false; // todos digitos iguais
  // Validacao dos digitos verificadores
  const calc = (base) => {
    let pos = base.length - 7;
    let sum = 0;
    for (let i = base.length; i >= 1; i--) {
      sum += parseInt(base[base.length - i]) * pos--;
      if (pos < 2) pos = 9;
    }
    const r = sum % 11;
    return r < 2 ? 0 : 11 - r;
  };
  const dig1 = calc(d.slice(0,12));
  const dig2 = calc(d.slice(0,12) + dig1);
  return dig1 === parseInt(d[12]) && dig2 === parseInt(d[13]);
}

async function consultarBrasilAPI(cnpjDigits) {
  const url = `https://brasilapi.com.br/api/cnpj/v1/${cnpjDigits}`;
  const r = await fetch(url, {
    headers: { 'Accept': 'application/json', 'User-Agent': 'Pronep-NF-System/1.0' }
  });
  if (!r.ok) throw new Error(`BrasilAPI HTTP ${r.status}`);
  const data = await r.json();
  return {
    cnpj: formatCNPJ(data.cnpj),
    razaoSocial: data.razao_social || '',
    nomeFantasia: data.nome_fantasia || data.razao_social || '',
    situacao: data.descricao_situacao_cadastral || '',
    uf: data.uf || '',
    municipio: data.municipio || '',
    endereco: [data.descricao_tipo_de_logradouro, data.logradouro, data.numero, data.complemento]
      .filter(Boolean).join(' ').trim(),
    bairro: data.bairro || '',
    cep: data.cep ? String(data.cep).replace(/(\d{5})(\d{3})/, '$1-$2') : '',
    atividadePrincipal: (data.cnae_fiscal_descricao || '').trim(),
    telefone: data.ddd_telefone_1 || '',
    email: data.email || '',
    source: 'brasilapi'
  };
}

async function consultarReceitaWS(cnpjDigits) {
  const url = `https://www.receitaws.com.br/v1/cnpj/${cnpjDigits}`;
  const r = await fetch(url, {
    headers: { 'Accept': 'application/json', 'User-Agent': 'Pronep-NF-System/1.0' }
  });
  if (!r.ok) throw new Error(`ReceitaWS HTTP ${r.status}`);
  const data = await r.json();
  if (data.status === 'ERROR') throw new Error(data.message || 'CNPJ nao encontrado');
  return {
    cnpj: data.cnpj || formatCNPJ(cnpjDigits),
    razaoSocial: data.nome || '',
    nomeFantasia: data.fantasia || data.nome || '',
    situacao: data.situacao || '',
    uf: data.uf || '',
    municipio: data.municipio || '',
    endereco: [data.logradouro, data.numero, data.complemento].filter(Boolean).join(', ').trim(),
    bairro: data.bairro || '',
    cep: data.cep || '',
    atividadePrincipal: (data.atividade_principal && data.atividade_principal[0] && data.atividade_principal[0].text) || '',
    telefone: data.telefone || '',
    email: data.email || '',
    source: 'receitaws'
  };
}

module.exports = async function (context, req) {
  try {
    const cnpjRaw = (req.query && req.query.cnpj) || (req.body && req.body.cnpj);
    if (!cnpjRaw) {
      context.res = {
        status: 400,
        headers: { 'Content-Type': 'application/json' },
        body: { error: 'Parametro cnpj obrigatorio. Use ?cnpj=00000000000000' }
      };
      return;
    }

    const cnpjDigits = onlyDigits(cnpjRaw);
    if (cnpjDigits.length !== 14) {
      context.res = {
        status: 400,
        headers: { 'Content-Type': 'application/json' },
        body: { error: 'CNPJ deve ter 14 digitos.', cnpj: cnpjRaw }
      };
      return;
    }

    if (!isValidCNPJ(cnpjDigits)) {
      context.res = {
        status: 400,
        headers: { 'Content-Type': 'application/json' },
        body: { error: 'CNPJ invalido (digitos verificadores nao conferem).', cnpj: formatCNPJ(cnpjDigits) }
      };
      return;
    }

    // Tenta BrasilAPI primeiro
    let result;
    let primaryError = null;
    try {
      result = await consultarBrasilAPI(cnpjDigits);
    } catch (err) {
      primaryError = err.message;
      context.log && context.log.warn && context.log.warn(`BrasilAPI falhou: ${err.message}, tentando ReceitaWS`);
      // Fallback: ReceitaWS
      try {
        result = await consultarReceitaWS(cnpjDigits);
      } catch (err2) {
        context.res = {
          status: 502,
          headers: { 'Content-Type': 'application/json' },
          body: {
            error: 'Nao foi possivel consultar o CNPJ em nenhuma fonte.',
            brasilapi: primaryError,
            receitaws: err2.message,
            cnpj: formatCNPJ(cnpjDigits)
          }
        };
        return;
      }
    }

    context.res = {
      status: 200,
      headers: { 'Content-Type': 'application/json', 'Cache-Control': 'public, max-age=86400' },
      body: result
    };
  } catch (err) {
    context.log && context.log.error && context.log.error('ConsultarCNPJ error:', err);
    context.res = {
      status: 500,
      headers: { 'Content-Type': 'application/json' },
      body: { error: (err && err.message) || String(err) }
    };
  }
};
