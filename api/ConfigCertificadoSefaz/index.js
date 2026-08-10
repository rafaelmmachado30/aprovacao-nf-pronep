/**
 * /api/ConfigCertificadoSefaz — trocar o certificado A1 da SEFAZ pela propria tela.
 *
 * GET   devolve o estado de cada CNPJ (validade, titular, quem trocou e quando)
 * POST  grava .pfx e/ou senha    body: { cnpj, pfxBase64?, senha }
 *
 * SO ADMIN. Chega a um certificado quem administra o sistema.
 *
 * POR QUE ISSO EXISTE
 * O A1 vence a cada 12 meses. Antes, renovar exigia entrar no portal do Azure para
 * subir o blob e reescrever a App Setting da senha — ou seja, dependia de mim ou de
 * quem tem acesso ao Azure. Agora e uma tela.
 *
 * O QUE NUNCA SAI DAQUI
 * A senha nao volta em nenhuma resposta, nem mascarada, nem em log, nem na auditoria.
 * O que a tela mostra e apenas SE existe senha gravada. O certificado tambem nunca e
 * devolvido para o navegador.
 *
 * A ORDEM DAS GRAVACOES, E O QUE ACONTECE SE UMA FALHAR
 * O par (arquivo, senha) e conferido ANTES de qualquer gravacao — entao o caso comum
 * de erro (senha errada, formato antigo) nunca chega a escrever nada. Se ainda assim
 * a segunda gravacao falhar, a resposta diz exatamente o que ficou pela metade, em
 * vez de um "erro ao salvar" que deixaria o administrador sem saber se pode repetir.
 */

require('isomorphic-fetch');

const { getUser } = require('../shared/auth');
const { requireAdmin } = require('../shared/authz');
const { registrar: auditRegistrar } = require('../shared/auditLog');
const { lerCnpjsConfigurados } = require('../shared/documentosFiscais');
const blobCert = require('../shared/blobCert');
const segredos = require('../shared/segredos');
const certA1 = require('../shared/certA1');

/* ~9 KB e o tamanho de um A1 da Pronep; 64 KB e folga generosa e ainda barra o
   engano de subir um PDF ou um zip no lugar do certificado. */
const MAX_BYTES = 64 * 1024;

function json(context, status, body) {
  context.res = { status: status, headers: { 'Content-Type': 'application/json' }, body: body };
}

function soDigitos(s) { return String(s || '').replace(/\D/g, ''); }

function nomeMeta(cnpj) { return cnpj + '.meta.json'; }
function nomeSenha(cnpj) { return cnpj + '.senha'; }

async function lerMeta(cnpj) {
  try {
    const txt = await blobCert.lerTexto(nomeMeta(cnpj));
    return txt ? JSON.parse(txt) : null;
  } catch (e) {
    return null;
  }
}

function diasAte(dataISO) {
  if (!dataISO) return null;
  const hoje = new Date(new Date().toISOString().slice(0, 10) + 'T00:00:00Z');
  const fim = new Date(dataISO + 'T00:00:00Z');
  if (isNaN(fim.getTime())) return null;
  return Math.round((fim - hoje) / 86400000);
}

/* ------------------------------------------------------------------------ GET */
async function estado(context) {
  const configurados = lerCnpjsConfigurados();

  const resposta = {
    prontoParaGravar: segredos.temChave(),
    storageConfigurado: !!process.env.SEFAZ_CERT_STORAGE,
    avisos: [],
    cnpjs: []
  };

  if (!resposta.storageConfigurado) {
    resposta.avisos.push('App Setting SEFAZ_CERT_STORAGE nao configurada — sem ela nao ' +
                         'da para ler nem gravar certificado.');
  }
  if (!resposta.prontoParaGravar) {
    resposta.avisos.push('App Setting CONFIG_CRYPTO_KEY nao configurada. Sem ela a senha ' +
                         'nao pode ser guardada cifrada, e o envio fica bloqueado.');
  }
  if (!configurados.length) {
    resposta.avisos.push('Nenhum CNPJ em SEFAZ_CNPJS.');
  }

  /* Uma listagem serve para todos os CNPJs — evita um GET por arquivo. */
  let existentes = [];
  if (resposta.storageConfigurado) {
    try {
      existentes = await blobCert.listar();
    } catch (e) {
      resposta.avisos.push('Nao consegui listar o container: ' + e.message);
    }
  }

  for (const c of configurados) {
    const meta = existentes.indexOf(nomeMeta(c.cnpj)) >= 0 ? await lerMeta(c.cnpj) : null;
    const dias = meta && meta.validadeFim ? diasAte(meta.validadeFim) : null;
    resposta.cnpjs.push({
      cnpj: c.cnpj,
      apelido: c.apelido,
      unidade: c.unidade,
      temArquivo: existentes.indexOf(c.cnpj + '.pfx') >= 0,
      /* A senha pode estar na App Setting (jeito antigo) ou no blob (pela tela).
         A tela precisa distinguir: so a do blob pode ser trocada aqui. */
      temSenhaNoBlob: existentes.indexOf(nomeSenha(c.cnpj)) >= 0,
      temSenhaEmAppSetting: !!process.env['SEFAZ_CERT_' + c.cnpj + '_SENHA'],
      titular: meta ? meta.titular || null : null,
      validadeFim: meta ? meta.validadeFim || null : null,
      diasRestantes: dias,
      bytes: meta ? meta.bytes || null : null,
      atualizadoEm: meta ? meta.atualizadoEm || null : null,
      atualizadoPor: meta ? meta.atualizadoPor || null : null,
      avisoValidade: meta ? meta.avisoValidade || null : null
    });
  }

  json(context, 200, resposta);
}

/* ----------------------------------------------------------------------- POST */
async function gravarCertificado(context, req, user) {
  const body = req.body || {};
  const cnpj = soDigitos(body.cnpj);

  const configurados = lerCnpjsConfigurados();
  const alvo = configurados.find(function (c) { return c.cnpj === cnpj; });
  if (!alvo) {
    return json(context, 400, { error: 'CNPJ ' + (cnpj || '(vazio)') +
      ' nao esta em SEFAZ_CNPJS. Cadastre a filial antes de enviar o certificado.' });
  }

  const senha = String(body.senha == null ? '' : body.senha);
  if (!senha) return json(context, 400, { error: 'Informe a senha do certificado.' });

  if (!segredos.temChave()) {
    return json(context, 503, { error: 'CONFIG_CRYPTO_KEY nao configurada — a senha nao ' +
      'pode ser guardada cifrada. Configure a App Setting antes de enviar.' });
  }

  /* Sem arquivo novo = "so trocar a senha". Nesse caso a senha e conferida contra o
     certificado que JA esta gravado; salvar uma senha sem conferir seria plantar uma
     falha para a proxima consulta a SEFAZ. */
  let pfx = null;
  let arquivoNovo = false;
  if (body.pfxBase64) {
    try {
      pfx = Buffer.from(String(body.pfxBase64), 'base64');
    } catch (e) {
      return json(context, 400, { error: 'pfxBase64 nao esta em base64 valido.' });
    }
    if (pfx.length > MAX_BYTES) {
      return json(context, 413, { error: 'O arquivo tem ' + pfx.length + ' bytes. ' +
        'Um certificado A1 tem cerca de 9 KB — confira se nao subiu o arquivo errado.' });
    }
    arquivoNovo = true;
  } else {
    try {
      pfx = await blobCert.lerPfx(cnpj);
    } catch (e) {
      return json(context, 404, { error: 'Nao existe certificado gravado para este CNPJ. ' +
        'Envie o arquivo .pfx junto com a senha.' });
    }
  }

  const exame = await certA1.inspecionar(pfx, senha);
  if (!exame.ok) {
    /* 422: o pedido esta bem formado, o conteudo e que nao serve. */
    return json(context, 422, { error: exame.mensagem, causa: exame.causa });
  }

  const feito = [];
  try {
    if (arquivoNovo) {
      await blobCert.gravarPfx(cnpj, pfx);
      feito.push('arquivo');
    }
    await blobCert.gravarTexto(nomeSenha(cnpj), segredos.cifrar(senha));
    feito.push('senha');
    await blobCert.gravarTexto(nomeMeta(cnpj), JSON.stringify({
      titular: exame.titular || null,
      validadeInicio: exame.validadeInicio || null,
      validadeFim: exame.validadeFim || null,
      avisoValidade: exame.aviso || null,
      bytes: exame.bytes,
      atualizadoEm: new Date().toISOString(),
      atualizadoPor: (user && (user.email || user.name)) || 'desconhecido'
    }, null, 2), 'application/json; charset=utf-8');
    feito.push('meta');
  } catch (e) {
    /* Diz o que ficou pela metade. "Erro ao salvar" deixaria o administrador sem
       saber se pode repetir — e aqui repetir e sempre seguro. */
    const parcial = feito.length
      ? 'Ja tinha gravado: ' + feito.join(', ') + '. '
      : 'Nada foi gravado. ';
    return json(context, 502, {
      error: parcial + 'Falhou em seguida: ' + e.message +
             ' Enviar de novo o mesmo arquivo e a mesma senha resolve.',
      gravado: feito
    });
  }

  /* Auditoria: o QUE mudou, nunca a senha. A assinatura e
     registrar(user, acao, objeto, resultado, detalhes). */
  try {
    await auditRegistrar(user, 'sefaz.certificado.atualizado',
      { tipo: 'certificado', id: cnpj },
      'sucesso',
      {
        cnpj: cnpj,
        apelido: alvo.apelido,
        arquivoTrocado: arquivoNovo,
        titular: exame.titular || null,
        validadeFim: exame.validadeFim || null,
        bytes: exame.bytes
      });
  } catch (e) {
    /* Auditoria e importante, mas nao desfaz uma troca que deu certo. */
    context.log('ConfigCertificadoSefaz: falha ao auditar: ' + (e && e.message));
  }

  json(context, 200, {
    ok: true,
    cnpj: cnpj,
    apelido: alvo.apelido,
    arquivoTrocado: arquivoNovo,
    titular: exame.titular || null,
    validadeFim: exame.validadeFim || null,
    diasRestantes: exame.diasRestantes == null ? null : exame.diasRestantes,
    aviso: exame.aviso || null,
    mensagem: arquivoNovo
      ? 'Certificado e senha atualizados.'
      : 'Senha atualizada (o arquivo continua o mesmo).'
  });
}

module.exports = async function (context, req) {
  try {
    const authz = await requireAdmin(context, req);
    if (!authz) return;
    const user = getUser(req);

    if (req.method === 'GET') return await estado(context);
    if (req.method === 'POST') return await gravarCertificado(context, req, user);
    json(context, 405, { error: 'Metodo nao suportado' });
  } catch (e) {
    context.log('ConfigCertificadoSefaz erro: ' + (e && e.stack || e));
    json(context, 500, { error: (e && e.message) || 'Erro inesperado' });
  }
};
