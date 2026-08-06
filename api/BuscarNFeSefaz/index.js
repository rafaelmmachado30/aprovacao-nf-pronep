/**
 * /api/BuscarNFeSefaz  (GET|POST) — cron/admin. Busca NF-e emitidas contra nossos CNPJs.
 *
 * Auth: X-Automacao-Secret == App Setting SEFAZ_SECRET (cron) OU sessao admin.
 * Mesmo padrao dos crons que ja existem (Aquecer, AlertaDiario, VarrerEmailsNF).
 *
 * Query:
 *   ?cnpj=00092929000198   so essa filial (default: todas as de SEFAZ_CNPJS)
 *   ?dryRun=1              CONSULTA a SEFAZ mas NAO grava nem avanca o ponteiro
 *   ?lotes=3               quantos lotes por CNPJ nesta execucao (default 3)
 *   ?ambiente=homologacao  default producao
 *
 * ORDEM QUE NAO PODE MUDAR:
 *   consultar -> GRAVAR os documentos -> SO ENTAO avancar o ponteiro
 * Avancar o ponteiro antes de gravar perde o lote para sempre: a SEFAZ so guarda
 * 90 dias e nao reentrega o que ja foi marcado como distribuido.
 *
 * TIMEOUT: a Function do Static Web Apps tem ~45s. Cada lote traz ate 50
 * documentos e cada um vira uma escrita no SharePoint. Por isso o teto de lotes
 * por execucao e a parada por tempo — o ponteiro guarda onde parou e a proxima
 * execucao continua. Nada se perde por interrupcao.
 */

require('isomorphic-fetch');
const { getGraphClient, resolveSiteId } = require('../shared/graph');
const {
  lerCnpjsConfigurados, lerCertificado, lerBase64Certificado, lerConfigSefaz, garantirPonteiro, gravarPonteiro,
  gravarDocumento, soDigitos
} = require('../shared/documentosFiscais');
const { consultarLote, extrairNFe } = require('../shared/sefaz');

/* Margem para fechar a resposta antes de a plataforma cortar em ~45s. */
const ORCAMENTO_MS = 32000;

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

module.exports = async function (context, req) {
  const t0 = Date.now();
  const q = (req.query || {});
  const dryRun = q.dryRun === '1' || q.dryRun === 'true';
  const maxLotes = Math.max(1, Math.min(20, parseInt(q.lotes, 10) || 3));
  const ambiente = q.ambiente === 'homologacao' ? 'homologacao' : 'producao';

  const diag = {
    step: 'init', dryRun: dryRun, ambiente: ambiente,
    cnpjs: [], avisos: [], timeMs: 0
  };

  try {
    if (!(await autorizado(req))) {
      context.res = { status: 403, headers: { 'Content-Type': 'application/json' },
        body: { error: 'Nao autorizado' } };
      return;
    }

    diag.step = 'config';
    let alvos = lerCnpjsConfigurados();
    if (q.cnpj) {
      const filtro = soDigitos(q.cnpj);
      alvos = alvos.filter(function (c) { return c.cnpj === filtro; });
      if (!alvos.length) throw new Error('CNPJ ' + filtro + ' nao esta em SEFAZ_CNPJS');
    }
    if (!alvos.length) throw new Error('SEFAZ_CNPJS vazio ou nao configurado');

    /* ?diagCert=1 — checa SO o certificado. Nao fala com a SEFAZ, nao le o
       SharePoint, nao grava nada. Existe porque um erro de TLS ("not enough data",
       "mac verify failure") nao diz SE o problema e o arquivo ou a senha, e sem
       essa separacao a investigacao vira tentativa e erro em producao.
       NAO EXPOE SEGREDO: reporta tamanho, cabecalho ASN.1 e o erro do OpenSSL —
       nunca o conteudo do .pfx nem a senha. */
    if (q.diagCert === '1' || q.diagCert === 'true') {
      const tls = require('tls');
      diag.step = 'diag_cert';
      for (const alvo of alvos) {
        const bloco = { cnpj: alvo.cnpj, apelido: alvo.apelido };
        let b64 = '';
        try {
          const lido = lerBase64Certificado(alvo.cnpj);
          b64 = lido.b64;
          bloco.partes = lido.partes;
        } catch (eSeq) {
          bloco.veredito = eSeq.message; diag.cnpjs.push(bloco); continue;
        }
        /* Sobre a senha reportamos FORMA, nunca conteudo: se existe, quantos
           caracteres tem e se veio com espaco nas pontas. Espaco invisivel colado
           junto da senha e erro classico de portal, e sozinho ele produz o mesmo
           "mac verify failure" de uma senha simplesmente errada — sem distinguir
           os dois, a investigacao vira tentativa e erro. */
        const _senha = process.env['SEFAZ_CERT_' + alvo.cnpj + '_SENHA'];
        bloco.senhaDefinida = !!_senha;
        if (_senha) {
          bloco.senhaTamanho = _senha.length;
          bloco.senhaComEspacoNasPontas = _senha !== _senha.trim();
        }

        /* O arquivo vem do Blob Storage; App Setting so como fallback de teste. */
        let buf;
        if (b64) {
          bloco.origem = 'appsetting';
          bloco.base64Chars = b64.length;
          buf = Buffer.from(b64, 'base64');
        } else {
          bloco.origem = 'blob';
          try {
            buf = await require('../shared/blobCert').lerPfx(alvo.cnpj);
          } catch (eBlob) {
            bloco.veredito = eBlob.message; diag.cnpjs.push(bloco); continue;
          }
        }
        bloco.bytes = buf.length;
        /* PKCS#12 e um SEQUENCE ASN.1: sempre comeca com 0x30 0x82. Se nao comecar,
           o que esta na App Setting nao e um .pfx — nem adianta olhar a senha. */
        bloco.cabecalho = buf.slice(0, 2).toString('hex');
        bloco.pareceP12 = buf.length > 2 && buf[0] === 0x30 && buf[1] === 0x82;
        if (bloco.pareceP12) {
          /* Byte 3-4 = tamanho declarado do SEQUENCE. Comparar com o tamanho real
             detecta truncamento na colagem, que e a hipotese principal. */
          const declarado = buf.readUInt16BE(2) + 4;
          bloco.tamanhoDeclarado = declarado;
          bloco.truncado = buf.length < declarado;
        }
        try {
          tls.createSecureContext({
            pfx: buf,
            passphrase: process.env['SEFAZ_CERT_' + alvo.cnpj + '_SENHA'] || ''
          });
          bloco.veredito = 'OK — certificado e senha aceitos pelo OpenSSL';
        } catch (eTls) {
          bloco.erroOpenSSL = eTls.message;
          bloco.codigo = eTls.code || null;
          const _senhaErrada = /mac verify|invalid password|wrong password/i.test(eTls.message);
          bloco.veredito = _senhaErrada
            ? 'Arquivo valido, SENHA errada'
            : (bloco.truncado ? 'Arquivo TRUNCADO na App Setting'
                              : 'Arquivo nao e um PKCS#12 valido');
          /* Se a senha sem espaco nas pontas abre, o problema e a colagem e nao a
             senha. Vale MUITO a pena separar: sao acoes completamente diferentes.
             So testa aqui no diagnostico — em producao a senha e usada como esta,
             porque senha pode legitimamente conter espaco e aparar por conta
             propria quebraria quem tem uma assim. */
          if (_senhaErrada && bloco.senhaComEspacoNasPontas) {
            try {
              tls.createSecureContext({ pfx: buf, passphrase: (_senha || '').trim() });
              bloco.veredito = 'A SENHA ESTA CERTA, mas foi gravada com espaco nas pontas — ' +
                               'regrave a App Setting _SENHA sem o espaco';
            } catch (e2) { /* segue: senha errada mesmo */ }
          }
        }
        diag.cnpjs.push(bloco);
      }
      diag.step = 'done';
      diag.timeMs = Date.now() - t0;
      context.res = { status: 200, headers: { 'Content-Type': 'application/json' },
        body: Object.assign({ ok: true, mensagem: 'Diagnostico de certificado — nenhuma chamada a SEFAZ.' }, diag) };
      return;
    }

    diag.step = 'graph';
    const client = await getGraphClient();
    const siteId = await resolveSiteId(client);

    /* FREIO DE MAO. Desligado pela tela de Configuracoes, vale a partir da proxima
       execucao do cron. Sai ANTES de qualquer consulta: nao toca a SEFAZ, nao move
       o ponteiro. Devolve 200 porque desligar e uma decisao, nao uma falha — o cron
       nao deve acender alarme vermelho todo dia por causa disso.
       dryRun e diagCert IGNORAM o freio de proposito: sao ferramentas de
       diagnostico, e precisar religar a integracao so para investigar um problema
       seria exatamente o contrario do que o freio serve. */
    diag.step = 'chave_liga_desliga';
    const cfgSefaz = await lerConfigSefaz(client, siteId);
    diag.integracao = cfgSefaz;
    if (!cfgSefaz.habilitado && !dryRun) {
      diag.step = 'done';
      diag.timeMs = Date.now() - t0;
      context.res = {
        status: 200, headers: { 'Content-Type': 'application/json' },
        body: Object.assign({
          ok: true, desligada: true,
          mensagem: 'Integracao com a SEFAZ esta DESLIGADA em Configuracoes' +
                    (cfgSefaz.motivo ? ' — ' + cfgSefaz.motivo : '') + '. Nada foi consultado.'
        }, diag)
      };
      return;
    }

    for (const alvo of alvos) {
      const bloco = {
        cnpj: alvo.cnpj, apelido: alvo.apelido,
        lotes: 0, documentos: 0, novos: 0, repetidos: 0, ignorados: 0,
        nsuInicial: null, nsuFinal: null, maxNSU: null,
        cStat: null, motivo: null, erro: null, parouPor: null
      };
      diag.cnpjs.push(bloco);

      let cert;
      try { cert = await lerCertificado(alvo.cnpj); }
      catch (eCert) { bloco.erro = eCert.message; continue; }

      let ponteiro;
      try { ponteiro = await garantirPonteiro(client, siteId, alvo.cnpj, alvo.apelido); }
      catch (eP) { bloco.erro = 'ponteiro: ' + eP.message; continue; }

      bloco.nsuInicial = ponteiro.ultimoNSU;
      let nsuCorrente = ponteiro.ultimoNSU;
      let baixadosNesta = 0;

      for (let i = 0; i < maxLotes; i++) {
        if (Date.now() - t0 > ORCAMENTO_MS) { bloco.parouPor = 'tempo'; break; }

        let lote;
        try {
          lote = await consultarLote({
            cnpj: alvo.cnpj, ultNSU: nsuCorrente, cert: cert,
            uf: alvo.unidade || 'RJ', ambiente: ambiente, timeoutMs: 20000
          });
        } catch (eL) {
          bloco.erro = eL.message;
          break;
        }

        bloco.lotes++;
        bloco.cStat = lote.cStat;
        bloco.motivo = lote.xMotivo;
        bloco.maxNSU = lote.maxNSU;

        /* 137 = nenhum documento localizado. 138 = documentos localizados.
           Qualquer outro cStat e recusa (certificado, CNPJ, consumo indevido) e
           NAO pode virar avanco de ponteiro. */
        if (lote.cStat !== '138') {
          if (lote.cStat === '137') bloco.parouPor = 'sem_novos';
          else { bloco.erro = 'cStat ' + lote.cStat + ': ' + lote.xMotivo; }
          break;
        }

        /* 1) GRAVA primeiro. */
        for (const doc of lote.documentos) {
          bloco.documentos++;
          if (doc.erro) { bloco.ignorados++; diag.avisos.push(alvo.apelido + ' NSU ' +
            doc.nsu + ': ' + doc.erro); continue; }

          const nfe = extrairNFe(doc);
          if (!nfe) { bloco.ignorados++; continue; }   /* evento, CT-e, etc. */

          if (dryRun) { bloco.novos++; continue; }

          try {
            const r = await gravarDocumento(client, siteId, {
              chaveAcesso: nfe.chaveAcesso,
              origem: 'sefaz',
              nsu: nfe.nsu,
              numeroNF: nfe.numeroNF,
              serie: nfe.serie,
              emitenteCNPJ: nfe.emitenteCNPJ,
              emitenteNome: nfe.emitenteNome,
              valor: nfe.valor,
              dataEmissao: nfe.dataEmissao,
              dataVencimento: nfe.dataVencimento,
              cnpjDestino: nfe.cnpjDestino || alvo.cnpj
            });
            if (r.novo) { bloco.novos++; baixadosNesta++; } else bloco.repetidos++;
          } catch (eG) {
            /* Falha de gravacao NAO pode deixar o ponteiro avancar por cima deste
               documento — para o lote aqui e tenta de novo na proxima execucao. */
            bloco.erro = 'gravacao NSU ' + doc.nsu + ': ' + eG.message;
            bloco.parouPor = 'erro_gravacao';
            break;
          }
        }
        if (bloco.parouPor === 'erro_gravacao') break;

        /* 2) SO AGORA avanca. */
        nsuCorrente = lote.ultNSU;
        if (!dryRun) {
          try {
            await gravarPonteiro(client, siteId, ponteiro, {
              ultimoNSU: nsuCorrente,
              maxNSU: lote.maxNSU,
              cStat: lote.cStat,
              motivo: lote.xMotivo,
              baixadosAcumulado: baixadosNesta
            });
            ponteiro.ultimoNSU = nsuCorrente;
          } catch (eP2) {
            bloco.erro = 'ponteiro: ' + eP2.message;
            bloco.parouPor = 'erro_ponteiro';
            break;
          }
        }

        if (lote.maxNSU && lote.ultNSU >= lote.maxNSU) {
          bloco.parouPor = 'fim';
          break;
        }
      }

      bloco.nsuFinal = nsuCorrente;
      if (!bloco.parouPor && bloco.lotes >= maxLotes) bloco.parouPor = 'limite_de_lotes';
    }

    diag.step = 'done';
    diag.timeMs = Date.now() - t0;

    const totalNovos = diag.cnpjs.reduce(function (s, c) { return s + c.novos; }, 0);
    const comErro = diag.cnpjs.filter(function (c) { return c.erro; });

    context.res = {
      status: 200, headers: { 'Content-Type': 'application/json' },
      body: Object.assign({
        ok: comErro.length === 0,
        mensagem: (dryRun ? '[SIMULACAO] ' : '') + totalNovos + ' documento(s) novo(s)' +
                  (comErro.length ? ' · ' + comErro.length + ' CNPJ(s) com erro' : '')
      }, diag)
    };
  } catch (err) {
    diag.timeMs = Date.now() - t0;
    context.res = {
      status: 500, headers: { 'Content-Type': 'application/json' },
      body: Object.assign({ error: (err && err.message) || String(err) }, diag)
    };
  }
};
