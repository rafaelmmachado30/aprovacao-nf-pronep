/**
 * shared/certA1.js — abre um .pfx (A1) para conferir senha e ler a validade.
 *
 * POR QUE ISSO EXISTE
 * O autoatendimento do certificado sem esta checagem seria pior do que nao ter:
 * o upload aceitaria qualquer arquivo, e o erro apareceria semanas depois, no meio
 * de uma consulta a SEFAZ, com a mensagem do OpenSSL. Aqui o administrador descobre
 * na hora se a senha esta errada ou se o arquivo esta no formato antigo.
 *
 * COMO SE LE A VALIDADE SEM ABRIR PKCS#12
 * O Node nao expoe parser de PKCS#12 — nao ha `new X509Certificate(pfx)`. Mas ele
 * sabe USAR um pfx num handshake TLS, e `socket.getCertificate()` devolve o
 * certificado LOCAL, com valid_from/valid_to. Entao o processo abre um servidor TLS
 * em 127.0.0.1 e conecta nele mesmo: nada sai da maquina, nao precisa do openssl
 * instalado (as Functions gerenciadas do SWA nao garantem binario externo) e nao
 * entra dependencia nova.
 *
 * FORMATO ANTIGO (o erro que mais assusta)
 * A1 exportado com encriptacao legada (RC2/40 bits) faz o Node dizer apenas
 * "Unsupported PKCS12 PFX data". Nao e senha errada nem arquivo corrompido: e o
 * OpenSSL 3 recusando um algoritmo obsoleto. Aqui isso vira uma mensagem com o
 * comando de reexportacao, em vez de uma caca ao erro.
 */

const tls = require('tls');

/* OID 1.2.840.113549.1.12 (pkcs-12), em DER. Todo PKCS#12 o contem: os tipos de bag
   (keyBag, pkcs8ShroudedKeyBag, certBag) vivem embaixo dele. */
const OID_PKCS12 = Buffer.from([0x2a, 0x86, 0x48, 0x86, 0xf7, 0x0d, 0x01, 0x0c]);

/**
 * E um PKCS#12 de verdade?
 *
 * A REGRA ANTERIOR ("comeca com 0x30 0x82") ERRAVA NOS DOIS SENTIDOS, e foi ela que
 * barrou o primeiro A1 real de um cliente:
 *   · FALSO NEGATIVO — 0x30 0x82 e SEQUENCE de comprimento DEFINIDO. Certificadora
 *     emite A1 com 0x30 0x80, comprimento INDEFINIDO (BER), que o PKCS#12 permite e
 *     o OpenSSL le sem reclamar. Arquivo legitimo, recusado na porta.
 *   · FALSO POSITIVO — um .cer/.crt (so o certificado, SEM a chave privada) tambem
 *     comeca com 0x30 0x82. Passava, e a falta da chave so aparecia na conversa com
 *     a SEFAZ. E e o arquivo errado mais facil de enviar por engano.
 * Conferir o OID resolve os dois: aceita qualquer codificacao e recusa .cer.
 */
function pareceCertificado(buf) {
  return Buffer.isBuffer(buf) && buf.length > 300 && buf[0] === 0x30 &&
         buf.indexOf(OID_PKCS12) >= 0;
}

/** Traduz o erro do OpenSSL para algo que o administrador consiga agir. */
function classificar(msg) {
  const m = String(msg || '');
  if (/Unsupported PKCS12/i.test(m)) {
    return {
      causa: 'formato-legado',
      mensagem: 'O arquivo esta no formato antigo (RC2/40 bits), que o servidor nao ' +
        'abre mais. Reexporte o certificado com criptografia atual e envie de novo. ' +
        'Pelo OpenSSL: openssl pkcs12 -in antigo.pfx -out novo.pfx -export ' +
        '-keypbe AES-256-CBC -certpbe AES-256-CBC -macalg sha256'
    };
  }
  if (/mac verify failure|invalid password|wrong.*password|incorrect password/i.test(m)) {
    return { causa: 'senha', mensagem: 'Senha do certificado incorreta.' };
  }
  if (/asn1|header too long|not enough data|decode error/i.test(m)) {
    return { causa: 'arquivo', mensagem: 'O arquivo nao parece um .pfx/.p12 valido.' };
  }
  return { causa: 'outro', mensagem: 'Nao consegui abrir o certificado: ' + m };
}

/** Handshake em loopback so para ler as datas do certificado. */
function lerDatas(pfx, senha, timeoutMs) {
  return new Promise(function (resolve) {
    let servidor;
    let respondido = false;
    function terminar(r) {
      if (respondido) return;
      respondido = true;
      try { if (servidor) servidor.close(); } catch (e) { /* ja fechado */ }
      resolve(r);
    }
    const relogio = setTimeout(function () { terminar(null); }, timeoutMs || 5000);

    try {
      servidor = tls.createServer({ pfx: pfx, passphrase: senha }, function (s) { s.end(); });
    } catch (e) {
      clearTimeout(relogio);
      return terminar(null);
    }
    servidor.on('error', function () { clearTimeout(relogio); terminar(null); });

    servidor.listen(0, '127.0.0.1', function () {
      const cliente = tls.connect({
        host: '127.0.0.1',
        port: servidor.address().port,
        pfx: pfx,
        passphrase: senha,
        /* Certificado da Receita nao valida contra a CA do sistema, e nem precisa:
           o unico objetivo aqui e o handshake acontecer para haver o que ler. */
        rejectUnauthorized: false
      }, function () {
        let cert = null;
        try { cert = cliente.getCertificate(); } catch (e) { cert = null; }
        try { cliente.destroy(); } catch (e) { /* ja fechado */ }
        clearTimeout(relogio);
        terminar(cert && cert.valid_to ? cert : null);
      });
      cliente.on('error', function () { clearTimeout(relogio); terminar(null); });
    });
  });
}

function paraISO(v) {
  if (!v) return null;
  const d = new Date(v);
  if (isNaN(d.getTime())) return null;
  return d.toISOString().slice(0, 10);
}

/**
 * Confere o par (arquivo, senha) e le o que der.
 * @returns {Promise<{ok:boolean, causa?:string, mensagem?:string, titular?:string,
 *                    validadeInicio?:string, validadeFim?:string, diasRestantes?:number,
 *                    aviso?:string, bytes:number}>}
 */
async function inspecionar(pfx, senha, opcoes) {
  const o = opcoes || {};
  const bytes = Buffer.isBuffer(pfx) ? pfx.length : 0;

  if (!pareceCertificado(pfx)) {
    return { ok: false, causa: 'arquivo', bytes: bytes,
             mensagem: 'O arquivo nao parece um .pfx/.p12 (deveria comecar com 30 82).' };
  }
  if (!senha) {
    return { ok: false, causa: 'senha', bytes: bytes,
             mensagem: 'Informe a senha do certificado.' };
  }

  /* Primeiro o teste barato: createSecureContext ja falha se a senha estiver
     errada ou o formato for legado, sem abrir socket nenhum. */
  try {
    tls.createSecureContext({ pfx: pfx, passphrase: senha });
  } catch (e) {
    const c = classificar(e && e.message);
    return { ok: false, causa: c.causa, mensagem: c.mensagem, bytes: bytes };
  }

  /* Senha conferida. Agora as datas — e se nao der, segue sem elas.
     Um e-CNPJ tem uso estendido de CLIENTE, e o loopback usa o mesmo arquivo do
     lado servidor; se algum ambiente recusar isso, bloquear o upload por causa de
     uma data seria trocar um problema real por um pior. */
  const cert = await lerDatas(pfx, senha, o.timeoutMs);
  if (!cert) {
    return {
      ok: true, bytes: bytes, validadeFim: null,
      aviso: 'Senha confere, mas nao consegui ler a validade deste arquivo. ' +
             'O certificado funciona; a tela so nao vai avisar a data de vencimento.'
    };
  }

  const fim = paraISO(cert.valid_to);
  let dias = null;
  if (fim) {
    const hoje = new Date(new Date().toISOString().slice(0, 10) + 'T00:00:00Z');
    dias = Math.round((new Date(fim + 'T00:00:00Z') - hoje) / 86400000);
  }

  return {
    ok: true,
    bytes: bytes,
    titular: (cert.subject && (cert.subject.CN || cert.subject.O)) || null,
    validadeInicio: paraISO(cert.valid_from),
    validadeFim: fim,
    diasRestantes: dias
  };
}

module.exports = { inspecionar, pareceCertificado, classificar };
