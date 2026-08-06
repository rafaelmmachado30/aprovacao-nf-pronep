/**
 * shared/blobCert.js — le os certificados A1 de um container privado no Blob Storage.
 *
 * POR QUE NAO EM APP SETTING (que era o plano original):
 * o Static Web Apps limita o TOTAL das application settings a 10 KB. Cada .pfx da
 * Pronep tem ~9 KB, ou ~12 KB em base64. Nao cabe um, muito menos tres — o Azure
 * recusa com "App settings cannot total to more than 10KB". Dividir em partes nao
 * resolve, porque o teto e sobre a soma.
 *
 * O container e PRIVADO e a chave de acesso continua numa App Setting (pequena),
 * entao o controle de acesso permanece o mesmo que foi decidido: quem chega ao
 * certificado e quem tem acesso ao Azure. Nada de certificado no repositorio.
 *
 * Assinatura Shared Key feita a mao, sem @azure/storage-blob: sao ~40 linhas de
 * HMAC e evita arrastar mais uma dependencia para dentro das Functions.
 */

const crypto = require('crypto');
const https = require('https');

const CONTAINER_PADRAO = 'certificados';
const API_VERSION = '2021-08-06';

/* Cache em memoria: a Function e reaproveitada entre chamadas, e baixar 9 KB a
   cada consulta de CNPJ seria desperdicio. Se a instancia reciclar, recarrega. */
const _cache = {};

function lerConexao() {
  const cs = process.env.SEFAZ_CERT_STORAGE;
  if (!cs) throw new Error('App Setting SEFAZ_CERT_STORAGE nao configurada ' +
                           '(connection string do storage dos certificados)');
  const partes = {};
  for (const p of cs.split(';')) {
    const i = p.indexOf('=');
    if (i > 0) partes[p.slice(0, i)] = p.slice(i + 1);
  }
  if (!partes.AccountName || !partes.AccountKey) {
    throw new Error('SEFAZ_CERT_STORAGE nao tem AccountName/AccountKey');
  }
  return {
    conta: partes.AccountName,
    chave: partes.AccountKey,
    sufixo: partes.EndpointSuffix || 'core.windows.net',
    protocolo: partes.DefaultEndpointsProtocol || 'https'
  };
}

/* StringToSign do Shared Key para GET de blob. A ordem dos campos vazios importa:
   o servico monta a mesma string do outro lado e compara byte a byte. */
function assinar(conta, chave, recurso, dataISO) {
  const stringToSign = [
    'GET', '', '', '', '', '', '', '', '', '', '', '',
    'x-ms-date:' + dataISO,
    'x-ms-version:' + API_VERSION,
    '/' + conta + recurso
  ].join('\n');
  const hmac = crypto.createHmac('sha256', Buffer.from(chave, 'base64'));
  hmac.update(stringToSign, 'utf8');
  return 'SharedKey ' + conta + ':' + hmac.digest('base64');
}

function baixar(cfg, container, nome, timeoutMs) {
  return new Promise(function (resolve, reject) {
    const recurso = '/' + container + '/' + nome;
    const dataISO = new Date().toUTCString();
    const req = https.request({
      host: cfg.conta + '.blob.' + cfg.sufixo,
      path: recurso,
      method: 'GET',
      headers: {
        'x-ms-date': dataISO,
        'x-ms-version': API_VERSION,
        'Authorization': assinar(cfg.conta, cfg.chave, recurso, dataISO)
      },
      timeout: timeoutMs || 15000
    }, function (res) {
      const partes = [];
      res.on('data', function (d) { partes.push(d); });
      res.on('end', function () {
        const corpo = Buffer.concat(partes);
        if (res.statusCode === 200) return resolve(corpo);
        if (res.statusCode === 404) {
          return reject(new Error('Certificado ' + nome + ' nao existe no container "' +
                                  container + '"'));
        }
        /* O corpo do erro do Storage e XML e pode citar a assinatura; corta curto. */
        reject(new Error('Blob Storage devolveu HTTP ' + res.statusCode + ' para ' + nome +
                         ': ' + corpo.toString('utf8').slice(0, 200)));
      });
    });
    req.on('timeout', function () { req.destroy(new Error('Blob Storage nao respondeu a tempo')); });
    req.on('error', reject);
    req.end();
  });
}

/**
 * Devolve o .pfx do CNPJ como Buffer. O blob se chama <cnpj>.pfx.
 * @param {string} cnpj  so digitos
 */
async function lerPfx(cnpj) {
  const doc = String(cnpj || '').replace(/\D/g, '');
  if (_cache[doc]) return _cache[doc];
  const cfg = lerConexao();
  const container = process.env.SEFAZ_CERT_CONTAINER || CONTAINER_PADRAO;
  const buf = await baixar(cfg, container, doc + '.pfx');
  _cache[doc] = buf;
  return buf;
}

/* Usado pelo diagnostico: lista o que existe no container, sem baixar conteudo.
   Serve para dizer "o certificado desse CNPJ nao foi enviado" em vez de deixar o
   erro aparecer so na hora da consulta a SEFAZ. */
function listar(timeoutMs) {
  return new Promise(function (resolve, reject) {
    let cfg;
    try { cfg = lerConexao(); } catch (e) { return reject(e); }
    const container = process.env.SEFAZ_CERT_CONTAINER || CONTAINER_PADRAO;
    const recurso = '/' + container;
    const dataISO = new Date().toUTCString();
    /* comp/restype entram na canonicalizacao como query ordenada. */
    const canonical = recurso + '\ncomp:list\nrestype:container';
    const stringToSign = [
      'GET', '', '', '', '', '', '', '', '', '', '', '',
      'x-ms-date:' + dataISO,
      'x-ms-version:' + API_VERSION,
      '/' + cfg.conta + canonical
    ].join('\n');
    const hmac = crypto.createHmac('sha256', Buffer.from(cfg.chave, 'base64'));
    hmac.update(stringToSign, 'utf8');
    const auth = 'SharedKey ' + cfg.conta + ':' + hmac.digest('base64');

    const req = https.request({
      host: cfg.conta + '.blob.' + cfg.sufixo,
      path: recurso + '?restype=container&comp=list',
      method: 'GET',
      headers: { 'x-ms-date': dataISO, 'x-ms-version': API_VERSION, 'Authorization': auth },
      timeout: timeoutMs || 15000
    }, function (res) {
      const partes = [];
      res.on('data', function (d) { partes.push(d); });
      res.on('end', function () {
        const xml = Buffer.concat(partes).toString('utf8');
        if (res.statusCode !== 200) {
          return reject(new Error('HTTP ' + res.statusCode + ': ' + xml.slice(0, 200)));
        }
        const nomes = [];
        const re = /<Name>([^<]+)<\/Name>/g;
        let m;
        while ((m = re.exec(xml))) nomes.push(m[1]);
        resolve(nomes);
      });
    });
    req.on('timeout', function () { req.destroy(new Error('Blob Storage nao respondeu a tempo')); });
    req.on('error', reject);
    req.end();
  });
}

module.exports = { lerPfx, listar };
