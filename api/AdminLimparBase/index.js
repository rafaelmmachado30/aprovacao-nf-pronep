// VERSAO MINIMA pra debugar 404
// Se essa retornar 200, o problema eh um dos requires/codigo removido
module.exports = async function (context, req) {
  context.res = {
    status: 200,
    headers: { 'Content-Type': 'application/json' },
    body: {
      ok: true,
      function: 'AdminLimparBase',
      version: 'minimal',
      method: req.method,
      timestamp: new Date().toISOString()
    }
  };
};
