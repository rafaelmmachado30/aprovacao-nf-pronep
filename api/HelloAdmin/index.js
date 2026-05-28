// Function de teste minima — sem requires externos
// Se ela retornar 200 mas AdminLimparBase ainda der 404, problema eh especifico
// da AdminLimparBase (algum require quebrando o boot).
// Se ela tambem der 404, problema eh limite do SWA ou config geral.
module.exports = async function (context, req) {
  context.res = {
    status: 200,
    headers: { 'Content-Type': 'application/json' },
    body: { ok: true, function: 'HelloAdmin', timestamp: new Date().toISOString() }
  };
};
