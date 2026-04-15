const { jsonResponse, clearSessionCookie } = require("./_auth");

exports.handler = async () =>
  jsonResponse(200, { ok: true }, { headers: { "set-cookie": clearSessionCookie() } });
