const { jsonResponse, createSessionCookie, normalizeText, verifyPassword } = require("./_auth");
const { readJson } = require("./_githubStore");

exports.handler = async (event) => {
  if (event.httpMethod !== "POST") {
    return jsonResponse(405, { ok: false, error: "Método não permitido." });
  }

  try {
    const { username, password } = JSON.parse(event.body || "{}");
    if (!username || !password) {
      return jsonResponse(400, { ok: false, error: "Informe usuário e senha." });
    }

    const users = await readJson("data/users.json", []);
    const user = users.find((item) => normalizeText(item.username) === normalizeText(username));

    if (!user || !user.active || !verifyPassword(password, user.passwordHash)) {
      return jsonResponse(401, { ok: false, error: "Usuário ou senha inválidos." });
    }

    return jsonResponse(200, {
      ok: true,
      user: {
        id: user.id,
        name: user.name,
        username: user.username,
        role: user.role,
        sector: user.sector,
      },
    }, {
      headers: {
        "set-cookie": createSessionCookie(user),
      },
    });
  } catch (error) {
    return jsonResponse(500, { ok: false, error: error.message || "Falha ao autenticar." });
  }
};
