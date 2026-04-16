const { jsonResponse, createSessionCookie, normalizeText, normalizeSectorList, verifyPassword } = require("./_auth");
const { readMergedJson } = require("./_githubStore");

exports.handler = async (event) => {
  if (event.httpMethod !== "POST") {
    return jsonResponse(405, { ok: false, error: "Método não permitido." });
  }

  try {
    const body = JSON.parse(event.body || "{}");
    const username = String(body.username || "").trim();
    const password = String(body.password || "").trim();
    if (!username || !password) {
      return jsonResponse(400, { ok: false, error: "Informe usuário e senha." });
    }

    const users = await readMergedJson("data/users.json", []);
    const defaultAdmin = {
      id: "u_admin_001",
      name: "Administrador",
      username: "admin",
      role: "admin",
      sector: "all",
      active: true,
    };
    let user = users.find((item) => normalizeText(item.username) === normalizeText(username));

    if ((!user || !user.active) && normalizeText(username) === "admin" && String(password) === "admin123") {
      user = defaultAdmin;
    }

    if (!user || !user.active || (user.passwordHash ? !verifyPassword(password, user.passwordHash) : String(password) !== "admin123")) {
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
        alertSectors: normalizeSectorList(user.sector, user.alertSectors),
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
