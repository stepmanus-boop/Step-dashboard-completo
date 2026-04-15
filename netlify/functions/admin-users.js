const crypto = require("crypto");
const { jsonResponse, requireAdmin, hashPassword, normalizeText } = require("./_auth");
const { readJson, writeJson, isGithubConfigured } = require("./_githubStore");

exports.handler = async (event) => {
  const admin = requireAdmin(event);
  if (!admin.ok) return admin.response;

  if (event.httpMethod === "GET") {
    const users = await readJson("data/users.json", []);
    return jsonResponse(200, {
      ok: true,
      githubSyncEnabled: isGithubConfigured(),
      users: users.map((user) => ({
        id: user.id,
        name: user.name,
        username: user.username,
        role: user.role,
        sector: user.sector,
        active: Boolean(user.active),
        createdAt: user.createdAt || null,
      })),
    });
  }

  if (event.httpMethod === "PATCH") {
    try {
      const body = JSON.parse(event.body || "{}");
      const userId = String(body.userId || "").trim();
      const nextRole = body.role === "admin" ? "admin" : "sector";
      if (!userId) {
        return jsonResponse(400, { ok: false, error: "Usuário não informado." });
      }
      const users = await readJson("data/users.json", []);
      const index = users.findIndex((user) => user.id === userId);
      if (index < 0) {
        return jsonResponse(404, { ok: false, error: "Usuário não encontrado." });
      }
      if (users[index].id === admin.session.sub && nextRole !== "admin") {
        return jsonResponse(400, { ok: false, error: "O admin atual não pode remover o próprio acesso." });
      }
      users[index].role = nextRole;
      users[index].sector = nextRole === "admin" ? "all" : (users[index].sector && users[index].sector !== "all" ? users[index].sector : "producao");
      await writeJson("data/users.json", users, `chore: atualiza perfil do usuário ${users[index].username}`);
      return jsonResponse(200, { ok: true, user: { id: users[index].id, role: users[index].role, sector: users[index].sector } });
    } catch (error) {
      return jsonResponse(500, { ok: false, error: error.message || "Falha ao atualizar usuário." });
    }
  }

  if (event.httpMethod !== "POST") {
    return jsonResponse(405, { ok: false, error: "Método não permitido." });
  }

  try {
    const body = JSON.parse(event.body || "{}");
    const name = String(body.name || "").trim();
    const username = String(body.username || "").trim();
    const password = String(body.password || "");
    const role = body.role === "admin" ? "admin" : "sector";
    const sector = role === "admin" ? "all" : String(body.sector || "").trim();

    if (!name || !username || !password) {
      return jsonResponse(400, { ok: false, error: "Preencha nome, usuário e senha." });
    }

    if (role !== "admin" && !sector) {
      return jsonResponse(400, { ok: false, error: "Selecione o setor do usuário." });
    }

    const users = await readJson("data/users.json", []);
    const exists = users.some((user) => normalizeText(user.username) === normalizeText(username));
    if (exists) {
      return jsonResponse(409, { ok: false, error: "Já existe um usuário com esse login." });
    }

    const nextUser = {
      id: `u_${crypto.randomBytes(6).toString("hex")}`,
      name,
      username,
      passwordHash: hashPassword(password),
      role,
      sector,
      active: true,
      createdAt: new Date().toISOString(),
    };

    users.push(nextUser);
    await writeJson("data/users.json", users, `feat: adiciona usuário ${username}`);

    return jsonResponse(200, {
      ok: true,
      user: {
        id: nextUser.id,
        name: nextUser.name,
        username: nextUser.username,
        role: nextUser.role,
        sector: nextUser.sector,
        active: nextUser.active,
      },
    });
  } catch (error) {
    return jsonResponse(500, { ok: false, error: error.message || "Falha ao criar usuário." });
  }
};
