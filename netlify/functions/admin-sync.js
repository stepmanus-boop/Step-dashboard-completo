
const { jsonResponse, requireAdmin, normalizeSectorList } = require("./_auth");
const { isGithubConfigured, writeJson } = require("./_githubStore");

exports.handler = async (event) => {
  const admin = requireAdmin(event);
  if (!admin.ok) return admin.response;

  if (event.httpMethod !== "POST") {
    return jsonResponse(405, { ok: false, error: "Método não permitido." });
  }

  if (!(await isGithubConfigured())) {
    return jsonResponse(400, { ok: false, error: "Configure GITHUB_TOKEN, GITHUB_REPO e GITHUB_BRANCH no Netlify antes de sincronizar." });
  }

  try {
    const body = JSON.parse(event.body || "{}");
    const users = Array.isArray(body.users) ? body.users : [];
    const alerts = Array.isArray(body.alerts) ? body.alerts : [];

    const sanitizedUsers = users.map((user) => ({
      id: String(user.id || "").trim(),
      name: String(user.name || "").trim(),
      username: String(user.username || "").trim(),
      passwordHash: String(user.passwordHash || ""),
      role: user.role === "admin" ? "admin" : "sector",
      sector: user.role === "admin" ? "all" : String(user.sector || "producao").trim(),
      alertSectors: user.role === "admin" ? [] : normalizeSectorList(user.sector || "producao", user.alertSectors),
      active: user.active !== false,
      createdAt: user.createdAt || new Date().toISOString(),
    })).filter((user) => user.id && user.name && user.username);

    const sanitizedAlerts = alerts.map((alert) => ({
      id: String(alert.id || "").trim(),
      sector: String(alert.sector || "").trim(),
      title: String(alert.title || "").trim(),
      message: String(alert.message || "").trim(),
      priority: String(alert.priority || "normal").trim(),
      requiresAck: Boolean(alert.requiresAck),
      acknowledgedBy: Array.isArray(alert.acknowledgedBy) ? alert.acknowledgedBy : [],
      createdAt: alert.createdAt || new Date().toISOString(),
      createdBy: alert.createdBy || admin.session.username,
      active: alert.active !== false,
    })).filter((alert) => alert.id && alert.sector && alert.title);

    await writeJson("data/users.json", sanitizedUsers, "chore: sincroniza usuários locais");
    await writeJson("data/manual-alerts.json", sanitizedAlerts, "chore: sincroniza alertas locais");

    return jsonResponse(200, {
      ok: true,
      githubSyncEnabled: true,
      users: sanitizedUsers.length,
      alerts: sanitizedAlerts.length,
    });
  } catch (error) {
    return jsonResponse(500, { ok: false, error: error.message || "Falha ao sincronizar com o GitHub." });
  }
};
