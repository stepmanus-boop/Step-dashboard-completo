const crypto = require("crypto");
const { jsonResponse, requireSession, requireAdmin } = require("./_auth");
const { readJson, writeJson, isGithubConfigured } = require("./_githubStore");

function alertVisibleToUser(alert, session) {
  if (!alert || alert.active === false) return false;
  if (session.role === "admin") return true;
  return String(alert.sector || "").toLowerCase() === String(session.sector || "").toLowerCase();
}

exports.handler = async (event) => {
  if (event.httpMethod === "GET") {
    const auth = requireSession(event);
    if (!auth.ok) return auth.response;
    const alerts = await readJson("data/manual-alerts.json", []);
    const acks = await readJson("data/alert-acks.json", []);
    const visible = alerts
      .filter((alert) => alertVisibleToUser(alert, auth.session))
      .map((alert) => {
        const acknowledgements = acks
          .filter((item) => item.alertId === alert.id)
          .sort((a, b) => new Date(b.acknowledgedAt || 0).getTime() - new Date(a.acknowledgedAt || 0).getTime())
          .map((item) => ({
            id: item.id,
            userId: item.userId,
            username: item.username,
            sector: item.sector,
            acknowledgedAt: item.acknowledgedAt,
          }));
        const acked = acknowledgements.some((item) => item.userId === auth.session.sub);
        return {
          ...alert,
          acknowledged: acked,
          ackCount: acknowledgements.length,
          lastAckAt: acknowledgements[0]?.acknowledgedAt || null,
          acknowledgements: auth.session.role === "admin" ? acknowledgements : undefined,
        };
      })
      .sort((a, b) => new Date(b.createdAt || 0).getTime() - new Date(a.createdAt || 0).getTime());

    return jsonResponse(200, {
      ok: true,
      githubSyncEnabled: await isGithubConfigured(),
      alerts: visible,
    });
  }

  if (event.httpMethod === "POST") {
    const admin = requireAdmin(event);
    if (!admin.ok) return admin.response;
    try {
      const body = JSON.parse(event.body || "{}");
      const title = String(body.title || "").trim();
      const message = String(body.message || "").trim();
      const sector = String(body.sector || "").trim().toLowerCase();
      const priority = String(body.priority || "normal").trim().toLowerCase();
      const requiresAck = body.requiresAck !== false;

      if (!title || !message || !sector) {
        return jsonResponse(400, { ok: false, error: "Informe setor, título e mensagem." });
      }

      const alerts = await readJson("data/manual-alerts.json", []);
      const nextAlert = {
        id: `a_${crypto.randomBytes(6).toString("hex")}`,
        sector,
        title,
        message,
        priority: ["low", "normal", "high", "urgent"].includes(priority) ? priority : "normal",
        requiresAck,
        active: true,
        createdAt: new Date().toISOString(),
        createdBy: admin.session.username,
      };

      alerts.unshift(nextAlert);
      await writeJson("data/manual-alerts.json", alerts, `feat: adiciona alerta manual para ${sector}`);
      return jsonResponse(200, { ok: true, alert: nextAlert });
    } catch (error) {
      return jsonResponse(500, { ok: false, error: error.message || "Falha ao criar alerta." });
    }
  }

  if (event.httpMethod === "PATCH") {
    const auth = requireSession(event);
    if (!auth.ok) return auth.response;
    try {
      const body = JSON.parse(event.body || "{}");
      const alertId = String(body.alertId || "").trim();
      if (!alertId) {
        return jsonResponse(400, { ok: false, error: "Alerta não informado." });
      }

      const alerts = await readJson("data/manual-alerts.json", []);
      const acks = await readJson("data/alert-acks.json", []);
      const alert = alerts.find((item) => item.id === alertId);
      if (!alert || !alertVisibleToUser(alert, auth.session)) {
        return jsonResponse(404, { ok: false, error: "Alerta não encontrado." });
      }

      const existing = acks.find((item) => item.alertId === alertId && item.userId === auth.session.sub);
      if (!existing) {
        acks.push({
          id: `ack_${crypto.randomBytes(6).toString("hex")}`,
          alertId,
          userId: auth.session.sub,
          username: auth.session.username,
          sector: auth.session.sector,
          acknowledgedAt: new Date().toISOString(),
        });
        await writeJson("data/alert-acks.json", acks, `chore: confirma leitura do alerta ${alertId}`);
      }

      return jsonResponse(200, { ok: true });
    } catch (error) {
      return jsonResponse(500, { ok: false, error: error.message || "Falha ao confirmar alerta." });
    }
  }

  return jsonResponse(405, { ok: false, error: "Método não permitido." });
};
