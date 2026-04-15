const { jsonResponse, requireAdmin } = require("./_auth");
const { getGithubConfig, saveGithubConfig, clearGithubConfig, syncLocalDataToGithub, isGithubConfigured } = require("./_githubStore");

function maskToken(token) {
  const value = String(token || "");
  if (!value) return "";
  if (value.length <= 8) return "********";
  return `${value.slice(0, 6)}••••${value.slice(-4)}`;
}

exports.handler = async (event) => {
  const admin = requireAdmin(event);
  if (!admin.ok) return admin.response;

  if (event.httpMethod === "GET") {
    const cfg = await getGithubConfig();
    return jsonResponse(200, {
      ok: true,
      configured: Boolean(cfg.repo && cfg.token),
      source: cfg.source,
      repo: cfg.repo || "",
      branch: cfg.branch || "main",
      tokenMasked: maskToken(cfg.token),
    });
  }

  if (event.httpMethod === "POST") {
    try {
      const body = JSON.parse(event.body || "{}");
      const action = String(body.action || "save");
      if (action === "clear") {
        await clearGithubConfig();
        return jsonResponse(200, { ok: true, message: "Configuração local do GitHub removida." });
      }
      if (action === "sync") {
        const synced = await syncLocalDataToGithub();
        const cfg = await getGithubConfig();
        return jsonResponse(200, {
          ok: true,
          message: "Sincronização com GitHub concluída com sucesso.",
          files: synced.files || [],
          configured: Boolean(cfg.repo && cfg.token),
          repo: cfg.repo || "",
          branch: cfg.branch || "main",
          tokenMasked: maskToken(cfg.token),
        });
      }

      const repo = String(body.repo || "").trim();
      const branch = String(body.branch || "main").trim() || "main";
      const token = String(body.token || "").trim();
      if (!repo || !branch || !token) {
        return jsonResponse(400, { ok: false, error: "Preencha token, repositório e branch." });
      }
      await saveGithubConfig({ repo, branch, token });
      const configured = await isGithubConfigured();
      return jsonResponse(200, {
        ok: true,
        configured,
        message: "Configuração do GitHub salva com sucesso.",
        repo,
        branch,
        tokenMasked: maskToken(token),
      });
    } catch (error) {
      return jsonResponse(500, { ok: false, error: error.message || "Falha ao salvar configuração do GitHub." });
    }
  }

  return jsonResponse(405, { ok: false, error: "Método não permitido." });
};
