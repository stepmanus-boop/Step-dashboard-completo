const fs = require("fs/promises");
const path = require("path");

const ENV_GITHUB_REPO = process.env.GITHUB_REPO || "";
const ENV_GITHUB_BRANCH = process.env.GITHUB_BRANCH || "main";
const ENV_GITHUB_TOKEN = process.env.GITHUB_TOKEN || "";
const LOCAL_DATA_ROOT = process.env.LOCAL_DATA_ROOT || "/tmp/step-gerencia-data";
const GITHUB_CONFIG_PATH = "data/github-config.json";

function resolveProjectPath(relativePath) {
  return path.resolve(__dirname, "..", "..", relativePath);
}

function resolveTempPath(relativePath) {
  return path.resolve(LOCAL_DATA_ROOT, relativePath.replace(/^data\//, ""));
}

async function pathExists(filePath) {
  try {
    await fs.access(filePath);
    return true;
  } catch (_) {
    return false;
  }
}

async function resolveBundledLocalPath(relativePath) {
  const candidates = [
    resolveProjectPath(relativePath),
    resolveProjectPath(relativePath.replace(/^data\//, "netlify/data/")),
  ];
  for (const filePath of candidates) {
    if (await pathExists(filePath)) {
      return filePath;
    }
  }
  return candidates[candidates.length - 1];
}

async function ensureTempSeed(relativePath) {
  const tempPath = resolveTempPath(relativePath);
  if (await pathExists(tempPath)) {
    return tempPath;
  }

  await fs.mkdir(path.dirname(tempPath), { recursive: true });

  const bundledPath = await resolveBundledLocalPath(relativePath);
  if (await pathExists(bundledPath)) {
    const raw = await fs.readFile(bundledPath, "utf8");
    await fs.writeFile(tempPath, raw, "utf8");
  }

  return tempPath;
}

async function readLocalRaw(relativePath) {
  const tempPath = await ensureTempSeed(relativePath);
  if (await pathExists(tempPath)) {
    return fs.readFile(tempPath, "utf8");
  }

  const bundledPath = await resolveBundledLocalPath(relativePath);
  return fs.readFile(bundledPath, "utf8");
}

async function writeLocalRaw(relativePath, content) {
  const filePath = await ensureTempSeed(relativePath);
  await fs.mkdir(path.dirname(filePath), { recursive: true });
  await fs.writeFile(filePath, content, "utf8");
  return { mode: "local" };
}

async function readLocalJson(relativePath, fallbackValue = []) {
  try {
    const raw = await readLocalRaw(relativePath);
    return JSON.parse(raw || JSON.stringify(fallbackValue));
  } catch (error) {
    if (String(error.message || "").includes("ENOENT")) {
      return fallbackValue;
    }
    throw error;
  }
}

async function getGithubConfig() {
  if (ENV_GITHUB_REPO && ENV_GITHUB_TOKEN) {
    return {
      repo: ENV_GITHUB_REPO,
      branch: ENV_GITHUB_BRANCH || "main",
      token: ENV_GITHUB_TOKEN,
      source: "env",
    };
  }

  try {
    const saved = await readLocalJson(GITHUB_CONFIG_PATH, {});
    if (saved?.repo && saved?.token) {
      return {
        repo: String(saved.repo),
        branch: String(saved.branch || "main"),
        token: String(saved.token),
        source: "local",
      };
    }
  } catch (_) {}

  return { repo: "", branch: "main", token: "", source: "none" };
}

async function saveGithubConfig(config = {}) {
  const payload = {
    repo: String(config.repo || "").trim(),
    branch: String(config.branch || "main").trim() || "main",
    token: String(config.token || "").trim(),
    updatedAt: new Date().toISOString(),
  };
  await writeLocalRaw(GITHUB_CONFIG_PATH, JSON.stringify(payload, null, 2));
  return payload;
}

async function clearGithubConfig() {
  await writeLocalRaw(GITHUB_CONFIG_PATH, JSON.stringify({}, null, 2));
  return { ok: true };
}

async function isGithubConfigured() {
  const cfg = await getGithubConfig();
  return Boolean(cfg.repo && cfg.token);
}

async function githubFetch(url, options = {}) {
  const cfg = await getGithubConfig();
  if (!cfg.repo || !cfg.token) {
    throw new Error("GitHub não configurado.");
  }

  const response = await fetch(url, {
    ...options,
    headers: {
      Authorization: `Bearer ${cfg.token}`,
      Accept: "application/vnd.github+json",
      "Content-Type": "application/json",
      ...(options.headers || {}),
    },
  });

  if (!response.ok) {
    const body = await response.text();
    throw new Error(`GitHub ${response.status}: ${body}`);
  }

  return response.json();
}

async function readGithubFile(relativePath) {
  const cfg = await getGithubConfig();
  const url = `https://api.github.com/repos/${cfg.repo}/contents/${relativePath}?ref=${encodeURIComponent(cfg.branch)}`;
  const payload = await githubFetch(url);
  const content = Buffer.from(String(payload.content || "").replace(/\n/g, ""), "base64").toString("utf8");
  return { content, sha: payload.sha, mode: "github" };
}

async function writeGithubFile(relativePath, content, message) {
  const cfg = await getGithubConfig();
  let current = null;
  try {
    current = await readGithubFile(relativePath);
  } catch (error) {
    if (!String(error.message || "").includes("GitHub 404")) {
      throw error;
    }
  }
  const url = `https://api.github.com/repos/${cfg.repo}/contents/${relativePath}`;
  await githubFetch(url, {
    method: "PUT",
    body: JSON.stringify({
      message,
      content: Buffer.from(content, "utf8").toString("base64"),
      ...(current?.sha ? { sha: current.sha } : {}),
      branch: cfg.branch,
    }),
  });
  return { mode: "github" };
}

async function readJson(relativePath, fallbackValue = []) {
  try {
    const raw = await isGithubConfigured()
      ? (async () => (await readGithubFile(relativePath)).content)()
      : readLocalRaw(relativePath);
    return JSON.parse(await raw || JSON.stringify(fallbackValue));
  } catch (error) {
    if (String(error.message || "").includes("ENOENT")) {
      return fallbackValue;
    }
    throw error;
  }
}

async function writeJson(relativePath, value, message = "chore: atualiza dados") {
  const content = JSON.stringify(value, null, 2);
  if (await isGithubConfigured()) {
    return writeGithubFile(relativePath, content, message);
  }
  return writeLocalRaw(relativePath, content);
}

async function syncLocalDataToGithub() {
  const cfg = await getGithubConfig();
  if (!cfg.repo || !cfg.token) {
    throw new Error("GitHub não configurado.");
  }
  const targets = [
    { path: "data/users.json", message: "chore: sincroniza usuários" },
    { path: "data/manual-alerts.json", message: "chore: sincroniza alertas manuais" },
    { path: "data/alert-reads.json", message: "chore: sincroniza leituras de alerta" },
  ];
  const results = [];
  for (const target of targets) {
    const raw = await readLocalRaw(target.path).catch(() => "[]");
    await writeGithubFile(target.path, raw || "[]", target.message);
    results.push(target.path);
  }
  return { ok: true, files: results };
}

module.exports = {
  readJson,
  writeJson,
  readLocalJson,
  readLocalRaw,
  writeLocalRaw,
  getGithubConfig,
  saveGithubConfig,
  clearGithubConfig,
  isGithubConfigured,
  syncLocalDataToGithub,
};
