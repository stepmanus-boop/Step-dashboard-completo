const fs = require("fs/promises");
const path = require("path");

const ENV_GITHUB_REPO = process.env.GITHUB_REPO || "";
const ENV_GITHUB_BRANCH = process.env.GITHUB_BRANCH || "main";
const ENV_GITHUB_TOKEN = process.env.GITHUB_TOKEN || "";
const LOCAL_DATA_ROOT = process.env.LOCAL_DATA_ROOT || "/tmp/step-gerencia-data";
const GITHUB_CONFIG_PATH = "config/github-sync.json";

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
    return parsePossiblyConflictedJson(raw || JSON.stringify(fallbackValue), fallbackValue, relativePath);
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
    const bundledPath = await resolveBundledLocalPath(GITHUB_CONFIG_PATH);
    if (await pathExists(bundledPath)) {
      const raw = await fs.readFile(bundledPath, "utf8");
      const saved = JSON.parse(raw || "{}");
      if (saved?.repo && saved?.token && !String(saved.token).includes("COLE_SEU_TOKEN_GITHUB_AQUI")) {
        return {
          repo: String(saved.repo),
          branch: String(saved.branch || "main"),
          token: String(saved.token),
          source: "file",
        };
      }
    }
  } catch (_) {}

  return { repo: "", branch: "main", token: "", source: "none" };
}

async function saveGithubConfig(config = {}) {
  const payload = {
    token: String(config.token || "").trim(),
    repo: String(config.repo || "").trim(),
    branch: String(config.branch || "main").trim() || "main",
    updatedAt: new Date().toISOString(),
  };
  const bundledPath = await resolveBundledLocalPath(GITHUB_CONFIG_PATH);
  await fs.mkdir(path.dirname(bundledPath), { recursive: true });
  await fs.writeFile(bundledPath, JSON.stringify(payload, null, 2), "utf8");
  return payload;
}

async function clearGithubConfig() {
  const bundledPath = await resolveBundledLocalPath(GITHUB_CONFIG_PATH);
  await fs.mkdir(path.dirname(bundledPath), { recursive: true });
  await fs.writeFile(bundledPath, JSON.stringify({}, null, 2), "utf8");
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
    return parsePossiblyConflictedJson(await raw || JSON.stringify(fallbackValue), fallbackValue, relativePath);
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

function normalizeMergeKey(value) {
  return String(value || "")
    .normalize("NFD")
    .replace(/[̀-ͯ]/g, "")
    .trim()
    .toLowerCase();
}

function isObject(value) {
  return value && typeof value === "object" && !Array.isArray(value);
}

function extractConflictBlocks(raw) {
  const text = String(raw || "");
  if (!text.includes("<<<<<<<")) return null;
  const lines = text.split(/\r?\n/);
  const blocks = [];
  let i = 0;
  while (i < lines.length) {
    if (!lines[i].startsWith("<<<<<<<")) {
      i += 1;
      continue;
    }
    i += 1;
    const head = [];
    while (i < lines.length && !lines[i].startsWith("=======")) {
      head.push(lines[i]);
      i += 1;
    }
    if (i < lines.length && lines[i].startsWith("=======")) i += 1;
    const incoming = [];
    while (i < lines.length && !lines[i].startsWith(">>>>>>>")) {
      incoming.push(lines[i]);
      i += 1;
    }
    if (i < lines.length && lines[i].startsWith(">>>>>>>")) i += 1;
    blocks.push({ head: head.join("\n").trim(), incoming: incoming.join("\n").trim() });
  }
  return blocks.length ? blocks : null;
}

function parseJsonOrNull(raw) {
  try {
    return JSON.parse(String(raw || "").trim());
  } catch (_) {
    return null;
  }
}

function mergeStructuredValues(left, right, relativePath = "") {
  if (Array.isArray(left) && Array.isArray(right)) {
    if (relativePath.includes("users")) return mergeUsers(left, right);
    return mergeById(left, right);
  }
  if (isObject(left) && isObject(right)) {
    return { ...left, ...right };
  }
  return right ?? left;
}

function parsePossiblyConflictedJson(raw, fallbackValue = [], relativePath = "") {
  const direct = parseJsonOrNull(raw);
  if (direct !== null) return direct;

  const blocks = extractConflictBlocks(raw);
  if (!blocks) {
    throw new Error(`JSON inválido em ${relativePath || "arquivo de dados"}.`);
  }

  let repaired = null;
  for (const block of blocks) {
    const left = parseJsonOrNull(block.head);
    const right = parseJsonOrNull(block.incoming);
    if (left === null && right === null) continue;
    repaired = repaired === null
      ? mergeStructuredValues(left ?? fallbackValue, right ?? fallbackValue, relativePath)
      : mergeStructuredValues(repaired, mergeStructuredValues(left ?? fallbackValue, right ?? fallbackValue, relativePath), relativePath);
  }

  if (repaired === null) {
    throw new Error(`JSON inválido em ${relativePath || "arquivo de dados"}: conflito de merge detectado.`);
  }

  return repaired;
}

function mergeUsers(localUsers = [], remoteUsers = []) {
  const merged = new Map();
  for (const user of [...remoteUsers, ...localUsers]) {
    if (!isObject(user)) continue;
    const key = normalizeMergeKey(user.id || user.username);
    if (!key) continue;
    const existing = merged.get(key) || {};
    merged.set(key, {
      ...existing,
      ...user,
      passwordHash: user.passwordHash || existing.passwordHash || "",
      active: user.active !== false,
    });
  }
  return [...merged.values()];
}

function mergeById(localItems = [], remoteItems = []) {
  const merged = new Map();
  for (const item of [...remoteItems, ...localItems]) {
    if (!isObject(item)) continue;
    const key = normalizeMergeKey(item.id);
    if (!key) continue;
    merged.set(key, { ...(merged.get(key) || {}), ...item });
  }
  return [...merged.values()];
}

async function readGithubJsonIfExists(relativePath, fallbackValue = []) {
  try {
    const payload = await readGithubFile(relativePath);
    return parsePossiblyConflictedJson(String(payload?.content || "") || JSON.stringify(fallbackValue), fallbackValue, relativePath);
  } catch (error) {
    if (String(error.message || "").includes("GitHub 404")) {
      return fallbackValue;
    }
    throw error;
  }
}

async function syncLocalDataToGithub() {
  const cfg = await getGithubConfig();
  if (!cfg.repo || !cfg.token) {
    throw new Error("GitHub não configurado. Defina GITHUB_TOKEN, GITHUB_REPO e GITHUB_BRANCH no Netlify para enviar ao repositório.");
  }

  const localUsers = await readLocalJson("data/users.json", []);
  const remoteUsers = await readGithubJsonIfExists("data/users.json", []);
  const mergedUsers = mergeUsers(localUsers, remoteUsers);

  const localAlerts = await readLocalJson("data/manual-alerts.json", []);
  const remoteAlerts = await readGithubJsonIfExists("data/manual-alerts.json", []);
  const mergedAlerts = mergeById(localAlerts, remoteAlerts);

  const localAcks = await readLocalJson("data/alert-acks.json", []);
  const remoteAcks = await readGithubJsonIfExists("data/alert-acks.json", []);
  const mergedAcks = mergeById(localAcks, remoteAcks);

  const targets = [
    { path: "data/users.json", value: mergedUsers, message: "chore: sincroniza usuários" },
    { path: "data/manual-alerts.json", value: mergedAlerts, message: "chore: sincroniza alertas manuais" },
    { path: "data/alert-acks.json", value: mergedAcks, message: "chore: sincroniza confirmações de leitura" },
  ];

  const results = [];
  for (const target of targets) {
    const raw = JSON.stringify(target.value, null, 2);
    await writeGithubFile(target.path, raw, target.message);
    const confirm = await readGithubFile(target.path);
    results.push({
      path: target.path,
      saved: Boolean(confirm && typeof confirm.content === "string"),
      size: String(confirm?.content || "").length,
    });
  }
  return { ok: true, files: results, repo: cfg.repo, branch: cfg.branch };
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
