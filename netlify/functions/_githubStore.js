
const fs = require("fs/promises");
const path = require("path");

const GITHUB_REPO = process.env.GITHUB_REPO || "";
const GITHUB_BRANCH = process.env.GITHUB_BRANCH || "main";
const GITHUB_TOKEN = process.env.GITHUB_TOKEN || "";
const LOCAL_DATA_ROOT = process.env.LOCAL_DATA_ROOT || "/tmp/step-gerencia-data";

function resolveProjectPath(relativePath) {
  return path.resolve(__dirname, "..", "..", relativePath);
}

function resolveTempPath(relativePath) {
  return path.resolve(LOCAL_DATA_ROOT, relativePath.replace(/^data\//, ""));
}

function isGithubConfigured() {
  return Boolean(GITHUB_REPO && GITHUB_TOKEN);
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

async function githubFetch(url, options = {}) {
  const response = await fetch(url, {
    ...options,
    headers: {
      Authorization: `Bearer ${GITHUB_TOKEN}`,
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
  const url = `https://api.github.com/repos/${GITHUB_REPO}/contents/${relativePath}?ref=${encodeURIComponent(GITHUB_BRANCH)}`;
  const payload = await githubFetch(url);
  const content = Buffer.from(String(payload.content || "").replace(/\n/g, ""), "base64").toString("utf8");
  return { content, sha: payload.sha, mode: "github" };
}

async function writeGithubFile(relativePath, content, message) {
  const current = await readGithubFile(relativePath);
  const url = `https://api.github.com/repos/${GITHUB_REPO}/contents/${relativePath}`;
  await githubFetch(url, {
    method: "PUT",
    body: JSON.stringify({
      message,
      content: Buffer.from(content, "utf8").toString("base64"),
      sha: current.sha,
      branch: GITHUB_BRANCH,
    }),
  });
  return { mode: "github" };
}

async function readJson(relativePath, fallbackValue = []) {
  try {
    const raw = isGithubConfigured()
      ? (await readGithubFile(relativePath)).content
      : await readLocalRaw(relativePath);
    return JSON.parse(raw || JSON.stringify(fallbackValue));
  } catch (error) {
    if (String(error.message || "").includes("ENOENT")) {
      return fallbackValue;
    }
    throw error;
  }
}

async function writeJson(relativePath, data, message) {
  const content = `${JSON.stringify(data, null, 2)}\n`;
  if (isGithubConfigured()) {
    return writeGithubFile(relativePath, content, message);
  }
  return writeLocalRaw(relativePath, content);
}

module.exports = {
  isGithubConfigured,
  readJson,
  writeJson,
};
