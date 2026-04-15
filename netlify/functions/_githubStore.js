
const fs = require("fs/promises");
const path = require("path");

const GITHUB_REPO = process.env.GITHUB_REPO || "";
const GITHUB_BRANCH = process.env.GITHUB_BRANCH || "main";
const GITHUB_TOKEN = process.env.GITHUB_TOKEN || "";

function isGithubConfigured() {
  return Boolean(GITHUB_REPO && GITHUB_TOKEN);
}

async function readLocalRaw(relativePath) {
  const filePath = path.join(process.cwd(), relativePath);
  return fs.readFile(filePath, "utf8");
}

async function writeLocalRaw(relativePath, content) {
  const filePath = path.join(process.cwd(), relativePath);
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
