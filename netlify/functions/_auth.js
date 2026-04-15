
const crypto = require("crypto");
const fs = require("fs/promises");
const path = require("path");

const SESSION_COOKIE_NAME = "step_session";
const SESSION_TTL_MS = Number(process.env.SESSION_TTL_MS || 1000 * 60 * 60 * 24 * 7);
const SESSION_SECRET = process.env.SESSION_SECRET || "step-dev-secret-change-me";

function jsonResponse(statusCode, data, extra = {}) {
  const headers = {
    "content-type": "application/json; charset=utf-8",
    "cache-control": "no-store",
    "access-control-allow-origin": "*",
    "access-control-allow-credentials": "true",
    ...extra.headers,
  };

  return {
    statusCode,
    headers,
    body: JSON.stringify(data),
  };
}

function toBase64Url(input) {
  return Buffer.from(input)
    .toString("base64")
    .replace(/\+/g, "-")
    .replace(/\//g, "_")
    .replace(/=+$/g, "");
}

function fromBase64Url(input) {
  let value = String(input || "").replace(/-/g, "+").replace(/_/g, "/");
  while (value.length % 4) value += "=";
  return Buffer.from(value, "base64").toString("utf8");
}

function signToken(payload) {
  const json = JSON.stringify(payload);
  const encoded = toBase64Url(json);
  const signature = crypto.createHmac("sha256", SESSION_SECRET).update(encoded).digest("hex");
  return `${encoded}.${signature}`;
}

function verifyToken(token) {
  if (!token || !String(token).includes(".")) return null;
  const [encoded, signature] = String(token).split(".");
  const expected = crypto.createHmac("sha256", SESSION_SECRET).update(encoded).digest("hex");
  if (!crypto.timingSafeEqual(Buffer.from(signature || ""), Buffer.from(expected))) return null;
  try {
    const payload = JSON.parse(fromBase64Url(encoded));
    if (!payload || !payload.exp || payload.exp < Date.now()) return null;
    return payload;
  } catch {
    return null;
  }
}

function createSessionCookie(user) {
  const payload = {
    sub: user.id,
    username: user.username,
    role: user.role,
    sector: user.sector,
    name: user.name,
    exp: Date.now() + SESSION_TTL_MS,
  };
  const token = signToken(payload);
  return `${SESSION_COOKIE_NAME}=${token}; Path=/; HttpOnly; SameSite=Lax; Max-Age=${Math.floor(SESSION_TTL_MS / 1000)}`;
}

function clearSessionCookie() {
  return `${SESSION_COOKIE_NAME}=; Path=/; HttpOnly; SameSite=Lax; Max-Age=0`;
}

function parseCookies(headers = {}) {
  const raw = headers.cookie || headers.Cookie || "";
  return raw.split(";").reduce((acc, part) => {
    const [key, ...rest] = part.trim().split("=");
    if (!key) return acc;
    acc[key] = rest.join("=");
    return acc;
  }, {});
}

function getSession(event) {
  const cookies = parseCookies(event?.headers || {});
  const token = cookies[SESSION_COOKIE_NAME];
  return verifyToken(token);
}

function requireSession(event) {
  const session = getSession(event);
  if (!session) {
    return { ok: false, response: jsonResponse(401, { ok: false, error: "Sessão expirada." }) };
  }
  return { ok: true, session };
}

function requireAdmin(event) {
  const result = requireSession(event);
  if (!result.ok) return result;
  if (result.session.role !== "admin") {
    return { ok: false, response: jsonResponse(403, { ok: false, error: "Acesso restrito ao administrador." }) };
  }
  return result;
}

function normalizeText(value) {
  return String(value || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .trim()
    .toLowerCase();
}

function hashPassword(password, saltHex) {
  const salt = saltHex ? Buffer.from(saltHex, "hex") : crypto.randomBytes(16);
  const derived = crypto.scryptSync(String(password), salt, 32);
  return `scrypt$${salt.toString("hex")}$${derived.toString("hex")}`;
}

function verifyPassword(password, passwordHash) {
  const [scheme, saltHex, hashHex] = String(passwordHash || "").split("$");
  if (scheme !== "scrypt" || !saltHex || !hashHex) return false;
  const derived = crypto.scryptSync(String(password), Buffer.from(saltHex, "hex"), 32).toString("hex");
  return crypto.timingSafeEqual(Buffer.from(derived, "hex"), Buffer.from(hashHex, "hex"));
}

async function readLocalJson(relativePath, fallbackValue = []) {
  const filePath = path.join(process.cwd(), relativePath);
  const raw = await fs.readFile(filePath, "utf8");
  return JSON.parse(raw || JSON.stringify(fallbackValue));
}

module.exports = {
  SESSION_COOKIE_NAME,
  jsonResponse,
  createSessionCookie,
  clearSessionCookie,
  getSession,
  requireSession,
  requireAdmin,
  normalizeText,
  hashPassword,
  verifyPassword,
  readLocalJson,
};
