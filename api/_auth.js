const crypto = require("crypto");

const COOKIE_NAME = "viserys_session";
const DEFAULT_USER = "jones260";
const DEFAULT_PASSWORD_SHA256 = "ed3b92471304a3ce7d70a6720001bdfa500a882cfd0d1aa8acb3500aee50c29d";
const SESSION_MAX_AGE_SECONDS = 60 * 60 * 24 * 7;

function configuredUser() {
  return process.env.APP_USER || DEFAULT_USER;
}

function verifyCredentials(username, password) {
  const expectedUser = configuredUser();
  const expectedPassword = process.env.APP_PASSWORD || "";
  const usernameOk = safeEqual(String(username || ""), expectedUser);

  if (!usernameOk) return false;

  if (expectedPassword) {
    return safeEqual(String(password || ""), expectedPassword);
  }

  return safeEqual(sha256(String(password || "")), DEFAULT_PASSWORD_SHA256);
}

function createSessionToken(username) {
  const payload = {
    u: username,
    exp: Math.floor(Date.now() / 1000) + SESSION_MAX_AGE_SECONDS,
  };
  const encoded = base64url(JSON.stringify(payload));
  return `${encoded}.${sign(encoded)}`;
}

function getUserFromRequest(req) {
  const token = parseCookies(req.headers.cookie || "")[COOKIE_NAME];
  if (!token) return null;

  const [encoded, signature] = token.split(".");
  if (!encoded || !signature || !safeEqual(signature, sign(encoded))) return null;

  try {
    const payload = JSON.parse(Buffer.from(encoded, "base64url").toString("utf8"));
    if (!payload.u || !payload.exp || payload.exp < Math.floor(Date.now() / 1000)) {
      return null;
    }
    return payload.u;
  } catch {
    return null;
  }
}

function setSessionCookie(res, username) {
  const token = createSessionToken(username);
  res.setHeader("Set-Cookie", serializeCookie(COOKIE_NAME, token, SESSION_MAX_AGE_SECONDS));
}

function clearSessionCookie(res) {
  res.setHeader("Set-Cookie", serializeCookie(COOKIE_NAME, "", 0));
}

function serializeCookie(name, value, maxAge) {
  const parts = [
    `${name}=${value}`,
    "Path=/",
    "HttpOnly",
    "SameSite=Lax",
    `Max-Age=${maxAge}`,
  ];

  if (process.env.VERCEL || process.env.NODE_ENV === "production") {
    parts.push("Secure");
  }

  return parts.join("; ");
}

function parseCookies(cookieHeader) {
  return cookieHeader.split(";").reduce((cookies, part) => {
    const index = part.indexOf("=");
    if (index < 0) return cookies;
    const key = part.slice(0, index).trim();
    const value = part.slice(index + 1).trim();
    if (key) cookies[key] = value;
    return cookies;
  }, {});
}

function sign(value) {
  return crypto.createHmac("sha256", sessionSecret()).update(value).digest("base64url");
}

function sessionSecret() {
  return process.env.APP_SESSION_SECRET || process.env.DATABASE_URL || "viserys-development-secret";
}

function sha256(value) {
  return crypto.createHash("sha256").update(value).digest("hex");
}

function safeEqual(a, b) {
  const left = Buffer.from(String(a));
  const right = Buffer.from(String(b));
  if (left.length !== right.length) return false;
  return crypto.timingSafeEqual(left, right);
}

module.exports = {
  configuredUser,
  verifyCredentials,
  getUserFromRequest,
  setSessionCookie,
  clearSessionCookie,
};
