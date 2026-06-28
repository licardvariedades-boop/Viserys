const {
  configuredUser,
  verifyCredentials,
  getUserFromRequest,
  setSessionCookie,
  clearSessionCookie,
} = require("./_auth");

module.exports = async function handler(req, res) {
  setJsonHeaders(res);

  if (req.method === "GET") {
    const user = getUserFromRequest(req);
    if (!user) return send(res, 200, { ok: false, authenticated: false });
    return send(res, 200, { ok: true, user });
  }

  if (req.method === "POST" && req.query.action === "logout") {
    clearSessionCookie(res);
    return send(res, 200, { ok: true });
  }

  if (req.method === "POST") {
    const body = await readBody(req);
    const username = String(body.username || "").trim();
    const password = String(body.password || "");

    if (!verifyCredentials(username, password)) {
      return send(res, 401, { ok: false, error: "Usuário ou senha inválidos." });
    }

    setSessionCookie(res, configuredUser());
    return send(res, 200, { ok: true, user: configuredUser() });
  }

  return send(res, 405, { ok: false, error: "Método não permitido." });
};

function setJsonHeaders(res) {
  res.setHeader("Content-Type", "application/json; charset=utf-8");
}

function send(res, status, payload) {
  res.status(status).send(JSON.stringify(payload));
}

async function readBody(req) {
  if (req.body && typeof req.body === "object") {
    return req.body;
  }

  if (typeof req.body === "string") {
    return JSON.parse(req.body || "{}");
  }

  const chunks = [];
  for await (const chunk of req) {
    chunks.push(Buffer.from(chunk));
  }

  const raw = Buffer.concat(chunks).toString("utf8");
  return raw ? JSON.parse(raw) : {};
}
