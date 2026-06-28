const { Pool } = require("pg");
const { getUserFromRequest } = require("./_auth");

let pool;
let tableReady = false;

module.exports = async function handler(req, res) {
  setJsonHeaders(res);

  const username = getUserFromRequest(req);
  if (!username) {
    return send(res, 401, { ok: false, error: "Faca login para acessar os dados." });
  }

  if (!process.env.DATABASE_URL) {
    return send(res, 500, {
      ok: false,
      error: "DATABASE_URL nao configurada na Vercel.",
    });
  }

  try {
    await ensureTable();

    if (req.method === "GET") {
      return getSession(req, res, username);
    }

    if (req.method === "POST") {
      return saveSession(req, res, username);
    }

    if (req.method === "DELETE") {
      return deleteSession(req, res, username);
    }

    return send(res, 405, { ok: false, error: "Método não permitido." });
  } catch (error) {
    console.error(error);
    return send(res, 500, {
      ok: false,
      error: databaseErrorMessage(error),
      code: error.code || "",
    });
  }
};

function getPool() {
  if (!pool) {
    pool = new Pool({
      connectionString: databaseUrlWithoutSslMode(),
      ssl: buildSslConfig(),
      max: 1,
      idleTimeoutMillis: 10000,
      connectionTimeoutMillis: 4000,
    });
  }

  return pool;
}

function databaseUrlWithoutSslMode() {
  const url = new URL(process.env.DATABASE_URL);
  url.searchParams.delete("sslmode");
  return url.toString();
}

function buildSslConfig() {
  if (process.env.PG_CA_CERT) {
    return {
      ca: process.env.PG_CA_CERT.replace(/\\n/g, "\n"),
      rejectUnauthorized: true,
    };
  }

  return { rejectUnauthorized: false };
}

function setJsonHeaders(res) {
  res.setHeader("Content-Type", "application/json; charset=utf-8");
}

function send(res, status, payload) {
  res.status(status).send(JSON.stringify(payload));
}

async function ensureTable() {
  if (tableReady) return;

  await runQuery(`
    create table if not exists dashboard_user_sessions (
      username text primary key,
      payload jsonb not null,
      created_at timestamptz not null default now(),
      updated_at timestamptz not null default now()
    )
  `);

  tableReady = true;
}

async function getSession(req, res, username) {
  const result = await runQuery(
    "select payload, updated_at from dashboard_user_sessions where username = $1",
    [username]
  );

  if (!result.rowCount) {
    return send(res, 200, { ok: true, payload: null });
  }

  return send(res, 200, {
    ok: true,
    payload: result.rows[0].payload,
    updatedAt: result.rows[0].updated_at,
  });
}

async function saveSession(req, res, username) {
  const body = await readBody(req);
  const payload = body.payload;

  if (!payload || payload.version !== 1) {
    return send(res, 400, { ok: false, error: "Sessão inválida." });
  }

  await runQuery(
    `
      insert into dashboard_user_sessions (username, payload, updated_at)
      values ($1, $2::jsonb, now())
      on conflict (username)
      do update set payload = excluded.payload, updated_at = now()
    `,
    [username, JSON.stringify(payload)]
  );

  return send(res, 200, { ok: true });
}

async function deleteSession(req, res, username) {
  await runQuery("delete from dashboard_user_sessions where username = $1", [username]);
  return send(res, 200, { ok: true });
}

function runQuery(sql, params = []) {
  return Promise.race([
    getPool().query(sql, params),
    new Promise((_, reject) => {
      setTimeout(() => {
        const error = new Error("Timeout ao conectar no banco.");
        error.code = "DB_TIMEOUT";
        reject(error);
      }, 7000);
    }),
  ]);
}

function databaseErrorMessage(error) {
  const message = String(error.message || "");

  if (error.code === "DB_TIMEOUT" || /timeout/i.test(message)) {
    return "Banco: timeout ao conectar. Verifique se o Aiven permite conexoes externas/Vercel.";
  }

  if (error.code === "ENOTFOUND" || error.code === "ECONNREFUSED") {
    return `Banco: conexao recusada (${error.code}). Confira host e porta do DATABASE_URL.`;
  }

  if (error.code === "28P01") {
    return "Banco: usuario ou senha do DATABASE_URL estao incorretos.";
  }

  if (/ssl|certificate/i.test(message)) {
    return "Banco: erro de SSL/certificado. Confira sslmode=require ou PG_CA_CERT.";
  }

  if (error.code) {
    return `Banco: ${error.code} - ${message}`;
  }

  return `Banco: ${message || "erro desconhecido"}`;
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
