const crypto = require("crypto");
const { Pool } = require("pg");

const pool = new Pool({
  connectionString: process.env.DATABASE_URL,
  ssl: buildSslConfig(),
  max: 3,
  idleTimeoutMillis: 30000,
  connectionTimeoutMillis: 10000,
});

let tableReady = false;

module.exports = async function handler(req, res) {
  setJsonHeaders(res);

  if (!process.env.DATABASE_URL) {
    return send(res, 500, {
      ok: false,
      error: "DATABASE_URL não configurada na Vercel.",
    });
  }

  try {
    await ensureTable();

    if (req.method === "GET") {
      return getSession(req, res);
    }

    if (req.method === "POST") {
      return saveSession(req, res);
    }

    if (req.method === "DELETE") {
      return deleteSession(req, res);
    }

    return send(res, 405, { ok: false, error: "Método não permitido." });
  } catch (error) {
    console.error(error);
    return send(res, 500, {
      ok: false,
      error: "Erro ao acessar o banco de dados.",
    });
  }
};

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

  await pool.query(`
    create table if not exists dashboard_sessions (
      session_key_hash text primary key,
      payload jsonb not null,
      created_at timestamptz not null default now(),
      updated_at timestamptz not null default now()
    )
  `);

  tableReady = true;
}

async function getSession(req, res) {
  const key = normalizeKey(req.query.key);
  if (!key) {
    return send(res, 400, { ok: false, error: "Chave da nuvem obrigatória." });
  }

  const result = await pool.query(
    "select payload, updated_at from dashboard_sessions where session_key_hash = $1",
    [hashKey(key)]
  );

  if (!result.rowCount) {
    return send(res, 404, { ok: false, error: "Nenhuma sessão encontrada para essa chave." });
  }

  return send(res, 200, {
    ok: true,
    payload: result.rows[0].payload,
    updatedAt: result.rows[0].updated_at,
  });
}

async function saveSession(req, res) {
  const body = await readBody(req);
  const key = normalizeKey(body.key);
  const payload = body.payload;

  if (!key) {
    return send(res, 400, { ok: false, error: "Chave da nuvem obrigatória." });
  }

  if (!payload || payload.version !== 1) {
    return send(res, 400, { ok: false, error: "Sessão inválida." });
  }

  await pool.query(
    `
      insert into dashboard_sessions (session_key_hash, payload, updated_at)
      values ($1, $2::jsonb, now())
      on conflict (session_key_hash)
      do update set payload = excluded.payload, updated_at = now()
    `,
    [hashKey(key), JSON.stringify(payload)]
  );

  return send(res, 200, { ok: true });
}

async function deleteSession(req, res) {
  const body = await readBody(req);
  const key = normalizeKey(body.key);

  if (!key) {
    return send(res, 400, { ok: false, error: "Chave da nuvem obrigatória." });
  }

  await pool.query("delete from dashboard_sessions where session_key_hash = $1", [hashKey(key)]);
  return send(res, 200, { ok: true });
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

function normalizeKey(key) {
  return String(key || "").trim();
}

function hashKey(key) {
  return crypto.createHash("sha256").update(key).digest("hex");
}
