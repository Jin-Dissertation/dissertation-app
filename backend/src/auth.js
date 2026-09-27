const MAX_FAILED_ATTEMPTS = 10;
const LOCKOUT_MS = 2 * 60 * 1000;

const encoder = new TextEncoder();

function normalizeCode(value) {
  return String(value ?? "").trim().toLowerCase();
}

function normalizeClientKey(value) {
  return String(value ?? "").trim();
}

function toHex(buffer) {
  return Array.from(new Uint8Array(buffer), (byte) =>
    byte.toString(16).padStart(2, "0")
  ).join("");
}

async function hmacHex(secret, value) {
  if (!secret) {
    throw new Error("ACCESS_CODE_PEPPER is not configured");
  }

  const key = await crypto.subtle.importKey(
    "raw",
    encoder.encode(secret),
    { name: "HMAC", hash: "SHA-256" },
    false,
    ["sign"]
  );

  const signature = await crypto.subtle.sign(
    "HMAC",
    key,
    encoder.encode(value)
  );

  return toHex(signature);
}

async function getFailureState(db, clientKeyHash) {
  return db
    .prepare(
      `SELECT failure_count, locked_until
       FROM auth_failure_state
       WHERE client_key_hash = ?1`
    )
    .bind(clientKeyHash)
    .first();
}

async function clearFailureState(db, clientKeyHash) {
  await db
    .prepare("DELETE FROM auth_failure_state WHERE client_key_hash = ?1")
    .bind(clientKeyHash)
    .run();
}

async function recordFailure(db, clientKeyHash, nowMs) {
  const current = await getFailureState(db, clientKeyHash);

  let count = Number(current?.failure_count || 0);
  const currentLockedUntilMs = current?.locked_until
    ? Date.parse(current.locked_until)
    : NaN;

  if (Number.isFinite(currentLockedUntilMs) && currentLockedUntilMs <= nowMs) {
    count = 0;
  }

  count += 1;

  const locked = count >= MAX_FAILED_ATTEMPTS;
  const lockedUntil = locked
    ? new Date(nowMs + LOCKOUT_MS).toISOString()
    : null;
  const updatedAt = new Date(nowMs).toISOString();

  await db
    .prepare(
      `INSERT INTO auth_failure_state
         (client_key_hash, failure_count, locked_until, updated_at)
       VALUES (?1, ?2, ?3, ?4)
       ON CONFLICT(client_key_hash) DO UPDATE SET
         failure_count = excluded.failure_count,
         locked_until = excluded.locked_until,
         updated_at = excluded.updated_at`
    )
    .bind(clientKeyHash, count, lockedUntil, updatedAt)
    .run();

  return {
    locked,
    retryAfterSeconds: locked
      ? Math.ceil((Date.parse(lockedUntil) - nowMs) / 1000)
      : 0,
    attemptsRemaining: locked
      ? 0
      : Math.max(0, MAX_FAILED_ATTEMPTS - count)
  };
}

export async function validateAccessCode(env, input, app) {
  const code = normalizeCode(
    input.code ?? input.access_code ?? input.user_id
  );
  const clientKey = normalizeClientKey(
    input.client_key ?? input.clientKey
  );

  if (!code) {
    return {
      status: 400,
      body: { ok: false, valid: false, error: "Missing code" }
    };
  }

  if (!clientKey) {
    return {
      status: 400,
      body: { ok: false, valid: false, error: "Missing client key" }
    };
  }

  if (code.length > 256 || clientKey.length > 256) {
    return {
      status: 400,
      body: { ok: false, valid: false, error: "Invalid request" }
    };
  }

  if (app !== "aqg" && app !== "training") {
    return {
      status: 400,
      body: { ok: false, valid: false, error: "Invalid app" }
    };
  }

  const clientKeyHash = await hmacHex(
    env.ACCESS_CODE_PEPPER,
    `client:${clientKey}`
  );

  const nowMs = Date.now();
  const failureState = await getFailureState(env.STUDY_DB, clientKeyHash);

  if (failureState?.locked_until) {
    const lockedUntilMs = Date.parse(failureState.locked_until);

    if (Number.isFinite(lockedUntilMs) && lockedUntilMs > nowMs) {
      return {
        status: 200,
        body: {
          ok: true,
          valid: false,
          locked: true,
          retryAfterSeconds: Math.ceil((lockedUntilMs - nowMs) / 1000),
          error: "Too many failed attempts. Please wait before trying again."
        }
      };
    }

    await clearFailureState(env.STUDY_DB, clientKeyHash);
  }

  const codeHash = await hmacHex(
    env.ACCESS_CODE_PEPPER,
    `access:${code}`
  );

  const row = await env.STUDY_DB
    .prepare(
      `SELECT participant_id, active, allow_aqg, allow_training
       FROM access_codes
       WHERE code_hash = ?1`
    )
    .bind(codeHash)
    .first();

  const permitted =
    row &&
    Number(row.active) === 1 &&
    (app === "aqg"
      ? Number(row.allow_aqg) === 1
      : Number(row.allow_training) === 1);

  if (permitted) {
    await clearFailureState(env.STUDY_DB, clientKeyHash);

    return {
      status: 200,
      body: {
        ok: true,
        valid: true,
        participant_id: String(row.participant_id)
      }
    };
  }

  const recorded = await recordFailure(
    env.STUDY_DB,
    clientKeyHash,
    nowMs
  );

  return {
    status: 200,
    body: {
      ok: true,
      valid: false,
      locked: recorded.locked,
      retryAfterSeconds: recorded.retryAfterSeconds,
      attemptsRemaining: recorded.attemptsRemaining
    }
  };
}


async function lookupAccessCode(env, code, app) {
  const codeHash = await hmacHex(
    env.ACCESS_CODE_PEPPER,
    `access:${code}`
  );

  const row = await env.STUDY_DB
    .prepare(
      `SELECT participant_id, active, allow_aqg, allow_training
       FROM access_codes
       WHERE code_hash = ?1`
    )
    .bind(codeHash)
    .first();

  if (
    !row ||
    Number(row.active) !== 1 ||
    (app === "aqg" && Number(row.allow_aqg) !== 1) ||
    (app === "training" && Number(row.allow_training) !== 1)
  ) {
    return null;
  }

  return row;
}

export async function authorizeAccessCode(env, input, app) {
  if (app !== "aqg" && app !== "training") {
    return {
      ok: false,
      status: 400,
      body: {
        ok: false,
        error: "Invalid app",
        code: "INVALID_REQUEST",
        retryable: false
      }
    };
  }

  const identifierKeys = [
    "user_id",
    "userId",
    "participant_code",
    "participantCode",
    "accessCode",
    "access_code",
    "code"
  ];

  const supplied = identifierKeys
    .filter((key) => input[key] !== undefined && input[key] !== null && String(input[key]).trim() !== "")
    .map((key) => normalizeCode(input[key]));

  const code = supplied[0] || "";

  if (!code) {
    return {
      ok: false,
      status: 401,
      body: {
        ok: false,
        error: "Please sign in with an active participant code.",
        code: "UNAUTHORIZED",
        retryable: false
      }
    };
  }

  if (code.length > 256 || supplied.some((value) => value !== code)) {
    return {
      ok: false,
      status: 401,
      body: {
        ok: false,
        error: "Participant identifiers do not match.",
        code: "UNAUTHORIZED",
        retryable: false
      }
    };
  }

  const row = await lookupAccessCode(env, code, app);

  if (!row) {
    return {
      ok: false,
      status: 401,
      body: {
        ok: false,
        error: "Please sign in with an active participant code.",
        code: "UNAUTHORIZED",
        retryable: false
      }
    };
  }

  return {
    ok: true,
    participantId: String(row.participant_id)
  };
}
