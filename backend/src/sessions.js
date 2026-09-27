import { authorizeAccessCode } from "./auth.js";

const encoder = new TextEncoder();

function formatSessionId(sequence) {
  return "S" + String(Math.max(1, Number(sequence) || 1)).padStart(4, "0");
}

async function sha256Hex(value) {
  const digest = await crypto.subtle.digest(
    "SHA-256",
    encoder.encode(String(value))
  );

  return Array.from(new Uint8Array(digest), (byte) =>
    byte.toString(16).padStart(2, "0")
  ).join("");
}

async function getReceipt(db, requestId) {
  return db
    .prepare(
      `SELECT request_id, operation, request_hash, response_status, response_json
       FROM request_receipts
       WHERE request_id = ?1`
    )
    .bind(requestId)
    .first();
}

function parseReceiptResponse(receipt) {
  if (!receipt?.response_json) return null;

  try {
    return JSON.parse(receipt.response_json);
  } catch {
    return null;
  }
}

async function claimRequest(db, requestId, operation, requestHash, createdAt) {
  const result = await db
    .prepare(
      `INSERT OR IGNORE INTO request_receipts
         (request_id, operation, request_hash, response_status, response_json, created_at)
       VALUES (?1, ?2, ?3, 0, NULL, ?4)`
    )
    .bind(requestId, operation, requestHash, createdAt)
    .run();

  return Number(result?.meta?.changes || 0) === 1;
}

async function completeReceipt(db, requestId, responseStatus, responseBody) {
  await db
    .prepare(
      `UPDATE request_receipts
       SET response_status = ?2,
           response_json = ?3
       WHERE request_id = ?1`
    )
    .bind(requestId, responseStatus, JSON.stringify(responseBody))
    .run();
}

async function allocateSequence(db, app, participantId) {
  const table =
    app === "training"
      ? "training_participant_counters"
      : "aqg_participant_counters";

  const row = await db
    .prepare(
      `INSERT INTO ${table} (participant_id, next_session_number)
       VALUES (?1, 2)
       ON CONFLICT(participant_id) DO UPDATE SET
         next_session_number = next_session_number + 1
       RETURNING next_session_number - 1 AS session_seq`
    )
    .bind(participantId)
    .first();

  const sequence = Number(row?.session_seq);

  if (!Number.isInteger(sequence) || sequence < 1) {
    throw new Error("Could not allocate a session number");
  }

  return sequence;
}

export async function createSessionId(env, input, app) {
  const auth = await authorizeAccessCode(env, input, app);

  if (!auth.ok) {
    return {
      status: auth.status,
      body: auth.body
    };
  }

  const requestId = String(input.request_id || "").trim();

  if (!requestId || requestId.length > 120) {
    return {
      status: 400,
      body: {
        ok: false,
        error: "A valid request identifier is required.",
        code: "INVALID_REQUEST",
        retryable: false
      }
    };
  }

  const participantId = auth.participantId;
  const operation = `${app}:createSessionId`;
  const requestHash = await sha256Hex(
    `${operation}|${participantId}`
  );

  const existing = await getReceipt(env.STUDY_DB, requestId);

  if (existing) {
    if (
      existing.operation !== operation ||
      existing.request_hash !== requestHash
    ) {
      return {
        status: 409,
        body: {
          ok: false,
          error: "This request identifier was already used for a different operation.",
          code: "REQUEST_ID_CONFLICT",
          retryable: false
        }
      };
    }

    const prior = parseReceiptResponse(existing);

    if (prior) {
      return {
        status: Number(existing.response_status || 200),
        body: {
          ...prior,
          duplicate: true
        }
      };
    }

    return {
      status: 503,
      body: {
        ok: false,
        error: "The earlier request is still being finalized. Please retry.",
        code: "REQUEST_PENDING",
        retryable: true
      }
    };
  }

  const createdAt = new Date().toISOString();
  const claimed = await claimRequest(
    env.STUDY_DB,
    requestId,
    operation,
    requestHash,
    createdAt
  );

  if (!claimed) {
    const racedReceipt = await getReceipt(env.STUDY_DB, requestId);
    const prior = parseReceiptResponse(racedReceipt);

    if (prior) {
      return {
        status: Number(racedReceipt.response_status || 200),
        body: {
          ...prior,
          duplicate: true
        }
      };
    }

    return {
      status: 503,
      body: {
        ok: false,
        error: "The request is already being processed. Please retry.",
        code: "REQUEST_PENDING",
        retryable: true
      }
    };
  }

  try {
    const sessionSeq = await allocateSequence(
      env.STUDY_DB,
      app,
      participantId
    );

    const responseBody = {
      ok: true,
      user_id: participantId,
      session_id: `${formatSessionId(sessionSeq)}-${crypto.randomUUID()}`,
      session_seq: sessionSeq,
      duplicate: false
    };

    await completeReceipt(
      env.STUDY_DB,
      requestId,
      200,
      responseBody
    );

    return {
      status: 200,
      body: responseBody
    };
  } catch (error) {
    await env.STUDY_DB
      .prepare(
        `DELETE FROM request_receipts
         WHERE request_id = ?1
           AND response_status = 0
           AND response_json IS NULL`
      )
      .bind(requestId)
      .run();

    throw error;
  }
}


function formatContextId(sequence) {
  return "C" + String(Math.max(1, Number(sequence) || 1)).padStart(4, "0");
}

async function sessionBelongsToParticipant(db, participantId, sessionId) {
  const row = await db
    .prepare(
      `SELECT 1 AS ok
       WHERE EXISTS (
         SELECT 1
         FROM request_receipts
         WHERE operation = 'aqg:createSessionId'
           AND json_extract(response_json, '$.user_id') = ?1
           AND json_extract(response_json, '$.session_id') = ?2
       )
       OR EXISTS (
         SELECT 1 FROM aqg_live_sessions
         WHERE participant_id = ?1 AND session_id = ?2
       )
       OR EXISTS (
         SELECT 1 FROM aqg_submissions
         WHERE participant_id = ?1 AND session_id = ?2
       )
       LIMIT 1`
    )
    .bind(participantId, sessionId)
    .first();

  return row?.ok === 1;
}

async function allocateContextSequence(db, participantId, sessionId) {
  const row = await db
    .prepare(
      `INSERT INTO aqg_session_context_counters
         (participant_id, session_id, next_context_number)
       VALUES (?1, ?2, 2)
       ON CONFLICT(participant_id, session_id) DO UPDATE SET
         next_context_number = next_context_number + 1
       RETURNING next_context_number - 1 AS context_seq`
    )
    .bind(participantId, sessionId)
    .first();

  const sequence = Number(row?.context_seq);
  if (!Number.isInteger(sequence) || sequence < 1) {
    throw new Error("Could not allocate a context number");
  }
  return sequence;
}

export async function createContextId(env, input) {
  const auth = await authorizeAccessCode(env, input, "aqg");
  if (!auth.ok) return { status: auth.status, body: auth.body };

  const participantId = auth.participantId;
  const sessionId = String(input.session_id || input.sessionId || "").trim();
  const requestId = String(input.request_id || "").trim();

  if (!sessionId) {
    return {
      status: 400,
      body: {
        ok: false,
        error: "Session ID is required.",
        code: "MISSING_SESSION",
        retryable: false
      }
    };
  }

  if (!requestId || requestId.length > 120) {
    return {
      status: 400,
      body: {
        ok: false,
        error: "A valid request identifier is required.",
        code: "INVALID_REQUEST",
        retryable: false
      }
    };
  }

  const belongs = await sessionBelongsToParticipant(
    env.STUDY_DB,
    participantId,
    sessionId
  );

  if (!belongs) {
    return {
      status: 403,
      body: {
        ok: false,
        error: "This session does not belong to this participant.",
        code: "SESSION_CHANGED",
        retryable: false
      }
    };
  }

  const operation = "aqg:createContextId";
  const requestHash = await sha256Hex(
    `${operation}|${participantId}|${sessionId}`
  );

  const existing = await getReceipt(env.STUDY_DB, requestId);
  if (existing) {
    if (existing.operation !== operation || existing.request_hash !== requestHash) {
      return {
        status: 409,
        body: {
          ok: false,
          error: "This request identifier was already used for a different operation.",
          code: "REQUEST_ID_CONFLICT",
          retryable: false
        }
      };
    }

    const prior = parseReceiptResponse(existing);
    if (prior) {
      return {
        status: Number(existing.response_status || 200),
        body: { ...prior, duplicate: true }
      };
    }

    return {
      status: 503,
      body: {
        ok: false,
        error: "The earlier request is still being finalized. Please retry.",
        code: "REQUEST_PENDING",
        retryable: true
      }
    };
  }

  const createdAt = new Date().toISOString();
  const claimed = await claimRequest(
    env.STUDY_DB,
    requestId,
    operation,
    requestHash,
    createdAt
  );

  if (!claimed) {
    const racedReceipt = await getReceipt(env.STUDY_DB, requestId);
    const prior = parseReceiptResponse(racedReceipt);

    if (prior) {
      return {
        status: Number(racedReceipt.response_status || 200),
        body: { ...prior, duplicate: true }
      };
    }

    return {
      status: 503,
      body: {
        ok: false,
        error: "The request is already being processed. Please retry.",
        code: "REQUEST_PENDING",
        retryable: true
      }
    };
  }

  try {
    const contextSeq = await allocateContextSequence(
      env.STUDY_DB,
      participantId,
      sessionId
    );

    const responseBody = {
      ok: true,
      user_id: participantId,
      session_id: sessionId,
      context_id: formatContextId(contextSeq),
      context_seq: contextSeq,
      duplicate: false
    };

    await completeReceipt(env.STUDY_DB, requestId, 200, responseBody);

    return { status: 200, body: responseBody };
  } catch (error) {
    await env.STUDY_DB
      .prepare(
        `DELETE FROM request_receipts
         WHERE request_id = ?1
           AND response_status = 0
           AND response_json IS NULL`
      )
      .bind(requestId)
      .run();

    throw error;
  }
}
