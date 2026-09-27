import { authorizeAccessCode } from "./auth.js";

const encoder = new TextEncoder();
const MAX_PAYLOAD_CHARS = 450000;
const MAX_BATCH_EVENTS = 80;

function textValue(...values) {
  for (const value of values) {
    if (value !== undefined && value !== null && String(value).trim() !== "") {
      return String(value).trim();
    }
  }
  return "";
}

function numberOrNull(value) {
  if (value === undefined || value === null || value === "") return null;
  const n = Number(value);
  return Number.isFinite(n) ? n : null;
}

function jsonText(value, fallback = null) {
  if (value === undefined || value === null || value === "") return fallback;
  if (typeof value === "string") {
    try {
      JSON.parse(value);
      return value;
    } catch {
      return JSON.stringify(value);
    }
  }
  return JSON.stringify(value);
}

function stableValue(value) {
  if (Array.isArray(value)) return value.map(stableValue);
  if (value && typeof value === "object") {
    return Object.fromEntries(
      Object.keys(value)
        .sort()
        .map((key) => [key, stableValue(value[key])])
    );
  }
  return value;
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

function receiptResponse(receipt) {
  if (!receipt?.response_json) return null;
  try {
    return JSON.parse(receipt.response_json);
  } catch {
    return null;
  }
}

async function sessionWasAllocated(db, participantId, sessionId) {
  const row = await db
    .prepare(
      `SELECT 1 AS ok
       FROM request_receipts
       WHERE operation = 'aqg:createSessionId'
         AND json_extract(response_json, '$.user_id') = ?1
         AND json_extract(response_json, '$.session_id') = ?2
       LIMIT 1`
    )
    .bind(participantId, sessionId)
    .first();

  return row?.ok === 1;
}

function normalizeStatus(value) {
  const status = String(value || "in_progress").trim().toLowerCase();
  if (status === "submitted") return "submitted";
  if (status === "completed") return "completed";
  return "in_progress";
}

function eventTimestamp(event, fallback) {
  return textValue(event?.timestamp, event?.t, fallback, new Date().toISOString());
}

function buildEventStatements(db, input, participantId, sessionId, requestId, nextRevision, previousLastEventAt) {
  const events = Array.isArray(input.batch_events) ? input.batch_events : [];
  if (!events.length) return { statements: [], lastEventAt: previousLastEventAt || "" };

  let priorMs = previousLastEventAt ? Date.parse(previousLastEventAt) : NaN;
  let lastEventAt = previousLastEventAt || "";
  const batchId = `${participantId}:${sessionId}:${requestId}`;
  const statements = [];

  for (const event of events) {
    const eventId = textValue(event?.event_id);
    if (!eventId) {
      throw Object.assign(new Error("An event identifier is missing."), {
        code: "INVALID_REQUEST",
        status: 400
      });
    }

    const timestamp = eventTimestamp(event, input.last_activity_at);
    const timeMs = Date.parse(timestamp);
    const msSincePrevious =
      Number.isFinite(priorMs) && Number.isFinite(timeMs)
        ? Math.max(0, timeMs - priorMs)
        : null;

    if (Number.isFinite(timeMs)) priorMs = timeMs;
    lastEventAt = timestamp;

    statements.push(
      db
        .prepare(
          `INSERT OR IGNORE INTO aqg_events (
             record_id, event_id, batch_id, event_timestamp, ms_since_previous,
             participant_id, session_id, context_id, event_type, button_id,
             detail_text, detail_json, device_label, created_at
           )
           SELECT ?1, ?2, ?3, ?4, ?5, ?6, ?7, ?8, ?9, ?10, ?11, ?12, ?13, ?14
           WHERE EXISTS (
             SELECT 1
             FROM aqg_live_sessions
             WHERE participant_id = ?6
               AND session_id = ?7
               AND revision = ?15
               AND last_save_id = ?16
           )`
        )
        .bind(
          crypto.randomUUID(),
          eventId,
          batchId,
          timestamp,
          msSincePrevious,
          participantId,
          sessionId,
          textValue(event?.context_id, input.context_id) || null,
          textValue(event?.event_type, event?.type),
          textValue(event?.button_id) || null,
          textValue(event?.detail_text) || null,
          jsonText(event?.detail_json ?? event?.details, null),
          textValue(event?.device_label, input.device_label) || null,
          new Date().toISOString(),
          nextRevision,
          requestId
        )
    );
  }

  return { statements, lastEventAt };
}

export async function saveAqgLiveSession(env, input) {
  if (JSON.stringify(input).length > MAX_PAYLOAD_CHARS) {
    return {
      status: 413,
      body: {
        ok: false,
        error: "This save is too large.",
        code: "PAYLOAD_TOO_LARGE",
        retryable: false
      }
    };
  }

  if (
    Array.isArray(input.batch_events) &&
    input.batch_events.length > MAX_BATCH_EVENTS
  ) {
    return {
      status: 413,
      body: {
        ok: false,
        error: "Too many events in one save.",
        code: "PAYLOAD_TOO_LARGE",
        retryable: false
      }
    };
  }

  const auth = await authorizeAccessCode(env, input, "aqg");
  if (!auth.ok) return { status: auth.status, body: auth.body };

  const participantId = auth.participantId;
  const sessionId = textValue(input.session_id, input.sessionId);
  const requestId = textValue(input.request_id);
  const baseRevision = Number(input.base_revision);

  if (!sessionId || !requestId || requestId.length > 120) {
    return {
      status: 400,
      body: {
        ok: false,
        error: "Participant, session, and request identifier are required.",
        code: "MISSING_SESSION",
        retryable: false
      }
    };
  }

  if (!Number.isInteger(baseRevision) || baseRevision < 0) {
    return {
      status: 400,
      body: {
        ok: false,
        error: "A valid base revision is required.",
        code: "INVALID_REQUEST",
        retryable: false
      }
    };
  }

  const operation = "aqg:updateLiveSession";
  const requestHash = await sha256Hex(
    JSON.stringify(stableValue({ ...input, user_id: participantId }))
  );

  const existingReceipt = await getReceipt(env.STUDY_DB, requestId);
  if (existingReceipt) {
    if (
      existingReceipt.operation !== operation ||
      existingReceipt.request_hash !== requestHash
    ) {
      return {
        status: 409,
        body: {
          ok: false,
          error: "This request identifier was already used for a different save.",
          code: "REQUEST_ID_CONFLICT",
          retryable: false
        }
      };
    }

    const prior = receiptResponse(existingReceipt);
    if (prior) {
      return {
        status: Number(existingReceipt.response_status || 200),
        body: { ...prior, duplicate: true }
      };
    }
  }

  const previous = await env.STUDY_DB
    .prepare(
      `SELECT *
       FROM aqg_live_sessions
       WHERE participant_id = ?1
         AND session_id = ?2`
    )
    .bind(participantId, sessionId)
    .first();

  if (!previous) {
    const allocated = await sessionWasAllocated(
      env.STUDY_DB,
      participantId,
      sessionId
    );

    if (!allocated) {
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
  }

  const currentRevision = Number(previous?.revision || 0);

  if (previous?.status === "submitted") {
    return {
      status: 409,
      body: {
        ok: false,
        error: "This session has already been submitted.",
        code: "SESSION_CLOSED",
        retryable: false
      }
    };
  }

  if (baseRevision !== currentRevision) {
    return {
      status: 409,
      body: {
        ok: false,
        error: "Another window has saved this session, or this page is outdated.",
        code: "REVISION_CONFLICT",
        retryable: false,
        revision: currentRevision
      }
    };
  }

  const now = new Date().toISOString();
  const nextRevision = currentRevision + 1;
  const status = normalizeStatus(
    textValue(input.status, previous?.status, "in_progress")
  );

  const fullEventsJson =
    Array.isArray(input.events) || typeof input.events === "object"
      ? jsonText(input.events, previous?.events_json || null)
      : previous?.events_json || null;

  const progressJson =
    input.progress_json !== undefined
      ? jsonText(input.progress_json, previous?.progress_json || null)
      : previous?.progress_json || null;

  const eventBuild = buildEventStatements(
    env.STUDY_DB,
    input,
    participantId,
    sessionId,
    requestId,
    nextRevision,
    previous?.last_event_at || ""
  );

  const lastEventAt = eventBuild.lastEventAt || previous?.last_event_at || null;
  const responseBody = {
    ok: true,
    updated: Boolean(previous),
    inserted: !previous,
    revision: nextRevision,
    duplicate: false
  };

  const liveStatement = env.STUDY_DB
    .prepare(
      `INSERT INTO aqg_live_sessions (
         record_id, participant_id, session_id, revision, last_save_id,
         last_event_at, context_id, status, session_start, last_activity_at,
         submitted_at, llm_product, llm_model, llm_description, model_used,
         course_context, question_context, extra_instructions, desired_questions,
         final_response, feedback_text, audio_object_key, audio_duration_seconds,
         active_seconds, progress_json, events_json, session_close_type,
         session_close_at, app_version, mode, created_at, updated_at
       )
       VALUES (
         ?1, ?2, ?3, ?4, ?5, ?6, ?7, ?8, ?9, ?10,
         ?11, ?12, ?13, ?14, ?15, ?16, ?17, ?18, ?19,
         ?20, ?21, ?22, ?23, ?24, ?25, ?26, ?27, ?28,
         ?29, ?30, ?31, ?32
       )
       ON CONFLICT(participant_id, session_id) DO UPDATE SET
         revision = excluded.revision,
         last_save_id = excluded.last_save_id,
         last_event_at = excluded.last_event_at,
         context_id = excluded.context_id,
         status = excluded.status,
         session_start = excluded.session_start,
         last_activity_at = excluded.last_activity_at,
         submitted_at = excluded.submitted_at,
         llm_product = excluded.llm_product,
         llm_model = excluded.llm_model,
         llm_description = excluded.llm_description,
         model_used = excluded.model_used,
         course_context = excluded.course_context,
         question_context = excluded.question_context,
         extra_instructions = excluded.extra_instructions,
         desired_questions = excluded.desired_questions,
         final_response = excluded.final_response,
         feedback_text = excluded.feedback_text,
         audio_object_key = excluded.audio_object_key,
         audio_duration_seconds = excluded.audio_duration_seconds,
         active_seconds = excluded.active_seconds,
         progress_json = excluded.progress_json,
         events_json = excluded.events_json,
         session_close_type = excluded.session_close_type,
         session_close_at = excluded.session_close_at,
         app_version = excluded.app_version,
         mode = excluded.mode,
         updated_at = excluded.updated_at
       WHERE aqg_live_sessions.revision = ?33
         AND aqg_live_sessions.status <> 'submitted'
       RETURNING revision`
    )
    .bind(
      previous?.record_id || crypto.randomUUID(),
      participantId,
      sessionId,
      nextRevision,
      requestId,
      lastEventAt,
      textValue(input.context_id, input.contextId, previous?.context_id) || null,
      status,
      textValue(input.session_start, input.sessionStart, previous?.session_start) || null,
      textValue(input.last_activity_at, input.lastActivityAt, previous?.last_activity_at, now) || now,
      previous?.submitted_at || null,
      textValue(input.llm_product, input.llmProduct, previous?.llm_product) || null,
      textValue(input.llm_expected_choice, input.llmExpectedChoice, previous?.llm_model) || null,
      textValue(input.llm_mismatch_description, input.llmMismatchDescription, previous?.llm_description) || null,
      textValue(input.model_used, input.modelUsed, previous?.model_used) || null,
      textValue(input.course_context, input.courseContext, previous?.course_context) || null,
      textValue(input.question_context, input.questionContext, previous?.question_context) || null,
      textValue(input.extra_instructions, input.extraInstructions, previous?.extra_instructions) || null,
      textValue(input.desired_questions, input.desiredQuestions, previous?.desired_questions) || null,
      textValue(input.llm_response, input.llmResponse, previous?.final_response) || null,
      textValue(input.session_notes, input.sessionNotes, previous?.feedback_text) || null,
      textValue(input.audio_object_key, previous?.audio_object_key) || null,
      numberOrNull(input.audio_duration ?? previous?.audio_duration_seconds),
      numberOrNull(input.active_seconds ?? input.activeSeconds ?? previous?.active_seconds),
      progressJson,
      fullEventsJson,
      textValue(input.session_close_type, input.sessionCloseType, previous?.session_close_type) || null,
      textValue(input.session_close_at, input.sessionCloseAt, previous?.session_close_at) || null,
      textValue(input.app_version, input.appVersion, previous?.app_version) || null,
      textValue(input.mode, previous?.mode, "standard") || "standard",
      previous?.created_at || now,
      now,
      baseRevision
    );

  const receiptStatement = env.STUDY_DB
    .prepare(
      `INSERT OR IGNORE INTO request_receipts
         (request_id, operation, request_hash, response_status, response_json, created_at)
       SELECT ?1, ?2, ?3, 200, ?4, ?5
       WHERE EXISTS (
         SELECT 1
         FROM aqg_live_sessions
         WHERE participant_id = ?6
           AND session_id = ?7
           AND revision = ?8
           AND last_save_id = ?1
       )`
    )
    .bind(
      requestId,
      operation,
      requestHash,
      JSON.stringify(responseBody),
      now,
      participantId,
      sessionId,
      nextRevision
    );

  const batchResults = await env.STUDY_DB.batch([
    liveStatement,
    ...eventBuild.statements,
    receiptStatement
  ]);

  const savedRevision = batchResults?.[0]?.results?.[0]?.revision;

  if (Number(savedRevision) !== nextRevision) {
    const racedReceipt = await getReceipt(env.STUDY_DB, requestId);
    const prior = receiptResponse(racedReceipt);

    if (
      racedReceipt &&
      racedReceipt.operation === operation &&
      racedReceipt.request_hash === requestHash &&
      prior
    ) {
      return {
        status: Number(racedReceipt.response_status || 200),
        body: { ...prior, duplicate: true }
      };
    }

    const latest = await env.STUDY_DB
      .prepare(
        `SELECT revision, status
         FROM aqg_live_sessions
         WHERE participant_id = ?1
           AND session_id = ?2`
      )
      .bind(participantId, sessionId)
      .first();

    return {
      status: 409,
      body: {
        ok: false,
        error:
          latest?.status === "submitted"
            ? "This session has already been submitted."
            : "Another window has saved this session, or this page is outdated.",
        code:
          latest?.status === "submitted"
            ? "SESSION_CLOSED"
            : "REVISION_CONFLICT",
        retryable: false,
        revision: Number(latest?.revision || currentRevision)
      }
    };
  }

  return { status: 200, body: responseBody };
}
