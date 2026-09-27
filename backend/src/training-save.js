import { authorizeAccessCode } from "./auth.js";

const MAX_PAYLOAD_CHARS = 450000;
const MAX_BATCH_EVENTS = 80;
const encoder = new TextEncoder();

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
    const trimmed = value.trim();
    if (!trimmed) return fallback;
    try {
      JSON.parse(trimmed);
      return trimmed;
    } catch {
      return JSON.stringify({ value: trimmed });
    }
  }
  try {
    return JSON.stringify(value);
  } catch {
    return fallback;
  }
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
       WHERE operation = 'training:createSessionId'
         AND json_extract(response_json, '$.user_id') = ?1
         AND json_extract(response_json, '$.session_id') = ?2
       LIMIT 1`
    )
    .bind(participantId, sessionId)
    .first();

  return row?.ok === 1;
}

function getItemCount(input) {
  let maxCount = 0;
  const explicit = Number(input?.item_count);
  if (Number.isFinite(explicit) && explicit > 0) maxCount = explicit;

  if (Array.isArray(input?.item_values)) {
    maxCount = Math.max(maxCount, input.item_values.length);
  }
  if (Array.isArray(input?.item_ms_values)) {
    maxCount = Math.max(maxCount, input.item_ms_values.length);
  }

  for (const source of [input?.item_map, input?.item_ms_map]) {
    if (!source || typeof source !== "object" || Array.isArray(source)) continue;
    for (const key of Object.keys(source)) {
      const match = key.match(/^item_(\d+)(?:_(?:response|ms))?$/);
      if (match) maxCount = Math.max(maxCount, Number(match[1]));
    }
  }

  for (const key of Object.keys(input || {})) {
    const match = key.match(/^item_(\d+)_(?:response|ms)$/);
    if (match) maxCount = Math.max(maxCount, Number(match[1]));
  }

  return Math.max(0, Math.floor(maxCount));
}

function extractItemSnapshot(input) {
  const itemCount = getItemCount(input);
  const responses = Array(itemCount).fill("");
  const msValues = Array(itemCount).fill(null);
  let hasSnapshot = false;

  if (Array.isArray(input?.item_values)) {
    hasSnapshot = true;
    for (let i = 0; i < Math.min(itemCount, input.item_values.length); i += 1) {
      const v = input.item_values[i];
      responses[i] = v === undefined || v === null ? "" : String(v);
    }
  }

  if (Array.isArray(input?.item_ms_values)) {
    hasSnapshot = true;
    for (let i = 0; i < Math.min(itemCount, input.item_ms_values.length); i += 1) {
      msValues[i] = numberOrNull(input.item_ms_values[i]);
    }
  }

  if (input?.item_map && typeof input.item_map === "object") {
    hasSnapshot = true;
    for (let i = 1; i <= itemCount; i += 1) {
      const v =
        input.item_map[`item_${i}`] ??
        input.item_map[`item_${i}_response`];
      if (v !== undefined && v !== null && String(v) !== "") {
        responses[i - 1] = String(v);
      }
    }
  }

  if (input?.item_ms_map && typeof input.item_ms_map === "object") {
    hasSnapshot = true;
    for (let i = 1; i <= itemCount; i += 1) {
      const v =
        input.item_ms_map[`item_${i}`] ??
        input.item_ms_map[`item_${i}_ms`];
      if (v !== undefined && v !== null && String(v) !== "") {
        msValues[i - 1] = numberOrNull(v);
      }
    }
  }

  for (let i = 1; i <= itemCount; i += 1) {
    const responseKey = `item_${i}_response`;
    const msKey = `item_${i}_ms`;

    if (Object.prototype.hasOwnProperty.call(input || {}, responseKey)) {
      hasSnapshot = true;
      const v = input[responseKey];
      responses[i - 1] = v === undefined || v === null ? "" : String(v);
    }

    if (Object.prototype.hasOwnProperty.call(input || {}, msKey)) {
      hasSnapshot = true;
      msValues[i - 1] = numberOrNull(input[msKey]);
    }
  }

  return { itemCount, responses, msValues, hasSnapshot };
}

function lastTrailLabel(trail) {
  const parts = String(trail || "")
    .split("|")
    .map((part) => part.trim())
    .filter(Boolean);
  if (!parts.length) return "";
  const last = parts[parts.length - 1];
  const colon = last.indexOf(":");
  return colon >= 0 ? last.slice(colon + 1).trim() : last;
}

function appendDeviceTrail(existingTrail, incomingLabel, incomingCardNumber) {
  const trail = String(existingTrail || "").trim();
  const label = String(incomingLabel || "").trim();
  if (!label) return trail;

  if (lastTrailLabel(trail) === label) {
    return trail || `start: ${label}`;
  }

  if (!trail) return `start: ${label}`;

  const n = Number(incomingCardNumber);
  const marker = Number.isFinite(n) && n > 0 ? `card ${n}` : "update";
  return `${trail} | ${marker}: ${label}`;
}

function eventTimestamp(event, fallback) {
  return textValue(event?.timestamp, event?.t, fallback, new Date().toISOString());
}

function buildEventStatements(
  db,
  input,
  participantId,
  sessionId,
  requestId,
  nextRevision,
  previousLastEventAt
) {
  const events = Array.isArray(input.batch_events) ? input.batch_events : [];
  if (!events.length) {
    return { statements: [], lastEventAt: previousLastEventAt || "" };
  }

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
          `INSERT OR IGNORE INTO training_events (
             record_id, event_id, batch_id, event_timestamp, ms_since_previous,
             participant_id, session_id, session_seq, event_type, section_index,
             card_index, detail_text, detail_json, device_label, created_at
           )
           SELECT ?1, ?2, ?3, ?4, ?5, ?6, ?7, ?8, ?9, ?10, ?11, ?12, ?13, ?14, ?15
           WHERE EXISTS (
             SELECT 1 FROM training_live_sessions
             WHERE participant_id = ?6
               AND session_id = ?7
               AND revision = ?16
               AND last_save_id = ?17
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
          numberOrNull(event?.session_seq ?? input.session_seq),
          textValue(event?.event_type, event?.type),
          numberOrNull(event?.section_index),
          numberOrNull(event?.card_index),
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

export async function saveTrainingLiveSession(env, input) {
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

  const auth = await authorizeAccessCode(env, input, "training");
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

  const operation = "training:updateLiveSession";
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
      `SELECT * FROM training_live_sessions
       WHERE participant_id = ?1 AND session_id = ?2`
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

  if (previous?.last_save_id === requestId) {
    return {
      status: 200,
      body: {
        ok: true,
        updated: true,
        inserted: false,
        revision: currentRevision,
        duplicate: true
      }
    };
  }

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
  const snapshot = extractItemSnapshot(input);

  const progressJson =
    input.progress_json !== undefined
      ? jsonText(input.progress_json, previous?.progress_json || null)
      : previous?.progress_json || null;

  const incomingDeviceLabel = textValue(input.device_label, input.deviceLabel);
  const incomingCardNumber = textValue(
    input.device_card_number,
    input.deviceCardNumber,
    input.current_card_number,
    input.currentCardNumber
  );

  const deviceTrail =
    input.device_trail !== undefined || input.deviceTrail !== undefined
      ? textValue(input.device_trail, input.deviceTrail)
      : appendDeviceTrail(
          previous?.device_trail || "",
          incomingDeviceLabel,
          incomingCardNumber
        );

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
  const lastActivityAt = textValue(
    input.last_activity_at,
    input.lastActivityAt,
    input.t,
    previous?.last_activity_at,
    now
  );

  const responseBody = {
    ok: true,
    updated: Boolean(previous),
    inserted: !previous,
    revision: nextRevision,
    duplicate: false
  };

  const liveStatement = env.STUDY_DB
    .prepare(
      `INSERT INTO training_live_sessions (
         record_id, participant_id, session_id, session_seq, revision,
         last_save_id, last_event_at, session_start, last_activity_at, status,
         app_version, content_version, current_card_number, device_trail,
         progress_json, progress_saved_at, last_event_type, created_at, updated_at
       )
       VALUES (
         ?1, ?2, ?3, ?4, ?5, ?6, ?7, ?8, ?9, ?10,
         ?11, ?12, ?13, ?14, ?15, ?16, ?17, ?18, ?19
       )
       ON CONFLICT(participant_id, session_id) DO UPDATE SET
         session_seq = COALESCE(excluded.session_seq, training_live_sessions.session_seq),
         revision = excluded.revision,
         last_save_id = excluded.last_save_id,
         last_event_at = COALESCE(excluded.last_event_at, training_live_sessions.last_event_at),
         session_start = COALESCE(excluded.session_start, training_live_sessions.session_start),
         last_activity_at = excluded.last_activity_at,
         status = excluded.status,
         app_version = COALESCE(excluded.app_version, training_live_sessions.app_version),
         content_version = COALESCE(excluded.content_version, training_live_sessions.content_version),
         current_card_number = COALESCE(excluded.current_card_number, training_live_sessions.current_card_number),
         device_trail = COALESCE(excluded.device_trail, training_live_sessions.device_trail),
         progress_json = COALESCE(excluded.progress_json, training_live_sessions.progress_json),
         progress_saved_at = COALESCE(excluded.progress_saved_at, training_live_sessions.progress_saved_at),
         last_event_type = COALESCE(excluded.last_event_type, training_live_sessions.last_event_type),
         updated_at = excluded.updated_at
       WHERE training_live_sessions.revision = ?20
         AND training_live_sessions.status <> 'submitted'
       RETURNING revision`
    )
    .bind(
      previous?.record_id || crypto.randomUUID(),
      participantId,
      sessionId,
      numberOrNull(input.session_seq ?? input.sessionSeq ?? previous?.session_seq),
      nextRevision,
      requestId,
      lastEventAt,
      textValue(input.session_start, input.sessionStart, previous?.session_start) || null,
      lastActivityAt,
      textValue(input.status, previous?.status, "in_progress") || "in_progress",
      textValue(input.app_version, input.appVersion, previous?.app_version) || null,
      textValue(input.content_version, input.contentVersion, previous?.content_version) || null,
      numberOrNull(
        input.current_card_number ??
          input.currentCardNumber ??
          input.device_card_number ??
          input.deviceCardNumber ??
          previous?.current_card_number
      ),
      deviceTrail || null,
      progressJson,
      textValue(
        input.progress_saved_at,
        input.progressSavedAt,
        lastActivityAt,
        previous?.progress_saved_at
      ) || null,
      textValue(
        input.last_event_type,
        input.lastEventType,
        input.event_type,
        input.eventType,
        previous?.last_event_type
      ) || null,
      previous?.created_at || now,
      now,
      baseRevision
    );

  const itemStatements = [];
  if (snapshot.hasSnapshot) {
    const contentVersion =
      textValue(input.content_version, input.contentVersion, previous?.content_version) ||
      null;

    for (let i = 1; i <= snapshot.itemCount; i += 1) {
      itemStatements.push(
        env.STUDY_DB
          .prepare(
            `INSERT INTO training_live_items (
               record_id, participant_id, session_id, content_version,
               item_number, response_text, response_ms, updated_at
             )
             SELECT ?1, ?2, ?3, ?4, ?5, ?6, ?7, ?8
             WHERE EXISTS (
               SELECT 1 FROM training_live_sessions
               WHERE participant_id = ?2
                 AND session_id = ?3
                 AND revision = ?9
                 AND last_save_id = ?10
             )
             ON CONFLICT(participant_id, session_id, item_number) DO UPDATE SET
               content_version = excluded.content_version,
               response_text = excluded.response_text,
               response_ms = excluded.response_ms,
               updated_at = excluded.updated_at`
          )
          .bind(
            crypto.randomUUID(),
            participantId,
            sessionId,
            contentVersion,
            i,
            snapshot.responses[i - 1],
            snapshot.msValues[i - 1],
            now,
            nextRevision,
            requestId
          )
      );
    }
  }

  const receiptStatement = env.STUDY_DB
    .prepare(
      `INSERT OR IGNORE INTO request_receipts
         (request_id, operation, request_hash, response_status, response_json, created_at)
       SELECT ?1, ?2, ?3, 200, ?4, ?5
       WHERE EXISTS (
         SELECT 1 FROM training_live_sessions
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
    ...itemStatements,
    receiptStatement
  ]);

  const savedRevision = batchResults?.[0]?.results?.[0]?.revision;
  if (Number(savedRevision) !== nextRevision) {
    const latest = await env.STUDY_DB
      .prepare(
        `SELECT revision, status
         FROM training_live_sessions
         WHERE participant_id = ?1 AND session_id = ?2`
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


export async function getLatestTrainingLiveSession(env, input) {
  const auth = await authorizeAccessCode(env, input, "training");
  if (!auth.ok) return { status: auth.status, body: auth.body };

  const row = await env.STUDY_DB
    .prepare(
      `SELECT *
       FROM training_live_sessions
       WHERE participant_id = ?1
         AND status <> 'submitted'
       ORDER BY COALESCE(last_activity_at, progress_saved_at, session_start, '') DESC,
                updated_at DESC
       LIMIT 1`
    )
    .bind(auth.participantId)
    .first();

  if (!row) {
    return { status: 200, body: { ok: true, found: false } };
  }

  let progress = null;
  try {
    if (row.progress_json) progress = JSON.parse(row.progress_json);
  } catch {
    progress = null;
  }

  return {
    status: 200,
    body: {
      ok: true,
      found: true,
      session: {
        revision: Number(row.revision || 0),
        last_save_id: row.last_save_id || "",
        user_id: auth.participantId,
        session_id: row.session_id || "",
        session_seq: row.session_seq == null ? "" : String(row.session_seq),
        session_start: row.session_start || "",
        last_activity_at: row.last_activity_at || "",
        status: row.status || "",
        app_version: row.app_version || "",
        content_version: row.content_version || "",
        current_card_number:
          row.current_card_number == null ? "" : String(row.current_card_number),
        device_trail: row.device_trail || "",
        progress_json: row.progress_json || "",
        progress_saved_at: row.progress_saved_at || ""
      },
      progress
    }
  };
}


export async function submitTrainingSession(env, input) {
  if (JSON.stringify(input).length > MAX_PAYLOAD_CHARS) {
    return {
      status: 413,
      body: {
        ok: false,
        error: "This submission is too large.",
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
        error: "Too many events in one submission.",
        code: "PAYLOAD_TOO_LARGE",
        retryable: false
      }
    };
  }

  const auth = await authorizeAccessCode(env, input, "training");
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

  const operation = "training:submitSession";
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
          error: "This request identifier was already used for a different submission.",
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

  const existingSubmission = await env.STUDY_DB
    .prepare(
      `SELECT submitted_at
       FROM training_submissions
       WHERE participant_id = ?1 AND session_id = ?2
       LIMIT 1`
    )
    .bind(participantId, sessionId)
    .first();

  if (existingSubmission) {
    const live = await env.STUDY_DB
      .prepare(
        `SELECT revision
         FROM training_live_sessions
         WHERE participant_id = ?1 AND session_id = ?2
         LIMIT 1`
      )
      .bind(participantId, sessionId)
      .first();

    return {
      status: 200,
      body: {
        ok: true,
        submitted: true,
        revision: Number(live?.revision || 0),
        duplicate: true
      }
    };
  }

  const previous = await env.STUDY_DB
    .prepare(
      `SELECT *
       FROM training_live_sessions
       WHERE participant_id = ?1 AND session_id = ?2
       LIMIT 1`
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

  if (previous?.status === "submitted") {
    return {
      status: 409,
      body: {
        ok: false,
        error: "The session is closed but its immutable submission is missing.",
        code: "SUBMISSION_NOT_COMMITTED",
        retryable: true,
        revision: Number(previous?.revision || 0)
      }
    };
  }

  const currentRevision = Number(previous?.revision || 0);

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
  const snapshot = extractItemSnapshot(input);

  const progressJson =
    input.progress_json !== undefined
      ? jsonText(input.progress_json, previous?.progress_json || null)
      : previous?.progress_json || null;

  const incomingDeviceLabel = textValue(input.device_label, input.deviceLabel);
  const incomingCardNumber = textValue(
    input.device_card_number,
    input.deviceCardNumber,
    input.current_card_number,
    input.currentCardNumber
  );

  const deviceTrail =
    input.device_trail !== undefined || input.deviceTrail !== undefined
      ? textValue(input.device_trail, input.deviceTrail)
      : appendDeviceTrail(
          previous?.device_trail || "",
          incomingDeviceLabel,
          incomingCardNumber
        );

  const eventBuild = buildEventStatements(
    env.STUDY_DB,
    input,
    participantId,
    sessionId,
    requestId,
    nextRevision,
    previous?.last_event_at || ""
  );

  const lastEventAt =
    eventBuild.lastEventAt || previous?.last_event_at || null;
  const lastActivityAt = textValue(
    input.last_activity_at,
    input.lastActivityAt,
    input.t,
    previous?.last_activity_at,
    now
  );

  const responseBody = {
    ok: true,
    submitted: true,
    revision: nextRevision,
    duplicate: false
  };

  const liveStatement = env.STUDY_DB
    .prepare(
      `INSERT INTO training_live_sessions (
         record_id, participant_id, session_id, session_seq, revision,
         last_save_id, last_event_at, session_start, last_activity_at, status,
         app_version, content_version, current_card_number, device_trail,
         progress_json, progress_saved_at, last_event_type, created_at, updated_at
       )
       VALUES (
         ?1, ?2, ?3, ?4, ?5, ?6, ?7, ?8, ?9, 'submitted',
         ?10, ?11, ?12, ?13, ?14, ?15, 'session_submitted', ?16, ?17
       )
       ON CONFLICT(participant_id, session_id) DO UPDATE SET
         session_seq = COALESCE(excluded.session_seq, training_live_sessions.session_seq),
         revision = excluded.revision,
         last_save_id = excluded.last_save_id,
         last_event_at = COALESCE(excluded.last_event_at, training_live_sessions.last_event_at),
         session_start = COALESCE(excluded.session_start, training_live_sessions.session_start),
         last_activity_at = excluded.last_activity_at,
         status = 'submitted',
         app_version = COALESCE(excluded.app_version, training_live_sessions.app_version),
         content_version = COALESCE(excluded.content_version, training_live_sessions.content_version),
         current_card_number = COALESCE(excluded.current_card_number, training_live_sessions.current_card_number),
         device_trail = COALESCE(excluded.device_trail, training_live_sessions.device_trail),
         progress_json = COALESCE(excluded.progress_json, training_live_sessions.progress_json),
         progress_saved_at = COALESCE(excluded.progress_saved_at, training_live_sessions.progress_saved_at),
         last_event_type = 'session_submitted',
         updated_at = excluded.updated_at
       WHERE training_live_sessions.revision = ?18
         AND training_live_sessions.status <> 'submitted'
       RETURNING revision`
    )
    .bind(
      previous?.record_id || crypto.randomUUID(),
      participantId,
      sessionId,
      numberOrNull(input.session_seq ?? input.sessionSeq ?? previous?.session_seq),
      nextRevision,
      requestId,
      lastEventAt,
      textValue(input.session_start, input.sessionStart, previous?.session_start) || null,
      lastActivityAt,
      textValue(input.app_version, input.appVersion, previous?.app_version) || null,
      textValue(input.content_version, input.contentVersion, previous?.content_version) || null,
      numberOrNull(
        input.current_card_number ??
          input.currentCardNumber ??
          input.device_card_number ??
          input.deviceCardNumber ??
          previous?.current_card_number
      ),
      deviceTrail || null,
      progressJson,
      textValue(
        input.progress_saved_at,
        input.progressSavedAt,
        lastActivityAt,
        previous?.progress_saved_at
      ) || null,
      previous?.created_at || now,
      now,
      baseRevision
    );

  const itemStatements = [];
  if (snapshot.hasSnapshot) {
    const contentVersion =
      textValue(
        input.content_version,
        input.contentVersion,
        previous?.content_version
      ) || null;

    for (let i = 1; i <= snapshot.itemCount; i += 1) {
      itemStatements.push(
        env.STUDY_DB
          .prepare(
            `INSERT INTO training_live_items (
               record_id, participant_id, session_id, content_version,
               item_number, response_text, response_ms, updated_at
             )
             SELECT ?1, ?2, ?3, ?4, ?5, ?6, ?7, ?8
             WHERE EXISTS (
               SELECT 1 FROM training_live_sessions
               WHERE participant_id = ?2
                 AND session_id = ?3
                 AND revision = ?9
                 AND last_save_id = ?10
                 AND status = 'submitted'
             )
             ON CONFLICT(participant_id, session_id, item_number) DO UPDATE SET
               content_version = excluded.content_version,
               response_text = excluded.response_text,
               response_ms = excluded.response_ms,
               updated_at = excluded.updated_at`
          )
          .bind(
            crypto.randomUUID(),
            participantId,
            sessionId,
            contentVersion,
            i,
            snapshot.responses[i - 1],
            snapshot.msValues[i - 1],
            now,
            nextRevision,
            requestId
          )
      );
    }
  }

  const details = { ...input, user_id: participantId };
  delete details.batch_events;

  const submissionStatement = env.STUDY_DB
    .prepare(
      `INSERT INTO training_submissions (
         record_id, participant_id, session_id, session_seq, session_start,
         session_end, duration_ms, duration_formatted, active_seconds,
         total_questions, correct_first, event_count, item_count, app_version,
         content_version, current_card_number, device_trail, submitted_at,
         details_json, created_at
       )
       SELECT
         ?1, ?2, ?3, ?4, ?5,
         ?6, ?7, ?8, ?9,
         ?10, ?11, ?12, ?13, ?14,
         ?15, ?16, ?17, ?18,
         ?19, ?18
       WHERE EXISTS (
         SELECT 1 FROM training_live_sessions
         WHERE participant_id = ?2
           AND session_id = ?3
           AND revision = ?20
           AND last_save_id = ?21
           AND status = 'submitted'
       )`
    )
    .bind(
      crypto.randomUUID(),
      participantId,
      sessionId,
      numberOrNull(input.session_seq ?? input.sessionSeq ?? previous?.session_seq),
      textValue(input.session_start, input.sessionStart, previous?.session_start) || null,
      textValue(input.session_end, input.sessionEnd, now),
      numberOrNull(input.duration_ms),
      textValue(input.duration_formatted) || null,
      numberOrNull(input.active_seconds),
      numberOrNull(input.total_questions),
      numberOrNull(input.correct_first),
      Array.isArray(input.events)
        ? input.events.length
        : numberOrNull(input.event_count),
      numberOrNull(input.item_count ?? snapshot.itemCount),
      textValue(input.app_version, input.appVersion, previous?.app_version) || null,
      textValue(
        input.content_version,
        input.contentVersion,
        previous?.content_version
      ) || null,
      numberOrNull(
        input.current_card_number ??
          input.currentCardNumber ??
          input.device_card_number ??
          input.deviceCardNumber ??
          previous?.current_card_number
      ),
      deviceTrail || null,
      now,
      JSON.stringify(details),
      nextRevision,
      requestId
    );

  const submissionItemsStatement = env.STUDY_DB
    .prepare(
      `INSERT INTO training_submission_items (
         record_id, participant_id, session_id, content_version,
         item_number, response_text, response_ms, created_at
       )
       SELECT
         lower(hex(randomblob(16))), i.participant_id, i.session_id,
         i.content_version, i.item_number, i.response_text, i.response_ms, ?1
       FROM training_live_items i
       WHERE i.participant_id = ?2
         AND i.session_id = ?3
         AND EXISTS (
           SELECT 1 FROM training_live_sessions l
           WHERE l.participant_id = ?2
             AND l.session_id = ?3
             AND l.revision = ?4
             AND l.last_save_id = ?5
             AND l.status = 'submitted'
         )`
    )
    .bind(now, participantId, sessionId, nextRevision, requestId);

  const receiptStatement = env.STUDY_DB
    .prepare(
      `INSERT OR IGNORE INTO request_receipts (
         request_id, operation, request_hash, response_status,
         response_json, created_at
       )
       SELECT ?1, ?2, ?3, 200, ?4, ?5
       WHERE EXISTS (
         SELECT 1 FROM training_submissions
         WHERE participant_id = ?6 AND session_id = ?7
       )
       AND EXISTS (
         SELECT 1 FROM training_live_sessions
         WHERE participant_id = ?6
           AND session_id = ?7
           AND revision = ?8
           AND last_save_id = ?1
           AND status = 'submitted'
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
    ...itemStatements,
    submissionStatement,
    submissionItemsStatement,
    receiptStatement
  ]);

  const committedRevision = batchResults?.[0]?.results?.[0]?.revision;

  if (Number(committedRevision) !== nextRevision) {
    const latest = await env.STUDY_DB
      .prepare(
        `SELECT revision, status
         FROM training_live_sessions
         WHERE participant_id = ?1 AND session_id = ?2`
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

  const committedSubmission = await env.STUDY_DB
    .prepare(
      `SELECT 1 AS ok
       FROM training_submissions
       WHERE participant_id = ?1 AND session_id = ?2
       LIMIT 1`
    )
    .bind(participantId, sessionId)
    .first();

  if (committedSubmission?.ok !== 1) {
    return {
      status: 500,
      body: {
        ok: false,
        error: "The immutable training submission was not created.",
        code: "SUBMISSION_NOT_COMMITTED",
        retryable: true
      }
    };
  }

  return { status: 200, body: responseBody };
}


export async function appendTrainingEvent(env, input) {
  const auth = await authorizeAccessCode(env, input, "training");
  if (!auth.ok) return { status: auth.status, body: auth.body };

  const participantId = auth.participantId;
  const sessionId = textValue(input.session_id, input.sessionId);
  const requestId = textValue(input.request_id, input.event_id);
  const eventId = textValue(input.event_id, input.request_id);
  const eventType = textValue(input.event_type, input.eventType, input.type);

  if (!sessionId || !requestId || !eventId || !eventType || requestId.length > 120 || eventId.length > 160) {
    return {
      status: 400,
      body: {
        ok: false,
        error: "Session, request, event identifier, and event type are required.",
        code: "INVALID_REQUEST",
        retryable: false
      }
    };
  }

  const operation = "training:appendEvent";
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
          error: "This request identifier was already used for a different operation.",
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

  const session = await env.STUDY_DB
    .prepare(
      `SELECT session_seq, status, last_event_at
       FROM training_live_sessions
       WHERE participant_id = ?1 AND session_id = ?2
       LIMIT 1`
    )
    .bind(participantId, sessionId)
    .first();

  if (!session) {
    if (!(await sessionWasAllocated(env.STUDY_DB, participantId, sessionId))) {
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

  const timestamp = eventTimestamp(input, input.last_activity_at);
  const latestEvent = await env.STUDY_DB
    .prepare(
      `SELECT event_timestamp
       FROM training_events
       WHERE participant_id = ?1 AND session_id = ?2
       ORDER BY event_timestamp DESC
       LIMIT 1`
    )
    .bind(participantId, sessionId)
    .first();

  const previousTimestamp = latestEvent?.event_timestamp || session?.last_event_at || "";
  const previousMs = previousTimestamp ? Date.parse(previousTimestamp) : NaN;
  const currentMs = Date.parse(timestamp);
  const msSincePrevious =
    Number.isFinite(previousMs) && Number.isFinite(currentMs)
      ? Math.max(0, currentMs - previousMs)
      : null;

  const responseBody = {
    ok: true,
    appended: true,
    event_id: eventId,
    duplicate: false
  };

  const now = new Date().toISOString();

  const eventStatement = env.STUDY_DB
    .prepare(
      `INSERT OR IGNORE INTO training_events (
         record_id, event_id, batch_id, event_timestamp, ms_since_previous,
         participant_id, session_id, session_seq, event_type, section_index,
         card_index, detail_text, detail_json, device_label, created_at
       )
       VALUES (?1, ?2, ?3, ?4, ?5, ?6, ?7, ?8, ?9, ?10, ?11, ?12, ?13, ?14, ?15)`
    )
    .bind(
      crypto.randomUUID(),
      eventId,
      participantId + ":" + sessionId + ":" + requestId,
      timestamp,
      msSincePrevious,
      participantId,
      sessionId,
      numberOrNull(input.session_seq ?? input.sessionSeq ?? session?.session_seq),
      eventType,
      numberOrNull(input.section_index ?? input.sectionIndex),
      numberOrNull(
        input.card_index ??
          input.cardIndex ??
          input.current_card_number ??
          input.currentCardNumber
      ),
      textValue(input.detail_text, input.detailText) || null,
      jsonText(input.detail_json ?? input.detailJson ?? input.details, null),
      textValue(input.device_label, input.deviceLabel) || null,
      now
    );

  const liveStatement = env.STUDY_DB
    .prepare(
      `UPDATE training_live_sessions
       SET last_event_at = CASE
             WHEN last_event_at IS NULL OR last_event_at = '' OR last_event_at < ?1 THEN ?1
             ELSE last_event_at
           END,
           last_event_type = ?2,
           last_activity_at = CASE
             WHEN last_activity_at IS NULL OR last_activity_at = '' OR last_activity_at < ?1 THEN ?1
             ELSE last_activity_at
           END,
           updated_at = ?3
       WHERE participant_id = ?4 AND session_id = ?5`
    )
    .bind(timestamp, eventType, now, participantId, sessionId);

  const receiptStatement = env.STUDY_DB
    .prepare(
      `INSERT OR IGNORE INTO request_receipts
       (request_id, operation, request_hash, response_status, response_json, created_at)
       VALUES (?1, ?2, ?3, 200, ?4, ?5)`
    )
    .bind(
      requestId,
      operation,
      requestHash,
      JSON.stringify(responseBody),
      now
    );

  const results = await env.STUDY_DB.batch([
    eventStatement,
    liveStatement,
    receiptStatement
  ]);

  const inserted = Number(results?.[0]?.meta?.changes || 0) > 0;

  if (!inserted) {
    const existing = await env.STUDY_DB
      .prepare(
        `SELECT participant_id, session_id
         FROM training_events
         WHERE event_id = ?1`
      )
      .bind(eventId)
      .first();

    if (
      existing &&
      (existing.participant_id !== participantId ||
        existing.session_id !== sessionId)
    ) {
      return {
        status: 409,
        body: {
          ok: false,
          error: "This event identifier was already used for a different request.",
          code: "REQUEST_ID_CONFLICT",
          retryable: false
        }
      };
    }

    return {
      status: 200,
      body: { ...responseBody, duplicate: true }
    };
  }

  return { status: 200, body: responseBody };
}
