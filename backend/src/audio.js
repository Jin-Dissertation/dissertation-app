/*
 * MAINTAINER GUIDE — PRIVATE AQG AUDIO
 *
 * This module authorizes the participant, confirms the session belongs to
 * them, validates audio type/size, and stores the recording in the private
 * STUDY_AUDIO R2 bucket.
 *
 * R2 audio is archived and verified separately from structured D1 reporting.
 * Do not make the bucket public or delete objects outside the guarded
 * archive/receipt workflow.
 */

import { authorizeAccessCode } from "./auth.js";
import { saveAqgFeedback } from "./feedback.js";

const MAX_AUDIO_BYTES = 8 * 1024 * 1024;
const MAX_BASE64_CHARS = 12 * 1024 * 1024;
const ALLOWED_AUDIO_TYPES = /^audio\/(webm|ogg|mp4|mpeg|wav)(;|$)/i;

function firstText(input, keys) {
  for (const key of keys) {
    const value = input?.[key];
    if (value !== undefined && value !== null && String(value).trim()) {
      return String(value).trim();
    }
  }
  return "";
}

function safePart(value, fallback = "audio") {
  const cleaned = String(value || "")
    .replace(/[^a-zA-Z0-9_.-]/g, "_")
    .replace(/_+/g, "_")
    .replace(/^[_\.]+|[_\.]+$/g, "")
    .slice(0, 180);
  return cleaned || fallback;
}

function decodeBase64(base64) {
  try {
    const binary = atob(base64);
    const bytes = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i += 1) {
      bytes[i] = binary.charCodeAt(i);
    }
    return bytes;
  } catch {
    return null;
  }
}

async function sessionBelongsToParticipant(db, participantId, sessionId) {
  if (!sessionId) return true;

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

export async function uploadAqgAudio(env, input) {
  const auth = await authorizeAccessCode(env, input, "aqg");
  if (!auth.ok) return { status: auth.status, body: auth.body };

  const participantId = auth.participantId;
  const sessionId = firstText(input, ["session_id", "sessionId"]);
  const requestId = firstText(input, ["request_id", "requestId"]);

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

  if (
    !(await sessionBelongsToParticipant(
      env.STUDY_DB,
      participantId,
      sessionId
    ))
  ) {
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

  const rawBase64 = firstText(input, ["audioBase64", "audio_base64"])
    .replace(/^data:.*;base64,/i, "")
    .trim();

  if (!rawBase64 || rawBase64.length > MAX_BASE64_CHARS) {
    return {
      status: 413,
      body: {
        ok: false,
        error: "Audio is missing or exceeds the 8 MB limit.",
        code: "PAYLOAD_TOO_LARGE",
        retryable: false
      }
    };
  }

  const mimeType =
    firstText(input, ["mimeType", "mime_type"]) || "audio/webm";

  if (!ALLOWED_AUDIO_TYPES.test(mimeType)) {
    return {
      status: 400,
      body: {
        ok: false,
        error: "Unsupported audio format.",
        code: "INVALID_REQUEST",
        retryable: false
      }
    };
  }

  const bytes = decodeBase64(rawBase64);
  if (!bytes) {
    return {
      status: 400,
      body: {
        ok: false,
        error: "Audio data is not valid base64.",
        code: "INVALID_REQUEST",
        retryable: false
      }
    };
  }

  if (bytes.byteLength > MAX_AUDIO_BYTES) {
    return {
      status: 413,
      body: {
        ok: false,
        error: "Audio exceeds the 8 MB limit. Please use a shorter recording.",
        code: "PAYLOAD_TOO_LARGE",
        retryable: false
      }
    };
  }

  const originalFilename =
    firstText(input, ["filename", "audio_filename"]) || "feedback.webm";
  const safeFilename = safePart(originalFilename, "feedback.webm");
  const safeRequestId = safePart(requestId, "request");
  const objectKey =
    `aqg/${safePart(participantId, "participant")}/${safePart(sessionId, "session")}/${safeRequestId}_${safeFilename}`;

  const existing = await env.STUDY_AUDIO.head(objectKey);
  let duplicate = Boolean(existing);

  if (!existing) {
    await env.STUDY_AUDIO.put(objectKey, bytes, {
      httpMetadata: {
        contentType: mimeType
      },
      customMetadata: {
        participant_id: participantId,
        session_id: sessionId,
        request_id: requestId,
        original_filename: safeFilename
      }
    });
  }

  const durationRaw =
    input?.audio_duration_seconds ??
    input?.audioDurationSeconds ??
    input?.audio_duration ??
    input?.audioDuration ??
    "";

  let feedbackResult;
  try {
    feedbackResult = await saveAqgFeedback(env, {
      ...input,
      feedback_id: requestId,
      audio_object_key: objectKey,
      audio_original_filename: safeFilename,
      audio_mime_type: mimeType,
      audio_duration_seconds: durationRaw
    });
  } catch (error) {
    if (!duplicate) {
      await env.STUDY_AUDIO.delete(objectKey);
    }
    throw error;
  }

  if (!feedbackResult?.body?.ok) {
    if (!duplicate) {
      await env.STUDY_AUDIO.delete(objectKey);
    }
    return feedbackResult;
  }

  duplicate = duplicate || Boolean(feedbackResult.body.duplicate);

  return {
    status: 200,
    body: {
      ok: true,
      fileId: objectKey,
      audioFileId: objectKey,
      objectKey,
      fileUrl: "",
      filename: safeFilename,
      mimeType,
      duplicate,
      notification_queued: true,
      notification_warning: ""
    }
  };
}
