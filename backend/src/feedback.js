import { authorizeAccessCode } from "./auth.js";

function firstText(input, keys) {
  for (const key of keys) {
    const value = input?.[key];
    if (value !== undefined && value !== null && String(value).trim()) {
      return String(value).trim();
    }
  }
  return "";
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

export async function saveAqgFeedback(env, input) {
  const auth = await authorizeAccessCode(env, input, "aqg");
  if (!auth.ok) return { status: auth.status, body: auth.body };

  const participantId = auth.participantId;
  const sessionId = firstText(input, [
    "session_id",
    "sessionId",
    "source_session_id",
    "sourceSessionId"
  ]);

  const textFeedback = firstText(input, [
    "text_feedback",
    "textFeedback",
    "session_notes",
    "sessionNotes"
  ]);

  const audioObjectKey = firstText(input, [
    "audio_object_key",
    "audioObjectKey"
  ]);

  const audioOriginalFilename = firstText(input, [
    "audio_original_filename",
    "audioOriginalFilename",
    "filename"
  ]);

  const audioMimeType = firstText(input, [
    "audio_mime_type",
    "audioMimeType",
    "mimeType"
  ]);

  const durationRaw =
    input?.audio_duration_seconds ??
    input?.audioDurationSeconds ??
    input?.audio_duration ??
    input?.audioDuration;

  const audioDurationSeconds =
    durationRaw === undefined || durationRaw === null || durationRaw === ""
      ? null
      : Number(durationRaw);

  if (!textFeedback && !audioObjectKey) {
    return {
      status: 400,
      body: {
        ok: false,
        error: "No feedback content provided.",
        code: "INVALID_REQUEST",
        retryable: false
      }
    };
  }

  if (
    audioDurationSeconds !== null &&
    (!Number.isFinite(audioDurationSeconds) || audioDurationSeconds < 0)
  ) {
    return {
      status: 400,
      body: {
        ok: false,
        error: "Audio duration is invalid.",
        code: "INVALID_REQUEST",
        retryable: false
      }
    };
  }

  if (
    sessionId &&
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

  const feedbackId = firstText(input, [
    "feedback_id",
    "feedbackId",
    "request_id"
  ]) || crypto.randomUUID();

  if (feedbackId.length > 120) {
    return {
      status: 400,
      body: {
        ok: false,
        error: "Feedback identifier is invalid.",
        code: "INVALID_REQUEST",
        retryable: false
      }
    };
  }

  const existing = await env.STUDY_DB
    .prepare(
      `SELECT feedback_id, participant_id, session_id
       FROM aqg_feedback
       WHERE feedback_id = ?1`
    )
    .bind(feedbackId)
    .first();

  if (existing) {
    if (
      existing.participant_id !== participantId ||
      String(existing.session_id || "") !== sessionId
    ) {
      return {
        status: 409,
        body: {
          ok: false,
          error: "This feedback identifier was already used for a different request.",
          code: "REQUEST_ID_CONFLICT",
          retryable: false
        }
      };
    }

    return {
      status: 200,
      body: {
        ok: true,
        feedback_id: feedbackId,
        duplicate: true
      }
    };
  }

  const now = new Date().toISOString();
  const recordId = crypto.randomUUID();
  const notificationId = crypto.randomUUID();

  const payload = JSON.stringify({
    feedback_id: feedbackId,
    participant_id: participantId,
    session_id: sessionId,
    text_feedback: textFeedback,
    audio_object_key: audioObjectKey,
    audio_original_filename: audioOriginalFilename,
    audio_mime_type: audioMimeType,
    audio_duration_seconds: audioDurationSeconds,
    submitted_at: now
  });

  await env.STUDY_DB.batch([
    env.STUDY_DB
      .prepare(
        `INSERT INTO aqg_feedback
           (record_id, feedback_id, participant_id, session_id, submitted_at,
            text_feedback, audio_object_key, audio_original_filename,
            audio_mime_type, audio_duration_seconds, notification_status, created_at)
         VALUES
           (?1, ?2, ?3, NULLIF(?4, ''), ?5, NULLIF(?6, ''),
            NULLIF(?7, ''), NULLIF(?8, ''), NULLIF(?9, ''), ?10, 'pending', ?5)`
      )
      .bind(
        recordId,
        feedbackId,
        participantId,
        sessionId,
        now,
        textFeedback,
        audioObjectKey,
        audioOriginalFilename,
        audioMimeType,
        audioDurationSeconds
      ),
    env.STUDY_DB
      .prepare(
        `INSERT INTO notification_outbox
           (notification_id, notification_type, participant_id, session_id,
            payload_json, status, attempt_count, created_at)
         VALUES
           (?1, 'aqg_feedback', ?2, NULLIF(?3, ''), ?4, 'pending', 0, ?5)`
      )
      .bind(
        notificationId,
        participantId,
        sessionId,
        payload,
        now
      )
  ]);

  return {
    status: 200,
    body: {
      ok: true,
      feedback_id: feedbackId,
      duplicate: false,
      notification_queued: true
    }
  };
}
