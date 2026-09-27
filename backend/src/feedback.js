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

function finiteNumberOrNull(...values) {
  for (const value of values) {
    if (value === undefined || value === null || value === "") continue;
    const n = Number(value);
    if (Number.isFinite(n)) return n;
  }
  return null;
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


async function trainingSessionBelongsToParticipant(db, participantId, sessionId) {
  if (!sessionId) return true;

  const row = await db
    .prepare(
      `SELECT 1 AS ok
       WHERE EXISTS (
         SELECT 1
         FROM request_receipts
         WHERE operation = 'training:createSessionId'
           AND json_extract(response_json, '$.user_id') = ?1
           AND json_extract(response_json, '$.session_id') = ?2
       )
       OR EXISTS (
         SELECT 1 FROM training_live_sessions
         WHERE participant_id = ?1 AND session_id = ?2
       )
       OR EXISTS (
         SELECT 1 FROM training_submissions
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
      `SELECT feedback_id, participant_id, session_id, text_feedback,
              audio_object_key, audio_original_filename, audio_mime_type,
              audio_duration_seconds
       FROM aqg_feedback
       WHERE feedback_id = ?1`
    )
    .bind(feedbackId)
    .first();

  if (existing) {
    const storedDuration =
      existing.audio_duration_seconds === null ||
      existing.audio_duration_seconds === undefined ||
      existing.audio_duration_seconds === ""
        ? null
        : Number(existing.audio_duration_seconds);

    const sameRequest =
      existing.participant_id === participantId &&
      String(existing.session_id || "") === sessionId &&
      String(existing.text_feedback || "") === textFeedback &&
      String(existing.audio_object_key || "") === audioObjectKey &&
      String(existing.audio_original_filename || "") === audioOriginalFilename &&
      String(existing.audio_mime_type || "") === audioMimeType &&
      storedDuration === audioDurationSeconds;

    if (!sameRequest) {
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
      notification_queued: true,
      notification_id: notificationId
    }
  };
}


export async function saveTrainingFeedback(env, input) {
  const auth = await authorizeAccessCode(env, input, "training");
  if (!auth.ok) return { status: auth.status, body: auth.body };

  const participantId = auth.participantId;
  const sessionId = firstText(input, ["session_id", "sessionId"]);
  const feedbackId =
    firstText(input, ["feedback_id", "feedbackId", "request_id"]) ||
    crypto.randomUUID();
  const textFeedback = firstText(input, [
    "text_feedback",
    "textFeedback",
    "feedback_text",
    "feedbackText",
    "detail_text",
    "detailText"
  ]);

  if (!textFeedback) {
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
    sessionId &&
    !(await trainingSessionBelongsToParticipant(
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

  const suppliedSessionSeq = finiteNumberOrNull(
    input.session_seq,
    input.sessionSeq
  );
  const sectionIndex = finiteNumberOrNull(
    input.section_index,
    input.sectionIndex
  );
  const cardIndex = finiteNumberOrNull(
    input.card_index,
    input.cardIndex,
    input.current_card_number,
    input.currentCardNumber
  );
  const sectionTitle = firstText(input, ["section_title", "sectionTitle"]);
  const feedbackSource = firstText(input, [
    "feedback_source",
    "feedbackSource"
  ]);
  const deviceLabel = firstText(input, ["device_label", "deviceLabel"]);

  const sessionRow = sessionId
    ? await env.STUDY_DB
        .prepare(
          `SELECT session_seq, last_event_at
           FROM training_live_sessions
           WHERE participant_id = ?1 AND session_id = ?2
           LIMIT 1`
        )
        .bind(participantId, sessionId)
        .first()
    : null;

  const sessionSeq =
    suppliedSessionSeq ?? finiteNumberOrNull(sessionRow?.session_seq);

  const existing = await env.STUDY_DB
    .prepare(
      `SELECT participant_id, session_id, session_seq, section_index,
              section_title, card_index, feedback_source, text_feedback,
              device_label
       FROM training_feedback
       WHERE feedback_id = ?1`
    )
    .bind(feedbackId)
    .first();

  if (existing) {
    const sameRequest =
      existing.participant_id === participantId &&
      String(existing.session_id || "") === sessionId &&
      finiteNumberOrNull(existing.session_seq) === sessionSeq &&
      finiteNumberOrNull(existing.section_index) === sectionIndex &&
      String(existing.section_title || "") === sectionTitle &&
      finiteNumberOrNull(existing.card_index) === cardIndex &&
      String(existing.feedback_source || "") === feedbackSource &&
      String(existing.text_feedback || "") === textFeedback &&
      String(existing.device_label || "") === deviceLabel;

    if (!sameRequest) {
      return {
        status: 409,
        body: {
          ok: false,
          error:
            "This feedback identifier was already used for a different request.",
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
  const notificationId = crypto.randomUUID();

  let msSincePrevious = null;
  if (sessionId) {
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

    const previousTimestamp =
      latestEvent?.event_timestamp || sessionRow?.last_event_at || "";
    const previousMs = previousTimestamp
      ? Date.parse(previousTimestamp)
      : NaN;
    const currentMs = Date.parse(now);

    if (Number.isFinite(previousMs) && Number.isFinite(currentMs)) {
      msSincePrevious = Math.max(0, currentMs - previousMs);
    }
  }

  const payload = JSON.stringify({
    feedback_id: feedbackId,
    participant_id: participantId,
    session_id: sessionId,
    session_seq: sessionSeq,
    section_index: sectionIndex,
    section_title: sectionTitle,
    card_index: cardIndex,
    feedback_source: feedbackSource,
    text_feedback: textFeedback,
    device_label: deviceLabel,
    submitted_at: now
  });

  const statements = [
    env.STUDY_DB.prepare(
      `INSERT INTO training_feedback (
         record_id, feedback_id, participant_id, session_id, session_seq,
         submitted_at, section_index, section_title, card_index,
         feedback_source, text_feedback, device_label,
         notification_status, created_at
       )
       VALUES (
         ?1, ?2, ?3, NULLIF(?4, ''), ?5, ?6, ?7, NULLIF(?8, ''),
         ?9, NULLIF(?10, ''), ?11, NULLIF(?12, ''), 'pending', ?6
       )`
    ).bind(
      crypto.randomUUID(),
      feedbackId,
      participantId,
      sessionId,
      sessionSeq,
      now,
      sectionIndex,
      sectionTitle,
      cardIndex,
      feedbackSource,
      textFeedback,
      deviceLabel
    ),

    env.STUDY_DB.prepare(
      `INSERT INTO notification_outbox (
         notification_id, notification_type, participant_id, session_id,
         payload_json, status, attempt_count, created_at
       )
       VALUES (
         ?1, 'training_feedback', ?2, NULLIF(?3, ''), ?4, 'pending', 0, ?5
       )`
    ).bind(
      notificationId,
      participantId,
      sessionId,
      payload,
      now
    )
  ];

  if (sessionId) {
    statements.push(
      env.STUDY_DB.prepare(
        `INSERT OR IGNORE INTO training_events (
           record_id, event_id, batch_id, event_timestamp, ms_since_previous,
           participant_id, session_id, session_seq, event_type,
           section_index, card_index, detail_text, detail_json,
           device_label, created_at
         )
         VALUES (
           ?1, ?2, ?3, ?4, ?5, ?6, ?7, ?8, 'feedback_submitted',
           ?9, ?10, ?11, ?12, NULLIF(?13, ''), ?4
         )`
      ).bind(
        crypto.randomUUID(),
        `feedback:${feedbackId}`,
        `feedback:${feedbackId}`,
        now,
        msSincePrevious,
        participantId,
        sessionId,
        sessionSeq,
        sectionIndex,
        cardIndex,
        textFeedback.slice(0, 500),
        JSON.stringify({
          feedback_text: textFeedback,
          section_title: sectionTitle,
          feedback_source: feedbackSource
        }),
        deviceLabel
      ),

      env.STUDY_DB.prepare(
        `UPDATE training_live_sessions
         SET last_event_at = ?1,
             last_event_type = 'feedback_submitted',
             last_activity_at = CASE
               WHEN last_activity_at IS NULL OR last_activity_at = '' OR last_activity_at < ?1
               THEN ?1 ELSE last_activity_at
             END,
             updated_at = ?1
         WHERE participant_id = ?2 AND session_id = ?3`
      ).bind(now, participantId, sessionId)
    );
  }

  await env.STUDY_DB.batch(statements);

  return {
    status: 200,
    body: {
      ok: true,
      feedback_id: feedbackId,
      duplicate: false,
      notification_queued: true,
      notification_id: notificationId
    }
  };
}

