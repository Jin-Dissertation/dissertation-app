function relayErrorMessage(value) {
  const text = String(value || "").trim();
  return text ? text.slice(0, 1000) : "Notification delivery failed.";
}

function parsePayloadJson(value) {
  if (!value) return {};
  try {
    const parsed = JSON.parse(value);
    return parsed && typeof parsed === "object" && !Array.isArray(parsed)
      ? parsed
      : {};
  } catch {
    return {};
  }
}

async function markNotificationFailure(env, notificationId, error) {
  const now = new Date().toISOString();
  const message = relayErrorMessage(error?.message || error);

  await env.STUDY_DB
    .prepare(
      `UPDATE notification_outbox
       SET status = 'pending',
           attempt_count = attempt_count + 1,
           last_error = ?1
       WHERE notification_id = ?2
         AND status <> 'sent'`
    )
    .bind(message, notificationId)
    .run();

  return {
    ok: false,
    notification_id: notificationId,
    error: message,
    attempted_at: now
  };
}

export async function deliverNotification(env, notificationId) {
  const id = String(notificationId || "").trim();
  if (!id) {
    return { ok: false, error: "Missing notification identifier." };
  }

  if (!env.NOTIFICATION_RELAY_URL || !env.NOTIFICATION_RELAY_SECRET) {
    return markNotificationFailure(
      env,
      id,
      "Notification relay is not configured."
    );
  }

  const row = await env.STUDY_DB
    .prepare(
      `SELECT notification_id, notification_type, participant_id, session_id,
              payload_json, status
       FROM notification_outbox
       WHERE notification_id = ?1
       LIMIT 1`
    )
    .bind(id)
    .first();

  if (!row) {
    return { ok: false, notification_id: id, error: "Notification not found." };
  }

  if (row.status === "sent") {
    return { ok: true, notification_id: id, sent: true, duplicate: true };
  }

  const relayBody = {
    action: "notificationRelay",
    relay_secret: env.NOTIFICATION_RELAY_SECRET,
    notification_id: row.notification_id,
    notification_type: row.notification_type,
    participant_id: row.participant_id || "",
    session_id: row.session_id || "",
    payload: parsePayloadJson(row.payload_json)
  };

  try {
    const response = await fetch(env.NOTIFICATION_RELAY_URL, {
      method: "POST",
      headers: {
        "content-type": "text/plain;charset=utf-8"
      },
      body: JSON.stringify(relayBody),
      redirect: "follow"
    });

    const responseText = await response.text();
    let relayResult = {};

    try {
      relayResult = responseText ? JSON.parse(responseText) : {};
    } catch {
      relayResult = {};
    }

    if (!response.ok || relayResult?.ok !== true || relayResult?.sent !== true) {
      throw new Error(
        relayResult?.error ||
          `Notification relay returned HTTP ${response.status}.`
      );
    }

    const sentAt = new Date().toISOString();

    await env.STUDY_DB
      .prepare(
        `UPDATE notification_outbox
         SET status = 'sent',
             attempt_count = attempt_count + 1,
             sent_at = ?1,
             last_error = NULL
         WHERE notification_id = ?2`
      )
      .bind(sentAt, id)
      .run();

    if (row.notification_type === "training_feedback") {
      await env.STUDY_DB
        .prepare(
          `UPDATE training_feedback
           SET notification_status = 'sent'
           WHERE feedback_id = json_extract(?1, '$.feedback_id')`
        )
        .bind(row.payload_json)
        .run();
    }

    if (row.notification_type === "aqg_feedback") {
      await env.STUDY_DB
        .prepare(
          `UPDATE aqg_feedback
           SET notification_status = 'sent'
           WHERE feedback_id = json_extract(?1, '$.feedback_id')`
        )
        .bind(row.payload_json)
        .run();
    }

    return {
      ok: true,
      notification_id: id,
      sent: true,
      duplicate: Boolean(relayResult.duplicate)
    };
  } catch (error) {
    return markNotificationFailure(env, id, error);
  }
}

export async function flushPendingNotifications(env, limit = 10) {
  const safeLimit = Math.max(1, Math.min(25, Number(limit) || 10));

  const result = await env.STUDY_DB
    .prepare(
      `SELECT notification_id
       FROM notification_outbox
       WHERE status = 'pending'
       ORDER BY created_at ASC
       LIMIT ?1`
    )
    .bind(safeLimit)
    .all();

  const notifications = result?.results || [];
  const outcomes = [];

  for (const row of notifications) {
    outcomes.push(await deliverNotification(env, row.notification_id));
  }

  return {
    ok: true,
    attempted: outcomes.length,
    outcomes
  };
}
