/*
 * MAINTAINER GUIDE — NON-PARTICIPANT AGGREGATE COUNTERS
 *
 * The public/non-participant tool does not create participant study sessions.
 * This module maintains the intentionally limited cumulative button-use counts.
 *
 * These counters are retained across reporting purge cycles and are not treated
 * like participant-level study records.
 */

function cleanButtonId(value) {
  return String(value ?? "").trim();
}

export async function incrementNonparticipantButtonCount(env, input) {
  const buttonId = cleanButtonId(input?.button_id ?? input?.buttonId);

  if (!buttonId) {
    return {
      status: 400,
      body: {
        ok: false,
        error: "button_id is required.",
        code: "INVALID_REQUEST",
        retryable: false
      }
    };
  }

  if (buttonId.length > 120) {
    return {
      status: 400,
      body: {
        ok: false,
        error: "button_id is too long.",
        code: "INVALID_REQUEST",
        retryable: false
      }
    };
  }

  const updatedAt = new Date().toISOString();
  const requestId = String(input?.request_id ?? "").trim();
  if (requestId.length > 120) {
    return { status: 400, body: { ok: false, code: "INVALID_REQUEST", retryable: false } };
  }

  if (requestId) {
    // Receipt and aggregate increment share a transaction. changes() observes
    // only the immediately preceding receipt insert; ignored retries add nothing.
    const results = await env.STUDY_DB.batch([
      env.STUDY_DB.prepare(`INSERT OR IGNORE INTO nonparticipant_press_receipts
        (request_id, button_id, created_at) VALUES (?1, ?2, ?3)`)
        .bind(requestId, buttonId, updatedAt),
      env.STUDY_DB.prepare(`INSERT INTO nonparticipant_button_counts (button_id, press_count, updated_at)
        SELECT ?1, 1, ?2 WHERE changes() = 1
        ON CONFLICT(button_id) DO UPDATE SET
          press_count = nonparticipant_button_counts.press_count + 1,
          updated_at = excluded.updated_at`).bind(buttonId, updatedAt)
    ]);
    if (Number(results[0].meta.changes) === 0) {
      const prior = await env.STUDY_DB.prepare("SELECT button_id FROM nonparticipant_press_receipts WHERE request_id = ?1")
        .bind(requestId).first();
      if (prior?.button_id !== buttonId) {
        return { status: 409, body: { ok: false, code: "REQUEST_ID_CONFLICT", retryable: false } };
      }
      return { status: 200, body: { ok: true, duplicate: true } };
    }
    return { status: 200, body: { ok: true, duplicate: false } };
  }

  // Compatibility for older non-participant pages: each POST is one press.
  await env.STUDY_DB
    .prepare(
      `INSERT INTO nonparticipant_button_counts
         (button_id, press_count, updated_at)
       VALUES (?1, 1, ?2)
       ON CONFLICT(button_id) DO UPDATE SET
         press_count = nonparticipant_button_counts.press_count + 1,
         updated_at = excluded.updated_at`
    )
    .bind(buttonId, updatedAt)
    .run();

  return {
    status: 200,
    body: {
      ok: true
    }
  };
}
