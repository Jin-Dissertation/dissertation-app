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
