import { authorizeAccessCode } from "./auth.js";

const CURRENT_TRAINING_CONTENT_VERSION = "2026-04-27-strict-content-a";

function contentKey(version) {
  return `content/${version}/training-content.json`;
}

export async function loadTrainingContent(env, input) {
  const auth = await authorizeAccessCode(env, input, "training");
  if (!auth.ok) return { status: auth.status, body: auth.body };

  const requestedVersion = String(
    input?.content_version ?? input?.contentVersion ?? CURRENT_TRAINING_CONTENT_VERSION
  ).trim();

  if (requestedVersion !== CURRENT_TRAINING_CONTENT_VERSION) {
    return {
      status: 404,
      body: {
        ok: false,
        error: "Training content version was not found.",
        code: "CONTENT_NOT_FOUND",
        retryable: false
      }
    };
  }

  const object = await env.TRAINING_MEDIA.get(contentKey(requestedVersion));

  if (!object) {
    return {
      status: 503,
      body: {
        ok: false,
        error: "Training content is temporarily unavailable.",
        code: "CONTENT_UNAVAILABLE",
        retryable: true
      }
    };
  }

  let parsed;
  try {
    parsed = JSON.parse(await object.text());
  } catch {
    return {
      status: 500,
      body: {
        ok: false,
        error: "Training content could not be read.",
        code: "CONTENT_INVALID",
        retryable: true
      }
    };
  }

  const rows = Array.isArray(parsed?.rows) ? parsed.rows : null;

  if (!rows) {
    return {
      status: 500,
      body: {
        ok: false,
        error: "Training content has an invalid structure.",
        code: "CONTENT_INVALID",
        retryable: true
      }
    };
  }

  return {
    status: 200,
    body: {
      ok: true,
      content_version: requestedVersion,
      rows
    }
  };
}
