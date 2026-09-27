import { authorizeAccessCode } from "./auth.js";

const CURRENT_TRAINING_CONTENT_VERSION = "2026-04-27-strict-content-a";

function contentKey(version) {
  return `content/${version}/training-content.json`;
}

export async function loadTrainingContent(env, input, requestUrl) {
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

  const origin = requestUrl ? new URL(requestUrl).origin : "";
  const resolvedRows = rows.map((row) => {
    const copy = { ...row };

    for (const field of ["media_url", "mobile_image_url"]) {
      const value = String(copy[field] || "").trim();
      if (value.startsWith("images/")) {
        const filename = value.slice("images/".length);
        if (ALLOWED_TRAINING_IMAGES.has(filename) && origin) {
          copy[field] =
            `${origin}/v1/training/media/${requestedVersion}/${filename}`;
        }
      }
    }

    return copy;
  });

  return {
    status: 200,
    body: {
      ok: true,
      content_version: requestedVersion,
      rows: resolvedRows
    }
  };
}


const ALLOWED_TRAINING_IMAGES = new Set([
  "stat-birkhead.png",
  "stat-flaws.png",
  "flawed-question.png",
  "flawed-question-annotated.png",
  "stat-confidence.png"
]);

export async function getTrainingMedia(env, version, filename) {
  if (
    version !== CURRENT_TRAINING_CONTENT_VERSION ||
    !ALLOWED_TRAINING_IMAGES.has(filename)
  ) {
    return new Response("Not found", { status: 404 });
  }

  const key = `content/${version}/images/${filename}`;
  const object = await env.TRAINING_MEDIA.get(key);

  if (!object) {
    return new Response("Not found", { status: 404 });
  }

  return new Response(object.body, {
    status: 200,
    headers: {
      "content-type": "image/png",
      "cache-control": "public, max-age=86400",
      "access-control-allow-origin": "*"
    }
  });
}
