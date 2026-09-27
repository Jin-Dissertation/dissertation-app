import { validateAccessCode } from "./auth.js";
import { createSessionId, createContextId } from "./sessions.js";
import { saveAqgLiveSession, getLatestAqgLiveSession, submitAqgSession, getLatestAqgSubmittedSettings } from "./aqg-save.js";
import { saveAqgFeedback } from "./feedback.js";
import { uploadAqgAudio } from "./audio.js";
import { saveTrainingLiveSession, getLatestTrainingLiveSession } from "./training-save.js";

function corsHeaders() {
  return {
    "access-control-allow-origin": "*",
    "access-control-allow-methods": "GET, POST, OPTIONS",
    "access-control-allow-headers": "content-type"
  };
}

function json(data, status = 200) {
  return new Response(JSON.stringify(data, null, 2), {
    status,
    headers: {
      "content-type": "application/json; charset=utf-8",
      ...corsHeaders()
    }
  });
}

async function parseBody(request) {
  const contentType = request.headers.get("content-type") || "";

  if (contentType.includes("application/json")) {
    const body = await request.json();
    return body && typeof body === "object" && !Array.isArray(body)
      ? body
      : {};
  }

  const text = await request.text();

  if (!text) return {};

  if (text.trim().startsWith("{")) {
    const body = JSON.parse(text);
    return body && typeof body === "object" && !Array.isArray(body) ? body : {};
  }

  const params = new URLSearchParams(text);
  return Object.fromEntries(params.entries());
}

export default {
  async fetch(request, env) {
    const url = new URL(request.url);

    if (request.method === "OPTIONS") {
      return new Response(null, {
        status: 204,
        headers: corsHeaders()
      });
    }

    if (request.method === "GET" && url.pathname === "/health") {
      try {
        const databaseCheck = await env.STUDY_DB
          .prepare("SELECT 1 AS ok")
          .first();

        return json({
          ok: true,
          service: "dissertation-study-api",
          database: databaseCheck?.ok === 1,
          studyAudioBinding: Boolean(env.STUDY_AUDIO),
          trainingMediaBinding: Boolean(env.TRAINING_MEDIA),
          authConfigured: Boolean(env.ACCESS_CODE_PEPPER)
        });
      } catch {
        return json(
          {
            ok: false,
            service: "dissertation-study-api",
            error: "Health check failed"
          },
          500
        );
      }
    }

    if (
      request.method === "POST" &&
      (url.pathname === "/v1/aqg/validate-code" ||
        url.pathname === "/v1/training/validate-code")
    ) {
      try {
        const body = await parseBody(request);
        const app = url.pathname.includes("/training/")
          ? "training"
          : "aqg";
        const result = await validateAccessCode(env, body, app);
        return json(result.body, result.status);
      } catch {
        return json(
          {
            ok: false,
            valid: false,
            error: "Access validation failed"
          },
          500
        );
      }
    }

    if (
      request.method === "POST" &&
      (url.pathname === "/v1/aqg/create-session" ||
        url.pathname === "/v1/training/create-session")
    ) {
      try {
        const body = await parseBody(request);
        const app = url.pathname.includes("/training/") ? "training" : "aqg";
        const result = await createSessionId(env, body, app);
        return json(result.body, result.status);
      } catch {
        return json({ ok: false, error: "Session allocation failed", code: "SERVER_ERROR", retryable: true }, 500);
      }
    }

    if (
      request.method === "POST" &&
      url.pathname === "/v1/training/latest-live-session"
    ) {
      try {
        const body = await parseBody(request);
        const result = await getLatestTrainingLiveSession(env, body);
        return json(result.body, result.status);
      } catch {
        return json(
          {
            ok: false,
            error: "Training session recovery failed",
            code: "SERVER_ERROR",
            retryable: true
          },
          500
        );
      }
    }


    if (
      request.method === "POST" &&
      url.pathname === "/v1/training/save-session"
    ) {
      try {
        const body = await parseBody(request);
        const result = await saveTrainingLiveSession(env, body);
        return json(result.body, result.status);
      } catch {
        return json(
          {
            ok: false,
            error: "Training session save failed",
            code: "SERVER_ERROR",
            retryable: true
          },
          500
        );
      }
    }


    if (
      request.method === "POST" &&
      url.pathname === "/v1/aqg/create-context"
    ) {
      try {
        const body = await parseBody(request);
        const result = await createContextId(env, body);
        return json(result.body, result.status);
      } catch {
        return json(
          {
            ok: false,
            error: "Context allocation failed",
            code: "SERVER_ERROR",
            retryable: true
          },
          500
        );
      }
    }


    if (
      request.method === "POST" &&
      url.pathname === "/v1/aqg/save-session"
    ) {
      try {
        const body = await parseBody(request);
        const result = await saveAqgLiveSession(env, body);
        return json(result.body, result.status);
      } catch {
        return json(
          {
            ok: false,
            error: "Session save failed",
            code: "SERVER_ERROR",
            retryable: true
          },
          500
        );
      }
    }

    if (
      request.method === "POST" &&
      url.pathname === "/v1/aqg/submit-session"
    ) {
      try {
        const body = await parseBody(request);
        const result = await submitAqgSession(env, body);
        return json(result.body, result.status);
      } catch {
        return json(
          {
            ok: false,
            error: "Session submission failed",
            code: "SERVER_ERROR",
            retryable: true
          },
          500
        );
      }
    }


    if (
      request.method === "POST" &&
      url.pathname === "/v1/aqg/upload-audio"
    ) {
      try {
        const body = await parseBody(request);
        const result = await uploadAqgAudio(env, body);
        return json(result.body, result.status);
      } catch {
        return json(
          {
            ok: false,
            error: "Audio upload failed",
            code: "SERVER_ERROR",
            retryable: true
          },
          500
        );
      }
    }


    if (
      request.method === "POST" &&
      url.pathname === "/v1/aqg/feedback"
    ) {
      try {
        const body = await parseBody(request);
        const result = await saveAqgFeedback(env, body);
        return json(result.body, result.status);
      } catch {
        return json(
          {
            ok: false,
            error: "Feedback save failed",
            code: "SERVER_ERROR",
            retryable: true
          },
          500
        );
      }
    }


    if (
      request.method === "POST" &&
      url.pathname === "/v1/aqg/latest-submitted-settings"
    ) {
      try {
        const body = await parseBody(request);
        const result = await getLatestAqgSubmittedSettings(env, body);
        return json(result.body, result.status);
      } catch {
        return json(
          {
            ok: false,
            error: "Submitted settings lookup failed",
            code: "SERVER_ERROR",
            retryable: true
          },
          500
        );
      }
    }


    if (
      request.method === "POST" &&
      url.pathname === "/v1/aqg/latest-live-session"
    ) {
      try {
        const body = await parseBody(request);
        const result = await getLatestAqgLiveSession(env, body);
        return json(result.body, result.status);
      } catch {
        return json(
          {
            ok: false,
            error: "Live session recovery failed",
            code: "SERVER_ERROR",
            retryable: true
          },
          500
        );
      }
    }

    return json(
      {
        ok: false,
        error: "Not found"
      },
      404
    );
  }
};
