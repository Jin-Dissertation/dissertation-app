import { validateAccessCode } from "./auth.js";
import { createSessionId, createContextId } from "./sessions.js";
import { saveAqgLiveSession, getLatestAqgLiveSession, submitAqgSession, getLatestAqgSubmittedSettings } from "./aqg-save.js";
import { saveAqgFeedback, saveTrainingFeedback } from "./feedback.js";
import { uploadAqgAudio } from "./audio.js";
import { saveTrainingLiveSession, getLatestTrainingLiveSession, submitTrainingSession, appendTrainingEvent } from "./training-save.js";
import { loadTrainingContent, getTrainingMedia } from "./training-content.js";
import { deliverNotification, flushPendingNotifications } from "./notifications.js";
import { incrementNonparticipantButtonCount } from "./nonparticipant.js";
import { handleReporting } from "./reporting.js";

const ALLOWED_BROWSER_ORIGINS = new Set([
  "https://jin-dissertation.github.io",
  "http://127.0.0.1:8000",
  "http://localhost:8000"
]);

function browserOriginAllowed(request) {
  const origin = request?.headers?.get("origin") || "";
  return !origin || ALLOWED_BROWSER_ORIGINS.has(origin);
}

function corsHeaders(request) {
  const origin = request?.headers?.get("origin") || "";
  const headers = {
    "access-control-allow-methods": "GET, POST, OPTIONS",
    "access-control-allow-headers": "content-type",
    "vary": "Origin"
  };

  if (origin && ALLOWED_BROWSER_ORIGINS.has(origin)) {
    headers["access-control-allow-origin"] = origin;
  }

  return headers;
}

function json(data, status = 200, request = null) {
  return new Response(JSON.stringify(data, null, 2), {
    status,
    headers: {
      "content-type": "application/json; charset=utf-8",
      ...corsHeaders(request)
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
  async fetch(request, env, ctx) {
    const url = new URL(request.url);
    const respond = (data, status = 200) => json(data, status, request);

    if (url.pathname === "/v1/reporting" || url.pathname.startsWith("/v1/reporting/")) {
      return handleReporting(request, env);
    }

    if (!browserOriginAllowed(request)) {
      return respond(
        {
          ok: false,
          error: "Browser origin is not allowed."
        },
        403
      );
    }

    if (request.method === "OPTIONS") {
      return new Response(null, {
        status: 204,
        headers: corsHeaders(request)
      });
    }

    if (request.method === "GET" && url.pathname === "/health") {
      try {
        const databaseCheck = await env.STUDY_DB
          .prepare("SELECT 1 AS ok")
          .first();

        return respond({
          ok: true,
          service: "dissertation-study-api",
          database: databaseCheck?.ok === 1,
          studyAudioBinding: Boolean(env.STUDY_AUDIO),
          trainingMediaBinding: Boolean(env.TRAINING_MEDIA),
          authConfigured: Boolean(env.ACCESS_CODE_PEPPER)
        });
      } catch {
        return respond(
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
      url.pathname === "/v1/nonparticipant/button-press"
    ) {
      try {
        const body = await parseBody(request);
        const result = await incrementNonparticipantButtonCount(env, body);
        return respond(result.body, result.status);
      } catch {
        return respond(
          {
            ok: false,
            error: "Non-participant button count update failed",
            code: "SERVER_ERROR",
            retryable: true
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
        return respond(result.body, result.status);
      } catch {
        return respond(
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
        return respond(result.body, result.status);
      } catch {
        return respond({ ok: false, error: "Session allocation failed", code: "SERVER_ERROR", retryable: true }, 500);
      }
    }

    if (
      request.method === "GET" &&
      url.pathname.startsWith("/v1/training/media/")
    ) {
      const parts = url.pathname.split("/").filter(Boolean);
      const version = parts[3] || "";
      const filename = parts[4] || "";

      if (parts.length !== 5) {
        return new Response("Not found", { status: 404 });
      }

      return getTrainingMedia(env, version, filename);
    }


    if (
      request.method === "POST" &&
      url.pathname === "/v1/training/load-content"
    ) {
      try {
        const body = await parseBody(request);
        const result = await loadTrainingContent(env, body, request.url);
        return respond(result.body, result.status);
      } catch {
        return respond(
          {
            ok: false,
            error: "Training content load failed",
            code: "SERVER_ERROR",
            retryable: true
          },
          500
        );
      }
    }


    if (
      request.method === "POST" &&
      url.pathname === "/v1/training/submit-session"
    ) {
      try {
        const body = await parseBody(request);
        const result = await submitTrainingSession(env, body);
        return respond(result.body, result.status);
      } catch {
        return respond(
          {
            ok: false,
            error: "Training session submission failed",
            code: "SERVER_ERROR",
            retryable: true
          },
          500
        );
      }
    }


    if (
      request.method === "POST" &&
      url.pathname === "/v1/training/latest-live-session"
    ) {
      try {
        const body = await parseBody(request);
        const result = await getLatestTrainingLiveSession(env, body);
        return respond(result.body, result.status);
      } catch {
        return respond(
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
        return respond(result.body, result.status);
      } catch {
        return respond(
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
        return respond(result.body, result.status);
      } catch {
        return respond(
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
        return respond(result.body, result.status);
      } catch {
        return respond(
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
        return respond(result.body, result.status);
      } catch {
        return respond(
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
        return respond(result.body, result.status);
      } catch {
        return respond(
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
      url.pathname === "/v1/training/append-event"
    ) {
      try {
        const body = await parseBody(request);
        const result = await appendTrainingEvent(env, body);

        if (
          result.status === 200 &&
          result.body?.completion_notification_id &&
          ctx?.waitUntil
        ) {
          ctx.waitUntil(
            deliverNotification(
              env,
              result.body.completion_notification_id
            )
          );
        }

        return respond(result.body, result.status);
      } catch {
        return respond(
          {
            ok: false,
            error: "Training event append failed",
            code: "SERVER_ERROR",
            retryable: true
          },
          500
        );
      }
    }

    if (
      request.method === "POST" &&
      url.pathname === "/v1/training/feedback"
    ) {
      try {
        const body = await parseBody(request);
        const result = await saveTrainingFeedback(env, body);

        if (
          result.status === 200 &&
          result.body?.notification_id &&
          ctx?.waitUntil
        ) {
          ctx.waitUntil(
            deliverNotification(env, result.body.notification_id)
          );
        }

        return respond(result.body, result.status);
      } catch {
        return respond(
          {
            ok: false,
            error: "Training feedback save failed",
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

        if (
          result.status === 200 &&
          result.body?.notification_id &&
          ctx?.waitUntil
        ) {
          ctx.waitUntil(
            deliverNotification(env, result.body.notification_id)
          );
        }

        return respond(result.body, result.status);
      } catch {
        return respond(
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
        return respond(result.body, result.status);
      } catch {
        return respond(
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
        return respond(result.body, result.status);
      } catch {
        return respond(
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

    return respond(
      {
        ok: false,
        error: "Not found"
      },
      404
    );
  },

  async scheduled(_event, env, ctx) {
    ctx.waitUntil(flushPendingNotifications(env, 10));
  }
};
