import { validateAccessCode } from "./auth.js";

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

    return json(
      {
        ok: false,
        error: "Not found"
      },
      404
    );
  }
};
