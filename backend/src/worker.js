function json(data, status = 200) {
  return new Response(JSON.stringify(data, null, 2), {
    status,
    headers: {
      "content-type": "application/json; charset=utf-8"
    }
  });
}

export default {
  async fetch(request, env) {
    const url = new URL(request.url);

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
          trainingMediaBinding: Boolean(env.TRAINING_MEDIA)
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

    return json(
      {
        ok: false,
        error: "Not found"
      },
      404
    );
  }
};
