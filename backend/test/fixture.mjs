import { Miniflare, Log, LogLevel, convertV4MiniflareOptions } from "miniflare";
import { build } from "esbuild";
import { readFile, mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { randomBytes, createHmac } from "node:crypto";
import { migrationStatements } from "../scripts/reporting-schema.mjs";

export async function fixture({ beforeMirror, tokenConfigured = true, provisioningTokenConfigured = false } = {}) {
  const directory = await mkdtemp(join(tmpdir(), "dissertation-synthetic-"));
  const token = randomBytes(32).toString("hex");
  const pepper = randomBytes(32).toString("hex");
  const provisioningToken = randomBytes(32).toString("hex");
  const code = "synthetic-" + randomBytes(12).toString("hex");
  const bundle = await build({ entryPoints: [new URL("../src/worker.js", import.meta.url).pathname], bundle: true, format: "esm", write: false, platform: "browser" });
  const mf = new Miniflare(convertV4MiniflareOptions({
    modules: true, script: bundle.outputFiles[0].text,
    compatibilityDate: "2026-09-26",
    d1Databases: ["STUDY_DB"], d1Persist: join(directory, "d1"),
    r2Buckets: ["STUDY_AUDIO", "TRAINING_MEDIA"], r2Persist: join(directory, "r2"),
    bindings: {
      ACCESS_CODE_PEPPER: pepper,
      ...(tokenConfigured ? { REPORTING_EXPORT_TOKEN: token } : {}),
      ...(provisioningTokenConfigured ? { PARTICIPANT_PROVISIONING_TOKEN: provisioningToken } : {})
    },
    // Any attempted external request fails; notification relay credentials absent.
    outboundService: () => new Response("Synthetic tests forbid external requests", { status: 502 }),
    log: new Log(LogLevel.ERROR)
  }));
  try {
    const db = await mf.getD1Database("STUDY_DB");
    for (const name of ["0001_operational_foundation.sql", "0002_study_schema.sql"]) {
      const sql = await readFile(new URL("../migrations/" + name, import.meta.url), "utf8");
      await db.batch(sql.split(";").map(s => s.trim()).filter(Boolean).map(s => db.prepare(s)));
    }
    if (beforeMirror) await beforeMirror(db);
    await db.batch(migrationStatements().map(s => db.prepare(s)));
    await db.prepare(`INSERT INTO access_codes (participant_id, code_hash, created_at, updated_at)
      VALUES ('TEST001', ?1, '2026-09-29T00:00:00Z', '2026-09-29T00:00:00Z')`)
      .bind(createHmac("sha256", pepper).update("access:" + code).digest("hex")).run();
    const request = async (path, { method = "GET", body, headers = {} } = {}) => {
      const response = await mf.dispatchFetch("https://synthetic.invalid" + path, {
        method, headers: { ...(body ? { "content-type": "application/json" } : {}), ...headers },
        ...(body ? { body: JSON.stringify(body) } : {})
      });
      return { status: response.status, headers: response.headers, body: await response.json() };
    };
    return {
      db, mf, code, token, provisioningToken, request,
      post: (path, body) => request(path, { method: "POST", body: { code, ...body } }),
      report: (path, options = {}) => request("/v1/reporting/" + path, { ...options, headers: { authorization: "Bearer " + token, ...options.headers } }),
      count: async () => (await db.prepare("SELECT COUNT(*) AS n FROM mirror_changes").first()).n,
      close: async () => { await mf.dispose(); await rm(directory, { recursive: true, force: true }); }
    };
  } catch (error) {
    await mf.dispose();
    await rm(directory, { recursive: true, force: true });
    throw error;
  }
}
