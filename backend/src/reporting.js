/*
 * MAINTAINER GUIDE — READ-ONLY REPORTING FEED
 *
 * This protected server-to-server endpoint is the Cloudflare side of the
 * UA OneDrive/Excel archive workflow.
 *
 * REPORTING_EXPORT_TOKEN authenticates the reporting client. The endpoint
 * reads the mirror feed in bounded pages and never accepts participant-browser
 * authentication. Structured detail JSON is scrubbed for credential-like keys.
 *
 * Keep this endpoint read-only. Deletion belongs only in the guarded purge
 * tools after a verified UA archive receipt exists.
 */

import { REPORTING_DATASETS, REPORTING_PROTOCOL_VERSION } from "./reporting-contract.js";

const encoder = new TextEncoder();
const PAGE_BYTES = 1024 * 1024;
const MAX_RESPONSE_BYTES = 4 * 1024 * 1024;
const datasets = new Map(REPORTING_DATASETS.map(dataset => [dataset.name, dataset]));
const credentialKeys = new Set([
  "code", "accesscode", "participantcode", "codeentered", "userid", "clientkey",
  "token", "accesstoken", "refreshtoken", "authorization", "password", "secret",
  "codehash", "accesscodepepper", "reportingexporttoken", "relaysecret"
]);

function reply(body, status = 200) {
  return new Response(JSON.stringify(body), {
    status,
    headers: {
      "content-type": "application/json; charset=utf-8",
      "cache-control": "no-store, private",
      "x-content-type-options": "nosniff",
      "referrer-policy": "no-referrer"
    }
  });
}

function error(code, status) {
  return reply({ ok: false, code }, status);
}

async function sameSecret(a, b) {
  const [left, right] = await Promise.all([a, b].map(value =>
    crypto.subtle.digest("SHA-256", encoder.encode(value))));
  const x = new Uint8Array(left);
  const y = new Uint8Array(right);
  let difference = 0;
  for (let i = 0; i < x.length; i++) difference |= x[i] ^ y[i];
  return difference === 0;
}

// Only used for structured event details, not participant-authored research text.
// Redact credential aliases at every object depth, including JSON nested as strings.
function cleanDetails(value, depth = 0) {
  if (depth > 40) return null;
  if (Array.isArray(value)) return value.map(item => cleanDetails(item, depth + 1));
  if (value && typeof value === "object") {
    return Object.fromEntries(Object.entries(value)
      .filter(([key]) => !credentialKeys.has(key.toLowerCase().replace(/[^a-z0-9]/g, "")))
      .map(([key, item]) => [key, cleanDetails(item, depth + 1)]));
  }
  if (typeof value === "string" && /^[\s]*[\[{]/.test(value)) {
    try { return JSON.stringify(cleanDetails(JSON.parse(value), depth + 1)); }
    catch { return null; }
  }
  return value;
}

function exportRecord(change) {
  const dataset = datasets.get(change.dataset);
  if (!dataset || !["upsert", "delete"].includes(change.operation)) throw new Error("Invalid feed record");
  if (change.operation === "delete") return null;
  const source = JSON.parse(change.record_json);
  if (!source || source[dataset.key] !== change.record_id) throw new Error("Invalid feed key");
  const record = Object.fromEntries(dataset.columns.map(column => [column, source[column] ?? null]));
  if (record.detail_json !== undefined && record.detail_json !== null) {
    try { record.detail_json = JSON.stringify(cleanDetails(JSON.parse(record.detail_json))); }
    catch { record.detail_json = null; }
  }
  return record;
}

function integerParam(params, name, fallback, maximum = Number.MAX_SAFE_INTEGER) {
  if (!params.has(name)) return fallback;
  const value = params.get(name);
  if (!/^(0|[1-9]\d*)$/.test(value)) throw new Error("Invalid integer");
  const number = Number(value);
  if (!Number.isSafeInteger(number) || number > maximum) throw new Error("Invalid integer");
  return number;
}

const metadataSql = `SELECT feed_generation, created_at,
  COALESCE((SELECT MAX(sequence) FROM mirror_changes c
    WHERE c.feed_generation = s.feed_generation), 0) AS latest_sequence
  FROM mirror_feed_state s WHERE singleton = 1`;

export async function handleReporting(request, env) {
  // Separate server-to-server boundary, before participant CORS or authentication.
  // No token in query strings, browser storage, workbook, or participant access codes.
  const secret = env.REPORTING_EXPORT_TOKEN;
  if (typeof secret !== "string" || secret.length < 32) return error("REPORTING_DISABLED", 503);
  const authorization = request.headers.get("authorization") || "";
  const match = /^Bearer ([^\s]{32,512})$/i.exec(authorization);
  if (!match || !(await sameSecret(match[1], secret))) return error("UNAUTHORIZED", 401);
  if (request.headers.has("origin")) return error("SERVER_TO_SERVER_ONLY", 403);
  if (request.method !== "GET") return error("METHOD_NOT_ALLOWED", 405);

  const url = new URL(request.url);
  const manifest = url.pathname === "/v1/reporting/manifest";
  if (!manifest && url.pathname !== "/v1/reporting/changes") return error("NOT_FOUND", 404);
  const allowed = manifest ? [] : ["generation", "after", "through", "limit"];
  for (const key of url.searchParams.keys()) {
    if (!allowed.includes(key) || url.searchParams.getAll(key).length !== 1) return error("INVALID_QUERY", 400);
  }

  let after, through, limit, generation;
  if (!manifest) {
    try {
      generation = url.searchParams.get("generation");
      if (!/^[a-f0-9]{32}$/.test(generation || "")) throw new Error("Invalid generation");
      after = integerParam(url.searchParams, "after", 0);
      through = integerParam(url.searchParams, "through", null);
      limit = integerParam(url.searchParams, "limit", 25, 100);
      if (limit < 1 || (through !== null && through < after)) throw new Error("Invalid range");
    } catch { return error("INVALID_QUERY", 400); }
  }

  try {
    // Pin reads to the primary if read replication is enabled in the future.
    const db = env.STUDY_DB.withSession("first-primary");
    if (manifest) {
      const meta = await db.prepare(metadataSql).first();
      if (!meta) return error("FEED_UNAVAILABLE", 503);
      return reply({ ok: true, protocol_version: REPORTING_PROTOCOL_VERSION, ...meta, datasets: REPORTING_DATASETS });
    }

    // Snapshot metadata and page in one read transaction. Limit bytes in SQL so a
    // page of large submissions cannot exhaust Worker memory before pagination.
    const [metaResult, pageResult] = await db.batch([
      db.prepare(metadataSql),
      db.prepare(`WITH candidates AS (
        SELECT sequence, COALESCE(length(CAST(record_json AS BLOB)), 4)
          + length(CAST(record_id AS BLOB)) * 6 + 1024 AS bytes
        FROM mirror_changes
        WHERE feed_generation = ?1 AND sequence > ?2
          AND sequence <= COALESCE(?3, (SELECT MAX(sequence) FROM mirror_changes WHERE feed_generation = ?1))
        ORDER BY sequence LIMIT ?4
      ), sized AS (
        SELECT sequence, SUM(bytes) OVER (ORDER BY sequence) AS page_bytes,
          ROW_NUMBER() OVER (ORDER BY sequence) AS row_number FROM candidates
      ) SELECT c.sequence, c.dataset, c.record_id, c.revision, c.operation, c.changed_at, c.record_json
        FROM sized s JOIN mirror_changes c ON c.sequence = s.sequence
        WHERE s.page_bytes <= ?5 OR s.row_number = 1 ORDER BY c.sequence`)
        .bind(generation, after, through, limit, PAGE_BYTES)
    ]);
    const meta = metaResult.results[0];
    if (!meta) return error("FEED_UNAVAILABLE", 503);
    if (meta.feed_generation !== generation) return error("FEED_GENERATION_CHANGED", 409);
    const upper = through ?? meta.latest_sequence;
    if (after > meta.latest_sequence || upper > meta.latest_sequence) return error("CURSOR_AHEAD", 409);
    const changes = pageResult.results.map(({ record_json, ...change }) => ({
      ...change, record: exportRecord({ ...change, record_json })
    }));
    const next = changes.length ? changes[changes.length - 1].sequence : upper;
    const body = {
      ok: true, protocol_version: REPORTING_PROTOCOL_VERSION,
      feed_generation: generation, after, through: upper, next_after: next,
      has_more: next < upper, changes
    };
    if (encoder.encode(JSON.stringify(body)).byteLength > MAX_RESPONSE_BYTES) return error("EXPORT_RECORD_TOO_LARGE", 413);
    return reply(body);
  } catch {
    // Never log or return record contents, tokens, SQL, or binding errors.
    return error("FEED_UNAVAILABLE", 503);
  }
}
