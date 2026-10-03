/*
 * MAINTAINER GUIDE — PARTICIPANT-CODE ADMINISTRATION
 *
 * This server-to-server administrative endpoint adds/reactivates or deactivates
 * participant codes. It is intentionally DISABLED unless the temporary
 * PARTICIPANT_PROVISIONING_TOKEN Worker secret exists.
 *
 * Current study workflow:
 *   participant code entered by participant
 *      → normalized here
 *      → same value used as deidentified participant_id
 *      → only HMAC(code) stored for authentication lookup
 *
 * Normal operation should leave PARTICIPANT_PROVISIONING_TOKEN deleted.
 * Responses deliberately never echo participant codes. Deactivation blocks
 * future access but does not delete prior research records.
 */

import { hashAccessCode, normalizeAccessCode } from "./auth.js";

const encoder = new TextEncoder();
const MAX_ENTRIES = 50;

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
  const [left, right] = await Promise.all(
    [a, b].map((value) =>
      crypto.subtle.digest("SHA-256", encoder.encode(value))
    )
  );
  const x = new Uint8Array(left);
  const y = new Uint8Array(right);
  let difference = 0;
  for (let i = 0; i < x.length; i += 1) difference |= x[i] ^ y[i];
  return difference === 0;
}

function normalizeParticipantId(value) {
  return String(value ?? "").trim();
}

function booleanFlag(value, fallback = 1) {
  if (value === undefined || value === null) return fallback;
  if (value === true || value === 1 || value === "1") return 1;
  if (value === false || value === 0 || value === "0") return 0;
  throw new Error("Invalid boolean flag");
}

export async function handleProvisioning(request, env) {
  const secret = env.PARTICIPANT_PROVISIONING_TOKEN;
  if (typeof secret !== "string" || secret.length < 32) {
    return error("PROVISIONING_DISABLED", 503);
  }

  const authorization = request.headers.get("authorization") || "";
  const match = /^Bearer ([^\\s]{32,512})$/i.exec(authorization);
  if (!match || !(await sameSecret(match[1], secret))) {
    return error("UNAUTHORIZED", 401);
  }

  if (request.headers.has("origin")) {
    return error("SERVER_TO_SERVER_ONLY", 403);
  }

  if (request.method !== "POST") {
    return error("METHOD_NOT_ALLOWED", 405);
  }

  let body;
  try {
    body = await request.json();
  } catch {
    return error("INVALID_JSON", 400);
  }

  if (!body || typeof body !== "object" || Array.isArray(body)) {
    return error("INVALID_REQUEST", 400);
  }

  const entries = body.entries;
  if (!Array.isArray(entries) || entries.length < 1 || entries.length > MAX_ENTRIES) {
    return error("INVALID_ENTRIES", 400);
  }

  const now = new Date().toISOString();
  const seenParticipants = new Set();
  const seenHashes = new Set();
  const prepared = [];

  try {
    for (const entry of entries) {
      if (!entry || typeof entry !== "object" || Array.isArray(entry)) {
        throw new Error("Invalid entry");
      }

      const participantCode = normalizeAccessCode(entry.participant_code);

      if (
        participantCode.length < 6 ||
        participantCode.length > 64 ||
        !/^[a-z0-9][a-z0-9._-]*$/.test(participantCode)
      ) {
        throw new Error("Invalid participant code");
      }
      if (seenParticipants.has(participantCode)) {
        throw new Error("Duplicate participant code");
      }

      const codeHash = await hashAccessCode(env, participantCode);
      if (seenHashes.has(codeHash)) {
        throw new Error("Duplicate participant code");
      }

      seenParticipants.add(participantCode);
      seenHashes.add(codeHash);

      prepared.push({
        participantId: participantCode,
        codeHash,
        active: booleanFlag(entry.active, 1),
        allowAqg: booleanFlag(entry.allow_aqg, 1),
        allowTraining: booleanFlag(entry.allow_training, 1)
      });
    }
  } catch {
    return error("INVALID_ENTRY", 400);
  }

  try {
    const statements = prepared.map((entry) =>
      env.STUDY_DB
        .prepare(
          `INSERT INTO access_codes
             (participant_id, code_hash, active, allow_aqg, allow_training, created_at, updated_at)
           VALUES (?1, ?2, ?3, ?4, ?5, ?6, ?6)
           ON CONFLICT(participant_id) DO UPDATE SET
             code_hash = excluded.code_hash,
             active = excluded.active,
             allow_aqg = excluded.allow_aqg,
             allow_training = excluded.allow_training,
             updated_at = excluded.updated_at`
        )
        .bind(
          entry.participantId,
          entry.codeHash,
          entry.active,
          entry.allowAqg,
          entry.allowTraining,
          now
        )
    );

    await env.STUDY_DB.batch(statements);

    return reply({
      ok: true,
      provisioned: prepared.length
    });
  } catch {
    return error("PROVISIONING_FAILED", 409);
  }
}
