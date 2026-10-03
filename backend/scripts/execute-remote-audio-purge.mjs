/*
 * OPERATOR GUIDE — GUARDED R2 AUDIO CLEANUP
 *
 * Preview mode revalidates the verified audio receipt, the current R2 bytes,
 * and current D1 references. Destructive mode requires explicit confirmation.
 *
 * A reporting receipt alone is never enough to delete audio. Audio has its own
 * independent archive/verification/receipt chain.
 */

import {
  mkdtemp,
  readFile,
  rm,
  stat
} from "node:fs/promises";
import { createHash } from "node:crypto";
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import { tmpdir } from "node:os";
import { join, resolve } from "node:path";

import {
  validateAudioArchiveManifest,
  validateAudioArchiveReceipt
} from "../src/audio-archive.js";
import { buildAudioPurgePlan } from "../src/audio-purge.js";

const execFileAsync = promisify(execFile);
const WRANGLER_VERSION = "4.145.0";
const DEFAULT_BUCKET = "dissertation-study-audio";
const DEFAULT_DATABASE = "dissertation-study-data";

function argument(name) {
  const index = process.argv.indexOf(name);
  if (index === -1) return null;
  const value = process.argv[index + 1];
  if (!value || value.startsWith("--")) {
    throw new Error(`${name} requires a value.`);
  }
  return value;
}

function flag(name) {
  return process.argv.includes(name);
}

function sqlLiteral(value) {
  return "'" + String(value).replace(/'/g, "''") + "'";
}

function parseWranglerJsonOutput(stdout) {
  const text = String(stdout ?? "");
  const lines = text.split(/\r?\n/);

  for (let i = 0; i < lines.length; i += 1) {
    const trimmed = lines[i].trimStart();
    if (!trimmed.startsWith("[") && !trimmed.startsWith("{")) continue;

    const candidate = lines.slice(i).join("\n").trim();
    try {
      return JSON.parse(candidate);
    } catch {
      // Wrangler may emit progress text before the JSON payload.
    }
  }

  throw new Error("Wrangler did not return a parseable JSON payload.");
}

async function sha256File(path) {
  const data = await readFile(path);
  return createHash("sha256").update(data).digest("hex");
}

async function readJson(path, label) {
  if (!path) throw new Error(`${label} path is required.`);
  return JSON.parse(await readFile(resolve(path), "utf8"));
}

async function readCurrentR2Object({
  bucket,
  object,
  stagingDirectory
}) {
  const path = join(
    stagingDirectory,
    `${String(object.object_key)
      .replace(/[^A-Za-z0-9_.-]/g, "_")}.bin`
  );

  try {
    await execFileAsync(
      "npx",
      [
        "--yes",
        `wrangler@${WRANGLER_VERSION}`,
        "r2",
        "object",
        "get",
        `${bucket}/${object.object_key}`,
        "--remote",
        "--file",
        path
      ],
      { maxBuffer: 16 * 1024 * 1024 }
    );
  } catch (error) {
    throw new Error(
      `Could not read remote R2 object ${object.object_key}: ` +
      `${String(error?.stderr || error?.message || error).trim()}`
    );
  }

  const info = await stat(path);

  return {
    exists: true,
    sha256: await sha256File(path),
    byte_length: info.size
  };
}

async function readOperationalReferences(database, objectKey) {
  const key = sqlLiteral(objectKey);
  const sql =
    "SELECT " +
    `(SELECT COUNT(*) FROM aqg_live_sessions WHERE audio_object_key = ${key}) AS live_sessions, ` +
    `(SELECT COUNT(*) FROM aqg_submissions WHERE audio_object_key = ${key}) AS submissions, ` +
    `(SELECT COUNT(*) FROM aqg_feedback WHERE audio_object_key = ${key}) AS feedback, ` +
    "(SELECT COUNT(*) FROM notification_outbox " +
    "WHERE status <> 'sent' " +
    `AND json_extract(payload_json, '$.audio_object_key') = ${key}) AS pending_notifications`;

  const { stdout } = await execFileAsync(
    "npx",
    [
      "--yes",
      `wrangler@${WRANGLER_VERSION}`,
      "d1",
      "execute",
      database,
      "--remote",
      "--json",
      "--command",
      sql
    ],
    { maxBuffer: 16 * 1024 * 1024 }
  );

  const parsed = parseWranglerJsonOutput(stdout);
  if (
    !Array.isArray(parsed) ||
    parsed.length !== 1 ||
    parsed[0]?.success !== true ||
    !Array.isArray(parsed[0]?.results) ||
    parsed[0]?.meta?.changed_db !== false
  ) {
    throw new Error(
      "Operational-reference check did not prove a successful read-only D1 query."
    );
  }

  const row = parsed[0].results[0] || {};

  return {
    live_sessions: Number(row.live_sessions || 0),
    submissions: Number(row.submissions || 0),
    feedback: Number(row.feedback || 0),
    pending_notifications: Number(row.pending_notifications || 0)
  };
}

async function buildRemotePlan({
  manifest,
  receipt,
  bucket,
  database
}) {
  validateAudioArchiveManifest(manifest);
  validateAudioArchiveReceipt({ manifest, receipt });

  const stagingDirectory = await mkdtemp(
    join(tmpdir(), "dissertation-audio-purge-plan-")
  );

  try {
    const objectStates = {};
    const operationalReferences = {};

    for (let index = 0; index < manifest.objects.length; index += 1) {
      const object = manifest.objects[index];

      process.stdout.write(
        `Revalidating R2 object ${index + 1}/${manifest.objects.length}...\n`
      );

      objectStates[object.object_key] = await readCurrentR2Object({
        bucket,
        object,
        stagingDirectory
      });

      operationalReferences[object.object_key] =
        await readOperationalReferences(database, object.object_key);
    }

    return buildAudioPurgePlan({
      manifest,
      receipt,
      objectStates,
      operationalReferences
    });
  } finally {
    await rm(stagingDirectory, { recursive: true, force: true });
  }
}

function printPlan(plan, executeRequested = false) {
  const lines = [
    "",
    "GUARDED AUDIO PURGE",
    "-------------------",
    `Audio export ID: ${plan.audio_export_id}`,
    `Verified receipt objects: ${plan.verified_object_count}`,
    `Would delete: ${plan.eligible_count}`,
    `Blocked: ${plan.blocked_count}`
  ];

  for (const blocked of plan.blocked_objects) {
    lines.push(
      `KEEP ${blocked.object_key}: ${blocked.keep_reasons.join(", ")}`
    );
  }

  if (!executeRequested) {
    lines.push(
      "",
      "PREVIEW ONLY.",
      "R2 objects were downloaded only for hash/size revalidation.",
      "D1 checks were SELECT-only.",
      "No R2 object was deleted.",
      ""
    );
  }

  process.stdout.write(lines.join("\n"));
}

async function main() {
  const manifestPath = argument("--manifest");
  const receiptPath = argument("--receipt");
  const execute = flag("--execute");
  const confirmAudioExportId = argument("--confirm-audio-export-id");
  const confirmDeleteCount = argument("--confirm-delete-count");

  const bucket = argument("--bucket") || DEFAULT_BUCKET;
  const database = argument("--database") || DEFAULT_DATABASE;

  const manifest = await readJson(manifestPath, "Manifest");
  const receipt = await readJson(receiptPath, "Receipt");

  const plan = await buildRemotePlan({
    manifest,
    receipt,
    bucket,
    database
  });

  printPlan(plan, execute);

  if (!execute) return;

  if (plan.blocked_count !== 0) {
    throw new Error(
      "Execution refused because at least one verified audio object is still blocked."
    );
  }

  if (confirmAudioExportId !== plan.audio_export_id) {
    throw new Error(
      "--confirm-audio-export-id must exactly match the verified audio export ID."
    );
  }

  if (Number(confirmDeleteCount) !== plan.eligible_count) {
    throw new Error(
      `--confirm-delete-count must exactly equal ${plan.eligible_count}.`
    );
  }

  // Rebuild the complete plan immediately before any destructive request.
  const revalidated = await buildRemotePlan({
    manifest,
    receipt,
    bucket,
    database
  });

  if (
    revalidated.blocked_count !== 0 ||
    revalidated.eligible_count !== plan.eligible_count
  ) {
    throw new Error(
      "Audio cleanup state changed during revalidation; no delete was attempted."
    );
  }

  process.stdout.write(
    [
      "",
      "EXECUTION AUTHORIZED.",
      "Receipt, current R2 bytes, and D1 operational references were revalidated.",
      `Deleting exactly ${revalidated.eligible_count} verified R2 object(s).`,
      ""
    ].join("\n")
  );

  let deleted = 0;

  for (const objectKey of revalidated.eligible_objects) {
    await execFileAsync(
      "npx",
      [
        "--yes",
        `wrangler@${WRANGLER_VERSION}`,
        "r2",
        "object",
        "delete",
        `${bucket}/${objectKey}`,
        "--remote"
      ],
      { maxBuffer: 16 * 1024 * 1024 }
    );
    deleted += 1;
    process.stdout.write(`Deleted ${objectKey}\n`);
  }

  process.stdout.write(
    [
      "",
      "AUDIO PURGE COMPLETE",
      "--------------------",
      `Deleted: ${deleted}`,
      "The verified OneDrive archive remains the durable audio copy.",
      ""
    ].join("\n")
  );
}

main().catch(error => {
  process.stderr.write(`GUARDED AUDIO PURGE FAILED: ${error.message}\n`);
  process.exitCode = 1;
});
