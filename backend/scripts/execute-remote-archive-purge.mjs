import { readFile, writeFile, chmod, unlink } from "node:fs/promises";
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import { fileURLToPath } from "node:url";
import { resolve } from "node:path";
import { tmpdir } from "node:os";
import { randomUUID } from "node:crypto";

import { validateArchiveReceipt } from "../src/archive-receipt.js";
import { buildArchivePurgePlan } from "../src/archive-purge-plan.js";
import {
  buildRemoteArchiveSnapshotQuery,
  parseWranglerArchiveSnapshot,
  createReadOnlySnapshotDb
} from "../src/archive-purge-remote.js";
import { buildGuardedArchivePurgeSql } from "../src/archive-purge-execute.js";

const execFileAsync = promisify(execFile);

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

async function readJson(path, label) {
  if (!path) throw new Error(`${label} path is required.`);
  return JSON.parse(await readFile(resolve(path), "utf8"));
}

async function remotePlan({ archiveExport, receipt, backendDirectory }) {
  const verified = validateArchiveReceipt({ archiveExport, receipt });

  const query = buildRemoteArchiveSnapshotQuery({
    feedGeneration: verified.feed_generation,
    verifiedRecords: verified.verified_records
  });

  const { stdout } = await execFileAsync(
    "npx",
    [
      "--yes",
      "wrangler@4.145.0",
      "d1",
      "execute",
      "dissertation-study-data",
      "--remote",
      "--json",
      "--command",
      query.sql
    ],
    {
      cwd: backendDirectory,
      maxBuffer: 64 * 1024 * 1024
    }
  );

  const snapshot = parseWranglerArchiveSnapshot({
    query,
    wranglerResults: JSON.parse(stdout)
  });

  return buildArchivePurgePlan({
    db: createReadOnlySnapshotDb(snapshot),
    archiveExport,
    receipt
  });
}

async function main() {
  const archivePath = argument("--archive");
  const receiptPath = argument("--receipt");
  const execute = flag("--execute");
  const confirmExportId = argument("--confirm-export-id");
  const confirmDeleteCount = argument("--confirm-delete-count");

  const archiveExport = await readJson(archivePath, "Archive");
  const receipt = await readJson(receiptPath, "Receipt");

  const backendDirectory = fileURLToPath(
    new URL("..", import.meta.url)
  );

  const plan = await remotePlan({
    archiveExport,
    receipt,
    backendDirectory
  });

  const guarded = buildGuardedArchivePurgeSql({
    plan,
    archiveExport
  });

  const lines = [
    "",
    "GUARDED ARCHIVE PURGE",
    "---------------------",
    `Export ID: ${plan.export_id}`,
    `Current generation matches receipt: ${plan.feed_generation_matches}`,
    `Verified receipt records: ${plan.verified_record_count}`,
    `Would delete: ${guarded.delete_count}`,
    `Retained by policy: ${guarded.retained_count}`,
    `Retained dataset(s): ${guarded.retained_datasets.join(", ")}`
  ];

  if (!execute) {
    lines.push(
      "",
      "PREVIEW ONLY.",
      "No DELETE, UPDATE, INSERT, DROP, or CREATE statements were sent to D1.",
      "Run again with the explicit execution confirmations only after reviewing this preview.",
      ""
    );
    process.stdout.write(lines.join("\n"));
    return;
  }

  if (confirmExportId !== plan.export_id) {
    throw new Error(
      "--confirm-export-id must exactly match the verified export id."
    );
  }

  if (Number(confirmDeleteCount) !== guarded.delete_count) {
    throw new Error(
      `--confirm-delete-count must exactly equal ${guarded.delete_count}.`
    );
  }

  const sqlPath = resolve(
    tmpdir(),
    `dissertation-archive-purge-${randomUUID()}.sql`
  );

  await writeFile(sqlPath, guarded.sql, { mode: 0o600 });
  await chmod(sqlPath, 0o600);

  try {
    lines.push(
      "",
      "EXECUTION AUTHORIZED.",
      "Revalidation passed immediately before execution.",
      `Deleting exactly ${guarded.delete_count} exact archived records.`,
      "Cumulative nonparticipant counters are retained.",
      "The reporting feed will rotate and the old feed generation will be removed."
    );
    process.stdout.write(lines.join("\n") + "\n");

    await execFileAsync(
      "npx",
      [
        "--yes",
        "wrangler@4.145.0",
        "d1",
        "execute",
        "dissertation-study-data",
        "--remote",
        "--yes",
        "--file",
        sqlPath
      ],
      {
        cwd: backendDirectory,
        maxBuffer: 64 * 1024 * 1024
      }
    );
  } finally {
    await unlink(sqlPath).catch(() => {});
  }

  const escapedOldGeneration = guarded.old_feed_generation.replace(/'/g, "''");
  const verifySql =
    "SELECT feed_generation FROM mirror_feed_state WHERE singleton = 1; " +
    "SELECT COUNT(*) AS old_generation_rows FROM mirror_changes " +
    `WHERE feed_generation = '${escapedOldGeneration}';`;

  const { stdout: verificationStdout } = await execFileAsync(
    "npx",
    [
      "--yes",
      "wrangler@4.145.0",
      "d1",
      "execute",
      "dissertation-study-data",
      "--remote",
      "--json",
      "--command",
      verifySql
    ],
    {
      cwd: backendDirectory,
      maxBuffer: 64 * 1024 * 1024
    }
  );

  const verification = JSON.parse(verificationStdout);
  const newGeneration =
    String(verification?.[0]?.results?.[0]?.feed_generation || "");
  const oldRows =
    Number(verification?.[1]?.results?.[0]?.old_generation_rows ?? -1);

  if (!/^[a-f0-9]{32}$/.test(newGeneration)) {
    throw new Error("Post-purge verification returned an invalid new generation.");
  }

  if (newGeneration === guarded.old_feed_generation || oldRows !== 0) {
    throw new Error(
      "Post-purge verification did not confirm feed rotation and old-feed removal."
    );
  }

  process.stdout.write(
    [
      "",
      "PURGE COMPLETE",
      "--------------",
      `Deleted: ${guarded.delete_count}`,
      `Retained by policy: ${guarded.retained_count}`,
      "Old reporting generation rows remaining: 0",
      "New reporting generation created successfully.",
      ""
    ].join("\n")
  );
}

main().catch(error => {
  process.stderr.write(`GUARDED PURGE FAILED: ${error.message}\n`);
  process.exitCode = 1;
});
