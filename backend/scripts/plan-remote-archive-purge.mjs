import { readFile } from "node:fs/promises";
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import { fileURLToPath } from "node:url";
import { resolve } from "node:path";

import { validateArchiveReceipt } from "../src/archive-receipt.js";
import { buildArchivePurgePlan } from "../src/archive-purge-plan.js";
import {
  buildRemoteArchiveSnapshotQuery,
  parseWranglerArchiveSnapshot,
  createReadOnlySnapshotDb
} from "../src/archive-purge-remote.js";

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

async function readJson(path, label) {
  if (!path) {
    throw new Error(`${label} path is required.`);
  }

  return JSON.parse(
    await readFile(resolve(path), "utf8")
  );
}

async function main() {
  const archivePath = argument("--archive");
  const receiptPath = argument("--receipt");

  const archiveExport =
    await readJson(archivePath, "Archive");

  const receipt =
    await readJson(receiptPath, "Receipt");

  const verified = validateArchiveReceipt({
    archiveExport,
    receipt
  });

  const query = buildRemoteArchiveSnapshotQuery({
    feedGeneration: verified.feed_generation,
    verifiedRecords: verified.verified_records
  });

  const backendDirectory = fileURLToPath(
    new URL("..", import.meta.url)
  );

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

  const wranglerResults = JSON.parse(stdout);

  const snapshot = parseWranglerArchiveSnapshot({
    query,
    wranglerResults
  });

  const db = createReadOnlySnapshotDb(snapshot);

  const plan = await buildArchivePurgePlan({
    db,
    archiveExport,
    receipt
  });

  const reasons = {};

  for (const record of plan.records) {
    reasons[record.reason] =
      (reasons[record.reason] || 0) + 1;
  }

  process.stdout.write(
    [
      "",
      "ARCHIVE PURGE DRY RUN",
      "---------------------",
      `Export ID: ${plan.export_id}`,
      `Receipt generation: ${plan.receipt_feed_generation}`,
      `Current generation: ${plan.current_feed_generation}`,
      `Generation matches: ${plan.feed_generation_matches}`,
      `Verified receipt records: ${plan.verified_record_count}`,
      `Eligible now: ${plan.eligible_count}`,
      `Kept: ${plan.keep_count}`,
      "",
      "Reason counts:",
      ...Object.entries(reasons)
        .sort(([a], [b]) => a.localeCompare(b))
        .map(([reason, count]) => `  ${reason}: ${count}`),
      "",
      "DRY RUN ONLY.",
      "Only SELECT statements were sent to D1.",
      "No Cloudflare records were modified or deleted.",
      ""
    ].join("\n")
  );
}

main().catch(error => {
  process.stderr.write(
    `PURGE DRY RUN FAILED: ${error.message}\n`
  );
  process.exitCode = 1;
});
