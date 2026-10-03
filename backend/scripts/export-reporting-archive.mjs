/*
 * OPERATOR GUIDE — EXPORT STRUCTURED STUDY DATA
 *
 * This script reads the protected reporting feed and creates one private JSON
 * archive pinned to an exact feed generation/checkpoint. It does not delete
 * anything from D1.
 *
 * The archive is what gets uploaded, unopened/unmodified, to the UA OneDrive
 * Incoming Reporting Exports folder for Power Automate + Office Script import.
 * The reporting credential itself is not stored in the archive.
 */

import { readFile, mkdir, writeFile, chmod } from "node:fs/promises";
import { resolve } from "node:path";
import { fetchReportingArchive } from "../src/archive-client.js";

const DEFAULT_BASE_URL =
  "https://dissertation-study-api.professor-jin.workers.dev";

function argument(name) {
  const index = process.argv.indexOf(name);
  if (index === -1) return null;
  const value = process.argv[index + 1];
  if (!value || value.startsWith("--")) {
    throw new Error(`${name} requires a value.`);
  }
  return value;
}

function safeTimestamp(iso) {
  return iso
    .replace(/\.\d{3}Z$/, "Z")
    .replace(/[:]/g, "")
    .replace("T", "-")
    .replace("Z", "Z");
}

async function main() {
  const baseUrl =
    argument("--base-url") ||
    process.env.REPORTING_BASE_URL ||
    DEFAULT_BASE_URL;

  const tokenFile =
    argument("--token-file") ||
    process.env.REPORTING_EXPORT_TOKEN_FILE ||
    "/tmp/reporting-export-token.txt";

  const outputDir = resolve(
    argument("--output-dir") ||
    process.env.REPORTING_EXPORT_OUTPUT_DIR ||
    "./private-reporting-exports"
  );

  let token;
  try {
    token = (await readFile(tokenFile, "utf8")).trim();
  } catch {
    throw new Error(
      `Could not read reporting token file: ${tokenFile}`
    );
  }

  if (token.length < 32) {
    throw new Error("Reporting token file does not contain a valid token.");
  }

  process.stdout.write("Reading reporting manifest and change feed...\n");

  const archive = await fetchReportingArchive({
    baseUrl,
    token
  });

  await mkdir(outputDir, {
    recursive: true,
    mode: 0o700
  });

  const filename =
    `dissertation-reporting-through-${String(archive.through_sequence)
      .padStart(6, "0")}-` +
    `${safeTimestamp(archive.generated_at)}-${archive.export_id}.json`;

  const outputPath = resolve(outputDir, filename);

  await writeFile(
    outputPath,
    `${JSON.stringify(archive, null, 2)}\n`,
    {
      encoding: "utf8",
      mode: 0o600,
      flag: "wx"
    }
  );

  await chmod(outputPath, 0o600);

  process.stdout.write(
    [
      "",
      "Reporting archive created.",
      `Records: ${archive.total_records}`,
      `Feed generation: ${archive.feed_generation}`,
      `Through sequence: ${archive.through_sequence}`,
      `File: ${outputPath}`,
      "",
      "The reporting credential is not stored in this archive.",
      "No Cloudflare data was modified or deleted.",
      ""
    ].join("\n")
  );
}

main().catch(error => {
  process.stderr.write(`EXPORT FAILED: ${error.message}\n`);
  process.exitCode = 1;
});
