/*
 * OPERATOR GUIDE — VERIFY AUDIO AFTER ONEDRIVE ROUND TRIP
 *
 * Run this against the manifest created during audio export and the bundle that
 * has been downloaded back from UA OneDrive. Verification checks the package
 * identity plus hashes/byte lengths before creating the verified audio receipt.
 *
 * Do not use the original local bundle as the "retrieved" bundle; the purpose
 * is to prove the UA-stored copy can be retrieved intact.
 */

import {
  chmod,
  mkdtemp,
  readFile,
  rm,
  stat,
  writeFile,
  mkdir
} from "node:fs/promises";
import { createHash } from "node:crypto";
import { spawnSync } from "node:child_process";
import { tmpdir } from "node:os";
import { dirname, join, resolve } from "node:path";
import {
  buildAudioArchiveReceipt,
  manifestsMatchCore,
  validateAudioArchiveManifest
} from "../src/audio-archive.js";

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
  return String(iso)
    .replace(/\.\d{3}Z$/, "Z")
    .replace(/:/g, "")
    .replace("T", "-");
}

function runCapture(command, args) {
  const result = spawnSync(command, args, {
    encoding: "utf8",
    stdio: ["ignore", "pipe", "pipe"]
  });
  if (result.status !== 0) {
    throw new Error(
      `${command} failed: ${String(result.stderr || result.error?.message || "").trim()}`
    );
  }
  return String(result.stdout || "");
}

function run(command, args) {
  const result = spawnSync(command, args, { stdio: "inherit" });
  if (result.status !== 0) {
    throw new Error(
      `${command} failed${result.error ? `: ${result.error.message}` : ""}`
    );
  }
}

async function sha256File(path) {
  const data = await readFile(path);
  return createHash("sha256").update(data).digest("hex");
}

function validateTarListing(listing, manifest) {
  const entries = listing
    .split(/\r?\n/)
    .map(value => value.trim())
    .filter(Boolean);

  for (const entry of entries) {
    if (
      entry.startsWith("/") ||
      entry.includes("../") ||
      entry.includes("\\")
    ) {
      throw new Error("Audio bundle contains an unsafe path.");
    }
  }

  const expectedFiles = new Set([
    "manifest.json",
    ...manifest.objects.map(object => `objects/${object.archive_filename}`)
  ]);

  const actualFiles = new Set(
    entries.filter(entry => !entry.endsWith("/"))
  );

  if (actualFiles.size !== expectedFiles.size) {
    throw new Error("Audio bundle contains unexpected or missing files.");
  }

  for (const expected of expectedFiles) {
    if (!actualFiles.has(expected)) {
      throw new Error(`Audio bundle is missing ${expected}.`);
    }
  }

  for (const actual of actualFiles) {
    if (!expectedFiles.has(actual)) {
      throw new Error(`Audio bundle contains unexpected file ${actual}.`);
    }
  }
}

async function main() {
  const manifestArgument = argument("--manifest");
  const bundleArgument = argument("--retrieved-bundle");

  if (!manifestArgument) throw new Error("--manifest is required.");
  if (!bundleArgument) throw new Error("--retrieved-bundle is required.");

  const manifestPath = resolve(manifestArgument);
  const bundlePath = resolve(bundleArgument);
  const manifest = JSON.parse(await readFile(manifestPath, "utf8"));
  validateAudioArchiveManifest(manifest);

  const bundleHash = await sha256File(bundlePath);
  if (bundleHash !== manifest.bundle_sha256) {
    throw new Error(
      "Retrieved OneDrive bundle SHA-256 does not match the original audio archive."
    );
  }

  const listing = runCapture("tar", ["-tzf", bundlePath]);
  validateTarListing(listing, manifest);

  const staging = await mkdtemp(
    join(tmpdir(), "dissertation-audio-verify-")
  );

  try {
    run("tar", ["-xzf", bundlePath, "-C", staging]);

    const innerCore = JSON.parse(
      await readFile(join(staging, "manifest.json"), "utf8")
    );

    if (!manifestsMatchCore(manifest, innerCore)) {
      throw new Error(
        "Manifest inside retrieved bundle does not match the external archive manifest."
      );
    }

    for (const object of manifest.objects) {
      const objectPath = join(
        staging,
        "objects",
        object.archive_filename
      );
      const info = await stat(objectPath);
      if (!info.isFile() || info.size !== object.byte_length) {
        throw new Error(
          `Retrieved audio object size mismatch: ${object.object_key}`
        );
      }

      const hash = await sha256File(objectPath);
      if (hash !== object.sha256) {
        throw new Error(
          `Retrieved audio object hash mismatch: ${object.object_key}`
        );
      }
    }

    const verifiedAt = new Date().toISOString();
    const receipt = buildAudioArchiveReceipt(manifest, { verifiedAt });

    const outputDir = resolve(
      argument("--output-dir") ||
      process.env.AUDIO_ARCHIVE_OUTPUT_DIR ||
      dirname(manifestPath)
    );
    await mkdir(outputDir, { recursive: true, mode: 0o700 });

    const receiptFilename =
      `verified-audio-receipt-${manifest.audio_export_id}-` +
      `${safeTimestamp(verifiedAt)}.json`;
    const receiptPath = join(outputDir, receiptFilename);

    await writeFile(
      receiptPath,
      `${JSON.stringify(receipt, null, 2)}\n`,
      {
        encoding: "utf8",
        mode: 0o600,
        flag: "wx"
      }
    );
    await chmod(receiptPath, 0o600);

    process.stdout.write(
      [
        "",
        "AUDIO ARCHIVE VERIFIED",
        "----------------------",
        `Audio export ID: ${manifest.audio_export_id}`,
        `Verified objects: ${manifest.object_count}`,
        `Bundle SHA-256: ${manifest.bundle_sha256}`,
        `Receipt: ${receiptPath}`,
        "",
        "This receipt verifies the retrieved OneDrive copy only.",
        "It does not delete or authorize automatic deletion of any R2 object.",
        "Keep the receipt private and do not paste its contents into chat.",
        ""
      ].join("\n")
    );
  } finally {
    await rm(staging, { recursive: true, force: true });
  }
}

main().catch(error => {
  process.stderr.write(`AUDIO VERIFY FAILED: ${error.message}\n`);
  process.exitCode = 1;
});
