import {
  chmod,
  mkdir,
  mkdtemp,
  readFile,
  rm,
  stat,
  writeFile
} from "node:fs/promises";
import { createHash } from "node:crypto";
import { spawnSync } from "node:child_process";
import { tmpdir } from "node:os";
import { basename, join, resolve } from "node:path";
import {
  buildAudioArchiveCore,
  collectAudioReferences,
  finalizeAudioArchiveManifest
} from "../src/audio-archive.js";

const DEFAULT_BUCKET = "dissertation-study-audio";
const WRANGLER_VERSION = "4.145.0";

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

function extensionFor(reference) {
  const candidates = [
    reference.audio_original_filename,
    reference.object_key
  ];

  for (const value of candidates) {
    const name = basename(String(value || ""));
    const match = name.match(/\.([A-Za-z0-9]{1,8})$/);
    if (match) return match[1].toLowerCase();
  }

  return "bin";
}

function run(command, args, options = {}) {
  const result = spawnSync(command, args, {
    stdio: "inherit",
    ...options
  });
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

async function main() {
  const reportingArchivePath = argument("--reporting-archive");
  if (!reportingArchivePath) {
    throw new Error("--reporting-archive is required.");
  }

  const bucket =
    argument("--bucket") ||
    process.env.STUDY_AUDIO_BUCKET ||
    DEFAULT_BUCKET;

  const outputDir = resolve(
    argument("--output-dir") ||
    process.env.AUDIO_ARCHIVE_OUTPUT_DIR ||
    "./private-audio-archives"
  );

  const reportingArchive = JSON.parse(
    await readFile(resolve(reportingArchivePath), "utf8")
  );

  const references = collectAudioReferences(reportingArchive);

  if (references.length === 0) {
    process.stdout.write(
      [
        "",
        "No referenced audio objects were found in this reporting archive.",
        "No R2 objects were read or modified.",
        ""
      ].join("\n")
    );
    return;
  }

  await mkdir(outputDir, { recursive: true, mode: 0o700 });

  const staging = await mkdtemp(
    join(tmpdir(), "dissertation-audio-archive-")
  );
  const objectDir = join(staging, "objects");
  await mkdir(objectDir, { recursive: true, mode: 0o700 });

  try {
    const archivedObjects = [];

    for (let index = 0; index < references.length; index += 1) {
      const reference = references[index];
      const archiveFilename =
        `${String(index + 1).padStart(4, "0")}.${extensionFor(reference)}`;
      const destination = join(objectDir, archiveFilename);

      process.stdout.write(
        `Downloading referenced R2 object ${index + 1}/${references.length}...\n`
      );

      run(
        process.platform === "win32" ? "npx.cmd" : "npx",
        [
          "--yes",
          `wrangler@${WRANGLER_VERSION}`,
          "r2",
          "object",
          "get",
          `${bucket}/${reference.object_key}`,
          "--remote",
          "--file",
          destination
        ]
      );

      const info = await stat(destination);
      archivedObjects.push({
        object_key: reference.object_key,
        archive_filename: archiveFilename,
        sha256: await sha256File(destination),
        byte_length: info.size
      });
    }

    const generatedAt = new Date().toISOString();
    const core = buildAudioArchiveCore({
      reportingArchive,
      archivedObjects,
      bucket,
      generatedAt
    });

    await writeFile(
      join(staging, "manifest.json"),
      `${JSON.stringify(core, null, 2)}\n`,
      { encoding: "utf8", mode: 0o600 }
    );

    const bundleFilename =
      `dissertation-audio-${safeTimestamp(generatedAt)}-` +
      `${core.audio_export_id}.tar.gz`;
    const bundlePath = join(outputDir, bundleFilename);

    run("tar", [
      "-czf",
      bundlePath,
      "-C",
      staging,
      "manifest.json",
      "objects"
    ]);

    await chmod(bundlePath, 0o600);

    const manifest = finalizeAudioArchiveManifest(core, {
      bundleFilename,
      bundleSha256: await sha256File(bundlePath)
    });

    const manifestFilename =
      bundleFilename.replace(/\.tar\.gz$/, ".manifest.json");
    const manifestPath = join(outputDir, manifestFilename);

    await writeFile(
      manifestPath,
      `${JSON.stringify(manifest, null, 2)}\n`,
      {
        encoding: "utf8",
        mode: 0o600,
        flag: "wx"
      }
    );
    await chmod(manifestPath, 0o600);

    process.stdout.write(
      [
        "",
        "AUDIO ARCHIVE CREATED",
        "---------------------",
        `Reporting export ID: ${manifest.reporting_export_id}`,
        `Audio export ID: ${manifest.audio_export_id}`,
        `Objects: ${manifest.object_count}`,
        `Bundle: ${bundlePath}`,
        `Manifest: ${manifestPath}`,
        "",
        "No R2 object was modified or deleted.",
        "Keep both files private. Do not paste the manifest contents into chat.",
        ""
      ].join("\n")
    );
  } finally {
    await rm(staging, { recursive: true, force: true });
  }
}

main().catch(error => {
  process.stderr.write(`AUDIO EXPORT FAILED: ${error.message}\n`);
  process.exitCode = 1;
});
