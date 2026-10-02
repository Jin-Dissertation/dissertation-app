import {
  ARCHIVE_EXPORT_FORMAT,
  ARCHIVE_EXPORT_VERSION,
  stableStringify
} from "./archive-export.js";

export const AUDIO_ARCHIVE_FORMAT = "dissertation_audio_archive_manifest";
export const AUDIO_ARCHIVE_VERSION = 1;
export const AUDIO_ARCHIVE_RECEIPT_FORMAT = "dissertation_audio_archive_receipt";
export const AUDIO_ARCHIVE_RECEIPT_VERSION = 1;

const AUDIO_DATASETS = new Set(["aqg_submissions", "aqg_feedback"]);
const SAFE_OBJECT_KEY = /^aqg\/[A-Za-z0-9_.-]+\/[A-Za-z0-9_.-]+\/[A-Za-z0-9_.-]+$/;
const SHA256_HEX = /^[a-f0-9]{64}$/;

function requireText(value, label) {
  const text = String(value ?? "").trim();
  if (!text) throw new Error(`Missing ${label}.`);
  return text;
}

function requireNonnegativeInteger(value, label) {
  const number = Number(value);
  if (!Number.isSafeInteger(number) || number < 0) {
    throw new Error(`Invalid ${label}.`);
  }
  return number;
}

function requirePositiveInteger(value, label) {
  const number = Number(value);
  if (!Number.isSafeInteger(number) || number < 1) {
    throw new Error(`Invalid ${label}.`);
  }
  return number;
}

function requireSha256(value, label) {
  const hash = requireText(value, label).toLowerCase();
  if (!SHA256_HEX.test(hash)) throw new Error(`Invalid ${label}.`);
  return hash;
}

function randomToken() {
  const bytes = new Uint8Array(24);
  crypto.getRandomValues(bytes);
  return Array.from(bytes, byte => byte.toString(16).padStart(2, "0")).join("");
}

function validateReportingArchiveEnvelope(archive) {
  if (
    !archive ||
    archive.format !== ARCHIVE_EXPORT_FORMAT ||
    Number(archive.format_version) !== ARCHIVE_EXPORT_VERSION
  ) {
    throw new Error("Invalid reporting archive export.");
  }

  requireText(archive.export_id, "reporting export ID");
  requireText(archive.feed_generation, "reporting feed generation");
  requireNonnegativeInteger(archive.through_sequence, "reporting through sequence");

  if (!Array.isArray(archive.datasets)) {
    throw new Error("Reporting archive datasets must be an array.");
  }
}

function mergeMetadata(target, source) {
  for (const field of [
    "audio_original_filename",
    "audio_mime_type",
    "audio_duration_seconds"
  ]) {
    const value = source[field];
    if (value === undefined || value === null || value === "") continue;

    if (
      target[field] !== undefined &&
      target[field] !== null &&
      target[field] !== "" &&
      String(target[field]) !== String(value)
    ) {
      throw new Error(`Conflicting ${field} for one audio object.`);
    }

    target[field] = value;
  }
}

export function collectAudioReferences(reportingArchive) {
  validateReportingArchiveEnvelope(reportingArchive);

  const byKey = new Map();

  for (const dataset of reportingArchive.datasets) {
    const datasetName = requireText(dataset?.name, "reporting dataset name");
    if (!AUDIO_DATASETS.has(datasetName)) continue;

    if (!Array.isArray(dataset.records)) {
      throw new Error(`Invalid records for ${datasetName}.`);
    }

    for (const archived of dataset.records) {
      const record = archived?.record;
      if (!record || typeof record !== "object" || Array.isArray(record)) {
        throw new Error(`Invalid record payload in ${datasetName}.`);
      }

      const objectKey = String(record.audio_object_key ?? "").trim();
      if (!objectKey) continue;

      if (!SAFE_OBJECT_KEY.test(objectKey)) {
        throw new Error(`Unsafe or unexpected R2 audio object key: ${objectKey}`);
      }

      const recordId = requireText(archived.record_id, "archive record ID");
      const revision = requirePositiveInteger(
        archived.revision,
        "archive record revision"
      );

      let entry = byKey.get(objectKey);
      if (!entry) {
        entry = {
          object_key: objectKey,
          references: [],
          audio_original_filename: "",
          audio_mime_type: "",
          audio_duration_seconds: null
        };
        byKey.set(objectKey, entry);
      }

      entry.references.push({
        dataset: datasetName,
        record_id: recordId,
        revision
      });

      mergeMetadata(entry, record);
    }
  }

  return [...byKey.values()]
    .map(entry => ({
      ...entry,
      references: entry.references.sort((a, b) =>
        a.dataset.localeCompare(b.dataset) ||
        a.record_id.localeCompare(b.record_id)
      )
    }))
    .sort((a, b) => a.object_key.localeCompare(b.object_key));
}

function validateArchivedObjectInput(object) {
  const objectKey = requireText(object?.object_key, "audio object key");
  if (!SAFE_OBJECT_KEY.test(objectKey)) {
    throw new Error(`Unsafe or unexpected R2 audio object key: ${objectKey}`);
  }

  const archiveFilename = requireText(
    object?.archive_filename,
    "audio archive filename"
  );
  if (
    archiveFilename.includes("/") ||
    archiveFilename.includes("\\") ||
    archiveFilename === "." ||
    archiveFilename === ".."
  ) {
    throw new Error("Invalid audio archive filename.");
  }

  return {
    object_key: objectKey,
    archive_filename: archiveFilename,
    sha256: requireSha256(object.sha256, "audio object SHA-256"),
    byte_length: requireNonnegativeInteger(
      object.byte_length,
      "audio object byte length"
    )
  };
}

export function buildAudioArchiveCore({
  reportingArchive,
  archivedObjects,
  bucket = "dissertation-study-audio",
  generatedAt = new Date().toISOString(),
  audioExportId = crypto.randomUUID(),
  archiveTokenFactory = randomToken
}) {
  validateReportingArchiveEnvelope(reportingArchive);

  if (!Array.isArray(archivedObjects)) {
    throw new Error("Archived audio objects must be an array.");
  }

  const references = collectAudioReferences(reportingArchive);
  const downloaded = new Map();

  for (const raw of archivedObjects) {
    const object = validateArchivedObjectInput(raw);
    if (downloaded.has(object.object_key)) {
      throw new Error("Duplicate downloaded audio object.");
    }
    downloaded.set(object.object_key, object);
  }

  if (downloaded.size !== references.length) {
    throw new Error("Downloaded audio object count does not match archive references.");
  }

  const objects = references.map((reference, index) => {
    const downloadedObject = downloaded.get(reference.object_key);
    if (!downloadedObject) {
      throw new Error(`Referenced audio object was not downloaded: ${reference.object_key}`);
    }

    return {
      object_key: reference.object_key,
      archive_filename: downloadedObject.archive_filename,
      sha256: downloadedObject.sha256,
      byte_length: downloadedObject.byte_length,
      archive_token: requireText(
        archiveTokenFactory(),
        "audio archive token"
      ),
      references: reference.references,
      audio_original_filename: reference.audio_original_filename || "",
      audio_mime_type: reference.audio_mime_type || "",
      audio_duration_seconds:
        reference.audio_duration_seconds === undefined
          ? null
          : reference.audio_duration_seconds
    };
  });

  const filenames = new Set();
  for (const object of objects) {
    if (filenames.has(object.archive_filename)) {
      throw new Error("Duplicate filename inside audio archive.");
    }
    filenames.add(object.archive_filename);
  }

  return {
    format: AUDIO_ARCHIVE_FORMAT,
    format_version: AUDIO_ARCHIVE_VERSION,
    audio_export_id: requireText(audioExportId, "audio export ID"),
    generated_at: requireText(generatedAt, "audio export timestamp"),
    bucket: requireText(bucket, "audio bucket"),
    reporting_export_id: String(reportingArchive.export_id),
    reporting_feed_generation: String(reportingArchive.feed_generation),
    reporting_through_sequence: Number(reportingArchive.through_sequence),
    object_count: objects.length,
    objects
  };
}

export function finalizeAudioArchiveManifest(
  core,
  { bundleFilename, bundleSha256 }
) {
  validateAudioArchiveCore(core);

  const filename = requireText(bundleFilename, "audio bundle filename");
  if (filename.includes("/") || filename.includes("\\")) {
    throw new Error("Invalid audio bundle filename.");
  }

  return {
    ...core,
    bundle_filename: filename,
    bundle_sha256: requireSha256(bundleSha256, "audio bundle SHA-256")
  };
}

export function audioArchiveCoreFromManifest(manifest) {
  const { bundle_filename, bundle_sha256, ...core } = manifest || {};
  return core;
}

export function validateAudioArchiveCore(core) {
  if (
    !core ||
    core.format !== AUDIO_ARCHIVE_FORMAT ||
    Number(core.format_version) !== AUDIO_ARCHIVE_VERSION
  ) {
    throw new Error("Invalid audio archive manifest.");
  }

  requireText(core.audio_export_id, "audio export ID");
  requireText(core.generated_at, "audio export timestamp");
  requireText(core.bucket, "audio bucket");
  requireText(core.reporting_export_id, "reporting export ID");
  requireText(core.reporting_feed_generation, "reporting feed generation");
  requireNonnegativeInteger(
    core.reporting_through_sequence,
    "reporting through sequence"
  );

  if (!Array.isArray(core.objects)) {
    throw new Error("Audio archive objects must be an array.");
  }

  if (
    requireNonnegativeInteger(core.object_count, "audio object count") !==
    core.objects.length
  ) {
    throw new Error("Audio archive object count does not match manifest.");
  }

  const keys = new Set();
  const filenames = new Set();

  for (const object of core.objects) {
    const checked = validateArchivedObjectInput(object);
    requireText(object.archive_token, "audio archive token");

    if (!Array.isArray(object.references) || object.references.length < 1) {
      throw new Error("Audio archive object is missing source references.");
    }

    for (const reference of object.references) {
      const dataset = requireText(reference?.dataset, "audio source dataset");
      if (!AUDIO_DATASETS.has(dataset)) {
        throw new Error("Unexpected audio source dataset.");
      }
      requireText(reference.record_id, "audio source record ID");
      requirePositiveInteger(reference.revision, "audio source record revision");
    }

    if (keys.has(checked.object_key)) {
      throw new Error("Duplicate audio object key in manifest.");
    }
    if (filenames.has(checked.archive_filename)) {
      throw new Error("Duplicate audio archive filename in manifest.");
    }
    keys.add(checked.object_key);
    filenames.add(checked.archive_filename);
  }

  return core;
}

export function validateAudioArchiveManifest(manifest) {
  validateAudioArchiveCore(audioArchiveCoreFromManifest(manifest));
  requireText(manifest.bundle_filename, "audio bundle filename");
  requireSha256(manifest.bundle_sha256, "audio bundle SHA-256");
  return manifest;
}

export function buildAudioArchiveReceipt(
  manifest,
  { verifiedAt = new Date().toISOString() } = {}
) {
  validateAudioArchiveManifest(manifest);

  return {
    format: AUDIO_ARCHIVE_RECEIPT_FORMAT,
    format_version: AUDIO_ARCHIVE_RECEIPT_VERSION,
    audio_export_id: manifest.audio_export_id,
    reporting_export_id: manifest.reporting_export_id,
    bundle_filename: manifest.bundle_filename,
    bundle_sha256: manifest.bundle_sha256,
    verified_at: requireText(verifiedAt, "audio verification timestamp"),
    verified_objects: manifest.objects.map(object => ({
      object_key: object.object_key,
      sha256: object.sha256,
      byte_length: object.byte_length,
      archive_token: object.archive_token
    }))
  };
}

export function validateAudioArchiveReceipt({ manifest, receipt }) {
  validateAudioArchiveManifest(manifest);

  if (
    !receipt ||
    receipt.format !== AUDIO_ARCHIVE_RECEIPT_FORMAT ||
    Number(receipt.format_version) !== AUDIO_ARCHIVE_RECEIPT_VERSION
  ) {
    throw new Error("Invalid audio archive receipt.");
  }

  for (const [field, label] of [
    ["audio_export_id", "audio export ID"],
    ["reporting_export_id", "reporting export ID"],
    ["bundle_filename", "audio bundle filename"],
    ["bundle_sha256", "audio bundle SHA-256"]
  ]) {
    if (String(receipt[field] ?? "") !== String(manifest[field] ?? "")) {
      throw new Error(`Receipt ${label} does not match audio archive.`);
    }
  }

  requireText(receipt.verified_at, "audio verification timestamp");

  if (!Array.isArray(receipt.verified_objects)) {
    throw new Error("Audio receipt verified_objects must be an array.");
  }

  if (receipt.verified_objects.length !== manifest.objects.length) {
    throw new Error("Audio receipt object count does not match archive.");
  }

  const archived = new Map(
    manifest.objects.map(object => [object.object_key, object])
  );
  const seen = new Set();

  for (const received of receipt.verified_objects) {
    const objectKey = requireText(received?.object_key, "receipt audio object key");
    if (seen.has(objectKey)) throw new Error("Duplicate object in audio receipt.");
    seen.add(objectKey);

    const expected = archived.get(objectKey);
    if (!expected) throw new Error("Audio receipt contains an unknown object.");

    if (
      requireSha256(received.sha256, "receipt audio object SHA-256") !==
      expected.sha256
    ) {
      throw new Error("Audio receipt object hash does not match archive.");
    }

    if (
      requireNonnegativeInteger(
        received.byte_length,
        "receipt audio byte length"
      ) !== expected.byte_length
    ) {
      throw new Error("Audio receipt object size does not match archive.");
    }

    if (
      requireText(received.archive_token, "receipt audio archive token") !==
      expected.archive_token
    ) {
      throw new Error("Audio receipt archive token does not match archive.");
    }
  }

  return {
    audio_export_id: receipt.audio_export_id,
    reporting_export_id: receipt.reporting_export_id,
    bundle_sha256: receipt.bundle_sha256,
    verified_at: receipt.verified_at,
    verified_object_count: receipt.verified_objects.length,
    verified_objects: receipt.verified_objects
  };
}

export function manifestsMatchCore(manifest, innerCore) {
  validateAudioArchiveManifest(manifest);
  validateAudioArchiveCore(innerCore);
  return stableStringify(audioArchiveCoreFromManifest(manifest)) ===
    stableStringify(innerCore);
}
