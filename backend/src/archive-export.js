import { REPORTING_DATASETS, REPORTING_PROTOCOL_VERSION } from "./reporting-contract.js";

export const ARCHIVE_EXPORT_FORMAT = "dissertation_reporting_archive_export";
export const ARCHIVE_EXPORT_VERSION = 1;

function canonicalize(value) {
  if (value === null || typeof value !== "object") return value;
  if (Array.isArray(value)) return value.map(canonicalize);

  return Object.fromEntries(
    Object.keys(value)
      .sort()
      .map(key => [key, canonicalize(value[key])])
  );
}

export function stableStringify(value) {
  return JSON.stringify(canonicalize(value));
}

function bytesToHex(bytes) {
  return Array.from(bytes, byte => byte.toString(16).padStart(2, "0")).join("");
}

export async function sha256Hex(value) {
  const bytes = new TextEncoder().encode(
    typeof value === "string" ? value : stableStringify(value)
  );
  const digest = await crypto.subtle.digest("SHA-256", bytes);
  return bytesToHex(new Uint8Array(digest));
}

function randomArchiveToken() {
  const bytes = new Uint8Array(24);
  crypto.getRandomValues(bytes);

  let binary = "";
  for (const byte of bytes) binary += String.fromCharCode(byte);

  return btoa(binary)
    .replace(/\+/g, "-")
    .replace(/\//g, "_")
    .replace(/=+$/g, "");
}

function validateManifest(manifest) {
  if (!manifest || manifest.ok !== true) {
    throw new Error("Reporting manifest is not valid.");
  }

  if (Number(manifest.protocol_version) !== REPORTING_PROTOCOL_VERSION) {
    throw new Error("Unsupported reporting protocol version.");
  }

  if (!/^[a-f0-9]{32}$/.test(String(manifest.feed_generation || ""))) {
    throw new Error("Invalid reporting feed generation.");
  }

  if (!Number.isSafeInteger(Number(manifest.latest_sequence)) || Number(manifest.latest_sequence) < 0) {
    throw new Error("Invalid reporting latest sequence.");
  }

  const received = Array.isArray(manifest.datasets) ? manifest.datasets : [];
  if (stableStringify(received) !== stableStringify(REPORTING_DATASETS)) {
    throw new Error("Reporting dataset contract does not match this application version.");
  }
}

function validateRecord(dataset, record) {
  if (!record || typeof record !== "object" || Array.isArray(record)) {
    throw new Error(`Invalid ${dataset.name} record.`);
  }

  if (String(record[dataset.key] ?? "") === "") {
    throw new Error(`Missing ${dataset.name} record key.`);
  }

  const keys = Object.keys(record).sort();
  const expected = [...dataset.columns].sort();

  if (stableStringify(keys) !== stableStringify(expected)) {
    throw new Error(`Unexpected columns in ${dataset.name} record.`);
  }
}

export async function buildArchiveExport({
  manifest,
  changes,
  throughSequence = Number(manifest?.latest_sequence),
  generatedAt = new Date().toISOString(),
  exportId = crypto.randomUUID(),
  archiveTokenFactory = randomArchiveToken
}) {
  validateManifest(manifest);

  const through = Number(throughSequence);
  if (!Number.isSafeInteger(through) || through < 0 || through > Number(manifest.latest_sequence)) {
    throw new Error("Invalid archive through sequence.");
  }

  if (!Array.isArray(changes)) throw new Error("Reporting changes must be an array.");

  const datasetByName = new Map(REPORTING_DATASETS.map(dataset => [dataset.name, dataset]));
  const current = new Map();
  let previousSequence = 0;

  for (const change of changes) {
    const sequence = Number(change?.sequence);
    if (!Number.isSafeInteger(sequence) || sequence <= previousSequence || sequence > through) {
      throw new Error("Reporting changes are not in valid sequence order.");
    }
    previousSequence = sequence;

    const dataset = datasetByName.get(change.dataset);
    if (!dataset) throw new Error(`Unknown reporting dataset: ${change.dataset}`);

    const recordId = String(change.record_id ?? "");
    const revision = Number(change.revision);
    if (!recordId || !Number.isSafeInteger(revision) || revision < 1) {
      throw new Error("Invalid reporting change key or revision.");
    }

    const key = `${change.dataset}\u0000${recordId}`;

    if (change.operation === "delete") {
      current.delete(key);
      continue;
    }

    if (change.operation !== "upsert") {
      throw new Error(`Unknown reporting operation: ${change.operation}`);
    }

    validateRecord(dataset, change.record);

    if (String(change.record[dataset.key]) !== recordId) {
      throw new Error("Reporting record ID does not match its record payload.");
    }

    current.set(key, {
      dataset: dataset.name,
      record_id: recordId,
      revision,
      record: change.record
    });
  }

  const grouped = new Map(REPORTING_DATASETS.map(dataset => [dataset.name, []]));

  for (const item of current.values()) {
    const recordSha256 = await sha256Hex(item.record);
    grouped.get(item.dataset).push({
      record_id: item.record_id,
      revision: item.revision,
      record_sha256: recordSha256,
      archive_token: archiveTokenFactory(),
      record: item.record
    });
  }

  const datasets = REPORTING_DATASETS.map(dataset => {
    const records = grouped.get(dataset.name)
      .sort((a, b) => a.record_id.localeCompare(b.record_id));

    return {
      name: dataset.name,
      key: dataset.key,
      columns: dataset.columns,
      record_count: records.length,
      records
    };
  });

  const totalRecords = datasets.reduce((sum, dataset) => sum + dataset.record_count, 0);

  return {
    format: ARCHIVE_EXPORT_FORMAT,
    format_version: ARCHIVE_EXPORT_VERSION,
    protocol_version: REPORTING_PROTOCOL_VERSION,
    export_id: exportId,
    generated_at: generatedAt,
    feed_generation: manifest.feed_generation,
    through_sequence: through,
    total_records: totalRecords,
    datasets
  };
}
