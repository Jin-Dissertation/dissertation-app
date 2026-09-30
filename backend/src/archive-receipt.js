import {
  ARCHIVE_EXPORT_FORMAT,
  ARCHIVE_EXPORT_VERSION
} from "./archive-export.js";

export const ARCHIVE_RECEIPT_FORMAT = "dissertation_reporting_archive_receipt";
export const ARCHIVE_RECEIPT_VERSION = 1;

function requireString(value, label) {
  const text = String(value ?? "").trim();
  if (!text) throw new Error(`Missing ${label}.`);
  return text;
}

function requirePositiveRevision(value) {
  const revision = Number(value);
  if (!Number.isSafeInteger(revision) || revision < 1) {
    throw new Error("Invalid archive record revision.");
  }
  return revision;
}

function indexArchiveExport(archiveExport) {
  if (
    !archiveExport ||
    archiveExport.format !== ARCHIVE_EXPORT_FORMAT ||
    Number(archiveExport.format_version) !== ARCHIVE_EXPORT_VERSION
  ) {
    throw new Error("Invalid archive export.");
  }

  const datasets = Array.isArray(archiveExport.datasets)
    ? archiveExport.datasets
    : [];

  const records = new Map();

  for (const dataset of datasets) {
    const datasetName = requireString(dataset?.name, "archive dataset name");
    const datasetRecords = Array.isArray(dataset?.records)
      ? dataset.records
      : [];

    for (const record of datasetRecords) {
      const recordId = requireString(record?.record_id, "archive record ID");
      const key = `${datasetName}\u0000${recordId}`;

      if (records.has(key)) {
        throw new Error("Duplicate record in archive export.");
      }

      records.set(key, {
        dataset: datasetName,
        record_id: recordId,
        revision: requirePositiveRevision(record.revision),
        record_sha256: requireString(record.record_sha256, "archive record hash"),
        archive_token: requireString(record.archive_token, "archive token")
      });
    }
  }

  return records;
}

export function validateArchiveReceipt({ archiveExport, receipt }) {
  const archiveRecords = indexArchiveExport(archiveExport);

  if (
    !receipt ||
    receipt.format !== ARCHIVE_RECEIPT_FORMAT ||
    Number(receipt.format_version) !== ARCHIVE_RECEIPT_VERSION
  ) {
    throw new Error("Invalid archive verification receipt.");
  }

  const exportId = requireString(receipt.export_id, "receipt export ID");
  const feedGeneration = requireString(
    receipt.feed_generation,
    "receipt feed generation"
  );

  if (exportId !== String(archiveExport.export_id)) {
    throw new Error("Receipt export ID does not match archive export.");
  }

  if (feedGeneration !== String(archiveExport.feed_generation)) {
    throw new Error("Receipt feed generation does not match archive export.");
  }

  const throughSequence = Number(receipt.through_sequence);
  if (
    !Number.isSafeInteger(throughSequence) ||
    throughSequence !== Number(archiveExport.through_sequence)
  ) {
    throw new Error("Receipt through sequence does not match archive export.");
  }

  requireString(receipt.verified_at, "receipt verification timestamp");

  if (!Array.isArray(receipt.verified_records)) {
    throw new Error("Receipt verified_records must be an array.");
  }

  const seen = new Set();
  const verifiedRecords = [];

  for (const received of receipt.verified_records) {
    const dataset = requireString(received?.dataset, "receipt dataset");
    const recordId = requireString(received?.record_id, "receipt record ID");
    const key = `${dataset}\u0000${recordId}`;

    if (seen.has(key)) {
      throw new Error("Duplicate record in archive verification receipt.");
    }
    seen.add(key);

    const archived = archiveRecords.get(key);
    if (!archived) {
      throw new Error("Receipt contains a record that was not in the archive export.");
    }

    const revision = requirePositiveRevision(received.revision);
    const recordSha256 = requireString(
      received.record_sha256,
      "receipt record hash"
    );
    const archiveToken = requireString(
      received.archive_token,
      "receipt archive token"
    );

    if (revision !== archived.revision) {
      throw new Error("Receipt record revision does not match archive export.");
    }

    if (recordSha256 !== archived.record_sha256) {
      throw new Error("Receipt record hash does not match archive export.");
    }

    if (archiveToken !== archived.archive_token) {
      throw new Error("Receipt archive token does not match archive export.");
    }

    verifiedRecords.push({
      dataset,
      record_id: recordId,
      revision,
      record_sha256: recordSha256,
      archive_token: archiveToken
    });
  }

  return {
    export_id: exportId,
    feed_generation: feedGeneration,
    through_sequence: throughSequence,
    verified_at: receipt.verified_at,
    verified_record_count: verifiedRecords.length,
    verified_records: verifiedRecords
  };
}
