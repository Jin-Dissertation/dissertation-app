import { REPORTING_DATASETS } from "./reporting-contract.js";
import { sha256Hex } from "./archive-export.js";
import { validateArchiveReceipt } from "./archive-receipt.js";

const datasetByName = new Map(
  REPORTING_DATASETS.map(dataset => [dataset.name, dataset])
);

function projectedRecord(dataset, row) {
  return Object.fromEntries(
    dataset.columns.map(column => [column, row?.[column] ?? null])
  );
}

export async function buildArchivePurgePlan({
  db,
  archiveExport,
  receipt
}) {
  if (!db) throw new Error("Archive purge planning requires a database.");

  const verified = validateArchiveReceipt({
    archiveExport,
    receipt
  });

  const feedState = await db
    .prepare(
      "SELECT feed_generation FROM mirror_feed_state WHERE singleton = 1"
    )
    .first();

  const currentGeneration = String(
    feedState?.feed_generation || ""
  );

  const generationMatches =
    currentGeneration === verified.feed_generation;

  const records = [];

  for (const verifiedRecord of verified.verified_records) {
    const dataset = datasetByName.get(verifiedRecord.dataset);

    if (!dataset) {
      records.push({
        ...verifiedRecord,
        status: "keep",
        reason: "unknown_dataset"
      });
      continue;
    }

    if (!generationMatches) {
      records.push({
        ...verifiedRecord,
        status: "keep",
        reason: "feed_generation_changed"
      });
      continue;
    }

    const latestMirror = await db
      .prepare(
        `SELECT revision, operation
         FROM mirror_changes
         WHERE feed_generation = ?1
           AND dataset = ?2
           AND record_id = ?3
         ORDER BY revision DESC, sequence DESC
         LIMIT 1`
      )
      .bind(
        verified.feed_generation,
        dataset.name,
        verifiedRecord.record_id
      )
      .first();

    if (!latestMirror) {
      records.push({
        ...verifiedRecord,
        status: "keep",
        reason: "mirror_record_missing"
      });
      continue;
    }

    if (latestMirror.operation !== "upsert") {
      records.push({
        ...verifiedRecord,
        status: "keep",
        reason: "source_already_deleted"
      });
      continue;
    }

    if (
      Number(latestMirror.revision) !==
      Number(verifiedRecord.revision)
    ) {
      records.push({
        ...verifiedRecord,
        status: "keep",
        reason: "revision_changed",
        current_revision: Number(latestMirror.revision)
      });
      continue;
    }

    const columnSql = dataset.columns
      .map(column => `"${column}"`)
      .join(", ");

    const sourceRow = await db
      .prepare(
        `SELECT ${columnSql}
         FROM "${dataset.name}"
         WHERE "${dataset.key}" = ?1
         LIMIT 1`
      )
      .bind(verifiedRecord.record_id)
      .first();

    if (!sourceRow) {
      records.push({
        ...verifiedRecord,
        status: "keep",
        reason: "source_record_missing"
      });
      continue;
    }

    const currentHash = await sha256Hex(
      projectedRecord(dataset, sourceRow)
    );

    if (currentHash !== verifiedRecord.record_sha256) {
      records.push({
        ...verifiedRecord,
        status: "keep",
        reason: "record_hash_changed",
        current_record_sha256: currentHash
      });
      continue;
    }

    records.push({
      ...verifiedRecord,
      status: "eligible",
      reason: "exact_archive_match",
      current_revision: Number(latestMirror.revision),
      current_record_sha256: currentHash
    });
  }

  const eligibleCount = records.filter(
    record => record.status === "eligible"
  ).length;

  return {
    export_id: verified.export_id,
    receipt_feed_generation: verified.feed_generation,
    current_feed_generation: currentGeneration,
    feed_generation_matches: generationMatches,
    through_sequence: verified.through_sequence,
    verified_record_count: verified.verified_record_count,
    eligible_count: eligibleCount,
    keep_count: records.length - eligibleCount,
    records
  };
}
