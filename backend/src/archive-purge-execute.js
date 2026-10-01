import { REPORTING_DATASETS } from "./reporting-contract.js";

export const ARCHIVE_PURGE_RETAIN_DATASETS = new Set([
  "nonparticipant_button_counts"
]);

const clock = "strftime('%Y-%m-%dT%H:%M:%fZ', 'now')";

function sqlIdentifier(value) {
  return '"' + String(value).replace(/"/g, '""') + '"';
}

function sqlLiteral(value) {
  if (value === null || value === undefined) return "NULL";

  if (typeof value === "number") {
    if (!Number.isFinite(value)) {
      throw new Error("Archive purge cannot encode a non-finite number.");
    }
    return String(value);
  }

  if (typeof value === "boolean") {
    return value ? "1" : "0";
  }

  return "'" + String(value).replace(/'/g, "''") + "'";
}

function projection(dataset, alias) {
  return (
    "json_object(" +
    dataset.columns
      .map(column => `${sqlLiteral(column)}, ${alias}.${sqlIdentifier(column)}`)
      .join(", ") +
    ")"
  );
}

function bootstrapStatement(dataset) {
  return (
    "INSERT INTO mirror_changes " +
    "(feed_generation, dataset, record_id, revision, operation, changed_at, record_json)\n" +
    `SELECT (SELECT feed_generation FROM mirror_feed_state WHERE singleton = 1), ${sqlLiteral(dataset.name)}, ` +
    `r.${sqlIdentifier(dataset.key)}, 1, 'upsert', ${clock},\n       ${projection(dataset, "r")}\n` +
    `FROM ${sqlIdentifier(dataset.name)} r ORDER BY r.${sqlIdentifier(dataset.key)};`
  );
}

function archivedRecordMap(archiveExport) {
  const map = new Map();

  for (const dataset of archiveExport?.datasets || []) {
    for (const archived of dataset?.records || []) {
      map.set(
        String(dataset.name) + "\u0000" + String(archived.record_id),
        archived
      );
    }
  }

  return map;
}

function exactSourcePredicate(dataset, record, alias = "r") {
  return dataset.columns
    .map(column => {
      if (!Object.prototype.hasOwnProperty.call(record, column)) {
        throw new Error(
          `Archive record is missing projected column ${dataset.name}.${column}.`
        );
      }

      return (
        `${alias}.${sqlIdentifier(column)} IS ` +
        sqlLiteral(record[column])
      );
    })
    .join("\n      AND ");
}

function latestMirrorPredicate({
  dataset,
  recordId,
  feedGeneration,
  revision
}) {
  const common =
    `feed_generation = ${sqlLiteral(feedGeneration)} ` +
    `AND dataset = ${sqlLiteral(dataset.name)} ` +
    `AND record_id = ${sqlLiteral(recordId)}`;

  return [
    `(SELECT revision FROM mirror_changes WHERE ${common} ORDER BY revision DESC, sequence DESC LIMIT 1) IS ${Number(revision)}`,
    `(SELECT operation FROM mirror_changes WHERE ${common} ORDER BY revision DESC, sequence DESC LIMIT 1) IS 'upsert'`
  ].join("\n      AND ");
}

export function buildGuardedArchivePurgeSql({
  plan,
  archiveExport
}) {
  if (!plan || !archiveExport) {
    throw new Error("Archive purge SQL requires a plan and archive export.");
  }

  if (!plan.feed_generation_matches) {
    throw new Error("Archive purge refused because the feed generation changed.");
  }

  const generation = String(plan.receipt_feed_generation || "");

  if (!/^[a-f0-9]{32}$/.test(generation)) {
    throw new Error("Archive purge received an invalid feed generation.");
  }

  const datasetByName = new Map(
    REPORTING_DATASETS.map(dataset => [dataset.name, dataset])
  );

  const archivedByKey = archivedRecordMap(archiveExport);

  const targetRecords = [];
  const retainedRecords = [];
  const blockedRecords = [];

  for (const record of plan.records || []) {
    if (ARCHIVE_PURGE_RETAIN_DATASETS.has(record.dataset)) {
      retainedRecords.push(record);
      continue;
    }

    if (record.status !== "eligible") {
      blockedRecords.push(record);
      continue;
    }

    targetRecords.push(record);
  }

  if (blockedRecords.length) {
    throw new Error(
      `Archive purge refused because ${blockedRecords.length} non-retained record(s) are not exact archive matches.`
    );
  }

  if (!targetRecords.length) {
    throw new Error("Archive purge found no deletable records.");
  }

  const guardName =
    "__archive_purge_guard_" +
    String(plan.export_id || "")
      .replace(/[^a-zA-Z0-9]/g, "")
      .slice(0, 16);

  if (guardName.length <= "__archive_purge_guard_".length) {
    throw new Error("Archive purge received an invalid export id.");
  }

  const statements = [
    `DROP TABLE IF EXISTS ${sqlIdentifier(guardName)};`,
    `CREATE TABLE ${sqlIdentifier(guardName)} (ok INTEGER NOT NULL CHECK (ok = 1));`
  ];

  const deleteStatements = [];
  const postDeleteAssertions = [];

  for (const verifiedRecord of targetRecords) {
    const dataset = datasetByName.get(verifiedRecord.dataset);

    if (!dataset) {
      throw new Error(
        `Archive purge received unknown dataset ${verifiedRecord.dataset}.`
      );
    }

    const archived = archivedByKey.get(
      dataset.name + "\u0000" + String(verifiedRecord.record_id)
    );

    if (!archived || !archived.record) {
      throw new Error(
        `Archive purge could not locate archived record ${dataset.name}/${verifiedRecord.record_id}.`
      );
    }

    const sourcePredicate = exactSourcePredicate(
      dataset,
      archived.record,
      "r"
    );

    const mirrorPredicate = latestMirrorPredicate({
      dataset,
      recordId: verifiedRecord.record_id,
      feedGeneration: generation,
      revision: verifiedRecord.revision
    });

    const generationPredicate =
      `(SELECT feed_generation FROM mirror_feed_state WHERE singleton = 1) IS ${sqlLiteral(generation)}`;

    const exactExists =
      `EXISTS (SELECT 1 FROM ${sqlIdentifier(dataset.name)} r\n` +
      `    WHERE r.${sqlIdentifier(dataset.key)} IS ${sqlLiteral(verifiedRecord.record_id)}\n` +
      `      AND ${sourcePredicate})`;

    statements.push(
      `INSERT INTO ${sqlIdentifier(guardName)} (ok)\n` +
      "SELECT CASE WHEN\n" +
      `  ${generationPredicate}\n` +
      `  AND ${mirrorPredicate}\n` +
      `  AND ${exactExists}\n` +
      "THEN 1 ELSE 0 END;"
    );

    deleteStatements.push(
      `DELETE FROM ${sqlIdentifier(dataset.name)} AS r\n` +
      `WHERE r.${sqlIdentifier(dataset.key)} IS ${sqlLiteral(verifiedRecord.record_id)}\n` +
      `  AND ${sourcePredicate}\n` +
      `  AND ${generationPredicate}\n` +
      `  AND ${mirrorPredicate};`
    );

    postDeleteAssertions.push(
      `INSERT INTO ${sqlIdentifier(guardName)} (ok)\n` +
      "SELECT CASE WHEN NOT EXISTS (" +
      `SELECT 1 FROM ${sqlIdentifier(dataset.name)} r ` +
      `WHERE r.${sqlIdentifier(dataset.key)} IS ${sqlLiteral(verifiedRecord.record_id)}` +
      ") THEN 1 ELSE 0 END;"
    );
  }

  statements.push(...deleteStatements);
  statements.push(...postDeleteAssertions);

  statements.push(
    `UPDATE mirror_feed_state\nSET feed_generation = lower(hex(randomblob(16))), created_at = ${clock}\nWHERE singleton = 1 AND feed_generation = ${sqlLiteral(generation)};`
  );

  statements.push(
    `INSERT INTO ${sqlIdentifier(guardName)} (ok)\n` +
    "SELECT CASE WHEN " +
    `(SELECT feed_generation FROM mirror_feed_state WHERE singleton = 1) <> ${sqlLiteral(generation)} ` +
    "THEN 1 ELSE 0 END;"
  );

  for (const dataset of REPORTING_DATASETS) {
    statements.push(bootstrapStatement(dataset));
  }

  statements.push(
    `DELETE FROM mirror_changes WHERE feed_generation = ${sqlLiteral(generation)};`
  );

  statements.push(
    `INSERT INTO ${sqlIdentifier(guardName)} (ok)\n` +
    "SELECT CASE WHEN NOT EXISTS (" +
    `SELECT 1 FROM mirror_changes WHERE feed_generation = ${sqlLiteral(generation)}` +
    ") THEN 1 ELSE 0 END;"
  );

  statements.push(`DROP TABLE ${sqlIdentifier(guardName)};`);

  statements.push(
    "SELECT " +
    `${sqlLiteral("archive_purge_complete")} AS status, ` +
    "(SELECT feed_generation FROM mirror_feed_state WHERE singleton = 1) AS current_generation, " +
    `${targetRecords.length} AS deleted_records, ` +
    `${retainedRecords.length} AS retained_records;`
  );

  return {
    sql: statements.join("\n\n") + "\n",
    delete_count: targetRecords.length,
    retained_count: retainedRecords.length,
    retained_datasets: [...ARCHIVE_PURGE_RETAIN_DATASETS],
    old_feed_generation: generation
  };
}
