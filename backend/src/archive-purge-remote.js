import { REPORTING_DATASETS } from "./reporting-contract.js";

function sqlLiteral(value) {
  return "'" + String(value).replace(/'/g, "''") + "'";
}

export function buildRemoteArchiveSnapshotQuery({
  feedGeneration,
  verifiedRecords
}) {
  if (!/^[a-f0-9]{32}$/.test(String(feedGeneration || ""))) {
    throw new Error("Invalid receipt feed generation.");
  }

  if (!Array.isArray(verifiedRecords)) {
    throw new Error("Verified records must be an array.");
  }

  const statements = [
    {
      kind: "feed",
      sql:
        "SELECT feed_generation " +
        "FROM mirror_feed_state WHERE singleton = 1"
    },
    {
      kind: "mirror",
      sql:
        "SELECT dataset, record_id, revision, operation, sequence " +
        "FROM mirror_changes " +
        `WHERE feed_generation = ${sqlLiteral(feedGeneration)} ` +
        "ORDER BY sequence ASC"
    }
  ];

  for (const dataset of REPORTING_DATASETS) {
    const ids = verifiedRecords
      .filter(record => record.dataset === dataset.name)
      .map(record => String(record.record_id));

    const columns = dataset.columns
      .map(column => `"${column}"`)
      .join(", ");

    const where = ids.length
      ? `"${dataset.key}" IN (${ids.map(sqlLiteral).join(", ")})`
      : "1 = 0";

    statements.push({
      kind: "source",
      dataset: dataset.name,
      sql:
        `SELECT ${columns} FROM "${dataset.name}" ` +
        `WHERE ${where}`
    });
  }

  for (const statement of statements) {
    if (!/^\s*SELECT\b/i.test(statement.sql)) {
      throw new Error("Remote archive snapshot attempted a non-SELECT statement.");
    }
  }

  return {
    feedGeneration,
    statements,
    sql: statements.map(statement => statement.sql).join(";\n") + ";"
  };
}

export function parseWranglerArchiveSnapshot({
  query,
  wranglerResults
}) {
  if (
    !query ||
    !Array.isArray(query.statements) ||
    !Array.isArray(wranglerResults) ||
    wranglerResults.length !== query.statements.length
  ) {
    throw new Error("Unexpected Wrangler snapshot result count.");
  }

  const sourceRowsByDataset = {};
  let currentFeedGeneration = "";
  let mirrorChanges = [];

  for (let i = 0; i < query.statements.length; i++) {
    const descriptor = query.statements[i];
    const result = wranglerResults[i];

    if (
      !result ||
      result.success !== true ||
      !Array.isArray(result.results)
    ) {
      throw new Error("Wrangler snapshot query failed.");
    }

    if (result.meta?.changed_db !== false) {
      throw new Error(
        "Remote snapshot did not prove that the database remained unchanged."
      );
    }

    if (descriptor.kind === "feed") {
      currentFeedGeneration = String(
        result.results[0]?.feed_generation || ""
      );
    } else if (descriptor.kind === "mirror") {
      mirrorChanges = result.results;
    } else if (descriptor.kind === "source") {
      sourceRowsByDataset[descriptor.dataset] = result.results;
    }
  }

  if (!/^[a-f0-9]{32}$/.test(currentFeedGeneration)) {
    throw new Error("Remote database returned an invalid feed generation.");
  }

  return {
    currentFeedGeneration,
    queriedFeedGeneration: query.feedGeneration,
    mirrorChanges,
    sourceRowsByDataset
  };
}

export function createReadOnlySnapshotDb(snapshot) {
  const latestMirror = new Map();

  for (const row of snapshot.mirrorChanges || []) {
    const key =
      String(row.dataset) + "\u0000" + String(row.record_id);

    const existing = latestMirror.get(key);

    const revision = Number(row.revision);
    const sequence = Number(row.sequence);

    if (
      !existing ||
      revision > Number(existing.revision) ||
      (
        revision === Number(existing.revision) &&
        sequence > Number(existing.sequence)
      )
    ) {
      latestMirror.set(key, row);
    }
  }

  const sourceMaps = new Map();

  for (const dataset of REPORTING_DATASETS) {
    const rows =
      snapshot.sourceRowsByDataset?.[dataset.name] || [];

    sourceMaps.set(
      dataset.name,
      new Map(
        rows.map(row => [
          String(row[dataset.key]),
          row
        ])
      )
    );
  }

  return {
    prepare(sql) {
      const normalized = String(sql)
        .replace(/\s+/g, " ")
        .trim();

      if (!/^SELECT\b/i.test(normalized)) {
        throw new Error(
          "Read-only snapshot DB rejected a non-SELECT statement."
        );
      }

      let bindings = [];

      return {
        bind(...values) {
          bindings = values;
          return this;
        },

        async first() {
          if (
            normalized.includes(
              "FROM mirror_feed_state WHERE singleton = 1"
            )
          ) {
            return {
              feed_generation:
                snapshot.currentFeedGeneration
            };
          }

          if (normalized.includes("FROM mirror_changes")) {
            const [
              feedGeneration,
              dataset,
              recordId
            ] = bindings;

            if (
              String(feedGeneration) !==
              snapshot.queriedFeedGeneration
            ) {
              return null;
            }

            return latestMirror.get(
              String(dataset) +
              "\u0000" +
              String(recordId)
            ) || null;
          }

          for (const dataset of REPORTING_DATASETS) {
            if (
              normalized.includes(
                `FROM "${dataset.name}"`
              )
            ) {
              const recordId = bindings[0];

              return (
                sourceMaps
                  .get(dataset.name)
                  .get(String(recordId)) ||
                null
              );
            }
          }

          throw new Error(
            "Read-only snapshot DB received an unexpected query."
          );
        }
      };
    }
  };
}
