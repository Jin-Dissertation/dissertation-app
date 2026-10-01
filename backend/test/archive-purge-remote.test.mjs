import test from "node:test";
import assert from "node:assert/strict";

import {
  ARCHIVE_EXPORT_FORMAT,
  ARCHIVE_EXPORT_VERSION,
  sha256Hex
} from "../src/archive-export.js";

import {
  ARCHIVE_RECEIPT_FORMAT,
  ARCHIVE_RECEIPT_VERSION
} from "../src/archive-receipt.js";

import { buildArchivePurgePlan } from "../src/archive-purge-plan.js";

import {
  buildRemoteArchiveSnapshotQuery,
  parseWranglerArchiveSnapshot,
  createReadOnlySnapshotDb
} from "../src/archive-purge-remote.js";

test(
  "remote purge snapshot is SELECT-only and integrates with purge planner",
  async () => {
    const generation = "a".repeat(32);

    const record = {
      button_id: "synthetic-button",
      press_count: 3,
      updated_at: "2026-10-01T00:00:00.000Z"
    };

    const hash = await sha256Hex(record);

    const archived = {
      record_id: record.button_id,
      revision: 1,
      record_sha256: hash,
      archive_token:
        "synthetic-archive-token-1234567890",
      record
    };

    const archiveExport = {
      format: ARCHIVE_EXPORT_FORMAT,
      format_version: ARCHIVE_EXPORT_VERSION,
      protocol_version: 1,
      export_id: "synthetic-remote-plan",
      generated_at: "2026-10-01T00:01:00.000Z",
      feed_generation: generation,
      through_sequence: 1,
      total_records: 1,
      datasets: [
        {
          name: "nonparticipant_button_counts",
          key: "button_id",
          columns: [
            "button_id",
            "press_count",
            "updated_at"
          ],
          record_count: 1,
          records: [archived]
        }
      ]
    };

    const receipt = {
      format: ARCHIVE_RECEIPT_FORMAT,
      format_version: ARCHIVE_RECEIPT_VERSION,
      export_id: archiveExport.export_id,
      feed_generation: generation,
      through_sequence: 1,
      verified_at: "2026-10-01T00:02:00.000Z",
      verified_records: [
        {
          dataset: "nonparticipant_button_counts",
          record_id: archived.record_id,
          revision: archived.revision,
          record_sha256: archived.record_sha256,
          archive_token: archived.archive_token
        }
      ]
    };

    const query = buildRemoteArchiveSnapshotQuery({
      feedGeneration: generation,
      verifiedRecords: receipt.verified_records
    });

    assert.equal(query.statements.length, 10);

    assert.ok(
      query.statements.every(statement =>
        /^\s*SELECT\b/i.test(statement.sql)
      )
    );

    const wranglerResults =
      query.statements.map(statement => {
        let results = [];

        if (statement.kind === "feed") {
          results = [
            { feed_generation: generation }
          ];
        } else if (statement.kind === "mirror") {
          results = [
            {
              dataset:
                "nonparticipant_button_counts",
              record_id:
                "synthetic-button",
              revision: 1,
              operation: "upsert",
              sequence: 1
            }
          ];
        } else if (
          statement.dataset ===
          "nonparticipant_button_counts"
        ) {
          results = [record];
        }

        return {
          success: true,
          results,
          meta: {
            changed_db: false
          }
        };
      });

    const snapshot =
      parseWranglerArchiveSnapshot({
        query,
        wranglerResults
      });

    const db =
      createReadOnlySnapshotDb(snapshot);

    const plan = await buildArchivePurgePlan({
      db,
      archiveExport,
      receipt
    });

    assert.equal(
      plan.feed_generation_matches,
      true
    );

    assert.equal(plan.eligible_count, 1);
    assert.equal(plan.keep_count, 0);
  }
);

test(
  "remote purge snapshot rejects any result that does not prove read-only behavior",
  () => {
    const generation = "b".repeat(32);

    const query =
      buildRemoteArchiveSnapshotQuery({
        feedGeneration: generation,
        verifiedRecords: []
      });

    const wranglerResults =
      query.statements.map(() => ({
        success: true,
        results: [],
        meta: {
          changed_db: false
        }
      }));

    wranglerResults[0].results = [
      { feed_generation: generation }
    ];

    wranglerResults[1].meta.changed_db = true;

    assert.throws(
      () =>
        parseWranglerArchiveSnapshot({
          query,
          wranglerResults
        }),
      /remained unchanged/
    );
  }
);
