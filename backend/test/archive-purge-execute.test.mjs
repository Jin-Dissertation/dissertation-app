import test from "node:test";
import assert from "node:assert/strict";

import { REPORTING_DATASETS } from "../src/reporting-contract.js";
import { buildGuardedArchivePurgeSql } from "../src/archive-purge-execute.js";

function recordFor(dataset, overrides = {}) {
  return Object.fromEntries(
    dataset.columns.map(column => [
      column,
      Object.prototype.hasOwnProperty.call(overrides, column)
        ? overrides[column]
        : null
    ])
  );
}

test("guarded purge deletes study rows but retains cumulative nonparticipant counters", () => {
  const generation = "a".repeat(32);
  const aqg = REPORTING_DATASETS.find(dataset => dataset.name === "aqg_submissions");
  const counts = REPORTING_DATASETS.find(dataset => dataset.name === "nonparticipant_button_counts");

  const aqgRecord = recordFor(aqg, {
    record_id: "synthetic-study-row",
    participant_id: "TEST001",
    session_id: "S0001",
    status: "submitted",
    submitted_at: "2026-10-01T00:00:00.000Z",
    created_at: "2026-10-01T00:00:00.000Z"
  });

  const countRecord = recordFor(counts, {
    button_id: "synthetic-button",
    press_count: 7,
    updated_at: "2026-10-01T00:00:00.000Z"
  });

  const archiveExport = {
    export_id: "synthetic-export-1234",
    datasets: [
      {
        name: aqg.name,
        records: [
          {
            record_id: aqgRecord.record_id,
            revision: 1,
            record: aqgRecord
          }
        ]
      },
      {
        name: counts.name,
        records: [
          {
            record_id: countRecord.button_id,
            revision: 1,
            record: countRecord
          }
        ]
      }
    ]
  };

  const plan = {
    export_id: archiveExport.export_id,
    receipt_feed_generation: generation,
    feed_generation_matches: true,
    records: [
      {
        dataset: aqg.name,
        record_id: aqgRecord.record_id,
        revision: 1,
        status: "eligible"
      },
      {
        dataset: counts.name,
        record_id: countRecord.button_id,
        revision: 1,
        status: "eligible"
      }
    ]
  };

  const result = buildGuardedArchivePurgeSql({
    plan,
    archiveExport
  });

  assert.equal(result.delete_count, 1);
  assert.equal(result.retained_count, 1);

  assert.match(result.sql, /DELETE FROM "aqg_submissions"/);
  assert.doesNotMatch(result.sql, /DELETE FROM "nonparticipant_button_counts"/);
  assert.match(result.sql, /UPDATE mirror_feed_state/);
  assert.match(result.sql, /DELETE FROM mirror_changes WHERE feed_generation/);

  const reseedIndex = result.sql.indexOf("INSERT INTO mirror_changes");
  const oldFeedDeleteIndex = result.sql.lastIndexOf(
    "DELETE FROM mirror_changes WHERE feed_generation"
  );

  assert.ok(reseedIndex !== -1);
  assert.ok(oldFeedDeleteIndex > reseedIndex);
});

test("guarded purge refuses partial deletion when any non-retained record is no longer exact", () => {
  const generation = "b".repeat(32);
  const aqg = REPORTING_DATASETS.find(dataset => dataset.name === "aqg_events");

  const record = recordFor(aqg, {
    record_id: "synthetic-event",
    event_id: "event-1",
    event_timestamp: "2026-10-01T00:00:00.000Z",
    participant_id: "TEST001",
    session_id: "S0001",
    event_type: "synthetic",
    created_at: "2026-10-01T00:00:00.000Z"
  });

  const archiveExport = {
    export_id: "synthetic-export-blocked",
    datasets: [
      {
        name: aqg.name,
        records: [
          {
            record_id: record.record_id,
            revision: 1,
            record
          }
        ]
      }
    ]
  };

  const plan = {
    export_id: archiveExport.export_id,
    receipt_feed_generation: generation,
    feed_generation_matches: true,
    records: [
      {
        dataset: aqg.name,
        record_id: record.record_id,
        revision: 2,
        status: "keep",
        reason: "revision_changed"
      }
    ]
  };

  assert.throws(
    () => buildGuardedArchivePurgeSql({ plan, archiveExport }),
    /not exact archive matches/
  );
});

test("guarded purge embeds exact source-value and mirror revision guards", () => {
  const generation = "c".repeat(32);
  const dataset = REPORTING_DATASETS.find(
    item => item.name === "training_submission_items"
  );

  const record = recordFor(dataset, {
    record_id: "synthetic-item",
    participant_id: "TEST001",
    session_id: "S0001",
    content_version: "v1",
    item_number: 3,
    response_text: "12",
    response_ms: 450,
    created_at: "2026-10-01T00:00:00.000Z"
  });

  const archiveExport = {
    export_id: "synthetic-export-guards",
    datasets: [
      {
        name: dataset.name,
        records: [
          {
            record_id: record.record_id,
            revision: 4,
            record
          }
        ]
      }
    ]
  };

  const plan = {
    export_id: archiveExport.export_id,
    receipt_feed_generation: generation,
    feed_generation_matches: true,
    records: [
      {
        dataset: dataset.name,
        record_id: record.record_id,
        revision: 4,
        status: "eligible"
      }
    ]
  };

  const result = buildGuardedArchivePurgeSql({
    plan,
    archiveExport
  });

  assert.match(result.sql, /"response_text" IS '12'/);
  assert.match(result.sql, /"item_number" IS 3/);
  assert.match(result.sql, /ORDER BY revision DESC, sequence DESC LIMIT 1/);
  assert.match(result.sql, /revision FROM mirror_changes/);
  assert.match(result.sql, /operation FROM mirror_changes/);
});
