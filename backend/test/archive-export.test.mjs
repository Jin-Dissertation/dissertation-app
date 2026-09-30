import test from "node:test";
import assert from "node:assert/strict";
import {
  buildArchiveExport,
  sha256Hex,
  stableStringify,
  ARCHIVE_EXPORT_FORMAT
} from "../src/archive-export.js";
import { REPORTING_DATASETS, REPORTING_PROTOCOL_VERSION } from "../src/reporting-contract.js";

function manifest(latestSequence = 4) {
  return {
    ok: true,
    protocol_version: REPORTING_PROTOCOL_VERSION,
    feed_generation: "0123456789abcdef0123456789abcdef",
    latest_sequence: latestSequence,
    datasets: REPORTING_DATASETS
  };
}

function fullRecord(datasetName, keyValue, overrides = {}) {
  const dataset = REPORTING_DATASETS.find(item => item.name === datasetName);
  return Object.fromEntries(
    dataset.columns.map(column => [
      column,
      column === dataset.key ? keyValue : null
    ]).concat(Object.entries(overrides))
  );
}

test("stable hashing ignores object key order", async () => {
  const a = { z: 1, a: "two", nested: { y: true, x: null } };
  const b = { nested: { x: null, y: true }, a: "two", z: 1 };

  assert.equal(stableStringify(a), stableStringify(b));
  assert.equal(await sha256Hex(a), await sha256Hex(b));
});

test("archive export reconstructs current state and excludes deleted records", async () => {
  const changes = [
    {
      sequence: 1,
      dataset: "aqg_submissions",
      record_id: "row-a",
      revision: 1,
      operation: "upsert",
      record: fullRecord("aqg_submissions", "row-a", { status: "submitted" })
    },
    {
      sequence: 2,
      dataset: "aqg_submissions",
      record_id: "row-a",
      revision: 2,
      operation: "upsert",
      record: fullRecord("aqg_submissions", "row-a", { status: "updated" })
    },
    {
      sequence: 3,
      dataset: "aqg_events",
      record_id: "row-b",
      revision: 1,
      operation: "upsert",
      record: fullRecord("aqg_events", "row-b", { event_id: "event-b" })
    },
    {
      sequence: 4,
      dataset: "aqg_events",
      record_id: "row-b",
      revision: 2,
      operation: "delete",
      record: null
    }
  ];

  let tokenNumber = 0;
  const archive = await buildArchiveExport({
    manifest: manifest(),
    changes,
    generatedAt: "2026-09-30T12:00:00.000Z",
    exportId: "synthetic-export",
    archiveTokenFactory: () => `token-${++tokenNumber}`
  });

  assert.equal(archive.format, ARCHIVE_EXPORT_FORMAT);
  assert.equal(archive.export_id, "synthetic-export");
  assert.equal(archive.through_sequence, 4);
  assert.equal(archive.total_records, 1);

  const submissions = archive.datasets.find(item => item.name === "aqg_submissions");
  assert.equal(submissions.record_count, 1);
  assert.equal(submissions.records[0].record.status, "updated");
  assert.equal(submissions.records[0].revision, 2);
  assert.equal(submissions.records[0].archive_token, "token-1");
  assert.match(submissions.records[0].record_sha256, /^[a-f0-9]{64}$/);

  const events = archive.datasets.find(item => item.name === "aqg_events");
  assert.equal(events.record_count, 0);
});

test("archive export rejects a mismatched dataset contract", async () => {
  const badManifest = manifest(0);
  badManifest.datasets = [];

  await assert.rejects(
    buildArchiveExport({ manifest: badManifest, changes: [] }),
    /dataset contract/
  );
});
