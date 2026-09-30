import test from "node:test";
import assert from "node:assert/strict";
import { fixture } from "./fixture.mjs";
import { buildArchiveExport } from "../src/archive-export.js";
import {
  ARCHIVE_RECEIPT_FORMAT,
  ARCHIVE_RECEIPT_VERSION
} from "../src/archive-receipt.js";
import { buildArchivePurgePlan } from "../src/archive-purge-plan.js";
import { reseedStatements } from "../scripts/reporting-schema.mjs";

async function archiveCurrentFeed(f) {
  const manifest = (await f.report("manifest")).body;

  const response = await f.report(
    `changes?generation=${manifest.feed_generation}` +
    `&after=0&through=${manifest.latest_sequence}&limit=100`
  );

  assert.equal(response.status, 200);
  assert.equal(response.body.has_more, false);

  const archive = await buildArchiveExport({
    manifest,
    changes: response.body.changes,
    throughSequence: manifest.latest_sequence,
    generatedAt: "2026-09-30T12:00:00.000Z",
    exportId: "synthetic-purge-export",
    archiveTokenFactory: (() => {
      let n = 0;
      return () => `synthetic-token-${++n}`;
    })()
  });

  return archive;
}

function receiptFor(archive, datasetName, recordId) {
  const dataset = archive.datasets.find(
    item => item.name === datasetName
  );

  const archived = dataset.records.find(
    item => item.record_id === recordId
  );

  assert.ok(archived, "expected archived record was not found");

  return {
    format: ARCHIVE_RECEIPT_FORMAT,
    format_version: ARCHIVE_RECEIPT_VERSION,
    export_id: archive.export_id,
    feed_generation: archive.feed_generation,
    through_sequence: archive.through_sequence,
    verified_at: "2026-09-30T13:00:00.000Z",
    verified_records: [
      {
        dataset: datasetName,
        record_id: archived.record_id,
        revision: archived.revision,
        record_sha256: archived.record_sha256,
        archive_token: archived.archive_token
      }
    ]
  };
}

test("exact archived record is eligible but nothing is deleted", async t => {
  const f = await fixture();
  t.after(f.close);

  await f.db.prepare(
    `INSERT INTO nonparticipant_button_counts
     VALUES ('synthetic-purge', 4, '2026-09-30T12:00:00.000Z')`
  ).run();

  const archive = await archiveCurrentFeed(f);
  const receipt = receiptFor(
    archive,
    "nonparticipant_button_counts",
    "synthetic-purge"
  );

  const plan = await buildArchivePurgePlan({
    db: f.db,
    archiveExport: archive,
    receipt
  });

  assert.equal(plan.feed_generation_matches, true);
  assert.equal(plan.eligible_count, 1);
  assert.equal(plan.keep_count, 0);
  assert.equal(plan.records[0].status, "eligible");
  assert.equal(plan.records[0].reason, "exact_archive_match");

  const stillThere = await f.db.prepare(
    `SELECT press_count
     FROM nonparticipant_button_counts
     WHERE button_id = 'synthetic-purge'`
  ).first();

  assert.equal(stillThere.press_count, 4);
});

test("record changed after export is kept", async t => {
  const f = await fixture();
  t.after(f.close);

  await f.db.prepare(
    `INSERT INTO nonparticipant_button_counts
     VALUES ('synthetic-purge', 4, '2026-09-30T12:00:00.000Z')`
  ).run();

  const archive = await archiveCurrentFeed(f);
  const receipt = receiptFor(
    archive,
    "nonparticipant_button_counts",
    "synthetic-purge"
  );

  await f.db.prepare(
    `UPDATE nonparticipant_button_counts
     SET press_count = 5,
         updated_at = '2026-09-30T14:00:00.000Z'
     WHERE button_id = 'synthetic-purge'`
  ).run();

  const plan = await buildArchivePurgePlan({
    db: f.db,
    archiveExport: archive,
    receipt
  });

  assert.equal(plan.eligible_count, 0);
  assert.equal(plan.keep_count, 1);
  assert.equal(plan.records[0].status, "keep");
  assert.equal(plan.records[0].reason, "revision_changed");

  const stillThere = await f.db.prepare(
    `SELECT press_count
     FROM nonparticipant_button_counts
     WHERE button_id = 'synthetic-purge'`
  ).first();

  assert.equal(stillThere.press_count, 5);
});

test("old receipt becomes ineligible after feed generation rotates", async t => {
  const f = await fixture();
  t.after(f.close);

  await f.db.prepare(
    `INSERT INTO nonparticipant_button_counts
     VALUES ('synthetic-purge', 4, '2026-09-30T12:00:00.000Z')`
  ).run();

  const archive = await archiveCurrentFeed(f);
  const receipt = receiptFor(
    archive,
    "nonparticipant_button_counts",
    "synthetic-purge"
  );

  await f.db.batch(
    reseedStatements().map(statement => f.db.prepare(statement))
  );

  const plan = await buildArchivePurgePlan({
    db: f.db,
    archiveExport: archive,
    receipt
  });

  assert.equal(plan.feed_generation_matches, false);
  assert.equal(plan.eligible_count, 0);
  assert.equal(plan.keep_count, 1);
  assert.equal(
    plan.records[0].reason,
    "feed_generation_changed"
  );

  const stillThere = await f.db.prepare(
    `SELECT press_count
     FROM nonparticipant_button_counts
     WHERE button_id = 'synthetic-purge'`
  ).first();

  assert.equal(stillThere.press_count, 4);
});
