import test from "node:test";
import assert from "node:assert/strict";

import { fixture } from "./fixture.mjs";
import { buildArchiveExport } from "../src/archive-export.js";
import {
  ARCHIVE_RECEIPT_FORMAT,
  ARCHIVE_RECEIPT_VERSION
} from "../src/archive-receipt.js";
import { buildArchivePurgePlan } from "../src/archive-purge-plan.js";
import { buildGuardedArchivePurgeSql } from "../src/archive-purge-execute.js";

async function seedPurgeRows(f) {
  await f.db.prepare(
    `INSERT INTO aqg_submissions (
       record_id, participant_id, session_id, status,
       submitted_at, feedback_text, created_at
     ) VALUES (
       'synthetic-study-row', 'TEST001', 'S9001', 'submitted',
       '2026-10-01T20:00:00.000Z', 'Archived; O''Brien',
       '2026-10-01T20:00:00.000Z'
     )`
  ).run();

  await f.db.prepare(
    `INSERT INTO nonparticipant_button_counts
     (button_id, press_count, updated_at)
     VALUES ('synthetic-button', 7, '2026-10-01T20:00:00.000Z')`
  ).run();
}

async function archiveCurrentFeed(f) {
  const manifest = (await f.report("manifest")).body;

  const response = await f.report(
    `changes?generation=${manifest.feed_generation}` +
    `&after=0&through=${manifest.latest_sequence}&limit=100`
  );

  assert.equal(response.status, 200);
  assert.equal(response.body.has_more, false);

  return buildArchiveExport({
    manifest,
    changes: response.body.changes,
    throughSequence: manifest.latest_sequence,
    generatedAt: "2026-10-01T20:05:00.000Z",
    exportId: "synthetic-execution-export",
    archiveTokenFactory: (() => {
      let n = 0;
      return () => `synthetic-execution-token-${++n}`;
    })()
  });
}

function receiptForBoth(archive) {
  const verifiedRecords = [
    ["aqg_submissions", "synthetic-study-row"],
    ["nonparticipant_button_counts", "synthetic-button"]
  ].map(([datasetName, recordId]) => {
    const dataset = archive.datasets.find(
      item => item.name === datasetName
    );

    const archived = dataset.records.find(
      item => item.record_id === recordId
    );

    assert.ok(archived);

    return {
      dataset: datasetName,
      record_id: archived.record_id,
      revision: archived.revision,
      record_sha256: archived.record_sha256,
      archive_token: archived.archive_token
    };
  });

  return {
    format: ARCHIVE_RECEIPT_FORMAT,
    format_version: ARCHIVE_RECEIPT_VERSION,
    export_id: archive.export_id,
    feed_generation: archive.feed_generation,
    through_sequence: archive.through_sequence,
    verified_at: "2026-10-01T20:10:00.000Z",
    verified_records: verifiedRecords
  };
}

async function planAndBuild(f, archive, receipt) {
  const plan = await buildArchivePurgePlan({
    db: f.db,
    archiveExport: archive,
    receipt
  });

  assert.equal(plan.eligible_count, 2);
  assert.equal(plan.keep_count, 0);

  return buildGuardedArchivePurgeSql({
    plan,
    archiveExport: archive
  });
}

test("guarded purge executes atomically: deletes archived study row, retains cumulative counter, rotates feed", async t => {
  const f = await fixture();
  t.after(f.close);

  await seedPurgeRows(f);

  const archive = await archiveCurrentFeed(f);
  const receipt = receiptForBoth(archive);
  const guarded = await planAndBuild(f, archive, receipt);

  assert.equal(guarded.delete_count, 1);
  assert.equal(guarded.retained_count, 1);

  await f.db.batch(
    guarded.statements.map(statement => f.db.prepare(statement))
  );

  const studyRow = await f.db.prepare(
    "SELECT record_id FROM aqg_submissions WHERE record_id = 'synthetic-study-row'"
  ).first();

  assert.equal(studyRow, null);

  const counter = await f.db.prepare(
    "SELECT press_count FROM nonparticipant_button_counts WHERE button_id = 'synthetic-button'"
  ).first();

  assert.equal(counter.press_count, 7);

  const feed = await f.db.prepare(
    "SELECT feed_generation FROM mirror_feed_state WHERE singleton = 1"
  ).first();

  assert.notEqual(feed.feed_generation, archive.feed_generation);

  const oldRows = await f.db.prepare(
    "SELECT COUNT(*) AS n FROM mirror_changes WHERE feed_generation = ?1"
  ).bind(archive.feed_generation).first();

  assert.equal(Number(oldRows.n), 0);

  const newCounter = await f.db.prepare(
    `SELECT revision, operation
     FROM mirror_changes
     WHERE feed_generation = ?1
       AND dataset = 'nonparticipant_button_counts'
       AND record_id = 'synthetic-button'
     ORDER BY sequence DESC
     LIMIT 1`
  ).bind(feed.feed_generation).first();

  assert.equal(Number(newCounter.revision), 1);
  assert.equal(newCounter.operation, "upsert");
});

test("guarded purge rolls back the entire transaction when a target changes after planning", async t => {
  const f = await fixture();
  t.after(f.close);

  await seedPurgeRows(f);

  const archive = await archiveCurrentFeed(f);
  const receipt = receiptForBoth(archive);
  const guarded = await planAndBuild(f, archive, receipt);

  await f.db.prepare(
    `UPDATE aqg_submissions
     SET feedback_text = 'changed after planning'
     WHERE record_id = 'synthetic-study-row'`
  ).run();

  const generationBefore = await f.db.prepare(
    "SELECT feed_generation FROM mirror_feed_state WHERE singleton = 1"
  ).first();

  await assert.rejects(
    f.db.batch(
      guarded.statements.map(statement => f.db.prepare(statement))
    )
  );

  const studyRow = await f.db.prepare(
    `SELECT feedback_text
     FROM aqg_submissions
     WHERE record_id = 'synthetic-study-row'`
  ).first();

  assert.equal(studyRow.feedback_text, "changed after planning");

  const counter = await f.db.prepare(
    "SELECT press_count FROM nonparticipant_button_counts WHERE button_id = 'synthetic-button'"
  ).first();

  assert.equal(counter.press_count, 7);

  const generationAfter = await f.db.prepare(
    "SELECT feed_generation FROM mirror_feed_state WHERE singleton = 1"
  ).first();

  assert.equal(
    generationAfter.feed_generation,
    generationBefore.feed_generation
  );

  const oldRows = await f.db.prepare(
    "SELECT COUNT(*) AS n FROM mirror_changes WHERE feed_generation = ?1"
  ).bind(archive.feed_generation).first();

  assert.ok(Number(oldRows.n) > 0);

  const guardTables = await f.db.prepare(
    `SELECT COUNT(*) AS n
     FROM sqlite_schema
     WHERE type = 'table'
       AND name LIKE '__archive_purge_guard_%'`
  ).first();

  assert.equal(Number(guardTables.n), 0);
});
