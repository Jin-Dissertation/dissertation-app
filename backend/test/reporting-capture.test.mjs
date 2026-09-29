import test from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { fixture } from "./fixture.mjs";
import { REPORTING_DATASETS } from "../src/reporting-contract.js";
import { migrationSql, reseedStatements } from "../scripts/reporting-schema.mjs";

test("checked-in SQL exactly matches the reporting contract", async () => {
  assert.equal(await readFile(new URL("../migrations/0003_reporting_mirror.sql", import.meta.url), "utf8"), migrationSql());
});

test("atomic capture: bootstrap, ignored inserts, no-op updates, revisions, deletes, rollback and reseed", async t => {
  const f = await fixture({ beforeMirror: async db => {
    await db.prepare("INSERT INTO nonparticipant_button_counts VALUES ('synthetic-before', 4, '2026-09-29')").run();
  } });
  t.after(f.close);
  const firstGeneration = (await f.report("manifest")).body.feed_generation;
  assert.equal(await f.count(), 1);
  const bootstrap = await f.db.prepare("SELECT * FROM mirror_changes").first();
  assert.equal(JSON.parse(bootstrap.record_json).press_count, 4);
  assert.equal(bootstrap.revision, 1);
  await f.db.prepare("INSERT OR IGNORE INTO nonparticipant_button_counts VALUES ('synthetic-before', 99, 'later')").run();
  assert.equal(await f.count(), 1);
  await f.db.prepare("UPDATE nonparticipant_button_counts SET press_count = press_count").run();
  assert.equal(await f.count(), 1);
  await f.db.prepare("UPDATE nonparticipant_button_counts SET press_count = 5").run();
  assert.equal(await f.count(), 2);
  await assert.rejects(f.db.batch([
    f.db.prepare("UPDATE nonparticipant_button_counts SET press_count = 6"),
    f.db.prepare("INSERT INTO nonparticipant_button_counts VALUES ('synthetic-before', 1, 'duplicate')")
  ]));
  assert.equal(await f.count(), 2, "mirror and source roll back together");
  assert.equal((await f.db.prepare("SELECT press_count FROM nonparticipant_button_counts").first()).press_count, 5);
  await assert.rejects(f.db.prepare("UPDATE nonparticipant_button_counts SET button_id = 'changed-key'").run());
  await f.db.prepare("DELETE FROM nonparticipant_button_counts").run();
  const deletion = await f.db.prepare("SELECT * FROM mirror_changes ORDER BY sequence DESC LIMIT 1").first();
  assert.equal(deletion.operation, "delete");
  assert.equal(deletion.record_json, null);
  assert.equal(deletion.revision, 3);
  await f.db.prepare("INSERT INTO nonparticipant_button_counts VALUES ('synthetic-before', 1, 'recreated')").run();
  assert.equal((await f.db.prepare("SELECT MAX(revision) AS r FROM mirror_changes").first()).r, 4);
  await f.db.batch(reseedStatements().map(s => f.db.prepare(s)));
  const manifest = (await f.report("manifest")).body;
  assert.notEqual(manifest.feed_generation, firstGeneration);
  const rows = (await f.report(`changes?generation=${manifest.feed_generation}&after=0`)).body.changes;
  assert.equal(rows.length, 1);
  assert.equal(rows[0].revision, 1);
  assert.equal(rows[0].record.press_count, 1);
  assert.equal((await f.report(`changes?generation=${firstGeneration}&after=0`)).status, 409);
});

test("all eight projections capture inserts, research edits, and deletions; operational fields excluded", async t => {
  const f = await fixture();
  t.after(f.close);
  for (const dataset of REPORTING_DATASETS) {
    const info = (await f.db.prepare(`PRAGMA table_info(${dataset.name})`).all()).results;
    const required = info.filter(c => c.pk || (c.notnull && c.dflt_value === null));
    const values = required.map(c => c.type === "INTEGER" ? 1 : `TEST-${dataset.name}-${c.name}`);
    const insert = `INSERT INTO ${dataset.name} (${required.map(c => c.name).join(",")}) VALUES (${values.map(() => "?").join(",")})`;
    await f.db.prepare(insert).bind(...values).run();
    const row = await f.db.prepare("SELECT * FROM mirror_changes ORDER BY sequence DESC LIMIT 1").first();
    assert.equal(row.dataset, dataset.name);
    assert.equal(row.revision, 1);
    assert.deepEqual(Object.keys(JSON.parse(row.record_json)), dataset.columns);
    const count = await f.count();
    if (info.some(c => c.name === "notification_status")) {
      await f.db.prepare(`UPDATE ${dataset.name} SET notification_status = 'sent'`).run();
      assert.equal(await f.count(), count, "notification delivery is operational");
    }
    if (dataset.name === "training_submissions") {
      await f.db.prepare("UPDATE training_submissions SET details_json = ?").bind(JSON.stringify({ code: f.code })).run();
      assert.equal(await f.count(), count, "raw request JSON is never mirrored");
    }
    const column = dataset.columns.find(c => c !== dataset.key && !required.some(r => r.name === c)) || "updated_at";
    await f.db.prepare(`UPDATE ${dataset.name} SET ${column} = ?`).bind("synthetic-edit").run();
    assert.equal(await f.count(), count + 1);
    await f.db.prepare(`DELETE FROM ${dataset.name}`).run();
    const last = await f.db.prepare("SELECT operation, revision FROM mirror_changes ORDER BY sequence DESC LIMIT 1").first();
    assert.deepEqual(last, { operation: "delete", revision: 3 });
  }
});
