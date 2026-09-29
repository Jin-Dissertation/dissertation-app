import test from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { transform } from "esbuild";
import vm from "node:vm";
import { Workbook } from "./office-workbook.mjs";
import { fixture } from "./fixture.mjs";
import { REPORTING_DATASETS } from "../src/reporting-contract.js";

const source = await readFile(new URL("../../reporting/office-script.ts", import.meta.url), "utf8");
const js = (await transform(source, { loader: "ts" })).code;
const apply = vm.runInNewContext(js + "; main;", {});
const run = (workbook, action, payload) => JSON.parse(apply(workbook, action, payload ? JSON.stringify(payload) : ""));
const manifest = { ok: true, protocol_version: 1, feed_generation: "a".repeat(32), datasets: REPORTING_DATASETS };
const spec = REPORTING_DATASETS.find(d => d.name === "aqg_feedback");
function change(sequence, id = "synthetic-feedback", feedback = "synthetic", revision = sequence) {
  const record = Object.fromEntries(spec.columns.map(column => [column, null]));
  Object.assign(record, { record_id: id, feedback_id: id, participant_id: "TEST001", text_feedback: feedback });
  return { sequence, dataset: spec.name, record_id: id, revision, operation: "upsert", changed_at: "2026-09-29", record };
}
function page(changes, after = 0) {
  const next = changes.at(-1)?.sequence ?? after;
  return { ok: true, protocol_version: 1, feed_generation: manifest.feed_generation, after, through: next, next_after: next, has_more: false, changes };
}

test("Office script: exact contract, initialization, duplicate pages, and checkpoint gap rejection", () => {
  const workbook = new Workbook();
  assert.equal(run(workbook, "status").initialized, false);
  run(workbook, "initialize", manifest);
  assert.equal(workbook.tables.size, 10);
  const first = page([change(1)]);
  assert.equal(run(workbook, "apply", first).last_sequence, 1);
  assert.equal(run(workbook, "apply", first).last_sequence, 1);
  assert.equal(workbook.getTable("report_aqg_feedback").rows.length, 1);
  assert.throws(() => run(workbook, "apply", page([change(3)], 2)), /gap/);
  assert.throws(() => run(workbook, "apply", { ...first, feed_generation: "b".repeat(32) }), /generation/);
  assert.throws(() => run(workbook, "initialize", { ...manifest, feed_generation: "b".repeat(32) }), /Generation changed/);
  assert.equal(run(workbook, "status").last_sequence, 1);
});

test("Office script: literal formulas, full long text with Unicode chunks, and deletion tombstones", () => {
  const workbook = new Workbook();
  run(workbook, "initialize", manifest);
  const long = "=" + "x".repeat(29998) + "😀" + "x".repeat(32000);
  run(workbook, "apply", page([change(1, "synthetic-large", long), change(2, "synthetic-formula", "=HYPERLINK(\"https://invalid\")")]));
  const table = workbook.getTable("report_aqg_feedback");
  assert.equal(table.rows.length, 2);
  const chunks = workbook.getTable("mirror_text_chunks");
  assert.equal(chunks.rows.map(row => row[5]).join(""), long);
  assert.ok(chunks.rows.every(row => row[5].length <= 30000));
  assert.deepEqual(workbook.formulas, []);
  run(workbook, "apply", page([{ ...change(3, "synthetic-large"), operation: "delete", record: null }], 2));
  assert.equal(chunks.rows.length, 0);
  assert.equal(table.rows[0][3], true);
  assert.equal(table.rows[0][2], 3);
  run(workbook, "apply", page([change(4, "synthetic-large", "recreated", 3)], 3));
  assert.equal(table.rows[0][3], false);
  assert.equal(table.rows.length, 2);
});

test("Office script: failure before checkpoint replays safely without duplicate rows or chunks", () => {
  const workbook = new Workbook();
  run(workbook, "initialize", manifest);
  const input = page([change(1, "synthetic-long", "a".repeat(61000)), change(2, "synthetic-other")]);
  workbook.failWrite = name => { if (name === "mirror_sync_state") throw new Error("Synthetic interrupted save"); };
  assert.throws(() => run(workbook, "apply", input), /interrupted/);
  assert.equal(run(workbook, "status").last_sequence, 0);
  assert.equal(workbook.getTable("report_aqg_feedback").rows.length, 2);
  workbook.failWrite = null;
  assert.equal(run(workbook, "apply", input).last_sequence, 2);
  assert.equal(workbook.getTable("report_aqg_feedback").rows.length, 2);
  assert.equal(workbook.getTable("mirror_text_chunks").rows.length, 3);
});

test("Office script: a failed main row after chunk writes is recoverable", () => {
  const workbook = new Workbook();
  run(workbook, "initialize", manifest);
  const input = page([change(1, "synthetic-long", "a".repeat(61000))]);
  workbook.failWrite = name => { if (name === "report_aqg_feedback") throw new Error("Synthetic row failure"); };
  assert.throws(() => run(workbook, "apply", input), /row failure/);
  assert.equal(workbook.getTable("mirror_text_chunks").rows.length, 3);
  assert.equal(run(workbook, "status").last_sequence, 0);
  workbook.failWrite = null;
  run(workbook, "apply", input);
  assert.equal(workbook.getTable("mirror_text_chunks").rows.length, 3);
  assert.equal(workbook.getTable("report_aqg_feedback").rows.length, 1);
});

test("Office script: malformed pages and missing workbook tables fail before advancing state", () => {
  const workbook = new Workbook();
  run(workbook, "initialize", manifest);
  const invalid = page([change(1), { ...change(2), dataset: "access_codes" }]);
  assert.throws(() => run(workbook, "apply", invalid), /Invalid change/);
  assert.equal(workbook.getTable("report_aqg_feedback").rows.length, 0);
  const unknownField = change(1);
  unknownField.record.access_code = "synthetic";
  assert.throws(() => run(workbook, "apply", page([unknownField])), /projection/);
  workbook.tables.delete("report_training_events");
  assert.throws(() => run(workbook, "apply", page([change(1)])), /Missing reporting table/);
  assert.equal(run(workbook, "status").last_sequence, 0);
});

test("actual local Worker feed can populate and increment the workbook importer", async t => {
  const f = await fixture();
  t.after(f.close);
  const workbook = new Workbook();
  for (let i = 0; i < 3; i++) {
    const result = await f.post("/v1/nonparticipant/button-press", { button_id: "synthetic-excel", request_id: `synthetic-excel-${i}` });
    assert.equal(result.status, 200);
  }
  const m = await f.report("manifest");
  run(workbook, "initialize", m.body);
  let state = run(workbook, "status");
  let hasMore = true;
  while (hasMore) {
    const result = await f.report(`changes?generation=${m.body.feed_generation}&after=${state.last_sequence}&limit=1`);
    assert.equal(result.status, 200);
    state = run(workbook, "apply", result.body);
    hasMore = result.body.has_more;
  }
  const rows = workbook.getTable("report_nonparticipant_button_counts").rows;
  assert.equal(rows.length, 1);
  assert.equal(rows[0][5], 3);
  assert.equal(state.last_sequence, m.body.latest_sequence);
});
