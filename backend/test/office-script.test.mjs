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

function archivedRecord(datasetName, recordId, overrides = {}, revision = 1, hashChar = "a") {
  const dataset = REPORTING_DATASETS.find(item => item.name === datasetName);
  const record = Object.fromEntries(
    dataset.columns.map(column => [
      column,
      column === dataset.key ? recordId : null
    ])
  );

  Object.assign(record, overrides);

  return {
    record_id: recordId,
    revision,
    record_sha256: hashChar.repeat(64),
    archive_token: `archive-token-synthetic-${recordId}`,
    record
  };
}

function archivePayload(
  entries,
  {
    exportId = "synthetic-archive-export-1",
    generation = "c".repeat(32),
    through = 20,
    generatedAt = "2026-09-30T18:00:00.000Z"
  } = {}
) {
  const datasets = REPORTING_DATASETS.map(dataset => {
    const records = entries
      .filter(entry => entry.dataset === dataset.name)
      .map(entry => entry.record);

    return {
      name: dataset.name,
      key: dataset.key,
      columns: dataset.columns,
      record_count: records.length,
      records
    };
  });

  return {
    format: "dissertation_reporting_archive_export",
    format_version: 1,
    protocol_version: 1,
    export_id: exportId,
    generated_at: generatedAt,
    feed_generation: generation,
    through_sequence: through,
    total_records: entries.length,
    datasets
  };
}

test("Office script: archive import writes master tables and returns verified receipt", () => {
  const workbook = new Workbook();

  const archive = archivePayload([
    {
      dataset: "aqg_feedback",
      record: archivedRecord(
        "aqg_feedback",
        "archive-feedback-1",
        {
          feedback_id: "archive-feedback-1",
          participant_id: "TEST001",
          text_feedback: "Synthetic archived feedback"
        }
      )
    },
    {
      dataset: "nonparticipant_button_counts",
      record: archivedRecord(
        "nonparticipant_button_counts",
        "archive-button-1",
        {
          press_count: 7,
          updated_at: "2026-09-30T17:00:00.000Z"
        },
        2,
        "b"
      )
    }
  ]);

  const receipt = run(workbook, "archive_import", archive);

  assert.equal(receipt.format, "dissertation_reporting_archive_receipt");
  assert.equal(receipt.export_id, archive.export_id);
  assert.equal(receipt.feed_generation, archive.feed_generation);
  assert.equal(receipt.through_sequence, 20);
  assert.equal(receipt.verified_records.length, 2);

  assert.equal(
    workbook.getTable("tbl_aqg_feedback").rows.length,
    1
  );

  assert.equal(
    workbook.getTable("tbl_nonparticipant_button_counts").rows.length,
    1
  );

  assert.equal(
    workbook.getTable("archive_verified_records").rows.length,
    2
  );

  assert.equal(
    workbook.getTable("archive_import_log").rows.length,
    1
  );

  assert.equal(
    workbook.getTable("mirror_sync_state"),
    undefined
  );
});

test("Office script: archive import preserves literal formulas and full long text", () => {
  const workbook = new Workbook();

  const longText =
    "=" +
    "x".repeat(30000) +
    "😀" +
    "y".repeat(31000);

  const archive = archivePayload([
    {
      dataset: "aqg_feedback",
      record: archivedRecord(
        "aqg_feedback",
        "archive-long",
        {
          feedback_id: "archive-long",
          participant_id: "TEST001",
          text_feedback: longText,
          audio_original_filename: "=HYPERLINK(\"https://invalid\")"
        }
      )
    }
  ]);

  const receipt = run(workbook, "archive_import", archive);

  assert.equal(receipt.verified_records.length, 1);
  assert.deepEqual(workbook.formulas, []);

  const chunks = workbook.getTable("mirror_text_chunks");
  assert.equal(
    chunks.rows.map(row => row[5]).join(""),
    longText
  );

  assert.ok(
    chunks.rows.every(row => String(row[5]).length <= 30000)
  );
});

test("Office script: interrupted archive import produces no completed log and replays safely", () => {
  const workbook = new Workbook();

  const archive = archivePayload([
    {
      dataset: "aqg_feedback",
      record: archivedRecord(
        "aqg_feedback",
        "archive-replay",
        {
          feedback_id: "archive-replay",
          participant_id: "TEST001",
          text_feedback: "Replay safety"
        }
      )
    }
  ]);

  workbook.failWrite = name => {
    if (name === "archive_import_log") {
      throw new Error("Synthetic archive completion failure");
    }
  };

  assert.throws(
    () => run(workbook, "archive_import", archive),
    /completion failure/
  );

  assert.equal(
    workbook.getTable("archive_import_log").rows.length,
    0
  );

  workbook.failWrite = null;

  const receipt = run(workbook, "archive_import", archive);

  assert.equal(receipt.verified_records.length, 1);
  assert.equal(
    workbook.getTable("tbl_aqg_feedback").rows.length,
    1
  );
  assert.equal(
    workbook.getTable("archive_verified_records").rows.length,
    1
  );
  assert.equal(
    workbook.getTable("archive_import_log").rows.length,
    1
  );
});

test("Office script: later archive can update master across feed generations without deleting older master rows", () => {
  const workbook = new Workbook();

  const first = archivePayload([
    {
      dataset: "aqg_feedback",
      record: archivedRecord(
        "aqg_feedback",
        "archive-existing",
        {
          feedback_id: "archive-existing",
          participant_id: "TEST001",
          text_feedback: "First version"
        }
      )
    },
    {
      dataset: "aqg_feedback",
      record: archivedRecord(
        "aqg_feedback",
        "archive-preserved",
        {
          feedback_id: "archive-preserved",
          participant_id: "TEST001",
          text_feedback: "Must remain in master"
        },
        1,
        "b"
      )
    }
  ]);

  run(workbook, "archive_import", first);

  const second = archivePayload(
    [
      {
        dataset: "aqg_feedback",
        record: archivedRecord(
          "aqg_feedback",
          "archive-existing",
          {
            feedback_id: "archive-existing",
            participant_id: "TEST001",
            text_feedback: "Second version"
          },
          1,
          "c"
        )
      }
    ],
    {
      exportId: "synthetic-archive-export-2",
      generation: "d".repeat(32),
      through: 4,
      generatedAt: "2026-09-30T19:00:00.000Z"
    }
  );

  run(workbook, "archive_import", second);

  const table = workbook.getTable("tbl_aqg_feedback");

  assert.equal(table.rows.length, 2);

  const idColumn = table.headers.indexOf("record_id");
  const textColumn = table.headers.indexOf("text_feedback");

  const existing = table.rows.find(
    row => row[idColumn] === "archive-existing"
  );

  const preserved = table.rows.find(
    row => row[idColumn] === "archive-preserved"
  );

  assert.equal(existing[textColumn], "Second version");
  assert.equal(preserved[textColumn], "Must remain in master");

  assert.equal(
    workbook.getTable("archive_import_log").rows.length,
    2
  );
});

test("Office script: archive import rejects an older completed checkpoint", () => {
  const workbook = new Workbook();

  const newer = archivePayload(
    [{
      dataset: "nonparticipant_button_counts",
      record: archivedRecord(
        "nonparticipant_button_counts",
        "archive-count",
        { press_count: 10, updated_at: "newer" }
      )
    }],
    {
      exportId: "newer-export",
      through: 30,
      generatedAt: "2026-09-30T20:00:00.000Z"
    }
  );

  run(workbook, "archive_import", newer);

  const older = archivePayload(
    [{
      dataset: "nonparticipant_button_counts",
      record: archivedRecord(
        "nonparticipant_button_counts",
        "archive-count",
        { press_count: 5, updated_at: "older" }
      )
    }],
    {
      exportId: "older-export",
      through: 20,
      generatedAt: "2026-09-30T21:00:00.000Z"
    }
  );

  assert.throws(
    () => run(workbook, "archive_import", older),
    /checkpoint is older/
  );

  const table =
    workbook.getTable("tbl_nonparticipant_button_counts");

  const pressCountIndex =
    table.headers.indexOf("press_count");

  assert.equal(table.rows[0][pressCountIndex], 10);
});

test("Office script: archive import converts existing header-only master worksheets into tables", () => {
  const workbook = new Workbook();

  for (const dataset of REPORTING_DATASETS) {
    const sheet = workbook.addWorksheet(dataset.name);
    sheet
      .getRangeByIndexes(0, 0, 1, dataset.columns.length)
      .setValues([dataset.columns]);
  }

  const archive = archivePayload([
    {
      dataset: "aqg_feedback",
      record: archivedRecord(
        "aqg_feedback",
        "existing-template-record",
        {
          feedback_id: "existing-template-record",
          participant_id: "TEST001",
          text_feedback: "Imported into existing workbook template"
        }
      )
    }
  ]);

  const receipt = run(workbook, "archive_import", archive);

  assert.equal(receipt.verified_records.length, 1);

  assert.ok(workbook.getTable("tbl_aqg_submissions"));
  assert.ok(workbook.getTable("tbl_aqg_events"));
  assert.ok(workbook.getTable("tbl_aqg_feedback"));
  assert.ok(workbook.getTable("tbl_training_submissions"));
  assert.ok(workbook.getTable("tbl_training_submission_items"));
  assert.ok(workbook.getTable("tbl_training_events"));
  assert.ok(workbook.getTable("tbl_training_feedback"));
  assert.ok(workbook.getTable("tbl_nonparticipant_button_counts"));

  assert.equal(
    workbook.getTable("tbl_aqg_feedback").rows.length,
    1
  );
});

test("Office script: archive import refuses an existing worksheet with incorrect headers", () => {
  const workbook = new Workbook();

  const sheet = workbook.addWorksheet("aqg_submissions");
  sheet
    .getRangeByIndexes(0, 0, 1, 1)
    .setValues([["wrong_header"]]);

  const archive = archivePayload([]);

  assert.throws(
    () => run(workbook, "archive_import", archive),
    /headers do not match/
  );
});
