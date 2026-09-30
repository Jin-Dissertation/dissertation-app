import test from "node:test";
import assert from "node:assert/strict";
import {
  validateArchiveReceipt,
  ARCHIVE_RECEIPT_FORMAT,
  ARCHIVE_RECEIPT_VERSION
} from "../src/archive-receipt.js";

function archiveExport() {
  return {
    format: "dissertation_reporting_archive_export",
    format_version: 1,
    protocol_version: 1,
    export_id: "synthetic-export-1",
    generated_at: "2026-09-30T12:00:00.000Z",
    feed_generation: "0123456789abcdef0123456789abcdef",
    through_sequence: 100,
    total_records: 2,
    datasets: [
      {
        name: "aqg_submissions",
        key: "record_id",
        columns: [],
        record_count: 2,
        records: [
          {
            record_id: "row-a",
            revision: 2,
            record_sha256: "a".repeat(64),
            archive_token: "token-a",
            record: {}
          },
          {
            record_id: "row-b",
            revision: 1,
            record_sha256: "b".repeat(64),
            archive_token: "token-b",
            record: {}
          }
        ]
      }
    ]
  };
}

function receipt(records) {
  return {
    format: ARCHIVE_RECEIPT_FORMAT,
    format_version: ARCHIVE_RECEIPT_VERSION,
    export_id: "synthetic-export-1",
    feed_generation: "0123456789abcdef0123456789abcdef",
    through_sequence: 100,
    verified_at: "2026-09-30T13:00:00.000Z",
    verified_records: records
  };
}

test("receipt accepts an exact partial verification", () => {
  const result = validateArchiveReceipt({
    archiveExport: archiveExport(),
    receipt: receipt([
      {
        dataset: "aqg_submissions",
        record_id: "row-a",
        revision: 2,
        record_sha256: "a".repeat(64),
        archive_token: "token-a"
      }
    ])
  });

  assert.equal(result.verified_record_count, 1);
  assert.equal(result.verified_records[0].record_id, "row-a");
});

test("receipt rejects a hash mismatch", () => {
  assert.throws(() => validateArchiveReceipt({
    archiveExport: archiveExport(),
    receipt: receipt([
      {
        dataset: "aqg_submissions",
        record_id: "row-a",
        revision: 2,
        record_sha256: "f".repeat(64),
        archive_token: "token-a"
      }
    ])
  }), /hash does not match/);
});

test("receipt rejects a revision or token mismatch", () => {
  assert.throws(() => validateArchiveReceipt({
    archiveExport: archiveExport(),
    receipt: receipt([
      {
        dataset: "aqg_submissions",
        record_id: "row-a",
        revision: 1,
        record_sha256: "a".repeat(64),
        archive_token: "token-a"
      }
    ])
  }), /revision does not match/);

  assert.throws(() => validateArchiveReceipt({
    archiveExport: archiveExport(),
    receipt: receipt([
      {
        dataset: "aqg_submissions",
        record_id: "row-a",
        revision: 2,
        record_sha256: "a".repeat(64),
        archive_token: "wrong-token"
      }
    ])
  }), /token does not match/);
});

test("receipt rejects unknown and duplicate records", () => {
  assert.throws(() => validateArchiveReceipt({
    archiveExport: archiveExport(),
    receipt: receipt([
      {
        dataset: "aqg_submissions",
        record_id: "row-missing",
        revision: 1,
        record_sha256: "c".repeat(64),
        archive_token: "token-c"
      }
    ])
  }), /not in the archive export/);

  const duplicate = {
    dataset: "aqg_submissions",
    record_id: "row-a",
    revision: 2,
    record_sha256: "a".repeat(64),
    archive_token: "token-a"
  };

  assert.throws(() => validateArchiveReceipt({
    archiveExport: archiveExport(),
    receipt: receipt([duplicate, duplicate])
  }), /Duplicate record/);
});

test("receipt must belong to the same export checkpoint", () => {
  const wrong = receipt([]);
  wrong.through_sequence = 99;

  assert.throws(() => validateArchiveReceipt({
    archiveExport: archiveExport(),
    receipt: wrong
  }), /through sequence does not match/);
});
