/**
 * Paste into Excel > Automate > New Script in the UA account.
 * Power Automate performs HTTP and archives each page before calling apply.
 * This script performs no network calls and accepts no credentials.
 * Use one flow writer (concurrency 1) and leave reporting tables unedited.
 */
type Cell = string | number | boolean;
type ResearchValue = Cell | null;
interface DatasetSpec { name: string; key: string; columns: string[]; }
interface Manifest { ok: boolean; protocol_version: number; feed_generation: string; datasets: DatasetSpec[]; }
interface Change {
  sequence: number; dataset: string; record_id: string; revision: number;
  operation: string; changed_at: string; record: { [key: string]: ResearchValue } | null;
}
interface Page {
  ok: boolean; protocol_version: number; feed_generation: string;
  after: number; through: number; next_after: number; has_more: boolean; changes: Change[];
}
interface SyncState { initialized: boolean; feed_generation: string; last_sequence: number; protocol_version: number; }

interface ArchiveRecord {
  record_id: string;
  revision: number;
  record_sha256: string;
  archive_token: string;
  record: { [key: string]: ResearchValue };
}
interface ArchiveDataset {
  name: string;
  key: string;
  columns: string[];
  record_count: number;
  records: ArchiveRecord[];
}
interface ArchiveExport {
  format: string;
  format_version: number;
  protocol_version: number;
  export_id: string;
  generated_at: string;
  feed_generation: string;
  through_sequence: number;
  total_records: number;
  datasets: ArchiveDataset[];
}
interface ArchiveReceiptRecord {
  dataset: string;
  record_id: string;
  revision: number;
  record_sha256: string;
  archive_token: string;
}

const DATASETS: DatasetSpec[] = [
  {
    "name": "aqg_submissions",
    "key": "record_id",
    "columns": [
      "record_id",
      "participant_id",
      "session_id",
      "context_id",
      "status",
      "session_start",
      "submitted_at",
      "llm_product",
      "llm_model",
      "llm_description",
      "course_context",
      "question_context",
      "extra_instructions",
      "desired_questions",
      "final_response",
      "feedback_text",
      "audio_object_key",
      "audio_duration_seconds",
      "active_seconds",
      "session_close_type",
      "session_close_at",
      "app_version",
      "mode",
      "created_at"
    ]
  },
  {
    "name": "aqg_events",
    "key": "record_id",
    "columns": [
      "record_id",
      "event_id",
      "batch_id",
      "event_timestamp",
      "ms_since_previous",
      "participant_id",
      "session_id",
      "context_id",
      "event_type",
      "button_id",
      "detail_text",
      "detail_json",
      "device_label",
      "created_at"
    ]
  },
  {
    "name": "aqg_feedback",
    "key": "record_id",
    "columns": [
      "record_id",
      "feedback_id",
      "participant_id",
      "session_id",
      "submitted_at",
      "text_feedback",
      "audio_object_key",
      "audio_original_filename",
      "audio_mime_type",
      "audio_duration_seconds",
      "created_at"
    ]
  },
  {
    "name": "training_submissions",
    "key": "record_id",
    "columns": [
      "record_id",
      "participant_id",
      "session_id",
      "session_seq",
      "session_start",
      "session_end",
      "duration_ms",
      "duration_formatted",
      "active_seconds",
      "total_questions",
      "correct_first",
      "event_count",
      "item_count",
      "app_version",
      "content_version",
      "current_card_number",
      "device_trail",
      "submitted_at",
      "created_at"
    ]
  },
  {
    "name": "training_submission_items",
    "key": "record_id",
    "columns": [
      "record_id",
      "participant_id",
      "session_id",
      "content_version",
      "item_number",
      "response_text",
      "response_ms",
      "created_at"
    ]
  },
  {
    "name": "training_events",
    "key": "record_id",
    "columns": [
      "record_id",
      "event_id",
      "batch_id",
      "event_timestamp",
      "ms_since_previous",
      "participant_id",
      "session_id",
      "session_seq",
      "event_type",
      "section_index",
      "card_index",
      "detail_text",
      "detail_json",
      "device_label",
      "created_at"
    ]
  },
  {
    "name": "training_feedback",
    "key": "record_id",
    "columns": [
      "record_id",
      "feedback_id",
      "participant_id",
      "session_id",
      "session_seq",
      "submitted_at",
      "section_index",
      "section_title",
      "card_index",
      "feedback_source",
      "text_feedback",
      "device_label",
      "created_at"
    ]
  },
  {
    "name": "nonparticipant_button_counts",
    "key": "button_id",
    "columns": [
      "button_id",
      "press_count",
      "updated_at"
    ]
  }
];
const META = ["mirror_record_id", "mirror_revision", "mirror_sequence", "mirror_deleted"];
const STATE_HEADERS = ["feed_generation", "last_sequence", "protocol_version", "last_synced_at"];
const CHUNK_HEADERS = ["dataset", "record_id", "revision", "field", "part", "text"];
const ARCHIVE_VERIFY_HEADERS = [
  "export_id", "feed_generation", "through_sequence", "dataset",
  "record_id", "revision", "record_sha256", "archive_token", "verified_at"
];
const ARCHIVE_LOG_HEADERS = [
  "export_id", "feed_generation", "through_sequence", "generated_at",
  "total_records", "verified_records", "imported_at"
];
const ARCHIVE_EXPORT_FORMAT = "dissertation_reporting_archive_export";
const ARCHIVE_RECEIPT_FORMAT = "dissertation_reporting_archive_receipt";

function main(workbook: ExcelScript.Workbook, action: string = "status", payloadJson: string = ""): string {
  const state = readState(workbook);
  if (action === "status") return JSON.stringify(state);
  if (action === "initialize") {
    const manifest = JSON.parse(payloadJson) as Manifest;
    if (!manifest.ok || manifest.protocol_version !== 1 || !/^[a-f0-9]{32}$/.test(manifest.feed_generation) ||
        JSON.stringify(manifest.datasets) !== JSON.stringify(DATASETS)) throw new Error("Unsupported reporting manifest");
    if (state.initialized) {
      if (state.feed_generation !== manifest.feed_generation) throw new Error("Generation changed: preserve this workbook and initialize a new workbook");
      return JSON.stringify(state);
    }
    for (const dataset of DATASETS) {
      const table = ensureTable(workbook, "report_" + dataset.name, dataset.name, META.concat(dataset.columns));
      if (table.getRowCount() > 0) throw new Error("Existing report data without sync state; use a new workbook");
    }
    const chunks = ensureTable(workbook, "mirror_text_chunks", "mirror_text_chunks", CHUNK_HEADERS);
    if (chunks.getRowCount() > 0) throw new Error("Existing text chunks without sync state; use a new workbook");
    const table = ensureTable(workbook, "mirror_sync_state", "mirror_sync_state", STATE_HEADERS);
    table.addRows(-1, [[manifest.feed_generation, 0, 1, new Date().toISOString()]]);
    return JSON.stringify(readState(workbook));
  }
  if (action === "archive_import") {
    const archive = JSON.parse(payloadJson) as ArchiveExport;
    return JSON.stringify(importArchive(workbook, archive));
  }

  if (action !== "apply") throw new Error("Expected status, initialize, archive_import, or apply");
  if (!state.initialized) throw new Error("Initialize a blank workbook with the manifest first");
  const page = JSON.parse(payloadJson) as Page;
  validatePage(page, state);
  if (page.next_after <= state.last_sequence) return JSON.stringify(state);
  if (page.after !== state.last_sequence) throw new Error("Page gap or overlap: request from the workbook checkpoint");

  // Validate all workbook structures before the first data mutation. Do not
  // silently recreate a deleted sheet after the checkpoint has advanced.
  const tables = DATASETS.map(dataset => requiredTable(workbook, "report_" + dataset.name, META.concat(dataset.columns)));
  const chunks = requiredTable(workbook, "mirror_text_chunks", CHUNK_HEADERS);
  for (const change of page.changes) {
    const datasetIndex = DATASETS.findIndex(dataset => dataset.name === change.dataset);
    const dataset = DATASETS[datasetIndex];
    const table = tables[datasetIndex];
    // Read only the key/sequence columns; large text cells stay out of this read.
    const keys = columnValues(table, "mirror_record_id");
    const rowIndex = keys.findIndex(value => String(value) === change.record_id);
    if (rowIndex >= 0) {
      const sequence = Number(columnValues(table, "mirror_sequence")[rowIndex]);
      if (sequence >= change.sequence) continue; // replay after a partial workbook write
    }

    // Replace chunks before committing the row's sequence. A failure here leaves
    // that row retryable, even if an earlier chunk write partially succeeded.
    removeChunks(chunks, change.dataset, change.record_id);
    const row: Cell[] = [safeText(change.record_id), change.revision, change.sequence, change.operation === "delete"];
    for (const column of dataset.columns) {
      const value = change.record === null ? null : change.record[column];
      if (typeof value === "string" && value.length > 30000) {
        const pieces = splitText(value);
        chunks.addRows(-1, pieces.map((piece, index) => [change.dataset, safeText(change.record_id), change.revision, column, index + 1, safeText(piece)]));
        row.push(`[Long text: ${value.length} characters; see mirror_text_chunks / ${column}]`);
      } else {
        row.push(value === null ? "" : typeof value === "string" ? safeText(value) : value);
      }
    }
    if (rowIndex < 0) table.addRows(-1, [row]);
    else table.getRangeBetweenHeaderAndTotal().getRow(rowIndex).setValues([row]);
  }
  // Last write only. If anything above fails, replay from this old checkpoint.
  requiredTable(workbook, "mirror_sync_state", STATE_HEADERS).getRangeBetweenHeaderAndTotal().getRow(0)
    .setValues([[state.feed_generation, page.next_after, 1, new Date().toISOString()]]);
  return JSON.stringify(readState(workbook));
}

function importArchive(workbook: ExcelScript.Workbook, archive: ArchiveExport): object {
  validateArchiveExport(archive);

  // Archive mode can use a blank workbook or the reporting tables created by
  // the existing importer. It never advances mirror_sync_state.
  const tables = DATASETS.map(dataset =>
    ensureTable(workbook, "tbl_" + dataset.name, dataset.name, dataset.columns)
  );
  const chunks = ensureTable(workbook, "mirror_text_chunks", "mirror_text_chunks", CHUNK_HEADERS);
  const verifiedTable = ensureTable(
    workbook,
    "archive_verified_records",
    "archive_verified_records",
    ARCHIVE_VERIFY_HEADERS
  );
  const logTable = ensureTable(
    workbook,
    "archive_import_log",
    "archive_import_log",
    ARCHIVE_LOG_HEADERS
  );

  rejectOlderArchive(logTable, archive);

  const verifiedAt = new Date().toISOString();
  const receiptRecords: ArchiveReceiptRecord[] = [];

  for (let datasetIndex = 0; datasetIndex < archive.datasets.length; datasetIndex++) {
    const archiveDataset = archive.datasets[datasetIndex];
    const dataset = DATASETS[datasetIndex];
    const table = tables[datasetIndex];

    for (const archived of archiveDataset.records) {
      writeArchiveRecord(
        table,
        chunks,
        dataset,
        archived
      );

      // Critical safety check: read the values back out of the workbook after
      // writing. A receipt is issued only if the persisted master row matches
      // the archive payload exactly.
      const verifiedHash = verifyArchiveRecord(
        table,
        chunks,
        dataset,
        archived
      );

      upsertArchiveVerification(
        verifiedTable,
        archive,
        dataset.name,
        archived,
        verifiedHash,
        verifiedAt
      );

      receiptRecords.push({
        dataset: dataset.name,
        record_id: archived.record_id,
        revision: archived.revision,
        record_sha256: verifiedHash,
        archive_token: archived.archive_token
      });
    }
  }

  if (receiptRecords.length !== archive.total_records) {
    throw new Error("Archive verification count does not match export");
  }

  // Completion marker is deliberately the final workbook write. If any record
  // import or read-back fails, there is no completed archive_import_log row
  // and no successful receipt returned to the caller.
  upsertArchiveLog(logTable, archive, receiptRecords.length, verifiedAt);

  return {
    format: ARCHIVE_RECEIPT_FORMAT,
    format_version: 1,
    export_id: archive.export_id,
    feed_generation: archive.feed_generation,
    through_sequence: archive.through_sequence,
    verified_at: verifiedAt,
    verified_records: receiptRecords
  };
}

function validateArchiveExport(archive: ArchiveExport): void {
  if (!archive || archive.format !== ARCHIVE_EXPORT_FORMAT ||
      archive.format_version !== 1 || archive.protocol_version !== 1) {
    throw new Error("Unsupported archive export");
  }

  if (typeof archive.export_id !== "string" || !archive.export_id ||
      archive.export_id.length > 200) {
    throw new Error("Invalid archive export ID");
  }

  if (!/^[a-f0-9]{32}$/.test(archive.feed_generation) ||
      !Number.isSafeInteger(archive.through_sequence) ||
      archive.through_sequence < 0 ||
      !Number.isSafeInteger(archive.total_records) ||
      archive.total_records < 0 ||
      typeof archive.generated_at !== "string" ||
      Number.isNaN(Date.parse(archive.generated_at))) {
    throw new Error("Invalid archive checkpoint");
  }

  if (!Array.isArray(archive.datasets) ||
      archive.datasets.length !== DATASETS.length) {
    throw new Error("Invalid archive dataset contract");
  }

  const seen = new Set<string>();
  let total = 0;

  for (let i = 0; i < DATASETS.length; i++) {
    const expected = DATASETS[i];
    const actual = archive.datasets[i];

    if (!actual ||
        actual.name !== expected.name ||
        actual.key !== expected.key ||
        JSON.stringify(actual.columns) !== JSON.stringify(expected.columns) ||
        !Array.isArray(actual.records) ||
        !Number.isSafeInteger(actual.record_count) ||
        actual.record_count !== actual.records.length) {
      throw new Error("Archive dataset contract mismatch");
    }

    for (const archived of actual.records) {
      if (!archived ||
          typeof archived.record_id !== "string" ||
          !archived.record_id ||
          archived.record_id.length > 1000 ||
          !Number.isSafeInteger(archived.revision) ||
          archived.revision < 1 ||
          !/^[a-f0-9]{64}$/.test(archived.record_sha256) ||
          !/^[A-Za-z0-9_-]{20,200}$/.test(archived.archive_token) ||
          !archived.record ||
          archived.record[expected.key] !== archived.record_id ||
          JSON.stringify(Object.keys(archived.record).sort()) !==
            JSON.stringify(expected.columns.slice().sort())) {
        throw new Error("Invalid archive record");
      }

      const unique = actual.name + "\u0000" + archived.record_id;
      if (seen.has(unique)) throw new Error("Duplicate archive record");
      seen.add(unique);

      for (const column of expected.columns) {
        const value = archived.record[column];
        if (value !== null &&
            typeof value !== "string" &&
            typeof value !== "number" &&
            typeof value !== "boolean") {
          throw new Error("Invalid archive cell value");
        }
        if (typeof value === "number" && !Number.isFinite(value)) {
          throw new Error("Invalid archive number");
        }
      }

      total++;
    }
  }

  if (total !== archive.total_records) {
    throw new Error("Archive total record count mismatch");
  }
}

function rejectOlderArchive(table: ExcelScript.Table, archive: ArchiveExport): void {
  if (!table.getRowCount()) return;

  const rows = table.getRangeBetweenHeaderAndTotal().getValues();

  for (const row of rows) {
    if (String(row[0]) === archive.export_id) return;
  }

  let latestGenerated = "";
  let sameGenerationThrough = -1;

  for (const row of rows) {
    const generatedAt = String(row[3] || "");
    if (generatedAt > latestGenerated) latestGenerated = generatedAt;

    if (String(row[1]) === archive.feed_generation) {
      sameGenerationThrough = Math.max(
        sameGenerationThrough,
        Number(row[2])
      );
    }
  }

  if (latestGenerated && archive.generated_at < latestGenerated) {
    throw new Error("Archive is older than the latest completed import");
  }

  if (sameGenerationThrough > archive.through_sequence) {
    throw new Error("Archive checkpoint is older than this generation's completed import");
  }
}

function writeArchiveRecord(
  table: ExcelScript.Table,
  chunks: ExcelScript.Table,
  dataset: DatasetSpec,
  archived: ArchiveRecord
): void {
  const keys = columnValues(table, dataset.key);
  const rowIndex = keys.findIndex(
    value => String(value) === archived.record_id
  );

  removeChunks(chunks, dataset.name, archived.record_id);

  const row: Cell[] = [];

  for (const column of dataset.columns) {
    const value = archived.record[column];

    if (typeof value === "string" && value.length > 30000) {
      const pieces = splitText(value);

      chunks.addRows(
        -1,
        pieces.map((piece, index) => [
          dataset.name,
          safeText(archived.record_id),
          archived.revision,
          column,
          index + 1,
          safeText(piece)
        ])
      );

      row.push(
        `[Long text: ${value.length} characters; see mirror_text_chunks / ${column}]`
      );
    } else {
      row.push(
        value === null
          ? ""
          : typeof value === "string"
            ? safeText(value)
            : value
      );
    }
  }

  if (rowIndex < 0) {
    table.addRows(-1, [row]);
  } else {
    table
      .getRangeBetweenHeaderAndTotal()
      .getRow(rowIndex)
      .setValues([row]);
  }
}

function verifyArchiveRecord(
  table: ExcelScript.Table,
  chunks: ExcelScript.Table,
  dataset: DatasetSpec,
  archived: ArchiveRecord
): string {
  const keys = columnValues(table, dataset.key);
  const rowIndex = keys.findIndex(
    value => String(value) === archived.record_id
  );

  if (rowIndex < 0) {
    throw new Error("Imported archive row is missing");
  }

  const row = table
    .getRangeBetweenHeaderAndTotal()
    .getRow(rowIndex)
    .getValues()[0];

  const keyIndex = dataset.columns.indexOf(dataset.key);

  if (
    keyIndex < 0 ||
    String(row[keyIndex]) !== archived.record_id
  ) {
    throw new Error("Imported archive key did not persist correctly");
  }

  const persisted: { [key: string]: ResearchValue } = {};

  for (let i = 0; i < dataset.columns.length; i++) {
    const column = dataset.columns[i];
    const expected = archived.record[column];

    if (typeof expected === "string" && expected.length > 30000) {
      const restored = readChunkedValue(
        chunks,
        dataset.name,
        archived.record_id,
        archived.revision,
        column
      );

      if (restored !== expected) {
        throw new Error("Imported long text did not persist correctly");
      }

      persisted[column] = restored;
      continue;
    }

    const actual = row[i];

    if (expected === null) {
      if (actual !== "") {
        throw new Error("Imported null value did not persist correctly");
      }

      persisted[column] = null;
      continue;
    }

    if (
      typeof actual !== typeof expected ||
      actual !== expected
    ) {
      throw new Error("Imported archive value did not persist correctly");
    }

    persisted[column] = actual;
  }

  const computedHash = sha256HexString(
    stableRecordStringify(persisted)
  );

  if (computedHash !== archived.record_sha256) {
    throw new Error(
      "Imported archive hash does not match persisted workbook data"
    );
  }

  return computedHash;
}

function stableRecordStringify(
  record: { [key: string]: ResearchValue }
): string {
  const keys = Object.keys(record).sort();
  const pieces: string[] = [];

  for (const key of keys) {
    const encodedKey = JSON.stringify(key);
    const encodedValue = JSON.stringify(record[key]);

    if (encodedValue === undefined) {
      throw new Error("Cannot hash unsupported archive value");
    }

    pieces.push(encodedKey + ":" + encodedValue);
  }

  return "{" + pieces.join(",") + "}";
}

function utf8Bytes(value: string): number[] {
  const bytes: number[] = [];

  for (let i = 0; i < value.length; i++) {
    let code = value.charCodeAt(i);

    if (code >= 0xd800 && code <= 0xdbff) {
      if (i + 1 < value.length) {
        const low = value.charCodeAt(i + 1);

        if (low >= 0xdc00 && low <= 0xdfff) {
          code =
            0x10000 +
            ((code - 0xd800) << 10) +
            (low - 0xdc00);
          i++;
        } else {
          code = 0xfffd;
        }
      } else {
        code = 0xfffd;
      }
    } else if (code >= 0xdc00 && code <= 0xdfff) {
      code = 0xfffd;
    }

    if (code <= 0x7f) {
      bytes.push(code);
    } else if (code <= 0x7ff) {
      bytes.push(
        0xc0 | (code >>> 6),
        0x80 | (code & 0x3f)
      );
    } else if (code <= 0xffff) {
      bytes.push(
        0xe0 | (code >>> 12),
        0x80 | ((code >>> 6) & 0x3f),
        0x80 | (code & 0x3f)
      );
    } else {
      bytes.push(
        0xf0 | (code >>> 18),
        0x80 | ((code >>> 12) & 0x3f),
        0x80 | ((code >>> 6) & 0x3f),
        0x80 | (code & 0x3f)
      );
    }
  }

  return bytes;
}

function rotateRight(value: number, amount: number): number {
  return (
    (value >>> amount) |
    (value << (32 - amount))
  ) >>> 0;
}

function sha256HexString(value: string): string {
  const constants = [
    0x428a2f98, 0x71374491, 0xb5c0fbcf, 0xe9b5dba5,
    0x3956c25b, 0x59f111f1, 0x923f82a4, 0xab1c5ed5,
    0xd807aa98, 0x12835b01, 0x243185be, 0x550c7dc3,
    0x72be5d74, 0x80deb1fe, 0x9bdc06a7, 0xc19bf174,
    0xe49b69c1, 0xefbe4786, 0x0fc19dc6, 0x240ca1cc,
    0x2de92c6f, 0x4a7484aa, 0x5cb0a9dc, 0x76f988da,
    0x983e5152, 0xa831c66d, 0xb00327c8, 0xbf597fc7,
    0xc6e00bf3, 0xd5a79147, 0x06ca6351, 0x14292967,
    0x27b70a85, 0x2e1b2138, 0x4d2c6dfc, 0x53380d13,
    0x650a7354, 0x766a0abb, 0x81c2c92e, 0x92722c85,
    0xa2bfe8a1, 0xa81a664b, 0xc24b8b70, 0xc76c51a3,
    0xd192e819, 0xd6990624, 0xf40e3585, 0x106aa070,
    0x19a4c116, 0x1e376c08, 0x2748774c, 0x34b0bcb5,
    0x391c0cb3, 0x4ed8aa4a, 0x5b9cca4f, 0x682e6ff3,
    0x748f82ee, 0x78a5636f, 0x84c87814, 0x8cc70208,
    0x90befffa, 0xa4506ceb, 0xbef9a3f7, 0xc67178f2
  ];

  const bytes = utf8Bytes(value);
  const totalLength =
    Math.ceil((bytes.length + 1 + 8) / 64) * 64;

  const padded: number[] =
    new Array(totalLength).fill(0);

  for (let i = 0; i < bytes.length; i++) {
    padded[i] = bytes[i];
  }

  padded[bytes.length] = 0x80;

  const bitLength = bytes.length * 8;
  const high =
    Math.floor(bitLength / 0x100000000);
  const low = bitLength >>> 0;

  for (let i = 0; i < 4; i++) {
    padded[totalLength - 8 + i] =
      (high >>> (24 - i * 8)) & 0xff;

    padded[totalLength - 4 + i] =
      (low >>> (24 - i * 8)) & 0xff;
  }

  let h0 = 0x6a09e667;
  let h1 = 0xbb67ae85;
  let h2 = 0x3c6ef372;
  let h3 = 0xa54ff53a;
  let h4 = 0x510e527f;
  let h5 = 0x9b05688c;
  let h6 = 0x1f83d9ab;
  let h7 = 0x5be0cd19;

  const words: number[] = new Array(64).fill(0);

  for (
    let offset = 0;
    offset < totalLength;
    offset += 64
  ) {
    for (let i = 0; i < 16; i++) {
      const j = offset + i * 4;

      words[i] = (
        (padded[j] << 24) |
        (padded[j + 1] << 16) |
        (padded[j + 2] << 8) |
        padded[j + 3]
      ) >>> 0;
    }

    for (let i = 16; i < 64; i++) {
      const s0 = (
        rotateRight(words[i - 15], 7) ^
        rotateRight(words[i - 15], 18) ^
        (words[i - 15] >>> 3)
      ) >>> 0;

      const s1 = (
        rotateRight(words[i - 2], 17) ^
        rotateRight(words[i - 2], 19) ^
        (words[i - 2] >>> 10)
      ) >>> 0;

      words[i] = (
        words[i - 16] +
        s0 +
        words[i - 7] +
        s1
      ) >>> 0;
    }

    let a = h0;
    let b = h1;
    let c = h2;
    let d = h3;
    let e = h4;
    let f = h5;
    let g = h6;
    let h = h7;

    for (let i = 0; i < 64; i++) {
      const sigma1 = (
        rotateRight(e, 6) ^
        rotateRight(e, 11) ^
        rotateRight(e, 25)
      ) >>> 0;

      const choose =
        ((e & f) ^ ((~e) & g)) >>> 0;

      const temp1 = (
        h +
        sigma1 +
        choose +
        constants[i] +
        words[i]
      ) >>> 0;

      const sigma0 = (
        rotateRight(a, 2) ^
        rotateRight(a, 13) ^
        rotateRight(a, 22)
      ) >>> 0;

      const majority =
        ((a & b) ^ (a & c) ^ (b & c)) >>> 0;

      const temp2 =
        (sigma0 + majority) >>> 0;

      h = g;
      g = f;
      f = e;
      e = (d + temp1) >>> 0;
      d = c;
      c = b;
      b = a;
      a = (temp1 + temp2) >>> 0;
    }

    h0 = (h0 + a) >>> 0;
    h1 = (h1 + b) >>> 0;
    h2 = (h2 + c) >>> 0;
    h3 = (h3 + d) >>> 0;
    h4 = (h4 + e) >>> 0;
    h5 = (h5 + f) >>> 0;
    h6 = (h6 + g) >>> 0;
    h7 = (h7 + h) >>> 0;
  }

  return [
    h0, h1, h2, h3,
    h4, h5, h6, h7
  ]
    .map(word =>
      word.toString(16).padStart(8, "0")
    )
    .join("");
}

function readChunkedValue(
  table: ExcelScript.Table,
  dataset: string,
  recordId: string,
  revision: number,
  field: string
): string {
  if (!table.getRowCount()) return "";

  const rows = table.getRangeBetweenHeaderAndTotal().getValues()
    .filter(row =>
      String(row[0]) === dataset &&
      String(row[1]) === recordId &&
      Number(row[2]) === revision &&
      String(row[3]) === field
    )
    .sort((a, b) => Number(a[4]) - Number(b[4]));

  return rows.map(row => String(row[5])).join("");
}

function upsertArchiveVerification(
  table: ExcelScript.Table,
  archive: ArchiveExport,
  dataset: string,
  archived: ArchiveRecord,
  verifiedHash: string,
  verifiedAt: string
): void {
  const rows = table.getRowCount()
    ? table.getRangeBetweenHeaderAndTotal().getValues()
    : [];

  const rowIndex = rows.findIndex(row =>
    String(row[0]) === archive.export_id &&
    String(row[3]) === dataset &&
    String(row[4]) === archived.record_id
  );

  const row: Cell[] = [
    safeText(archive.export_id),
    archive.feed_generation,
    archive.through_sequence,
    dataset,
    safeText(archived.record_id),
    archived.revision,
    verifiedHash,
    archived.archive_token,
    verifiedAt
  ];

  if (rowIndex < 0) table.addRows(-1, [row]);
  else table.getRangeBetweenHeaderAndTotal().getRow(rowIndex).setValues([row]);
}

function upsertArchiveLog(
  table: ExcelScript.Table,
  archive: ArchiveExport,
  verifiedRecords: number,
  importedAt: string
): void {
  const rows = table.getRowCount()
    ? table.getRangeBetweenHeaderAndTotal().getValues()
    : [];

  const rowIndex = rows.findIndex(
    row => String(row[0]) === archive.export_id
  );

  const row: Cell[] = [
    safeText(archive.export_id),
    archive.feed_generation,
    archive.through_sequence,
    archive.generated_at,
    archive.total_records,
    verifiedRecords,
    importedAt
  ];

  if (rowIndex < 0) table.addRows(-1, [row]);
  else table.getRangeBetweenHeaderAndTotal().getRow(rowIndex).setValues([row]);
}

function readState(workbook: ExcelScript.Workbook): SyncState {
  const table = workbook.getTable("mirror_sync_state");
  if (!table || table.getRowCount() === 0) return { initialized: false, feed_generation: "", last_sequence: 0, protocol_version: 1 };
  assertHeaders(table, STATE_HEADERS);
  if (table.getRowCount() !== 1) throw new Error("Invalid sync state row count");
  const row = table.getRangeBetweenHeaderAndTotal().getValues()[0];
  if (!/^[a-f0-9]{32}$/.test(String(row[0])) || !Number.isSafeInteger(row[1]) || Number(row[1]) < 0 || row[2] !== 1) throw new Error("Invalid sync state");
  return { initialized: true, feed_generation: String(row[0]), last_sequence: Number(row[1]), protocol_version: 1 };
}

function assertHeaders(table: ExcelScript.Table, headers: string[]): void {
  const actual = table.getHeaderRowRange().getValues()[0];
  if (JSON.stringify(actual) !== JSON.stringify(headers)) throw new Error("Reporting table columns changed: " + table.getName());
}

function requiredTable(workbook: ExcelScript.Workbook, name: string, headers: string[]): ExcelScript.Table {
  const table = workbook.getTable(name);
  if (!table) throw new Error("Missing reporting table: " + name);
  assertHeaders(table, headers);
  return table;
}

function ensureTable(workbook: ExcelScript.Workbook, name: string, sheetName: string, headers: string[]): ExcelScript.Table {
  const existing = workbook.getTable(name);
  if (existing) {
    assertHeaders(existing, headers);
    return existing;
  }

  const existingSheet = workbook.getWorksheet(sheetName);

  if (existingSheet) {
    const headerRange = existingSheet.getRangeByIndexes(
      0,
      0,
      1,
      headers.length
    );

    const actualHeaders = headerRange.getValues()[0];

    if (JSON.stringify(actualHeaders) !== JSON.stringify(headers)) {
      throw new Error(
        "Existing worksheet headers do not match expected reporting columns: " +
        sheetName
      );
    }

    const table = existingSheet.addTable(headerRange, true);
    table.setName(name);
    return table;
  }

  const sheet = workbook.addWorksheet(sheetName);
  const range = sheet.getRangeByIndexes(0, 0, 1, headers.length);
  range.setValues([headers]);

  const table = sheet.addTable(range, true);
  table.setName(name);

  return table;
}

function columnValues(table: ExcelScript.Table, name: string): Cell[] {
  if (!table.getRowCount()) return [];
  const column = table.getColumnByName(name);
  if (!column) throw new Error("Missing reporting column: " + name);
  return column.getRangeBetweenHeaderAndTotal().getValues().map(row => row[0]);
}

function removeChunks(table: ExcelScript.Table, dataset: string, recordId: string): void {
  const datasetNames = columnValues(table, "dataset");
  const recordIds = columnValues(table, "record_id");
  for (let i = datasetNames.length - 1; i >= 0; i--) {
    if (datasetNames[i] === dataset && recordIds[i] === recordId) table.deleteRowsAt(i, 1);
  }
}

function safeText(value: string): string {
  // Force every non-empty archive string to be stored as literal Excel text.
  // This prevents Excel from changing values such as "12" into the number 12,
  // interpreting dates, or treating formula-like strings as formulas.
  return value === "" ? "" : "'" + value;
}

function splitText(value: string): string[] {
  const parts: string[] = [];
  for (let start = 0; start < value.length;) {
    let end = Math.min(start + 30000, value.length);
    const last = value.charCodeAt(end - 1);
    if (end < value.length && last >= 0xd800 && last <= 0xdbff) end--;
    parts.push(value.slice(start, end));
    start = end;
  }
  return parts;
}

function validatePage(page: Page, state: SyncState): void {
  if (!page.ok || page.protocol_version !== 1 || page.feed_generation !== state.feed_generation) throw new Error("Feed generation or protocol changed");
  for (const value of [page.after, page.through, page.next_after]) {
    if (!Number.isSafeInteger(value) || value < 0) throw new Error("Invalid cursor");
  }
  if (page.after > page.next_after || page.next_after > page.through || page.has_more !== (page.next_after < page.through) ||
      !Array.isArray(page.changes) || page.changes.length > 100) throw new Error("Invalid page bounds");
  let sequence = page.after;
  for (const change of page.changes) {
    const dataset = DATASETS.find(spec => spec.name === change.dataset);
    if (!dataset || !Number.isSafeInteger(change.sequence) || change.sequence <= sequence || change.sequence > page.next_after ||
        !Number.isSafeInteger(change.revision) || change.revision < 1 || typeof change.record_id !== "string" ||
        !change.record_id || change.record_id.length > 1000 || !["upsert", "delete"].includes(change.operation)) throw new Error("Invalid change");
    if (change.operation === "delete") {
      if (change.record !== null) throw new Error("Deletion payload must be null");
    } else {
      if (!change.record || change.record[dataset.key] !== change.record_id ||
          JSON.stringify(Object.keys(change.record).sort()) !== JSON.stringify(dataset.columns.slice().sort())) throw new Error("Invalid record projection");
      for (const column of dataset.columns) {
        const value = change.record[column];
        if (value !== null && typeof value !== "string" && typeof value !== "number" && typeof value !== "boolean") throw new Error("Invalid cell value");
        if (typeof value === "number" && !Number.isFinite(value)) throw new Error("Invalid number");
      }
    }
    sequence = change.sequence;
  }
  if (page.changes.length ? sequence !== page.next_after : page.next_after !== page.through) throw new Error("Page checkpoint does not cover its records");
}
