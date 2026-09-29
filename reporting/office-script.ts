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
  if (action !== "apply") throw new Error("Expected status, initialize, or apply");
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
  if (existing) { assertHeaders(existing, headers); return existing; }
  if (workbook.getWorksheet(sheetName)) throw new Error("Existing worksheet without expected table: " + sheetName);
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
  // Excel interprets formula-like strings passed to setValues/addRows. An Excel
  // apostrophe escapes them as literal text; escape a leading apostrophe too.
  return /^[\s\u0000-\u001f]*[=+\-@']/.test(value) ? "'" + value : value;
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
