# Cloudflare → UA OneDrive / Excel reporting mirror

## Status and boundary

Implemented and remotely validated on `cloudflare-migration` with **synthetic data only**.
The reporting migration is active on the migration Worker/D1 and the UA OneDrive/Excel
archive path has been tested end to end. **Production cutover is still not authorized.**
The existing pilot's GitHub Pages site, Google Apps Script deployments, Google Sheets,
and production `main` remain untouched.

The Worker still exposes the authenticated pull API, but UA's Power Automate environment
did not provide the HTTP action without a Premium license. The validated UA workflow
therefore uses a secure archive export followed by a OneDrive file-trigger flow:

`Cloudflare reporting API → private archive export → UA OneDrive Incoming → Power Automate → Office Script archive_import → verified receipt → guarded D1 purge`.

Cloudflare receives no Microsoft permissions. The Office Script makes no network calls
and stores no Cloudflare credential. The reporting credential remains separate from
participant access codes and is used only by the export client.

Validated UA folders and workbook:

- `/AQG Dissertation/Incoming Reporting Exports`
- `/AQG Dissertation/Processed Reporting Exports`
- `/AQG Dissertation/Verified Reporting Receipts`
- `/AQG Dissertation/D1 Recovery Backups`
- `/AQG Dissertation/Audio Archives`
- `/AQG Dissertation/Verified Audio Receipts`
- `/AQG Dissertation/UA_OneDrive_Reporting_Mirror.xlsx`

The current reporting flow is intentionally single-writer (trigger concurrency 1). R2 audio is
not part of this archive/purge authorization. It now uses a separate, independently
verified archive/receipt/purge workflow so a D1 reporting receipt can never authorize
an R2 deletion.

## Captured datasets

| Dataset / worksheet | Stable source key | Report content |
| --- | --- | --- |
| `aqg_submissions` | `record_id` | Submitted settings, final responses, timing, audio references |
| `aqg_events` | `record_id` | Durable event records and sanitized structured details |
| `aqg_feedback` | `record_id` | Text feedback and audio metadata |
| `training_submissions` | `record_id` | Submission summary, version, timing, question counts |
| `training_submission_items` | `record_id` | Per-item response and response time |
| `training_events` | `record_id` | Durable events, including append and feedback paths |
| `training_feedback` | `record_id` | Text feedback and section/card context |
| `nonparticipant_button_counts` | `button_id` | Anonymous cumulative button counts |

The exact column allowlist lives in `backend/src/reporting-contract.js` and is returned
by the manifest. Authentication tables, counters, request receipts, live saves,
notification queues, and latest settings are excluded. `notification_status` is not
research data and changes to it do not advance the mirror. Raw training
`details_json` is excluded because the existing save code copies credential aliases
and recovery state into it. The normalized item and submission fields remain available.

Event `detail_json` has credential keys removed recursively, including JSON encoded
inside strings; malformed structured details become null. Participant-authored free
text is preserved. This is not a general de-identification service: authorized research
exports still contain participant IDs and submitted research content.

Audio object keys, filenames, MIME types, and durations are metadata only. This API
does **not** download private R2 audio or grant bucket access. Private recordings are
handled by a separate archive tool that reads only audio keys referenced by a reporting
archive, downloads those exact R2 objects, computes SHA-256 and byte length, and packages
them with a private manifest for transfer to UA OneDrive.

## AQG shortcut tracking

Both participant and nonparticipant pages include these fixed button IDs. No sheet
row or model-specific configuration is needed for either button.

| Button ID | Participant event | Nonparticipant measure |
| --- | --- | --- |
| `btnSkipContext` | `context_skipped` | One cumulative count per enabled press |
| `btnShowFirstQuestion` | `first_question_prompt_copied`, or `first_question_copy_failed` if automatic copying fails | One cumulative count per enabled press, including failed automatic copies |

The first-question event includes `prompt_id: "first_question"`, the prompt title,
and a `copy_succeeded` boolean in `detail_json`. Its hard-coded text is identical for
ChatGPT, Gemini, Copilot, and Other AI:

> I am done setting up questions, please show me the first question

These events use the existing revision-safe AQG save path and appear in `aqg_events`;
anonymous counts appear in `nonparticipant_button_counts`. The existing capture
triggers and Excel importer require no schema change. Disabled buttons do not count.

Skipping requires an AI-product selection and unlocks the sample, making-questions,
and final prompts without required context or a starter-prompt copy. The starter
prompt itself still needs course/topic. The separate `contextSkipped` recovery flag
does not falsely mark a starter copy. It survives regular-mode recovery and clears on
reset or a successful starter copy. Training Mode ignores/removes this flag and keeps
the skip button visibly disabled with the full-setup explanation. The first-question
button is at the end of Band 1, enabled after skipping or copying the current starter.

## Atomic capture and revisions

`0003_reporting_mirror.sql` creates the feed generation, adds `record_json` to
`mirror_changes`, backfills the eight existing datasets, and installs database
triggers. D1 applies a migration as a transaction. Successful source inserts, updates
to report fields, and explicit deletes create matching snapshots/tombstones inside
the same transaction. Failed batches and ignored duplicate inserts leave no new changes.

- A new record begins at revision 1. Each later reportable mutation increments it.
- A deletion advances the revision and emits `operation: "delete", record: null`.
- Recreating the same source key advances its revision again within that generation.
- No-op updates and changes solely to excluded operational fields do not emit changes.
- Source primary keys cannot be updated; normal inserts/updates/deletes are supported.
  Do not introduce `INSERT OR REPLACE` on reportable tables.
- Historical snapshots are returned, not a join to a record's latest value. Retrying
  a page with the same generation and `through` therefore preserves its content.
- The mirror is a reporting synchronization feed, not a separate study outcome measure.

The nonparticipant endpoint now accepts an optional anonymous `request_id`. The
migration frontend supplies one per press. Receipt insertion and count increment
are atomic; concurrent retries with the same ID count once. Reusing an ID for a
different button gives 409. Legacy requests without an ID still count once per POST.
These receipts contain no participant/session identifier and are not exported.

## Authentication and API

Create an independent, randomly generated reporting secret of at least 32 characters
(recommended: 32 random bytes represented as 64 hex characters). Configure it through
Cloudflare's secure secret prompt as `REPORTING_EXPORT_TOKEN`; store the corresponding
value in UA's approved secret facility. Never put it in Git, workbook cells, an Office
Script, a URL, a browser frontend, or chat. Participant access codes are not export keys.

Use `Authorization: Bearer <secret>` over HTTPS. No browser Origin or CORS is supported.
All reporting responses use `Cache-Control: no-store, private`. The API is disabled
with 503 when the separate secret is missing or too short. Authentication is checked
before data access; the API exposes GET operations only.

### `GET /v1/reporting/manifest`

Returns `ok`, `protocol_version: 1`, `feed_generation`, `created_at`, `latest_sequence`,
and `datasets` (name, source key, allowed columns). Save the generation as part of
the destination's synchronization identity; never use a bare sequence across generations.

### `GET /v1/reporting/changes`

| Parameter | Rule |
| --- | --- |
| `generation` | Required; the 32-character generation from the manifest |
| `after` | Nonnegative safe integer, exclusive cursor; defaults to 0 |
| `through` | Optional inclusive upper sequence; fixes the end of this catch-up run |
| `limit` | Integer 1–100; defaults to 25 |

Unknown, repeated, malformed, or unsupported parameters return 400. There is no
arbitrary table selector, SQL parameter, participant selector, or filter that could
silently skip datasets. Generation mismatch and cursors ahead of the feed return 409.

The response includes `feed_generation`, `protocol_version`, `after`, `through`,
`next_after`, `has_more`, and `changes`. Each change has `sequence`, `dataset`,
`record_id`, `revision`, `operation`, `changed_at`, and `record`.

Read from the primary and capture metadata/page in one D1 read transaction. Pages
target at most 1 MiB of snapshots, bounded in SQL before materializing the rows. A
single larger record is allowed, with a hard 4 MiB response ceiling; an oversized
record fails explicitly with 413 and does not advance the cursor. There is no truncation.

## Bootstrap and steady-state algorithm

1. Fetch the manifest; initialize a new blank workbook with it.
2. Read the workbook's `mirror_sync_state` using the script's `status` action.
3. Fix `through` to the manifest's `latest_sequence` for this run.
4. Fetch changes after the stored checkpoint, using its generation and that `through`.
5. Archive the complete response JSON in the restricted UA OneDrive archive.
6. Apply the same response through the script's `apply` action.
7. Read the returned checkpoint. Continue while `has_more`, within the run's page cap.
8. Start later runs from the persisted workbook checkpoint with a fresh manifest.

On any archive or workbook failure, do not advance a separate cursor. Retry the same
page or resume from the workbook checkpoint. The script upserts by dataset/key and
uses each row's source sequence to recover from partial workbook writes. The global
checkpoint is its final write. Concurrent writers are unsupported: use concurrency 1.

An empty page can advance across sequence gaps. Never infer the cursor from row count
or timestamps. Empty tables are normal. A workbook and archive belong to exactly one
generation; generation changes stop ingestion for an intentional rebootstrap.

## Excel importer

Paste the entire `reporting/office-script.ts` into a new script in the UA account.
Its `main` parameters are `action` and `payloadJson`, and it returns JSON text.

| Action | `payloadJson` | Result |
| --- | --- | --- |
| `status` | Empty | Initialization status, generation, protocol, last sequence |
| `initialize` | Manifest response JSON | Creates eight data tables, text-chunk table, sync-state table |
| `apply` | Change-page response JSON | Updates rows/chunks; returns the persisted checkpoint |
| `archive_import` | Complete archive-export JSON | Upserts master `tbl_<dataset>` tables, verifies persisted rows by SHA-256/readback, logs completion, and returns a verified receipt |

Each data table has the projected research fields plus `mirror_record_id`,
`mirror_revision`, `mirror_sequence`, and `mirror_deleted`. Filter out rows where
`mirror_deleted = TRUE` in active reports. Tombstones retain their key/sequence so
replayed older records cannot resurrect a deletion. Recreated records clear that flag.

Excel permits 32,767 characters in a cell. Fields over 30,000 characters are preserved
in `mirror_text_chunks`, split without breaking Unicode surrogate pairs. Join by
dataset, record ID, revision, and field, and concatenate in `part` order. The data
cell contains a readable pointer. Null and empty text display as blank in Excel;
the archived JSON preserves their distinction and the complete typed record.

Formula-like strings are escaped as literal Excel text, including in chunk rows.
The importer checks its fixed schema, cursor order, and generation before writing.
Missing/renamed reporting tables stop the import instead of silently losing prior rows.
Keep analysis, formulas, and annotations on separate worksheets.

### Validated UA Power Automate archive flow

The originally proposed scheduled HTTP pull was not used because the UA tenant's HTTP
action requires Premium licensing. The pull API remains available, but the validated
institutional workflow is file-triggered:

1. From an authenticated Codespace/admin environment, create a private archive with
   `node scripts/export-reporting-archive.mjs`. The archive pins its feed generation
   and through-sequence and contains per-record revision, SHA-256, and a random archive
   token. Keep the archive private and do not paste its contents into chat.
2. Upload the unopened archive JSON to
   `/AQG Dissertation/Incoming Reporting Exports`.
3. Power Automate's OneDrive for Business **When a file is created** trigger runs with
   concurrency set to 1, then **Get file content** reads the original archive.
4. A Compose action converts the OneDrive content with
   `base64ToString(body('Get_file_content')?['$content'])`.
5. Excel Online (Business) **Run script** calls the dissertation Office Script with
   `action = archive_import` and the Compose output as `payloadJson`.
6. The Office Script validates the archive contract, writes each master row, reads the
   persisted values back, verifies SHA-256, writes archive verification/log tables, and
   returns a receipt only after the entire archive is successfully verified.
7. OneDrive **Create file** writes that returned JSON to
   `Verified Reporting Receipts`, named
   `verified-receipt-<export_id>-<timestamp>.json`.
8. OneDrive **Move or rename a file** moves the *original trigger file identifier* to
   `Processed Reporting Exports` and uses the trigger's original **File name** as the
   destination filename. Do not use the receipt Create-file `Name` token here.

This exact filename distinction was tested after correcting an initial configuration
error that renamed the processed archive to the receipt filename.

### Validated receipt and guarded-purge behavior

A downloaded verified receipt may be supplied to the local purge tooling without
opening or pasting its contents. The default command is preview-only. Destructive mode
requires `--execute`, the exact export ID, and the exact expected deletion count.

The guarded purge:

- revalidates export/receipt generation, revision, hash, archive token, latest mirror
  revision/operation, and exact source-row values immediately before execution;
- refuses the whole purge if any non-retained target is no longer exact;
- never treats `nonparticipant_button_counts` as purgeable study rows because those
  are cumulative counters whose historical value must not be reset;
- deletes only exact verified study records, rotates the reporting feed generation,
  reseeds the new generation from current retained records, and removes the old
  generation so archive deletions do not propagate back into the UA master workbook;
- does not touch R2 audio or operational/live-session/authentication tables.

Remote synthetic validation completed two successful cycles: the first archived and
purged the existing synthetic study rows while retaining 8 cumulative counters; the
second created one new synthetic AQG event after rotation, archived it through the UA
flow, verified a 9-record receipt (1 event + 8 counters), purged only that event, and
ended with 0 study-data rows plus the same 8 counters in the new generation.

### Validated private R2 audio archive flow

Private audio uses a deliberately separate authorization chain from D1 reporting.
The validated synthetic workflow is:

1. Start from a private reporting archive and collect only nonblank
   `audio_object_key` values from `aqg_submissions` and `aqg_feedback`.
   Duplicate references to the same R2 key collapse to one archive object.
2. `node scripts/export-audio-archive.mjs --reporting-archive <file>` downloads
   each exact referenced key from the private `dissertation-study-audio` bucket.
   The exporter rejects keys outside the expected `aqg/<participant>/<session>/<file>`
   namespace.
3. The exporter computes each object's SHA-256 and byte length, creates a private
   per-object archive token, writes an internal manifest, and packages the manifest
   plus audio files into a `.tar.gz` bundle. An external manifest records the bundle
   SHA-256. No R2 object is modified or deleted by export.
4. Upload the unopened bundle and external manifest to
   `/AQG Dissertation/Audio Archives`.
5. Download the bundle back from UA OneDrive and run
   `node scripts/verify-audio-archive.mjs --manifest <manifest> --retrieved-bundle <bundle>`.
   Verification requires the returned bundle SHA-256, internal manifest, individual
   object hashes, and byte lengths to match before a separate verified audio receipt
   is issued.
6. Store that private receipt in
   `/AQG Dissertation/Verified Audio Receipts`.
7. `node scripts/execute-remote-audio-purge.mjs` defaults to preview-only. It
   validates the audio receipt, re-downloads current R2 bytes to recheck hash/size,
   and runs SELECT-only D1 checks for current references in `aqg_live_sessions`,
   `aqg_submissions`, `aqg_feedback`, and unsent `notification_outbox` payloads.
8. Destructive mode additionally requires `--execute`, the exact audio export ID,
   and the exact deletion count. The complete plan is rebuilt immediately before
   deletion; any changed bytes or current D1 reference blocks deletion.

Remote synthetic validation completed one full audio cycle: one synthetic R2 object
was exported, uploaded to UA OneDrive, downloaded back, verified byte-for-byte, then
revalidated against current R2 and D1 state and deleted by exact key. A final remote
R2 read confirmed the key no longer existed.

The private audio receipt is independent of the reporting receipt. Neither receipt is
interchangeable with the other, and a D1 archive receipt is never sufficient evidence
for deleting an R2 object.

### Original direct-pull design

The authenticated `GET /v1/reporting/manifest` and `GET /v1/reporting/changes`
endpoints and the Office Script `status`, `initialize`, and `apply` actions remain
implemented and tested. They support a future institution-approved scheduled pull
mechanism if one becomes available. They are **not** the currently validated UA
Power Automate path and must not be described as active automation.

## Restore, reseed, retention, and rollback

The initial migration backfills existing source rows automatically; clients bootstrap
by reading from `after=0`. For a restored database or deliberate feed rebuild, create
a new numbered D1 migration from:

```bash
node scripts/generate-reporting-schema.mjs --reseed
```

The command only prints SQL. Review/save it as a new migration; apply the entire file
atomically using D1 migrations after stopping the consumer. It changes the generation
and snapshots the current eight source tables. Existing changes are retained. Do not
execute its statements individually or edit/reapply migration 0003. Back up/preserve
the previous workbook, then initialize a new destination after reconciliation.

Verified archive-based source cleanup is now implemented for the eight reporting
datasets, but it is deliberately conservative. A receipt alone is insufficient:
the purge tool independently revalidates the current feed generation, latest mirror
revision/operation, source revision, exact projected source values, SHA-256, and archive
token immediately before deletion. It defaults to preview-only and requires explicit
execution confirmations.

`nonparticipant_button_counts` are retained by policy because they are cumulative
counters; deleting them would allow later lower counts to overwrite historical totals.
After a successful study-row purge, the feed is rotated/reseeded and the old generation
is removed so the UA archive remains historical rather than receiving purge tombstones.

Private R2 recording transfer and guarded cleanup are now implemented and validated
remotely with synthetic data. The audio workflow remains separate from reporting:
it requires a private audio manifest, a OneDrive round-trip verification receipt,
unchanged current R2 bytes, and zero current D1 operational references before an
object can be eligible for deletion. A D1 reporting receipt must never be used as
authorization to delete an R2 object.

Before production use, document the operating cadence/owner and the study's approved
retention schedule for archived audio. The validated tooling proves safe transfer and
exact-object cleanup behavior; it does not by itself define how long real recordings
must be retained.

To suspend reporting, stop the UA flow and remove/rotate the reporting secret. Existing
participant saves can continue capturing changes for later catch-up. Reverting the
Worker code does not require dropping the additive migration. Keep the feed and
source records for recovery; no destructive down migration is supplied.

## Verification and sources

Run `npm ci && npm run validate` from `backend/`. Tests create only disposable local
D1/R2 state, random synthetic credentials, and `TEST001` records. They call actual
Worker routes; outbound network requests are blocked. Fixture cleanup disposes the
runtime and removes its temporary data directory even when assertions fail.

See [validation-report.md](validation-report.md) and [cutover-and-rollback.md](cutover-and-rollback.md).

Primary references checked for this implementation:

- [Cloudflare D1 batch transactions](https://developers.cloudflare.com/d1/worker-api/d1-database/)
- [Cloudflare D1 migrations](https://developers.cloudflare.com/d1/reference/migrations/)
- [Office Scripts external-call restrictions](https://learn.microsoft.com/en-us/office/dev/scripts/develop/external-calls)
- [Office Scripts platform limits](https://learn.microsoft.com/en-us/office/dev/scripts/testing/platform-limits)
- [Excel cell limits](https://support.microsoft.com/en-us/excel/excel-specifications-and-limits)
- [Office Scripts range values](https://learn.microsoft.com/en-us/javascript/api/office-scripts/excelscript/excelscript.range)
- [Power Automate secure inputs/outputs](https://learn.microsoft.com/en-us/power-automate/guidance/coding-guidelines/use-secure-inputs-outputs-triggers)
