# Cloudflare → UA OneDrive / Excel reporting mirror

## Status and boundary

Implemented and tested on `cloudflare-migration`. **Not deployed or connected to UA yet.**
The existing pilot's GitHub Pages site, Google Apps Script deployments, and Google
Sheets remain active. Production cutover requires Taemin's explicit approval.

The Worker exposes a pull API. A UA-controlled flow reads it, archives the returned
JSON in UA OneDrive, then calls `reporting/office-script.ts` to update Excel. Cloudflare
receives no Microsoft permissions. The script has no network calls or credentials.

Activation needs an authenticated Cloudflare administrator and confirmation that the
UA account can use a scheduled Power Automate flow with an HTTP action, OneDrive for
Business, and Excel Online (Business)'s **Run script** action. UA licensing and connector
policies have not been verified. If unavailable, keep this API contract and use an
institution-approved scheduled process; do not provision a personal Microsoft app.

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
does **not** download private R2 audio, grant bucket access, or copy recordings to
OneDrive. An approved media-copy/retention workflow remains a separate activation task.

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

### Suggested UA Power Automate flow

This is a configuration recipe, not a falsely preconnected/importable flow package.
Connection IDs, tenant policy, secret storage, workbook location, and licensing must
come from the UA account.

1. Use **Recurrence**, initially every 30 minutes. Set trigger concurrency to **1**.
   Use one flow writer for this workbook and no parallel branches that edit it.
2. Obtain the reporting secret from the UA-approved secret facility. Enable secure
   inputs/outputs on every action that handles the credential or research payload,
   including HTTP, JSON archive, and Run script actions. Restrict flow ownership.
3. Use an HTTP GET action for the manifest, passing the Authorization header. Treat
   the returned dataset contract/protocol as fixed version 1.
4. Use Excel Online (Business) **Run script**, `action = status`. To interpret its
   JSON-text return, use `json(outputs('Read_sync_state')?['body/result'])` with the
   actual action name. If uninitialized, run `initialize` with
   `string(body('Get_manifest'))` against the intended blank workbook.
5. If the workbook generation differs from the manifest, terminate without resetting
   anything. Otherwise set `after` from its checkpoint and `through` from the manifest.
6. Loop sequentially, at most **10 pages per run**, requesting `limit=25`. Construct
   the changes URL from the documented query fields. An illustrative expression is:

   ```text
   concat(variables('ApiBase'), '/v1/reporting/changes?generation=',
     variables('Generation'), '&after=', string(variables('After')),
     '&through=', string(variables('Through')), '&limit=25')
   ```

7. Archive the exact HTTP response body under a restricted folder such as
   `Dissertation/Reporting/<generation>/pages/<after>-<next_after>.json`. Use a
   deterministic name: Get file metadata by path, create only if absent, otherwise
   update the existing file with the identical replayed page. Only a confirmed 404
   should take the create branch; other storage errors stop the run.
8. Once the archive action succeeds, run `apply` with `string(body('Get_changes'))`.
   Update the flow's temporary `After` variable from the script's returned
   `last_sequence`. Stop when `has_more` is false or the page cap is reached.
9. Configure failure handling that preserves the checkpoint and reports only generic
   status/correlation metadata to an authorized operator. Do not send study responses,
   page bodies, or credentials in notifications.

Office Scripts cannot call external APIs when run by Power Automate; that is why the
HTTP action is separate. Microsoft documents a 120-second Run script timeout and
1,600 calls/user/day. The page cap and initial interval leave room for bounded catch-up;
measure actual performance with synthetic data in the UA tenant before adjusting.

### Required tenant smoke test (pending)

- Run initialization and two synthetic pages in a separate UA test workbook/archive.
- Replay a page; confirm row and button-count totals stay unchanged.
- Test a field beginning `=`, a 61,000-character response, and an emoji at a chunk
  boundary; verify literal display and full reconstruction.
- Simulate an archive failure and a Run script failure; verify the saved cursor.
- Confirm connector permissions, secure run history, and scheduled execution.
- Remove the synthetic test workbook/archive or retain only as explicitly labeled test
  fixtures. Do not substitute real participants to validate the connection.

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

No automated source-data, feed, receipt, or R2 deletion is implemented. **Do not delete
source rows merely to reclaim temporary Cloudflare storage:** explicit deletes emit
tombstones, and the archive still contains earlier snapshots. Retention, withdrawal,
archival purge, and media transfer need a coordinated policy and verified durable
copies. Resetting a generation is not a deletion/withdrawal mechanism.

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
