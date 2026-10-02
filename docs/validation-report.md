# Reporting migration validation

Validation updated: 2026-10-01 America/Los_Angeles (2026-10-02 UTC).
Branch: `cloudflare-migration`.
Starting migration commit: `7c0586d0ed269112985ce0f996f1dd9152921562`.
Pre-documentation validation checkpoint: `069e2311d4fe7f6eecba44490c48086058382c97`.

## Verified locally

| Check | Result / scope |
| --- | --- |
| Backend JavaScript syntax | PASS, every source/helper/test JavaScript file |
| Current frontend inline JavaScript | PASS, parsed as classic scripts or modules as declared |
| Tracked AQG/FAQ JSON | PASS |
| Current frontend Apps Script dependency scan | PASS for index, training, NP, and both help pages |
| `git diff --check` | PASS |
| Full migration diff vs current `origin/main` | PASS; current main is an ancestor |
| `npx wrangler deploy --dry-run` | PASS; no deployment performed |
| Wrangler local migration application | PASS, all three SQL migrations including trigger bodies; temporary database removed |
| Reporting / backend tests | 51 PASS, 0 failures after reporting + audio archive/purge safeguards |
| AQG browser tests | 7 PASS, 0 failures; real headless Chromium, local Worker/D1/R2 |
| AQG visual inspection | PASS, desktop setup and mobile Training Mode; optional external font CSS suppressed |
| D1/R2 testing | Actual Worker route handlers in local workerd/Miniflare, disposable database/buckets |
| Synthetic cleanup | Test runtime disposed and temporary D1/R2 directories removed |
| Workbook behavior | PASS in the model and in the actual UA Excel/Office Scripts tenant with synthetic archives |
| Remote reporting Worker / D1 | PASS with synthetic data; reporting migration active on the migration Worker/D1 without changing production GitHub Pages |
| UA OneDrive / Power Automate archive flow | PASS; file-trigger flow imports archives, creates verified receipts, and moves the original archive to Processed with its original filename |
| Guarded archive purge | PASS remotely; exact receipt/hash/revision checks, cumulative NP counters retained, feed rotated, old generation removed |
| Private R2 audio archive | PASS remotely with synthetic audio; exact referenced-object export, SHA-256 manifest, UA OneDrive round-trip verification, and separate verified audio receipt |
| Guarded R2 audio purge | PASS remotely with one synthetic object; current R2 hash/size and D1 operational references revalidated before exact-object deletion; post-delete read confirmed the key no longer existed |
| Remote notification relay | PASS with fresh synthetic AQG feedback through the deployed Worker using the intended GitHub Pages Origin header; outbox reached `sent`, attempt count 1, `sent_at` populated, `last_error` null |

## Synthetic coverage

1. Source inserts, reportable updates, ignored inserts, no-op updates, and explicit
   deletes for all eight datasets; operational/authentication data excluded.
2. Backfill of existing records, unique generation creation, reseeding, and rejection
   of the previous generation.
3. Atomic rollback of both source and mirror changes when a later batch statement fails.
4. AQG allocation, live saves, duplicated events, stale revisions, final submission,
   text feedback, and synthetic audio upload/metadata; request retries do not duplicate
   the corresponding feed records or audio objects.
5. Training live saves, saved item snapshots, append-event path, feedback-generated
   event, final submission and submission items; retry and duplicate-event behavior.
6. Nonparticipant count updates, sequential/concurrent request retries, conflicting
   request IDs, monotonic revisions and cumulative counts.
7. Separate reporting authorization, disabled-secret behavior, method/origin restrictions,
   allowed dataset boundary, no-store headers, malformed/duplicate query rejection.
8. Pagination with a fixed high-water mark while new writes continue, exact page replay,
   historical count values, tombstones, empty pages, oversized-page byte limits.
9. Nested access-code alias redaction and exclusion of raw training request JSON.
10. Office Script initialization against the actual manifest, idempotent upserts,
    out-of-order/generation rejection, partial-write recovery, and checkpoint-last writes.
11. Literal formula handling, lossless long-response chunks with Unicode boundaries,
    deletions/recreation, malformed input and missing-table rejection.
12. Actual local Worker changes imported through the Office Script into the workbook model.
13. Participant and nonparticipant pages: context skip with blank fields, all four AI
    products, identical hard-coded first-question clipboard text, all prompt bands,
    copy history/preview, direct band navigation, reload and reset.
14. Distinct skip and first-question button IDs reach the reporting feed exactly once
    per enabled press. NP counts stay anonymous; participant events retain their session.
    Failed automatic copies remain manually copyable and are counted with a failure event.
15. Training Mode rejects a saved skip flag from both local storage and server recovery;
    the disabled skip button cannot bypass the handler guard. First-question/proceed
    buttons require the current starter copy, and editing context requires recopying.
16. Server recovery of a regular skipped session, mobile layout, keyboard activation,
    and refresh of prompt/sample content after asynchronous configuration loading.
17. Private R2 audio archive contract: reporting-archive references are deduplicated,
    object keys are restricted to the expected AQG namespace, every referenced object
    must be present exactly once, and manifests/receipts bind bundle hash, object hash,
    byte length, and a separate random audio archive token.
18. Guarded audio purge planning: deletion eligibility requires an exact verified audio
    receipt, unchanged current R2 bytes, and zero current references from AQG live sessions,
    submissions, feedback, or pending notification payloads. Hash/reference mismatches
    block deletion.

The local harness used only `TEST001`, explicitly synthetic field values, and randomly
generated local-only test credentials. It did not access deployed resources. The later
remote validation also used only synthetic/test records; no real participant records
were used for the archive, receipt, purge, or post-purge checks. Synthetic audio in the
local harness was a small artificial byte string, not a human recording. The harness
blocks external Worker fetches and has no notification-relay credentials.

Browser requests to the configured deployed Worker hostname are intercepted and
dispatched to the local test Worker. No request reaches that deployment. Optional
Google Fonts CSS is replaced with empty CSS; unexpected external requests fail the
test. The browser run used Chromium 153.0.8010.0 with an explicit executable path.

## Repeat the local checks

From `backend/`:

```sh
npm ci
npx playwright install chromium
npm run validate
npm run test:ui
```

`test:ui` starts a temporary read-only static server on `127.0.0.1:8000`, creates
synthetic credentials automatically, and disposes its browser and databases afterward.
It requires that local port to be free. An existing compatible Chromium installation
can be used by setting `BROWSER_EXECUTABLE_PATH`. `UI_SCREENSHOT_DIR` optionally saves
local QA images. The browser harness is the isolated local test entry point: serving
the HTML alone retains its configured deployed Worker URL.

## Remote / UA validation completed

The reporting migration was applied to the migration D1/Worker and exercised with
synthetic data only. A UA OneDrive/Excel archive flow was then validated end to end:

- a full D1 recovery backup was created and copied to the restricted UA OneDrive backup folder;
- reporting archives were exported with a pinned generation/checkpoint and uploaded to
  `Incoming Reporting Exports`;
- Power Automate invoked the Office Script `archive_import` action, which wrote and
  read back master rows before issuing a verified receipt;
- the receipt was stored separately in `Verified Reporting Receipts`, while the
  original archive moved to `Processed Reporting Exports`;
- guarded purge previews matched the verified archives exactly;
- remote guarded purges deleted only verified study rows, retained cumulative
  `nonparticipant_button_counts`, rotated the feed generation, and removed the old
  generation without touching R2 or operational/live-session tables;
- a new synthetic AQG event was created after the first purge, appeared in the new
  reporting generation, was archived through UA, and was then safely purged through
  the same verified cycle;
- final remote state was 0 rows in the seven study-data reporting tables, 8 retained
  cumulative nonparticipant counters, and 8 matching rows in the current mirror generation.

The UA tenant did not provide the originally planned Premium HTTP action, so the
validated flow is file-triggered after a secure archive export rather than a scheduled
Power Automate HTTP pull. Production GitHub Pages were not changed.

Private R2 audio was also validated separately with synthetic data only. A synthetic
recording was written to the remote `dissertation-study-audio` bucket, downloaded
through the new exact-reference exporter, hashed, packaged with a private manifest,
uploaded to `/AQG Dissertation/Audio Archives`, downloaded back from UA OneDrive,
and independently verified byte-for-byte. A separate verified audio receipt was saved
to `/AQG Dissertation/Verified Audio Receipts`. The guarded audio purge then
revalidated the receipt, current R2 object SHA-256/byte length, and read-only D1
operational-reference counts immediately before deleting exactly that one synthetic
object. A final remote read returned "The specified key does not exist," confirming
the deletion. The D1 reporting receipt was never accepted as audio-deletion authority.

## Not yet validated / still unfinished

- Full browser-based intended-origin frontend smoke test from the migration frontend, including access, saves/resume, submission, audio, feedback, and notification behavior together.
- Production cutover and live GitHub Pages-origin smoke tests: explicitly not authorized.
- Final operational cadence/ownership and retention schedule for ongoing reporting
  exports and private audio archives after cutover.
- Full UA tenant failure/retry procedure for the ongoing operational workflow.

The local and remote synthetic tests establish the reporting/archive behavior; they do
not authorize a production release. See [reporting-mirror.md](reporting-mirror.md) for
the validated UA archive workflow and [cutover-and-rollback.md](cutover-and-rollback.md)
for the remaining release gates.
