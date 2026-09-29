# Reporting migration validation

Validation date: 2026-09-29 (UTC and America/Los_Angeles).
Branch: `cloudflare-migration`.
Starting migration commit: `7c0586d0ed269112985ce0f996f1dd9152921562`.

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
| Reporting / backend tests | 13 PASS, 0 failures |
| AQG browser tests | 7 PASS, 0 failures; real headless Chromium, local Worker/D1/R2 |
| AQG visual inspection | PASS, desktop setup and mobile Training Mode; optional external font CSS suppressed |
| D1/R2 testing | Actual Worker route handlers in local workerd/Miniflare, disposable database/buckets |
| Synthetic cleanup | Test runtime disposed and temporary D1/R2 directories removed |
| Workbook behavior | Tested in a purpose-built Excel API model; actual UA tenant still pending |

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

Only `TEST001`, explicitly synthetic field values, and randomly generated local-only
test credentials were used. No deployed database, real participant records, account
credentials, or live production Pages were accessed for these tests. Synthetic audio
was a small artificial byte string, not a human recording. The harness blocks external
Worker fetches and has no notification-relay credentials.

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

## Not yet validated

- Remote Cloudflare migration application and deployment: CLI authentication is absent.
- Actual UA Power Automate license, HTTP connector policy, Excel/Office Scripts access.
- UA workbook/OneDrive permissions, secure flow history, real Excel rendering, and scheduled runs.
- Private R2 recording transfer to institutional storage and coordinated retention/purge.
- Production cutover and live-origin smoke tests: explicitly not authorized.

The local tests establish code behavior; they do not claim a completed institutional
connection or production release. See [reporting-mirror.md](reporting-mirror.md) for the
tenant smoke test and [cutover-and-rollback.md](cutover-and-rollback.md) for release gates.
