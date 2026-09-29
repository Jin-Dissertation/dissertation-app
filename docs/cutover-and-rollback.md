# Migration activation, future cutover, and rollback

**No production cutover is authorized.** The pilot is still using the existing system.
All development and reporting work remains on `cloudflare-migration`.

## Current anchors

- Repository: `Jin-Dissertation/dissertation-app`
- Existing production `main` observed before this work: `8895af36c5e3eef7306bde176d4431372f1db81a`
- Remote rollback tag: `pre-cloudflare-cutover-2026-09-27`
- Worker: `dissertation-study-api.professor-jin.workers.dev`
- D1: `dissertation-study-data`, binding `STUDY_DB`
- Private audio R2: `dissertation-study-audio`, binding `STUDY_AUDIO`
- Training media R2: `dissertation-training-media`, binding `TRAINING_MEDIA`

These are identifiers, not credentials. Recheck branch heads and the exact deployment
versions when scheduling cutover; do not assume today's main or rollback tag is the
final pre-cutover version.

## Reporting activation before production cutover

Cloudflare is unauthenticated in the current Work terminal. UA account entitlements,
workbook connections, and secret storage have not been configured. Complete these
steps through the authorized account session; never send credentials through chat.

1. Sign in to Cloudflare on the administrator's workstation using `npx wrangler login`.
2. Fetch and check out only `cloudflare-migration`. Confirm a clean tree and run
   `npm ci && npm run validate` from `backend/`.
3. Confirm the target Worker/D1 are the migration infrastructure, that no live pilot
   page has been redirected to them, and that current scheduled notification settings
   are preserved. Record a database backup and existing Worker version.
4. Apply the additive schema with
   `npx wrangler d1 migrations apply dissertation-study-data --remote`.
   Review pending migration names first. Do not reapply the first two migrations to
   an already initialized database or execute trigger fragments separately.
5. Configure `REPORTING_EXPORT_TOKEN` using
   `npx wrangler secret put REPORTING_EXPORT_TOKEN`, entering the generated secret
   only into the secure prompt. Preserve existing access/notification secrets.
6. Run `npx wrangler deploy` from this branch's `backend/`. This updates the migration
   Worker; it does not change GitHub Pages. Do not run this against a different target.
7. Configure a separate synthetic UA workbook/archive and the flow described in
   [reporting-mirror.md](reporting-mirror.md). Complete the tenant smoke tests.
8. For a remote synthetic test, use a dedicated test participant and randomized request
   IDs created securely on the administrator's machine. Exercise the existing routes,
   inspect only those synthetic IDs, and verify D1/changes/Excel/OneDrive results.
   Remove only the test rows, receipts, and test R2 objects. Deletions must be consumed
   if the test mirror is retained. Never run broad DELETE statements against remote D1.

Remote deployment and synthetic testing above are **pending**, not reported as passed.
The installed UA automation mechanism is a genuine external-account decision.

## Resources that must remain active

- All current Google Apps Script deployments and their backing Google Sheets.
- Current GitHub Pages production content from `main` throughout the pilot.
- Any existing notification relay and its configuration; the migration Worker still
  depends on its configured relay for notification delivery.
- Existing pilot audio/media stores, links, and exports until verified transfer and
  approved retention decisions.
- D1 source rows, mirror snapshots, current training content/media, and private R2 audio.
- The rollback tag and the exact final pre-cutover commit/deployment records.

The separate `v2/index.html`, `v2/training.html`, and `learning/platformwip.html` paths
retain legacy Apps Script URLs and are outside this migration task. They are not removed.

## Final readiness checklist

- [ ] Taemin confirms the pilot workflow is finished and explicitly approves cutover.
- [ ] Reporting Worker/migration and UA scheduled ingestion tested with synthetic data.
- [ ] Separate production reporting destination initialized and permissions verified.
- [ ] Required private audio transfer/retention arrangements resolved and tested.
- [ ] Remote synthetic participant data and audio cleaned up without touching study rows.
- [ ] Production participant-code provisioning is handled securely outside chat.
- [ ] All current frontend paths pass syntax/JSON/dependency checks; training media resolves.
- [ ] AQG and training access, saves, resume, final submissions, feedback, audio, and
  notification delivery work from the intended Pages origin in a synthetic smoke test.
- [ ] Both event and record counts reconcile across D1, export pages, and UA destinations.
- [ ] UA flow failures/retries/duplicate pages have been exercised in the actual tenant.
- [ ] Main/branch differences reviewed against the then-current main, with no pilot edits lost.
- [ ] Exact final pre-cutover Git commit, Pages artifact, Worker version, D1 backup, and
  rollback owner/window recorded.

## Exact future release sequence — only after explicit approval

This sequence is documentation for the future release; it has not been executed.

1. Fetch current `main` and `cloudflare-migration`, reconcile any new production edits
   into the migration branch, and repeat `npm run validate`. Review the complete diff.
2. Record the exact current production commit and a new immutable pre-cutover tag if
   `main` advanced since `pre-cloudflare-cutover-2026-09-27`. Preserve the old tag.
3. Back up D1, record the active Worker version/configuration, and confirm legacy
   Apps Script and Sheets are still available for rollback.
4. Apply pending D1 migrations, configure secrets, and deploy the tested Worker from
   the approved migration commit. Validate the migration backend before changing Pages.
5. Start/verify the UA reporting consumer with the expected generation and destination.
   Confirm the first completed synthetic cycle; then clean up the test fixtures.
6. Merge the reviewed migration branch into `main` using the repository's release
   procedure, record the resulting merge SHA, and allow the existing Pages process to
   publish it. Do not change Pages settings or delete legacy resources during this step.
7. Verify the live Pages origin using synthetic data: access, training media, recovery,
   AQG generation workflow, submissions, audio, feedback, notifications, and mirror lag.
8. Record the results and the time at which participant use of Cloudflare begins.
   Keep the legacy deployments active through the agreed rollback window.

## Rollback sequence

1. Pause new participant activity and identify whether the issue affects frontend,
   backend, notifications, or reporting only. For a reporting-only failure, pause its
   flow; keep D1 source data and the checkpoint for replay.
2. Restore the exact recorded pre-cutover Pages content via a reviewed revert of the
   migration release on `main`, then let Pages publish. If the release was a merge
   commit, the corresponding command is `git revert -m 1 <recorded-merge-sha>` after
   verifying parent 1 is the intended pre-cutover production line. Do not force-push.
3. Restore the recorded compatible Worker version only if necessary. Leave additive
   D1 schema and records intact; do not run a destructive down migration.
4. Verify the legacy Apps Script workflow with a synthetic account. Preserve all
   records created after cutover in Cloudflare and UA; reconcile them deliberately.
5. Resume the selected system only after verification and Taemin's decision. Document
   the incident, record boundaries, and a revised release plan.
