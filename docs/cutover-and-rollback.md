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

## Reporting/archive status before production cutover

The migration Worker/D1 reporting path and UA archive/receipt workflow have now been
validated remotely with synthetic data while production GitHub Pages remained on the
existing system. This is **not** a production cutover.

Completed synthetic validation:

1. Migration `0003_reporting_mirror.sql` is present remotely and the authenticated
   reporting API is reachable on the migration Worker.
2. A full D1 recovery backup was created before destructive testing and copied to the
   restricted UA OneDrive `D1 Recovery Backups` folder.
3. Private archive exports were imported into
   `UA_OneDrive_Reporting_Mirror.xlsx` through a file-trigger Power Automate flow.
4. The Office Script `archive_import` path issued verified receipts only after row
   write/readback and SHA-256 verification.
5. The flow stores receipts in `Verified Reporting Receipts` and moves original
   archives to `Processed Reporting Exports` under their original filenames.
6. Guarded D1 purge previews and executions were validated remotely. Exact archived
   study rows were deleted only after receipt/current-state revalidation; cumulative
   nonparticipant counters were retained; the feed rotated; old-generation mirror rows
   were removed; operational/live-session tables and R2 were untouched.
7. A post-purge synthetic AQG event successfully entered the new generation, completed
   the same UA archive/receipt cycle, and was safely purged. Final D1 reporting state
   returned to 0 study rows plus 8 retained cumulative counters.

The UA tenant does not provide the originally planned Power Automate HTTP action without
Premium licensing, so the currently validated workflow begins with a secure archive
export and OneDrive upload. The reporting pull API remains available for a future
institution-approved scheduler.

Still required before production cutover: resolve private R2 audio transfer/retention,
decide the ongoing operational cadence/owner for archive exports, complete the final
intended-origin frontend smoke test (including notifications and audio), reconcile any
new production edits, and obtain Taemin's separate explicit approval.

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
- [x] Reporting Worker/migration and UA archive/receipt ingestion tested with synthetic data.
- [x] UA OneDrive/Excel reporting destination and file-trigger archive flow validated with synthetic data.
- [ ] Required private audio transfer/retention arrangements resolved and tested.
- [ ] Remote synthetic participant data and audio cleaned up without touching study rows.
- [ ] Production participant-code provisioning is handled securely outside chat.
- [ ] All current frontend paths pass syntax/JSON/dependency checks; training media resolves.
- [ ] AQG and training access, saves, resume, final submissions, feedback, audio, and
  notification delivery work from the intended Pages origin in a synthetic smoke test.
- [x] Synthetic event/record counts reconciled across D1, archive export, UA workbook receipt, and post-purge state.
- [ ] UA failure/retry behavior has been exercised sufficiently for the final production procedure (the validated archive path is replay-safe in tests, but a full tenant failure simulation remains pending).
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
