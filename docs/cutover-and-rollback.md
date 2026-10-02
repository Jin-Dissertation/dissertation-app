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
8. Private R2 audio archival was validated separately with a synthetic recording:
   exact referenced-object export, per-object SHA-256/byte length, private manifest,
   UA OneDrive upload/download round trip, independent byte-for-byte verification,
   and a separate verified audio receipt.
9. Guarded R2 cleanup was then validated against that synthetic object. Immediately
   before deletion, the tool revalidated the receipt, current R2 bytes, and SELECT-only
   D1 counts for live-session, submission, feedback, and pending-notification references.
   Exactly one eligible synthetic key was deleted, and a subsequent remote read confirmed
   that the key no longer existed.
10. Remote notification delivery was revalidated with fresh synthetic AQG feedback sent
    to the deployed Worker using the intended GitHub Pages Origin header. The Worker
    queued the notification and the outbox reached `sent` on the first attempt with
    a populated `sent_at` value and no error.
11. A real-Chromium AQG smoke test ran the branch-local frontend on loopback against the
    deployed migration Worker. Access/fresh session, skip-context, first-question prompt,
    quick feedback plus notification delivery, fake-microphone audio upload with remote
    R2 readback, final submission, and D1 persistence all passed.
12. The corresponding AQG reporting archive completed the UA OneDrive receipt flow and
    guarded purge: 25 verified records were reconciled, 17 study rows were deleted, and
    8 cumulative nonparticipant counters were retained. Its browser-smoke audio object
    completed the separate OneDrive audio round trip and guarded R2 deletion, followed
    by an independent absence check.
13. A real-Chromium training smoke test ran the branch-local training frontend against
    the deployed migration Worker. Access/content load, Save and Exit, server-side
    resume, feedback plus notification delivery, final submission transport, and D1
    persistence passed. Final completion was triggered programmatically after the
    resume check, so this is not a card-by-card instructional-content test.
14. The training reporting archive then completed the UA OneDrive receipt flow and
    guarded purge: 39 verified records were reconciled, 31 study rows were deleted, and
    the same 8 cumulative nonparticipant counters were retained.
15. Two older synthetic AQG audio objects referenced by the authoritative reporting
    archive were recovered from R2, archived together, uploaded/downloaded through UA
    OneDrive, verified byte-for-byte, then deleted by the guarded R2 purge after the
    single stale live-session blocker was removed. Independent reads confirmed both
    keys no longer exist.
16. Synthetic operational cleanup removed 159 TEST001 live-session/live-item/context/
    latest-settings/sent-notification rows after the reporting and audio archives were
    verified. The synthetic access code, AQG/training participant counters, and
    idempotency request receipts remain intentionally available for any final validation.

These browser smokes used a branch-local loopback origin, not the live GitHub Pages
site. Separately, the deployed Worker has been exercised with the actual GitHub Pages
Origin header. A live Pages browser check remains a post-cutover verification step.

The UA tenant does not provide the originally planned Power Automate HTTP action without
Premium licensing, so the currently validated workflow begins with a secure archive
export and OneDrive upload. The reporting pull API remains available for a future
institution-approved scheduler.

Still required before production cutover: document the ongoing operational cadence/owner
and approved retention schedule for reporting and audio archives, reconcile any new
production edits from `main`, record the final release/rollback anchors, and obtain
Taemin's separate explicit approval. The full UA tenant failure/retry simulation was
deliberately deferred after repeated successful archive/receipt/replay-safe cycles; it
remains an unexercised scenario rather than a validated result.

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
- [x] Private R2 audio export, UA OneDrive round-trip verification, separate receipt, and guarded exact-object cleanup tested with synthetic data.
- [x] Synthetic reporting rows, live/session state, sent notifications, and archived R2 audio cleaned up after verified transfer; minimal TEST001 credential/counter/idempotency fixtures remain intentionally available.
- [ ] Production participant-code provisioning is handled securely outside chat.
- [ ] All current frontend paths pass syntax/JSON/dependency checks; training media resolves.
- [x] Branch-local real-browser AQG and training smoke tests passed against the deployed migration Worker, including AQG audio and notification delivery; the actual Pages Origin header was separately validated against the Worker.
- [ ] Post-cutover live GitHub Pages browser smoke test completed from the deployed production origin.
- [x] Synthetic event/record counts reconciled across D1, archive export, UA workbook receipt, and post-purge state.
- [ ] Ongoing reporting/audio archive cadence, ownership, and approved retention schedule documented for production operation.
- [ ] Full UA tenant failure/retry simulation completed. This was deliberately deferred after repeated successful replay-safe archive cycles and is not claimed as validated.
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
