# Dissertation Site Owner Guide

This is the practical operating manual for the dissertation website. It is
written for the study owner rather than for a software developer.

## 1. What you normally need to do

### For each new participant

1. Temporarily enable the secure participant-code administration route.
2. Add the participant code using the hidden-input terminal tool.
3. Verify the expected active-code count.
4. Delete the temporary administration secret again.

### If a participant should no longer have access

Deactivate the participant code. Do **not** delete their prior study records.

### About weekly during active data collection

Review whether new study activity exists. If there is new activity:

1. Export the new D1 reporting archive.
2. Upload the unopened archive JSON to UA OneDrive.
3. Wait for Power Automate/Office Script to create the verified reporting receipt
   and move the original archive to Processed Reporting Exports.
4. Export any R2 audio referenced by that reporting archive.
5. Upload the audio archive + manifest to UA OneDrive.
6. Download the audio bundle back and run the independent verification step.
7. Store the verified audio receipt.
8. Only after verification, optionally run guarded D1/R2 purge previews and then
   explicit confirmed purges.

If there has been **no new activity**, no archive or purge action is required.

### After code changes

Run:

```bash
cd /workspaces/dissertation-app/backend
npm run validate
```

Backend changes should be deployed to the Worker before a frontend production
release. Frontend publication occurs through the repository's normal
`main`/GitHub Pages process.

## 2. What usually runs by itself

You normally do **not** need to manually manage:

- participant session IDs;
- request receipts;
- save revisions;
- local browser recovery state;
- D1 reporting mirror rows;
- notification retries;
- R2 object naming;
- Power Automate's file-trigger import once the archive is uploaded.

## 3. Add or reactivate participant codes

Never paste participant codes into chat, Git issues, screenshots, or source files.

```bash
cd /workspaces/dissertation-app/backend

export PARTICIPANT_PROVISIONING_TOKEN="$(openssl rand -hex 32)"
printf '%s' "$PARTICIPANT_PROVISIONING_TOKEN" |   npx wrangler secret put PARTICIPANT_PROVISIONING_TOKEN

node scripts/provision-access-codes.mjs
```

The script hides the participant code while it is typed and asks for it twice.

When finished:

```bash
npx wrangler secret delete PARTICIPANT_PROVISIONING_TOKEN
unset PARTICIPANT_PROVISIONING_TOKEN
npx wrangler secret list --format pretty
```

Optional disabled-route check:

```bash
curl -sS -o /dev/null -w '%{http_code}\n'   -X POST   https://dissertation-study-api.professor-jin.workers.dev/v1/admin/access-codes/provision
```

Expected while disabled: `503`.

## 4. Deactivate a participant code

Use the same temporary-token setup, then run:

```bash
node scripts/provision-access-codes.mjs --deactivate
```

This blocks future access but preserves existing study records. Then delete and
unset the temporary provisioning token exactly as above.

To reactivate later, run the normal command without `--deactivate`.

## 5. Check active-code counts without exposing codes

```bash
npx wrangler d1 execute dissertation-study-data --remote --command "SELECT COUNT(*) AS active_codes FROM access_codes WHERE active = 1;"
```

Avoid dumping `code_hash` or participant-code lists unless there is a specific
administrative reason.

## 6. Weekly reporting/archive review

The detailed contract is in `docs/reporting-mirror.md`.

```text
D1 reporting feed
   ↓
private reporting archive JSON
   ↓
UA OneDrive / Incoming Reporting Exports
   ↓
Power Automate + Office Script
   ↓
Verified Reporting Receipt
   ↓
optional guarded D1 purge

Referenced R2 audio
   ↓
private audio bundle + manifest
   ↓
UA OneDrive / Audio Archives
   ↓
download-back verification
   ↓
Verified Audio Receipt
   ↓
optional guarded R2 purge
```

### Reporting archive export

```bash
node scripts/export-reporting-archive.mjs
```

It requires the matching `REPORTING_EXPORT_TOKEN` through the secure token-file
mechanism documented in `docs/reporting-mirror.md`. Never put that token in a
URL, frontend file, workbook, Git, or chat.

Upload the resulting private JSON **unopened/unmodified** to:

`/AQG Dissertation/Incoming Reporting Exports`

A successful UA flow creates a receipt in:

`/AQG Dissertation/Verified Reporting Receipts`

and moves the original archive to:

`/AQG Dissertation/Processed Reporting Exports`

### Audio archive

```bash
node scripts/export-audio-archive.mjs --reporting-archive <reporting-archive-file>
```

Upload the resulting bundle/manifest to:

`/AQG Dissertation/Audio Archives`

After downloading the bundle back from OneDrive:

```bash
node scripts/verify-audio-archive.mjs   --manifest <manifest-file>   --retrieved-bundle <downloaded-bundle>
```

Store the resulting verified receipt in:

`/AQG Dissertation/Verified Audio Receipts`

### Purging Cloudflare copies

Purge tools default to preview-only and use explicit confirmation guards. Use the
detailed commands in `docs/reporting-mirror.md`. If anything does not match
exactly, stop rather than forcing a delete.

## 7. Notification health

```bash
npx wrangler d1 execute dissertation-study-data --remote --command "SELECT status, COUNT(*) AS n FROM notification_outbox GROUP BY status ORDER BY status;"
```

Persistent pending rows deserve investigation; do not delete them just to clear
the count.

## 8. Before changing the site

1. Work on a non-production branch.
2. Keep the change focused.
3. Run `npm run validate`.
4. Review the branch difference against current `main`.
5. If backend logic changed, deploy the reviewed Worker.
6. Before a major release, record:
   - release commit SHA;
   - current `main` SHA;
   - active Worker version;
   - fresh D1 recovery backup filename/size/SHA-256.
7. Merge through the normal release procedure.
8. Let existing GitHub Pages settings publish `main`.
9. Smoke-test the live site.

## 9. Make a D1 recovery backup

```bash
cd /workspaces/dissertation-app/backend
mkdir -p private-d1-backups

STAMP="$(date -u +%Y%m%dT%H%M%SZ)"
BACKUP="private-d1-backups/dissertation-study-data-${STAMP}.sql"

npx wrangler d1 export dissertation-study-data   --remote   --output="$BACKUP"   --skip-confirmation

wc -c "$BACKUP"
sha256sum "$BACKUP"
```

Copy the exact SQL file to:

`/AQG Dissertation/D1 Recovery Backups`

Do not paste SQL contents into chat.

## 10. Live-site smoke test after a release

Use synthetic/test data, not a real participant, to verify:

- participant access;
- training content/media;
- Save & Exit / resume;
- AQG setup and generation workflow;
- Skip contextual details;
- Show first question;
- live saves;
- final submissions;
- feedback;
- audio when relevant;
- notifications;
- reporting/archive visibility.

## 11. Secrets

| Secret | Purpose | Normal handling |
|---|---|---|
| `ACCESS_CODE_PEPPER` | HMAC participant codes | Long-lived; do not rotate casually |
| `REPORTING_EXPORT_TOKEN` | Server-to-server reporting export | Keep private; never browser-facing |
| `NOTIFICATION_RELAY_SECRET` | Authenticate notification relay | Keep private |
| `NOTIFICATION_RELAY_URL` | Notification relay destination | Keep private/configured |
| `PARTICIPANT_PROVISIONING_TOKEN` | Temporary code administration | Normally absent |

## 12. Things not to do casually

Do not:

- rotate `ACCESS_CODE_PEPPER`;
- manually delete D1 study rows;
- manually delete R2 participant audio;
- bypass archive receipt/hash checks;
- make the participant-audio bucket public;
- force-push production history;
- delete legacy rollback resources before the rollback window is closed;
- edit production `main` for experimentation;
- put access codes or secrets into source files.

## 13. If something breaks

### Site loads, but saves fail
Check the Worker health/deployment first, then browser/network errors.

### Participant code does not work
Confirm the code is active; reactivate through the secure administration tool.

### Data are in Cloudflare but not Excel
Pause cleanup. Check archive → OneDrive trigger → Office Script → verified receipt.

### Audio is in R2 but not OneDrive
Do not delete R2. Complete the separate audio archive/verification process.

### Feedback notification did not arrive
Check the notification-outbox counts. Pending notifications are designed to retry.

### A release caused a serious problem
Use the recorded release/main commit and Worker version in
`docs/cutover-and-rollback.md`.

## 14. End-of-study closeout

1. Final reporting archive.
2. Final audio archive/verification.
3. Verify UA OneDrive/Excel completeness.
4. Retain receipts/backups per approved requirements.
5. Only then perform final guarded Cloudflare cleanup.
6. Disable/remove study-only access paths when no longer needed.
7. Preserve promised public non-participant materials.

The safe rule is: **archive and verify first; delete second.**
