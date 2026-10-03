# Dissertation App Code Map

This is the short orientation document. For regular operating instructions, see
`SITE-OWNER-GUIDE.md`.

## The system in one picture

```text
Participant browser
(GitHub Pages)
      |
      | HTTPS
      v
Cloudflare Worker
  |        |         |
  v        v         v
 D1       R2     Notification relay
  |
  | verified archive/export
  v
UA OneDrive / Excel
```

### GitHub Pages

The participant-facing website.

- `index.html` — AQG/question-generation application.
- `training/training.html` — professional-development module.

The browser keeps some temporary state in `localStorage` so a participant can
navigate, refresh, or resume. That browser state is not the institutional
research archive.

### Cloudflare Worker

`backend/src/worker.js` is the traffic director. It receives HTTPS requests and
routes them to the appropriate module.

### D1

Cloudflare's structured operational database. It contains participant sessions,
events, submissions, feedback, training progress, authentication hashes, and the
reporting mirror machinery.

### R2

Private object storage.

- `STUDY_AUDIO` — private participant AQG audio.
- `TRAINING_MEDIA` — training-module media.

### UA OneDrive / Excel

The durable institutional reporting/archive destination. Structured data and
private audio follow separate verified archive/receipt workflows.

## Participant AQG flow

```text
Participant enters code
      ↓
auth.js validates HMAC(code)
      ↓
worker.js creates/routes session
      ↓
index.html builds prompts + logs interactions
      ↓
aqg-save.js writes resumable D1 live state
      ↓
optional audio.js → private R2
      ↓
feedback.js / final AQG submission
      ↓
D1 reporting mirror
      ↓
verified UA archive
```

## Training flow

```text
Participant enters same code
      ↓
auth.js
      ↓
training-content.js loads content/media
      ↓
training.html renders module
      ↓
training-save.js records progress/events/live state
      ↓
feedback/completion
      ↓
D1 reporting mirror
      ↓
verified UA archive
```

## Files to look at first

| File | Purpose |
|---|---|
| `index.html` | AQG participant UI and browser workflow |
| `training/training.html` | Training participant UI |
| `backend/src/worker.js` | API routing |
| `backend/src/auth.js` | Participant-code authentication |
| `backend/src/aqg-save.js` | AQG live saves/recovery/submission |
| `backend/src/training-save.js` | Training live saves/events/completion |
| `backend/src/audio.js` | Private AQG audio upload |
| `backend/src/feedback.js` | Feedback + notification-outbox creation |
| `backend/src/notifications.js` | Notification delivery/retry |
| `backend/src/reporting.js` | Read-only reporting feed |
| `backend/src/provisioning.js` | Temporarily enabled code administration |
| `backend/scripts/provision-access-codes.mjs` | Hidden-input code admin tool |
| `docs/reporting-mirror.md` | Detailed archive/reporting design |
| `docs/cutover-and-rollback.md` | Release/rollback procedure |

## Why request receipts and revisions exist

Participant networks are imperfect. A browser can refresh, retry, or send the
same request twice. The backend therefore uses:

- `request_id` + request receipts to make retries idempotent;
- `revision` + `base_revision` to reject stale overwrites;
- live-session rows for Save & Exit / recovery;
- verified archive receipts before destructive Cloudflare cleanup.

Do not remove those protections just because the code looks repetitive.

## Participant-code model

There is one participant code per participant.

```text
participant code → deidentified participant identifier
participant code → HMAC hash for authentication lookup
```

D1 authentication does not need a plaintext-code column. The long-lived
`ACCESS_CODE_PEPPER` secret must not be changed casually because existing code
hashes depend on it.

## Things not to change casually

- `ACCESS_CODE_PEPPER`
- reporting/notification secrets
- request receipt + revision logic
- verified archive/purge checks
- R2 privacy
- GitHub Pages settings during a rollback window
- legacy Apps Script/Sheets resources until migration rollback is no longer needed
