import test from "node:test";
import assert from "node:assert/strict";
import {
  buildAudioArchiveCore,
  buildAudioArchiveReceipt,
  finalizeAudioArchiveManifest
} from "../src/audio-archive.js";
import { buildAudioPurgePlan } from "../src/audio-purge.js";

const AUDIO_KEY = "aqg/TEST001/S0001/request-1_feedback.webm";
const HASH = "a".repeat(64);
const BUNDLE_HASH = "b".repeat(64);

function reportingArchive() {
  return {
    format: "dissertation_reporting_archive_export",
    format_version: 1,
    protocol_version: 1,
    export_id: "reporting-export",
    generated_at: "2026-10-02T12:00:00.000Z",
    feed_generation: "0123456789abcdef0123456789abcdef",
    through_sequence: 10,
    total_records: 1,
    datasets: [
      {
        name: "aqg_feedback",
        records: [
          {
            record_id: "feedback-row",
            revision: 1,
            record: {
              audio_object_key: AUDIO_KEY,
              audio_original_filename: "feedback.webm",
              audio_mime_type: "audio/webm",
              audio_duration_seconds: 2
            }
          }
        ]
      }
    ]
  };
}

function verifiedPair() {
  const core = buildAudioArchiveCore({
    reportingArchive: reportingArchive(),
    archivedObjects: [
      {
        object_key: AUDIO_KEY,
        archive_filename: "0001.webm",
        sha256: HASH,
        byte_length: 321
      }
    ],
    audioExportId: "audio-export",
    archiveTokenFactory: () => "private-audio-token"
  });

  const manifest = finalizeAudioArchiveManifest(core, {
    bundleFilename: "audio.tar.gz",
    bundleSha256: BUNDLE_HASH
  });

  const receipt = buildAudioArchiveReceipt(manifest, {
    verifiedAt: "2026-10-02T13:00:00.000Z"
  });

  return { manifest, receipt };
}

test("audio purge planner allows an exact verified object with no live references", () => {
  const { manifest, receipt } = verifiedPair();

  const plan = buildAudioPurgePlan({
    manifest,
    receipt,
    objectStates: {
      [AUDIO_KEY]: {
        exists: true,
        sha256: HASH,
        byte_length: 321
      }
    },
    operationalReferences: {
      [AUDIO_KEY]: {
        live_sessions: 0,
        submissions: 0,
        feedback: 0,
        pending_notifications: 0
      }
    }
  });

  assert.equal(plan.verified_object_count, 1);
  assert.equal(plan.eligible_count, 1);
  assert.equal(plan.blocked_count, 0);
  assert.deepEqual(plan.eligible_objects, [AUDIO_KEY]);
});

test("audio purge planner blocks changed R2 bytes", () => {
  const { manifest, receipt } = verifiedPair();

  const plan = buildAudioPurgePlan({
    manifest,
    receipt,
    objectStates: {
      [AUDIO_KEY]: {
        exists: true,
        sha256: "c".repeat(64),
        byte_length: 321
      }
    },
    operationalReferences: {
      [AUDIO_KEY]: {}
    }
  });

  assert.equal(plan.eligible_count, 0);
  assert.deepEqual(
    plan.blocked_objects[0].keep_reasons,
    ["r2_object_hash_changed"]
  );
});

test("audio purge planner blocks current D1 and pending-notification references", () => {
  const { manifest, receipt } = verifiedPair();

  const plan = buildAudioPurgePlan({
    manifest,
    receipt,
    objectStates: {
      [AUDIO_KEY]: {
        exists: true,
        sha256: HASH,
        byte_length: 321
      }
    },
    operationalReferences: {
      [AUDIO_KEY]: {
        live_sessions: 1,
        submissions: 2,
        feedback: 3,
        pending_notifications: 4
      }
    }
  });

  assert.equal(plan.eligible_count, 0);
  assert.deepEqual(
    plan.blocked_objects[0].keep_reasons,
    [
      "referenced_by_live_session",
      "referenced_by_submission",
      "referenced_by_feedback",
      "referenced_by_pending_notification"
    ]
  );
});

test("audio purge planner rejects a receipt that does not match the manifest", () => {
  const { manifest, receipt } = verifiedPair();
  const badReceipt = structuredClone(receipt);
  badReceipt.verified_objects[0].archive_token = "wrong-token";

  assert.throws(
    () =>
      buildAudioPurgePlan({
        manifest,
        receipt: badReceipt,
        objectStates: {},
        operationalReferences: {}
      }),
    /archive token does not match/
  );
});
