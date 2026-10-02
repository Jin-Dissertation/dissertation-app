import test from "node:test";
import assert from "node:assert/strict";
import {
  AUDIO_ARCHIVE_FORMAT,
  AUDIO_ARCHIVE_RECEIPT_FORMAT,
  audioArchiveCoreFromManifest,
  buildAudioArchiveCore,
  buildAudioArchiveReceipt,
  collectAudioReferences,
  finalizeAudioArchiveManifest,
  manifestsMatchCore,
  validateAudioArchiveReceipt
} from "../src/audio-archive.js";

const HASH_A = "a".repeat(64);
const HASH_B = "b".repeat(64);

function reportingArchive({ unsafe = false } = {}) {
  const objectKey = unsafe
    ? "../unexpected/audio.webm"
    : "aqg/TEST001/S0001/request-1_feedback.webm";

  return {
    format: "dissertation_reporting_archive_export",
    format_version: 1,
    protocol_version: 1,
    export_id: "reporting-export-1",
    generated_at: "2026-10-02T12:00:00.000Z",
    feed_generation: "0123456789abcdef0123456789abcdef",
    through_sequence: 42,
    total_records: 2,
    datasets: [
      {
        name: "aqg_submissions",
        records: [
          {
            record_id: "submission-row",
            revision: 2,
            record: {
              audio_object_key: objectKey,
              audio_duration_seconds: 12.5
            }
          }
        ]
      },
      {
        name: "aqg_feedback",
        records: [
          {
            record_id: "feedback-row",
            revision: 1,
            record: {
              audio_object_key: objectKey,
              audio_original_filename: "feedback.webm",
              audio_mime_type: "audio/webm",
              audio_duration_seconds: 12.5
            }
          }
        ]
      },
      {
        name: "aqg_events",
        records: []
      }
    ]
  };
}

test("audio references deduplicate one R2 object and preserve reporting references", () => {
  const references = collectAudioReferences(reportingArchive());

  assert.equal(references.length, 1);
  assert.equal(
    references[0].object_key,
    "aqg/TEST001/S0001/request-1_feedback.webm"
  );
  assert.equal(references[0].audio_original_filename, "feedback.webm");
  assert.equal(references[0].audio_mime_type, "audio/webm");
  assert.equal(references[0].audio_duration_seconds, 12.5);
  assert.deepEqual(
    references[0].references.map(item => item.dataset),
    ["aqg_feedback", "aqg_submissions"]
  );
});

test("audio archive rejects object keys outside the application AQG namespace", () => {
  assert.throws(
    () => collectAudioReferences(reportingArchive({ unsafe: true })),
    /Unsafe or unexpected R2 audio object key/
  );
});

test("audio archive core requires every referenced object exactly once", () => {
  assert.throws(
    () =>
      buildAudioArchiveCore({
        reportingArchive: reportingArchive(),
        archivedObjects: [],
        audioExportId: "audio-export-1",
        archiveTokenFactory: () => "token-1"
      }),
    /count does not match/
  );

  assert.throws(
    () =>
      buildAudioArchiveCore({
        reportingArchive: reportingArchive(),
        archivedObjects: [
          {
            object_key: "aqg/TEST001/S0001/request-1_feedback.webm",
            archive_filename: "0001.webm",
            sha256: HASH_A,
            byte_length: 123
          },
          {
            object_key: "aqg/TEST001/S0001/request-1_feedback.webm",
            archive_filename: "0002.webm",
            sha256: HASH_A,
            byte_length: 123
          }
        ],
        audioExportId: "audio-export-1",
        archiveTokenFactory: () => "token-1"
      }),
    /Duplicate downloaded audio object/
  );
});

test("audio manifest and receipt bind package, object hash, size, and archive token", () => {
  const core = buildAudioArchiveCore({
    reportingArchive: reportingArchive(),
    archivedObjects: [
      {
        object_key: "aqg/TEST001/S0001/request-1_feedback.webm",
        archive_filename: "0001.webm",
        sha256: HASH_A,
        byte_length: 123
      }
    ],
    generatedAt: "2026-10-02T12:10:00.000Z",
    audioExportId: "audio-export-1",
    archiveTokenFactory: () => "private-token"
  });

  assert.equal(core.format, AUDIO_ARCHIVE_FORMAT);
  assert.equal(core.object_count, 1);

  const manifest = finalizeAudioArchiveManifest(core, {
    bundleFilename: "dissertation-audio-audio-export-1.tar.gz",
    bundleSha256: HASH_B
  });

  assert.equal(manifest.bundle_sha256, HASH_B);
  assert.equal(
    manifestsMatchCore(manifest, audioArchiveCoreFromManifest(manifest)),
    true
  );

  const receipt = buildAudioArchiveReceipt(manifest, {
    verifiedAt: "2026-10-02T12:20:00.000Z"
  });

  assert.equal(receipt.format, AUDIO_ARCHIVE_RECEIPT_FORMAT);
  const verified = validateAudioArchiveReceipt({ manifest, receipt });
  assert.equal(verified.verified_object_count, 1);
  assert.equal(verified.verified_objects[0].archive_token, "private-token");
});

test("audio receipt rejects a mismatched token, hash, or package identity", () => {
  const core = buildAudioArchiveCore({
    reportingArchive: reportingArchive(),
    archivedObjects: [
      {
        object_key: "aqg/TEST001/S0001/request-1_feedback.webm",
        archive_filename: "0001.webm",
        sha256: HASH_A,
        byte_length: 123
      }
    ],
    audioExportId: "audio-export-2",
    archiveTokenFactory: () => "token-2"
  });

  const manifest = finalizeAudioArchiveManifest(core, {
    bundleFilename: "bundle.tar.gz",
    bundleSha256: HASH_B
  });
  const receipt = buildAudioArchiveReceipt(manifest);

  const badToken = structuredClone(receipt);
  badToken.verified_objects[0].archive_token = "wrong";
  assert.throws(
    () => validateAudioArchiveReceipt({ manifest, receipt: badToken }),
    /archive token does not match/
  );

  const badHash = structuredClone(receipt);
  badHash.verified_objects[0].sha256 = "c".repeat(64);
  assert.throws(
    () => validateAudioArchiveReceipt({ manifest, receipt: badHash }),
    /hash does not match/
  );

  const badPackage = structuredClone(receipt);
  badPackage.bundle_sha256 = "d".repeat(64);
  assert.throws(
    () => validateAudioArchiveReceipt({ manifest, receipt: badPackage }),
    /bundle SHA-256 does not match/
  );
});
