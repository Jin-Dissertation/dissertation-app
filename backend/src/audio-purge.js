import { validateAudioArchiveReceipt } from "./audio-archive.js";

function integer(value) {
  const number = Number(value);
  return Number.isSafeInteger(number) && number >= 0 ? number : 0;
}

export function buildAudioPurgePlan({
  manifest,
  receipt,
  objectStates,
  operationalReferences
}) {
  const verified = validateAudioArchiveReceipt({ manifest, receipt });

  if (!objectStates || typeof objectStates !== "object") {
    throw new Error("Audio object states are required.");
  }

  if (!operationalReferences || typeof operationalReferences !== "object") {
    throw new Error("Operational audio reference counts are required.");
  }

  const items = manifest.objects.map(object => {
    const state = objectStates[object.object_key] || null;
    const refs = operationalReferences[object.object_key] || {};

    const reasons = [];

    if (!state || state.exists !== true) {
      reasons.push("r2_object_missing");
    } else {
      if (String(state.sha256 || "") !== String(object.sha256)) {
        reasons.push("r2_object_hash_changed");
      }
      if (Number(state.byte_length) !== Number(object.byte_length)) {
        reasons.push("r2_object_size_changed");
      }
    }

    const liveSessions = integer(refs.live_sessions);
    const submissions = integer(refs.submissions);
    const feedback = integer(refs.feedback);
    const pendingNotifications = integer(refs.pending_notifications);

    if (liveSessions > 0) reasons.push("referenced_by_live_session");
    if (submissions > 0) reasons.push("referenced_by_submission");
    if (feedback > 0) reasons.push("referenced_by_feedback");
    if (pendingNotifications > 0) {
      reasons.push("referenced_by_pending_notification");
    }

    return {
      object_key: object.object_key,
      sha256: object.sha256,
      byte_length: object.byte_length,
      eligible: reasons.length === 0,
      keep_reasons: reasons,
      references: {
        live_sessions: liveSessions,
        submissions,
        feedback,
        pending_notifications: pendingNotifications
      }
    };
  });

  const eligible = items.filter(item => item.eligible);
  const blocked = items.filter(item => !item.eligible);

  return {
    audio_export_id: verified.audio_export_id,
    reporting_export_id: verified.reporting_export_id,
    verified_object_count: verified.verified_object_count,
    eligible_count: eligible.length,
    blocked_count: blocked.length,
    items,
    eligible_objects: eligible.map(item => item.object_key),
    blocked_objects: blocked.map(item => ({
      object_key: item.object_key,
      keep_reasons: item.keep_reasons
    }))
  };
}
