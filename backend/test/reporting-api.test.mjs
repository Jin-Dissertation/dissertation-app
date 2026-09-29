import test from "node:test";
import assert from "node:assert/strict";
import { fixture } from "./fixture.mjs";
import { REPORTING_DATASETS } from "../src/reporting-contract.js";

async function ok(promise) {
  const result = await promise;
  assert.equal(result.status, 200, JSON.stringify(result.body));
  assert.equal(result.body.ok, true);
  return result.body;
}

test("synthetic endpoint journey: AQG/training saves, submissions, items, all event paths, feedback, audio, NP retries", async t => {
  const f = await fixture();
  t.after(f.close);
  const aqg = await ok(f.post("/v1/aqg/create-session", { request_id: "synthetic-aqg-allocate" }));
  const training = await ok(f.post("/v1/training/create-session", { request_id: "synthetic-training-allocate" }));
  assert.equal(await f.count(), 0, "allocations and authentication are operational");
  const retry = async (path, body) => {
    const result = await ok(f.post(path, body));
    const count = await f.count();
    const again = await ok(f.post(path, body));
    assert.equal(again.duplicate, true, path);
    assert.equal(await f.count(), count, path + " retry created a duplicate change");
    return result;
  };
  const event = { event_id: "synthetic-aqg-event", timestamp: "2026-09-29T00:00:00Z", event_type: "starter_prompt_copied",
    details: { prompt: "synthetic", nested: { accessCode: f.code, keep: "research" }, encoded: JSON.stringify({ participant_code: f.code, keep: 1 }) } };
  await retry("/v1/aqg/save-session", {
    session_id: aqg.session_id, request_id: "synthetic-aqg-save", base_revision: 0, batch_events: [event],
    course_context: "Synthetic course", question_context: "Synthetic topic"
  });
  await retry("/v1/aqg/save-session", {
    session_id: aqg.session_id, request_id: "synthetic-aqg-repeat-event", base_revision: 1, batch_events: [event]
  });
  assert.equal(await f.count(), 1, "duplicate event in a new save does not mirror again");
  const stale = await f.post("/v1/aqg/save-session", {
    session_id: aqg.session_id, request_id: "synthetic-aqg-stale", base_revision: 0, batch_events: [{ ...event, event_id: "synthetic-stale-event" }]
  });
  assert.equal(stale.status, 409);
  assert.equal(await f.count(), 1);
  await retry("/v1/aqg/feedback", { session_id: aqg.session_id, feedback_id: "synthetic-aqg-feedback", text_feedback: "Synthetic feedback" });
  await retry("/v1/aqg/upload-audio", {
    session_id: aqg.session_id, request_id: "synthetic-audio", audioBase64: "U1lOVEhFVElD",
    filename: "synthetic.webm", mimeType: "audio/webm", audio_duration_seconds: 1
  });
  await retry("/v1/aqg/submit-session", {
    session_id: aqg.session_id, request_id: "synthetic-aqg-submit", base_revision: 2,
    final_response: "Synthetic final questions", batch_events: [{ ...event, event_id: "synthetic-aqg-final", event_type: "session_closed" }]
  });

  const trainingEvent = { event_id: "synthetic-training-save-event", timestamp: "2026-09-29T00:00:01Z", event_type: "card_entered" };
  await retry("/v1/training/save-session", {
    session_id: training.session_id, request_id: "synthetic-training-save", base_revision: 0,
    content_version: "synthetic-v1", item_values: ["synthetic answer 1", "synthetic answer 2"], item_ms_values: [1000, 2000],
    batch_events: [trainingEvent]
  });
  await retry("/v1/training/append-event", {
    session_id: training.session_id, request_id: "synthetic-append", event_id: "synthetic-append-event",
    event_type: "card_entered", card_index: 2, timestamp: "2026-09-29T00:00:02Z"
  });
  const beforeIgnored = await f.count();
  await ok(f.post("/v1/training/append-event", {
    session_id: training.session_id, request_id: "synthetic-append-duplicate-new-request", event_id: "synthetic-append-event",
    event_type: "card_entered", card_index: 2, timestamp: "2026-09-29T00:00:02Z"
  }));
  assert.equal(await f.count(), beforeIgnored);
  await retry("/v1/training/feedback", {
    session_id: training.session_id, feedback_id: "synthetic-training-feedback", text_feedback: "Synthetic training feedback", section_index: 1
  });
  const trainingSubmit = {
    session_id: training.session_id, request_id: "synthetic-training-submit", base_revision: 1,
    participant_code: f.code, progress_json: { participantCode: f.code }, total_questions: 2, correct_first: 1,
    batch_events: [{ ...trainingEvent, event_id: "synthetic-training-final", event_type: "completed" }]
  };
  await retry("/v1/training/submit-session", trainingSubmit);

  const npPath = "/v1/nonparticipant/button-press";
  await retry(npPath, { button_id: "synthetic-button", request_id: "synthetic-press" });
  await retry(npPath, { button_id: "synthetic-button", request_id: "synthetic-press-2" });
  const conflict = await f.post(npPath, { button_id: "synthetic-other", request_id: "synthetic-press" });
  assert.equal(conflict.status, 409);
  const countBeforeRace = await f.count();
  await Promise.all(Array.from({ length: 5 }, () => ok(f.post(npPath, { button_id: "synthetic-button", request_id: "synthetic-press-race" }))));
  assert.equal(await f.count(), countBeforeRace + 1);

  const manifest = await ok(f.report("manifest"));
  const page = await ok(f.report(`changes?generation=${manifest.feed_generation}&after=0&limit=100`));
  assert.equal(page.has_more, false);
  assert.deepEqual([...new Set(page.changes.map(c => c.dataset))].sort(), REPORTING_DATASETS.map(d => d.name).sort());
  assert.equal(page.changes.filter(c => c.dataset === "training_submission_items").length, 2);
  assert.equal(page.changes.filter(c => c.dataset === "aqg_feedback").length, 2);
  const np = page.changes.filter(c => c.dataset === "nonparticipant_button_counts");
  assert.deepEqual(np.map(c => [c.revision, c.record.press_count]), [[1, 1], [2, 2], [3, 3]]);
  assert.equal(JSON.stringify(page).includes(f.code), false, "access code leaked in export");
  assert.equal(JSON.stringify(page).includes("details_json"), false);
  const detail = JSON.parse(page.changes.find(c => c.record?.event_id === event.event_id).record.detail_json);
  assert.deepEqual(detail.nested, { keep: "research" });
  assert.equal(detail.encoded, '{"keep":1}');
  for (const dataset of REPORTING_DATASETS) {
    const sourceCount = (await f.db.prepare(`SELECT COUNT(*) AS n FROM ${dataset.name}`).first()).n;
    assert.ok(sourceCount > 0);
    assert.ok(page.changes.some(c => c.dataset === dataset.name));
  }
  const audio = await f.mf.getR2Bucket("STUDY_AUDIO");
  assert.equal((await audio.list()).objects.length, 1);
});

test("reporting credential boundary, query validation, missing configuration, and private response headers", async t => {
  const f = await fixture();
  const disabled = await fixture({ tokenConfigured: false });
  t.after(f.close); t.after(disabled.close);
  assert.equal((await disabled.report("manifest")).status, 503);
  assert.equal((await f.request("/v1/reporting/manifest")).status, 401);
  assert.equal((await f.request("/v1/reporting/manifest", { headers: { authorization: "Bearer " + f.code } })).status, 401);
  assert.equal((await f.report("manifest", { method: "POST" })).status, 405);
  assert.equal((await f.report("manifest", { headers: { origin: "https://jin-dissertation.github.io" } })).status, 403);
  assert.equal((await f.report("access_codes")).status, 404);
  const manifest = await f.report("manifest");
  assert.equal(manifest.headers.get("access-control-allow-origin"), null);
  assert.match(manifest.headers.get("cache-control"), /no-store/);
  const g = manifest.body.feed_generation;
  for (const query of ["after=-1", "after=1e2", "after=9007199254740992", "limit=0", "limit=101", "dataset=access_codes", "after=0&after=0", "token=secret", "after=1&through=0"]) {
    assert.equal((await f.report(`changes?generation=${g}&${query}`)).status, 400, query);
  }
  assert.equal((await f.report("changes?after=0")).status, 400);
  assert.equal((await f.report(`changes?generation=${g}&after=1`)).status, 409);
  assert.equal((await f.report(`changes?generation=${g}&through=1`)).status, 409);
  const empty = await ok(f.report(`changes?generation=${g}&after=0`));
  assert.deepEqual([empty.next_after, empty.through, empty.has_more, empty.changes.length], [0, 0, false, 0]);
});

test("bounded cursor replay, concurrent updates, historical values, and deletion tombstones", async t => {
  const f = await fixture();
  t.after(f.close);
  for (let i = 0; i < 5; i++) await ok(f.post("/v1/nonparticipant/button-press", { button_id: "synthetic-pagination", request_id: `synthetic-page-${i}` }));
  const manifest = await ok(f.report("manifest"));
  const path = `changes?generation=${manifest.feed_generation}&after=0&limit=2&through=${manifest.latest_sequence}`;
  const first = await ok(f.report(path));
  await ok(f.post("/v1/nonparticipant/button-press", { button_id: "synthetic-pagination", request_id: "synthetic-after-watermark" }));
  assert.deepEqual(await ok(f.report(path)), first, "replayed page must retain the original values");
  const received = [...first.changes];
  let page = first;
  while (page.has_more) {
    page = await ok(f.report(`changes?generation=${manifest.feed_generation}&after=${page.next_after}&through=${first.through}&limit=2`));
    received.push(...page.changes);
  }
  assert.deepEqual(received.map(c => c.record.press_count), [1, 2, 3, 4, 5]);
  assert.equal(page.next_after, first.through);
  await f.db.prepare("DELETE FROM nonparticipant_button_counts WHERE button_id = 'synthetic-pagination'").run();
  const tail = await ok(f.report(`changes?generation=${manifest.feed_generation}&after=${first.through}`));
  assert.deepEqual(tail.changes.map(c => c.operation), ["upsert", "delete"]);
  assert.equal(tail.changes[1].record, null);
  assert.equal(tail.changes[1].revision, 7);
});

test("large records paginate by bytes without skipping data", async t => {
  const f = await fixture();
  t.after(f.close);
  const text = "Synthetic " + "x".repeat(600000);
  for (let i = 0; i < 3; i++) {
    await f.db.prepare(`INSERT INTO aqg_feedback (record_id, feedback_id, participant_id, submitted_at, text_feedback, created_at)
      VALUES (?1, ?1, 'TEST001', '2026-09-29', ?2, '2026-09-29')`).bind(`synthetic-large-${i}`, text).run();
  }
  const manifest = await ok(f.report("manifest"));
  let after = 0;
  for (let i = 0; i < 3; i++) {
    const page = await ok(f.report(`changes?generation=${manifest.feed_generation}&after=${after}&limit=100`));
    assert.equal(page.changes.length, 1);
    assert.equal(page.changes[0].record.text_feedback.length, text.length);
    after = page.next_after;
  }
  assert.equal(after, manifest.latest_sequence);
});
