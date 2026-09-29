// Real browser + local Worker/D1/R2 only. Every browser request is intercepted;
// the deployed Worker URL is dispatched into the disposable fixture, never sent.
import test from "node:test";
import assert from "node:assert/strict";
import { createServer } from "node:http";
import { readFile, mkdir } from "node:fs/promises";
import { join } from "node:path";
import { setTimeout as delay } from "node:timers/promises";
import { chromium } from "playwright";
import { fixture } from "./fixture.mjs";

const root = new URL("../../", import.meta.url).pathname;
const origin = "http://127.0.0.1:8000";
const firstQuestion = "I am done setting up questions, please show me the first question";
const models = ["ChatGPT", "Gemini", "Copilot", "Other AI"];

async function eventually(check, message) {
  const deadline = Date.now() + 10000;
  do {
    if (await check()) return;
    await delay(50);
  } while (Date.now() < deadline);
  assert.fail(message);
}

async function browserFixture(t, { mobile = false } = {}) {
  const f = await fixture();
  let browser;
  let context;
  const failures = [];
  const server = createServer(async (req, res) => {
    const path = new URL(req.url, origin).pathname;
    const relative = path === "/" ? "index.html" : path.slice(1);
    // Serve public app files only, never credentials or test/runtime files.
    if (!/^(?:index\.html|np\/npindex\.html|aqg-config\.json|(?:np\/)?favicon\.ico|(?:np\/)?images\/[\w.-]+)$/.test(relative)) {
      res.writeHead(404).end(); return;
    }
    try {
      const body = await readFile(join(root, relative));
      const type = relative.endsWith(".html") ? "text/html" : relative.endsWith(".json") ? "application/json" : "application/octet-stream";
      res.writeHead(200, { "content-type": type, "cache-control": "no-store" }).end(body);
    } catch { res.writeHead(404).end(); }
  });
  t.after(async () => {
    await context?.close();
    await browser?.close();
    await new Promise(resolve => server.close(resolve));
    await f.close();
    assert.deepEqual(failures, [], "browser/Worker failures");
  });
  await new Promise((resolve, reject) => {
    server.once("error", reject);
    server.listen(8000, "127.0.0.1", resolve);
  });
  browser = await chromium.launch({
    headless: true,
    ...(process.env.BROWSER_EXECUTABLE_PATH ? { executablePath: process.env.BROWSER_EXECUTABLE_PATH } : {}),
    args: ["--disable-dev-shm-usage"]
  });
  context = await browser.newContext({
    viewport: mobile ? { width: 390, height: 844 } : { width: 1440, height: 1000 },
    permissions: ["clipboard-read", "clipboard-write"], serviceWorkers: "block"
  });
  await context.route("**/*", async route => {
    const request = route.request();
    const url = new URL(request.url());
    if (url.origin === origin) { await route.continue(); return; }
    // Optional font CSS is deliberately empty: no external font downloads.
    if (url.origin === "https://fonts.googleapis.com") {
      await route.fulfill({ status: 200, contentType: "text/css", body: "" }); return;
    }
    if (url.origin !== "https://dissertation-study-api.professor-jin.workers.dev") {
      failures.push("Blocked unexpected external request: " + url.origin + url.pathname);
      await route.abort(); return;
    }
    try {
      const headers = await request.allHeaders();
      const response = await f.mf.dispatchFetch("https://synthetic.invalid" + url.pathname + url.search, {
        method: request.method(),
        headers: { ...(headers.origin ? { origin: headers.origin } : {}), ...(headers["content-type"] ? { "content-type": headers["content-type"] } : {}) },
        ...(request.postData() ? { body: request.postData() } : {})
      });
      if (!response.ok) failures.push(`${url.pathname}: HTTP ${response.status}`);
      await route.fulfill({ status: response.status, headers: Object.fromEntries(response.headers), body: Buffer.from(await response.arrayBuffer()) });
    } catch (error) { failures.push(String(error)); await route.abort().catch(() => {}); }
  });
  const page = await context.newPage();
  page.setDefaultTimeout(10000);
  page.on("pageerror", error => failures.push(error.message));
  const open = async (np = false, training = false) => {
    await page.goto(origin + (np ? "/np/npindex.html" : "/") + (training ? "?mode=training" : ""));
    if (np) {
      await page.locator("#btnOpenTool").click();
    } else if (await page.locator("#accessCodeInput").isVisible()) {
      await page.locator("#accessCodeInput").fill(f.code);
      await page.locator("#btnVerifyAccess").click();
    }
    await page.locator("#appShell").waitFor({ state: "visible" });
  };
  const screenshot = async name => {
    if (!process.env.UI_SCREENSHOT_DIR) return;
    await mkdir(process.env.UI_SCREENSHOT_DIR, { recursive: true });
    await page.locator("#bandSetup").screenshot({ path: join(process.env.UI_SCREENSHOT_DIR, name + ".png") });
  };
  return { ...f, page, context, open, screenshot, failures };
}

async function assertLocked(page) {
  assert.equal(await page.locator("#tagPrompts").textContent(), "Locked");
  assert.equal(await page.locator("#tagFinalize").textContent(), "Locked");
  assert.equal(await page.locator("#btnShowFirstQuestion").isDisabled(), true);
}

async function copyFirstQuestion(page) {
  await page.locator("#btnShowFirstQuestion").click();
  await eventually(async () => (await page.evaluate(() => navigator.clipboard.readText())) === firstQuestion, "generic first-question text was not copied");
  assert.equal(await page.locator("#recentCopiedText").textContent(), firstQuestion);
}

for (const np of [false, true]) {
  test(`${np ? "NP" : "participant"}: skip empty context, all four models, first-question text, reporting, reload and reset`, async t => {
    const f = await browserFixture(t);
    const { page } = f;
    await f.open(np);
    await page.locator('.modelOption[data-model="ChatGPT"]').click();
    await assertLocked(page);
    assert.equal(await page.locator("#ctxCourse").inputValue(), "");
    assert.equal(await page.locator("#ctxArea").inputValue(), "");
    await page.locator("#btnSkipContext").click();
    for (const model of models) {
      await page.locator(`.modelOption[data-model="${model}"]`).click();
      assert.equal(await page.locator("#tagPrompts").textContent(), "Open");
      assert.equal(await page.locator("#tagFinalize").textContent(), "Open");
      assert.equal(await page.locator("#starterCopyCard").getAttribute("aria-disabled"), "true", "skip must not enable an empty starter prompt");
      assert.equal(await page.locator("#setupSamplePromptBox").isVisible(), true);
      await copyFirstQuestion(page);
    }
    const prefix = np ? "dissertation.np" : "dissertation.participant";
    assert.equal(await page.evaluate(key => localStorage.getItem(key), prefix + ".starterCopied"), "false", "skip must not fake a starter copy");
    await f.screenshot(np ? "np-skip" : "participant-skip");
    await page.locator("#sumFinalize").click();
    await page.locator("#finalCopyCard").waitFor({ state: "visible" });
    assert.ok((await page.locator("#finalCopyPreview").textContent()).trim());
    await page.locator("#finalCopyTitle").click();
    assert.equal(await page.evaluate(key => localStorage.getItem(key), prefix + ".finalizeVisited"), "true");
    await page.locator("#sumPrompts").click();
    await page.locator("#promptsInner .promptCard").first().click();
    await page.locator("#btnEditContext").click();

    if (np) {
      await eventually(async () => (await f.db.prepare("SELECT press_count FROM nonparticipant_button_counts WHERE button_id='btnShowFirstQuestion'").first())?.press_count === 4, "NP first-question count missing or doubled");
      assert.equal((await f.db.prepare("SELECT press_count FROM nonparticipant_button_counts WHERE button_id='btnSkipContext'").first()).press_count, 1);
      assert.equal((await f.db.prepare("SELECT COUNT(*) AS n FROM aqg_events").first()).n, 0);
    } else {
      await eventually(async () => (await f.db.prepare("SELECT COUNT(*) AS n FROM aqg_events WHERE button_id='btnShowFirstQuestion'").first()).n === 4, "first-question participant events missing or doubled");
      const skipped = await f.db.prepare("SELECT event_type FROM aqg_events WHERE button_id='btnSkipContext'").all();
      assert.deepEqual(skipped.results.map(r => r.event_type), ["context_skipped"]);
      const live = await f.post("/v1/aqg/latest-live-session", {});
      assert.equal(live.body.progress.contextSkipped, true);
      assert.equal(live.body.progress.starterPromptCopied, false);
      assert.equal(live.body.session.course_context, "");
    }
    const manifest = (await f.report("manifest")).body;
    const report = (await f.report(`changes?generation=${manifest.feed_generation}&after=0&limit=100`)).body;
    const tracked = report.changes.filter(c => c.record?.button_id === "btnSkipContext" || c.record?.button_id === "btnShowFirstQuestion");
    assert.equal(tracked.filter(c => c.record.button_id === "btnSkipContext").length, 1);
    assert.equal(tracked.filter(c => c.record.button_id === "btnShowFirstQuestion").length, 4);
    assert.equal(JSON.stringify(report).includes(f.code), false);

    // Page reload must preserve the skip choice; resetting must clear it.
    await f.open(np);
    assert.equal(await page.locator("#tagPrompts").textContent(), "Open");
    assert.equal(await page.locator("#tagFinalize").textContent(), "Open");
    await eventually(async () => (await page.locator("#promptsInner .promptCard").count()) > 0, "reload lost the unlocked prompts");
    assert.ok((await page.locator("#setupSamplePromptList .promptCard").count()) > 0);
    await page.locator("#btnResetSession").click();
    await page.locator("#btnResetConfirm").click();
    await eventually(async () => await page.evaluate(key => localStorage.getItem(key) !== "true", prefix + ".contextSkipped"), "reset retained the skip choice");
    await assertLocked(page);
  });

  test(`${np ? "NP mobile" : "participant"}: training blocks saved skip, requires current starter, and logs only enabled actions`, async t => {
    const f = await browserFixture(t, { mobile: np });
    const { page } = f;
    const prefix = np ? "dissertation.np" : "dissertation.participant";
    await f.open(np);
    await page.evaluate(key => localStorage.setItem(key, "true"), prefix + ".contextSkipped");
    await f.open(np, true);
    await page.locator('.modelOption[data-model="ChatGPT"]').click();
    await assertLocked(page);
    assert.equal(await page.locator("#btnSkipContext").isVisible(), true);
    assert.equal(await page.locator("#btnSkipContext").isDisabled(), true);
    assert.equal(await page.locator("#skipContextHint").textContent(), "Disabled during training so you can experience the full setup process.");
    // Even a synthetic click event cannot bypass the handler's training guard.
    await page.locator("#btnSkipContext").dispatchEvent("click");
    await assertLocked(page);
    await f.screenshot(np ? "np-training-mobile" : "participant-training");
    await page.locator("#ctxCourse").fill("Synthetic course");
    await page.locator("#ctxArea").fill("Synthetic topic");
    assert.equal(await page.locator("#btnShowFirstQuestion").isDisabled(), true);
    await page.locator("#starterCopyTitle").click();
    await eventually(async () => !(await page.locator("#btnShowFirstQuestion").isDisabled()), "starter did not enable first-question button");
    await copyFirstQuestion(page);
    await page.locator("#ctxArea").fill("Changed synthetic topic");
    assert.equal(await page.locator("#btnShowFirstQuestion").isDisabled(), true);
    assert.equal(await page.locator("#btnGoPrompts").isDisabled(), true);
    await page.locator("#starterCopyTitle").click();
    await page.locator("#btnGoPrompts").click();
    await page.locator("#btnReadyFinal").click();
    await page.locator("#finalCopyCard").waitFor({ state: "visible" });
    if (np) {
      await eventually(async () => (await f.db.prepare("SELECT press_count FROM nonparticipant_button_counts WHERE button_id='btnShowFirstQuestion'").first())?.press_count === 1, "NP first-question tracking missing or duplicated");
      assert.equal(await f.db.prepare("SELECT press_count FROM nonparticipant_button_counts WHERE button_id='btnSkipContext'").first(), null);
    } else {
      await eventually(async () => (await f.db.prepare("SELECT COUNT(*) AS n FROM aqg_events WHERE event_type='first_question_prompt_copied'").first()).n === 1, "participant first-question tracking missing or duplicated");
      assert.equal((await f.db.prepare("SELECT COUNT(*) AS n FROM aqg_events WHERE event_type='context_skipped'").first()).n, 0);
    }
    assert.equal(await page.evaluate(key => localStorage.getItem(key) === "true", prefix + ".contextSkipped"), false);
  });

  test(`${np ? "NP" : "participant"}: a failed clipboard copy still counts one enabled first-question press`, async t => {
    const f = await browserFixture(t);
    await f.open(np);
    await f.page.locator('.modelOption[data-model="Gemini"]').click();
    await f.page.locator("#btnSkipContext").click();
    await f.page.evaluate(() => {
      navigator.clipboard.writeText = async () => { throw new Error("Synthetic clipboard failure"); };
      document.execCommand = () => false;
    });
    await f.page.locator("#btnShowFirstQuestion").focus();
    await f.page.keyboard.press("Enter");
    assert.equal(await f.page.locator("#recentCopiedText").textContent(), firstQuestion, "manual-copy fallback is still available");
    if (np) {
      await eventually(async () => (await f.db.prepare("SELECT press_count FROM nonparticipant_button_counts WHERE button_id='btnShowFirstQuestion'").first())?.press_count === 1, "failed copy press was not counted exactly once");
    } else {
      await eventually(async () => (await f.db.prepare("SELECT COUNT(*) AS n FROM aqg_events WHERE button_id='btnShowFirstQuestion'").first()).n === 1, "failed copy press was not recorded exactly once");
      const event = await f.db.prepare("SELECT event_type, detail_json FROM aqg_events WHERE button_id='btnShowFirstQuestion'").first();
      assert.equal(event.event_type, "first_question_copy_failed");
      assert.equal(JSON.parse(event.detail_json).copy_succeeded, false);
    }
  });
}

test("participant: recover a skipped session from D1, but reject its skip choice in Training Mode", async t => {
  const f = await browserFixture(t);
  const allocation = await f.post("/v1/aqg/create-session", { request_id: "synthetic-skip-restore" });
  assert.equal(allocation.status, 200);
  const saved = await f.post("/v1/aqg/save-session", {
    session_id: allocation.body.session_id, request_id: "synthetic-skip-restore-save", base_revision: 0,
    model_used: "Gemini", llm_product: "gemini", course_context: "", question_context: "",
    progress_json: { stage: "prompts", contextSkipped: true, starterPromptCopied: false }
  });
  assert.equal(saved.status, 200);
  await f.open();
  assert.equal(await f.page.locator("#tagPrompts").textContent(), "Open");
  assert.equal(await f.page.locator("#ctxCourse").inputValue(), "");
  await f.page.locator("#btnEditContext").click();
  await copyFirstQuestion(f.page);
  await eventually(async () => (await f.db.prepare("SELECT COUNT(*) AS n FROM aqg_events WHERE event_type='first_question_prompt_copied'").first()).n === 1, "restored session copy event missing");
  // Clear browser state so the next login really exercises the server restore path.
  await f.page.evaluate(() => localStorage.clear());
  await f.open(false, true);
  await assertLocked(f.page);
  assert.equal(await f.page.locator("#btnSkipContext").isDisabled(), true);
});
