// Guarded synthetic browser smoke for the training module.
// Serves the migration branch locally and sends browser requests to the deployed Worker.
// Never use with real participant credentials or data.
import assert from "node:assert/strict";
import { execFileSync } from "node:child_process";
import { createServer } from "node:http";
import { readFile } from "node:fs/promises";
import { extname, join, normalize } from "node:path";
import { setTimeout as delay } from "node:timers/promises";
import { chromium } from "playwright";

const LOCAL_ORIGIN = "http://127.0.0.1:8000";
const WORKER_ORIGIN = "https://dissertation-study-api.professor-jin.workers.dev";
const ACCESS_CODE = String(process.env.SMOKE_ACCESS_CODE || "").trim();
const CONFIRM = "--execute-synthetic-remote";
const root = new URL("../../", import.meta.url).pathname;

if (!process.argv.includes(CONFIRM)) {
  console.error(`Refusing to create remote synthetic data. Re-run with ${CONFIRM}.`);
  process.exit(2);
}
if (!ACCESS_CODE) {
  console.error("SMOKE_ACCESS_CODE is required. Use only the synthetic test code.");
  process.exit(2);
}
const branch = execFileSync("git", ["branch", "--show-current"], { cwd: root, encoding: "utf8" }).trim();
if (branch !== "cloudflare-migration") {
  console.error(`Refusing to run from branch "${branch}". Expected "cloudflare-migration".`);
  process.exit(2);
}

function typeFor(pathname) {
  switch (extname(pathname).toLowerCase()) {
    case ".html": return "text/html; charset=utf-8";
    case ".csv": return "text/csv; charset=utf-8";
    case ".vtt": return "text/vtt; charset=utf-8";
    case ".mp4": return "video/mp4";
    case ".webm": return "video/webm";
    case ".png": return "image/png";
    case ".jpg":
    case ".jpeg": return "image/jpeg";
    case ".svg": return "image/svg+xml";
    case ".ico": return "image/x-icon";
    default: return "application/octet-stream";
  }
}

function localPath(pathname) {
  if (pathname === "/training/" || pathname === "/training/training.html") return "training/training.html";
  if (pathname === "/training/about.html") return "training/about.html";
  if (pathname === "/training/about.csv") return "training/about.csv";
  if (pathname === "/favicon.ico") return "favicon.ico";
  if (/^\/training\/(?:images|videos)\/[A-Za-z0-9_./-]+$/.test(pathname)) return pathname.slice(1);
  return "";
}

async function startServer() {
  const server = createServer(async (req, res) => {
    const pathname = new URL(req.url || "/", LOCAL_ORIGIN).pathname;
    const relative = localPath(pathname);
    if (!relative) {
      res.writeHead(404, { "cache-control": "no-store" });
      res.end("Not found");
      return;
    }
    try {
      const safe = normalize(relative).replace(/^\.\.(?:\/|\\)/, "");
      const body = await readFile(join(root, safe));
      res.writeHead(200, { "content-type": typeFor(relative), "cache-control": "no-store" });
      res.end(body);
    } catch {
      res.writeHead(404, { "cache-control": "no-store" });
      res.end("Not found");
    }
  });
  await new Promise((resolve, reject) => {
    server.once("error", reject);
    server.listen(8000, "127.0.0.1", resolve);
  });
  return server;
}

async function closeServer(server) {
  if (!server) return;
  await new Promise(resolve => server.close(resolve));
}

async function eventually(check, message, timeoutMs = 20000, intervalMs = 100) {
  const end = Date.now() + timeoutMs;
  let last;
  while (Date.now() < end) {
    try {
      if (await check()) return;
    } catch (error) {
      last = error;
    }
    await delay(intervalMs);
  }
  if (last) throw last;
  assert.fail(message);
}

function d1Read(sql) {
  const raw = execFileSync(
    "npx",
    ["--yes", "wrangler@4.145.0", "d1", "execute", "dissertation-study-data", "--remote", "--json", "--command", sql],
    {
      cwd: new URL("../", import.meta.url).pathname,
      encoding: "utf8",
      stdio: ["ignore", "pipe", "pipe"],
      maxBuffer: 4 * 1024 * 1024
    }
  );
  const parsed = JSON.parse(raw.trim());
  assert.ok(Array.isArray(parsed) && parsed[0]?.success === true, "Wrangler D1 query failed");
  assert.equal(parsed[0]?.meta?.changed_db, false, "Read-only verification unexpectedly changed D1");
  return parsed[0]?.results || [];
}

async function advanceOneCard(page) {
  const before = await page.evaluate(() => getCurrentCardNumber());
  for (let attempt = 0; attempt < 8; attempt += 1) {
    const block = page.locator(".section-block.revealed").last();
    const selectors = [".video-done-btn", ".continue-btn", ".feedback-next-btn.show", ".option-btn:not([disabled])"];
    let clicked = false;
    for (const selector of selectors) {
      const candidate = block.locator(selector).first();
      if (await candidate.isVisible().catch(() => false) && await candidate.isEnabled().catch(() => false)) {
        await candidate.click();
        clicked = true;
        break;
      }
    }
    if (!clicked) throw new Error("Could not find a safe visible action on the current training card.");
    await delay(150);
    const current = await page.evaluate(() => getCurrentCardNumber());
    if (current > before) return;
  }
  throw new Error("Training did not advance after interacting with the first card.");
}

async function main() {
  const failures = [];
  let notificationId = "";
  let sessionId = "";
  let browser, context, server;

  try {
    server = await startServer();
    browser = await chromium.launch({ headless: true, args: ["--disable-dev-shm-usage"] });
    context = await browser.newContext({ viewport: { width: 1440, height: 1000 }, serviceWorkers: "block" });
    const page = await context.newPage();
    page.setDefaultTimeout(20000);

    page.on("pageerror", error => failures.push(`pageerror: ${error.message}`));

    await context.route("**/*", async route => {
      const url = new URL(route.request().url());
      if (url.origin === LOCAL_ORIGIN || url.origin === WORKER_ORIGIN) {
        await route.continue();
        return;
      }
      if (url.origin === "https://fonts.googleapis.com") {
        await route.fulfill({ status: 200, contentType: "text/css", body: "" });
        return;
      }
      if (url.origin === "https://fonts.gstatic.com" ||
          url.hostname.endsWith("youtube.com") ||
          url.hostname === "youtu.be" ||
          url.hostname.endsWith("vimeo.com")) {
        await route.abort().catch(() => {});
        return;
      }
      failures.push(`unexpected external request: ${url.origin}${url.pathname}`);
      await route.abort().catch(() => {});
    });

    page.on("response", async response => {
      const url = new URL(response.url());
      if (url.origin !== WORKER_ORIGIN) return;
      let body = null;
      try { body = await response.json(); } catch {}
      if (response.status() >= 400) failures.push(`${url.pathname}: HTTP ${response.status()}`);
      if (url.pathname === "/v1/training/feedback" && body?.ok) {
        notificationId = String(body.notification_id || notificationId || "");
      }
    });

    await page.goto(`${LOCAL_ORIGIN}/training/training.html`, { waitUntil: "domcontentloaded" });
    await page.locator("#code-input").fill(ACCESS_CODE);
    await page.locator("#gate-submit-btn").click();

    await Promise.race([
      page.locator("#resume-modal").waitFor({ state: "visible" }),
      page.locator("#training").waitFor({ state: "visible" })
    ]);
    if (await page.locator("#resume-modal").isVisible()) {
      await page.locator("#resume-startover-btn").click();
    }

    await page.locator("#training").waitFor({ state: "visible" });
    await page.locator("#sections-container .section-block.revealed").first().waitFor({ state: "visible" });

    sessionId = await page.evaluate(() => String(sessionId || ""));
    assert.match(sessionId, /^[A-Za-z0-9_.-]+$/, "Unexpected training session identifier");

    await advanceOneCard(page);

    const saved = page.waitForResponse(
      r => r.url() === `${WORKER_ORIGIN}/v1/training/save-session` &&
           r.request().method() === "POST" &&
           r.status() === 200,
      { timeout: 30000 }
    );
    await page.locator("#btn-save-exit").click();
    await saved;
    await page.locator("#gate").waitFor({ state: "visible" });

    await page.locator("#code-input").fill(ACCESS_CODE);
    await page.locator("#gate-submit-btn").click();
    await page.locator("#resume-modal").waitFor({ state: "visible" });
    await page.locator("#resume-continue-btn").click();
    await page.locator("#training").waitFor({ state: "visible" });

    assert.equal(
      await page.evaluate(() => String(sessionId || "")),
      sessionId,
      "Training resumed a different session"
    );

    await page.locator("#btn-report-issue").click();
    await page.locator("#feedback-modal").waitFor({ state: "visible" });
    await page.locator("#feedback-textarea").fill(
      "Synthetic remote training browser smoke-test feedback only. No participant data."
    );
    await page.locator("#feedback-submit-btn").click();
    await eventually(
      async () => (await page.locator("#feedback-status").textContent())?.includes("Thanks"),
      "Training feedback did not report success",
      30000
    );

    const submitted = page.waitForResponse(
      r => r.url() === `${WORKER_ORIGIN}/v1/training/submit-session` &&
           r.request().method() === "POST",
      { timeout: 30000 }
    );

    // Deliberately accelerates only the final completion trigger. This smoke
    // validates browser transport, save/resume, feedback, notification, and
    // submission persistence; it is not a card-by-card instructional test.
    await page.evaluate(() => showCompletion());
    const submitResponse = await submitted;
    assert.equal(submitResponse.status(), 200, "Training submission returned a non-200 response");
    const submitBody = await submitResponse.json();
    assert.ok(submitBody?.ok || submitBody?.duplicate, "Training submission did not report success");

    await eventually(
      async () => page.evaluate(() => hasSuccessfulSubmission === true),
      "Training page did not enter submitted state",
      30000
    );

    assert.deepEqual(failures, [], "Browser/Worker failures were observed");

    await context.close(); context = null;
    await browser.close(); browser = null;
    await closeServer(server); server = null;

    const rows = d1Read(
      `SELECT
        (SELECT COUNT(*) FROM training_submissions WHERE participant_id='TEST001' AND session_id='${sessionId}') AS submissions,
        (SELECT COUNT(*) FROM training_submission_items WHERE participant_id='TEST001' AND session_id='${sessionId}') AS submission_items,
        (SELECT COUNT(*) FROM training_events WHERE participant_id='TEST001' AND session_id='${sessionId}') AS events,
        (SELECT COUNT(*) FROM training_feedback WHERE participant_id='TEST001' AND session_id='${sessionId}') AS feedback_rows,
        (SELECT COUNT(*) FROM training_live_sessions WHERE participant_id='TEST001' AND session_id='${sessionId}' AND status='submitted') AS submitted_live;`
    );

    assert.equal(Number(rows[0]?.submissions || 0), 1, "Training submission was not persisted");
    assert.ok(Number(rows[0]?.submission_items || 0) > 0, "Training submission items were not persisted");
    assert.ok(Number(rows[0]?.events || 0) > 0, "Training events were not persisted");
    assert.ok(Number(rows[0]?.feedback_rows || 0) >= 1, "Training feedback was not persisted");
    assert.equal(Number(rows[0]?.submitted_live || 0), 1, "Training live row was not submitted");

    assert.match(notificationId, /^[0-9a-f-]{36}$/i, "No training feedback notification ID was returned");
    await eventually(() => {
      const n = d1Read(
        `SELECT status, attempt_count, sent_at, last_error
         FROM notification_outbox
         WHERE notification_id='${notificationId}' LIMIT 1;`
      )[0];
      return n?.status === "sent" &&
        Number(n?.attempt_count || 0) >= 1 &&
        Boolean(n?.sent_at) &&
        n?.last_error == null;
    }, "Training feedback notification did not reach sent state", 30000, 1000);

    console.log("");
    console.log("REMOTE TRAINING BROWSER SMOKE: PASS");
    console.log("-----------------------------------");
    console.log("Frontend source: local cloudflare-migration branch on 127.0.0.1");
    console.log("Backend: deployed migration Worker");
    console.log("Access + training content load: PASS");
    console.log("Save and exit + server-side resume: PASS");
    console.log("Training feedback + notification delivery: PASS");
    console.log("Training submission transport + D1 persistence: PASS");
    console.log("");
    console.log("Scope note: final completion was triggered programmatically after the resume check;");
    console.log("this does not replace participant-facing card-by-card module testing.");
    console.log("");
    console.log("Synthetic remote training records were intentionally retained.");
    console.log("Archive/verify them before guarded cleanup.");
  } finally {
    await context?.close().catch(() => {});
    await browser?.close().catch(() => {});
    await closeServer(server).catch(() => {});
  }
}

main().catch(error => {
  console.error("");
  console.error("REMOTE TRAINING BROWSER SMOKE: FAILED");
  console.error(error?.stack || error);
  process.exitCode = 1;
});
