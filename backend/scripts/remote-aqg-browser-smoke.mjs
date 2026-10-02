// Synthetic browser smoke test against the deployed migration Worker.
// This serves the cloudflare-migration frontend only on 127.0.0.1 and lets
// browser requests reach the real Worker. It must never be used with real
// participant credentials or data.
import assert from "node:assert/strict";
import { execFileSync } from "node:child_process";
import { createServer } from "node:http";
import { mkdtemp, readFile, rm, stat } from "node:fs/promises";
import { tmpdir } from "node:os";
import { extname, join, normalize } from "node:path";
import { setTimeout as delay } from "node:timers/promises";
import { chromium } from "playwright";

const LOCAL_ORIGIN = "http://127.0.0.1:8000";
const WORKER_ORIGIN = "https://dissertation-study-api.professor-jin.workers.dev";
const EXPECTED_BRANCH = "cloudflare-migration";
const REMOTE_CONFIRMATION = "--execute-synthetic-remote";
const ACCESS_CODE = String(process.env.SMOKE_ACCESS_CODE || "").trim();
const root = new URL("../../", import.meta.url).pathname;

if (!process.argv.includes(REMOTE_CONFIRMATION)) {
  console.error(
    `Refusing to create remote synthetic test data. Re-run with ${REMOTE_CONFIRMATION} after reviewing the command.`
  );
  process.exit(2);
}

if (!ACCESS_CODE) {
  console.error("SMOKE_ACCESS_CODE is required. Use only the synthetic test code.");
  process.exit(2);
}

const branch = execFileSync("git", ["branch", "--show-current"], {
  cwd: root,
  encoding: "utf8"
}).trim();

if (branch !== EXPECTED_BRANCH) {
  console.error(
    `Refusing to run from branch "${branch}". Expected "${EXPECTED_BRANCH}".`
  );
  process.exit(2);
}

function contentType(pathname) {
  switch (extname(pathname).toLowerCase()) {
    case ".html": return "text/html; charset=utf-8";
    case ".json": return "application/json; charset=utf-8";
    case ".png": return "image/png";
    case ".jpg":
    case ".jpeg": return "image/jpeg";
    case ".svg": return "image/svg+xml";
    case ".ico": return "image/x-icon";
    default: return "application/octet-stream";
  }
}

function allowedStaticPath(pathname) {
  if (pathname === "/" || pathname === "/index.html") return "index.html";
  if (pathname === "/aqg-config.json") return "aqg-config.json";
  if (pathname === "/favicon.ico") return "favicon.ico";
  if (/^\/images\/[A-Za-z0-9_.-]+$/.test(pathname)) return pathname.slice(1);
  return "";
}

async function startStaticServer() {
  const server = createServer(async (req, res) => {
    const pathname = new URL(req.url || "/", LOCAL_ORIGIN).pathname;
    const relative = allowedStaticPath(pathname);

    if (!relative) {
      res.writeHead(404, { "cache-control": "no-store" });
      res.end("Not found");
      return;
    }

    try {
      const safeRelative = normalize(relative).replace(/^\.\.(?:\/|\\)/, "");
      const body = await readFile(join(root, safeRelative));
      res.writeHead(200, {
        "content-type": contentType(relative),
        "cache-control": "no-store"
      });
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
  const deadline = Date.now() + timeoutMs;
  let lastError = null;

  while (Date.now() < deadline) {
    try {
      if (await check()) return;
    } catch (error) {
      lastError = error;
    }
    await delay(intervalMs);
  }

  if (lastError) throw lastError;
  assert.fail(message);
}

function parseWranglerJson(raw) {
  const parsed = JSON.parse(String(raw || "").trim());
  assert.ok(Array.isArray(parsed) && parsed[0]?.success === true, "Wrangler D1 query failed");
  assert.equal(parsed[0]?.meta?.changed_db, false, "Read-only verification unexpectedly changed D1");
  return parsed[0]?.results || [];
}

function d1Read(sql) {
  const raw = execFileSync(
    "npx",
    [
      "--yes",
      "wrangler@4.145.0",
      "d1",
      "execute",
      "dissertation-study-data",
      "--remote",
      "--json",
      "--command",
      sql
    ],
    {
      cwd: new URL("../", import.meta.url).pathname,
      encoding: "utf8",
      stdio: ["ignore", "pipe", "pipe"],
      maxBuffer: 4 * 1024 * 1024
    }
  );
  return parseWranglerJson(raw);
}

async function verifyRemoteAudioObject(objectKey) {
  assert.match(
    objectKey,
    /^aqg\/[A-Za-z0-9_.-]+\/[A-Za-z0-9_.-]+\/[A-Za-z0-9_.-]+$/,
    "Worker returned an unexpected R2 object key"
  );

  const dir = await mkdtemp(join(tmpdir(), "aqg-remote-smoke-"));
  const file = join(dir, "audio.bin");

  try {
    execFileSync(
      "npx",
      [
        "--yes",
        "wrangler@4.145.0",
        "r2",
        "object",
        "get",
        `dissertation-study-audio/${objectKey}`,
        "--remote",
        "--file",
        file
      ],
      {
        cwd: new URL("../", import.meta.url).pathname,
        encoding: "utf8",
        stdio: ["ignore", "pipe", "pipe"]
      }
    );

    const info = await stat(file);
    assert.ok(info.size > 0, "Retrieved R2 audio object was empty");
    return info.size;
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
}

async function main() {
  const failures = [];
  const workerResponses = [];
  let browser;
  let context;
  let server;
  let sessionId = "";
  let audioObjectKey = "";
  let notificationId = "";

  try {
    server = await startStaticServer();

    browser = await chromium.launch({
      headless: true,
      args: [
        "--disable-dev-shm-usage",
        "--use-fake-ui-for-media-stream",
        "--use-fake-device-for-media-stream"
      ]
    });

    context = await browser.newContext({
      viewport: { width: 1440, height: 1000 },
      permissions: ["clipboard-read", "clipboard-write", "microphone"],
      serviceWorkers: "block"
    });

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

      failures.push(`unexpected external request: ${url.origin}${url.pathname}`);
      await route.abort().catch(() => {});
    });

    page.on("response", async response => {
      const url = new URL(response.url());
      if (url.origin !== WORKER_ORIGIN) return;

      let body = null;
      try {
        body = await response.json();
      } catch {}

      workerResponses.push({
        path: url.pathname,
        status: response.status(),
        body
      });

      if (response.status() >= 400) {
        failures.push(`${url.pathname}: HTTP ${response.status()}`);
      }

      if (url.pathname === "/v1/aqg/upload-audio" && body?.ok) {
        audioObjectKey = String(body.objectKey || body.audioFileId || body.fileId || "");
        notificationId = String(body.notification_id || notificationId || "");
      }

      if (url.pathname === "/v1/aqg/feedback" && body?.ok) {
        notificationId = String(body.notification_id || notificationId || "");
      }
    });

    await page.goto(`${LOCAL_ORIGIN}/`, { waitUntil: "domcontentloaded" });
    await page.locator("#accessCodeInput").fill(ACCESS_CODE);
    await page.locator("#btnVerifyAccess").click();
    await page.locator("#appShell").waitFor({ state: "visible" });

    // Always start from a fresh synthetic AQG session rather than inheriting
    // an older TEST001 live session.
    const oldSessionId = await page.evaluate(() =>
      localStorage.getItem("dissertation.participant.sessionId") || ""
    );

    await page.locator("#btnResetSession").click();
    await page.locator("#modalReset").waitFor({ state: "visible" });
    await page.locator("#btnResetConfirm").click();
    await page.locator("#modalReset").waitFor({ state: "hidden" });

    await eventually(async () => {
      sessionId = await page.evaluate(() =>
        localStorage.getItem("dissertation.participant.sessionId") || ""
      );
      return Boolean(sessionId && sessionId !== oldSessionId);
    }, "AQG reset did not allocate a fresh synthetic session");

    await page.locator('.modelOption[data-model="ChatGPT"]').click();
    await eventually(
      async () => !(await page.locator("#btnSkipContext").isDisabled()),
      "Skip contextual details did not become available"
    );

    await page.locator("#btnSkipContext").click();
    await eventually(
      async () => (await page.locator("#tagPrompts").textContent()) === "Open",
      "Prompt band did not unlock after context skip"
    );

    const firstQuestion =
      "I am done setting up questions, please show me the first question";
    await page.locator("#btnShowFirstQuestion").click();
    await eventually(
      async () => (await page.locator("#recentCopiedText").textContent()) === firstQuestion,
      "First-question prompt was not copied"
    );

    await page.locator("#sumFinalize").click();
    await page.locator("#finalCopyCard").waitFor({ state: "visible" });
    await page.locator("#finalCopyTitle").click();
    await eventually(
      async () => !(await page.locator("#llmResponse").isDisabled()),
      "Final-response field did not unlock after final prompt copy"
    );

    // The optional question-set textarea lives inside a collapsed <details>.
    // Open it exactly as a participant would before trying to type.
    if (!(await page.locator("#finalModelOutputSection").getAttribute("open"))) {
      await page.locator("#finalModelOutputSection > summary").click();
    }
    await page.locator("#llmResponse").waitFor({ state: "visible" });
    await page.locator("#llmResponse").fill(
      "Synthetic remote browser smoke-test question output only. No participant data."
    );
    await page.locator("#llmResponse").press("Tab");

    // Exercise browser -> Worker -> D1 -> notification relay through the real
    // quick-feedback UI.
    await page.locator("#sidebarBtnFeedback").click();
    await page.locator("#sidebarSessionNotes").fill(
      "Synthetic remote browser smoke-test feedback only. No participant data."
    );
    await page.locator("#sidebarSubmitTextFeedback").click();
    await page.locator("#modalPrivacy").waitFor({ state: "visible" });
    await page.locator("#chkPrivacyConfirm").check();
    await page.locator("#btnPrivacyConfirm").click();
    await eventually(
      async () => (await page.locator("#sidebarFeedbackSavedNote").textContent())?.includes("Text feedback saved"),
      "Synthetic quick feedback did not report success"
    );

    // Exercise actual MediaRecorder UI with Chromium's fake microphone,
    // upload to the deployed Worker/R2, and confirm the remote object exists.
    await page.locator("#sidebarBtnAudioStart").click();
    await eventually(
      async () => (await page.locator("#sidebarBtnAudioState").textContent())?.includes("Recording"),
      "Fake microphone recording did not start"
    );
    await delay(1200);
    await page.locator("#sidebarBtnAudioStop").click();
    await page.locator("#sidebarAudioPlaybackRow").waitFor({ state: "visible" });
    await page.locator("#sidebarBtnAudioSubmit").click();
    await page.locator("#modalPrivacy").waitFor({ state: "visible" });
    await page.locator("#chkPrivacyConfirm").check();
    await page.locator("#btnPrivacyConfirm").click();

    await eventually(
      async () => workerResponses.some(r => r.path === "/v1/aqg/upload-audio" && r.status === 200 && r.body?.ok),
      "Audio upload did not complete successfully",
      30000
    );
    assert.ok(audioObjectKey, "Audio upload did not return an object key");

    // End the same synthetic session so the final submission path is exercised.
    await page.locator("#btnEndSession").waitFor({ state: "visible" });
    await page.locator("#btnEndSession").click();
    await page.locator("#modalEnd").waitFor({ state: "visible" });
    await page.locator("#chkEndPrivacy").check();

    const submitted = page.waitForResponse(
      response =>
        response.url() === `${WORKER_ORIGIN}/v1/aqg/submit-session` &&
        response.request().method() === "POST",
      { timeout: 30000 }
    );

    await page.locator("#btnEndConfirm").click();
    const submitResponse = await submitted;
    assert.equal(submitResponse.status(), 200, "AQG final submission returned a non-200 response");
    const submitBody = await submitResponse.json();
    assert.ok(submitBody?.ok || submitBody?.duplicate, "AQG final submission did not report success");

    await eventually(
      async () => !(await page.locator("#modalEnd").isVisible()),
      "End-session modal did not close after successful submission",
      30000
    );

    assert.deepEqual(failures, [], "Browser/Worker failures were observed");

    await context.close();
    context = null;
    await browser.close();
    browser = null;
    await closeServer(server);
    server = null;

    const audioBytes = await verifyRemoteAudioObject(audioObjectKey);

    assert.match(sessionId, /^[A-Za-z0-9_.-]+$/, "Unexpected synthetic session identifier");

    const sessionRows = d1Read(
      `SELECT
         (SELECT COUNT(*) FROM aqg_submissions WHERE participant_id = 'TEST001' AND session_id = '${sessionId}') AS submissions,
         (SELECT COUNT(*) FROM aqg_feedback WHERE participant_id = 'TEST001' AND session_id = '${sessionId}') AS feedback_rows,
         (SELECT COUNT(*) FROM aqg_events WHERE participant_id = 'TEST001' AND session_id = '${sessionId}') AS events;`
    );

    assert.equal(Number(sessionRows[0]?.submissions || 0), 1, "Final AQG submission was not persisted");
    assert.ok(Number(sessionRows[0]?.feedback_rows || 0) >= 2, "Expected text + audio feedback rows were not persisted");
    assert.ok(Number(sessionRows[0]?.events || 0) > 0, "Expected AQG events were not persisted");

    if (notificationId) {
      assert.match(notificationId, /^[0-9a-f-]{36}$/i, "Unexpected notification identifier");

      await eventually(() => {
        const rows = d1Read(
          `SELECT status, attempt_count, sent_at, last_error
             FROM notification_outbox
            WHERE notification_id = '${notificationId}'
            LIMIT 1;`
        );
        return rows[0]?.status === "sent" &&
          Number(rows[0]?.attempt_count || 0) >= 1 &&
          Boolean(rows[0]?.sent_at) &&
          rows[0]?.last_error == null;
      }, "Browser-generated notification did not reach sent state", 30000, 1000);
    } else {
      assert.fail("No notification ID was returned by browser-generated feedback");
    }

    console.log("");
    console.log("REMOTE AQG BROWSER SMOKE: PASS");
    console.log("------------------------------");
    console.log("Frontend source: local cloudflare-migration branch on 127.0.0.1");
    console.log("Backend: deployed migration Worker");
    console.log("Access + fresh session: PASS");
    console.log("Context skip + first-question prompt: PASS");
    console.log("Quick text feedback + notification delivery: PASS");
    console.log(`Fake-microphone audio upload + remote R2 readback: PASS (${audioBytes} bytes)`);
    console.log("Final AQG submission + D1 persistence: PASS");
    console.log("");
    console.log("Synthetic remote records and the uploaded R2 object were intentionally retained.");
    console.log("Do not delete them directly; archive/verify them before guarded cleanup.");
  } finally {
    await context?.close().catch(() => {});
    await browser?.close().catch(() => {});
    await closeServer(server).catch(() => {});
  }
}

main().catch(error => {
  console.error("");
  console.error("REMOTE AQG BROWSER SMOKE: FAILED");
  console.error(error?.stack || error);
  process.exitCode = 1;
});
