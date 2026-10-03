/*
 * OPERATOR GUIDE — SECURE PARTICIPANT-CODE ADMINISTRATION
 *
 * Default:
 *   node scripts/provision-access-codes.mjs
 *     → add/reactivate participant codes
 *
 * Deactivate access without deleting study records:
 *   node scripts/provision-access-codes.mjs --deactivate
 *
 * Before either action, PARTICIPANT_PROVISIONING_TOKEN must temporarily exist
 * both in this shell and as a Worker secret. Delete the Worker secret again
 * immediately afterward.
 *
 * Codes are hidden while typed and are not written to a file by this script.
 */

import { createInterface } from "node:readline/promises";
import { execFileSync } from "node:child_process";

const endpoint =
  process.env.PARTICIPANT_PROVISIONING_ENDPOINT ||
  "https://dissertation-study-api.professor-jin.workers.dev/v1/admin/access-codes/provision";

const deactivate = process.argv.includes("--deactivate");
const token = process.env.PARTICIPANT_PROVISIONING_TOKEN || "";
if (token.length < 32) {
  console.error("PARTICIPANT_PROVISIONING_TOKEN is not set in this terminal.");
  process.exit(1);
}

if (!process.stdin.isTTY || !process.stdout.isTTY) {
  console.error("Run this script from an interactive terminal.");
  process.exit(1);
}

const rl = createInterface({ input: process.stdin, output: process.stdout });

// Temporarily disable terminal echo while the participant code is typed.
async function hiddenQuestion(label) {
  process.stdout.write(label);
  try {
    execFileSync("stty", ["-echo"], { stdio: ["inherit", "ignore", "inherit"] });
    return await rl.question("");
  } finally {
    try {
      execFileSync("stty", ["echo"], { stdio: ["inherit", "ignore", "inherit"] });
    } catch {}
    process.stdout.write("\n");
  }
}

// ACCESS_CODE_PEPPER stays inside the Worker; this local script never needs it.
async function provision(participantCode) {
  const response = await fetch(endpoint, {
    method: "POST",
    headers: {
      authorization: "Bearer " + token,
      "content-type": "application/json"
    },
    body: JSON.stringify({
      entries: [{
        participant_code: participantCode,
        active: !deactivate,
        allow_aqg: true,
        allow_training: true
      }]
    })
  });

  let body = null;
  try { body = await response.json(); } catch {}

  if (!response.ok || !body?.ok) {
    throw new Error(body?.code || ("HTTP_" + response.status));
  }
  return body;
}

try {
  console.log(deactivate
    ? "Secure participant-code deactivation"
    : "Secure participant-code provisioning");
  console.log("Each participant uses one participant code for both access and deidentified study identification.");
  console.log(deactivate
    ? "Deactivation blocks future access but preserves existing study records."
    : "Provisioning adds/reactivates access.");
  console.log("Codes are hidden while typed and are never written by this script.\n");

  while (true) {
    const first = (await hiddenQuestion("Participant code (blank to finish): ")).trim();
    if (!first) break;

    const second = (await hiddenQuestion("Re-enter participant code: ")).trim();

    if (first !== second) {
      console.log("Codes did not match; nothing was sent.\n");
      continue;
    }

    try {
      const result = await provision(first);
      console.log(deactivate
        ? "Deactivated 1 participant code.\n"
        : "Provisioned/reactivated 1 participant code.\n");
    } catch (error) {
      console.error("Provisioning failed:", error.message, "\n");
    }
  }
} finally {
  rl.close();
}
