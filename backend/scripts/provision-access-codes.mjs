import { createInterface } from "node:readline/promises";
import { execFileSync } from "node:child_process";

const endpoint =
  process.env.PARTICIPANT_PROVISIONING_ENDPOINT ||
  "https://dissertation-study-api.professor-jin.workers.dev/v1/admin/access-codes/provision";

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

async function hiddenQuestion(label) {
  process.stdout.write(label);
  try {
    execFileSync("stty", ["-echo"], { stdio: ["inherit", "ignore", "inherit"] });
    const answer = await rl.question("");
    return answer;
  } finally {
    try {
      execFileSync("stty", ["echo"], { stdio: ["inherit", "ignore", "inherit"] });
    } catch {}
    process.stdout.write("\n");
  }
}

async function provision(participantId, code) {
  const response = await fetch(endpoint, {
    method: "POST",
    headers: {
      authorization: "Bearer " + token,
      "content-type": "application/json"
    },
    body: JSON.stringify({
      entries: [{
        participant_id: participantId,
        code,
        active: true,
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
  console.log("Secure participant-code provisioning");
  console.log("Codes are hidden while typed and are never written by this script.\n");

  while (true) {
    const participantId = (await rl.question("Participant ID (blank to finish): ")).trim();
    if (!participantId) break;

    const first = (await hiddenQuestion("Access code: ")).trim();
    const second = (await hiddenQuestion("Re-enter access code: ")).trim();

    if (!first || first !== second) {
      console.log("Codes did not match; nothing was sent.\n");
      continue;
    }

    try {
      const result = await provision(participantId, first);
      console.log("Provisioned:", result.participant_ids.join(", "), "\n");
    } catch (error) {
      console.error("Provisioning failed for", participantId + ":", error.message, "\n");
    }
  }
} finally {
  rl.close();
}
