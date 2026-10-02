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
    return await rl.question("");
  } finally {
    try {
      execFileSync("stty", ["echo"], { stdio: ["inherit", "ignore", "inherit"] });
    } catch {}
    process.stdout.write("\n");
  }
}

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
  console.log("Each participant uses one participant code for both access and deidentified study identification.");
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
      console.log("Provisioned 1 participant code.\n");
    } catch (error) {
      console.error("Provisioning failed:", error.message, "\n");
    }
  }
} finally {
  rl.close();
}
