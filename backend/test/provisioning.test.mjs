import test from "node:test";
import assert from "node:assert/strict";
import { fixture } from "./fixture.mjs";

test("secure participant provisioning hashes codes server-side", async () => {
  const fx = await fixture({ provisioningTokenConfigured: true });
  try {
    const syntheticCode = "synthetic-production-code-123";
    const participantId = syntheticCode;

    const unauthorized = await fx.request("/v1/admin/access-codes/provision", {
      method: "POST",
      body: { entries: [{ participant_id: participantId, code: syntheticCode }] }
    });
    assert.equal(unauthorized.status, 401);

    const browserBlocked = await fx.request("/v1/admin/access-codes/provision", {
      method: "POST",
      headers: {
        authorization: "Bearer " + fx.provisioningToken,
        origin: "https://jin-dissertation.github.io"
      },
      body: { entries: [{ participant_id: participantId, code: syntheticCode }] }
    });
    assert.equal(browserBlocked.status, 403);

    const provisioned = await fx.request("/v1/admin/access-codes/provision", {
      method: "POST",
      headers: { authorization: "Bearer " + fx.provisioningToken },
      body: { entries: [{ participant_id: participantId, code: syntheticCode }] }
    });
    assert.equal(provisioned.status, 200);
    assert.deepEqual(provisioned.body, {
      ok: true,
      provisioned: 1,
      participant_ids: [participantId]
    });
    assert.equal(JSON.stringify(provisioned.body).includes(syntheticCode), false);

    const row = await fx.db
      .prepare("SELECT participant_id, code_hash, active, allow_aqg, allow_training FROM access_codes WHERE participant_id = ?1")
      .bind(participantId)
      .first();

    assert.equal(row.participant_id, syntheticCode);
    assert.notEqual(row.code_hash, syntheticCode);
    assert.equal(row.code_hash.length, 64);
    assert.equal(Number(row.active), 1);
    assert.equal(Number(row.allow_aqg), 1);
    assert.equal(Number(row.allow_training), 1);

    const validated = await fx.request("/v1/aqg/validate-code", {
      method: "POST",
      body: {
        code: syntheticCode,
        client_key: "synthetic-provisioning-test"
      }
    });
    assert.equal(validated.status, 200);
    assert.equal(validated.body.valid, true);
    assert.equal(validated.body.participant_id, participantId);

    const retained = await fx.db
      .prepare("SELECT COUNT(*) AS n FROM access_codes WHERE participant_id = 'TEST001'")
      .first();
    assert.equal(Number(retained.n), 1);
  } finally {
    await fx.close();
  }
});
