import test from "node:test";
import assert from "node:assert/strict";
import { fetchReportingArchive } from "../src/archive-client.js";
import {
  REPORTING_DATASETS,
  REPORTING_PROTOCOL_VERSION
} from "../src/reporting-contract.js";

const generation = "0123456789abcdef0123456789abcdef";
const token = "x".repeat(40);

function fullRecord(datasetName, keyValue, overrides = {}) {
  const dataset = REPORTING_DATASETS.find(
    item => item.name === datasetName
  );

  return Object.fromEntries(
    dataset.columns
      .map(column => [
        column,
        column === dataset.key ? keyValue : null
      ])
      .concat(Object.entries(overrides))
  );
}

function response(body, status = 200) {
  return {
    ok: status >= 200 && status < 300,
    status,
    async json() {
      return body;
    }
  };
}

test("client pins checkpoint and reconstructs multiple pages", async () => {
  const seenUrls = [];

  const fakeFetch = async url => {
    seenUrls.push(String(url));

    if (String(url).endsWith("/v1/reporting/manifest")) {
      return response({
        ok: true,
        protocol_version: REPORTING_PROTOCOL_VERSION,
        feed_generation: generation,
        latest_sequence: 2,
        datasets: REPORTING_DATASETS
      });
    }

    const parsed = new URL(url);
    const after = Number(parsed.searchParams.get("after"));

    assert.equal(parsed.searchParams.get("generation"), generation);
    assert.equal(parsed.searchParams.get("through"), "2");

    if (after === 0) {
      return response({
        ok: true,
        protocol_version: REPORTING_PROTOCOL_VERSION,
        feed_generation: generation,
        after: 0,
        through: 2,
        next_after: 1,
        has_more: true,
        changes: [
          {
            sequence: 1,
            dataset: "nonparticipant_button_counts",
            record_id: "button-a",
            revision: 1,
            operation: "upsert",
            changed_at: "2026-09-30T12:00:00.000Z",
            record: fullRecord(
              "nonparticipant_button_counts",
              "button-a",
              {
                press_count: 1,
                updated_at: "2026-09-30T12:00:00.000Z"
              }
            )
          }
        ]
      });
    }

    if (after === 1) {
      return response({
        ok: true,
        protocol_version: REPORTING_PROTOCOL_VERSION,
        feed_generation: generation,
        after: 1,
        through: 2,
        next_after: 2,
        has_more: false,
        changes: [
          {
            sequence: 2,
            dataset: "nonparticipant_button_counts",
            record_id: "button-b",
            revision: 1,
            operation: "upsert",
            changed_at: "2026-09-30T12:01:00.000Z",
            record: fullRecord(
              "nonparticipant_button_counts",
              "button-b",
              {
                press_count: 2,
                updated_at: "2026-09-30T12:01:00.000Z"
              }
            )
          }
        ]
      });
    }

    throw new Error(`Unexpected URL: ${url}`);
  };

  let n = 0;
  const archive = await fetchReportingArchive({
    baseUrl: "https://example.test/",
    token,
    fetchImpl: fakeFetch,
    generatedAt: "2026-09-30T13:00:00.000Z",
    exportId: "synthetic-client-export",
    archiveTokenFactory: () => `token-${++n}`
  });

  assert.equal(seenUrls.length, 3);
  assert.equal(archive.through_sequence, 2);
  assert.equal(archive.total_records, 2);

  const counts = archive.datasets.find(
    item => item.name === "nonparticipant_button_counts"
  );

  assert.deepEqual(
    counts.records.map(item => item.record_id),
    ["button-a", "button-b"]
  );
});

test("client rejects pagination that stops before checkpoint", async () => {
  const fakeFetch = async url => {
    if (String(url).endsWith("/v1/reporting/manifest")) {
      return response({
        ok: true,
        protocol_version: REPORTING_PROTOCOL_VERSION,
        feed_generation: generation,
        latest_sequence: 2,
        datasets: REPORTING_DATASETS
      });
    }

    return response({
      ok: true,
      protocol_version: REPORTING_PROTOCOL_VERSION,
      feed_generation: generation,
      after: 0,
      through: 2,
      next_after: 1,
      has_more: false,
      changes: [
        {
          sequence: 1,
          dataset: "nonparticipant_button_counts",
          record_id: "button-a",
          revision: 1,
          operation: "upsert",
          changed_at: "2026-09-30T12:00:00.000Z",
          record: fullRecord(
            "nonparticipant_button_counts",
            "button-a",
            {
              press_count: 1,
              updated_at: "2026-09-30T12:00:00.000Z"
            }
          )
        }
      ]
    });
  };

  await assert.rejects(
    fetchReportingArchive({
      baseUrl: "https://example.test",
      token,
      fetchImpl: fakeFetch
    }),
    /ended before the export checkpoint/
  );
});
