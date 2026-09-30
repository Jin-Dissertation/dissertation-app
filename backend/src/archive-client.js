import { buildArchiveExport } from "./archive-export.js";

function normalizeBaseUrl(value) {
  return String(value || "").replace(/\/+$/, "");
}

async function getJson({ fetchImpl, url, token }) {
  const response = await fetchImpl(url, {
    method: "GET",
    headers: {
      authorization: `Bearer ${token}`,
      accept: "application/json"
    }
  });

  let body;
  try {
    body = await response.json();
  } catch {
    throw new Error(`Reporting API returned non-JSON HTTP ${response.status}.`);
  }

  if (!response.ok || body?.ok !== true) {
    const code = String(body?.error || `HTTP_${response.status}`);
    throw new Error(`Reporting API request failed: ${code}.`);
  }

  return body;
}

export async function fetchReportingArchive({
  baseUrl,
  token,
  fetchImpl = fetch,
  pageLimit = 100,
  generatedAt,
  exportId,
  archiveTokenFactory
}) {
  const root = normalizeBaseUrl(baseUrl);

  if (!/^https?:\/\//i.test(root)) {
    throw new Error("Reporting API base URL must be http or https.");
  }

  if (typeof token !== "string" || token.length < 32) {
    throw new Error("A reporting export token of at least 32 characters is required.");
  }

  if (!Number.isSafeInteger(pageLimit) || pageLimit < 1 || pageLimit > 100) {
    throw new Error("Reporting page limit must be between 1 and 100.");
  }

  const manifest = await getJson({
    fetchImpl,
    url: `${root}/v1/reporting/manifest`,
    token
  });

  const generation = String(manifest.feed_generation || "");
  const through = Number(manifest.latest_sequence);

  if (!/^[a-f0-9]{32}$/.test(generation)) {
    throw new Error("Reporting manifest returned an invalid feed generation.");
  }

  if (!Number.isSafeInteger(through) || through < 0) {
    throw new Error("Reporting manifest returned an invalid latest sequence.");
  }

  const changes = [];
  let after = 0;

  while (after < through) {
    const params = new URLSearchParams({
      generation,
      after: String(after),
      through: String(through),
      limit: String(pageLimit)
    });

    const page = await getJson({
      fetchImpl,
      url: `${root}/v1/reporting/changes?${params.toString()}`,
      token
    });

    if (String(page.feed_generation || "") !== generation) {
      throw new Error("Reporting feed generation changed during export.");
    }

    if (Number(page.after) !== after) {
      throw new Error("Reporting API returned an unexpected page cursor.");
    }

    if (Number(page.through) !== through) {
      throw new Error("Reporting API changed the export checkpoint.");
    }

    if (!Array.isArray(page.changes)) {
      throw new Error("Reporting API page did not contain a changes array.");
    }

    let expectedSequence = after;

    for (const change of page.changes) {
      const sequence = Number(change?.sequence);

      if (!Number.isSafeInteger(sequence) || sequence <= expectedSequence) {
        throw new Error("Reporting API returned invalid change sequence ordering.");
      }

      expectedSequence = sequence;
      changes.push(change);
    }

    const nextAfter = Number(page.next_after);

    if (
      !Number.isSafeInteger(nextAfter) ||
      nextAfter < after ||
      nextAfter > through
    ) {
      throw new Error("Reporting API returned an invalid next cursor.");
    }

    if (page.changes.length && nextAfter !== expectedSequence) {
      throw new Error("Reporting API next cursor does not match the final change.");
    }

    if (page.has_more === true && nextAfter <= after) {
      throw new Error("Reporting pagination did not advance.");
    }

    if (page.has_more !== true && nextAfter !== through) {
      throw new Error("Reporting pagination ended before the export checkpoint.");
    }

    after = nextAfter;
  }

  if (after !== through) {
    throw new Error("Reporting export did not reach its checkpoint.");
  }

  return buildArchiveExport({
    manifest,
    changes,
    throughSequence: through,
    generatedAt,
    exportId,
    archiveTokenFactory
  });
}
