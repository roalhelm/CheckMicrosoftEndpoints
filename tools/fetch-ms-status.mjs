#!/usr/bin/env node
// Writes web/data/ms-status.generated.json from Microsoft's public status documents.
// The web page reads Microsoft live first and falls back to this file when the browser
// cannot (for example when CORS blocks the request).
//
// To keep the commit history small, the file is only rewritten when the incidents or
// source states change, or when the last check is older than REFRESH_AFTER_HOURS.
import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { readMicrosoftStatus, STATUS_PAGE_URL } from "../web/ms-status.mjs";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const outputPath = path.join(repoRoot, "web", "data", "ms-status.generated.json");
const REFRESH_AFTER_HOURS = 6;

function comparable(data) {
  return JSON.stringify({
    sources: (data?.sources ?? []).map(({ id, ok }) => ({ id, ok })),
    incidents: data?.incidents ?? [],
  });
}

async function readExisting() {
  try {
    return JSON.parse(await fs.readFile(outputPath, "utf8"));
  } catch {
    return null;
  }
}

const status = await readMicrosoftStatus({ timeoutMs: 15000 });
const next = {
  schemaVersion: 1,
  checkedAt: new Date().toISOString(),
  statusPage: STATUS_PAGE_URL,
  ...status,
};

for (const source of status.sources) {
  console.log(`${source.id}: ${source.ok ? "ok" : `failed (${source.error})`}`);
}
console.log(`Active incidents: ${status.incidents.length}`);

const previous = await readExisting();
const ageHours = previous?.checkedAt ? (Date.now() - Date.parse(previous.checkedAt)) / 36e5 : Infinity;

if (previous && comparable(previous) === comparable(next) && ageHours < REFRESH_AFTER_HOURS) {
  console.log(`No change since ${previous.checkedAt}; file left as is.`);
} else {
  await fs.mkdir(path.dirname(outputPath), { recursive: true });
  await fs.writeFile(outputPath, `${JSON.stringify(next, null, 2)}\n`, "utf8");
  console.log(`Wrote ${path.relative(repoRoot, outputPath)}`);
}

if (status.sources.every((source) => !source.ok)) {
  // Not a failure of the job: the page then shows that the status could not be read.
  console.log("::warning::No Microsoft status document could be read.");
}
