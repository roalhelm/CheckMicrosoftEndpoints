// Microsoft public service status (https://status.cloud.microsoft).
// Shared by the web page (live read in the browser) and tools/fetch-ms-status.mjs
// (scheduled GitHub Action that writes web/data/ms-status.generated.json as fallback).
//
// The status page is a single-page app that reads these public JSON documents:
//   /api/posts/mac           Microsoft 365 for business (one post, Status "Available" = no incident)
//   /api/posts/azure         Azure (one post)
//   /api/posts/ppac          Power Platform (one post)
//   /api/posts/m365Consumer  Microsoft 365 consumer workloads (array, Status "Operational" = healthy)
// No tenant, token or sign-in is involved. Tenant-specific health stays in the admin center.

export const STATUS_PAGE_URL = "https://status.cloud.microsoft";
export const STATUS_HOSTS = ["https://status.cloud.microsoft", "https://msit.status.cloud.microsoft"];

export const STATUS_SOURCES = [
  { id: "mac", name: "Microsoft 365", path: "/api/posts/mac", kind: "post" },
  { id: "azure", name: "Azure", path: "/api/posts/azure", kind: "post" },
  { id: "ppac", name: "Power Platform", path: "/api/posts/ppac", kind: "post" },
  { id: "m365Consumer", name: "Microsoft 365 consumer", path: "/api/posts/m365Consumer", kind: "workloads" },
];

const HEALTHY_STATES = new Set(["available", "operational", "service restored", "resolved", "restored", "false positive"]);

function clean(value) {
  return typeof value === "string" ? value.trim() : "";
}

export function htmlToText(html) {
  return clean(
    String(html ?? "")
      .replace(/<br\s*\/?>/gi, " ")
      .replace(/<\/(div|p|li)>/gi, " ")
      .replace(/<[^>]+>/g, "")
      .replace(/&nbsp;/g, " ")
      .replace(/&amp;/g, "&")
      .replace(/&lt;/g, "<")
      .replace(/&gt;/g, ">")
      .replace(/&quot;/g, '"')
      .replace(/&#39;/g, "'")
      .replace(/\s+/g, " "),
  );
}

function truncate(text, max) {
  if (text.length <= max) return text;
  const cut = text.slice(0, max);
  return `${cut.slice(0, Math.max(cut.lastIndexOf(" "), max - 20)).trim()} ...`;
}

export function severityOf(status) {
  const s = clean(status).toLowerCase();
  if (!s || HEALTHY_STATES.has(s)) return "ok";
  if (s.includes("interruption") || s.includes("outage")) return "outage";
  if (s.includes("maintenance")) return "maintenance";
  return "degraded";
}

function incidentFromPost(source, raw, nameOverride, keepTime = true) {
  const status = clean(raw?.Status);
  const severity = severityOf(status);
  if (severity === "ok") return null;
  const service = clean(nameOverride) || source.name;
  return {
    source: source.id,
    service,
    status: status || "Unknown",
    severity,
    title: htmlToText(raw?.Title) || `${service}: ${status}`,
    summary: truncate(htmlToText(raw?.Message), 360),
    updatedAt: keepTime ? clean(raw?.LastUpdatedTime) || null : null,
  };
}

// Returns the incidents of one document. Throws when the shape is not recognized,
// so a changed format shows as "could not be read" instead of a false all-clear.
export function parseStatusDocument(source, json) {
  if (source.kind === "workloads") {
    if (!Array.isArray(json)) throw new Error("Unexpected format (expected a list of workloads)");
    const rows = json.filter((row) => row && typeof row.Status === "string");
    if (json.length && !rows.length) throw new Error("Unexpected format (no workload status)");
    return rows
      // LastUpdatedTime of a workload changes on every request, so it is not a useful incident time.
      .map((row) => incidentFromPost(source, row, clean(row.ServiceDisplayName) || clean(row.ServiceWorkloadName), false))
      .filter(Boolean);
  }
  if (!json || typeof json !== "object" || Array.isArray(json) || typeof json.Status !== "string") {
    throw new Error("Unexpected format (expected a post with Status)");
  }
  const incident = incidentFromPost(source, json);
  return incident ? [incident] : [];
}

const SEVERITY_ORDER = { outage: 0, degraded: 1, maintenance: 2 };

export function sortIncidents(incidents) {
  return [...incidents].sort(
    (a, b) =>
      (SEVERITY_ORDER[a.severity] ?? 9) - (SEVERITY_ORDER[b.severity] ?? 9) ||
      String(b.updatedAt ?? "").localeCompare(String(a.updatedAt ?? "")),
  );
}

async function fetchWithTimeout(fetchImpl, url, timeoutMs) {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), timeoutMs);
  try {
    const response = await fetchImpl(url, {
      signal: controller.signal,
      cache: "no-store",
      headers: { Accept: "application/json" },
    });
    if (!response.ok) throw new Error(`HTTP ${response.status}`);
    return await response.json();
  } catch (error) {
    if (error?.name === "AbortError") throw new Error("Timeout");
    throw error;
  } finally {
    clearTimeout(timer);
  }
}

// Reads every source, trying each host in turn. Never throws; failures are reported per source.
export async function readMicrosoftStatus({ fetchImpl = fetch, timeoutMs = 8000, hosts = STATUS_HOSTS } = {}) {
  const results = await Promise.all(
    STATUS_SOURCES.map(async (source) => {
      let lastError = null;
      for (const host of hosts) {
        try {
          const json = await fetchWithTimeout(fetchImpl, `${host}${source.path}`, timeoutMs);
          return { id: source.id, name: source.name, ok: true, error: null, incidents: parseStatusDocument(source, json) };
        } catch (error) {
          lastError = error instanceof Error ? error.message : String(error);
        }
      }
      return { id: source.id, name: source.name, ok: false, error: lastError, incidents: [] };
    }),
  );
  return {
    sources: results.map(({ id, name, ok, error }) => ({ id, name, ok, error })),
    incidents: sortIncidents(results.flatMap((result) => result.incidents)),
  };
}
