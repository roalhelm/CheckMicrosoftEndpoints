const ENDPOINT_DATA_PATH ="./data/endpoints.generated.json";
const SOURCE_HEALTH_PATH = "./data/source-health.generated.json";
const REQUEST_TIMEOUT_MS = 8000;
const CHECK_CONCURRENCY = 6;

const state = {
  generatedAt: null,
  whitelistMode: null,
  services: [],
  endpointRows: [],
  sourceHealth: [],
  filters: {
    search: "",
    service: "all",
    region: "all",
    status: "all",
    sortBy: "service",
    complianceMode: false,
  },
};

const elements = {
  runCheckBtn: document.getElementById("runCheckBtn"),
  exportJsonBtn: document.getElementById("exportJsonBtn"),
  exportCsvBtn: document.getElementById("exportCsvBtn"),
  shareLinkBtn: document.getElementById("shareLinkBtn"),
  searchInput: document.getElementById("searchInput"),
  serviceFilter: document.getElementById("serviceFilter"),
  regionFilter: document.getElementById("regionFilter"),
  statusFilter: document.getElementById("statusFilter"),
  sortBy: document.getElementById("sortBy"),
  complianceModeToggle: document.getElementById("complianceModeToggle"),
  sourceHealthList: document.getElementById("sourceHealthList"),
  sourcePanelDetails: document.getElementById("sourcePanelDetails"),
  dataOriginSummary: document.getElementById("dataOriginSummary"),
  endpointTableBody: document.getElementById("endpointTableBody"),
  endpointServiceGroups: document.getElementById("endpointServiceGroups"),
  expandAllGroupsBtn: document.getElementById("expandAllGroupsBtn"),
  collapseAllGroupsBtn: document.getElementById("collapseAllGroupsBtn"),
  overallProgress: document.getElementById("overallProgress"),
  overallProgressLabel: document.getElementById("overallProgressLabel"),
  serviceProgressGrid: document.getElementById("serviceProgressGrid"),
  summaryTotal: document.getElementById("summaryTotal"),
  summaryReachable: document.getElementById("summaryReachable"),
  summaryTimeout: document.getElementById("summaryTimeout"),
  summaryUnreachable: document.getElementById("summaryUnreachable"),
  proxyHintsPanel: document.getElementById("proxyHintsPanel"),
  infoDialog: document.getElementById("infoDialog"),
  infoDialogTitle: document.getElementById("infoDialogTitle"),
  infoDialogPurpose: document.getElementById("infoDialogPurpose"),
  infoDialogSource: document.getElementById("infoDialogSource"),
  closeInfoDialogBtn: document.getElementById("closeInfoDialogBtn"),
};

function withTimeout(promiseFactory, timeoutMs) {
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), timeoutMs);
  return promiseFactory(controller.signal).finally(() => clearTimeout(timeout));
}

async function fetchJson(url) {
  const response = await fetch(url, { cache: "no-store" });
  if (!response.ok) throw new Error(`HTTP ${response.status} for ${url}`);
  return response.json();
}

function buildEndpointRowsFromServices(services) {
  return services.flatMap((service) =>
    service.endpoints.map((endpoint) => {
      const testable = Boolean(endpoint.url);
      return {
        id: `${service.id}::${endpoint.host}`,
        serviceId: service.id,
        serviceName: service.name,
        region: service.region ?? "global",
        host: endpoint.host,
        testHost: endpoint.testHost ?? null,
        url: endpoint.url ?? null,
        purposeShort: endpoint.purposeShort ?? "",
        source: endpoint.source ?? service.sourceUrl ?? "generated",
        status: testable ? "untested" : "untestable",
        durationMs: null,
        details: testable ? "Noch nicht getestet" : "Wildcard ohne bekannten Testhost",
      };
    }),
  );
}

function updateFiltersFromQuery() {
  const params = new URLSearchParams(window.location.search);
  const keys = ["search", "service", "region", "status", "sortBy"];
  for (const key of keys) {
    if (params.has(key)) state.filters[key] = params.get(key);
  }
  state.filters.complianceMode = params.get("complianceMode") === "1";
}

function syncFilterControls() {
  elements.searchInput.value = state.filters.search;
  elements.serviceFilter.value = state.filters.service;
  elements.regionFilter.value = state.filters.region;
  elements.statusFilter.value = state.filters.status;
  elements.sortBy.value = state.filters.sortBy;
  elements.complianceModeToggle.checked = state.filters.complianceMode;
}

function buildShareUrl() {
  const params = new URLSearchParams();
  Object.entries(state.filters).forEach(([key, value]) => {
    if (!value || value === "all" || value === false) return;
    params.set(key, value === true ? "1" : String(value));
  });
  return `${window.location.origin}${window.location.pathname}${params.toString() ? `?${params}` : ""}`;
}

function applyFilters(rows) {
  const search = state.filters.search.toLowerCase().trim();
  return rows.filter((row) => {
    if (state.filters.service !== "all" && row.serviceId !== state.filters.service) return false;
    if (state.filters.region !== "all" && row.region !== state.filters.region) return false;
    if (state.filters.status !== "all" && row.status !== state.filters.status) return false;
    if (state.filters.complianceMode && ["reachable", "untested", "untestable"].includes(row.status)) return false;
    if (search && !`${row.host} ${row.serviceName} ${row.purposeShort}`.toLowerCase().includes(search)) return false;
    return true;
  });
}

function sortRows(rows) {
  const clone = [...rows];
  const sortBy = state.filters.sortBy;
  const statusRank = { unreachable: 0, timeout: 1, untested: 2, untestable: 3, reachable: 4 };
  clone.sort((a, b) => {
    if (sortBy === "host") return a.host.localeCompare(b.host);
    if (sortBy === "status") return (statusRank[a.status] ?? 9) - (statusRank[b.status] ?? 9) || a.host.localeCompare(b.host);
    if (sortBy === "duration") return (a.durationMs ?? Number.MAX_SAFE_INTEGER) - (b.durationMs ?? Number.MAX_SAFE_INTEGER);
    return a.serviceName.localeCompare(b.serviceName) || a.host.localeCompare(b.host);
  });
  return clone;
}

function statusLabel(status) {
  if (status === "reachable") return "Erreichbar";
  if (status === "timeout") return "Timeout";
  if (status === "unreachable") return "Nicht erreichbar";
  if (status === "untestable") return "Nicht direkt testbar";
  return "Ungetestet";
}

function renderSourceHealth() {
  const generatedLabel = state.generatedAt ? new Date(state.generatedAt).toLocaleString("de-DE") : "unbekannt";
  elements.dataOriginSummary.textContent =
    `Quelle: Generierte Daten aus offiziellen Microsoft-Quellen, täglich per GitHub Action aktualisiert. ` +
    `Stand: ${generatedLabel}. Whitelist: ${state.whitelistMode ?? "unbekannt"}.`;

  elements.sourceHealthList.innerHTML = "";
  for (const entry of state.sourceHealth) {
    const div = document.createElement("article");
    div.className = "source-item";
    const qualityClass = entry.qualityScore >= 80 ? "reachable" : entry.qualityScore >= 50 ? "timeout" : "unreachable";
    div.innerHTML = `
      <h3>${entry.service}</h3>
      <p class="small-text">Status: <strong>${entry.sourceStatus}</strong></p>
      <p class="small-text">Live: ${entry.liveCount} | Fallback: ${entry.baselineCount} | Effektiv: ${entry.mergedCount}</p>
      <p class="quality status-text ${qualityClass}">Qualität: ${entry.qualityScore}/100</p>
      ${
        entry.unknownLiveHostsCount
          ? `<p class="small-text">Whitelist-Warnung: ${entry.unknownLiveHostsCount} unbekannte Live-Hosts${
              entry.unknownLiveHostsSample?.length ? ` (z. B. ${entry.unknownLiveHostsSample.join(", ")})` : ""
            }</p>`
          : ""
      }
      ${entry.error ? `<p class="small-text">Fehler: ${entry.error}</p>` : ""}
    `;
    elements.sourceHealthList.appendChild(div);
  }
}

function renderServiceProgress() {
  const grouped = new Map();
  state.endpointRows.forEach((row) => {
    if (!grouped.has(row.serviceId)) {
      grouped.set(row.serviceId, {
        serviceName: row.serviceName,
        total: 0,
        done: 0,
      });
    }
    const item = grouped.get(row.serviceId);
    item.total += 1;
    if (row.status !== "untested") item.done += 1;
  });

  elements.serviceProgressGrid.innerHTML = "";
  for (const item of grouped.values()) {
    const percent = item.total ? Math.round((item.done / item.total) * 100) : 0;
    const card = document.createElement("article");
    card.className = "service-progress";
    card.innerHTML = `
      <strong>${item.serviceName}</strong>
      <p class="small-text">${item.done}/${item.total} geprüft (${percent}%)</p>
      <progress max="100" value="${percent}"></progress>
    `;
    elements.serviceProgressGrid.appendChild(card);
  }
}

function renderSummary() {
  const rows = state.endpointRows;
  const total = rows.length;
  const reachable = rows.filter((row) => row.status === "reachable").length;
  const timeout = rows.filter((row) => row.status === "timeout").length;
  const unreachable = rows.filter((row) => row.status === "unreachable").length;
  const tested = rows.filter((row) => row.status !== "untested").length;
  const percent = total ? Math.round((tested / total) * 100) : 0;

  elements.summaryTotal.textContent = `${total} Endpunkte`;
  elements.summaryReachable.textContent = `Erreichbar: ${reachable}`;
  elements.summaryTimeout.textContent = `Timeout: ${timeout}`;
  elements.summaryUnreachable.textContent = `Nicht erreichbar: ${unreachable}`;

  elements.overallProgress.value = percent;
  elements.overallProgressLabel.textContent = `${percent}%`;
  elements.proxyHintsPanel.hidden = unreachable + timeout < 3;

  renderServiceProgress();
}

function renderTable() {
  renderGroupedByService();
  return;
}

function renderGroupedByService() {
  const filtered = sortRows(applyFilters(state.endpointRows));
  const groups = new Map();
  for (const row of filtered) {
    if (!groups.has(row.serviceId)) {
      groups.set(row.serviceId, {
        serviceName: row.serviceName,
        region: row.region,
        rows: [],
      });
    }
    groups.get(row.serviceId).rows.push(row);
  }

  elements.endpointServiceGroups.innerHTML = "";

  for (const [, group] of groups) {
    const reachable = group.rows.filter((row) => row.status === "reachable").length;
    const timeout = group.rows.filter((row) => row.status === "timeout").length;
    const unreachable = group.rows.filter((row) => row.status === "unreachable").length;

    const details = document.createElement("details");
    details.className = "endpoint-group";
    details.open = true;
    details.innerHTML = `
      <summary>
        <strong>${group.serviceName}</strong>
        <span class="group-meta">Region: ${group.region} • ${group.rows.length} Endpunkte • ✅ ${reachable} • ⏱ ${timeout} • ❌ ${unreachable}</span>
      </summary>
      <div class="endpoint-group-content">
        <div class="table-wrap">
          <table class="endpoint-service-table">
            <thead>
              <tr>
                <th>Host</th>
                <th>Info</th>
                <th>Status</th>
                <th>Dauer (ms)</th>
                <th>Quelle</th>
              </tr>
            </thead>
            <tbody>
              ${group.rows
                .map(
                  (row) => `
                <tr>
                  <td>${row.host}${
                    row.testHost && row.testHost !== row.host ? `<br /><span class="small-text">getestet via ${row.testHost}</span>` : ""
                  }</td>
                  <td><button class="info-btn" data-row-id="${row.id}" aria-label="Info für ${row.host}">i</button></td>
                  <td><span class="status-badge ${row.status}">${statusLabel(row.status)}</span></td>
                  <td>${row.durationMs ?? "-"}</td>
                  <td>${row.source}</td>
                </tr>
              `,
                )
                .join("")}
            </tbody>
          </table>
        </div>
      </div>
    `;
    elements.endpointServiceGroups.appendChild(details);
  }

  elements.endpointServiceGroups.querySelectorAll(".info-btn").forEach((button) => {
    button.addEventListener("click", () => {
      const row = state.endpointRows.find((entry) => entry.id === button.dataset.rowId);
      if (!row) return;
      elements.infoDialogTitle.textContent = `${row.host}`;
      elements.infoDialogPurpose.textContent = row.purposeShort;
      const testInfo = row.url ? `Test-URL: ${row.url}` : "Nicht direkt testbar (Wildcard ohne bekannten Testhost)";
      elements.infoDialogSource.textContent = `Service: ${row.serviceName} • Quelle: ${row.source} • ${testInfo}`;
      elements.infoDialog.showModal();
    });
  });
}

function setAllServiceGroupsOpen(isOpen) {
  elements.endpointServiceGroups.querySelectorAll("details.endpoint-group").forEach((group) => {
    group.open = isOpen;
  });
}

function updateFilterOptions() {
  const serviceValues = [...new Map(state.services.map((service) => [service.id, service.name]))];
  const regionValues = [...new Set(state.services.map((service) => service.region ?? "global"))].sort();

  elements.serviceFilter.innerHTML = `<option value="all">Alle Services</option>${serviceValues
    .map(([id, name]) => `<option value="${id}">${name}</option>`)
    .join("")}`;

  elements.regionFilter.innerHTML = `<option value="all">Alle Regionen</option>${regionValues
    .map((region) => `<option value="${region}">${region}</option>`)
    .join("")}`;
}

function rerender() {
  renderSummary();
  renderTable();
}

function updateFilterStateFromControls() {
  state.filters.search = elements.searchInput.value;
  state.filters.service = elements.serviceFilter.value;
  state.filters.region = elements.regionFilter.value;
  state.filters.status = elements.statusFilter.value;
  state.filters.sortBy = elements.sortBy.value;
  state.filters.complianceMode = elements.complianceModeToggle.checked;
}

async function runEndpointCheck(endpoint) {
  const start = performance.now();
  try {
    const corsResult = await withTimeout(
      async (signal) => fetch(endpoint.url, { method: "HEAD", mode: "cors", cache: "no-store", signal }),
      REQUEST_TIMEOUT_MS,
    );
    const durationMs = Math.round(performance.now() - start);
    if (corsResult) {
      return {
        status: "reachable",
        durationMs,
        details: "Antwort erhalten (cors)",
      };
    }
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    if (message.toLowerCase().includes("abort")) {
      return { status: "timeout", durationMs: Math.round(performance.now() - start), details: "Timeout bei cors-Request" };
    }

    try {
      await withTimeout(
        async (signal) =>
          fetch(endpoint.url, {
            method: "GET",
            mode: "no-cors",
            cache: "no-store",
            signal,
          }),
        REQUEST_TIMEOUT_MS,
      );
      return {
        status: "reachable",
        durationMs: Math.round(performance.now() - start),
        details: "Antwort erhalten (no-cors, opaque)",
      };
    } catch (fallbackError) {
      const fallbackMessage = fallbackError instanceof Error ? fallbackError.message : String(fallbackError);
      if (fallbackMessage.toLowerCase().includes("abort")) {
        return { status: "timeout", durationMs: Math.round(performance.now() - start), details: "Timeout bei no-cors-Fallback" };
      }
      return {
        status: "unreachable",
        durationMs: Math.round(performance.now() - start),
        details: `Keine Antwort: ${fallbackMessage}`,
      };
    }
  }

  return {
    status: "unreachable",
    durationMs: Math.round(performance.now() - start),
    details: "Keine verwertbare Antwort",
  };
}

async function runChecks() {
  elements.runCheckBtn.disabled = true;
  const queue = state.endpointRows.filter((row) => row.url);
  const totalTestable = queue.length;
  let done = 0;
  const workers = Array.from({ length: CHECK_CONCURRENCY }, async () => {
    while (queue.length > 0) {
      const next = queue.shift();
      if (!next) break;
      const result = await runEndpointCheck(next);
      next.status = result.status;
      next.durationMs = result.durationMs;
      next.details = result.details;
      done += 1;
      const percent = Math.round((done / totalTestable) * 100);
      elements.overallProgress.value = percent;
      elements.overallProgressLabel.textContent = `${percent}%`;
      renderServiceProgress();
      renderTable();
    }
  });

  await Promise.all(workers);
  elements.runCheckBtn.disabled = false;
  rerender();
}

function download(filename, content, mimeType) {
  const blob = new Blob([content], { type: mimeType });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = filename;
  anchor.click();
  URL.revokeObjectURL(url);
}

function exportJson() {
  const rows = sortRows(applyFilters(state.endpointRows));
  download(
    `endpoint-results-${new Date().toISOString().replace(/[:.]/g, "-")}.json`,
    JSON.stringify({ generatedAt: new Date().toISOString(), filters: state.filters, rows }, null, 2),
    "application/json",
  );
}

function exportCsv() {
  const rows = sortRows(applyFilters(state.endpointRows));
  const headers = ["service", "region", "host", "testHost", "url", "status", "durationMs", "purposeShort", "source", "details"];
  const lines = [
    headers.join(","),
    ...rows.map((row) =>
      headers
        .map((header) => `"${String(row[header] ?? "").replaceAll('"', '""')}"`)
        .join(","),
    ),
  ];
  download(
    `endpoint-results-${new Date().toISOString().replace(/[:.]/g, "-")}.csv`,
    lines.join("\n"),
    "text/csv;charset=utf-8",
  );
}

function attachEventHandlers() {
  const rerenderFromControls = () => {
    updateFilterStateFromControls();
    rerender();
  };

  elements.searchInput.addEventListener("input", rerenderFromControls);
  elements.serviceFilter.addEventListener("change", rerenderFromControls);
  elements.regionFilter.addEventListener("change", rerenderFromControls);
  elements.statusFilter.addEventListener("change", rerenderFromControls);
  elements.sortBy.addEventListener("change", rerenderFromControls);
  elements.complianceModeToggle.addEventListener("change", rerenderFromControls);

  elements.runCheckBtn.addEventListener("click", runChecks);
  elements.exportJsonBtn.addEventListener("click", exportJson);
  elements.exportCsvBtn.addEventListener("click", exportCsv);
  elements.expandAllGroupsBtn.addEventListener("click", () => setAllServiceGroupsOpen(true));
  elements.collapseAllGroupsBtn.addEventListener("click", () => setAllServiceGroupsOpen(false));
  elements.shareLinkBtn.addEventListener("click", async () => {
    updateFilterStateFromControls();
    const link = buildShareUrl();
    await navigator.clipboard.writeText(link);
    elements.shareLinkBtn.textContent = "Link kopiert";
    setTimeout(() => {
      elements.shareLinkBtn.textContent = "Share-Link kopieren";
    }, 1400);
  });
  elements.closeInfoDialogBtn.addEventListener("click", () => elements.infoDialog.close());
}

async function init() {
  updateFiltersFromQuery();

  const [endpointData, sourceHealthData] = await Promise.all([fetchJson(ENDPOINT_DATA_PATH), fetchJson(SOURCE_HEALTH_PATH)]);
  state.generatedAt = endpointData.generatedAt ?? sourceHealthData.generatedAt ?? null;
  state.whitelistMode = endpointData.whitelistMode ?? sourceHealthData.whitelistMode ?? null;
  state.services = endpointData.services ?? [];
  state.endpointRows = buildEndpointRowsFromServices(state.services);
  state.sourceHealth = sourceHealthData.sourceHealth ?? [];

  updateFilterOptions();
  syncFilterControls();
  attachEventHandlers();
  renderSourceHealth();
  rerender();
}

init().catch((error) => {
  console.error("Fehler beim Initialisieren der Web-App:", error);
  elements.dataOriginSummary.textContent = `Initialisierung fehlgeschlagen: ${
    error instanceof Error ? error.message : String(error)
  }`;
});
