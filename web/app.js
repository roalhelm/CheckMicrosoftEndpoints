const FALLBACK_DATA_PATH = "./data/endpoints.generated.json";
const SOURCE_HEALTH_PATH = "./data/source-health.generated.json";
const REQUEST_TIMEOUT_MS = 8000;
const CHECK_CONCURRENCY = 6;
const LIVE_REFRESH_TIMEOUT_MS = 10000;
const DEFAULT_WHITELIST_MODE = "strict";

const state = {
  dataOrigin: "fallback",
  whitelistMode: DEFAULT_WHITELIST_MODE,
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
  refreshLiveBtn: document.getElementById("refreshLiveBtn"),
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
  liveStatus: document.getElementById("liveStatus"),
};

const purposeRules = [
  { pattern: /manage\.microsoft\.com|dm\.microsoft\.com/i, purpose: "Intune Geräteverwaltung und MDM-Kommunikation" },
  { pattern: /graph\.microsoft\.com|graph\.windows\.net/i, purpose: "Microsoft Graph API für Richtlinien, Geräte- und Benutzerdaten" },
  { pattern: /login\.microsoftonline\.com|enterpriseregistration\.windows\.net|aadcdn\./i, purpose: "Anmeldung, Entra ID Authentifizierung und Geräte-Registrierung" },
  { pattern: /delivery\.mp\.microsoft\.com|windowsupdate\.|update\.microsoft\.com/i, purpose: "Windows Update- und Content-Delivery-Endpunkte" },
  { pattern: /defender|wdcp|security\.microsoft\.com|definitionupdates/i, purpose: "Defender Security- und Signatur-Updates" },
  { pattern: /office|office365|sharepoint|onedrive|admin\.microsoft\.com/i, purpose: "Microsoft 365 Portale, Services und Inhalte" },
  { pattern: /store|displaycatalog|licensing\.mp\.microsoft\.com|purchase\.md/i, purpose: "Microsoft Store Katalog, Lizenzierung und Kaufdienste" },
  { pattern: /activation\.sls\.microsoft\.com|validation\.sls\.microsoft\.com|crl\.microsoft\.com/i, purpose: "Windows Aktivierung und Zertifikats-/Validierungsdienste" },
  { pattern: /edge\.microsoft\.com|smartscreen|msedge\./i, purpose: "Microsoft Edge Konfiguration, Updates und SmartScreen-Schutz" },
  { pattern: /events\.data\.microsoft\.com|watson|watcab/i, purpose: "Windows Telemetrie- und Diagnosedatenübertragung" },
];

const serviceHostWhitelists = {
  "windows-update": [/windowsupdate/i, /update\.microsoft\.com/i, /delivery\.mp\.microsoft\.com/i, /dl\.delivery\.mp\.microsoft\.com/i],
  "windows-autopatch": [/autopatch/i, /mmd/i, /manage\.microsoft\.com/i, /blob\.core\.windows\.net/i, /webpubsub\.azure\.com/i, /microsoft\.com/i, /windows\.net/i],
  intune: [
    /manage\.microsoft\.com/i,
    /dm\.microsoft\.com/i,
    /microsoftonline\.com/i,
    /enterpriseregistration\.windows\.net/i,
    /graph\.microsoft\.com/i,
    /(?:^|\.)live\.com$/i,
    /do\.dsp\.mp\.microsoft\.com/i,
    /dl\.delivery\.mp\.microsoft\.com/i,
    /edge\.skype\.com/i,
    /ecs\.office\.com/i,
    /fd\.api\.orgmsg\.microsoft\.com/i,
    /ris\.prod\.api\.personalization\.ideas\.microsoft\.com/i,
  ],
  "microsoft-defender": [/defender/i, /security\.microsoft\.com/i, /wdcp\.microsoft\.com/i, /definitionupdates\.microsoft\.com/i, /\.cp\.wd\.microsoft\.com/i, /protection\.outlook\.com/i],
  "azure-ad": [/microsoftonline\.com/i, /enterpriseregistration\.windows\.net/i, /graph\.windows\.net/i, /aadcdn\./i, /pas\.windows\.net/i, /management\.azure\.com/i, /ad\.msft\.net/i],
  "microsoft-365": [
    /office/i,
    /office365/i,
    /sharepoint/i,
    /onedrive/i,
    /microsoft\.com/i,
    /cloud\.microsoft/i,
    /usercontent\.microsoft/i,
    /\.mx\.microsoft$/i,
    /\.static\.microsoft$/i,
    /\.wns\.windows\.com$/i,
    /storage\.live\.com$/i,
    /g\.live\.com$/i,
    /mediaservices\.windows\.net$/i,
    /adl\.windows\.com$/i,
  ],
  "microsoft-store": [/store/i, /displaycatalog\.mp\.microsoft\.com/i, /purchase\.md\.mp\.microsoft\.com/i, /licensing\.mp\.microsoft\.com/i, /dsx\.mp\.microsoft\.com/i, /akamaized\.net/i, /s-microsoft\.com/i],
  "windows-activation": [/activation\.sls\.microsoft\.com/i, /validation\.sls\.microsoft\.com/i, /crl\.microsoft\.com/i, /purchase\.md\.mp\.microsoft\.com/i, /licensing\.mp\.microsoft\.com/i],
  "microsoft-edge": [
    /edge\.microsoft\.com/i,
    /msedge\./i,
    /smartscreen/i,
    /edge-enterprise\.activity\.windows\.com/i,
    /skype\.com/i,
    /(?:^|\.)dl\.delivery\.mp\.microsoft\.com$/i,
    /msedgeextensions\./i,
  ],
  "windows-telemetry": [
    /events\.data\.microsoft\.com/i,
    /watson/i,
    /watcab/i,
    /blob\.core\.windows\.net/i,
    /vortex-win\.data\.microsoft\.com$/i,
    /oca\.telemetry\.microsoft\.com$/i,
    /oca\.microsoft\.com$/i,
    /settings-win\.data\.microsoft\.com$/i,
    /(?:^|\.)live\.com$/i,
  ],
};

function inferPurpose(host, serviceName) {
  for (const rule of purposeRules) {
    if (rule.pattern.test(host)) {
      return rule.purpose;
    }
  }
  return `${serviceName}: Endpoint für Dienstkommunikation`;
}

function normalizeHost(input) {
  return input
    .trim()
    .replace(/^https?:\/\//i, "")
    .replace(/\/.*$/, "")
    .replace(/^\[|\]$/g, "")
    .replace(/`/g, "")
    .replace(/,+$/, "")
    .toLowerCase();
}

function isLikelyHostname(value) {
  if (!value || value.length < 4) return false;
  if (value.includes("<") || value.includes(">")) return false;
  if (value.includes("/")) return false;
  if (/^\d+\.\d+\.\d+\.\d+(\/\d+)?$/.test(value)) return false;
  if (value.includes(":") && !value.includes(".")) return false;
  if (!/^(?:\*\.)?[a-z0-9][a-z0-9.-]+\.[a-z]{2,12}$/i.test(value)) return false;
  if (/\.(md|json|png|jpg|jpeg|svg|html|htm|xml)$/i.test(value)) return false;
  return true;
}

function isEndpointCandidateHost(host) {
  const excludeExact = new Set(["learn.microsoft.com", "github.com", "go.microsoft.com", "localhost", "www.microsoft.com"]);
  if (excludeExact.has(host)) return false;
  const excludePatterns = [
    /poolparty\.biz$/i,
    /js\.monitor\.azure\.com$/i,
    /portal\.azure\.com$/i,
    /wcpstatic\.microsoft\.com$/i,
    /support\.microsoft\.com$/i,
    /techcommunity\.microsoft\.com$/i,
    /office\.edge$/i,
  ];
  if (excludePatterns.some((pattern) => pattern.test(host))) return false;

  const allowedPatterns = [
    /microsoft/i,
    /windows/i,
    /office/i,
    /sharepoint/i,
    /onedrive/i,
    /azure/i,
    /msauth/i,
    /live\.com$/i,
    /skype\.com$/i,
    /blob\.core\.windows\.net$/i,
    /msocdn/i,
  ];
  return allowedPatterns.some((pattern) => pattern.test(host));
}

function resolveWhitelistMode(value) {
  const mode = String(value ?? DEFAULT_WHITELIST_MODE).trim().toLowerCase();
  if (mode === "relaxed") return "relaxed";
  return "strict";
}

function extractHostnamesFromText(text) {
  const hosts = new Set();
  let hadCodeSpanHosts = false;

  const codeSpanRegex = /`([^`]+)`/g;
  let codeMatch = codeSpanRegex.exec(text);
  while (codeMatch) {
    const segmentHosts = extractHostnamesFromPlainText(codeMatch[1]);
    if (segmentHosts.length > 0) hadCodeSpanHosts = true;
    for (const host of segmentHosts) hosts.add(host);
    codeMatch = codeSpanRegex.exec(text);
  }

  if (!hadCodeSpanHosts) {
    for (const host of extractHostnamesFromPlainText(text)) {
      hosts.add(host);
    }
  }

  return [...hosts].filter(isEndpointCandidateHost);
}

function splitKnownAndUnknownHosts(hosts, serviceId) {
  if (state.whitelistMode === "relaxed") {
    return { knownHosts: hosts, unknownHosts: [] };
  }
  const whitelist = serviceHostWhitelists[serviceId] ?? [];
  if (whitelist.length === 0) {
    return { knownHosts: hosts, unknownHosts: [] };
  }
  const knownHosts = [];
  const unknownHosts = [];
  for (const host of hosts) {
    if (whitelist.some((pattern) => pattern.test(host))) {
      knownHosts.push(host);
    } else {
      unknownHosts.push(host);
    }
  }
  return { knownHosts, unknownHosts };
}

function extractScopedText(text, scope) {
  if (!scope) return text;
  const candidates = [];
  if (scope.anchorId) {
    const escaped = scope.anchorId.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    const anchorPatterns = [
      new RegExp(`id=["']${escaped}["']`, "i"),
      new RegExp(`name=["']${escaped}["']`, "i"),
      new RegExp(`#${escaped}\\b`, "i"),
    ];
    for (const pattern of anchorPatterns) {
      const idx = text.search(pattern);
      if (idx >= 0) candidates.push(idx);
    }
  }
  if (scope.headingHint) {
    const escapedHeading = scope.headingHint.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    const headingPatterns = [
      new RegExp(`\\n#{2,4}\\s+${escapedHeading}\\b`, "i"),
      new RegExp(`<h[2-4][^>]*>[^<]*${escapedHeading}[^<]*<\\/h[2-4]>`, "i"),
    ];
    for (const pattern of headingPatterns) {
      const idx = text.search(pattern);
      if (idx >= 0) candidates.push(idx);
    }
  }
  if (candidates.length === 0) return text;

  const start = Math.min(...candidates);
  const sliceStart = Math.max(0, start - 60);
  const endCandidates = [];
  const endPatterns = [/\n##\s+/g, /\n###\s+/g, /<h2\b/gi, /<h3\b/gi];
  for (const pattern of endPatterns) {
    pattern.lastIndex = start + 120;
    const match = pattern.exec(text);
    if (match?.index && match.index > start + 120) {
      endCandidates.push(match.index);
    }
  }
  if (endCandidates.length === 0) return text.slice(sliceStart);
  return text.slice(sliceStart, Math.min(...endCandidates));
}

function filterHostsForService(hosts, serviceId) {
  if (serviceId === "microsoft-store") {
    return hosts.filter((host) =>
      /store|displaycatalog|purchase\.md|licensing\.mp\.microsoft\.com|dsx\.mp\.microsoft\.com/i.test(host),
    );
  }
  if (serviceId === "intune") {
    return hosts.filter(
      (host) =>
        !/store|displaycatalog|purchase\.md|licensing\.mp\.microsoft\.com|dsx\.mp\.microsoft\.com/i.test(host),
    );
  }
  return hosts;
}

function extractHostnamesFromPlainText(text) {
  const hosts = new Set();
  const regex = /(?:https?:\/\/)?(?:\*\.)?[a-z0-9][a-z0-9.-]*\.[a-z]{2,}/gi;
  let match = regex.exec(text);
  while (match) {
    const host = normalizeHost(match[0]);
    if (isLikelyHostname(host)) hosts.add(host);
    match = regex.exec(text);
  }
  return [...hosts].filter(isEndpointCandidateHost);
}

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

async function fetchTextWithTimeout(url, timeoutMs = LIVE_REFRESH_TIMEOUT_MS) {
  return withTimeout(async (signal) => {
    const response = await fetch(url, {
      signal,
      cache: "no-store",
      headers: { Accept: "text/markdown,text/plain,*/*" },
    });
    if (!response.ok) throw new Error(`HTTP ${response.status}`);
    return response.text();
  }, timeoutMs);
}

function buildEndpointRowsFromServices(services) {
  return services.flatMap((service) =>
    service.endpoints.map((endpoint) => ({
      id: `${service.id}::${endpoint.host}`,
      serviceId: service.id,
      serviceName: service.name,
      region: service.region ?? "global",
      host: endpoint.host,
      url: endpoint.url ?? `https://${endpoint.host.replace(/^\*\./, "www.")}`,
      purposeShort: endpoint.purposeShort ?? inferPurpose(endpoint.host, service.name),
      source: endpoint.source ?? service.sourceUrl ?? "generated",
      status: "untested",
      durationMs: null,
      details: "Noch nicht getestet",
    })),
  );
}

function updateFiltersFromQuery() {
  const params = new URLSearchParams(window.location.search);
  const keys = ["search", "service", "region", "status", "sortBy"];
  for (const key of keys) {
    if (params.has(key)) state.filters[key] = params.get(key);
  }
  state.filters.complianceMode = params.get("complianceMode") === "1";
  state.whitelistMode = resolveWhitelistMode(params.get("whitelistMode"));
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
  if (state.whitelistMode !== DEFAULT_WHITELIST_MODE) {
    params.set("whitelistMode", state.whitelistMode);
  }
  return `${window.location.origin}${window.location.pathname}${params.toString() ? `?${params}` : ""}`;
}

function applyFilters(rows) {
  const search = state.filters.search.toLowerCase().trim();
  return rows.filter((row) => {
    if (state.filters.service !== "all" && row.serviceId !== state.filters.service) return false;
    if (state.filters.region !== "all" && row.region !== state.filters.region) return false;
    if (state.filters.status !== "all" && row.status !== state.filters.status) return false;
    if (state.filters.complianceMode && (row.status === "reachable" || row.status === "untested")) return false;
    if (search && !`${row.host} ${row.serviceName} ${row.purposeShort}`.toLowerCase().includes(search)) return false;
    return true;
  });
}

function sortRows(rows) {
  const clone = [...rows];
  const sortBy = state.filters.sortBy;
  const statusRank = { unreachable: 0, timeout: 1, untested: 2, reachable: 3 };
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
  return "Ungetestet";
}

function setLiveStatus(message, type = "info") {
  const liveStatus = elements.liveStatus;
  if (!liveStatus) return;

  liveStatus.textContent = message;
  liveStatus.classList.remove("loading", "success", "error");

  if (type === "loading") liveStatus.classList.add("loading");
  else if (type === "success") liveStatus.classList.add("success");
  else if (type === "error") liveStatus.classList.add("error");
}

function renderSourceHealth() {
  const modeLabel = state.whitelistMode === "relaxed" ? "relaxed" : "strict";
  elements.dataOriginSummary.textContent =
    state.dataOrigin === "live"
      ? `Quelle: Live-Abruf von Microsoft-Seiten (mit Fallback pro Service). Whitelist: ${modeLabel}.`
      : `Quelle: Generierte Fallback-Daten aus dem Repository. Whitelist: ${modeLabel}.`;

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
                  <td>${row.host}</td>
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
      elements.infoDialogSource.textContent = `Service: ${row.serviceName} • Quelle: ${row.source}`;
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
  const queue = [...state.endpointRows];
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
      const percent = Math.round((done / state.endpointRows.length) * 100);
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

function mergeLiveHostsIntoServices(liveByService) {
  const nextServices = state.services.map((service) => {
    const liveHosts = liveByService.get(service.id);
    if (!liveHosts || liveHosts.length === 0) return service;

    const mergedHosts = new Set(service.endpoints.map((endpoint) => endpoint.host));
    liveHosts.forEach((host) => mergedHosts.add(host));
    const mergedEndpoints = [...mergedHosts].sort().map((host) => ({
      host,
      url: `https://${host.replace(/^\*\./, "www.")}`,
      purposeShort: inferPurpose(host, service.name),
      ports: [443],
      source: "live+fallback-merged",
    }));

    return {
      ...service,
      sourceStatus: "live+fallback-merged",
      endpoints: mergedEndpoints,
    };
  });

  state.services = nextServices;
  state.endpointRows = buildEndpointRowsFromServices(nextServices);
}

async function refreshLiveData() {
  elements.refreshLiveBtn.disabled = true;
  setLiveStatus("Live-Daten werden aktualisiert…", "loading");

  const liveByService = new Map();
  const sourceHealthMap = new Map(state.sourceHealth.map((entry) => [entry.serviceId, { ...entry }]));

  try {
    await Promise.all(
      state.services.map(async (service) => {
        const health = sourceHealthMap.get(service.id) ?? {
          service: service.name,
          serviceId: service.id,
          baselineCount: service.endpoints.length,
          sourceUrl: service.sourceUrl,
        };

        if (!service.sourceUrl) {
          health.sourceStatus = "fallback";
          sourceHealthMap.set(service.id, health);
          return;
        }

        try {
          const text = await fetchTextWithTimeout(service.sourceUrl);
          const scopedText = extractScopedText(text, service.sourceScope ?? null);
          const hosts = filterHostsForService(extractHostnamesFromText(scopedText), service.id);
          const { knownHosts, unknownHosts } = splitKnownAndUnknownHosts(hosts, service.id);
          health.unknownLiveHostsCount = unknownHosts.length;
          health.unknownLiveHostsSample = unknownHosts.slice(0, 8);
          health.whitelistApplied = Boolean(serviceHostWhitelists[service.id]?.length);
          if (knownHosts.length > 0) {
            liveByService.set(service.id, knownHosts);
            health.sourceStatus = "live+fallback-merged";
            health.liveCount = knownHosts.length;
            health.error = null;
            const whitelistPenalty = Math.min(25, unknownHosts.length * 3);
            health.qualityScore = Math.max(
              0,
              Math.min(100, 50 + Math.round((knownHosts.length / Math.max(1, health.baselineCount)) * 50) - whitelistPenalty),
            );
          } else {
            health.sourceStatus = "fallback";
            health.error = unknownHosts.length > 0 ? "Live-Quelle lieferte nur nicht-whitelistete Hosts" : "Live-Quelle ohne auswertbare Hosts";
            health.liveCount = 0;
            health.qualityScore = 20;
          }
        } catch (error) {
          health.sourceStatus = "fallback";
          health.error = error instanceof Error ? error.message : String(error);
          health.liveCount = 0;
          health.qualityScore = 15;
        }

        sourceHealthMap.set(service.id, health);
      }),
    );

    mergeLiveHostsIntoServices(liveByService);
    state.sourceHealth = [...sourceHealthMap.values()];
    state.dataOrigin = liveByService.size > 0 ? "live" : "fallback";
    if (elements.sourcePanelDetails) {
      elements.sourcePanelDetails.open = false;
    }
    updateFilterOptions();
    syncFilterControls();
    renderSourceHealth();
    rerender();
    setLiveStatus(
      liveByService.size > 0
        ? `Live-Daten aktualisiert: ${liveByService.size} Service(s) aktualisiert.`
        : "Live-Daten neu geladen: Fallback-Daten werden weiter verwendet.",
      liveByService.size > 0 ? "success" : "error",
    );
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    setLiveStatus(`Fehler beim Laden der Live-Daten: ${message}`, "error");
    console.error("refreshLiveData failed:", error);
  } finally {
    elements.refreshLiveBtn.disabled = false;
  }
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
  const headers = ["service", "region", "host", "url", "status", "durationMs", "purposeShort", "source", "details"];
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
  elements.refreshLiveBtn.addEventListener("click", refreshLiveData);
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

  const [fallbackData, sourceHealthData] = await Promise.all([fetchJson(FALLBACK_DATA_PATH), fetchJson(SOURCE_HEALTH_PATH)]);
  state.whitelistMode = resolveWhitelistMode(state.whitelistMode ?? sourceHealthData.whitelistMode ?? fallbackData.whitelistMode);
  state.services = fallbackData.services ?? [];
  state.endpointRows = buildEndpointRowsFromServices(state.services);
  state.sourceHealth = sourceHealthData.sourceHealth ?? [];

  updateFilterOptions();
  syncFilterControls();
  attachEventHandlers();
  renderSourceHealth();
  rerender();

  await refreshLiveData();
}

init().catch((error) => {
  console.error("Fehler beim Initialisieren der Web-App:", error);
  elements.dataOriginSummary.textContent = `Initialisierung fehlgeschlagen: ${
    error instanceof Error ? error.message : String(error)
  }`;
});
