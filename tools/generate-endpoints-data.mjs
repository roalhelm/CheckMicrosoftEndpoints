#!/usr/bin/env node
import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const repoRoot = path.resolve(__dirname, "..");

const psScriptPath = path.join(repoRoot, "CheckMicrosoftEndpointsV2.ps1");
const outputDataPath = path.join(repoRoot, "web", "data", "endpoints.generated.json");
const outputHealthPath = path.join(repoRoot, "web", "data", "source-health.generated.json");
const DEFAULT_WHITELIST_MODE = "strict";

const serviceSourceConfig = {
  "Windows Update for Business": {
    id: "windows-update",
    region: "global",
    sourceUrl:
      "https://learn.microsoft.com/en-us/windows/deployment/update/waas-wu-settings",
  },
  "Windows Autopatch": {
    id: "windows-autopatch",
    region: "global",
    sourceUrl: "https://learn.microsoft.com/en-us/windows/deployment/windows-autopatch/",
  },
  Intune: {
    id: "intune",
    region: "north-america",
    sourceUrl:
      "https://learn.microsoft.com/en-us/intune/fundamentals/endpoints?tabs=north-america#endpoints",
    sourceScope: {
      anchorId: "intune-core-service",
      headingHint: "Intune core service",
    },
  },
  "Microsoft Defender": {
    id: "microsoft-defender",
    region: "global",
    sourceUrl:
      "https://learn.microsoft.com/en-us/microsoft-365/security/defender-endpoint/configure-proxy-internet",
  },
  "Azure Active Directory": {
    id: "azure-ad",
    region: "global",
    sourceUrl:
      "https://learn.microsoft.com/en-us/azure/active-directory/hybrid/reference-connect-ports",
  },
  "Microsoft 365": {
    id: "microsoft-365",
    region: "global",
    sourceUrl:
      "https://learn.microsoft.com/en-us/microsoft-365/enterprise/urls-and-ip-address-ranges",
  },
  "Microsoft Store": {
    id: "microsoft-store",
    region: "global",
    sourceUrl:
      "https://learn.microsoft.com/en-us/intune/fundamentals/endpoints?tabs=north-america#microsoft-store-app",
    sourceScope: {
      anchorId: "microsoft-store-app",
      headingHint: "Microsoft Store app",
    },
  },
  "Windows Activation": {
    id: "windows-activation",
    region: "global",
    sourceUrl:
      "https://learn.microsoft.com/en-us/windows/deployment/volume-activation/activate-using-key-management-service-vamt",
  },
  "Microsoft Edge": {
    id: "microsoft-edge",
    region: "global",
    sourceUrl:
      "https://learn.microsoft.com/en-us/deployedge/microsoft-edge-security-endpoints",
  },
  "Windows Telemetry": {
    id: "windows-telemetry",
    region: "global",
    sourceUrl:
      "https://learn.microsoft.com/en-us/windows/privacy/configure-windows-diagnostic-data-in-your-organization",
  },
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
  "windows-update": [
    /windowsupdate/i,
    /update\.microsoft\.com/i,
    /delivery\.mp\.microsoft\.com/i,
    /dl\.delivery\.mp\.microsoft\.com/i,
  ],
  "windows-autopatch": [
    /autopatch/i,
    /mmd/i,
    /manage\.microsoft\.com/i,
    /blob\.core\.windows\.net/i,
    /webpubsub\.azure\.com/i,
    /microsoft\.com/i,
    /windows\.net/i,
  ],
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
  "microsoft-defender": [
    /defender/i,
    /security\.microsoft\.com/i,
    /wdcp\.microsoft\.com/i,
    /definitionupdates\.microsoft\.com/i,
    /\.cp\.wd\.microsoft\.com/i,
    /protection\.outlook\.com/i,
  ],
  "azure-ad": [
    /microsoftonline\.com/i,
    /enterpriseregistration\.windows\.net/i,
    /graph\.windows\.net/i,
    /aadcdn\./i,
    /pas\.windows\.net/i,
    /management\.azure\.com/i,
    /ad\.msft\.net/i,
  ],
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
  "microsoft-store": [
    /store/i,
    /displaycatalog\.mp\.microsoft\.com/i,
    /purchase\.md\.mp\.microsoft\.com/i,
    /licensing\.mp\.microsoft\.com/i,
    /dsx\.mp\.microsoft\.com/i,
    /akamaized\.net/i,
    /s-microsoft\.com/i,
  ],
  "windows-activation": [
    /activation\.sls\.microsoft\.com/i,
    /validation\.sls\.microsoft\.com/i,
    /crl\.microsoft\.com/i,
    /purchase\.md\.mp\.microsoft\.com/i,
    /licensing\.mp\.microsoft\.com/i,
  ],
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
    .replace(/^www\./i, "www.")
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
  const excludeExact = new Set([
    "learn.microsoft.com",
    "github.com",
    "go.microsoft.com",
    "localhost",
    "www.microsoft.com",
  ]);
  if (excludeExact.has(host)) {
    return false;
  }
  const excludePatterns = [
    /poolparty\.biz$/i,
    /js\.monitor\.azure\.com$/i,
    /portal\.azure\.com$/i,
    /wcpstatic\.microsoft\.com$/i,
    /support\.microsoft\.com$/i,
    /techcommunity\.microsoft\.com$/i,
    /office\.edge$/i,
  ];
  if (excludePatterns.some((pattern) => pattern.test(host))) {
    return false;
  }

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

const whitelistMode = resolveWhitelistMode(process.env.ENDPOINT_WHITELIST_MODE);

function extractHostnamesFromText(text) {
  const hosts = new Set();
  let hadCodeSpanHosts = false;

  const codeSpanRegex = /`([^`]+)`/g;
  let codeMatch = codeSpanRegex.exec(text);
  while (codeMatch) {
    const segment = codeMatch[1];
    const segmentHosts = extractHostnamesFromPlainText(segment);
    if (segmentHosts.length > 0) hadCodeSpanHosts = true;
    for (const host of segmentHosts) {
      hosts.add(host);
    }
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
  if (whitelistMode === "relaxed") {
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
  if (endCandidates.length === 0) {
    return text.slice(sliceStart);
  }
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
    if (isLikelyHostname(host)) {
      hosts.add(host);
    }
    match = regex.exec(text);
  }
  return [...hosts].filter(isEndpointCandidateHost);
}

function parseBackendUrlsFromPs(psContent) {
  const services = {};
  const blockStart = psContent.indexOf("$backendUrls = @{");
  if (blockStart === -1) {
    throw new Error("Konnte $backendUrls-Block in CheckMicrosoftEndpointsV2.ps1 nicht finden.");
  }

  const block = psContent.slice(blockStart);
  const serviceRegex = /'([^']+)'\s*=\s*@\(([\s\S]*?)\)\s*\|\s*Sort-Object\s+-Unique/g;
  let match = serviceRegex.exec(block);
  while (match) {
    const serviceName = match[1];
    const entries = match[2];
    const endpointRegex = /'https:\/\/([^']+)'/g;
    const endpoints = new Set();
    let endpointMatch = endpointRegex.exec(entries);
    while (endpointMatch) {
      const host = normalizeHost(endpointMatch[1]);
      if (isLikelyHostname(host)) {
        endpoints.add(host);
      }
      endpointMatch = endpointRegex.exec(entries);
    }
    services[serviceName] = [...endpoints].sort();
    match = serviceRegex.exec(block);
  }
  return services;
}

async function fetchSourceText(url, timeoutMs = 12000) {
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), timeoutMs);
  try {
    const response = await fetch(url, {
      signal: controller.signal,
      headers: { "User-Agent": "CheckMicrosoftEndpoints/endpoint-generator" },
    });
    if (!response.ok) {
      throw new Error(`HTTP ${response.status}`);
    }
    return await response.text();
  } finally {
    clearTimeout(timeout);
  }
}

function extractUpdatedAt(text) {
  const match = text.match(/^updated_at:\s*(.+)$/m) || text.match(/^ms\.date:\s*(.+)$/m);
  return match ? String(match[1]).trim() : null;
}

function mergeBaselineAndLiveHosts(baselineHosts, liveHosts) {
  const merged = new Set(baselineHosts);
  for (const host of liveHosts) {
    merged.add(host);
  }
  return [...merged].sort();
}

async function ensureDir(dirPath) {
  await fs.mkdir(dirPath, { recursive: true });
}

async function main() {
  const psContent = await fs.readFile(psScriptPath, "utf-8");
  const baselineByService = parseBackendUrlsFromPs(psContent);

  const services = [];
  const sourceHealth = [];
  const generatedAt = new Date().toISOString();

  for (const [serviceName, baselineHosts] of Object.entries(baselineByService)) {
    const config = serviceSourceConfig[serviceName] ?? {
      id: serviceName.toLowerCase().replace(/\s+/g, "-"),
      region: "global",
      sourceUrl: null,
    };

    let sourceText = "";
    let liveHosts = [];
    let updatedAt = null;
    let sourceStatus = "fallback";
    let error = null;
    let unknownLiveHosts = [];

    if (config.sourceUrl) {
      try {
        sourceText = await fetchSourceText(config.sourceUrl);
        const scopedText = extractScopedText(sourceText, config.sourceScope);
        const serviceFilteredHosts = filterHostsForService(extractHostnamesFromText(scopedText), config.id);
        const { knownHosts, unknownHosts } = splitKnownAndUnknownHosts(serviceFilteredHosts, config.id);
        unknownLiveHosts = unknownHosts;
        liveHosts = knownHosts;
        updatedAt = extractUpdatedAt(sourceText);
        sourceStatus = liveHosts.length > 0 ? "live+fallback-merged" : "fallback";
      } catch (fetchError) {
        error = fetchError instanceof Error ? fetchError.message : String(fetchError);
      }
    }

    const mergedHosts = mergeBaselineAndLiveHosts(baselineHosts, liveHosts);
    const endpoints = mergedHosts.map((host) => ({
      host,
      url: `https://${host.replace(/^\*\./, "www.")}`,
      purposeShort: inferPurpose(host, serviceName),
      ports: [443],
      source: config.sourceUrl ? "microsoft-official+baseline" : "baseline",
    }));

    const baselineCount = baselineHosts.length;
    const liveCount = liveHosts.length;
    const unknownLiveHostsCount = unknownLiveHosts.length;
    const mergedCount = mergedHosts.length;
    const whitelistPenalty = Math.min(25, unknownLiveHostsCount * 3);
    const qualityScore = Math.max(
      0,
      Math.min(
        100,
        (sourceStatus.startsWith("live") ? 40 : 10) +
          Math.round((Math.min(liveCount, baselineCount) / Math.max(1, baselineCount)) * 45) +
          (updatedAt ? 15 : 5) -
          whitelistPenalty,
      ),
    );

    services.push({
      id: config.id,
      name: serviceName,
      region: config.region,
      sourceUrl: config.sourceUrl,
      sourceScope: config.sourceScope ?? null,
      sourceStatus,
      sourceUpdatedAt: updatedAt,
      endpoints,
    });

    sourceHealth.push({
      service: serviceName,
      serviceId: config.id,
      sourceUrl: config.sourceUrl,
      sourceStatus,
      sourceUpdatedAt: updatedAt,
      baselineCount,
      liveCount,
      unknownLiveHostsCount,
      unknownLiveHostsSample: unknownLiveHosts.slice(0, 8),
      mergedCount,
      qualityScore,
      error,
      whitelistApplied: Boolean(serviceHostWhitelists[config.id]?.length),
    });
  }

  const payload = {
    schemaVersion: 1,
    generatedAt,
    whitelistMode,
    totalServices: services.length,
    totalEndpoints: services.reduce((sum, service) => sum + service.endpoints.length, 0),
    services,
  };

  await ensureDir(path.dirname(outputDataPath));
  await fs.writeFile(outputDataPath, JSON.stringify(payload, null, 2) + "\n", "utf-8");
  await fs.writeFile(
    outputHealthPath,
    JSON.stringify({ schemaVersion: 1, generatedAt, whitelistMode, sourceHealth }, null, 2) + "\n",
    "utf-8",
  );

  const withLiveSources = sourceHealth.filter((entry) => entry.sourceStatus.startsWith("live")).length;
  console.log(`Whitelist mode: ${whitelistMode}`);
  console.log(`Generated ${payload.totalEndpoints} endpoints across ${payload.totalServices} services.`);
  console.log(`Live source merge successful for ${withLiveSources}/${sourceHealth.length} services.`);
}

main().catch((error) => {
  console.error("Failed to generate endpoint data:", error);
  process.exitCode = 1;
});
