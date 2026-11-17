// Maps endpoints to the Microsoft service feature they belong to and turns
// failed reachability checks into a plain-language "what will break" summary.
//
// Keep IMPACT_RULES in sync with $endpointImpactRules in CheckMicrosoftEndpointsV2.ps1.
// Rules are evaluated top to bottom, the first match wins, so specific hosts
// must come before broad domain rules.
//
// Severity:
//   critical  core function of the service stops working
//   degraded  service works, but a feature is missing or slower
//   low       optional, admin-only or telemetry; users will rarely notice

export const SEVERITY_ORDER = { critical: 0, degraded: 1, low: 2 };

export const SEVERITY_LABELS = {
  critical: "Critical",
  degraded: "Degraded",
  low: "Low",
};

export const IMPACT_RULES = [
  // --- Identity -------------------------------------------------------------
  {
    id: "entra-certauth",
    area: "Entra ID",
    feature: "Certificate-based authentication",
    severity: "degraded",
    impact: "Sign-in with certificates or smart cards fails; password and passwordless sign-in still work.",
    pattern: /(?:^|\.)certauth\.[a-z0-9.-]+$/i,
  },
  {
    id: "entra-signin",
    area: "Entra ID",
    feature: "Sign-in and token issuance",
    severity: "critical",
    impact: "Users cannot sign in or renew tokens. Almost every Microsoft cloud service is affected.",
    pattern: /(?:^|\.)(?:login\.microsoft\.com|login\.windows\.net|msauth\.net|msftauth\.net|msauthimages\.net|auth\.microsoft\.com|microsoftonline-p\.com|microsoftonline-p\.net|accesscontrol\.windows\.net|activedirectory\.windowsazure\.com)$|^(?:[a-z0-9-]+\.)?login[a-z0-9-]*\.microsoftonline\.com$|^microsoftonline\.com$/i,
  },
  {
    id: "entra-sspr",
    area: "Entra ID",
    feature: "Self-service password reset",
    severity: "degraded",
    impact: "Users cannot reset or change their password themselves.",
    pattern: /(?:^|\.)passwordreset\.microsoftonline\.com$/i,
  },
  {
    id: "entra-seamless-sso",
    area: "Entra ID",
    feature: "Seamless single sign-on",
    severity: "degraded",
    impact: "Seamless SSO on domain-joined devices fails; users are asked for their password.",
    pattern: /(?:^|\.)autologon\.microsoftazuread-sso\.com$/i,
  },
  {
    id: "entra-admin",
    area: "Entra ID",
    feature: "Directory administration",
    severity: "low",
    impact: "Legacy admin tools and Entra Connect provisioning calls fail. Users are not affected.",
    pattern: /(?:^|\.)(?:adminwebservice|companymanager|provisioningapi)\.microsoftonline\.com$/i,
  },
  {
    id: "entra-other",
    area: "Entra ID",
    feature: "Sign-in and token issuance",
    severity: "critical",
    impact: "Users cannot sign in or renew tokens. Almost every Microsoft cloud service is affected.",
    pattern: /(?:^|\.)microsoftonline\.com$/i,
  },
  {
    id: "entra-device-registration",
    area: "Entra ID",
    feature: "Device registration",
    severity: "critical",
    impact: "Entra join, hybrid join and device registration fail. Conditional Access rules that require a compliant or joined device block access.",
    pattern: /(?:^|\.)enterpriseregistration\.windows\.net$/i,
  },
  {
    id: "entra-pas",
    area: "Entra ID",
    feature: "Windows sign-in services",
    severity: "degraded",
    impact: "Entra ID sign-in to Windows devices and VMs and some passwordless scenarios can fail.",
    pattern: /(?:^|\.)pas\.windows\.net$/i,
  },
  {
    id: "graph",
    area: "Microsoft Graph",
    feature: "Graph API",
    severity: "degraded",
    impact: "Admin portals, scripts and apps that use Microsoft Graph fail. Some Intune and Autopilot features rely on it.",
    pattern: /(?:^|\.)graph\.(?:microsoft\.com|windows\.net)$/i,
  },
  {
    id: "azure-arm",
    area: "Azure",
    feature: "Azure Resource Manager",
    severity: "low",
    impact: "Azure portal and Azure management scripts fail. Client devices are not affected.",
    pattern: /(?:^|\.)management\.azure\.com$/i,
  },

  // --- Windows Autopilot / Autopatch -----------------------------------------
  {
    id: "autopilot",
    area: "Windows Autopilot",
    feature: "Deployment profile",
    severity: "critical",
    impact: "Devices cannot download their Autopilot profile and go through a normal, unmanaged setup.",
    pattern: /(?:^|\.)(?:ztd|cs)\.dds\.microsoft\.com$|(?:^|\.)dds\.microsoft\.com$/i,
  },
  {
    id: "autopatch",
    area: "Windows Autopatch",
    feature: "Autopatch service",
    severity: "degraded",
    impact: "Autopatch device registration, readiness checks and update reporting fail. Updates themselves still flow through Windows Update.",
    pattern: /(?:^|\.)(?:autopatch\.microsoft\.com|mmdcustomer\.microsoft\.com|mmdls\.microsoft\.com|devicelistenerprod\.microsoft\.com|webpubsub\.azure\.com)$|^payloadprod[a-z0-9*-]*\.blob\.core\.windows\.net$/i,
  },

  // --- Intune ---------------------------------------------------------------
  {
    id: "intune-enrollment",
    area: "Intune",
    feature: "Device enrollment",
    severity: "critical",
    impact: "New devices cannot enroll in Intune (Autopilot, user and bulk enrollment).",
    pattern: /(?:^|\.)enrollment\.manage\.microsoft\.com$|(?:^|\.)enterpriseenrollment[a-z0-9.-]*$/i,
  },
  {
    id: "intune-content",
    area: "Intune",
    feature: "App and script downloads",
    severity: "degraded",
    impact: "Win32 apps, PowerShell scripts and remediations cannot be downloaded; policies still apply.",
    pattern: /mscdn\.manage\.microsoft\.com$|^swd[a-z0-9-]*\.manage\.microsoft\.com$|(?:^|\.)naprodimedatapri\.azureedge\.net$|^mmcswd[a-z0-9-]*\.blob\.core\.windows\.net$/i,
  },
  {
    id: "intune-org-messages",
    area: "Intune",
    feature: "Organizational messages",
    severity: "low",
    impact: "Organizational messages are not shown on devices.",
    pattern: /(?:^|\.)fd\.api\.orgmsg\.microsoft\.com$|(?:^|\.)ris\.prod\.api\.personalization\.ideas\.microsoft\.com$/i,
  },
  {
    id: "intune-company-portal",
    area: "Intune",
    feature: "Company Portal",
    severity: "degraded",
    impact: "Company Portal app and website cannot load; self-service app installs fail.",
    pattern: /(?:^|\.)portal\.manage\.microsoft\.com$/i,
  },
  {
    id: "intune-admin",
    area: "Intune",
    feature: "Admin center",
    severity: "low",
    impact: "The Intune admin center cannot be opened from this network. Managed devices are not affected.",
    pattern: /(?:^|\.)(?:endpoint|intune)\.microsoft\.com$/i,
  },
  {
    id: "intune-checkin",
    area: "Intune",
    feature: "Device check-in and policies",
    severity: "critical",
    impact: "Devices stop syncing with Intune: no new policies, apps, compliance status or remote actions.",
    pattern: /(?:^|\.)manage\.microsoft\.com$|(?:^|\.)dm\.microsoft\.com$/i,
  },

  // --- Microsoft Defender ---------------------------------------------------
  {
    id: "defender-signatures",
    area: "Microsoft Defender",
    feature: "Signature updates",
    severity: "critical",
    impact: "Antivirus signatures and engine updates are not downloaded; protection gets outdated.",
    pattern: /(?:^|\.)definitionupdates\.microsoft\.com$/i,
  },
  {
    id: "defender-cloud",
    area: "Microsoft Defender",
    feature: "Cloud protection (MAPS)",
    severity: "critical",
    impact: "Cloud-delivered protection, block at first sight and sample submission stop working.",
    pattern: /(?:^|\.)wdcp\.microsoft\.com$|(?:^|\.)wdcpalt\.microsoft\.com$|(?:^|\.)cp\.wd\.microsoft\.com$|(?:^|\.)wd\.microsoft\.com$/i,
  },
  {
    id: "defender-edr",
    area: "Microsoft Defender",
    feature: "Defender for Endpoint sensor and portal",
    severity: "degraded",
    impact: "Onboarding, EDR telemetry, live response and the Defender portal are affected.",
    pattern: /(?:^|\.)(?:security|defender)\.microsoft\.com$|^winatp-gw-[a-z0-9-]+\.microsoft\.com$|(?:^|\.)endpoint\.security\.microsoft\.com$/i,
  },
  {
    id: "smartscreen",
    area: "Microsoft Defender",
    feature: "SmartScreen",
    severity: "degraded",
    impact: "SmartScreen reputation checks for downloads, apps and websites fail.",
    pattern: /smartscreen|(?:^|\.)urs\.microsoft\.com$/i,
  },

  // --- Microsoft 365: Exchange ---------------------------------------------
  {
    id: "exchange-protection",
    area: "Exchange Online",
    feature: "Exchange Online Protection",
    severity: "degraded",
    impact: "Mail flow through Exchange Online Protection and Safe Links/Attachments checks are affected.",
    pattern: /(?:^|\.)protection\.outlook\.com$/i,
  },
  {
    id: "exchange",
    area: "Exchange Online",
    feature: "Mail, calendar and Autodiscover",
    severity: "critical",
    impact: "Outlook cannot connect: no mail, calendar or contacts, Autodiscover fails.",
    pattern: /(?:^|\.)outlook\.(?:office|office365|com|cloud\.microsoft)(?:\.com)?$|^autodiscover\.|(?:^|\.)smtp\.office365\.com$|(?:^|\.)outlook\.office365\.com$/i,
  },

  // --- Microsoft 365: Teams -------------------------------------------------
  {
    id: "office-config",
    area: "Microsoft 365 Apps",
    feature: "Feature configuration (ECS)",
    severity: "degraded",
    impact: "Office, Teams and Edge cannot load feature flags and configuration; some features stay off.",
    pattern: /(?:^|\.)config\.edge\.skype\.com$|(?:^|\.)ecs\.office\.com$/i,
  },
  {
    id: "teams",
    area: "Teams",
    feature: "Chat, meetings and calls",
    severity: "critical",
    impact: "Teams chat, meetings, calls and media are disrupted.",
    pattern: /(?:^|\.)teams\.(?:microsoft\.com|cloud\.microsoft|live\.com)$|(?:^|\.)skype\.com$|(?:^|\.)lync\.com$|(?:^|\.)teams\.cdn\.office\.net$|(?:^|\.)skypeforbusiness\.com$/i,
  },

  // --- Microsoft 365: SharePoint / OneDrive --------------------------------
  {
    id: "onedrive-client",
    area: "SharePoint & OneDrive",
    feature: "OneDrive sync client updates",
    severity: "low",
    impact: "The OneDrive sync app cannot update itself; sync keeps working.",
    pattern: /(?:^|\.)oneclient\.sfx\.ms$|(?:^|\.)g\.live\.com$/i,
  },
  {
    id: "sharepoint",
    area: "SharePoint & OneDrive",
    feature: "Files, sync and sharing",
    severity: "critical",
    impact: "SharePoint sites and OneDrive files cannot be opened or synced; Teams files are affected too.",
    pattern: /(?:^|\.)sharepoint\.com$|(?:^|\.)sharepointonline\.com$|(?:^|\.)onedrive\.(?:com|live\.com)$|(?:^|\.)svc\.ms$|(?:^|\.)storage\.live\.com$/i,
  },

  // --- Microsoft 365: Apps and portals --------------------------------------
  {
    id: "m365-apps-install",
    area: "Microsoft 365 Apps",
    feature: "Install and updates",
    severity: "degraded",
    impact: "Microsoft 365 Apps cannot be installed or updated (Click-to-Run).",
    pattern: /(?:^|\.)officecdn\.microsoft\.com(?:\.edgesuite\.net)?$|(?:^|\.)officecdn-microsoft-com\.akamaized\.net$/i,
  },
  {
    id: "m365-apps-licensing",
    area: "Microsoft 365 Apps",
    feature: "Licensing and cloud policy",
    severity: "critical",
    impact: "Office apps cannot activate or refresh their license and switch to reduced functionality; Cloud Policy settings do not apply.",
    pattern: /(?:^|\.)officeclient\.microsoft\.com$|(?:^|\.)config\.office\.com$|(?:^|\.)config\.office\.net$|(?:^|\.)ols\.officeapps\.live\.com$/i,
  },
  {
    id: "m365-admin",
    area: "Microsoft 365",
    feature: "Admin and compliance portals",
    severity: "low",
    impact: "Microsoft 365 admin, Purview and compliance portals cannot be opened. End users are not affected.",
    pattern: /(?:^|\.)admin\.microsoft\.com$|(?:^|\.)protection\.office\.com$|(?:^|\.)compliance\.microsoft\.com$|(?:^|\.)purview\.microsoft\.com$/i,
  },
  {
    id: "m365-telemetry",
    area: "Microsoft 365 Apps",
    feature: "Diagnostics",
    severity: "low",
    impact: "Office diagnostic data is not sent. Users will not notice.",
    pattern: /(?:^|\.)nexus(?:rules)?\.officeapps\.live\.com$|(?:^|\.)office\.events\.data\.microsoft\.com$|(?:^|\.)aria\.microsoft\.com$|(?:^|\.)otelrules\.azureedge\.net$/i,
  },
  {
    id: "purview-mip",
    area: "Microsoft Purview",
    feature: "Sensitivity labels and encryption",
    severity: "degraded",
    impact: "Sensitivity labels cannot be applied and protected (encrypted) documents and mails cannot be opened.",
    pattern: /informationprotection/i,
  },
  {
    id: "power-platform",
    area: "Power Platform",
    feature: "Power Automate and connectors",
    severity: "degraded",
    impact: "Power Automate flows and Power Platform connectors used in Microsoft 365 fail.",
    pattern: /(?:^|\.)flow\.microsoft\.com$|(?:^|\.)azure-apim\.net$|(?:^|\.)azurerms\.com$|(?:^|\.)powerautomate\.com$/i,
  },
  {
    id: "m365-media",
    area: "Microsoft 365",
    feature: "Stream and meeting recordings",
    severity: "low",
    impact: "Videos and meeting recordings may not play.",
    pattern: /(?:^|\.)mediaservices\.windows\.net$|(?:^|\.)amp\.azure\.net$/i,
  },
  {
    id: "m365-dictation",
    area: "Microsoft 365 Apps",
    feature: "Dictation and speech",
    severity: "low",
    impact: "Dictation and read-aloud features in Office apps do not work.",
    pattern: /(?:^|\.)officespeech\.platform\.bing\.com$/i,
  },
  {
    id: "m365-redirects",
    area: "Microsoft 365 Apps",
    feature: "Add-ins, help and redirects",
    severity: "low",
    impact: "Office add-ins store, in-app help and some links inside Office apps do not open.",
    pattern: /officeredir|officepreviewredir|(?:^|\.)appsforoffice\.microsoft\.com$|(?:^|\.)office15client\.microsoft\.com$|^r\.office\.microsoft\.com$|^c1\.microsoft\.com$|^office\.microsoft$/i,
  },
  {
    id: "m365-web",
    area: "Microsoft 365",
    feature: "Web apps and portal",
    severity: "degraded",
    impact: "office.com, Office for the web and parts of the desktop apps (add-ins, templates, sharing) fail.",
    pattern: /(?:^|\.)(?:office|office365|office\.net|officeapps\.live\.com|online\.office\.com|cloud\.microsoft|static\.microsoft|usercontent\.microsoft|microsoft365\.com|m365\.cloud\.microsoft)(?:\.com)?$|(?:^|\.)portal\.office\.com$|(?:^|\.)msocdn\.com$|(?:^|\.)wns\.windows\.com$|(?:^|\.)mx\.microsoft$|(?:^|\.)onmicrosoft\.com$/i,
  },

  // --- Microsoft Edge -------------------------------------------------------
  {
    id: "edge-extensions",
    area: "Microsoft Edge",
    feature: "Extensions",
    severity: "low",
    impact: "Edge extensions cannot be installed or updated from the add-ons store.",
    pattern: /msedgeextensions/i,
  },
  {
    id: "edge-activity",
    area: "Microsoft Edge",
    feature: "Enterprise sync and activity",
    severity: "low",
    impact: "Edge enterprise sync and activity features are limited.",
    pattern: /edge-enterprise\.activity\.windows\.com$/i,
  },
  {
    id: "edge-misc",
    area: "Microsoft Edge",
    feature: "Optional browser features",
    severity: "low",
    impact: "Optional Edge features such as rewards, privacy dashboard or passkey sync are unavailable.",
    pattern: /(?:^|\.)(?:edgepasskeysenclave|privacy|rewardsplatform|prod\.rewardsplatform)\.microsoft\.com$/i,
  },
  {
    id: "edge-updates",
    area: "Microsoft Edge",
    feature: "Updates and services",
    severity: "degraded",
    impact: "Edge cannot update itself or load online services.",
    pattern: /(?:^|\.)edge\.microsoft\.com$|(?:^|\.)msedge\.(?:api\.cdp\.microsoft\.com|net)$|(?:^|\.)api\.cdp\.microsoft\.com$|^msedge\.[a-z0-9.]*dl\.delivery\.mp\.microsoft\.com$/i,
  },

  // --- Windows Update / Store / Activation ----------------------------------
  {
    id: "delivery-optimization",
    area: "Windows Update",
    feature: "Delivery Optimization",
    severity: "low",
    impact: "Peer caching is not available; devices download updates directly, which uses more bandwidth.",
    pattern: /(?:^|\.)do\.dsp\.mp\.microsoft\.com$|(?:^|\.)dsp\.mp\.microsoft\.com$/i,
  },
  {
    id: "store-catalog",
    area: "Microsoft Store",
    feature: "Catalog, purchase and licensing",
    severity: "degraded",
    impact: "Store apps, including Store apps deployed by Intune, cannot be installed, updated or licensed.",
    pattern: /(?:^|\.)(?:displaycatalog|purchase\.md|licensing|storeedgefd\.dsx)\.mp\.microsoft\.com$/i,
  },
  {
    id: "store-media",
    area: "Microsoft Store",
    feature: "Images and live tiles",
    severity: "low",
    impact: "Store images, tiles and revocation lists do not load; installs still work.",
    pattern: /(?:^|\.)livetileedge\.dsx\.mp\.microsoft\.com$|(?:^|\.)store-images\.s-microsoft\.com$|akamaized\.net$|(?:^|\.)storequality\.microsoft\.com$/i,
  },
  {
    id: "wu-scan",
    area: "Windows Update",
    feature: "Update scan and approval",
    severity: "critical",
    impact: "Devices cannot check for updates. No security, feature or driver updates are offered.",
    pattern: /(?:^|\.)fe[0-9]\.(?:ws\.)?(?:delivery\.mp\.microsoft\.com)$|(?:^|\.)update\.microsoft\.com$|(?:^|\.)windowsupdate\.microsoft\.com$/i,
  },
  {
    id: "wu-download",
    area: "Windows Update",
    feature: "Update downloads",
    severity: "critical",
    impact: "Update packages, feature updates, drivers and Store app packages cannot be downloaded.",
    pattern: /(?:^|\.)delivery\.mp\.microsoft\.com$|(?:^|\.)windowsupdate\.com$/i,
  },
  {
    id: "windows-activation",
    area: "Windows Activation",
    feature: "Activation and license validation",
    severity: "degraded",
    impact: "Windows cannot activate or validate its license (MAK, retail, subscription activation).",
    pattern: /(?:^|\.)(?:activation|activation-v2|validation|validation-v2)\.sls\.microsoft\.com$/i,
  },
  {
    id: "crl",
    area: "Certificates",
    feature: "Revocation checks (CRL/OCSP)",
    severity: "degraded",
    impact: "TLS certificate revocation checks fail or time out. Sign-in, activation and downloads can hang.",
    pattern: /(?:^|\.)crl\.microsoft\.com$|(?:^|\.)ocsp\.|(?:^|\.)oneocsp\.microsoft\.com$|(?:^|\.)mscrl\.microsoft\.com$/i,
  },

  // --- Telemetry ------------------------------------------------------------
  {
    id: "windows-telemetry",
    area: "Windows diagnostics",
    feature: "Diagnostic data and crash reports",
    severity: "low",
    impact: "Diagnostic data and crash reports are not sent. Windows Update reports and Endpoint analytics may be incomplete.",
    pattern: /(?:^|\.)events\.data\.microsoft\.com$|(?:^|\.)vortex[a-z0-9-]*\.data\.microsoft\.com$|(?:^|\.)settings-win\.data\.microsoft\.com$|watson|watcab|(?:^|\.)oca\.(?:telemetry\.)?microsoft\.com$/i,
  },

  // --- Generic fallbacks ----------------------------------------------------
  {
    id: "microsoft-web",
    area: "Microsoft web",
    feature: "Links and redirects",
    severity: "low",
    impact: "Help links and redirects from Microsoft products do not open.",
    pattern: /^(?:go|www|aka)\.microsoft\.com$|^aka\.ms$/i,
  },
  {
    id: "microsoft-account",
    area: "Microsoft account",
    feature: "Consumer services",
    severity: "low",
    impact: "Personal Microsoft account features (consumer sign-in, settings sync) are limited.",
    pattern: /(?:^|\.)live\.com$/i,
  },
];

const M365_SERVICE_AREAS = {
  Exchange: "Exchange Online",
  SharePoint: "SharePoint & OneDrive",
  Skype: "Teams",
  Common: "Microsoft 365",
};

function fallbackRule(serviceName) {
  return {
    id: `other:${serviceName ?? "unknown"}`,
    area: serviceName || "Other",
    feature: "Other endpoints",
    severity: "degraded",
    impact: `Parts of ${serviceName || "the service"} may not work.`,
  };
}

function capSeverity(severity, maxSeverity) {
  return SEVERITY_ORDER[severity] < SEVERITY_ORDER[maxSeverity] ? maxSeverity : severity;
}

/**
 * Returns the feature an endpoint belongs to.
 * `m365` is optional metadata from the Microsoft 365 endpoint web service
 * ({ serviceArea, category }); when present it refines the severity:
 * Optimize/Allow endpoints are required for the service to work, Default ones are not.
 */
export function classifyEndpoint(host, serviceName, m365 = null) {
  const normalized = String(host ?? "").trim().toLowerCase().replace(/^\*\./, "");
  const rule = IMPACT_RULES.find((entry) => entry.pattern.test(normalized));
  const base = rule ? { ...rule } : fallbackRule(serviceName);
  delete base.pattern;

  if (m365?.category) {
    const category = String(m365.category);
    if (!rule && m365.serviceArea && M365_SERVICE_AREAS[m365.serviceArea]) {
      base.area = M365_SERVICE_AREAS[m365.serviceArea];
      base.id = `m365:${m365.serviceArea}`;
    }
    if (category === "Optimize" || category === "Allow") {
      base.severity = "critical";
    } else if (category === "Default") {
      base.severity = capSeverity(base.severity, "degraded");
    }
    base.m365Category = category;
  }
  return base;
}

function isFailed(status) {
  return status === "unreachable" || status === "timeout";
}

function isTested(status) {
  return status === "reachable" || isFailed(status);
}

/**
 * Groups check results by feature and returns one entry per affected feature.
 * rows: [{ host, serviceName, status, m365? }]
 * A host that appears in several services is counted once per feature.
 */
export function summarizeImpacts(rows) {
  const groups = new Map();
  for (const row of rows) {
    if (!isTested(row.status)) continue;
    const feature = row.impact ?? classifyEndpoint(row.host, row.serviceName, row.m365);
    const key = `${feature.area}::${feature.feature}`;
    if (!groups.has(key)) {
      groups.set(key, {
        key,
        area: feature.area,
        feature: feature.feature,
        severity: feature.severity,
        impact: feature.impact,
        hosts: new Map(),
        services: new Set(),
      });
    }
    const group = groups.get(key);
    if (SEVERITY_ORDER[feature.severity] < SEVERITY_ORDER[group.severity]) {
      group.severity = feature.severity;
    }
    const hostKey = String(row.host).toLowerCase();
    const previous = group.hosts.get(hostKey);
    // If the same host was tested twice, a failure wins over success.
    if (!previous || isFailed(row.status)) {
      group.hosts.set(hostKey, row.status);
    }
    if (row.serviceName) group.services.add(row.serviceName);
  }

  const result = [];
  for (const group of groups.values()) {
    const failedHosts = [...group.hosts].filter(([, status]) => isFailed(status)).map(([host]) => host).sort();
    if (failedHosts.length === 0) continue;
    const total = group.hosts.size;
    result.push({
      area: group.area,
      feature: group.feature,
      severity: group.severity,
      impact: group.impact,
      failed: failedHosts.length,
      total,
      allFailed: failedHosts.length === total,
      failedHosts,
      services: [...group.services].sort(),
    });
  }

  result.sort(
    (a, b) =>
      SEVERITY_ORDER[a.severity] - SEVERITY_ORDER[b.severity] ||
      Number(b.allFailed) - Number(a.allFailed) ||
      b.failed - a.failed ||
      a.area.localeCompare(b.area) ||
      a.feature.localeCompare(b.feature),
  );
  return result;
}

/** One-line wording of how likely the impact is. */
export function likelihoodText(entry) {
  if (entry.allFailed) return "likely affected";
  if (entry.failed / entry.total >= 0.5) return "probably affected";
  return "possibly affected";
}
