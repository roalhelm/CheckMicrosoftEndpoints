
[CmdletBinding()]
param(
    [Parameter(HelpMessage="Specify which services to test. Use 'All' for all services or choose specific ones.")]
    [ValidateSet('All', 'WindowsUpdate', 'Autopatch', 'Intune', 'Defender', 'AzureAD', 'Microsoft365', 'Store', 'Activation', 'Edge', 'Telemetry', 'Interactive')]
    [string[]]$Services = 'Interactive',
    
    [Parameter(HelpMessage="Skip ping latency tests to speed up execution")]
    [switch]$SkipPing,
    
    [Parameter(HelpMessage="Skip download speed tests to speed up execution")]
    [switch]$SkipSpeed,
    
    [Parameter(HelpMessage="Run in quiet mode with minimal output")]
    [switch]$Quiet,
    
    [Parameter(HelpMessage="Generate HTML report and save to specified path")]
    [string]$HtmlReport,
    
    [Parameter(HelpMessage="Open HTML report in browser after generation")]
    [switch]$OpenReport,

    [Parameter(HelpMessage="Do not query the Microsoft 365 endpoint web service for endpoint categories")]
    [switch]$SkipM365Lookup
)

<#
.SYNOPSIS
    Comprehensive connectivity, latency, and performance test for selectable Microsoft cloud service endpoints including Autopatch, Intune, Defender, and Microsoft 365.

    GitHub Repository: https://github.com/roalhelm/PowershellScripts

.DESCRIPTION
    This extensive connectivity testing script validates network access to critical Microsoft service endpoints
    required for modern Windows environments and Microsoft cloud services. It covers all major services including
    Windows Update for Business, Windows Autopatch, Microsoft Intune, Microsoft Defender, Azure AD, Microsoft 365,
    Microsoft Store, Windows Activation, Microsoft Edge, and Windows Telemetry services.
    
    The script performs the following operations:
    1. Allows selective testing of specific Microsoft services or all services
    2. Tests TCP connectivity to essential Microsoft service URLs on port 443 (HTTPS)
    3. Measures ping latency to all reachable endpoints for performance analysis (optional)
    4. Tests download speed/response time to evaluate network quality (optional)
    5. Groups results by service backend for organized reporting
    6. Provides color-coded output for easy identification of connectivity issues
    7. Reports IP addresses of reachable endpoints for network troubleshooting
    8. Displays comprehensive performance statistics and network quality ratings
    9. Returns appropriate exit codes for automation and monitoring scenarios
    10. Interactive menu system for easy service selection
    
    Tested Service Categories:
    - Windows Update for Business (WUfB)
    - Windows Autopatch (automated patch management)
    - Microsoft Intune (device management)
    - Microsoft Defender (security services)
    - Azure Active Directory (identity services)
    - Microsoft 365 (productivity suite)
    - Microsoft Store (app distribution)
    - Windows Activation (licensing services)
    - Microsoft Edge (browser services)
    - Windows Telemetry (diagnostic data)
    
    This tool is particularly useful for:
    - Network administrators validating comprehensive firewall configurations
    - Enterprise deployment teams ensuring all Microsoft services are accessible
    - Troubleshooting connectivity issues across multiple Microsoft platforms
    - Compliance verification for complete Microsoft cloud service access
    - Pre-deployment network validation in complex enterprise environments
    - Performance monitoring and network quality assessment
    - Identifying slow connections or network bottlenecks to Microsoft services
    - Targeted testing of specific Microsoft services to save time
    - Automated testing scenarios with service-specific focus

.PARAMETER Services
    Specifies which Microsoft services to test. Options:
    - 'All': Test all available services
    - 'WindowsUpdate': Windows Update for Business endpoints
    - 'Autopatch': Windows Autopatch endpoints  
    - 'Intune': Microsoft Intune device management endpoints
    - 'Defender': Microsoft Defender security endpoints
    - 'AzureAD': Azure Active Directory authentication endpoints
    - 'Microsoft365': Microsoft 365 productivity suite endpoints
    - 'Store': Microsoft Store app distribution endpoints
    - 'Activation': Windows activation and licensing endpoints
    - 'Edge': Microsoft Edge browser service endpoints
    - 'Telemetry': Windows diagnostic and telemetry endpoints
    - 'Interactive': Show interactive menu to select services (default)

.PARAMETER SkipPing
    Skips ping latency testing to reduce execution time. Only basic connectivity will be tested.

.PARAMETER SkipSpeed
    Skips download speed/response time testing to reduce execution time.

.PARAMETER Quiet
    Runs in quiet mode with minimal console output. Useful for automated scenarios.

.PARAMETER HtmlReport
    Generates a comprehensive HTML report and saves it to the specified path. 
    If no extension is provided, '.html' will be added automatically.
    If no path is provided, saves to current directory with timestamp.

.PARAMETER OpenReport
    Automatically opens the generated HTML report in the default web browser.
    Only works when -HtmlReport parameter is also specified.

.PARAMETER SkipM365Lookup
    Skips the query to the official Microsoft 365 endpoint web service (endpoints.office.com).
    The script uses that service to rate Microsoft 365 endpoints: Optimize and Allow endpoints
    count as critical. Without it, the built-in rules are used.

.NOTES
    File Name     : CheckMicrosoftEndpointsV2.ps1
    Author        : Ronny Alhelm
    Version       : 2.2
    Creation Date : October 13, 2025
    Last Modified : October 14, 2025
    Requirements  : PowerShell 5.1 or higher
    Network       : Internet connectivity required
    Ports         : HTTPS (443) access to Microsoft endpoints, ICMP for ping tests
    Permissions   : No elevated privileges required
    Duration      : 1-10 minutes depending on selected services and test options

.CHANGES
    Version 2.2 (2026-11-17):
    - Each endpoint is mapped to the service feature it belongs to (e.g. Intune check-in, Teams calls,
      Windows Update downloads) with a severity (critical, degraded, low)
    - New "Likely service impact" summary in console and HTML report: which features are affected,
      how many of their endpoints failed and what users will notice
    - Microsoft 365 endpoints are rated with the categories from the official endpoint web service

    Version 2.1 (2025-10-14):
    - Added comprehensive HTML report generation with modern responsive design
    - Implemented automatic browser opening for generated reports
    - Enhanced visual presentation with color-coded status indicators
    - Added performance statistics and service impact analysis in HTML format
    - Improved report accessibility and professional appearance
    
    Version 2.0 (2025-10-14):
    - Added selective service testing with parameter support
    - Implemented interactive menu system for service selection
    - Added options to skip ping and speed tests for faster execution
    - Added quiet mode for automated scenarios
    - Enhanced parameter validation and help documentation
    - Improved user experience with service-specific testing options

    Version 1.1 (2025-10-14):
    - Added comprehensive ping latency testing for all reachable endpoints
    - Implemented download speed/response time measurements
    - Enhanced performance statistics with min/max/average calculations
    - Added network quality ratings and performance recommendations
    - Improved progress indicators during extended testing phases
    - Extended result tables to include latency and speed metrics
    - Added color-coded performance indicators for quick assessment

    Version 1.0 (2025-10-13):
    - Initial release with comprehensive endpoint testing
    - Added support for Windows Update for Business endpoints
    - Included Microsoft Intune service endpoints
    - Added Microsoft Defender endpoint validation
    - Implemented color-coded output for better readability
    - Added grouped results display by service backend
    - Implemented proper exit codes for automation
    - Added IP address resolution for troubleshooting

.VERSION
    2.1

.EXAMPLE
    .\CheckMicrosoftEndpointsV2.ps1
    Runs in interactive mode, showing a menu to select which services to test.

.EXAMPLE
    .\CheckMicrosoftEndpointsV2.ps1 -Services All
    Tests all available Microsoft services with full connectivity, ping, and speed tests.

.EXAMPLE
    .\CheckMicrosoftEndpointsV2.ps1 -Services Intune,Defender -SkipSpeed
    Tests only Intune and Microsoft Defender endpoints with ping tests but skips speed tests.

.EXAMPLE
    .\CheckMicrosoftEndpointsV2.ps1 -Services WindowsUpdate,Microsoft365 -Quiet
    Tests Windows Update and Microsoft 365 endpoints in quiet mode with minimal output.

.EXAMPLE
    .\CheckMicrosoftEndpointsV2.ps1 -Services AzureAD -SkipPing -SkipSpeed
    Tests only Azure AD endpoints with basic connectivity testing only (fastest execution).

.EXAMPLE
    # Use in automated monitoring for specific services
    .\CheckMicrosoftEndpointsV2.ps1 -Services Intune,Defender -Quiet
    if ($LASTEXITCODE -eq 0) {
        Write-Host "Intune and Defender endpoints accessible"
    } else {
        Write-Host "Connectivity issues detected" -ForegroundColor Red
    }

.EXAMPLE
    # Test only critical services for quick validation
    .\CheckMicrosoftEndpointsV2.ps1 -Services WindowsUpdate,AzureAD,Intune -SkipSpeed

.EXAMPLE
    # Generate HTML report and open in browser
    .\CheckMicrosoftEndpointsV2.ps1 -Services All -HtmlReport "NetworkReport.html" -OpenReport

.EXAMPLE
    # Create HTML report with timestamp in filename
    .\CheckMicrosoftEndpointsV2.ps1 -Services Intune,Defender -HtmlReport "Microsoft-Endpoints-Report-$(Get-Date -Format 'yyyyMMdd-HHmmss').html" -Quiet

.EXAMPLE
    # Full test with comprehensive HTML report
    .\CheckMicrosoftEndpointsV2.ps1 -Services All -HtmlReport "C:\Reports\Microsoft-Endpoints.html"

#>



# Function to show interactive service selection menu
function Show-ServiceSelectionMenu {
    $selectedServices = @()
    
        Write-Host "`nMICROSOFT ENDPOINT CONNECTIVITY TESTER V2.1" -ForegroundColor Cyan
    Write-Host "="*60 -ForegroundColor DarkCyan
    Write-Host "Select the Microsoft Services to test:" -ForegroundColor Yellow
    Write-Host ""
    
    $menuOptions = @(
        @{ Key = "1"; Name = "Windows Update for Business"; Param = "WindowsUpdate" }
        @{ Key = "2"; Name = "Windows Autopatch"; Param = "Autopatch" }
        @{ Key = "3"; Name = "Microsoft Intune"; Param = "Intune" }
        @{ Key = "4"; Name = "Microsoft Defender"; Param = "Defender" }
        @{ Key = "5"; Name = "Azure Active Directory"; Param = "AzureAD" }
        @{ Key = "6"; Name = "Microsoft 365"; Param = "Microsoft365" }
        @{ Key = "7"; Name = "Microsoft Store"; Param = "Store" }
        @{ Key = "8"; Name = "Windows Activation"; Param = "Activation" }
        @{ Key = "9"; Name = "Microsoft Edge"; Param = "Edge" }
        @{ Key = "10"; Name = "Windows Telemetry"; Param = "Telemetry" }
        @{ Key = "A"; Name = "ALL SERVICES"; Param = "All" }
    )
    
    foreach ($option in $menuOptions) {
        if ($option.Key -eq "A") {
            Write-Host ""
            Write-Host "[$($option.Key)]  $($option.Name) " -ForegroundColor Green
        } else {
            Write-Host "[$($option.Key)]   $($option.Name)" -ForegroundColor White
        }
    }
    
    Write-Host ""
    Write-Host "Enter the numbers of the desired services (e.g. 1,3,5 or A for all):" -ForegroundColor Cyan
    Write-Host "Additional options:" -ForegroundColor Yellow
    Write-Host "  - Press 'P' to skip ping tests" -ForegroundColor Gray
    Write-Host "  - Press 'S' to skip speed tests" -ForegroundColor Gray
    Write-Host "  - Press 'H' to generate HTML report" -ForegroundColor Gray
    Write-Host "  - Press 'Q' for quiet mode" -ForegroundColor Gray
    
    $userInput = Read-Host "`nYour selection"
    
    # Parse input for service selection
    $selections = $userInput -split ',' | ForEach-Object { $_.Trim().ToUpper() }
    
    foreach ($selection in $selections) {
        $option = $menuOptions | Where-Object { $_.Key -eq $selection }
        if ($option) {
            if ($option.Param -eq "All") {
                return @("All")
            } else {
                $selectedServices += $option.Param
            }
        }
    }
    
    if ($selectedServices.Count -eq 0) {
        Write-Host "No valid services selected. Using all services." -ForegroundColor Yellow
        return @("All")
    }
    
    return $selectedServices
}

# Handle interactive mode
if ($Services -contains 'Interactive') {
    $Services = Show-ServiceSelectionMenu
    
    # Ask for additional options
    Write-Host ""
    $skipOptions = Read-Host "Options: Skip (P)ing, Skip (S)peed, Generate (H)TML-Report, (Q)uiet, or Enter for default"
    
    if ($skipOptions -match 'P') { $SkipPing = $true }
    if ($skipOptions -match 'S') { $SkipSpeed = $true }
    if ($skipOptions -match 'Q') { $Quiet = $true }
    if ($skipOptions -match 'H') { 
        $HtmlReport = "Microsoft-Endpoints-Report-$(Get-Date -Format 'yyyyMMdd-HHmmss').html"
        $htmlChoice = Read-Host "Open HTML report in browser? (y/n)"
        if ($htmlChoice -match '^(y|yes)') { $OpenReport = $true }
    }
    
    Clear-Host
}

# List of required Microsoft service endpoints (reviewed 2026)
$backendUrls = @{
    'Windows Update for Business' = @(
        'https://dl.delivery.mp.microsoft.com',
        'https://download.windowsupdate.com',
        'https://fe3.delivery.mp.microsoft.com',
        'https://sws.update.microsoft.com',
        'https://update.microsoft.com',
        'https://windowsupdate.microsoft.com',
        'https://www.update.microsoft.com',
        'https://*.delivery.mp.microsoft.com',
        'https://*.dl.delivery.mp.microsoft.com',
        'https://*.update.microsoft.com',
        'https://*.windowsupdate.com',
        'https://*.windowsupdate.microsoft.com'
    ) | Sort-Object -Unique
    
    'Windows Autopatch' = @(
<#        'https://api.update.microsoft.com',
        'https://autopatch.microsoft.com',
        'https://config.edge.skype.com',
        'https://graph.microsoft.com',
        'https://login.microsoftonline.com',
        'https://manage.microsoft.com',
        'https://nexusrules.officeapps.live.com'
#>
#########################################################################
        'https://mmdcustomer.microsoft.com',
        'https://mmdls.microsoft.com',
        'https://devicelistenerprod.microsoft.com',
        'https://login.windows.net',
        'https://device.autopatch.microsoft.com',
        'https://services.autopatch.microsoft.com',
        'https://payloadprod*.blob.core.windows.net',
        'https://*.webpubsub.azure.com'
#########################################################################
    ) | Sort-Object -Unique
    
    'Intune' = @(
        'https://device.login.microsoftonline.com',
        'https://endpoint.microsoft.com',
        'https://graph.microsoft.com',
        'https://login.microsoftonline.com',
        'https://manage.microsoft.com',
        'https://*.manage.microsoft.com',
        'https://*.dm.microsoft.com',
        'https://portal.manage.microsoft.com',
        'https://enrollment.manage.microsoft.com',
        'https://enterpriseregistration.windows.net',
        'https://certauth.enterpriseregistration.windows.net',
        'https://r.manage.microsoft.com',
        'https://i.manage.microsoft.com',
        'https://p.manage.microsoft.com',
        'https://c.manage.microsoft.com',
        'https://m.manage.microsoft.com'
    ) | Sort-Object -Unique
    
    'Microsoft Defender' = @(
        'https://*.defender.microsoft.com',
        'https://*.protection.outlook.com',
        'https://*.security.microsoft.com',
        'https://*.wdcp.microsoft.com',
        'https://go.microsoft.com',
        'https://wdcp.microsoft.com',
        'https://definitionupdates.microsoft.com',
        'https://www.microsoft.com',
        'https://unitedstates.cp.wd.microsoft.com',
        'https://europe.cp.wd.microsoft.com',
        'https://asia.cp.wd.microsoft.com'
    ) | Sort-Object -Unique
    
    'Azure Active Directory' = @(
        'https://login.microsoftonline.com',
        'https://device.login.microsoftonline.com',
        'https://enterpriseregistration.windows.net',
        'https://pas.windows.net',
        'https://management.azure.com',
        'https://aadcdn.msauth.net',
        'https://aadcdn.msftauth.net'
    ) | Sort-Object -Unique
    
    'Microsoft 365' = @(
        'https://admin.microsoft.com',
        'https://config.office.com',
        'https://graph.microsoft.com',
        'https://login.microsoftonline.com',
        'https://officecdn.microsoft.com',
        'https://protection.office.com',
        'https://portal.office.com',
        'https://*.cloud.microsoft',
        'https://*.static.microsoft',
        'https://*.usercontent.microsoft',
        'https://*.office.com',
        'https://*.office365.com',
        'https://*.sharepoint.com',
        'https://*.onedrive.com'
    ) | Sort-Object -Unique
    
    'Microsoft Store' = @(
        'https://storeedgefd.dsx.mp.microsoft.com',
        'https://livetileedge.dsx.mp.microsoft.com',
        'https://storecatalogrevocation.storequality.microsoft.com',
        'https://img-prod-cms-rt-microsoft-com.akamaized.net',
        'https://store-images.s-microsoft.com',
        'https://displaycatalog.mp.microsoft.com',
        'https://licensing.mp.microsoft.com',
        'https://purchase.md.mp.microsoft.com'
    ) | Sort-Object -Unique
    
    'Windows Activation' = @(
        'https://activation.sls.microsoft.com',
        'https://crl.microsoft.com',
        'https://validation.sls.microsoft.com',
        'https://activation-v2.sls.microsoft.com',
        'https://purchase.md.mp.microsoft.com',
        'https://licensing.mp.microsoft.com'
    ) | Sort-Object -Unique
    
    'Microsoft Edge' = @(
        'https://config.edge.skype.com',
        'https://edge.microsoft.com',
        'https://msedge.api.cdp.microsoft.com',
        'https://dual-s-01.dual.dualstack.edge-enterprise.activity.windows.com',
        'https://nav.smartscreen.microsoft.com',
        'https://unitedstates.smartscreen-prod.microsoft.com'
    ) | Sort-Object -Unique
    
    'Windows Telemetry' = @(
        'https://v10c.events.data.microsoft.com',
        'https://v20.events.data.microsoft.com',
        'https://watson.telemetry.microsoft.com',
        'https://umwatsonc.events.data.microsoft.com',
        'https://ceuswatcab01.blob.core.windows.net',
        'https://ceuswatcab02.blob.core.windows.net',
        'https://eaus2watcab01.blob.core.windows.net',
        'https://eaus2watcab02.blob.core.windows.net',
        'https://weus2watcab01.blob.core.windows.net',
        'https://weus2watcab02.blob.core.windows.net'
    ) | Sort-Object -Unique
}

# Service impact descriptions for failed connectivity
$serviceImpacts = @{
    'Windows Update for Business' = @{
        'Impact' = 'Windows Updates and Feature Updates cannot be downloaded'
        'Symptoms' = 'Update errors, outdated security patches, missing feature updates'
    }
    'Windows Autopatch' = @{
        'Impact' = 'Automatic patch management does not work'
        'Symptoms' = 'Manual update management required, no automatic deployment schedules'
    }
    'Intune' = @{
        'Impact' = 'Device enrollment, app deployment, and compliance checks do not work'
        'Symptoms' = 'New devices cannot be enrolled, apps do not install, compliance reports missing'
    }
    'Microsoft Defender' = @{
        'Impact' = 'Antivirus updates and cloud protection do not work'
        'Symptoms' = 'Outdated virus signatures, no cloud-based threat analysis, limited protection'
    }
    'Azure Active Directory' = @{
        'Impact' = 'Sign-in and device authentication may fail'
        'Symptoms' = 'Login problems, device registration failed, SSO does not work'
    }
    'Microsoft 365' = @{
        'Impact' = 'Office apps, SharePoint, and OneDrive do not work properly'
        'Symptoms' = 'Office apps do not start, documents do not sync, Admin Center not accessible'
    }
    'Microsoft Store' = @{
        'Impact' = 'App installation and updates from Microsoft Store do not work'
        'Symptoms' = 'Store does not open, apps cannot be installed/updated'
    }
    'Windows Activation' = @{
        'Impact' = 'Windows activation and license validation do not work'
        'Symptoms' = 'Activation errors, license warnings, limited functionality'
    }
    'Microsoft Edge' = @{
        'Impact' = 'Browser updates and enterprise features work with limitations'
        'Symptoms' = 'SmartScreen disabled, no automatic updates, enterprise policies do not work'
    }
    'Windows Telemetry' = @{
        'Impact' = 'Diagnostic data and error reports are not transmitted'
        'Symptoms' = 'No Windows diagnostic data, error reports lost, update quality may be impaired'
    }
}

# BEGIN GENERATED IMPACT RULES
# Generated by tools/sync-impact-rules.mjs from web/service-impact.mjs. Do not edit by hand.
# First match wins. Severity: critical = core function stops, degraded = feature missing, low = optional/telemetry.
$endpointImpactRules = @(
    [PSCustomObject]@{ Id = 'entra-certauth'; Area = 'Entra ID'; Feature = 'Certificate-based authentication'; Severity = 'degraded'; Pattern = '(?:^|\.)certauth\.[a-z0-9.-]+$'; Impact = 'Sign-in with certificates or smart cards fails; password and passwordless sign-in still work.' },
    [PSCustomObject]@{ Id = 'entra-signin'; Area = 'Entra ID'; Feature = 'Sign-in and token issuance'; Severity = 'critical'; Pattern = '(?:^|\.)(?:login\.microsoft\.com|login\.windows\.net|msauth\.net|msftauth\.net|msauthimages\.net|auth\.microsoft\.com|microsoftonline-p\.com|microsoftonline-p\.net|accesscontrol\.windows\.net|activedirectory\.windowsazure\.com)$|^(?:[a-z0-9-]+\.)?login[a-z0-9-]*\.microsoftonline\.com$|^microsoftonline\.com$'; Impact = 'Users cannot sign in or renew tokens. Almost every Microsoft cloud service is affected.' },
    [PSCustomObject]@{ Id = 'entra-sspr'; Area = 'Entra ID'; Feature = 'Self-service password reset'; Severity = 'degraded'; Pattern = '(?:^|\.)passwordreset\.microsoftonline\.com$'; Impact = 'Users cannot reset or change their password themselves.' },
    [PSCustomObject]@{ Id = 'entra-seamless-sso'; Area = 'Entra ID'; Feature = 'Seamless single sign-on'; Severity = 'degraded'; Pattern = '(?:^|\.)autologon\.microsoftazuread-sso\.com$'; Impact = 'Seamless SSO on domain-joined devices fails; users are asked for their password.' },
    [PSCustomObject]@{ Id = 'entra-admin'; Area = 'Entra ID'; Feature = 'Directory administration'; Severity = 'low'; Pattern = '(?:^|\.)(?:adminwebservice|companymanager|provisioningapi)\.microsoftonline\.com$'; Impact = 'Legacy admin tools and Entra Connect provisioning calls fail. Users are not affected.' },
    [PSCustomObject]@{ Id = 'entra-other'; Area = 'Entra ID'; Feature = 'Sign-in and token issuance'; Severity = 'critical'; Pattern = '(?:^|\.)microsoftonline\.com$'; Impact = 'Users cannot sign in or renew tokens. Almost every Microsoft cloud service is affected.' },
    [PSCustomObject]@{ Id = 'entra-device-registration'; Area = 'Entra ID'; Feature = 'Device registration'; Severity = 'critical'; Pattern = '(?:^|\.)enterpriseregistration\.windows\.net$'; Impact = 'Entra join, hybrid join and device registration fail. Conditional Access rules that require a compliant or joined device block access.' },
    [PSCustomObject]@{ Id = 'entra-pas'; Area = 'Entra ID'; Feature = 'Windows sign-in services'; Severity = 'degraded'; Pattern = '(?:^|\.)pas\.windows\.net$'; Impact = 'Entra ID sign-in to Windows devices and VMs and some passwordless scenarios can fail.' },
    [PSCustomObject]@{ Id = 'graph'; Area = 'Microsoft Graph'; Feature = 'Graph API'; Severity = 'degraded'; Pattern = '(?:^|\.)graph\.(?:microsoft\.com|windows\.net)$'; Impact = 'Admin portals, scripts and apps that use Microsoft Graph fail. Some Intune and Autopilot features rely on it.' },
    [PSCustomObject]@{ Id = 'azure-arm'; Area = 'Azure'; Feature = 'Azure Resource Manager'; Severity = 'low'; Pattern = '(?:^|\.)management\.azure\.com$'; Impact = 'Azure portal and Azure management scripts fail. Client devices are not affected.' },
    [PSCustomObject]@{ Id = 'autopilot'; Area = 'Windows Autopilot'; Feature = 'Deployment profile'; Severity = 'critical'; Pattern = '(?:^|\.)(?:ztd|cs)\.dds\.microsoft\.com$|(?:^|\.)dds\.microsoft\.com$'; Impact = 'Devices cannot download their Autopilot profile and go through a normal, unmanaged setup.' },
    [PSCustomObject]@{ Id = 'autopatch'; Area = 'Windows Autopatch'; Feature = 'Autopatch service'; Severity = 'degraded'; Pattern = '(?:^|\.)(?:autopatch\.microsoft\.com|mmdcustomer\.microsoft\.com|mmdls\.microsoft\.com|devicelistenerprod\.microsoft\.com|webpubsub\.azure\.com)$|^payloadprod[a-z0-9*-]*\.blob\.core\.windows\.net$'; Impact = 'Autopatch device registration, readiness checks and update reporting fail. Updates themselves still flow through Windows Update.' },
    [PSCustomObject]@{ Id = 'intune-enrollment'; Area = 'Intune'; Feature = 'Device enrollment'; Severity = 'critical'; Pattern = '(?:^|\.)enrollment\.manage\.microsoft\.com$|(?:^|\.)enterpriseenrollment[a-z0-9.-]*$'; Impact = 'New devices cannot enroll in Intune (Autopilot, user and bulk enrollment).' },
    [PSCustomObject]@{ Id = 'intune-content'; Area = 'Intune'; Feature = 'App and script downloads'; Severity = 'degraded'; Pattern = 'mscdn\.manage\.microsoft\.com$|^swd[a-z0-9-]*\.manage\.microsoft\.com$|(?:^|\.)naprodimedatapri\.azureedge\.net$|^mmcswd[a-z0-9-]*\.blob\.core\.windows\.net$'; Impact = 'Win32 apps, PowerShell scripts and remediations cannot be downloaded; policies still apply.' },
    [PSCustomObject]@{ Id = 'intune-org-messages'; Area = 'Intune'; Feature = 'Organizational messages'; Severity = 'low'; Pattern = '(?:^|\.)fd\.api\.orgmsg\.microsoft\.com$|(?:^|\.)ris\.prod\.api\.personalization\.ideas\.microsoft\.com$'; Impact = 'Organizational messages are not shown on devices.' },
    [PSCustomObject]@{ Id = 'intune-company-portal'; Area = 'Intune'; Feature = 'Company Portal'; Severity = 'degraded'; Pattern = '(?:^|\.)portal\.manage\.microsoft\.com$'; Impact = 'Company Portal app and website cannot load; self-service app installs fail.' },
    [PSCustomObject]@{ Id = 'intune-admin'; Area = 'Intune'; Feature = 'Admin center'; Severity = 'low'; Pattern = '(?:^|\.)(?:endpoint|intune)\.microsoft\.com$'; Impact = 'The Intune admin center cannot be opened from this network. Managed devices are not affected.' },
    [PSCustomObject]@{ Id = 'intune-checkin'; Area = 'Intune'; Feature = 'Device check-in and policies'; Severity = 'critical'; Pattern = '(?:^|\.)manage\.microsoft\.com$|(?:^|\.)dm\.microsoft\.com$'; Impact = 'Devices stop syncing with Intune: no new policies, apps, compliance status or remote actions.' },
    [PSCustomObject]@{ Id = 'defender-signatures'; Area = 'Microsoft Defender'; Feature = 'Signature updates'; Severity = 'critical'; Pattern = '(?:^|\.)definitionupdates\.microsoft\.com$'; Impact = 'Antivirus signatures and engine updates are not downloaded; protection gets outdated.' },
    [PSCustomObject]@{ Id = 'defender-cloud'; Area = 'Microsoft Defender'; Feature = 'Cloud protection (MAPS)'; Severity = 'critical'; Pattern = '(?:^|\.)wdcp\.microsoft\.com$|(?:^|\.)wdcpalt\.microsoft\.com$|(?:^|\.)cp\.wd\.microsoft\.com$|(?:^|\.)wd\.microsoft\.com$'; Impact = 'Cloud-delivered protection, block at first sight and sample submission stop working.' },
    [PSCustomObject]@{ Id = 'defender-edr'; Area = 'Microsoft Defender'; Feature = 'Defender for Endpoint sensor and portal'; Severity = 'degraded'; Pattern = '(?:^|\.)(?:security|defender)\.microsoft\.com$|^winatp-gw-[a-z0-9-]+\.microsoft\.com$|(?:^|\.)endpoint\.security\.microsoft\.com$'; Impact = 'Onboarding, EDR telemetry, live response and the Defender portal are affected.' },
    [PSCustomObject]@{ Id = 'smartscreen'; Area = 'Microsoft Defender'; Feature = 'SmartScreen'; Severity = 'degraded'; Pattern = 'smartscreen|(?:^|\.)urs\.microsoft\.com$'; Impact = 'SmartScreen reputation checks for downloads, apps and websites fail.' },
    [PSCustomObject]@{ Id = 'exchange-protection'; Area = 'Exchange Online'; Feature = 'Exchange Online Protection'; Severity = 'degraded'; Pattern = '(?:^|\.)protection\.outlook\.com$'; Impact = 'Mail flow through Exchange Online Protection and Safe Links/Attachments checks are affected.' },
    [PSCustomObject]@{ Id = 'exchange'; Area = 'Exchange Online'; Feature = 'Mail, calendar and Autodiscover'; Severity = 'critical'; Pattern = '(?:^|\.)outlook\.(?:office|office365|com|cloud\.microsoft)(?:\.com)?$|^autodiscover\.|(?:^|\.)smtp\.office365\.com$|(?:^|\.)outlook\.office365\.com$'; Impact = 'Outlook cannot connect: no mail, calendar or contacts, Autodiscover fails.' },
    [PSCustomObject]@{ Id = 'office-config'; Area = 'Microsoft 365 Apps'; Feature = 'Feature configuration (ECS)'; Severity = 'degraded'; Pattern = '(?:^|\.)config\.edge\.skype\.com$|(?:^|\.)ecs\.office\.com$'; Impact = 'Office, Teams and Edge cannot load feature flags and configuration; some features stay off.' },
    [PSCustomObject]@{ Id = 'teams'; Area = 'Teams'; Feature = 'Chat, meetings and calls'; Severity = 'critical'; Pattern = '(?:^|\.)teams\.(?:microsoft\.com|cloud\.microsoft|live\.com)$|(?:^|\.)skype\.com$|(?:^|\.)lync\.com$|(?:^|\.)teams\.cdn\.office\.net$|(?:^|\.)skypeforbusiness\.com$'; Impact = 'Teams chat, meetings, calls and media are disrupted.' },
    [PSCustomObject]@{ Id = 'onedrive-client'; Area = 'SharePoint & OneDrive'; Feature = 'OneDrive sync client updates'; Severity = 'low'; Pattern = '(?:^|\.)oneclient\.sfx\.ms$|(?:^|\.)g\.live\.com$'; Impact = 'The OneDrive sync app cannot update itself; sync keeps working.' },
    [PSCustomObject]@{ Id = 'sharepoint'; Area = 'SharePoint & OneDrive'; Feature = 'Files, sync and sharing'; Severity = 'critical'; Pattern = '(?:^|\.)sharepoint\.com$|(?:^|\.)sharepointonline\.com$|(?:^|\.)onedrive\.(?:com|live\.com)$|(?:^|\.)svc\.ms$|(?:^|\.)storage\.live\.com$'; Impact = 'SharePoint sites and OneDrive files cannot be opened or synced; Teams files are affected too.' },
    [PSCustomObject]@{ Id = 'm365-apps-install'; Area = 'Microsoft 365 Apps'; Feature = 'Install and updates'; Severity = 'degraded'; Pattern = '(?:^|\.)officecdn\.microsoft\.com(?:\.edgesuite\.net)?$|(?:^|\.)officecdn-microsoft-com\.akamaized\.net$'; Impact = 'Microsoft 365 Apps cannot be installed or updated (Click-to-Run).' },
    [PSCustomObject]@{ Id = 'm365-apps-licensing'; Area = 'Microsoft 365 Apps'; Feature = 'Licensing and cloud policy'; Severity = 'critical'; Pattern = '(?:^|\.)officeclient\.microsoft\.com$|(?:^|\.)config\.office\.com$|(?:^|\.)config\.office\.net$|(?:^|\.)ols\.officeapps\.live\.com$'; Impact = 'Office apps cannot activate or refresh their license and switch to reduced functionality; Cloud Policy settings do not apply.' },
    [PSCustomObject]@{ Id = 'm365-admin'; Area = 'Microsoft 365'; Feature = 'Admin and compliance portals'; Severity = 'low'; Pattern = '(?:^|\.)admin\.microsoft\.com$|(?:^|\.)protection\.office\.com$|(?:^|\.)compliance\.microsoft\.com$|(?:^|\.)purview\.microsoft\.com$'; Impact = 'Microsoft 365 admin, Purview and compliance portals cannot be opened. End users are not affected.' },
    [PSCustomObject]@{ Id = 'm365-telemetry'; Area = 'Microsoft 365 Apps'; Feature = 'Diagnostics'; Severity = 'low'; Pattern = '(?:^|\.)nexus(?:rules)?\.officeapps\.live\.com$|(?:^|\.)office\.events\.data\.microsoft\.com$|(?:^|\.)aria\.microsoft\.com$|(?:^|\.)otelrules\.azureedge\.net$'; Impact = 'Office diagnostic data is not sent. Users will not notice.' },
    [PSCustomObject]@{ Id = 'purview-mip'; Area = 'Microsoft Purview'; Feature = 'Sensitivity labels and encryption'; Severity = 'degraded'; Pattern = 'informationprotection'; Impact = 'Sensitivity labels cannot be applied and protected (encrypted) documents and mails cannot be opened.' },
    [PSCustomObject]@{ Id = 'power-platform'; Area = 'Power Platform'; Feature = 'Power Automate and connectors'; Severity = 'degraded'; Pattern = '(?:^|\.)flow\.microsoft\.com$|(?:^|\.)azure-apim\.net$|(?:^|\.)azurerms\.com$|(?:^|\.)powerautomate\.com$'; Impact = 'Power Automate flows and Power Platform connectors used in Microsoft 365 fail.' },
    [PSCustomObject]@{ Id = 'm365-media'; Area = 'Microsoft 365'; Feature = 'Stream and meeting recordings'; Severity = 'low'; Pattern = '(?:^|\.)mediaservices\.windows\.net$|(?:^|\.)amp\.azure\.net$'; Impact = 'Videos and meeting recordings may not play.' },
    [PSCustomObject]@{ Id = 'm365-dictation'; Area = 'Microsoft 365 Apps'; Feature = 'Dictation and speech'; Severity = 'low'; Pattern = '(?:^|\.)officespeech\.platform\.bing\.com$'; Impact = 'Dictation and read-aloud features in Office apps do not work.' },
    [PSCustomObject]@{ Id = 'm365-redirects'; Area = 'Microsoft 365 Apps'; Feature = 'Add-ins, help and redirects'; Severity = 'low'; Pattern = 'officeredir|officepreviewredir|(?:^|\.)appsforoffice\.microsoft\.com$|(?:^|\.)office15client\.microsoft\.com$|^r\.office\.microsoft\.com$|^c1\.microsoft\.com$|^office\.microsoft$'; Impact = 'Office add-ins store, in-app help and some links inside Office apps do not open.' },
    [PSCustomObject]@{ Id = 'm365-web'; Area = 'Microsoft 365'; Feature = 'Web apps and portal'; Severity = 'degraded'; Pattern = '(?:^|\.)(?:office|office365|office\.net|officeapps\.live\.com|online\.office\.com|cloud\.microsoft|static\.microsoft|usercontent\.microsoft|microsoft365\.com|m365\.cloud\.microsoft)(?:\.com)?$|(?:^|\.)portal\.office\.com$|(?:^|\.)msocdn\.com$|(?:^|\.)wns\.windows\.com$|(?:^|\.)mx\.microsoft$|(?:^|\.)onmicrosoft\.com$'; Impact = 'office.com, Office for the web and parts of the desktop apps (add-ins, templates, sharing) fail.' },
    [PSCustomObject]@{ Id = 'edge-extensions'; Area = 'Microsoft Edge'; Feature = 'Extensions'; Severity = 'low'; Pattern = 'msedgeextensions'; Impact = 'Edge extensions cannot be installed or updated from the add-ons store.' },
    [PSCustomObject]@{ Id = 'edge-activity'; Area = 'Microsoft Edge'; Feature = 'Enterprise sync and activity'; Severity = 'low'; Pattern = 'edge-enterprise\.activity\.windows\.com$'; Impact = 'Edge enterprise sync and activity features are limited.' },
    [PSCustomObject]@{ Id = 'edge-misc'; Area = 'Microsoft Edge'; Feature = 'Optional browser features'; Severity = 'low'; Pattern = '(?:^|\.)(?:edgepasskeysenclave|privacy|rewardsplatform|prod\.rewardsplatform)\.microsoft\.com$'; Impact = 'Optional Edge features such as rewards, privacy dashboard or passkey sync are unavailable.' },
    [PSCustomObject]@{ Id = 'edge-updates'; Area = 'Microsoft Edge'; Feature = 'Updates and services'; Severity = 'degraded'; Pattern = '(?:^|\.)edge\.microsoft\.com$|(?:^|\.)msedge\.(?:api\.cdp\.microsoft\.com|net)$|(?:^|\.)api\.cdp\.microsoft\.com$|^msedge\.[a-z0-9.]*dl\.delivery\.mp\.microsoft\.com$'; Impact = 'Edge cannot update itself or load online services.' },
    [PSCustomObject]@{ Id = 'delivery-optimization'; Area = 'Windows Update'; Feature = 'Delivery Optimization'; Severity = 'low'; Pattern = '(?:^|\.)do\.dsp\.mp\.microsoft\.com$|(?:^|\.)dsp\.mp\.microsoft\.com$'; Impact = 'Peer caching is not available; devices download updates directly, which uses more bandwidth.' },
    [PSCustomObject]@{ Id = 'store-catalog'; Area = 'Microsoft Store'; Feature = 'Catalog, purchase and licensing'; Severity = 'degraded'; Pattern = '(?:^|\.)(?:displaycatalog|purchase\.md|licensing|storeedgefd\.dsx)\.mp\.microsoft\.com$'; Impact = 'Store apps, including Store apps deployed by Intune, cannot be installed, updated or licensed.' },
    [PSCustomObject]@{ Id = 'store-media'; Area = 'Microsoft Store'; Feature = 'Images and live tiles'; Severity = 'low'; Pattern = '(?:^|\.)livetileedge\.dsx\.mp\.microsoft\.com$|(?:^|\.)store-images\.s-microsoft\.com$|akamaized\.net$|(?:^|\.)storequality\.microsoft\.com$'; Impact = 'Store images, tiles and revocation lists do not load; installs still work.' },
    [PSCustomObject]@{ Id = 'wu-scan'; Area = 'Windows Update'; Feature = 'Update scan and approval'; Severity = 'critical'; Pattern = '(?:^|\.)fe[0-9]\.(?:ws\.)?(?:delivery\.mp\.microsoft\.com)$|(?:^|\.)update\.microsoft\.com$|(?:^|\.)windowsupdate\.microsoft\.com$'; Impact = 'Devices cannot check for updates. No security, feature or driver updates are offered.' },
    [PSCustomObject]@{ Id = 'wu-download'; Area = 'Windows Update'; Feature = 'Update downloads'; Severity = 'critical'; Pattern = '(?:^|\.)delivery\.mp\.microsoft\.com$|(?:^|\.)windowsupdate\.com$'; Impact = 'Update packages, feature updates, drivers and Store app packages cannot be downloaded.' },
    [PSCustomObject]@{ Id = 'windows-activation'; Area = 'Windows Activation'; Feature = 'Activation and license validation'; Severity = 'degraded'; Pattern = '(?:^|\.)(?:activation|activation-v2|validation|validation-v2)\.sls\.microsoft\.com$'; Impact = 'Windows cannot activate or validate its license (MAK, retail, subscription activation).' },
    [PSCustomObject]@{ Id = 'crl'; Area = 'Certificates'; Feature = 'Revocation checks (CRL/OCSP)'; Severity = 'degraded'; Pattern = '(?:^|\.)crl\.microsoft\.com$|(?:^|\.)ocsp\.|(?:^|\.)oneocsp\.microsoft\.com$|(?:^|\.)mscrl\.microsoft\.com$'; Impact = 'TLS certificate revocation checks fail or time out. Sign-in, activation and downloads can hang.' },
    [PSCustomObject]@{ Id = 'windows-telemetry'; Area = 'Windows diagnostics'; Feature = 'Diagnostic data and crash reports'; Severity = 'low'; Pattern = '(?:^|\.)events\.data\.microsoft\.com$|(?:^|\.)vortex[a-z0-9-]*\.data\.microsoft\.com$|(?:^|\.)settings-win\.data\.microsoft\.com$|watson|watcab|(?:^|\.)oca\.(?:telemetry\.)?microsoft\.com$'; Impact = 'Diagnostic data and crash reports are not sent. Windows Update reports and Endpoint analytics may be incomplete.' },
    [PSCustomObject]@{ Id = 'microsoft-web'; Area = 'Microsoft web'; Feature = 'Links and redirects'; Severity = 'low'; Pattern = '^(?:go|www|aka)\.microsoft\.com$|^aka\.ms$'; Impact = 'Help links and redirects from Microsoft products do not open.' },
    [PSCustomObject]@{ Id = 'microsoft-account'; Area = 'Microsoft account'; Feature = 'Consumer services'; Severity = 'low'; Pattern = '(?:^|\.)live\.com$'; Impact = 'Personal Microsoft account features (consumer sign-in, settings sync) are limited.' }
)
# END GENERATED IMPACT RULES

$impactSeverityOrder = @{ 'critical' = 0; 'degraded' = 1; 'low' = 2 }
$m365ServiceAreaNames = @{
    'Exchange'   = 'Exchange Online'
    'SharePoint' = 'SharePoint & OneDrive'
    'Skype'      = 'Teams'
    'Common'     = 'Microsoft 365'
}

$servicePlatformSupport = @{
    'Windows Activation' = @('Windows')
}

function Test-IsWindowsPlatform {
    return [System.Runtime.InteropServices.RuntimeInformation]::IsOSPlatform([System.Runtime.InteropServices.OSPlatform]::Windows)
}

function Test-IsMacOSPlatform {
    return [System.Runtime.InteropServices.RuntimeInformation]::IsOSPlatform([System.Runtime.InteropServices.OSPlatform]::OSX)
}

function Test-IsLinuxPlatform {
    return [System.Runtime.InteropServices.RuntimeInformation]::IsOSPlatform([System.Runtime.InteropServices.OSPlatform]::Linux)
}

function Test-ServiceSupportedOnCurrentPlatform {
    param (
        [string]$ServiceName
    )

    if (-not $servicePlatformSupport.ContainsKey($ServiceName)) {
        return $true
    }

    $supportedPlatforms = $servicePlatformSupport[$ServiceName]

    if (Test-IsWindowsPlatform) {
        return $supportedPlatforms -contains 'Windows'
    }

    if (Test-IsMacOSPlatform) {
        return $supportedPlatforms -contains 'macOS'
    }

    if (Test-IsLinuxPlatform) {
        return $supportedPlatforms -contains 'Linux'
    }

    return $false
}

# Function to test ping latency
function Test-PingLatency {
    param (
        [string]$HostName,
        [int]$Count = 4
    )
    
    try {
        $pingResults = Test-Connection -ComputerName $HostName -Count $Count -ErrorAction Stop
        $avgLatency = ($pingResults | Measure-Object ResponseTime -Average).Average
        return [math]::Round($avgLatency, 2)
    }
    catch {
        return $null
    }
}

# Function to test HTTPS/TCP connectivity without Windows-specific cmdlets
function Test-TcpConnectivity {
    param (
        [string]$HostName,
        [int]$Port = 443,
        [int]$TimeoutMilliseconds = 5000
    )

    $addresses = @()

    try {
        $addresses = [System.Net.Dns]::GetHostAddresses($HostName) |
            Where-Object { $_.AddressFamily -in @([System.Net.Sockets.AddressFamily]::InterNetwork, [System.Net.Sockets.AddressFamily]::InterNetworkV6) }
    }
    catch {
        return [PSCustomObject]@{
            TcpTestSucceeded = $false
            RemoteAddress = $null
        }
    }

    foreach ($address in $addresses) {
        $client = [System.Net.Sockets.TcpClient]::new()
        try {
            $asyncResult = $client.BeginConnect($address, $Port, $null, $null)
            if ($asyncResult.AsyncWaitHandle.WaitOne($TimeoutMilliseconds, $false) -and $client.Connected) {
                $client.EndConnect($asyncResult)

                $remoteAddress = $null
                try {
                    $remoteEndpoint = $client.Client.RemoteEndPoint
                    if ($remoteEndpoint -and $remoteEndpoint.Address) {
                        $remoteAddress = $remoteEndpoint.Address.IPAddressToString
                    }
                }
                catch {
                    $remoteAddress = $address.IPAddressToString
                }

                return [PSCustomObject]@{
                    TcpTestSucceeded = $true
                    RemoteAddress = $remoteAddress
                }
            }
        }
        catch {
        }
        finally {
            $client.Close()
            $client.Dispose()
        }
    }

    return [PSCustomObject]@{
        TcpTestSucceeded = $false
        RemoteAddress = $null
    }
}

# Function to test HTTPS response time (ms)
function Test-DownloadSpeed {
    param (
        [string]$Url,
        [int]$TimeoutSeconds = 10
    )
    
    try {
        # Create a test URL by adding a small file path (many Microsoft endpoints serve content)
        $testUrl = $Url.TrimEnd('/') + "/favicon.ico"
        
        $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
        $webClient = New-Object System.Net.WebClient
        $webClient.Headers.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36")
        
        try {
            $data = $webClient.DownloadData($testUrl)
            $stopwatch.Stop()
            
            if ($stopwatch.ElapsedMilliseconds -gt 0 -and $data.Length -gt 0) {
                return [math]::Round($stopwatch.ElapsedMilliseconds, 2)
            }
        }
        catch {
            # If favicon.ico doesn't work, try a HEAD request to measure server response time
            $request = [System.Net.WebRequest]::Create($Url)
            $request.Method = "HEAD"
            $request.Timeout = $TimeoutSeconds * 1000
            
            $stopwatch.Restart()
            $response = $request.GetResponse()
            $stopwatch.Stop()
            $response.Close()
            
            # Return response time in ms as a "speed" indicator (lower is better)
            return [math]::Round($stopwatch.ElapsedMilliseconds, 2)
        }
        finally {
            $webClient.Dispose()
        }
    }
    catch {
        return $null
    }
    
    return $null
}

# Function to display service impact warnings
function Show-ServiceImpactWarning {
    param (
        [string]$ServiceName,
        [array]$FailedEndpoints
    )
    
    if ($serviceImpacts.ContainsKey($ServiceName)) {
        Write-Host "    [WARN] IMPACT:" -ForegroundColor Yellow
        Write-Host "       • $($serviceImpacts[$ServiceName].Impact)" -ForegroundColor Yellow
        Write-Host "    SYMPTOMS:" -ForegroundColor Cyan
        Write-Host "       • $($serviceImpacts[$ServiceName].Symptoms)" -ForegroundColor Cyan
        Write-Host "    Affected endpoints: $($FailedEndpoints.Count)" -ForegroundColor Red
        Write-Host ""
    }
}

# Loads serviceArea and category (Optimize/Allow/Default) of Microsoft 365 URLs
# from the official Microsoft 365 endpoint web service. Returns $null if offline.
function Get-M365EndpointLookup {
    $uri = 'https://endpoints.office.com/endpoints/worldwide?clientrequestid=b10c5ed1-bad1-445f-b386-b919946339a7'
    try {
        $entries = Invoke-RestMethod -Uri $uri -TimeoutSec 15 -ErrorAction Stop
    } catch {
        return $null
    }

    $exact = @{}
    $wildcards = New-Object System.Collections.ArrayList
    foreach ($entry in @($entries)) {
        $info = [PSCustomObject]@{
            ServiceArea = [string]$entry.serviceArea
            Category    = [string]$entry.category
            Required    = [bool]$entry.required
        }
        foreach ($rawUrl in @($entry.urls)) {
            if (-not $rawUrl) { continue }
            $url = ([string]$rawUrl).Trim().ToLowerInvariant()
            if ($url.StartsWith('*.')) {
                [void]$wildcards.Add([PSCustomObject]@{ Suffix = $url.Substring(1); Info = $info })
            }
            if (-not $exact.ContainsKey($url)) { $exact[$url] = $info }
        }
    }

    return [PSCustomObject]@{
        Exact     = $exact
        # Longest suffix first, so *.teams.microsoft.com wins over *.microsoft.com
        Wildcards = @($wildcards | Sort-Object -Property @{ Expression = { $_.Suffix.Length }; Descending = $true })
    }
}

function Get-M365EndpointInfo {
    param (
        [string]$HostName,
        $Lookup
    )

    if (-not $Lookup) { return $null }
    $name = $HostName.ToLowerInvariant()
    if ($Lookup.Exact.ContainsKey($name)) { return $Lookup.Exact[$name] }
    $bare = $name -replace '^\*\.', ''
    foreach ($wildcard in $Lookup.Wildcards) {
        # Too broad to say anything about the service
        if ($wildcard.Suffix -eq '.microsoft.com' -or $wildcard.Suffix -eq '.windows.net') { continue }
        if ($bare.EndsWith($wildcard.Suffix) -or ".$bare" -eq $wildcard.Suffix) { return $wildcard.Info }
    }
    return $null
}

# Maps an endpoint to the service feature it belongs to (see $endpointImpactRules)
function Get-EndpointImpact {
    param (
        [string]$HostName,
        [string]$ServiceName,
        $M365Info
    )

    $name = $HostName.Trim().ToLowerInvariant() -replace '^\*\.', ''
    $rule = $null
    foreach ($candidate in $endpointImpactRules) {
        if ($name -match $candidate.Pattern) { $rule = $candidate; break }
    }

    if ($rule) {
        $impact = [PSCustomObject]@{
            Area         = $rule.Area
            Feature      = $rule.Feature
            Severity     = $rule.Severity
            Impact       = $rule.Impact
            M365Category = $null
        }
    } else {
        $area = if ($ServiceName) { $ServiceName } else { 'Other' }
        $impact = [PSCustomObject]@{
            Area         = $area
            Feature      = 'Other endpoints'
            Severity     = 'degraded'
            Impact       = "Parts of $area may not work."
            M365Category = $null
        }
    }

    # Microsoft marks endpoints that are required for a service as Optimize or Allow
    if ($M365Info -and $M365Info.Category) {
        if (-not $rule -and $m365ServiceAreaNames.ContainsKey($M365Info.ServiceArea)) {
            $impact.Area = $m365ServiceAreaNames[$M365Info.ServiceArea]
        }
        if ($M365Info.Category -eq 'Optimize' -or $M365Info.Category -eq 'Allow') {
            $impact.Severity = 'critical'
        } elseif ($M365Info.Category -eq 'Default' -and $impact.Severity -eq 'critical') {
            $impact.Severity = 'degraded'
        }
        $impact.M365Category = $M365Info.Category
    }

    return $impact
}

# Groups failed endpoints by feature: one entry per affected feature,
# with how many of its endpoints failed and what users will notice.
function Get-ServiceImpactSummary {
    param ([array]$Results)

    $groups = [ordered]@{}
    foreach ($result in $Results) {
        if (-not $result.ImpactFeature) { continue }
        $key = "$($result.ImpactArea)::$($result.ImpactFeature)"
        if (-not $groups.Contains($key)) {
            $groups[$key] = [PSCustomObject]@{
                Area     = $result.ImpactArea
                Feature  = $result.ImpactFeature
                Severity = $result.ImpactSeverity
                Impact   = $result.ImpactText
                Hosts    = @{}
            }
        }
        $group = $groups[$key]
        if ($impactSeverityOrder[$result.ImpactSeverity] -lt $impactSeverityOrder[$group.Severity]) {
            $group.Severity = $result.ImpactSeverity
        }
        $hostKey = $result.Host.ToLowerInvariant()
        if (-not $group.Hosts.ContainsKey($hostKey) -or $result.Status -eq 'FAILED') {
            $group.Hosts[$hostKey] = $result.Status
        }
    }

    $summary = foreach ($group in $groups.Values) {
        $failedHosts = @($group.Hosts.Keys | Where-Object { $group.Hosts[$_] -eq 'FAILED' } | Sort-Object)
        if ($failedHosts.Count -eq 0) { continue }
        $total = $group.Hosts.Count
        $likelihood = if ($failedHosts.Count -eq $total) { 'likely affected' }
                      elseif (($failedHosts.Count / $total) -ge 0.5) { 'probably affected' }
                      else { 'possibly affected' }
        [PSCustomObject]@{
            Area        = $group.Area
            Feature     = $group.Feature
            Severity    = $group.Severity
            Impact      = $group.Impact
            Failed      = $failedHosts.Count
            Total       = $total
            AllFailed   = ($failedHosts.Count -eq $total)
            Likelihood  = $likelihood
            FailedHosts = $failedHosts
        }
    }

    return @($summary | Sort-Object -Property `
        @{ Expression = { $impactSeverityOrder[$_.Severity] } },
        @{ Expression = { $_.AllFailed }; Descending = $true },
        @{ Expression = { $_.Failed }; Descending = $true },
        @{ Expression = { $_.Area } },
        @{ Expression = { $_.Feature } })
}

function Show-ServiceImpactSummary {
    param ([array]$ImpactSummary)

    if (-not $ImpactSummary -or $ImpactSummary.Count -eq 0) { return }
    Write-Host "LIKELY SERVICE IMPACT (by feature):" -ForegroundColor Yellow
    Write-Host "-"*50 -ForegroundColor Yellow
    foreach ($entry in $ImpactSummary) {
        $color = switch ($entry.Severity) { 'critical' { 'Red' } 'degraded' { 'Yellow' } default { 'Gray' } }
        Write-Host ("[{0}] {1}: {2} - {3} of {4} endpoint(s) unreachable, {5}" -f $entry.Severity.ToUpper(), $entry.Area, $entry.Feature, $entry.Failed, $entry.Total, $entry.Likelihood) -ForegroundColor $color
        Write-Host "   └─ $($entry.Impact)" -ForegroundColor $color
        Write-Host "      Endpoints: $($entry.FailedHosts -join ', ')" -ForegroundColor DarkGray
    }
}

function ConvertTo-HtmlText {
    param ([string]$Text)
    return [System.Net.WebUtility]::HtmlEncode($Text)
}

# Function to generate HTML report
function New-HtmlReport {
    param (
        [array]$Results,
        [hashtable]$SelectedServices,
        [array]$FailedServices,
        [hashtable]$ServiceImpacts,
        [array]$ImpactSummary,
        [bool]$SkipPing,
        [bool]$SkipSpeed,
        [string]$FilePath
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $testDuration = (Get-Date) - $script:startTime
    
    # Calculate statistics
    $totalEndpoints = $Results.Count
    $successfulEndpoints = ($Results | Where-Object { $_.Status -eq 'OK' }).Count
    $failedEndpoints = $totalEndpoints - $successfulEndpoints
    $successRate = if ($totalEndpoints -gt 0) { [math]::Round(($successfulEndpoints / $totalEndpoints) * 100, 1) } else { 0 }
    
    $pingStats = $null
    $speedStats = $null
    
    if (-not $SkipPing) {
        $pingResults = $Results | Where-Object { $_.Status -eq 'OK' -and $null -ne $_.PingLatency_ms }
        if ($pingResults.Count -gt 0) {
            $pingStats = @{
                Average = [math]::Round(($pingResults | Measure-Object PingLatency_ms -Average).Average, 2)
                Min = [math]::Round(($pingResults | Measure-Object PingLatency_ms -Minimum).Minimum, 2)
                Max = [math]::Round(($pingResults | Measure-Object PingLatency_ms -Maximum).Maximum, 2)
                Count = $pingResults.Count
            }
        }
    }
    
    if (-not $SkipSpeed) {
        $speedResults = $Results | Where-Object { $_.Status -eq 'OK' -and $null -ne $_.DownloadSpeed_Kbps }
        if ($speedResults.Count -gt 0) {
            $speedStats = @{
                Average = [math]::Round(($speedResults | Measure-Object DownloadSpeed_Kbps -Average).Average, 2)
                Min = [math]::Round(($speedResults | Measure-Object DownloadSpeed_Kbps -Minimum).Minimum, 2)
                Max = [math]::Round(($speedResults | Measure-Object DownloadSpeed_Kbps -Maximum).Maximum, 2)
                Count = $speedResults.Count
            }
        }
    }

    # Generate HTML content
    $htmlContent = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Microsoft Endpoint Connectivity Report</title>
    <style>
        * { box-sizing: border-box; }
        body { 
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; 
            margin: 0; 
            padding: 20px; 
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
        }
        .container { 
            max-width: 1200px; 
            margin: 0 auto; 
            background: white; 
            border-radius: 10px; 
            box-shadow: 0 10px 30px rgba(0,0,0,0.3);
            overflow: hidden;
        }
        .header { 
            background: linear-gradient(135deg, #0078d4 0%, #106ebe 100%); 
            color: white; 
            padding: 30px; 
            text-align: center; 
        }
        .header h1 { margin: 0; font-size: 2.5em; font-weight: 300; }
        .header p { margin: 10px 0 0 0; opacity: 0.9; font-size: 1.1em; }
        
        .summary { 
            display: grid; 
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); 
            gap: 20px; 
            padding: 30px; 
            background: #f8f9fa;
        }
        .stat-card { 
            background: white; 
            padding: 25px; 
            border-radius: 8px; 
            text-align: center; 
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
            border-left: 4px solid #0078d4;
        }
        .stat-value { 
            font-size: 2.5em; 
            font-weight: bold; 
            margin-bottom: 5px; 
        }
        .stat-label { 
            color: #666; 
            font-size: 0.9em; 
            text-transform: uppercase; 
            letter-spacing: 1px; 
        }
        .success { color: #28a745; border-left-color: #28a745 !important; }
        .danger { color: #dc3545; border-left-color: #dc3545 !important; }
        .warning { color: #ffc107; border-left-color: #ffc107 !important; }
        .info { color: #17a2b8; border-left-color: #17a2b8 !important; }
        
        .content { padding: 30px; }
        .section { margin-bottom: 40px; }
        .section h2 { 
            color: #333; 
            border-bottom: 2px solid #0078d4; 
            padding-bottom: 10px; 
            margin-bottom: 20px;
            font-size: 1.8em;
            font-weight: 300;
        }
        
        .service-group { 
            margin-bottom: 30px; 
            border: 1px solid #e0e0e0; 
            border-radius: 8px; 
            overflow: hidden;
        }
        .service-header { 
            background: #f1f3f4; 
            padding: 15px 20px; 
            font-weight: bold; 
            font-size: 1.2em;
            color: #333;
            display: flex;
            align-items: center;
            justify-content: space-between;
        }
        .service-status { 
            padding: 3px 12px; 
            border-radius: 20px; 
            color: white; 
            font-size: 0.8em; 
            font-weight: bold;
        }
        
        table { 
            width: 100%; 
            border-collapse: collapse; 
            margin-top: 0;
        }
        th, td { 
            text-align: left; 
            padding: 12px 15px; 
            border-bottom: 1px solid #e0e0e0; 
        }
        th { 
            background: #f8f9fa; 
            font-weight: 600; 
            color: #333;
            font-size: 0.9em;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }
        tr:hover { background-color: #f8f9fa; }
        
        .status-ok { 
            background: #d4edda; 
            color: #155724; 
            padding: 4px 8px; 
            border-radius: 4px; 
            font-size: 0.8em; 
            font-weight: bold;
        }
        .status-failed { 
            background: #f8d7da; 
            color: #721c24; 
            padding: 4px 8px; 
            border-radius: 4px; 
            font-size: 0.8em; 
            font-weight: bold;
        }
        
        .ping-excellent { color: #28a745; font-weight: bold; }
        .ping-good { color: #ffc107; font-weight: bold; }
        .ping-poor { color: #dc3545; font-weight: bold; }
        
        .footer { 
            background: #f8f9fa; 
            padding: 20px; 
            text-align: center; 
            color: #666; 
            font-size: 0.9em;
            border-top: 1px solid #e0e0e0;
        }
        
        .alert { 
            padding: 15px 20px; 
            margin: 20px 0; 
            border-radius: 5px; 
            border-left: 4px solid;
        }
        .alert-danger { 
            background: #f8d7da; 
            border-left-color: #dc3545; 
            color: #721c24; 
        }
        .alert-success { 
            background: #d4edda; 
            border-left-color: #28a745; 
            color: #155724; 
        }
        
        .impact-section {
            background: #fff3cd;
            border: 1px solid #ffeaa7;
            border-radius: 5px;
            padding: 15px;
            margin-top: 10px;
        }
        
        .impact-title {
            font-weight: bold;
            color: #856404;
            margin-bottom: 10px;
        }
        
        .impact-list {
            list-style: none;
            padding: 0;
            margin: 0;
        }
        
        .impact-list li {
            padding: 5px 0;
            border-bottom: 1px solid #ffeaa7;
        }
        
        .impact-list li:last-child {
            border-bottom: none;
        }

        .feature-impact td { vertical-align: top; }
        .feature-impact .hosts { color: #6c757d; font-size: 0.85em; word-break: break-all; }
        .severity {
            display: inline-block;
            padding: 2px 8px;
            border-radius: 4px;
            font-size: 0.8em;
            font-weight: bold;
            text-transform: uppercase;
        }
        .severity-critical { background: #f8d7da; color: #721c24; }
        .severity-degraded { background: #fff3cd; color: #856404; }
        .severity-low { background: #e2e3e5; color: #383d41; }
        .legend { color: #6c757d; font-size: 0.85em; margin-top: 10px; }

        @media (max-width: 768px) {
            .summary { grid-template-columns: repeat(2, 1fr); }
            .container { margin: 10px; }
            body { padding: 10px; }
            table { font-size: 0.8em; }
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>Microsoft Endpoint Connectivity Report</h1>
            <p>Generated on $timestamp | Test Duration: $($testDuration.ToString("hh\:mm\:ss"))</p>
        </div>
        
        <div class="summary">
            <div class="stat-card">
                <div class="stat-value">$totalEndpoints</div>
                <div class="stat-label">Tested Endpoints</div>
            </div>
            <div class="stat-card success">
                <div class="stat-value">$successfulEndpoints</div>
                <div class="stat-label">Successful</div>
            </div>
            <div class="stat-card danger">
                <div class="stat-value">$failedEndpoints</div>
                <div class="stat-label">Failed</div>
            </div>
            <div class="stat-card info">
                <div class="stat-value">$successRate%</div>
                <div class="stat-label">Success Rate</div>
            </div>
"@

    # Add ping statistics if available
    if ($pingStats) {
        $htmlContent += @"
            <div class="stat-card">
                <div class="stat-value">$($pingStats.Average)ms</div>
                <div class="stat-label">Avg Latency</div>
            </div>
"@
    }

    # Add speed statistics if available
    if ($speedStats) {
        $htmlContent += @"
            <div class="stat-card">
                <div class="stat-value">$($speedStats.Average)</div>
                <div class="stat-label">Avg Response (ms)</div>
            </div>
"@
    }

    $htmlContent += @"
        </div>
        
        <div class="content">
"@

    # Add overall status alert
    if ($failedEndpoints -eq 0) {
        $htmlContent += @"
            <div class="alert alert-success">
                <strong>[OK] Excellent!</strong> All tested Microsoft endpoints are reachable and functional.
            </div>
"@
    } else {
        $htmlContent += @"
            <div class="alert alert-danger">
                <strong>Warning!</strong> $failedEndpoints of $totalEndpoints Microsoft endpoints are not reachable. 
                This may lead to functional limitations in the affected services.
            </div>
"@
    }

    # Add feature-level impact section
    if ($ImpactSummary -and $ImpactSummary.Count -gt 0) {
        $htmlContent += @"
            <div class="section">
                <h2>Likely Service Impact</h2>
                <p>Each endpoint belongs to one service feature. The more of a feature's endpoints fail, the more likely it is broken.</p>
                <table class="feature-impact">
                    <thead>
                        <tr>
                            <th>Severity</th>
                            <th>Feature</th>
                            <th>Unreachable</th>
                            <th>What users will notice</th>
                        </tr>
                    </thead>
                    <tbody>
"@
        foreach ($entry in $ImpactSummary) {
            $hostList = ($entry.FailedHosts | ForEach-Object { ConvertTo-HtmlText $_ }) -join ', '
            $htmlContent += @"
                        <tr>
                            <td><span class="severity severity-$($entry.Severity)">$($entry.Severity)</span></td>
                            <td><strong>$(ConvertTo-HtmlText $entry.Area)</strong><br>$(ConvertTo-HtmlText $entry.Feature)</td>
                            <td>$($entry.Failed) of $($entry.Total)<br><small>$($entry.Likelihood)</small></td>
                            <td>$(ConvertTo-HtmlText $entry.Impact)<div class="hosts">$hostList</div></td>
                        </tr>
"@
        }
        $htmlContent += @"
                    </tbody>
                </table>
                <p class="legend">Critical: the core function stops working. Degraded: a feature is missing or slower. Low: optional, admin-only or telemetry endpoints. Microsoft 365 endpoints marked Optimize or Allow by Microsoft count as critical.</p>
            </div>
"@
    }

    # Add services section
    $htmlContent += @"
            <div class="section">
                <h2>Service Details</h2>
"@

    # Generate service-specific tables
    foreach ($serviceName in $SelectedServices.Keys) {
        $serviceResults = $Results | Where-Object { $SelectedServices[$serviceName] -contains $_.URL }
        $serviceFailures = ($serviceResults | Where-Object { $_.Status -eq 'FAILED' }).Count

        
        $statusClass = if ($serviceFailures -eq 0) { "success" } else { "danger" }
        $statusText = if ($serviceFailures -eq 0) { "[OK] All OK" } else { "[FAIL] $serviceFailures Errors" }
        
        $htmlContent += @"
                <div class="service-group">
                    <div class="service-header">
                        $serviceName
                        <span class="service-status $statusClass">$statusText</span>
                    </div>
                    <table>
                        <thead>
                            <tr>
                                <th>URL</th>
                                <th>Status</th>
                                <th>IP-Adresse</th>
                                <th>Feature</th>
"@

        if (-not $SkipPing) {
            $htmlContent += "<th>Latenz (ms)</th>"
        }
        
        if (-not $SkipSpeed) {
            $htmlContent += "<th>Response (ms)</th>"
        }

        $htmlContent += @"
                            </tr>
                        </thead>
                        <tbody>
"@

        foreach ($result in $serviceResults) {
            $statusClass = if ($result.Status -eq 'OK') { 'status-ok' } else { 'status-failed' }
            
            $pingClass = ""
            $pingValue = "-"
            if (-not $SkipPing -and $result.PingLatency_ms) {
                $pingValue = $result.PingLatency_ms
                if ($result.PingLatency_ms -lt 50) { $pingClass = "ping-excellent" }
                elseif ($result.PingLatency_ms -lt 100) { $pingClass = "ping-good" }
                else { $pingClass = "ping-poor" }
            }
            
            $speedValue = if (-not $SkipSpeed -and $result.DownloadSpeed_Kbps) { $result.DownloadSpeed_Kbps } else { "-" }
            $ipValue = if ($result.IPAddress) { $result.IPAddress } else { "-" }
            
            $htmlContent += @"
                            <tr>
                                <td>$($result.URL)</td>
                                <td><span class="$statusClass">$($result.Status)</span></td>
                                <td>$ipValue</td>
                                <td><span class="severity severity-$($result.ImpactSeverity)">$($result.ImpactSeverity)</span> $(ConvertTo-HtmlText "$($result.ImpactArea): $($result.ImpactFeature)")</td>
"@

            if (-not $SkipPing) {
                $htmlContent += "<td class=`"$pingClass`">$pingValue</td>"
            }
            
            if (-not $SkipSpeed) {
                $htmlContent += "<td>$speedValue</td>"
            }

            $htmlContent += "</tr>"
        }

        $htmlContent += "</tbody></table>"
        
        # Add service impact information if there are failures
        if ($serviceFailures -gt 0 -and $ServiceImpacts.ContainsKey($serviceName)) {
            $impact = $ServiceImpacts[$serviceName]
            $htmlContent += @"
                    <div class="impact-section">
                        <div class="impact-title">[WARN] Possible Impact of Failures:</div>
                        <ul class="impact-list">
                            <li><strong>Functional Impact:</strong> $($impact.Impact)</li>
                            <li><strong>Possible Symptoms:</strong> $($impact.Symptoms)</li>
                        </ul>
                    </div>
"@
        }
        
        $htmlContent += "</div>"
    }

    $htmlContent += "</div>"

    # Add performance statistics section if available
    if ($pingStats -or $speedStats) {
        $htmlContent += @"
            <div class="section">
                <h2>Performance Statistics</h2>
"@
        
        if ($pingStats) {
            $htmlContent += @"
                <div class="service-group">
                    <div class="service-header">Ping Latency Statistics</div>
                    <table>
                        <tr>
                            <th>Metric</th>
                            <th>Value</th>
                            <th>Rating</th>
                        </tr>
                        <tr>
                            <td>Average Latency</td>
                            <td class="$(if($pingStats.Average -lt 50){"ping-excellent"}elseif($pingStats.Average -lt 100){"ping-good"}else{"ping-poor"})">$($pingStats.Average) ms</td>
                            <td>$(if($pingStats.Average -lt 50){"Excellent"}elseif($pingStats.Average -lt 100){"Good"}else{"[WARN] Needs Improvement"})</td>
                        </tr>
                        <tr>
                            <td>Best Latency</td>
                            <td class="ping-excellent">$($pingStats.Min) ms</td>
                            <td>-</td>
                        </tr>
                        <tr>
                            <td>Worst Latency</td>
                            <td class="ping-poor">$($pingStats.Max) ms</td>
                            <td>-</td>
                        </tr>
                        <tr>
                            <td>Tested Hosts</td>
                            <td>$($pingStats.Count)</td>
                            <td>-</td>
                        </tr>
                    </table>
                </div>
"@
        }

        if ($speedStats) {
            $htmlContent += @"
                <div class="service-group">
                    <div class="service-header">Response Time Statistics</div>
                    <table>
                        <tr>
                            <th>Metric</th>
                            <th>Value</th>
                        </tr>
                        <tr>
                            <td>Average Response Time</td>
                            <td>$($speedStats.Average) ms</td>
                        </tr>
                        <tr>
                            <td>Best Performance</td>
                            <td>$($speedStats.Min) ms</td>
                        </tr>
                        <tr>
                            <td>Worst Performance</td>
                            <td>$($speedStats.Max) ms</td>
                        </tr>
                        <tr>
                            <td>Tested Endpoints</td>
                            <td>$($speedStats.Count)</td>
                        </tr>
                    </table>
                </div>
"@
        }

        $htmlContent += "</div>"
    }

    # Footer
    $computerName = $env:COMPUTERNAME
    $userName = $env:USERNAME
    $htmlContent += @"
        </div>
        <div class="footer">
            <p><strong>Microsoft Endpoint Connectivity Tester V2.1</strong> | Executed on: $computerName by $userName</p>
            <p>Generated by PowerShell Script: CheckMicrosoftEndpointsV2.ps1 | © 2025 Ronny Alhelm</p>
        </div>
    </div>
</body>
</html>
"@

    # Write HTML file
    try {
        $htmlContent | Out-File -FilePath $FilePath -Encoding UTF8
        return $true
    }
    catch {
        Write-Error "Error creating HTML file: $_"
        return $false
    }
}

# Filter services based on selection
$selectedBackendUrls = @{}

if ($Services -contains 'All') {
    $selectedBackendUrls = $backendUrls
    if (-not $Quiet) {
        Write-Host "All Microsoft Services have been selected for testing" -ForegroundColor Green
    }
} else {
    # Service name mapping for parameter validation
    $serviceMapping = @{
        'WindowsUpdate' = 'Windows Update for Business'
        'Autopatch' = 'Windows Autopatch'
        'Intune' = 'Intune'
        'Defender' = 'Microsoft Defender'
        'AzureAD' = 'Azure Active Directory'
        'Microsoft365' = 'Microsoft 365'
        'Store' = 'Microsoft Store'
        'Activation' = 'Windows Activation'
        'Edge' = 'Microsoft Edge'
        'Telemetry' = 'Windows Telemetry'
    }
    
    foreach ($service in $Services) {
        $serviceName = $serviceMapping[$service]
        if ($serviceName -and $backendUrls.ContainsKey($serviceName)) {
            $selectedBackendUrls[$serviceName] = $backendUrls[$serviceName]
            if (-not $Quiet) {
                Write-Host "[OK] $serviceName has been selected for testing" -ForegroundColor Cyan
            }
        } else {
            Write-Warning "Service '$service' not found or invalid"
        }
    }
}

$unsupportedServices = @()
foreach ($serviceName in @($selectedBackendUrls.Keys)) {
    if (-not (Test-ServiceSupportedOnCurrentPlatform -ServiceName $serviceName)) {
        $unsupportedServices += $serviceName
        $selectedBackendUrls.Remove($serviceName)
    }
}

if ($unsupportedServices.Count -gt 0 -and -not $Quiet) {
    Write-Warning "The following services are Windows-only and were skipped on this platform: $($unsupportedServices -join ', ')"
}

if ($selectedBackendUrls.Count -eq 0) {
    Write-Error "No valid services selected. Script will exit."
    exit 1
}

# Display test configuration
if (-not $Quiet) {
    Write-Host "`nTEST CONFIGURATION:" -ForegroundColor Yellow
    Write-Host "   • Selected Services: $($selectedBackendUrls.Keys -join ', ')" -ForegroundColor White
    Write-Host "   • Ping Tests: $(if($SkipPing){'Skipped'}else{'Enabled'})" -ForegroundColor $(if($SkipPing){'Red'}else{'Green'})
    Write-Host "   • Speed Tests: $(if($SkipSpeed){'Skipped'}else{'Enabled'})" -ForegroundColor $(if($SkipSpeed){'Red'}else{'Green'})
    Write-Host "   • Mode: $(if($Quiet){'Quiet'}else{'Verbose'})" -ForegroundColor $(if($Quiet){'Gray'}else{'Cyan'})
}

# Flatten selected URLs for testing
$urls = $selectedBackendUrls.Values | ForEach-Object { $_ } 

# Remove wildcards for direct test (cannot resolve * in Test-NetConnection)
$seenTestUrls = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
$script:skippedWildcardUrls = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
$testUrls = foreach ($url in $urls) {
    # Wildcard entries (*.domain) are firewall/proxy rules, not resolvable hosts.
    # Testing a made-up host such as www.<domain> mostly fails with NXDOMAIN and
    # reports false errors, so wildcards are skipped; their concrete hosts are tested.
    if ($url -match '\*') {
        [void]$script:skippedWildcardUrls.Add($url)
        continue
    }
    $normalizedUrl = $url

    if ($seenTestUrls.Add($normalizedUrl)) {
        $normalizedUrl
    }
}

if (-not $Quiet -and $script:skippedWildcardUrls.Count -gt 0) {
    Write-Host "`nSkipped $($script:skippedWildcardUrls.Count) wildcard entries (allow them as *.domain rules in firewall/proxy):" -ForegroundColor DarkGray
    $script:skippedWildcardUrls | Sort-Object | ForEach-Object { Write-Host "   • $_" -ForegroundColor DarkGray }
}

# Service each URL belongs to (first match), used for the impact fallback
$urlServiceMap = @{}
foreach ($serviceName in ($selectedBackendUrls.Keys | Sort-Object)) {
    foreach ($serviceUrl in $selectedBackendUrls[$serviceName]) {
        if (-not $urlServiceMap.ContainsKey($serviceUrl)) { $urlServiceMap[$serviceUrl] = $serviceName }
    }
}

$m365Lookup = $null
if (-not $SkipM365Lookup) {
    if (-not $Quiet) {
        Write-Host "`nLoading Microsoft 365 endpoint categories from endpoints.office.com..." -ForegroundColor DarkGray
    }
    $m365Lookup = Get-M365EndpointLookup
    if (-not $m365Lookup -and -not $Quiet) {
        Write-Host "   Not available, using built-in rules only." -ForegroundColor DarkGray
    }
}

# Store script start time for duration calculation
$script:startTime = Get-Date

$results = @()

# Build test description
$testTypes = @("Connectivity")
if (-not $SkipPing) { $testTypes += "Ping" }
if (-not $SkipSpeed) { $testTypes += "Speed" }

if (-not $Quiet) {
    Write-Host "`nStarting endpoint tests ($($testTypes -join ', '))..." -ForegroundColor Cyan
    
    $estimatedTime = $testUrls.Count * (1 + $(if(-not $SkipPing){2}else{0}) + $(if(-not $SkipSpeed){3}else{0}))
    Write-Host "Estimated duration: approx. $([math]::Round($estimatedTime/60, 1)) minutes for $($testUrls.Count) endpoints" -ForegroundColor Yellow
    Write-Host ""
}

$totalUrls = $testUrls.Count
$currentUrl = 0

foreach ($url in $testUrls) {
    $currentUrl++
    $testHost = ($url -replace 'https://', '').Split('/')[0]
    
    # Progress indicator
    Write-Progress -Activity "Testing Microsoft Endpoints" -Status "Testing $testHost ($currentUrl of $totalUrls)" -PercentComplete (($currentUrl / $totalUrls) * 100)
    
    # Test basic connectivity
    $connectivity = Test-TcpConnectivity -HostName $testHost -Port 443
    $status = if ($connectivity.TcpTestSucceeded) { 'OK' } else { 'FAILED' }
    
    # Initialize additional test results
    $pingLatency = $null
    $downloadSpeed = $null
    
    # If basic connectivity works, test ping and speed based on parameters
    if ($connectivity.TcpTestSucceeded) {
        if (-not $Quiet) {
            Write-Host "  [OK] $testHost - Connection OK" -ForegroundColor Green
        }
        
        # Test ping if not skipped
        if (-not $SkipPing) {
            if (-not $Quiet) {
                Write-Host "  $testHost - Testing Ping..." -ForegroundColor Blue
            }
            $pingLatency = Test-PingLatency -HostName $testHost
        }
        
        # Test speed if not skipped
        if (-not $SkipSpeed) {
            if (-not $Quiet) {
                Write-Host "  $testHost - Testing Speed..." -ForegroundColor Magenta
            }
            $downloadSpeed = Test-DownloadSpeed -Url $url
        }
    } else {
        if (-not $Quiet) {
            Write-Host "  [FAIL] $testHost - Not reachable" -ForegroundColor Red
        }
    }
    
    $endpointImpact = Get-EndpointImpact -HostName $testHost -ServiceName $urlServiceMap[$url] -M365Info (Get-M365EndpointInfo -HostName $testHost -Lookup $m365Lookup)

    $results += [PSCustomObject]@{
        URL = $url
        Host = $testHost
        Status = $status
        IPAddress = $connectivity.RemoteAddress
        PingLatency_ms = $pingLatency
        DownloadSpeed_Kbps = $downloadSpeed
        Service = $urlServiceMap[$url]
        ImpactArea = $endpointImpact.Area
        ImpactFeature = $endpointImpact.Feature
        ImpactSeverity = $endpointImpact.Severity
        ImpactText = $endpointImpact.Impact
    }
}

Write-Progress -Activity "Testing Microsoft Endpoints" -Completed

$impactSummary = Get-ServiceImpactSummary -Results $results

# Output grouped results by backend with enhanced information
$failedServices = @()
foreach ($backend in $selectedBackendUrls.Keys) {
    if (-not $Quiet) {
        Write-Host "`nBackend: $backend" -ForegroundColor Cyan
        Write-Host "="*60 -ForegroundColor DarkCyan
    }
    
    $backendResults = $results | Where-Object { $selectedBackendUrls[$backend] -contains $_.URL }
    
    # Display results with dynamic formatting based on enabled tests
    $tableColumns = @(
        @{Name="URL"; Expression={$_.URL}; Width=40},
        @{Name="Status"; Expression={$_.Status}; Width=8},
        @{Name="IP Address"; Expression={$_.IPAddress}; Width=15}
    )
    
    if (-not $SkipPing) {
        $tableColumns += @{Name="Ping (ms)"; Expression={if($_.PingLatency_ms) {"$($_.PingLatency_ms)"} else {"-"}}; Width=10}
    }
    
    if (-not $SkipSpeed) {
        $tableColumns += @{Name="Response (ms)"; Expression={if($_.DownloadSpeed_Kbps) {"$($_.DownloadSpeed_Kbps)"} else {"-"}}; Width=12}
    }
    
    if (-not $Quiet) {
        $backendResults | Format-Table $tableColumns -AutoSize
    }
    
    # Performance analysis (only if tests were performed)
    $successfulEndpoints = $backendResults | Where-Object { $_.Status -eq 'OK' }
    if ($successfulEndpoints.Count -gt 0 -and (-not $SkipPing -or -not $SkipSpeed)) {
        
        if (-not $SkipPing) {
            $avgPing = ($successfulEndpoints | Where-Object { $_.PingLatency_ms -ne $null } | Measure-Object PingLatency_ms -Average).Average
            if ($avgPing) {
                $pingColor = if ($avgPing -lt 50) { "Green" } elseif ($avgPing -lt 100) { "Yellow" } else { "Red" }
                Write-Host "Average Latency: $([math]::Round($avgPing, 2)) ms" -ForegroundColor $pingColor
            }
        }
        
        if (-not $SkipSpeed) {
            $avgSpeed = ($successfulEndpoints | Where-Object { $_.DownloadSpeed_Kbps -ne $null } | Measure-Object DownloadSpeed_Kbps -Average).Average
            if ($avgSpeed) {
                $speedColor = if ($avgSpeed -gt 1000) { "Green" } elseif ($avgSpeed -gt 100) { "Yellow" } else { "Red" }
                Write-Host "Average Response Time: $([math]::Round($avgSpeed, 2)) ms" -ForegroundColor $speedColor
            }
        }
    }
    
    $failedEndpoints = $backendResults | Where-Object { $_.Status -eq 'FAILED' }
    if ($failedEndpoints.Count -gt 0) {
        if (-not $Quiet) {
            Write-Host "[FAIL] $($failedEndpoints.Count) endpoint(s) for $backend are NOT reachable!" -ForegroundColor Red
        }
        $failedServices += $backend
        if (-not $Quiet) {
            Show-ServiceImpactWarning -ServiceName $backend -FailedEndpoints $failedEndpoints
        }
    } else {
        if (-not $Quiet) {
            Write-Host "[OK] All endpoints for $backend are reachable." -ForegroundColor Green
        }
        
        # Performance rating (only if ping tests were performed)
        if (-not $SkipPing -and $avgPing) {
            if ($avgPing -lt 50) {
                if (-not $Quiet) { Write-Host "Excellent network performance!" -ForegroundColor Green }
            } elseif ($avgPing -lt 100) {
                if (-not $Quiet) { Write-Host "Good network performance" -ForegroundColor Yellow }
            } else {
                if (-not $Quiet) { Write-Host "[WARN] Slow network performance - check recommended" -ForegroundColor Red }
            }
        }
    }
}

# Summary table of all results with enhanced metrics
if (-not $Quiet) {
    Write-Host "`n" + "="*100 -ForegroundColor White
    Write-Host "OVERALL SUMMARY OF ALL TESTED ENDPOINTS" -ForegroundColor White
    Write-Host "="*100 -ForegroundColor White

    # Create dynamic summary table columns
    $summaryColumns = @(
        @{Name="URL"; Expression={$_.URL}; Width=45},
        @{Name="Host"; Expression={$_.Host}; Width=30},
        @{Name="Status"; Expression={$_.Status}; Width=8},
        @{Name="IP Address"; Expression={$_.IPAddress}; Width=15}
    )
    
    if (-not $SkipPing) {
        $summaryColumns += @{Name="Ping (ms)"; Expression={if($_.PingLatency_ms) {"$($_.PingLatency_ms)"} else {"-"}}; Width=10}
    }
    
    if (-not $SkipSpeed) {
        $summaryColumns += @{Name="Response (ms)"; Expression={if($_.DownloadSpeed_Kbps) {"$($_.DownloadSpeed_Kbps)"} else {"-"}}; Width=12}
    }

    $results | Format-Table $summaryColumns -AutoSize
}

# Overall performance statistics
if (-not $Quiet -and (-not $SkipPing -or -not $SkipSpeed)) {
    Write-Host "`nOVERALL NETWORK PERFORMANCE STATISTICS:" -ForegroundColor Cyan
    Write-Host "-"*60 -ForegroundColor DarkCyan

    $successfulTests = $results | Where-Object { $_.Status -eq 'OK' }
    
    if (-not $SkipPing) {
        $testsWithPing = $successfulTests | Where-Object { $_.PingLatency_ms -ne $null }
        
        if ($testsWithPing.Count -gt 0) {
            $overallAvgPing = ($testsWithPing | Measure-Object PingLatency_ms -Average).Average
            $minPing = ($testsWithPing | Measure-Object PingLatency_ms -Minimum).Minimum
            $maxPing = ($testsWithPing | Measure-Object PingLatency_ms -Maximum).Maximum
            
            Write-Host "Ping Statistics:" -ForegroundColor Yellow
            Write-Host "   • Average: $([math]::Round($overallAvgPing, 2)) ms" -ForegroundColor White
            Write-Host "   • Minimum: $([math]::Round($minPing, 2)) ms" -ForegroundColor Green
            Write-Host "   • Maximum: $([math]::Round($maxPing, 2)) ms" -ForegroundColor Red
            Write-Host "   • Tested Hosts: $($testsWithPing.Count)" -ForegroundColor Cyan
        }
    }

    if (-not $SkipSpeed) {
        $testsWithSpeed = $successfulTests | Where-Object { $_.DownloadSpeed_Kbps -ne $null }
        
        if ($testsWithSpeed.Count -gt 0) {
            $overallAvgSpeed = ($testsWithSpeed | Measure-Object DownloadSpeed_Kbps -Average).Average
            $minSpeed = ($testsWithSpeed | Measure-Object DownloadSpeed_Kbps -Minimum).Minimum
            $maxSpeed = ($testsWithSpeed | Measure-Object DownloadSpeed_Kbps -Maximum).Maximum
            
            Write-Host "Response Time Statistics:" -ForegroundColor Yellow
            Write-Host "   • Average: $([math]::Round($overallAvgSpeed, 2)) ms" -ForegroundColor White
            Write-Host "   • Minimum: $([math]::Round($minSpeed, 2)) ms" -ForegroundColor Green
            Write-Host "   • Maximum: $([math]::Round($maxSpeed, 2)) ms" -ForegroundColor Red
            Write-Host "   • Tested Endpoints: $($testsWithSpeed.Count)" -ForegroundColor Cyan
        }
    }
}

# Generate HTML report if requested
if ($HtmlReport) {
    if (-not $HtmlReport.EndsWith('.html')) {
        $HtmlReport += '.html'
    }
    
    if (-not $Quiet) {
        Write-Host "`nGenerating HTML Report..." -ForegroundColor Cyan
    }
    
    $htmlSuccess = New-HtmlReport -Results $results -SelectedServices $selectedBackendUrls -FailedServices $failedServices -ServiceImpacts $serviceImpacts -ImpactSummary $impactSummary -SkipPing $SkipPing -SkipSpeed $SkipSpeed -FilePath $HtmlReport
    
    if ($htmlSuccess) {
        $fullPath = (Resolve-Path $HtmlReport -ErrorAction SilentlyContinue).Path
        if (-not $fullPath) {
            $fullPath = (Get-Item $HtmlReport -ErrorAction SilentlyContinue).FullName
        }
        if (-not $fullPath) {
            $fullPath = $HtmlReport
        }
        
        if (-not $Quiet) {
            Write-Host "[OK] HTML Report successfully created: $fullPath" -ForegroundColor Green
        }
        
        # Open report in browser if requested
        if ($OpenReport) {
            try {
                if (Test-IsWindowsPlatform) {
                    Start-Process $fullPath
                }
                elseif (Test-IsMacOSPlatform -and (Get-Command open -ErrorAction SilentlyContinue)) {
                    & open $fullPath
                }
                elseif (Test-IsLinuxPlatform -and (Get-Command xdg-open -ErrorAction SilentlyContinue)) {
                    & xdg-open $fullPath
                }
                else {
                    Start-Process $fullPath
                }
                if (-not $Quiet) {
                    Write-Host "Opening HTML Report in browser..." -ForegroundColor Cyan
                }
            }
            catch {
                Write-Warning "Could not open HTML Report in browser: $_"
            }
        }
    } else {
        Write-Error "Error creating HTML Report"
    }
}

# Final summary with service impact overview
if ($results.Status -contains 'FAILED') {
    $totalFailed = ($results | Where-Object { $_.Status -eq 'FAILED' }).Count
    $totalTested = $results.Count
    
    if (-not $Quiet) {
        Write-Host "`nWARNING: $totalFailed of $totalTested Microsoft endpoints are NOT reachable!" -ForegroundColor Red
        Write-Host "="*80 -ForegroundColor Red
        
        Show-ServiceImpactSummary -ImpactSummary $impactSummary
        Write-Host ""

        if ($failedServices.Count -gt 0) {
            Write-Host "AFFECTED SERVICES AND POSSIBLE IMPACTS:" -ForegroundColor Yellow
            Write-Host "-"*50 -ForegroundColor Yellow
            foreach ($service in $failedServices) {
                Write-Host "$service" -ForegroundColor Red
                if ($serviceImpacts.ContainsKey($service)) {
                    Write-Host "   └─ $($serviceImpacts[$service].Impact)" -ForegroundColor Yellow
                }
            }
            Write-Host "`nRECOMMENDATION: Check your firewall and proxy configuration for the services mentioned above." -ForegroundColor Cyan
        }
    }
    exit 1
} else {
    if (-not $Quiet) {
        Write-Host "`nSUCCESS: All $($results.Count) Microsoft endpoints are reachable!" -ForegroundColor Green
        Write-Host "All Microsoft Cloud Services should function properly." -ForegroundColor Green
    }
    exit 0
}
