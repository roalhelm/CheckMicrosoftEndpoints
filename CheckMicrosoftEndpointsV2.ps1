
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
    [switch]$OpenReport
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

.NOTES
    File Name     : CheckMicrosoftEndpointsV2.ps1
    Author        : Ronny Alhelm
    Version       : 2.1
    Creation Date : October 13, 2025
    Last Modified : October 14, 2025
    Requirements  : PowerShell 5.1 or higher
    Network       : Internet connectivity required
    Ports         : HTTPS (443) access to Microsoft endpoints, ICMP for ping tests
    Permissions   : No elevated privileges required
    Duration      : 1-10 minutes depending on selected services and test options

.CHANGES
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
    .\CheckMicrosoftEndpointsV2.ps1 -Services Intune,Defender -HtmlReport -Quiet

.EXAMPLE
    # Full test with comprehensive HTML report
    .\CheckMicrosoftEndpointsV2.ps1 -Services All -HtmlReport "C:\Reports\Microsoft-Endpoints.html"

#>



# Function to show interactive service selection menu
function Show-ServiceSelectionMenu {
    $selectedServices = @()
    
    Write-Host "`n🔧 MICROSOFT ENDPOINT CONNECTIVITY TESTER V2.0" -ForegroundColor Cyan
    Write-Host "="*60 -ForegroundColor DarkCyan
    Write-Host "Wählen Sie die zu testenden Microsoft Services aus:" -ForegroundColor Yellow
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
        @{ Key = "A"; Name = "ALLE SERVICES"; Param = "All" }
    )
    
    foreach ($option in $menuOptions) {
        if ($option.Key -eq "A") {
            Write-Host ""
            Write-Host "[$($option.Key)]  $($option.Name) 🚀" -ForegroundColor Green
        } else {
            Write-Host "[$($option.Key)]   $($option.Name)" -ForegroundColor White
        }
    }
    
    Write-Host ""
    Write-Host "Geben Sie die Nummern der gewünschten Services ein (z.B. 1,3,5 oder A für alle):" -ForegroundColor Cyan
    Write-Host "Zusätzliche Optionen:" -ForegroundColor Yellow
    Write-Host "  - Drücken Sie 'P' um Ping-Tests zu überspringen" -ForegroundColor Gray
    Write-Host "  - Drücken Sie 'S' um Geschwindigkeitstests zu überspringen" -ForegroundColor Gray
    Write-Host "  - Drücken Sie 'H' um HTML-Report zu generieren" -ForegroundColor Gray
    Write-Host "  - Drücken Sie 'Q' für stillen Modus" -ForegroundColor Gray
    
    $userInput = Read-Host "`nIhre Auswahl"
    
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
        Write-Host "Keine gültigen Services ausgewählt. Verwende alle Services." -ForegroundColor Yellow
        return @("All")
    }
    
    return $selectedServices
}

# Handle interactive mode
if ($Services -contains 'Interactive') {
    $Services = Show-ServiceSelectionMenu
    
    # Ask for additional options
    Write-Host ""
    $skipOptions = Read-Host "Optionen: (P)ing überspringen, (S)peed überspringen, (H)TML-Report, (Q)uiet, oder Enter für Standard"
    
    if ($skipOptions -match 'P') { $SkipPing = $true }
    if ($skipOptions -match 'S') { $SkipSpeed = $true }
    if ($skipOptions -match 'Q') { $Quiet = $true }
    if ($skipOptions -match 'H') { 
        $HtmlReport = "Microsoft-Endpoints-Report-$(Get-Date -Format 'yyyyMMdd-HHmmss').html"
        $htmlChoice = Read-Host "HTML-Report im Browser öffnen? (j/n)"
        if ($htmlChoice -match '^(j|y|yes|ja)') { $OpenReport = $true }
    }
    
    Clear-Host
}

# List of required Microsoft service endpoints (as of 2025)
$backendUrls = @{
    'Windows Update for Business' = @(
        'https://dl.delivery.mp.microsoft.com',
        'https://download.windowsupdate.com',
        'https://fe3.delivery.mp.microsoft.com',
        'https://sws.update.microsoft.com',
        'https://update.microsoft.com',
        'https://windowsupdate.microsoft.com',
        'https://www.delivery.mp.microsoft.com',
        'https://www.dl.delivery.mp.microsoft.com',
        'https://www.update.microsoft.com',
        'https://www.windowsupdate.com',
        'https://www.windowsupdate.microsoft.com',
        'https://*.delivery.mp.microsoft.com',
        'https://*.dl.delivery.mp.microsoft.com',
        'https://*.update.microsoft.com',
        'https://*.windowsupdate.com',
        'https://*.windowsupdate.microsoft.com'
    ) | Sort-Object -Unique
    
    'Windows Autopatch' = @(
        'https://api.update.microsoft.com',
        'https://autopatch.microsoft.com',
        'https://config.edge.skype.com',
        'https://graph.microsoft.com',
        'https://login.microsoftonline.com',
        'https://manage.microsoft.com',
        'https://nexusrules.officeapps.live.com'
    ) | Sort-Object -Unique
    
    'Intune' = @(
        'https://device.login.microsoftonline.com',
        'https://endpoint.microsoft.com',
        'https://graph.microsoft.com',
        'https://login.microsoftonline.com',
        'https://manage.microsoft.com',
        'https://portal.manage.microsoft.com',
        'https://enrollment.manage.microsoft.com',
        'https://enterpriseregistration.windows.net',
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
        'https://policykeyservice.dc.ad.msft.net',
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
        'https://purchase.mp.microsoft.com'
    ) | Sort-Object -Unique
    
    'Windows Activation' = @(
        'https://activation.sls.microsoft.com',
        'https://crl.microsoft.com',
        'https://validation.sls.microsoft.com',
        'https://activation-v2.sls.microsoft.com',
        'https://purchase.mp.microsoft.com',
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
        'Impact' = 'Windows Updates und Feature Updates können nicht heruntergeladen werden'
        'Symptoms' = 'Update-Fehler, veraltete Sicherheitspatches, fehlende Feature-Updates'
    }
    'Windows Autopatch' = @{
        'Impact' = 'Automatische Patch-Verwaltung funktioniert nicht'
        'Symptoms' = 'Manuelle Update-Verwaltung erforderlich, keine automatischen Deployment-Zeitpläne'
    }
    'Intune' = @{
        'Impact' = 'Geräte-Enrollment, App-Deployment und Compliance-Checks funktionieren nicht'
        'Symptoms' = 'Neue Geräte können nicht registriert werden, Apps installieren nicht, Compliance-Berichte fehlen'
    }
    'Microsoft Defender' = @{
        'Impact' = 'Antivirus-Updates und Cloud-Schutz funktionieren nicht'
        'Symptoms' = 'Veraltete Virensignaturen, keine Cloud-basierten Bedrohungsanalysen, eingeschränkter Schutz'
    }
    'Azure Active Directory' = @{
        'Impact' = 'Anmeldung und Geräte-Authentication können fehlschlagen'
        'Symptoms' = 'Login-Probleme, Geräte-Registrierung fehlgeschlagen, SSO funktioniert nicht'
    }
    'Microsoft 365' = @{
        'Impact' = 'Office-Apps, SharePoint und OneDrive funktionieren nicht vollständig'
        'Symptoms' = 'Office-Apps starten nicht, Dokumente synchronisieren nicht, Admin Center nicht erreichbar'
    }
    'Microsoft Store' = @{
        'Impact' = 'App-Installation und -Updates aus dem Microsoft Store funktionieren nicht'
        'Symptoms' = 'Store öffnet nicht, Apps können nicht installiert/aktualisiert werden'
    }
    'Windows Activation' = @{
        'Impact' = 'Windows-Aktivierung und Lizenzvalidierung funktionieren nicht'
        'Symptoms' = 'Aktivierungsfehler, Lizenz-Warnungen, eingeschränkte Funktionalität'
    }
    'Microsoft Edge' = @{
        'Impact' = 'Browser-Updates und Enterprise-Features funktionieren eingeschränkt'
        'Symptoms' = 'SmartScreen deaktiviert, keine automatischen Updates, Enterprise-Policies funktionieren nicht'
    }
    'Windows Telemetry' = @{
        'Impact' = 'Diagnose-Daten und Fehlerberichte werden nicht übertragen'
        'Symptoms' = 'Keine Windows-Diagnosedaten, Fehlerberichte gehen verloren, Update-Qualität kann beeinträchtigt sein'
    }
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

# Function to test download speed (simplified test)
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
                $speedBps = ($data.Length * 8) / ($stopwatch.ElapsedMilliseconds / 1000)
                $speedKbps = [math]::Round($speedBps / 1024, 2)
                return $speedKbps
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
        Write-Host "    ⚠️  AUSWIRKUNGEN:" -ForegroundColor Yellow
        Write-Host "       • $($serviceImpacts[$ServiceName].Impact)" -ForegroundColor Yellow
        Write-Host "    🔍 SYMPTOME:" -ForegroundColor Cyan
        Write-Host "       • $($serviceImpacts[$ServiceName].Symptoms)" -ForegroundColor Cyan
        Write-Host "    🚫 Betroffene Endpunkte: $($FailedEndpoints.Count)" -ForegroundColor Red
        Write-Host ""
    }
}

# Function to generate HTML report
function New-HtmlReport {
    param (
        [array]$Results,
        [hashtable]$SelectedServices,
        [array]$FailedServices,
        [hashtable]$ServiceImpacts,
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
<html lang="de">
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
            <h1>🌐 Microsoft Endpoint Connectivity Report</h1>
            <p>Erstellt am $timestamp | Testdauer: $($testDuration.ToString("hh\:mm\:ss"))</p>
        </div>
        
        <div class="summary">
            <div class="stat-card">
                <div class="stat-value">$totalEndpoints</div>
                <div class="stat-label">Getestete Endpunkte</div>
            </div>
            <div class="stat-card success">
                <div class="stat-value">$successfulEndpoints</div>
                <div class="stat-label">Erfolgreich</div>
            </div>
            <div class="stat-card danger">
                <div class="stat-value">$failedEndpoints</div>
                <div class="stat-label">Fehlgeschlagen</div>
            </div>
            <div class="stat-card info">
                <div class="stat-value">$successRate%</div>
                <div class="stat-label">Erfolgsrate</div>
            </div>
"@

    # Add ping statistics if available
    if ($pingStats) {
        $htmlContent += @"
            <div class="stat-card">
                <div class="stat-value">$($pingStats.Average)ms</div>
                <div class="stat-label">Ø Latenz</div>
            </div>
"@
    }

    # Add speed statistics if available
    if ($speedStats) {
        $htmlContent += @"
            <div class="stat-card">
                <div class="stat-value">$($speedStats.Average)</div>
                <div class="stat-label">Ø Speed (Kbps)</div>
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
                <strong>✅ Ausgezeichnet!</strong> Alle getesteten Microsoft Endpunkte sind erreichbar und funktionsfähig.
            </div>
"@
    } else {
        $htmlContent += @"
            <div class="alert alert-danger">
                <strong>⚠️ Warnung!</strong> $failedEndpoints von $totalEndpoints Microsoft Endpunkten sind nicht erreichbar. 
                Dies kann zu Funktionseinschränkungen bei den betroffenen Services führen.
            </div>
"@
    }

    # Add services section
    $htmlContent += @"
            <div class="section">
                <h2>📋 Service-Details</h2>
"@

    # Generate service-specific tables
    foreach ($serviceName in $SelectedServices.Keys) {
        $serviceResults = $Results | Where-Object { $SelectedServices[$serviceName] -contains $_.URL }
        $serviceFailures = ($serviceResults | Where-Object { $_.Status -eq 'FAILED' }).Count

        
        $statusClass = if ($serviceFailures -eq 0) { "success" } else { "danger" }
        $statusText = if ($serviceFailures -eq 0) { "✅ Alle OK" } else { "❌ $serviceFailures Fehler" }
        
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
"@

        if (-not $SkipPing) {
            $htmlContent += "<th>Latenz (ms)</th>"
        }
        
        if (-not $SkipSpeed) {
            $htmlContent += "<th>Speed (Kbps)</th>"
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
                        <div class="impact-title">⚠️ Mögliche Auswirkungen bei Fehlern:</div>
                        <ul class="impact-list">
                            <li><strong>Funktionsbeeinträchtigung:</strong> $($impact.Impact)</li>
                            <li><strong>Mögliche Symptome:</strong> $($impact.Symptoms)</li>
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
                <h2>📊 Performance-Statistiken</h2>
"@
        
        if ($pingStats) {
            $htmlContent += @"
                <div class="service-group">
                    <div class="service-header">🏓 Ping-Latenz Statistiken</div>
                    <table>
                        <tr>
                            <th>Metrik</th>
                            <th>Wert</th>
                            <th>Bewertung</th>
                        </tr>
                        <tr>
                            <td>Durchschnittliche Latenz</td>
                            <td class="$(if($pingStats.Average -lt 50){"ping-excellent"}elseif($pingStats.Average -lt 100){"ping-good"}else{"ping-poor"})">$($pingStats.Average) ms</td>
                            <td>$(if($pingStats.Average -lt 50){"🌟 Exzellent"}elseif($pingStats.Average -lt 100){"👍 Gut"}else{"⚠️ Verbesserungsbedarf"})</td>
                        </tr>
                        <tr>
                            <td>Beste Latenz</td>
                            <td class="ping-excellent">$($pingStats.Min) ms</td>
                            <td>-</td>
                        </tr>
                        <tr>
                            <td>Schlechteste Latenz</td>
                            <td class="ping-poor">$($pingStats.Max) ms</td>
                            <td>-</td>
                        </tr>
                        <tr>
                            <td>Getestete Hosts</td>
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
                    <div class="service-header">🚀 Geschwindigkeits-/Antwortzeit Statistiken</div>
                    <table>
                        <tr>
                            <th>Metrik</th>
                            <th>Wert</th>
                        </tr>
                        <tr>
                            <td>Durchschnittliche Speed/Antwortzeit</td>
                            <td>$($speedStats.Average) Kbps/ms</td>
                        </tr>
                        <tr>
                            <td>Beste Performance</td>
                            <td>$($speedStats.Min) Kbps/ms</td>
                        </tr>
                        <tr>
                            <td>Schlechteste Performance</td>
                            <td>$($speedStats.Max) Kbps/ms</td>
                        </tr>
                        <tr>
                            <td>Getestete Endpoints</td>
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
            <p><strong>Microsoft Endpoint Connectivity Tester V2.0</strong> | Ausgeführt auf: $computerName von $userName</p>
            <p>Generiert von PowerShell Script: CheckMicrosoftEndpointsV2.ps1 | © 2025 Ronny Alhelm</p>
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
        Write-Error "Fehler beim Erstellen der HTML-Datei: $_"
        return $false
    }
}

# Filter services based on selection
$selectedBackendUrls = @{}

if ($Services -contains 'All') {
    $selectedBackendUrls = $backendUrls
    if (-not $Quiet) {
        Write-Host "🚀 Alle Microsoft Services wurden für Tests ausgewählt" -ForegroundColor Green
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
                Write-Host "✅ $serviceName wurde für Tests ausgewählt" -ForegroundColor Cyan
            }
        } else {
            Write-Warning "Service '$service' nicht gefunden oder ungültig"
        }
    }
}

if ($selectedBackendUrls.Count -eq 0) {
    Write-Error "Keine gültigen Services ausgewählt. Script wird beendet."
    exit 1
}

# Display test configuration
if (-not $Quiet) {
    Write-Host "`n📋 TEST-KONFIGURATION:" -ForegroundColor Yellow
    Write-Host "   • Ausgewählte Services: $($selectedBackendUrls.Keys -join ', ')" -ForegroundColor White
    Write-Host "   • Ping-Tests: $(if($SkipPing){'❌ Übersprungen'}else{'✅ Aktiviert'})" -ForegroundColor $(if($SkipPing){'Red'}else{'Green'})
    Write-Host "   • Geschwindigkeits-Tests: $(if($SkipSpeed){'❌ Übersprungen'}else{'✅ Aktiviert'})" -ForegroundColor $(if($SkipSpeed){'Red'}else{'Green'})
    Write-Host "   • Modus: $(if($Quiet){'🔇 Still'}else{'🔊 Verbose'})" -ForegroundColor $(if($Quiet){'Gray'}else{'Cyan'})
}

# Flatten selected URLs for testing
$urls = $selectedBackendUrls.Values | ForEach-Object { $_ } 

# Remove wildcards for direct test (cannot resolve * in Test-NetConnection)
$testUrls = $urls | ForEach-Object {
    if ($_ -like '*.*.*.*') { $_ } else { $_ -replace '\*\.', 'www.' }
}

# Store script start time for duration calculation
$script:startTime = Get-Date

$results = @()

# Build test description
$testTypes = @("Konnektivität")
if (-not $SkipPing) { $testTypes += "Ping" }
if (-not $SkipSpeed) { $testTypes += "Geschwindigkeit" }

if (-not $Quiet) {
    Write-Host "`n🔍 Starte Endpunkt-Tests ($($testTypes -join ', '))..." -ForegroundColor Cyan
    
    $estimatedTime = $testUrls.Count * (1 + $(if(-not $SkipPing){2}else{0}) + $(if(-not $SkipSpeed){3}else{0}))
    Write-Host "Geschätzte Dauer: ca. $([math]::Round($estimatedTime/60, 1)) Minuten für $($testUrls.Count) Endpunkte" -ForegroundColor Yellow
    Write-Host ""
}

$totalUrls = $testUrls.Count
$currentUrl = 0

foreach ($url in $testUrls) {
    $currentUrl++
    $testHost = ($url -replace 'https://', '').Split('/')[0]
    
    # Progress indicator
    Write-Progress -Activity "Teste Microsoft Endpunkte" -Status "Teste $testHost ($currentUrl von $totalUrls)" -PercentComplete (($currentUrl / $totalUrls) * 100)
    
    # Test basic connectivity
    $result = Test-NetConnection -ComputerName $testHost -Port 443 -WarningAction SilentlyContinue
    $status = if ($result.TcpTestSucceeded) { 'OK' } else { 'FAILED' }
    
    # Initialize additional test results
    $pingLatency = $null
    $downloadSpeed = $null
    
    # If basic connectivity works, test ping and speed based on parameters
    if ($result.TcpTestSucceeded) {
        if (-not $Quiet) {
            Write-Host "  ✅ $testHost - Verbindung OK" -ForegroundColor Green
        }
        
        # Test ping if not skipped
        if (-not $SkipPing) {
            if (-not $Quiet) {
                Write-Host "  🏓 $testHost - Teste Ping..." -ForegroundColor Blue
            }
            $pingLatency = Test-PingLatency -HostName $testHost
        }
        
        # Test speed if not skipped
        if (-not $SkipSpeed) {
            if (-not $Quiet) {
                Write-Host "  📊 $testHost - Teste Geschwindigkeit..." -ForegroundColor Magenta
            }
            $downloadSpeed = Test-DownloadSpeed -Url $url
        }
    } else {
        if (-not $Quiet) {
            Write-Host "  ❌ $testHost - Nicht erreichbar" -ForegroundColor Red
        }
    }
    
    $results += [PSCustomObject]@{
        URL = $url
        Host = $testHost
        Status = $status
        IPAddress = $result.RemoteAddress
        PingLatency_ms = $pingLatency
        DownloadSpeed_Kbps = $downloadSpeed
    }
}

Write-Progress -Activity "Teste Microsoft Endpunkte" -Completed

# Output grouped results by backend with enhanced information
$failedServices = @()
foreach ($backend in $selectedBackendUrls.Keys) {
    Write-Host "`nBackend: $backend" -ForegroundColor Cyan
    Write-Host "="*60 -ForegroundColor DarkCyan
    
    $backendResults = $results | Where-Object { $selectedBackendUrls[$backend] -contains $_.URL }
    
    # Display results with dynamic formatting based on enabled tests
    $tableColumns = @(
        @{Name="URL"; Expression={$_.URL}; Width=40},
        @{Name="Status"; Expression={$_.Status}; Width=8},
        @{Name="IP-Adresse"; Expression={$_.IPAddress}; Width=15}
    )
    
    if (-not $SkipPing) {
        $tableColumns += @{Name="Ping (ms)"; Expression={if($_.PingLatency_ms) {"$($_.PingLatency_ms)"} else {"-"}}; Width=10}
    }
    
    if (-not $SkipSpeed) {
        $tableColumns += @{Name="Speed (Kbps)"; Expression={if($_.DownloadSpeed_Kbps) {"$($_.DownloadSpeed_Kbps)"} else {"-"}}; Width=12}
    }
    
    $backendResults | Format-Table $tableColumns -AutoSize
    
    # Performance analysis (only if tests were performed)
    $successfulEndpoints = $backendResults | Where-Object { $_.Status -eq 'OK' }
    if ($successfulEndpoints.Count -gt 0 -and (-not $SkipPing -or -not $SkipSpeed)) {
        
        if (-not $SkipPing) {
            $avgPing = ($successfulEndpoints | Where-Object { $_.PingLatency_ms -ne $null } | Measure-Object PingLatency_ms -Average).Average
            if ($avgPing) {
                $pingColor = if ($avgPing -lt 50) { "Green" } elseif ($avgPing -lt 100) { "Yellow" } else { "Red" }
                Write-Host "📊 Durchschnittliche Latenz: $([math]::Round($avgPing, 2)) ms" -ForegroundColor $pingColor
            }
        }
        
        if (-not $SkipSpeed) {
            $avgSpeed = ($successfulEndpoints | Where-Object { $_.DownloadSpeed_Kbps -ne $null } | Measure-Object DownloadSpeed_Kbps -Average).Average
            if ($avgSpeed) {
                $speedColor = if ($avgSpeed -gt 1000) { "Green" } elseif ($avgSpeed -gt 100) { "Yellow" } else { "Red" }
                Write-Host "🚀 Durchschnittliche Antwortzeit/Speed: $([math]::Round($avgSpeed, 2)) ms/Kbps" -ForegroundColor $speedColor
            }
        }
    }
    
    $failedEndpoints = $backendResults | Where-Object { $_.Status -eq 'FAILED' }
    if ($failedEndpoints.Count -gt 0) {
        Write-Host "❌ $($failedEndpoints.Count) Endpunkt(e) für $backend sind NICHT erreichbar!" -ForegroundColor Red
        $failedServices += $backend
        Show-ServiceImpactWarning -ServiceName $backend -FailedEndpoints $failedEndpoints
    } else {
        Write-Host "✅ Alle Endpunkte für $backend sind erreichbar." -ForegroundColor Green
        
        # Performance rating (only if ping tests were performed)
        if (-not $SkipPing -and $avgPing) {
            if ($avgPing -lt 50) {
                Write-Host "🌟 Exzellente Netzwerk-Performance!" -ForegroundColor Green
            } elseif ($avgPing -lt 100) {
                Write-Host "👍 Gute Netzwerk-Performance" -ForegroundColor Yellow
            } else {
                Write-Host "⚠️  Langsame Netzwerk-Performance - Überprüfung empfohlen" -ForegroundColor Red
            }
        }
    }
}

# Summary table of all results with enhanced metrics
if (-not $Quiet) {
    Write-Host "`n" + "="*100 -ForegroundColor White
    Write-Host "GESAMTÜBERSICHT ALLER GETESTETEN ENDPUNKTE" -ForegroundColor White
    Write-Host "="*100 -ForegroundColor White

    # Create dynamic summary table columns
    $summaryColumns = @(
        @{Name="URL"; Expression={$_.URL}; Width=45},
        @{Name="Host"; Expression={$_.Host}; Width=30},
        @{Name="Status"; Expression={$_.Status}; Width=8},
        @{Name="IP-Adresse"; Expression={$_.IPAddress}; Width=15}
    )
    
    if (-not $SkipPing) {
        $summaryColumns += @{Name="Ping (ms)"; Expression={if($_.PingLatency_ms) {"$($_.PingLatency_ms)"} else {"-"}}; Width=10}
    }
    
    if (-not $SkipSpeed) {
        $summaryColumns += @{Name="Speed (Kbps)"; Expression={if($_.DownloadSpeed_Kbps) {"$($_.DownloadSpeed_Kbps)"} else {"-"}}; Width=12}
    }

    $results | Format-Table $summaryColumns -AutoSize
}

# Overall performance statistics
if (-not $Quiet -and (-not $SkipPing -or -not $SkipSpeed)) {
    Write-Host "`n📈 GESAMTE NETZWERK-PERFORMANCE STATISTIKEN:" -ForegroundColor Cyan
    Write-Host "-"*60 -ForegroundColor DarkCyan

    $successfulTests = $results | Where-Object { $_.Status -eq 'OK' }
    
    if (-not $SkipPing) {
        $testsWithPing = $successfulTests | Where-Object { $_.PingLatency_ms -ne $null }
        
        if ($testsWithPing.Count -gt 0) {
            $overallAvgPing = ($testsWithPing | Measure-Object PingLatency_ms -Average).Average
            $minPing = ($testsWithPing | Measure-Object PingLatency_ms -Minimum).Minimum
            $maxPing = ($testsWithPing | Measure-Object PingLatency_ms -Maximum).Maximum
            
            Write-Host "🏓 Ping-Statistiken:" -ForegroundColor Yellow
            Write-Host "   • Durchschnitt: $([math]::Round($overallAvgPing, 2)) ms" -ForegroundColor White
            Write-Host "   • Minimum: $([math]::Round($minPing, 2)) ms" -ForegroundColor Green
            Write-Host "   • Maximum: $([math]::Round($maxPing, 2)) ms" -ForegroundColor Red
            Write-Host "   • Getestete Hosts: $($testsWithPing.Count)" -ForegroundColor Cyan
        }
    }

    if (-not $SkipSpeed) {
        $testsWithSpeed = $successfulTests | Where-Object { $_.DownloadSpeed_Kbps -ne $null }
        
        if ($testsWithSpeed.Count -gt 0) {
            $overallAvgSpeed = ($testsWithSpeed | Measure-Object DownloadSpeed_Kbps -Average).Average
            $minSpeed = ($testsWithSpeed | Measure-Object DownloadSpeed_Kbps -Minimum).Minimum
            $maxSpeed = ($testsWithSpeed | Measure-Object DownloadSpeed_Kbps -Maximum).Maximum
            
            Write-Host "🚀 Geschwindigkeits-/Antwortzeit-Statistiken:" -ForegroundColor Yellow
            Write-Host "   • Durchschnitt: $([math]::Round($overallAvgSpeed, 2)) ms/Kbps" -ForegroundColor White
            Write-Host "   • Minimum: $([math]::Round($minSpeed, 2)) ms/Kbps" -ForegroundColor Green
            Write-Host "   • Maximum: $([math]::Round($maxSpeed, 2)) ms/Kbps" -ForegroundColor Red
            Write-Host "   • Getestete Endpoints: $($testsWithSpeed.Count)" -ForegroundColor Cyan
        }
    }
}

# Generate HTML report if requested
if ($HtmlReport) {
    if (-not $HtmlReport.EndsWith('.html')) {
        $HtmlReport += '.html'
    }
    
    # Use default filename if no path specified
    if (-not (Split-Path $HtmlReport -Parent)) {
        $timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
        $HtmlReport = "Microsoft-Endpoints-Report-$timestamp.html"
    }
    
    if (-not $Quiet) {
        Write-Host "`n📄 Generiere HTML-Report..." -ForegroundColor Cyan
    }
    
    $htmlSuccess = New-HtmlReport -Results $results -SelectedServices $selectedBackendUrls -FailedServices $failedServices -ServiceImpacts $serviceImpacts -SkipPing $SkipPing -SkipSpeed $SkipSpeed -FilePath $HtmlReport
    
    if ($htmlSuccess) {
        $fullPath = (Resolve-Path $HtmlReport -ErrorAction SilentlyContinue).Path
        if (-not $fullPath) {
            $fullPath = (Get-Item $HtmlReport -ErrorAction SilentlyContinue).FullName
        }
        if (-not $fullPath) {
            $fullPath = $HtmlReport
        }
        
        if (-not $Quiet) {
            Write-Host "✅ HTML-Report erfolgreich erstellt: $fullPath" -ForegroundColor Green
        }
        
        # Open report in browser if requested
        if ($OpenReport) {
            try {
                Start-Process $fullPath
                if (-not $Quiet) {
                    Write-Host "🌐 HTML-Report wird im Browser geöffnet..." -ForegroundColor Cyan
                }
            }
            catch {
                Write-Warning "Konnte HTML-Report nicht im Browser öffnen: $_"
            }
        }
    } else {
        Write-Error "Fehler beim Erstellen des HTML-Reports"
    }
}

# Final summary with service impact overview
if ($results.Status -contains 'FAILED') {
    $totalFailed = ($results | Where-Object { $_.Status -eq 'FAILED' }).Count
    $totalTested = $results.Count
    
    if (-not $Quiet) {
        Write-Host "`n🚨 WARNUNG: $totalFailed von $totalTested Microsoft-Endpunkten sind NICHT erreichbar!" -ForegroundColor Red
        Write-Host "="*80 -ForegroundColor Red
        
        if ($failedServices.Count -gt 0) {
            Write-Host "📋 BETROFFENE SERVICES UND MÖGLICHE AUSWIRKUNGEN:" -ForegroundColor Yellow
            Write-Host "-"*50 -ForegroundColor Yellow
            foreach ($service in $failedServices) {
                Write-Host "🔴 $service" -ForegroundColor Red
                if ($serviceImpacts.ContainsKey($service)) {
                    Write-Host "   └─ $($serviceImpacts[$service].Impact)" -ForegroundColor Yellow
                }
            }
            Write-Host "`n💡 EMPFEHLUNG: Überprüfen Sie Ihre Firewall- und Proxy-Konfiguration für die oben genannten Services." -ForegroundColor Cyan
        }
    }
    exit 1
} else {
    if (-not $Quiet) {
        Write-Host "`n✅ ERFOLG: Alle $($results.Count) Microsoft-Endpunkte sind erreichbar!" -ForegroundColor Green
        Write-Host "🎉 Alle Microsoft Cloud Services sollten ordnungsgemäß funktionieren." -ForegroundColor Green
    }
    exit 0
}
