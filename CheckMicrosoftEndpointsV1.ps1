
<#
.SYNOPSIS
    Comprehensive connectivity, latency, and performance test for all critical Microsoft cloud service endpoints including Autopatch, Intune, Defender, and Microsoft 365.

    GitHub Repository: https://github.com/roalhelm/PowershellScripts

.DESCRIPTION
    This extensive connectivity testing script validates network access to critical Microsoft service endpoints
    required for modern Windows environments and Microsoft cloud services. It covers all major services including
    Windows Update for Business, Windows Autopatch, Microsoft Intune, Microsoft Defender, Azure AD, Microsoft 365,
    Microsoft Store, Windows Activation, Microsoft Edge, and Windows Telemetry services.
    
    The script performs the following operations:
    1. Tests TCP connectivity to essential Microsoft service URLs on port 443 (HTTPS)
    2. Measures ping latency to all reachable endpoints for performance analysis
    3. Tests download speed/response time to evaluate network quality
    4. Groups results by service backend for organized reporting
    5. Provides color-coded output for easy identification of connectivity issues
    6. Reports IP addresses of reachable endpoints for network troubleshooting
    7. Displays comprehensive performance statistics and network quality ratings
    8. Returns appropriate exit codes for automation and monitoring scenarios
    
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

.NOTES
    File Name     : CheckMicrosoftEndpointsV1.ps1
    Author        : Ronny Alhelm
    Version       : 1.1
    Creation Date : October 13, 2025
    Last Modified : October 14, 2025
    Requirements  : PowerShell 5.1 or higher
    Network       : Internet connectivity required
    Ports         : HTTPS (443) access to Microsoft endpoints, ICMP for ping tests
    Permissions   : No elevated privileges required
    Duration      : 3-10 minutes depending on network conditions and endpoint count

.CHANGES
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
    1.1

.EXAMPLE
    .\CheckMicrosoftEndpointsV1.ps1
    Runs the complete endpoint connectivity, ping, and performance check and displays results grouped by service backend.
    
    Sample Output:
    Backend: Windows Update for Business
    URL                                    Status  IP-Adresse      Ping (ms)  Response (ms)
    ---                                    ------  ----------      ---------  -------------
    https://dl.delivery.mp.microsoft.com  OK      52.97.144.85    23.45      1250.75
    https://download.windowsupdate.com     OK      204.79.197.200  18.92      890.33
    
    📊 Durchschnittliche Latenz: 21.19 ms
    🚀 Durchschnittliche Antwortzeit: 542.10 ms
    ✅ Alle Endpunkte für Windows Update for Business sind erreichbar.
    🌟 Exzellente Netzwerk-Performance!

.EXAMPLE
    # Use in automated monitoring
    .\CheckMicrosoftEndpointsV1.ps1
    if ($LASTEXITCODE -eq 0) {
        Write-Host "All Microsoft endpoints accessible"
    } else {
        Write-Host "Connectivity issues detected" -ForegroundColor Red
    }

.EXAMPLE
    # Run with error capture for logging
    $output = .\CheckMicrosoftEndpointsV1.ps1 2>&1
    $exitCode = $LASTEXITCODE
    Add-Content -Path "C:\Logs\EndpointCheck.log" -Value "$(Get-Date): Exit Code $exitCode"

#>


# List of required Microsoft service endpoints (reviewed 2026)
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
        'https://mmdcustomer.microsoft.com',
        'https://mmdls.microsoft.com',
        'https://devicelistenerprod.microsoft.com',
        'https://login.windows.net',
        'https://device.autopatch.microsoft.com',
        'https://services.autopatch.microsoft.com',
        'https://payloadprod*.blob.core.windows.net',
        'https://*.webpubsub.azure.com'
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
        Write-Host "    ⚠️  AUSWIRKUNGEN:" -ForegroundColor Yellow
        Write-Host "       • $($serviceImpacts[$ServiceName].Impact)" -ForegroundColor Yellow
        Write-Host "    🔍 SYMPTOME:" -ForegroundColor Cyan
        Write-Host "       • $($serviceImpacts[$ServiceName].Symptoms)" -ForegroundColor Cyan
        Write-Host "    🚫 Betroffene Endpunkte: $($FailedEndpoints.Count)" -ForegroundColor Red
        Write-Host ""
    }
}

# Flatten all URLs for testing
$urls = $backendUrls.Values | ForEach-Object { $_ } 


# Replace wildcards with testable hostnames for direct test
$testUrls = $urls | ForEach-Object {
    if ($_ -like '*`**') {
        ($_ -replace '^https://\*\.', 'https://www.' -replace '\*', 'prod')
    } else {
        $_
    }
}

$results = @()

Write-Host "🔍 Starte erweiterte Endpunkt-Tests (Konnektivität, Ping, Geschwindigkeit)..." -ForegroundColor Cyan
Write-Host "Dies kann einige Minuten dauern..." -ForegroundColor Yellow
Write-Host ""

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
    
    # If basic connectivity works, test ping and speed
    if ($result.TcpTestSucceeded) {
        Write-Host "  ✅ $testHost - Teste Ping..." -ForegroundColor Green
        $pingLatency = Test-PingLatency -HostName $testHost
        
        Write-Host "  📊 $testHost - Teste Geschwindigkeit..." -ForegroundColor Blue
        $downloadSpeed = Test-DownloadSpeed -Url $url
    } else {
        Write-Host "  ❌ $testHost - Nicht erreichbar" -ForegroundColor Red
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
foreach ($backend in $backendUrls.Keys) {
    Write-Host "`nBackend: $backend" -ForegroundColor Cyan
    Write-Host "="*60 -ForegroundColor DarkCyan
    
    $backendResults = $results | Where-Object { $backendUrls[$backend] -contains $_.URL }
    
    # Display results with enhanced formatting
    $backendResults | Format-Table @(
        @{Name="URL"; Expression={$_.URL}; Width=40},
        @{Name="Status"; Expression={$_.Status}; Width=8},
        @{Name="IP-Adresse"; Expression={$_.IPAddress}; Width=15},
        @{Name="Ping (ms)"; Expression={if($_.PingLatency_ms) {"$($_.PingLatency_ms)"} else {"-"}}; Width=10},
        @{Name="Response (ms)"; Expression={if($_.DownloadSpeed_Kbps) {"$($_.DownloadSpeed_Kbps)"} else {"-"}}; Width=12}
    ) -AutoSize
    
    # Performance analysis
    $successfulEndpoints = $backendResults | Where-Object { $_.Status -eq 'OK' }
    if ($successfulEndpoints.Count -gt 0) {
        $avgPing = ($successfulEndpoints | Where-Object { $_.PingLatency_ms -ne $null } | Measure-Object PingLatency_ms -Average).Average
        $avgSpeed = ($successfulEndpoints | Where-Object { $_.DownloadSpeed_Kbps -ne $null } | Measure-Object DownloadSpeed_Kbps -Average).Average
        
        if ($avgPing) {
            $pingColor = if ($avgPing -lt 50) { "Green" } elseif ($avgPing -lt 100) { "Yellow" } else { "Red" }
            Write-Host "📊 Durchschnittliche Latenz: $([math]::Round($avgPing, 2)) ms" -ForegroundColor $pingColor
        }
        
        if ($avgSpeed) {
            $speedColor = if ($avgSpeed -gt 1000) { "Green" } elseif ($avgSpeed -gt 100) { "Yellow" } else { "Red" }
            Write-Host "🚀 Durchschnittliche Antwortzeit: $([math]::Round($avgSpeed, 2)) ms" -ForegroundColor $speedColor
        }
    }
    
    $failedEndpoints = $backendResults | Where-Object { $_.Status -eq 'FAILED' }
    if ($failedEndpoints.Count -gt 0) {
        Write-Host "❌ $($failedEndpoints.Count) Endpunkt(e) für $backend sind NICHT erreichbar!" -ForegroundColor Red
        $failedServices += $backend
        Show-ServiceImpactWarning -ServiceName $backend -FailedEndpoints $failedEndpoints
    } else {
        Write-Host "✅ Alle Endpunkte für $backend sind erreichbar." -ForegroundColor Green
        
        # Performance rating
        if ($avgPing) {
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
Write-Host "`n" + "="*100 -ForegroundColor White
Write-Host "GESAMTÜBERSICHT ALLER ENDPUNKT-TESTS MIT PERFORMANCE-METRIKEN" -ForegroundColor White
Write-Host "="*100 -ForegroundColor White

# Create enhanced summary table
$results | Format-Table @(
    @{Name="URL"; Expression={$_.URL}; Width=45},
    @{Name="Host"; Expression={$_.Host}; Width=30},
    @{Name="Status"; Expression={$_.Status}; Width=8},
    @{Name="IP-Adresse"; Expression={$_.IPAddress}; Width=15},
    @{Name="Ping (ms)"; Expression={if($_.PingLatency_ms) {"$($_.PingLatency_ms)"} else {"-"}}; Width=10},
    @{Name="Response (ms)"; Expression={if($_.DownloadSpeed_Kbps) {"$($_.DownloadSpeed_Kbps)"} else {"-"}}; Width=12}
) -AutoSize

# Overall performance statistics
Write-Host "`n📈 GESAMTE NETZWERK-PERFORMANCE STATISTIKEN:" -ForegroundColor Cyan
Write-Host "-"*60 -ForegroundColor DarkCyan

$successfulTests = $results | Where-Object { $_.Status -eq 'OK' }
$testsWithPing = $successfulTests | Where-Object { $_.PingLatency_ms -ne $null }
$testsWithSpeed = $successfulTests | Where-Object { $_.DownloadSpeed_Kbps -ne $null }

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

if ($testsWithSpeed.Count -gt 0) {
    $overallAvgSpeed = ($testsWithSpeed | Measure-Object DownloadSpeed_Kbps -Average).Average
    $minSpeed = ($testsWithSpeed | Measure-Object DownloadSpeed_Kbps -Minimum).Minimum
    $maxSpeed = ($testsWithSpeed | Measure-Object DownloadSpeed_Kbps -Maximum).Maximum
    
    Write-Host "🚀 Antwortzeit-Statistiken:" -ForegroundColor Yellow
    Write-Host "   • Durchschnitt: $([math]::Round($overallAvgSpeed, 2)) ms" -ForegroundColor White
    Write-Host "   • Minimum: $([math]::Round($minSpeed, 2)) ms" -ForegroundColor Green
    Write-Host "   • Maximum: $([math]::Round($maxSpeed, 2)) ms" -ForegroundColor Red
    Write-Host "   • Getestete Endpoints: $($testsWithSpeed.Count)" -ForegroundColor Cyan
}

# Final summary with service impact overview
if ($results.Status -contains 'FAILED') {
    $totalFailed = ($results | Where-Object { $_.Status -eq 'FAILED' }).Count
    $totalTested = $results.Count
    
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
    exit 1
} else {
    Write-Host "`n✅ ERFOLG: Alle $($results.Count) Microsoft-Endpunkte sind erreichbar!" -ForegroundColor Green
    Write-Host "🎉 Alle Microsoft Cloud Services sollten ordnungsgemäß funktionieren." -ForegroundColor Green
    exit 0
}
