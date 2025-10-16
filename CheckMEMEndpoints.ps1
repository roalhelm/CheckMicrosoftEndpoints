[CmdletBinding()]
param(
    [Parameter(HelpMessage="Skip ping latency tests to speed up execution")]
    [switch]$SkipPing,
    
    [Parameter(HelpMessage="Skip download speed tests to speed up execution")]
    [switch]$SkipSpeed,
    
    [Parameter(HelpMessage="Run in quiet mode with minimal output")]
    [switch]$Quiet,
    
    [Parameter(HelpMessage="Generate detailed log file")]
    [string]$LogFile = "MEM-Endpoints-Test-$(Get-Date -Format 'yyyyMMdd-HHmmss').log",
    
    [Parameter(HelpMessage="Test timeout in seconds")]
    [int]$TimeoutSeconds = 10
)

<#
.SYNOPSIS
    Microsoft Endpoint Manager (MEM) Connectivity and Performance Tester

.DESCRIPTION
    This script dynamically retrieves the current Microsoft Endpoint Manager (MEM) IP addresses and URLs 
    from the official Microsoft API and tests their reachability through your firewall. It provides detailed 
    connectivity, latency, and speed measurements with comprehensive logging for comparison purposes.
    
    The script performs the following operations:
    1. Retrieves current MEM IPs and URLs from Microsoft's official API
    2. Tests TCP connectivity to all endpoints on port 443 (HTTPS)
    3. Measures ping latency for performance analysis (optional)
    4. Tests download speed/response time to evaluate network quality (optional)
    5. Generates detailed log files for comparison and troubleshooting
    6. Provides color-coded output for easy identification of connectivity issues
    7. Returns appropriate exit codes for automation scenarios
    
    The log file contains:
    - Test execution timestamp and duration
    - Complete list of tested IPs and URLs with results
    - Performance statistics (latency, speed)
    - Connectivity summary for easy comparison
    - Detailed error information for failed connections

.PARAMETER SkipPing
    Skips ping latency testing to reduce execution time. Only basic connectivity will be tested.

.PARAMETER SkipSpeed
    Skips download speed/response time testing to reduce execution time.

.PARAMETER Quiet
    Runs in quiet mode with minimal console output. Results are still logged.

.PARAMETER LogFile
    Specifies the path for the log file. If not provided, creates a timestamped log file in current directory.

.PARAMETER TimeoutSeconds
    Timeout in seconds for network tests. Default is 10 seconds.

.NOTES
    File Name     : CheckMEMEndpoints.ps1
    Author        : Created based on requirements
    Version       : 1.0
    Creation Date : October 16, 2025
    Requirements  : PowerShell 5.1 or higher, Internet connectivity
    Network       : HTTPS (443) access to Microsoft endpoints, ICMP for ping tests
    Permissions   : No elevated privileges required

.EXAMPLE
    .\CheckMEMEndpoints.ps1
    Runs complete connectivity test with ping and speed tests, creates timestamped log file.

.EXAMPLE
    .\CheckMEMEndpoints.ps1 -SkipPing -SkipSpeed -Quiet
    Runs basic connectivity test only in quiet mode with logging.

.EXAMPLE
    .\CheckMEMEndpoints.ps1 -LogFile "C:\Logs\MEM-Test.log"
    Runs complete test and saves results to specified log file.
#>

# Function to write to log file with timestamp
function Write-Log {
    param(
        [string]$Message,
        [string]$Level = "INFO"
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"
    
    # Write to log file
    Add-Content -Path $script:LogFile -Value $logEntry -Encoding UTF8
    
    # Also write to console if not quiet
    if (-not $Quiet) {
        switch ($Level) {
            "ERROR" { Write-Host $logEntry -ForegroundColor Red }
            "WARN"  { Write-Host $logEntry -ForegroundColor Yellow }
            "SUCCESS" { Write-Host $logEntry -ForegroundColor Green }
            "INFO"  { Write-Host $logEntry -ForegroundColor Cyan }
            default { Write-Host $logEntry }
        }
    }
}

# Function to test ping latency
function Test-PingLatency {
    param (
        [string]$HostName,
        [int]$Count = 4,
        [int]$TimeoutSeconds = 5
    )
    
    try {
        $pingResults = Test-Connection -ComputerName $HostName -Count $Count -TimeoutSeconds $TimeoutSeconds -ErrorAction Stop
        $avgLatency = ($pingResults | Measure-Object ResponseTime -Average).Average
        return [math]::Round($avgLatency, 2)
    }
    catch {
        return $null
    }
}

# Function to test download speed/response time
function Test-DownloadSpeed {
    param (
        [string]$Url,
        [int]$TimeoutSeconds = 10
    )
    
    try {
        $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
        
        # Try different approaches for speed testing
        $webClient = New-Object System.Net.WebClient
        $webClient.Headers.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36")
        
        try {
            # First try favicon.ico
            $testUrl = $Url.TrimEnd('/') + "/favicon.ico"
            $data = $webClient.DownloadData($testUrl)
            $stopwatch.Stop()
            
            if ($stopwatch.ElapsedMilliseconds -gt 0 -and $data.Length -gt 0) {
                $speedBps = ($data.Length * 8) / ($stopwatch.ElapsedMilliseconds / 1000)
                $speedKbps = [math]::Round($speedBps / 1024, 2)
                return $speedKbps
            }
        }
        catch {
            # If favicon doesn't work, try HEAD request for response time
            try {
                $request = [System.Net.WebRequest]::Create($Url)
                $request.Method = "HEAD"
                $request.Timeout = $TimeoutSeconds * 1000
                
                $stopwatch.Restart()
                $response = $request.GetResponse()
                $stopwatch.Stop()
                $response.Close()
                
                return [math]::Round($stopwatch.ElapsedMilliseconds, 2)
            }
            catch {
                # Last resort: try GET request with timeout
                $request = [System.Net.WebRequest]::Create($Url)
                $request.Method = "GET" 
                $request.Timeout = $TimeoutSeconds * 1000
                
                $stopwatch.Restart()
                $response = $request.GetResponse()
                $stopwatch.Stop()
                $response.Close()
                
                return [math]::Round($stopwatch.ElapsedMilliseconds, 2)
            }
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

# Function to get MEM IPs from Microsoft API
function Get-MEMIPs {
    Write-Log "Retrieving MEM IP addresses from Microsoft API..." "INFO"
    
    try {
        $uri = "https://endpoints.office.com/endpoints/WorldWide?ServiceAreas=MEM&clientrequestid=" + ([GUID]::NewGuid()).Guid
        $response = Invoke-RestMethod -Uri $uri -ErrorAction Stop
        $ips = $response | Where-Object {$_.ServiceArea -eq "MEM" -and $_.ips} | Select-Object -Unique -ExpandProperty ips
        
        Write-Log "Retrieved $($ips.Count) unique IP addresses/ranges from Microsoft API" "SUCCESS"
        return $ips
    }
    catch {
        Write-Log "Error retrieving MEM IPs: $($_.Exception.Message)" "ERROR"
        return @()
    }
}

# Function to get MEM URLs from Microsoft API
function Get-MEMURLs {
    Write-Log "Retrieving MEM URLs from Microsoft API..." "INFO"
    
    try {
        $uri = "https://endpoints.office.com/endpoints/WorldWide?ServiceAreas=MEM&clientrequestid=" + ([GUID]::NewGuid()).Guid
        $response = Invoke-RestMethod -Uri $uri -ErrorAction Stop
        $urls = $response | Where-Object {$_.ServiceArea -eq "MEM" -and $_.urls} | Select-Object -Unique -ExpandProperty urls
        
        Write-Log "Retrieved $($urls.Count) unique URLs from Microsoft API" "SUCCESS"
        return $urls
    }
    catch {
        Write-Log "Error retrieving MEM URLs: $($_.Exception.Message)" "ERROR"
        return @()
    }
}

# Function to test IP connectivity (convert CIDR to testable IPs)
function Test-IPConnectivity {
    param(
        [string]$IPRange,
        [int]$Port = 443
    )
    
    # Handle CIDR notation - for testing purposes, we'll test the network address
    if ($IPRange -match '^(\d+\.\d+\.\d+\.\d+)/(\d+)$') {
        $ip = $matches[1]
        $prefixLength = [int]$matches[2]
        
        # For /32, test the exact IP. For others, test the network address
        if ($prefixLength -eq 32) {
            $testIP = $ip
        } else {
            $testIP = $ip
        }
    } elseif ($IPRange -match '^\d+\.\d+\.\d+\.\d+$') {
        # Single IP address
        $testIP = $IPRange
    } else {
        Write-Log "Unsupported IP format: $IPRange" "WARN"
        return $null
    }
    
    try {
        $result = Test-NetConnection -ComputerName $testIP -Port $Port -WarningAction SilentlyContinue
        return $result
    }
    catch {
        Write-Log "Error testing IP $testIP : $($_.Exception.Message)" "ERROR"
        return $null
    }
}

# Function to test URL connectivity  
function Test-URLConnectivity {
    param(
        [string]$Url,
        [int]$Port = 443
    )
    
    # Extract hostname from URL
    $hostname = ($Url -replace 'https?://', '').Split('/')[0]
    
    # Handle wildcard URLs
    if ($hostname.StartsWith('*.')) {
        $hostname = $hostname -replace '^\*\.', 'www.'
    }
    
    try {
        $result = Test-NetConnection -ComputerName $hostname -Port $Port -WarningAction SilentlyContinue
        return $result
    }
    catch {
        Write-Log "Error testing URL $hostname : $($_.Exception.Message)" "ERROR"  
        return $null
    }
}

# Main script execution
$script:startTime = Get-Date
$script:LogFile = $LogFile

# Initialize log file
Write-Log "=== Microsoft Endpoint Manager (MEM) Connectivity Test ===" "INFO"
Write-Log "Script Version: 1.0" "INFO"
Write-Log "Test Configuration:" "INFO"
Write-Log "  - Ping Tests: $(if($SkipPing){'Disabled'}else{'Enabled'})" "INFO"
Write-Log "  - Speed Tests: $(if($SkipSpeed){'Disabled'}else{'Enabled'})" "INFO" 
Write-Log "  - Timeout: $TimeoutSeconds seconds" "INFO"
Write-Log "  - Quiet Mode: $(if($Quiet){'Enabled'}else{'Disabled'})" "INFO"
Write-Log "" "INFO"

if (-not $Quiet) {
    Write-Host "`n🌐 MICROSOFT ENDPOINT MANAGER (MEM) CONNECTIVITY TESTER" -ForegroundColor Cyan
    Write-Host "="*70 -ForegroundColor DarkCyan
    Write-Host "Retrieving current MEM endpoints from Microsoft API..." -ForegroundColor Yellow
}

# Get MEM IPs and URLs from Microsoft API
$memIPs = Get-MEMIPs
$memURLs = Get-MEMURLs

if ($memIPs.Count -eq 0 -and $memURLs.Count -eq 0) {
    Write-Log "No MEM endpoints retrieved from Microsoft API. Exiting." "ERROR"
    exit 1
}

Write-Log "Starting connectivity tests..." "INFO"
Write-Log "IPs to test: $($memIPs.Count)" "INFO"
Write-Log "URLs to test: $($memURLs.Count)" "INFO"

$results = @()
$totalEndpoints = $memIPs.Count + $memURLs.Count
$currentEndpoint = 0

if (-not $Quiet) {
    Write-Host "`n🔍 Testing $totalEndpoints MEM endpoints (IPs: $($memIPs.Count), URLs: $($memURLs.Count))" -ForegroundColor Cyan
    Write-Host "This may take several minutes..." -ForegroundColor Yellow
    Write-Host ""
}

# Test IP addresses
foreach ($ip in $memIPs) {
    $currentEndpoint++
    Write-Progress -Activity "Testing MEM Endpoints" -Status "Testing IP $ip ($currentEndpoint of $totalEndpoints)" -PercentComplete (($currentEndpoint / $totalEndpoints) * 100)
    
    Write-Log "Testing IP: $ip" "INFO"
    
    $connectResult = Test-IPConnectivity -IPRange $ip
    
    if ($connectResult -and $connectResult.TcpTestSucceeded) {
        $status = 'OK'
        $ipAddress = $connectResult.RemoteAddress
        Write-Log "IP $ip - Connection successful (Resolved to: $ipAddress)" "SUCCESS"
        
        # Test ping if enabled
        $pingLatency = $null
        if (-not $SkipPing -and $ipAddress) {
            Write-Log "Testing ping to $ipAddress..." "INFO"
            $pingLatency = Test-PingLatency -HostName $ipAddress -TimeoutSeconds $TimeoutSeconds
            if ($pingLatency) {
                Write-Log "Ping to $ipAddress : $pingLatency ms" "SUCCESS"
            } else {
                Write-Log "Ping to $ipAddress failed" "WARN"
            }
        }
        
        # Test speed if enabled
        $downloadSpeed = $null
        if (-not $SkipSpeed) {
            Write-Log "Testing response time to https://$ipAddress..." "INFO"
            $downloadSpeed = Test-DownloadSpeed -Url "https://$ipAddress" -TimeoutSeconds $TimeoutSeconds
            if ($downloadSpeed) {
                Write-Log "Response time to $ipAddress : $downloadSpeed ms/Kbps" "SUCCESS"
            } else {
                Write-Log "Speed test to $ipAddress failed" "WARN"
            }
        }
        
    } else {
        $status = 'FAILED'
        $ipAddress = $null
        Write-Log "IP $ip - Connection FAILED" "ERROR"
    }
    
    $results += [PSCustomObject]@{
        Type = 'IP'
        Endpoint = $ip
        Status = $status
        ResolvedIP = $ipAddress
        PingLatency_ms = $pingLatency
        DownloadSpeed_Kbps = $downloadSpeed
    }
}

# Test URLs
foreach ($url in $memURLs) {
    $currentEndpoint++
    $hostname = ($url -replace 'https?://', '').Split('/')[0]
    
    # Handle wildcards for display
    $displayUrl = $url
    if ($url.StartsWith('*.')) {
        $displayUrl = $url -replace '^\*\.', 'www.'
    }
    
    Write-Progress -Activity "Testing MEM Endpoints" -Status "Testing URL $displayUrl ($currentEndpoint of $totalEndpoints)" -PercentComplete (($currentEndpoint / $totalEndpoints) * 100)
    
    Write-Log "Testing URL: $url" "INFO"
    
    $connectResult = Test-URLConnectivity -Url $url
    
    if ($connectResult -and $connectResult.TcpTestSucceeded) {
        $status = 'OK' 
        $ipAddress = $connectResult.RemoteAddress
        Write-Log "URL $url - Connection successful (Resolved to: $ipAddress)" "SUCCESS"
        
        # Test ping if enabled
        $pingLatency = $null
        if (-not $SkipPing -and $ipAddress) {
            Write-Log "Testing ping to $hostname..." "INFO"
            $pingLatency = Test-PingLatency -HostName $hostname -TimeoutSeconds $TimeoutSeconds
            if ($pingLatency) {
                Write-Log "Ping to $hostname : $pingLatency ms" "SUCCESS"
            } else {
                Write-Log "Ping to $hostname failed" "WARN"
            }
        }
        
        # Test speed if enabled
        $downloadSpeed = $null
        if (-not $SkipSpeed) {
            Write-Log "Testing response time to $displayUrl..." "INFO"
            $downloadSpeed = Test-DownloadSpeed -Url $displayUrl -TimeoutSeconds $TimeoutSeconds
            if ($downloadSpeed) {
                Write-Log "Response time to $displayUrl : $downloadSpeed ms/Kbps" "SUCCESS"
            } else {
                Write-Log "Speed test to $displayUrl failed" "WARN"
            }
        }
        
    } else {
        $status = 'FAILED'
        $ipAddress = $null
        Write-Log "URL $url - Connection FAILED" "ERROR"
    }
    
    $results += [PSCustomObject]@{
        Type = 'URL'
        Endpoint = $url
        Status = $status
        ResolvedIP = $ipAddress
        PingLatency_ms = $pingLatency
        DownloadSpeed_Kbps = $downloadSpeed
    }
}

Write-Progress -Activity "Testing MEM Endpoints" -Completed

# Generate results summary
$endTime = Get-Date
$testDuration = $endTime - $script:startTime

$successfulTests = ($results | Where-Object { $_.Status -eq 'OK' }).Count
$failedTests = ($results | Where-Object { $_.Status -eq 'FAILED' }).Count
$successRate = if ($totalEndpoints -gt 0) { [math]::Round(($successfulTests / $totalEndpoints) * 100, 1) } else { 0 }

Write-Log "" "INFO"
Write-Log "=== TEST RESULTS SUMMARY ===" "INFO"
Write-Log "Test Duration: $($testDuration.ToString('hh\:mm\:ss'))" "INFO"
Write-Log "Total Endpoints Tested: $totalEndpoints" "INFO"
Write-Log "Successful Connections: $successfulTests" "SUCCESS"
Write-Log "Failed Connections: $failedTests" $(if($failedTests -eq 0){"SUCCESS"}else{"ERROR"})
Write-Log "Success Rate: $successRate%" $(if($successRate -eq 100){"SUCCESS"}else{"WARN"})
Write-Log "" "INFO"

# Log detailed results table
Write-Log "=== DETAILED RESULTS ===" "INFO"
$tableHeader = "Type".PadRight(5) + "Endpoint".PadRight(45) + "Status".PadRight(8) + "Resolved IP".PadRight(16)

if (-not $SkipPing) {
    $tableHeader += "Ping(ms)".PadRight(10)
}
if (-not $SkipSpeed) {
    $tableHeader += "Speed(Kbps/ms)".PadRight(15)
}

Write-Log $tableHeader "INFO"
Write-Log ("-" * $tableHeader.Length) "INFO"

foreach ($result in $results) {
    $ipValue = if($result.ResolvedIP){$result.ResolvedIP.ToString()}else{"-"}
    
    $line = $result.Type.PadRight(5) + 
            $result.Endpoint.PadRight(45) + 
            $result.Status.PadRight(8) + 
            $ipValue.PadRight(16)
            
    if (-not $SkipPing) {
        $line += $(if($result.PingLatency_ms){$result.PingLatency_ms}else{"-"}).ToString().PadRight(10)
    }
    if (-not $SkipSpeed) {
        $line += $(if($result.DownloadSpeed_Kbps){$result.DownloadSpeed_Kbps}else{"-"}).ToString().PadRight(15)
    }
    
    $logLevel = if ($result.Status -eq 'OK') { "SUCCESS" } else { "ERROR" }
    Write-Log $line $logLevel
}

# Performance statistics
if (-not $SkipPing -or -not $SkipSpeed) {
    Write-Log "" "INFO"
    Write-Log "=== PERFORMANCE STATISTICS ===" "INFO"
    
    $successfulResults = $results | Where-Object { $_.Status -eq 'OK' }
    
    if (-not $SkipPing) {
        $pingResults = $successfulResults | Where-Object { $null -ne $_.PingLatency_ms }
        if ($pingResults.Count -gt 0) {
            $avgPing = [math]::Round(($pingResults | Measure-Object PingLatency_ms -Average).Average, 2)
            $minPing = [math]::Round(($pingResults | Measure-Object PingLatency_ms -Minimum).Minimum, 2)  
            $maxPing = [math]::Round(($pingResults | Measure-Object PingLatency_ms -Maximum).Maximum, 2)
            
            Write-Log "Ping Statistics:" "INFO"
            Write-Log "  Average Latency: $avgPing ms" "INFO"
            Write-Log "  Minimum Latency: $minPing ms" "INFO"
            Write-Log "  Maximum Latency: $maxPing ms" "INFO"
            Write-Log "  Tested Hosts: $($pingResults.Count)" "INFO"
        }
    }
    
    if (-not $SkipSpeed) {
        $speedResults = $successfulResults | Where-Object { $null -ne $_.DownloadSpeed_Kbps }
        if ($speedResults.Count -gt 0) {
            $avgSpeed = [math]::Round(($speedResults | Measure-Object DownloadSpeed_Kbps -Average).Average, 2)
            $minSpeed = [math]::Round(($speedResults | Measure-Object DownloadSpeed_Kbps -Minimum).Minimum, 2)
            $maxSpeed = [math]::Round(($speedResults | Measure-Object DownloadSpeed_Kbps -Maximum).Maximum, 2)
            
            Write-Log "Speed/Response Time Statistics:" "INFO"
            Write-Log "  Average: $avgSpeed Kbps/ms" "INFO"
            Write-Log "  Minimum: $minSpeed Kbps/ms" "INFO"
            Write-Log "  Maximum: $maxSpeed Kbps/ms" "INFO"
            Write-Log "  Tested Endpoints: $($speedResults.Count)" "INFO"
        }
    }
}

# Failed endpoints analysis
if ($failedTests -gt 0) {
    Write-Log "" "INFO"
    Write-Log "=== FAILED ENDPOINTS ANALYSIS ===" "ERROR"
    
    $failedIPs = $results | Where-Object { $_.Type -eq 'IP' -and $_.Status -eq 'FAILED' }
    $failedURLs = $results | Where-Object { $_.Type -eq 'URL' -and $_.Status -eq 'FAILED' }
    
    if ($failedIPs.Count -gt 0) {
        Write-Log "Failed IP Addresses ($($failedIPs.Count)):" "ERROR"
        foreach ($failed in $failedIPs) {
            Write-Log "  - $($failed.Endpoint)" "ERROR"
        }
    }
    
    if ($failedURLs.Count -gt 0) {
        Write-Log "Failed URLs ($($failedURLs.Count)):" "ERROR"
        foreach ($failed in $failedURLs) {
            Write-Log "  - $($failed.Endpoint)" "ERROR"
        }
    }
    
    Write-Log "" "INFO"
    Write-Log "RECOMMENDATION: Check firewall rules for the failed endpoints above." "WARN"
    Write-Log "These endpoints are required for Microsoft Endpoint Manager functionality." "WARN"
}

Write-Log "" "INFO"
Write-Log "=== END OF TEST ===" "INFO"
Write-Log "Log file saved to: $script:LogFile" "SUCCESS"

# Console output summary if not quiet
if (-not $Quiet) {
    Write-Host "`n" + "="*70 -ForegroundColor White
    Write-Host "TEST RESULTS SUMMARY" -ForegroundColor White
    Write-Host "="*70 -ForegroundColor White
    
    Write-Host "Test Duration: $($testDuration.ToString('hh\:mm\:ss'))" -ForegroundColor Cyan
    Write-Host "Total Endpoints: $totalEndpoints" -ForegroundColor Cyan
    Write-Host "Successful: $successfulTests" -ForegroundColor $(if($successfulTests -eq $totalEndpoints){"Green"}else{"Yellow"})
    Write-Host "Failed: $failedTests" -ForegroundColor $(if($failedTests -eq 0){"Green"}else{"Red"})
    Write-Host "Success Rate: $successRate%" -ForegroundColor $(if($successRate -eq 100){"Green"}elseif($successRate -gt 80){"Yellow"}else{"Red"})
    
    # Display results table
    Write-Host "`nDetailed Results:" -ForegroundColor Cyan
    
    $tableColumns = @(
        @{Name="Type"; Expression={$_.Type}; Width=5},
        @{Name="Endpoint"; Expression={$_.Endpoint}; Width=45},
        @{Name="Status"; Expression={$_.Status}; Width=8},
        @{Name="IP Address"; Expression={if($_.ResolvedIP){$_.ResolvedIP.ToString()}else{"-"}}; Width=15}
    )
    
    if (-not $SkipPing) {
        $tableColumns += @{Name="Ping (ms)"; Expression={if($_.PingLatency_ms){$_.PingLatency_ms}else{"-"}}; Width=10}
    }
    
    if (-not $SkipSpeed) {
        $tableColumns += @{Name="Speed (Kbps)"; Expression={if($_.DownloadSpeed_Kbps){$_.DownloadSpeed_Kbps}else{"-"}}; Width=12}
    }
    
    $results | Format-Table $tableColumns -AutoSize
    
    Write-Host "`n📄 Detailed log file created: $script:LogFile" -ForegroundColor Green
    
    if ($failedTests -eq 0) {
        Write-Host "✅ SUCCESS: All MEM endpoints are reachable!" -ForegroundColor Green
        Write-Host "🎉 Microsoft Endpoint Manager should function properly." -ForegroundColor Green
    } else {
        Write-Host "⚠️ WARNING: $failedTests MEM endpoints are not reachable!" -ForegroundColor Red
        Write-Host "💡 Check your firewall configuration for the failed endpoints." -ForegroundColor Yellow
    }
}

# Set exit code
if ($failedTests -eq 0) {
    exit 0
} else {
    exit 1
}