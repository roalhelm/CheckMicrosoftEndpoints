[CmdletBinding()]
param(
    [Parameter(HelpMessage = "Path to endpoint script that contains endpoint service maps")]
    [string]$ScriptPath = "./CheckMicrosoftEndpointsV2.ps1",

    [Parameter(HelpMessage = "Path to JSON baseline file")]
    [string]$BaselinePath = "./.endpoint-baseline.json",

    [Parameter(HelpMessage = "Days between mandatory review cycles")]
    [int]$ReviewIntervalDays = 31,

    [Parameter(HelpMessage = "Update baseline file with current script values")]
    [switch]$UpdateBaseline,

    [Parameter(HelpMessage = "Return non-zero exit code when endpoint drift is detected")]
    [switch]$FailOnDrift,

    [Parameter(HelpMessage = "Reduce output noise")]
    [switch]$Quiet
)

<#!
.SYNOPSIS
    Monthly maintenance routine for endpoint baseline drift detection.

.DESCRIPTION
    Parses endpoint URL literals from the local endpoint script and compares them
    against a JSON baseline file. The routine warns when endpoints changed or when
    the baseline age exceeded the configured review interval.

.PARAMETER ScriptPath
    Endpoint script to parse. Default: ./CheckMicrosoftEndpointsV2.ps1

.PARAMETER BaselinePath
    Baseline JSON path. Default: ./.endpoint-baseline.json

.PARAMETER ReviewIntervalDays
    Threshold for age warning. Default: 31 days.

.PARAMETER UpdateBaseline
    Persist current endpoint set to the baseline file.

.PARAMETER FailOnDrift
    Exit code 2 when drift is detected.

.PARAMETER Quiet
    Reduce informational output.

.EXAMPLE
    .\MaintainEndpointBaseline.ps1

.EXAMPLE
    .\MaintainEndpointBaseline.ps1 -UpdateBaseline

.EXAMPLE
    .\MaintainEndpointBaseline.ps1 -FailOnDrift
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Write-Info {
    param([string]$Message)
    if (-not $Quiet) {
        Write-Host "[INFO] $Message" -ForegroundColor Cyan
    }
}

function Write-Warn {
    param([string]$Message)
    Write-Host "[WARN] $Message" -ForegroundColor Yellow
}

function Write-Ok {
    param([string]$Message)
    if (-not $Quiet) {
        Write-Host "[OK]   $Message" -ForegroundColor Green
    }
}

function Write-Err {
    param([string]$Message)
    Write-Host "[ERR]  $Message" -ForegroundColor Red
}

function Get-EndpointMapFromScript {
    param([string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        throw "Script file not found: $Path"
    }

    $lines = Get-Content -LiteralPath $Path -ErrorAction Stop

    $startIndex = -1
    for ($i = 0; $i -lt $lines.Count; $i++) {
        if ($lines[$i] -match '^\$backendUrls\s*=\s*@\{') {
            $startIndex = $i
            break
        }
    }

    if ($startIndex -lt 0) {
        throw "Could not find endpoint map ('$backendUrls = @{') in script: $Path"
    }

    $services = @{}
    $currentService = $null
    $depth = 0
    $inBlockComment = $false

    for ($i = $startIndex; $i -lt $lines.Count; $i++) {
        $line = $lines[$i]
        $trimmed = $line.Trim()

        if ($trimmed -match '<#') {
            $inBlockComment = $true
        }

        if (-not $inBlockComment) {
            $openCount = ([regex]::Matches($line, '\{')).Count
            $closeCount = ([regex]::Matches($line, '\}')).Count
            $depth += $openCount
            $depth -= $closeCount

            if ($trimmed -match "^'([^']+)'\s*=\s*@\(") {
                $currentService = $matches[1]
                if (-not $services.ContainsKey($currentService)) {
                    $services[$currentService] = New-Object System.Collections.Generic.List[string]
                }
                continue
            }

            if ($currentService -and $trimmed -match "^'([^']+)'\s*,?\s*$") {
                $services[$currentService].Add($matches[1])
                continue
            }

            if ($trimmed -match '^\)\s*\|\s*Sort-Object\s+-Unique') {
                $currentService = $null
                continue
            }

            if ($depth -le 0 -and $i -gt $startIndex) {
                break
            }
        }

        if ($trimmed -match '#>') {
            $inBlockComment = $false
        }
    }

    if ($services.Count -eq 0) {
        throw "No services parsed from endpoint map in script: $Path"
    }

    $normalized = @{}
    foreach ($serviceName in ($services.Keys | Sort-Object)) {
        $normalized[$serviceName] = @($services[$serviceName] | Sort-Object -Unique)
    }

    return $normalized
}

function Get-DriftReport {
    param(
        [hashtable]$Current,
        [hashtable]$Baseline
    )

    $report = [ordered]@{
        AddedServices = @()
        RemovedServices = @()
        EndpointAdds = @()
        EndpointRemovals = @()
        DriftDetected = $false
    }

    $currentServices = @($Current.Keys | Sort-Object)
    $baselineServices = @($Baseline.Keys | Sort-Object)

    $report.AddedServices = @($currentServices | Where-Object { $_ -notin $baselineServices })
    $report.RemovedServices = @($baselineServices | Where-Object { $_ -notin $currentServices })

    foreach ($service in @($currentServices | Where-Object { $_ -in $baselineServices })) {
        $currEndpoints = @($Current[$service])
        $baseEndpoints = @($Baseline[$service])

        $added = @($currEndpoints | Where-Object { $_ -notin $baseEndpoints })
        $removed = @($baseEndpoints | Where-Object { $_ -notin $currEndpoints })

        if ($added.Count -gt 0) {
            $report.EndpointAdds += [PSCustomObject]@{
                Service = $service
                Endpoints = $added
            }
        }

        if ($removed.Count -gt 0) {
            $report.EndpointRemovals += [PSCustomObject]@{
                Service = $service
                Endpoints = $removed
            }
        }
    }

    if (
        $report.AddedServices.Count -gt 0 -or
        $report.RemovedServices.Count -gt 0 -or
        $report.EndpointAdds.Count -gt 0 -or
        $report.EndpointRemovals.Count -gt 0
    ) {
        $report.DriftDetected = $true
    }

    return $report
}

function Save-Baseline {
    param(
        [hashtable]$ServiceMap,
        [string]$OutputPath,
        [string]$SourceScript
    )

    $obj = [ordered]@{
        GeneratedAt = (Get-Date).ToString('o')
        SourceScript = $SourceScript
        Services = $ServiceMap
    }

    $obj | ConvertTo-Json -Depth 10 | Out-File -LiteralPath $OutputPath -Encoding UTF8
}

try {
    Write-Info "Parsing endpoints from script: $ScriptPath"
    $currentMap = Get-EndpointMapFromScript -Path $ScriptPath

    $currentServiceCount = $currentMap.Count
    $currentEndpointCount = (@($currentMap.Values | ForEach-Object { $_ }).Count)
    Write-Ok "Parsed $currentServiceCount services and $currentEndpointCount endpoints"

    if (-not (Test-Path -LiteralPath $BaselinePath)) {
        Write-Warn "Baseline not found. Creating initial baseline at: $BaselinePath"
        Save-Baseline -ServiceMap $currentMap -OutputPath $BaselinePath -SourceScript $ScriptPath
        Write-Ok "Initial baseline created"
        exit 0
    }

    Write-Info "Loading baseline: $BaselinePath"
    $baselineRaw = Get-Content -LiteralPath $BaselinePath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop

    if (-not $baselineRaw.Services) {
        throw "Invalid baseline format: missing 'Services'"
    }

    $baselineMap = @{}
    foreach ($prop in $baselineRaw.Services.PSObject.Properties) {
        $baselineMap[$prop.Name] = @($prop.Value)
    }

    $baselineDate = $null
    if ($baselineRaw.GeneratedAt) {
        try {
            $baselineDate = [datetime]$baselineRaw.GeneratedAt
        }
        catch {
            Write-Warn "Could not parse baseline timestamp: $($baselineRaw.GeneratedAt)"
        }
    }

    $ageDays = $null
    $overdue = $false
    if ($baselineDate) {
        $ageDays = [math]::Floor(((Get-Date) - $baselineDate).TotalDays)
        if ($ageDays -ge $ReviewIntervalDays) {
            $overdue = $true
            Write-Warn "Baseline review overdue: $ageDays days (limit: $ReviewIntervalDays)"
        } else {
            Write-Ok "Baseline age: $ageDays days"
        }
    } else {
        Write-Warn "Baseline age unknown because timestamp is missing or invalid"
    }

    $drift = Get-DriftReport -Current $currentMap -Baseline $baselineMap

    if ($drift.DriftDetected) {
        Write-Warn "Endpoint drift detected between current script and baseline"

        if ($drift.AddedServices.Count -gt 0) {
            Write-Host "  Added services:" -ForegroundColor Yellow
            $drift.AddedServices | ForEach-Object { Write-Host "    + $_" -ForegroundColor Yellow }
        }

        if ($drift.RemovedServices.Count -gt 0) {
            Write-Host "  Removed services:" -ForegroundColor Yellow
            $drift.RemovedServices | ForEach-Object { Write-Host "    - $_" -ForegroundColor Yellow }
        }

        foreach ($entry in $drift.EndpointAdds) {
            Write-Host "  Added endpoints in [$($entry.Service)]:" -ForegroundColor Yellow
            foreach ($ep in $entry.Endpoints) {
                Write-Host "    + $ep" -ForegroundColor Yellow
            }
        }

        foreach ($entry in $drift.EndpointRemovals) {
            Write-Host "  Removed endpoints in [$($entry.Service)]:" -ForegroundColor Yellow
            foreach ($ep in $entry.Endpoints) {
                Write-Host "    - $ep" -ForegroundColor Yellow
            }
        }
    } else {
        Write-Ok "No endpoint drift detected"
    }

    if ($UpdateBaseline) {
        Write-Info "Updating baseline with current endpoint set"
        Save-Baseline -ServiceMap $currentMap -OutputPath $BaselinePath -SourceScript $ScriptPath
        Write-Ok "Baseline updated"
    }

    if ($drift.DriftDetected -and $FailOnDrift) {
        Write-Err "Failing because drift was detected and -FailOnDrift is enabled"
        exit 2
    }

    if ($overdue) {
        exit 1
    }

    exit 0
}
catch {
    Write-Err $_.Exception.Message
    exit 99
}
