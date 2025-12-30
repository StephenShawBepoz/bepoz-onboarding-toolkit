#requires -version 5.1
<#
Venue Check (Read-only)
- Uses Toolkit-provided SQL context:
  BEPOZ_SQL_SERVER, BEPOZ_SQL_DSN, BEPOZ_TOOLKIT_RUNDIR
  and/or <RunDir>\ToolkitContext.json

Outputs (in RunDir):
- Venues.csv
- VenueSummary.csv
- Report.json

Safe by default: no changes, no prompts, no secrets written.
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Write-Log {
    param(
        [Parameter(Mandatory=$true)][string]$Message,
        [ValidateSet('INFO','WARN','ERROR')][string]$Level = 'INFO'
    )
    $ts = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    Write-Host "[$ts][$Level] $Message"
}

function Get-RunDir {
    $rd = $env:BEPOZ_TOOLKIT_RUNDIR
    if ([string]::IsNullOrWhiteSpace($rd)) {
        # Fallback: current directory (still write outputs somewhere)
        $rd = (Get-Location).Path
        Write-Log "BEPOZ_TOOLKIT_RUNDIR not set; using current directory: $rd" 'WARN'
    }
    if (-not (Test-Path -LiteralPath $rd)) {
        New-Item -ItemType Directory -Path $rd -Force | Out-Null
    }
    return $rd
}

function Read-ToolkitContextJson {
    param([Parameter(Mandatory=$true)][string]$RunDir)

    $ctxPath = Join-Path $RunDir 'ToolkitContext.json'
    if (Test-Path -LiteralPath $ctxPath) {
        try {
            return (Get-Content -LiteralPath $ctxPath -Raw | ConvertFrom-Json)
        } catch {
            Write-Log "Failed to parse ToolkitContext.json: $($_.Exception.Message)" 'WARN'
            return $null
        }
    }
    return $null
}

function Redact-ConnectionString {
    param([string]$ConnString)

    if ([string]::IsNullOrWhiteSpace($ConnString)) { return $ConnString }
    $redacted = $ConnString

    # Redact common secret keys if present
    $redacted = [regex]::Replace($redacted, '(?i)(Password|Pwd)\s*=\s*[^;]*', '$1=REDACTED')
    $redacted = [regex]::Replace($redacted, '(?i)(User\s*ID|UID)\s*=\s*[^;]*', '$1=REDACTED')
    return $redacted
}

function Get-SqlConnectionString {
    param($ToolkitContext)

    # Prefer env var
    $dsn = $env:BEPOZ_SQL_DSN

    # If missing, try ToolkitContext.json
    if ([string]::IsNullOrWhiteSpace($dsn) -and $ToolkitContext -ne $null) {
        try {
            if ($ToolkitContext.BEPOZ_SQL_DSN) { $dsn = [string]$ToolkitContext.BEPOZ_SQL_DSN }
        } catch { }
    }

    if ([string]::IsNullOrWhiteSpace($dsn)) {
        throw "No SQL DSN/connection string found. Expected BEPOZ_SQL_DSN env var (or ToolkitContext.json)."
    }

    # Most deployments provide a real SQL connection string. We'll treat it as such.
    return $dsn
}

function Invoke-SqlQuery {
    param(
        [Parameter(Mandatory=$true)][string]$ConnectionString,
        [Parameter(Mandatory=$true)][string]$Query,
        [hashtable]$Parameters
    )

    $conn = New-Object System.Data.SqlClient.SqlConnection $ConnectionString
    $cmd  = $conn.CreateCommand()
    $cmd.CommandText = $Query
    $cmd.CommandTimeout = 30

    if ($Parameters) {
        foreach ($k in $Parameters.Keys) {
            $p = $cmd.Parameters.Add("@$k", [System.Data.SqlDbType]::VarChar, 4000)
            $p.Value = [string]$Parameters[$k]
        }
    }

    $dt = New-Object System.Data.DataTable
    try {
        $conn.Open()
        $rdr = $cmd.ExecuteReader()
        $dt.Load($rdr)
        $rdr.Close()
    } finally {
        $conn.Close()
        $conn.Dispose()
    }
    return $dt
}

function Convert-IPv4IntToDotted {
    param([Nullable[int]]$IntIp)

    if ($null -eq $IntIp) { return $null }
    try {
        # Interpret as unsigned 32-bit, big-endian dotted quad (best-effort)
        $u = [uint32]$IntIp
        $o1 = ($u -shr 24) -band 0xFF
        $o2 = ($u -shr 16) -band 0xFF
        $o3 = ($u -shr 8)  -band 0xFF
        $o4 = $u -band 0xFF
        return "$o1.$o2.$o3.$o4"
    } catch {
        return $null
    }
}

function Write-JsonFile {
    param(
        [Parameter(Mandatory=$true)][string]$Path,
        [Parameter(Mandatory=$true)]$Object
    )
    $json = $Object | ConvertTo-Json -Depth 8
    $json | Set-Content -LiteralPath $Path -Encoding UTF8
}

# ---------------- MAIN ----------------
$runDir = Get-RunDir
$toolkitCtx = Read-ToolkitContextJson -RunDir $runDir

$report = [ordered]@{
    toolId      = 'venue-check'
    toolVersion = '1.0.0'
    startedAt   = (Get-Date).ToString('o')
    runDir      = $runDir
    sqlServer   = $env:BEPOZ_SQL_SERVER
    sqlDsnRedacted = $null
    status      = 'Unknown'
    checks      = @()
    counts      = @{}
    outputs     = @{}
    venuesSample = @()
    errors      = @()
}

try {
    Write-Log "RunDir: $runDir"
    if ($toolkitCtx) { Write-Log "ToolkitContext.json found." }

    $connStr = Get-SqlConnectionString -ToolkitContext $toolkitCtx
    $report.sqlDsnRedacted = Redact-ConnectionString $connStr

    Write-Log "Connecting to SQL (DSN/connection string provided by Toolkit)..."
    # Quick connectivity check
    $null = Invoke-SqlQuery -ConnectionString $connStr -Query "SELECT TOP 1 1 AS Ok;"

    Write-Log "Querying dbo.Venue..."
    # Avoid sensitive columns like MailPassword; select only minimum required columns.
    $qVenues = @"
SELECT
    v.VenueID,
    v.Name,
    v.SiteCode,
    v.SmartControllerID,
    v.MS_DNSName,
    v.MSIPAddress
FROM dbo.Venue v
ORDER BY v.VenueID;
"@

    $dtVenues = Invoke-SqlQuery -ConnectionString $connStr -Query $qVenues

    $venues = @()
    foreach ($row in $dtVenues.Rows) {
        $msIpInt = $null
        try { $msIpInt = [int]$row.MSIPAddress } catch { $msIpInt = $null }

        $venues += [pscustomobject]@{
            VenueID          = [int]$row.VenueID
            Name             = [string]$row.Name
            SiteCode         = [string]$row.SiteCode
            SmartControllerID= $row.SmartControllerID
            MS_DNSName       = [string]$row.MS_DNSName
            MSIPAddress_Int  = $msIpInt
            MSIPAddress      = Convert-IPv4IntToDotted -IntIp $msIpInt
        }
    }

    $report.counts.venueRowCount = $venues.Count

    if ($venues.Count -eq 0) {
        $report.checks += [pscustomobject]@{ name='VenuesExist'; result='Fail'; detail='dbo.Venue returned 0 rows.' }
        throw "dbo.Venue returned 0 rows."
    } else {
        $report.checks += [pscustomobject]@{ name='VenuesExist'; result='Pass'; detail="Found $($venues.Count) venue row(s)." }
    }

    Write-Log "Querying stores per venue..."
    $qStores = @"
SELECT
    v.VenueID,
    COUNT(s.StoreID) AS StoreCount
FROM dbo.Venue v
LEFT JOIN dbo.Store s ON s.VenueID = v.VenueID
GROUP BY v.VenueID
ORDER BY v.VenueID;
"@
    $dtStores = Invoke-SqlQuery -ConnectionString $connStr -Query $qStores

    $storeMap = @{}
    foreach ($row in $dtStores.Rows) {
        $storeMap[[int]$row.VenueID] = [int]$row.StoreCount
    }

    # Basic sanity checks
    $blankNames = $venues | Where-Object { [string]::IsNullOrWhiteSpace($_.Name) }
    if ($blankNames.Count -gt 0) {
        $report.checks += [pscustomobject]@{ name='BlankVenueNames'; result='Warn'; detail="Found $($blankNames.Count) venue(s) with blank Name." }
    } else {
        $report.checks += [pscustomobject]@{ name='BlankVenueNames'; result='Pass'; detail='No blank venue names found.' }
    }

    $noStores = @()
    foreach ($v in $venues) {
        $sc = 0
        if ($storeMap.ContainsKey($v.VenueID)) { $sc = $storeMap[$v.VenueID] }
        if ($sc -eq 0) { $noStores += $v }
    }
    if ($noStores.Count -gt 0) {
        $report.checks += [pscustomobject]@{ name='VenuesWithNoStores'; result='Warn'; detail="Found $($noStores.Count) venue(s) with 0 stores." }
    } else {
        $report.checks += [pscustomobject]@{ name='VenuesWithNoStores'; result='Pass'; detail='All venues have at least 1 store.' }
    }

    $dupNames = $venues | Group-Object -Property Name | Where-Object { $_.Count -gt 1 -and -not [string]::IsNullOrWhiteSpace($_.Name) }
    if ($dupNames.Count -gt 0) {
        $names = ($dupNames | Select-Object -ExpandProperty Name) -join ', '
        $report.checks += [pscustomobject]@{ name='DuplicateVenueNames'; result='Warn'; detail="Duplicate venue Name(s): $names" }
    } else {
        $report.checks += [pscustomobject]@{ name='DuplicateVenueNames'; result='Pass'; detail='No duplicate venue names found.' }
    }

    # Export CSVs
    $venuesCsv = Join-Path $runDir 'Venues.csv'
    $venues | Export-Csv -LiteralPath $venuesCsv -NoTypeInformation -Encoding UTF8
    $report.outputs.venuesCsv = $venuesCsv

    $summary = foreach ($v in $venues) {
        $sc = 0
        if ($storeMap.ContainsKey($v.VenueID)) { $sc = $storeMap[$v.VenueID] }
        [pscustomobject]@{
            VenueID    = $v.VenueID
            Name       = $v.Name
            StoreCount = $sc
        }
    }
    $summaryCsv = Join-Path $runDir 'VenueSummary.csv'
    $summary | Export-Csv -LiteralPath $summaryCsv -NoTypeInformation -Encoding UTF8
    $report.outputs.venueSummaryCsv = $summaryCsv

    # Console output
    Write-Log "Venues found: $($venues.Count)"
    $venues | Select-Object -First 20 | Format-Table -AutoSize | Out-String | Write-Host

    # Sample in report (keep it small)
    $report.venuesSample = $venues | Select-Object -First 25

    # Decide status
    $hasFail = ($report.checks | Where-Object { $_.result -eq 'Fail' }).Count -gt 0
    $hasWarn = ($report.checks | Where-Object { $_.result -eq 'Warn' }).Count -gt 0

    if ($hasFail) { $report.status = 'Fail' }
    elseif ($hasWarn) { $report.status = 'Warn' }
    else { $report.status = 'Pass' }

    Write-Log "Completed with status: $($report.status)"
}
catch {
    $msg = $_.Exception.Message
    Write-Log $msg 'ERROR'
    $report.status = 'Fail'
    $report.errors += $msg
}
finally {
    $report.finishedAt = (Get-Date).ToString('o')
    $reportPath = Join-Path $runDir 'Report.json'
    try {
        Write-JsonFile -Path $reportPath -Object $report
        Write-Log "Wrote report: $reportPath"
    } catch {
        Write-Log "Failed to write Report.json: $($_.Exception.Message)" 'ERROR'
    }
}

if ($report.status -eq 'Fail') { exit 1 }
exit 0
