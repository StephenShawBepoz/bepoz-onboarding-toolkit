param([string]$RunDir)

# =========================
# Bepoz Onboarding Toolkit Tool
# Tool: Printer Fallback Audit (and optional Fix)
# Requirements:
# - Built-in PowerShell/.NET only
# - LOCAL SQL only (no remote SQL)
# - Read SQL config from HKCU:\SOFTWARE\Backoffice (SQL_Server, SQL_DSN)
# - Write-Host only for output
# - Create <RunDir>\Report.json
# - Exit 0 success, non-zero failure
# - No changes unless user explicitly chooses FIX and types YES
# - If changing SQL, require user to type YES and write Before/After into Report.json
# =========================

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Write-Out([string]$msg) { Write-Host $msg }

function Fail([string]$msg, [int]$code = 1) {
    Write-Out ""
    Write-Out "ERROR: $msg"
    try {
        if ($script:Report -and $script:Report.status) {
            $script:Report.status = "failed"
            $script:Report.error = $msg
            Write-Report
        }
    } catch {}
    exit $code
}

function Ensure-RunDir {
    if ([string]::IsNullOrWhiteSpace($RunDir)) {
        $RunDir = $env:BEPoz_TOOLKIT_RUNDIR
    }
    if ([string]::IsNullOrWhiteSpace($RunDir)) {
        Fail "RunDir was not provided and BEPoz_TOOLKIT_RUNDIR is not set."
    }
    if (-not (Test-Path -LiteralPath $RunDir)) {
        New-Item -Path $RunDir -ItemType Directory -Force | Out-Null
    }
    $script:RunDirResolved = (Resolve-Path -LiteralPath $RunDir).Path
    $script:ReportPath = Join-Path $script:RunDirResolved "Report.json"
}

function Write-Report {
    $json = $script:Report | ConvertTo-Json -Depth 20
    $json | Out-File -LiteralPath $script:ReportPath -Encoding UTF8 -Force
    Write-Out ""
    Write-Out "Report written: $script:ReportPath"
}

function Assert-ToolkitLocation {
    # Everything runs from C:\Bepoz\OnboardingToolkit (best-effort validation)
    try {
        $root = "C:\Bepoz\OnboardingToolkit"
        $here = $PSScriptRoot
        if (-not $here) { return }
        if ($here -notlike "$root*") {
            Write-Out "WARN: Tool is not running from $root (current: $here). Continuing anyway."
            $script:Report.warnings += "Not running from C:\Bepoz\OnboardingToolkit (current: $here)."
        }
    } catch {}
}

function Assert-NotPOS {
    # Heuristic: BackOffice/Server typically has BackOffice.exe; POS typically has SmartPOS.exe.
    $programs = "C:\Bepoz\Programs"
    $backOfficeExe = Join-Path $programs "BackOffice.exe"
    $smartPosExe   = Join-Path $programs "SmartPOS.exe"

    $hasBO  = Test-Path -LiteralPath $backOfficeExe
    $hasPOS = Test-Path -LiteralPath $smartPosExe

    if (-not $hasBO -and $hasPOS) {
        Fail "This tool is intended for BackOffice PCs/Servers only. Detected SmartPOS.exe without BackOffice.exe."
    }

    # If neither exists, don't hard fail (some environments differ), but warn.
    if (-not $hasBO -and -not $hasPOS) {
        Write-Out "WARN: Could not confirm BackOffice vs POS (BackOffice.exe/SmartPOS.exe not found). Continuing."
        $script:Report.warnings += "Could not confirm BackOffice vs POS (BackOffice.exe/SmartPOS.exe not found)."
    }
}

function Get-SqlConfigFromRegistry {
    $key = "HKCU:\SOFTWARE\Backoffice"
    if (-not (Test-Path -LiteralPath $key)) {
        Fail "Registry key not found: $key"
    }

    $bo = Get-ItemProperty -LiteralPath $key
    $server = ("" + $bo.SQL_Server).Trim()
    $db     = ("" + $bo.SQL_DSN).Trim()

    if ([string]::IsNullOrWhiteSpace($server)) { Fail "SQL_Server is missing/blank in $key" }
    if ([string]::IsNullOrWhiteSpace($db))     { Fail "SQL_DSN is missing/blank in $key" }

    $script:Report.sql.server = $server
    $script:Report.sql.database = $db

    Write-Out "SQL config from registry:"
    Write-Out "  SQL_Server = $server"
    Write-Out "  SQL_DSN    = $db"

    # LOCAL SQL only: ensure server points to local machine
    # Acceptable forms:
    # - .\INSTANCE
    # - localhost\INSTANCE
    # - COMPUTERNAME\INSTANCE
    # - (local)\INSTANCE
    $serverHost = $server
    if ($serverHost.Contains("\")) { $serverHost = $serverHost.Split("\")[0] }
    $serverHost = $serverHost.Trim()

    $localOk = $false
    $cn = $env:COMPUTERNAME

    if ($serverHost -eq "." -or $serverHost -eq "(local)" -or $serverHost -eq "localhost") { $localOk = $true }
    elseif ($serverHost -ieq $cn) { $localOk = $true }
    else {
        # Sometimes SQL_Server is set to the machine name even if running local; if mismatch, treat as remote.
        $localOk = $false
    }

    if (-not $localOk) {
        Fail "LOCAL SQL only: SQL_Server host '$serverHost' does not match this computer '$cn'. Refusing to connect."
    }

    return [PSCustomObject]@{ Server = $server; Database = $db }
}

function Open-SqlConnection([string]$Server, [string]$Database) {
    $cs = "Server=$Server;Database=$Database;Integrated Security=True;TrustServerCertificate=True;"
    $conn = New-Object System.Data.SqlClient.SqlConnection $cs
    $conn.Open()
    return $conn
}

function Invoke-Query([System.Data.SqlClient.SqlConnection]$Conn, [string]$Sql, [hashtable]$Params = $null) {
    $cmd = $Conn.CreateCommand()
    $cmd.CommandText = $Sql
    $cmd.CommandTimeout = 30

    if ($Params) {
        foreach ($k in $Params.Keys) {
            $p = $cmd.Parameters.Add("@$k", [System.Data.SqlDbType]::Variant)
            $p.Value = $Params[$k]
        }
    }

    $da = New-Object System.Data.SqlClient.SqlDataAdapter $cmd
    $dt = New-Object System.Data.DataTable
    [void]$da.Fill($dt)
    return $dt
}

function Invoke-NonQuery([System.Data.SqlClient.SqlConnection]$Conn, [string]$Sql, [hashtable]$Params = $null) {
    $cmd = $Conn.CreateCommand()
    $cmd.CommandText = $Sql
    $cmd.CommandTimeout = 30

    if ($Params) {
        foreach ($k in $Params.Keys) {
            $p = $cmd.Parameters.Add("@$k", [System.Data.SqlDbType]::Variant)
            $p.Value = $Params[$k]
        }
    }

    return $cmd.ExecuteNonQuery()
}

function DataTable-ToObjects([System.Data.DataTable]$dt) {
    $list = @()
    foreach ($row in $dt.Rows) {
        $o = [PSCustomObject]@{}
        foreach ($col in $dt.Columns) {
            $o | Add-Member -NotePropertyName $col.ColumnName -NotePropertyValue $row[$col.ColumnName]
        }
        $list += $o
    }
    return $list
}

function Prompt-Int([string]$prompt, [int]$min = 1) {
    while ($true) {
        $raw = Read-Host $prompt
        if ($raw -match '^\d+$') {
            $v = [int]$raw
            if ($v -ge $min) { return $v }
        }
        Write-Out "Please enter a valid integer >= $min."
    }
}

function Prompt-Choice([string]$prompt, [string[]]$valid) {
    $validUpper = $valid | ForEach-Object { $_.ToUpperInvariant() }
    while ($true) {
        $raw = (Read-Host $prompt)
        if ($null -eq $raw) { $raw = "" }
        $u = $raw.ToUpperInvariant().Trim()
        if ($validUpper -contains $u) { return $u }
        Write-Out ("Valid options: " + ($valid -join ", "))
    }
}

function Print-Table([object[]]$rows) {
    if (-not $rows -or $rows.Count -eq 0) {
        Write-Out "(no rows)"
        return
    }
    $rows | Format-Table -AutoSize | Out-String | ForEach-Object { $_.TrimEnd() } | ForEach-Object { Write-Out $_ }
}

# =========================
# Main
# =========================
Ensure-RunDir

$script:Report = [ordered]@{
    toolId      = "printer-fallback-audit"
    toolVersion = "1.0.0"
    startedUtc  = [DateTime]::UtcNow.ToString("o")
    runDir      = $script:RunDirResolved
    status      = "running"
    warnings    = @()
    sql         = [ordered]@{ server = ""; database = "" }
    inputs      = [ordered]@{}
    summary     = [ordered]@{}
    results     = [ordered]@{
        venuesListed = 0
        selectedVenue = $null
        devicesFound = 0
        devices = @()
    }
    fix         = [ordered]@{
        requested = $false
        actions = @()
        beforeAfter = @()
    }
}

try {
    Write-Out "Printer Fallback Audit"
    Write-Out "RunDir: $script:RunDirResolved"
    Assert-ToolkitLocation
    Assert-NotPOS

    $cfg = Get-SqlConfigFromRegistry

    Write-Out ""
    Write-Out "Opening SQL connection (local only)..."
    $conn = Open-SqlConnection -Server $cfg.Server -Database $cfg.Database
    Write-Out "Connected."

    # List venues
    $dtVenues = Invoke-Query -Conn $conn -Sql "SELECT VenueID, Name FROM dbo.Venue ORDER BY Name;"
    $venues = DataTable-ToObjects $dtVenues
    $script:Report.results.venuesListed = $venues.Count

    Write-Out ""
    Write-Out "Venues:"
    Print-Table $venues

    if ($venues.Count -eq 0) {
        Fail "No venues found in dbo.Venue."
    }

    Write-Out ""
    $venueId = Prompt-Int "Enter VenueID to audit"
    $venue = $venues | Where-Object { [int]$_.VenueID -eq $venueId } | Select-Object -First 1
    if (-not $venue) {
        Fail "VenueID $venueId was not found."
    }

    $script:Report.inputs.VenueID = $venueId
    $script:Report.results.selectedVenue = [ordered]@{ VenueID = $venueId; Name = $venue.Name }

    Write-Out ""
    Write-Out "Auditing Venue: [$venueId] $($venue.Name)"

    # Fetch printing devices across the whole venue:
    # - Venue -> Store -> Workstation -> Device
    # - Exclude EFTPOS (4), CashDrawer (6), KDS (28)
    # - Identify printers by typical port/subtype/name heuristics
    $sqlDevices = @"
SELECT
    v.VenueID,
    v.Name AS VenueName,
    st.StoreID,
    st.Name AS StoreName,
    ws.WorkstationID,
    ws.Name AS WorkstationName,

    d.DeviceID,
    d.Name AS DeviceName,
    d.Disabled,
    d.DeviceType,
    d.SubType,
    d.PortName,
    d.IPAddress,
    d.TcpPort,
    d.FallbackID,
    fb.Name AS FallbackDeviceName
FROM dbo.Venue v
JOIN dbo.Store st ON st.VenueID = v.VenueID
JOIN dbo.Workstation ws ON ws.StoreID = st.StoreID
JOIN dbo.Device d ON d.WorkstationID = ws.WorkstationID
LEFT JOIN dbo.Device fb ON fb.DeviceID = d.FallbackID
WHERE v.VenueID = @VenueID
  AND ISNULL(d.DeviceType,0) NOT IN (4,6,28)
  AND (
        d.PortName IN ('COM1','COM2','COM3','COM4','TCP-P','WIN')
     OR d.PortName LIKE 'COM%'
     OR d.PortName LIKE 'TCP%'
     OR d.PortName LIKE 'WIN%'
     OR d.SubType IN (70,84,85,3005,3032)
     OR d.Name LIKE '%Printer%'
     OR d.Name LIKE '%Expo%'
     OR d.Name LIKE '%Grill%'
     OR d.Name LIKE '%Pass%'
  )
ORDER BY st.Name, ws.Name, d.Name;
"@

    $dtDevices = Invoke-Query -Conn $conn -Sql $sqlDevices -Params @{ VenueID = $venueId }
    $devices = DataTable-ToObjects $dtDevices

    $script:Report.results.devicesFound = $devices.Count

    Write-Out ""
    Write-Out "Printing devices found (venue-wide): $($devices.Count)"
    Print-Table ($devices | Select-Object StoreName, WorkstationName, DeviceID, DeviceName, PortName, DeviceType, SubType, Disabled, FallbackID, FallbackDeviceName)

    # Save devices in report (structured, not table string)
    $script:Report.results.devices = @(
        $devices | ForEach-Object {
            [ordered]@{
                StoreName = $_.StoreName
                WorkstationName = $_.WorkstationName
                DeviceID = [int]$_.DeviceID
                DeviceName = $_.DeviceName
                PortName = $_.PortName
                DeviceType = $_.DeviceType
                SubType = $_.SubType
                Disabled = $_.Disabled
                FallbackID = if ($_.FallbackID -ne $null -and $_.FallbackID -ne "") { [int]$_.FallbackID } else { $null }
                FallbackDeviceName = $_.FallbackDeviceName
            }
        }
    )

    # Summary
    $withFallback = ($devices | Where-Object { $_.FallbackID -ne $null -and $_.FallbackID -ne "" }).Count
    $noFallback   = $devices.Count - $withFallback
    $disabled     = ($devices | Where-Object { [int]$_.Disabled -eq 1 }).Count

    $script:Report.summary = [ordered]@{
        totalPrintingDevices = $devices.Count
        withFallback = $withFallback
        withoutFallback = $noFallback
        disabledDevices = $disabled
    }

    Write-Out ""
    Write-Out "Summary:"
    Write-Out "  Total:           $($devices.Count)"
    Write-Out "  With fallback:   $withFallback"
    Write-Out "  Without fallback:$noFallback"
    Write-Out "  Disabled:        $disabled"

    # Optional FIX mode
    Write-Out ""
    Write-Out "No changes will be made unless you explicitly choose FIX."
    $mode = Prompt-Choice "Type AUDIT to finish, or FIX to set/clear a fallback" @("AUDIT","FIX")
    if ($mode -eq "AUDIT") {
        $script:Report.status = "success"
        $script:Report.finishedUtc = [DateTime]::UtcNow.ToString("o")
        Write-Report
        exit 0
    }

    $script:Report.fix.requested = $true

    if ($devices.Count -eq 0) {
        Fail "FIX requested, but there are no printing devices to modify for this venue."
    }

    Write-Out ""
    $targetDeviceId = Prompt-Int "Enter DeviceID to modify fallback for"
    $target = $devices | Where-Object { [int]$_.DeviceID -eq $targetDeviceId } | Select-Object -First 1
    if (-not $target) {
        Fail "DeviceID $targetDeviceId was not found in the printing-device list for this venue."
    }

    Write-Out ""
    Write-Out "Selected device:"
    Write-Out "  [$($target.DeviceID)] $($target.DeviceName)  |  $($target.StoreName) > $($target.WorkstationName)"
    Write-Out "  Current FallbackID: $($target.FallbackID)  ($($target.FallbackDeviceName))"

    $action = Prompt-Choice "Type SET to set a fallback, or CLEAR to remove fallback" @("SET","CLEAR")

    $newFallbackId = $null
    if ($action -eq "SET") {
        Write-Out ""
        Write-Out "Choose a fallback device from the SAME venue list:"
        Print-Table ($devices | Select-Object DeviceID, DeviceName, StoreName, WorkstationName)

        $newFallbackId = Prompt-Int "Enter Fallback DeviceID"
        if ($newFallbackId -eq [int]$target.DeviceID) {
            Fail "Fallback cannot be the same DeviceID as the device itself."
        }
        $fb = $devices | Where-Object { [int]$_.DeviceID -eq $newFallbackId } | Select-Object -First 1
        if (-not $fb) {
            Fail "Fallback DeviceID $newFallbackId was not found in the venue printing-device list."
        }

        Write-Out ""
        Write-Out "Proposed change:"
        Write-Out "  Device:   [$($target.DeviceID)] $($target.DeviceName)"
        Write-Out "  Fallback: [$($fb.DeviceID)] $($fb.DeviceName)"

        $confirm = Read-Host "Type YES to continue (anything else cancels)"
        if ($confirm.Trim().ToUpperInvariant() -ne "YES") {
            Fail "User cancelled (did not type YES). No changes were made." 2
        }
    }
    elseif ($action -eq "CLEAR") {
        Write-Out ""
        Write-Out "Proposed change:"
        Write-Out "  Device:   [$($target.DeviceID)] $($target.DeviceName)"
        Write-Out "  Fallback: (clear / NULL)"

        $confirm = Read-Host "Type YES to continue (anything else cancels)"
        if ($confirm.Trim().ToUpperInvariant() -ne "YES") {
            Fail "User cancelled (did not type YES). No changes were made." 2
        }
    }

    # Before snapshot
    $before = @{
        DeviceID = [int]$target.DeviceID
        DeviceName = $target.DeviceName
        BeforeFallbackID = if ($target.FallbackID -ne $null -and $target.FallbackID -ne "") { [int]$target.FallbackID } else { $null }
        BeforeFallbackDeviceName = $target.FallbackDeviceName
    }

    # Apply change
    Write-Out ""
    Write-Out "Applying change..."
    $sqlUpdate = "UPDATE dbo.Device SET FallbackID = @FallbackID WHERE DeviceID = @DeviceID;"
    $affected = Invoke-NonQuery -Conn $conn -Sql $sqlUpdate -Params @{ DeviceID = [int]$target.DeviceID; FallbackID = $newFallbackId }
    Write-Out "Rows affected: $affected"

    # After snapshot
    $dtAfter = Invoke-Query -Conn $conn -Sql "SELECT d.DeviceID, d.Name AS DeviceName, d.FallbackID, fb.Name AS FallbackDeviceName FROM dbo.Device d LEFT JOIN dbo.Device fb ON fb.DeviceID = d.FallbackID WHERE d.DeviceID = @DeviceID;" -Params @{ DeviceID = [int]$target.DeviceID }
    $afterObj = (DataTable-ToObjects $dtAfter | Select-Object -First 1)

    $after = @{
        DeviceID = [int]$afterObj.DeviceID
        DeviceName = $afterObj.DeviceName
        AfterFallbackID = if ($afterObj.FallbackID -ne $null -and $afterObj.FallbackID -ne "") { [int]$afterObj.FallbackID } else { $null }
        AfterFallbackDeviceName = $afterObj.FallbackDeviceName
    }

    $script:Report.fix.actions += [ordered]@{
        action = $action
        deviceId = [int]$target.DeviceID
        newFallbackId = $newFallbackId
        rowsAffected = $affected
        whenUtc = [DateTime]::UtcNow.ToString("o")
    }

    $script:Report.fix.beforeAfter += [ordered]@{
        before = $before
        after = $after
    }

    Write-Out ""
    Write-Out "Done."
    Write-Out "Before: FallbackID=$($before.BeforeFallbackID) ($($before.BeforeFallbackDeviceName))"
    Write-Out "After : FallbackID=$($after.AfterFallbackID) ($($after.AfterFallbackDeviceName))"

    $script:Report.status = "success"
    $script:Report.finishedUtc = [DateTime]::UtcNow.ToString("o")
    Write-Report
    exit 0
}
catch {
    $msg = $_.Exception.Message
    Fail $msg 1
}
finally {
    try { if ($conn) { $conn.Close(); $conn.Dispose() } } catch {}
}
