param([string]$RunDir)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# =========================
# Toolkit scaffolding
# =========================
function Out([string]$m) { Write-Host $m }

$script:Report = [ordered]@{
    toolId      = "printer-fallback-audit"
    toolVersion = "1.0.0"
    startedUtc  = [DateTime]::UtcNow.ToString("o")
    status      = "running"
    runDir      = $null
    reportPath  = $null
    inputs      = [ordered]@{ VenueID = $null; ApplyFix = $false }
    sql         = [ordered]@{ SQL_Server = $null; SQL_DSN = $null; LocalSqlOnlyVerified = $false }
    results     = [ordered]@{
        venue     = $null
        printers  = @()
        counts    = [ordered]@{ printers = 0; withFallback = 0; withoutFallback = 0; anomalies = 0 }
        anomalies = @()
        fix       = [ordered]@{ attempted = $false; applied = $false; reason = $null; beforeAfter = @() }
    }
    finishedUtc = $null
    error       = $null
}

function Write-Report {
    try {
        $script:Report.finishedUtc = [DateTime]::UtcNow.ToString("o")
        ($script:Report | ConvertTo-Json -Depth 40) | Out-File -LiteralPath $script:Report.reportPath -Encoding UTF8 -Force
        Out ""
        Out "Report written: $($script:Report.reportPath)"
    } catch {
        Out "WARN: Failed to write Report.json: $($_.Exception.Message)"
    }
}

function Fail([string]$msg, [int]$code = 1) {
    Out ""
    Out "ERROR: $msg"
    $script:Report.status = "failed"
    $script:Report.error  = $msg
    Write-Report
    exit $code
}

function Succeed {
    $script:Report.status = "success"
    Write-Report
    exit 0
}

try {
    if ([string]::IsNullOrWhiteSpace($RunDir)) { $RunDir = $env:BEPoz_TOOLKIT_RUNDIR }
    if ([string]::IsNullOrWhiteSpace($RunDir)) { Fail "RunDir missing. Provide param RunDir or set BEPoz_TOOLKIT_RUNDIR." 2 }

    if (-not (Test-Path -LiteralPath $RunDir)) { New-Item -Path $RunDir -ItemType Directory -Force | Out-Null }
    $script:Report.runDir = (Resolve-Path -LiteralPath $RunDir).Path
    $script:Report.reportPath = Join-Path $script:Report.runDir "Report.json"
} catch {
    Fail "Failed to initialise RunDir: $($_.Exception.Message)" 2
}

# =========================
# GUI helpers (no Read-Host)
# =========================
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

function Show-InputBox {
    param(
        [Parameter(Mandatory)] [string] $Title,
        [Parameter(Mandatory)] [string] $Prompt,
        [string] $DefaultValue = ""
    )

    $form = New-Object System.Windows.Forms.Form
    $form.Text = $Title
    $form.StartPosition = "CenterScreen"
    $form.Size = New-Object System.Drawing.Size(520, 180)
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    $form.TopMost = $true

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.AutoSize = $false
    $lbl.Text = $Prompt
    $lbl.Location = New-Object System.Drawing.Point(12, 12)
    $lbl.Size = New-Object System.Drawing.Size(480, 40)
    $form.Controls.Add($lbl)

    $txt = New-Object System.Windows.Forms.TextBox
    $txt.Text = $DefaultValue
    $txt.Location = New-Object System.Drawing.Point(12, 58)
    $txt.Size = New-Object System.Drawing.Size(480, 22)
    $form.Controls.Add($txt)

    $ok = New-Object System.Windows.Forms.Button
    $ok.Text = "OK"
    $ok.Location = New-Object System.Drawing.Point(316, 95)
    $ok.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton = $ok
    $form.Controls.Add($ok)

    $cancel = New-Object System.Windows.Forms.Button
    $cancel.Text = "Cancel"
    $cancel.Location = New-Object System.Drawing.Point(402, 95)
    $cancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.CancelButton = $cancel
    $form.Controls.Add($cancel)

    $result = $form.ShowDialog()
    if ($result -ne [System.Windows.Forms.DialogResult]::OK) { return $null }
    return $txt.Text
}

function Show-ConfirmFixDialog {
    param(
        [Parameter(Mandatory)] [string] $Title,
        [Parameter(Mandatory)] [string] $Message
    )

    $form = New-Object System.Windows.Forms.Form
    $form.Text = $Title
    $form.StartPosition = "CenterScreen"
    $form.Size = New-Object System.Drawing.Size(640, 250)
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    $form.TopMost = $true

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.AutoSize = $false
    $lbl.Text = $Message
    $lbl.Location = New-Object System.Drawing.Point(12, 12)
    $lbl.Size = New-Object System.Drawing.Size(600, 90)
    $form.Controls.Add($lbl)

    $chk = New-Object System.Windows.Forms.CheckBox
    $chk.Text = "I understand this will change SQL data"
    $chk.Location = New-Object System.Drawing.Point(12, 108)
    $chk.Size = New-Object System.Drawing.Size(600, 22)
    $form.Controls.Add($chk)

    $lblYes = New-Object System.Windows.Forms.Label
    $lblYes.Text = "Type YES to proceed:"
    $lblYes.Location = New-Object System.Drawing.Point(12, 136)
    $lblYes.Size = New-Object System.Drawing.Size(170, 22)
    $form.Controls.Add($lblYes)

    $txt = New-Object System.Windows.Forms.TextBox
    $txt.Location = New-Object System.Drawing.Point(170, 136)
    $txt.Size = New-Object System.Drawing.Size(160, 22)
    $form.Controls.Add($txt)

    $ok = New-Object System.Windows.Forms.Button
    $ok.Text = "Apply Fix"
    $ok.Location = New-Object System.Drawing.Point(420, 170)
    $ok.Enabled = $false
    $form.Controls.Add($ok)

    $cancel = New-Object System.Windows.Forms.Button
    $cancel.Text = "Cancel"
    $cancel.Location = New-Object System.Drawing.Point(520, 170)
    $cancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.CancelButton = $cancel
    $form.Controls.Add($cancel)

    $handler = { $ok.Enabled = ($chk.Checked -and ($txt.Text.Trim().ToUpperInvariant() -eq "YES")) }
    $chk.add_CheckedChanged($handler)
    $txt.add_TextChanged($handler)

    $ok.add_Click({
        $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $form.Close()
    })

    return ($form.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK)
}

# =========================
# DataTable normaliser (THIS is the fix)
# =========================
function Ensure-DataTable {
    param([object]$Value)

    if ($null -eq $Value) { throw "Expected DataTable but got: <null>" }

    if ($Value -is [System.Data.DataTable]) { return $Value }

    if ($Value -is [System.Data.DataSet]) {
        if ($Value.Tables.Count -gt 0) { return $Value.Tables[0] }
        throw "Expected DataTable but got DataSet with 0 tables."
    }

    if ($Value -is [System.Array] -and $Value.Count -gt 0) {
        if ($Value[0] -is [System.Data.DataTable]) { return $Value[0] }

        if ($Value[0] -is [System.Data.DataRow]) {
            $t = $Value[0].Table.Clone()
            foreach ($r in $Value) { [void]$t.ImportRow($r) }
            return $t
        }
    }

    throw "Expected DataTable but got: $($Value.GetType().FullName)"
}

function DataTable-ToObjects {
    param([Parameter(Mandatory)][object]$dt)

    # Normalise anything (DataSet/DataTable/Object[]/DataRow[]) into a DataTable
    $dt = Ensure-DataTable $dt

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

# =========================
# SQL config from HKCU + local-only enforcement
# =========================
function Get-SqlConfig {
    $key = "HKCU:\SOFTWARE\Backoffice"
    if (-not (Test-Path -LiteralPath $key)) { Fail "Registry key not found: $key" 4 }

    $bo = Get-ItemProperty -LiteralPath $key
    $server = ("" + $bo.SQL_Server).Trim()
    $db     = ("" + $bo.SQL_DSN).Trim()

    if ([string]::IsNullOrWhiteSpace($server)) { Fail "SQL_Server missing/blank in $key" 4 }
    if ([string]::IsNullOrWhiteSpace($db))     { Fail "SQL_DSN missing/blank in $key" 4 }

    $script:Report.sql.SQL_Server = $server
    $script:Report.sql.SQL_DSN    = $db

    Out "SQL config from registry:"
    Out "  SQL_Server = $server"
    Out "  SQL_DSN    = $db"
    Out ""

    $hostPart = $server
    if ($hostPart.Contains("\")) { $hostPart = $hostPart.Split("\")[0] }
    $hostPart = $hostPart.Trim()

    $cn = $env:COMPUTERNAME
    $localOk =
        ($hostPart -eq ".") -or
        ($hostPart -ieq "(local)") -or
        ($hostPart -ieq "localhost") -or
        ($hostPart -ieq $cn)

    if (-not $localOk) {
        Fail "LOCAL SQL only: SQL_Server host '$hostPart' does not match this computer '$cn'. Refusing to connect." 5
    }

    $script:Report.sql.LocalSqlOnlyVerified = $true
    return [PSCustomObject]@{ Server = $server; Database = $db }
}

function Open-SqlConnection([string]$Server, [string]$Database) {
    $cs = "Server=$Server;Database=$Database;Integrated Security=True;TrustServerCertificate=True;"
    $conn = New-Object System.Data.SqlClient.SqlConnection $cs
    $conn.Open()
    return $conn
}

function Invoke-SqlQuery {
    param(
        [Parameter(Mandatory)][System.Data.SqlClient.SqlConnection]$Conn,
        [Parameter(Mandatory)][string]$Sql,
        [hashtable]$Params = $null
    )

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
    $ds = New-Object System.Data.DataSet
    [void]$da.Fill($ds)

    # IMPORTANT: return only ONE DataTable, not the whole Tables collection
    if ($ds.Tables.Count -eq 0) { return (New-Object System.Data.DataTable) }
    return $ds.Tables[0]
}

function Invoke-SqlNonQuery {
    param(
        [Parameter(Mandatory)][System.Data.SqlClient.SqlConnection]$Conn,
        [Parameter(Mandatory)][string]$Sql,
        [hashtable]$Params = $null
    )

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

# =========================
# Audit queries
# =========================
function Get-Venue([System.Data.SqlClient.SqlConnection]$Conn, [int]$VenueID) {
    $dt = Invoke-SqlQuery -Conn $Conn -Sql "SELECT VenueID, Name FROM dbo.Venue WHERE VenueID = @VenueID;" -Params @{ VenueID = $VenueID }
    return (DataTable-ToObjects $dt | Select-Object -First 1)
}

function Get-VenuePrintingDevices([System.Data.SqlClient.SqlConnection]$Conn, [int]$VenueID) {
    # NOTE: printer identification is heuristic; we can tighten once you confirm real printer DeviceType/SubType values.
    $sql = @"
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
  AND ISNULL(d.DeviceType,0) NOT IN (4,6,28) -- exclude EFTPOS/drawers/KDS-ish based on your samples
  AND (
        d.PortName IN ('TCP-P','WIN')
     OR d.PortName LIKE 'COM%'
     OR d.PortName LIKE 'TCP%'
     OR d.PortName LIKE 'WIN%'
     OR d.SubType IN (70,84,85,3005,3032)
     OR d.Name LIKE '%Printer%'
  )
ORDER BY st.Name, ws.Name, d.Name;
"@
    $dt = Invoke-SqlQuery -Conn $Conn -Sql $sql -Params @{ VenueID = $VenueID }
    return (DataTable-ToObjects $dt)
}

function Find-Anomalies([object[]]$devices) {
    $anoms = New-Object System.Collections.Generic.List[object]

    foreach ($d in $devices) {
        $fid = $d.FallbackID
        if ($fid -ne $null -and (""+$fid).Trim() -ne "") {
            if ([string]::IsNullOrWhiteSpace("" + $d.FallbackDeviceName)) {
                $anoms.Add([ordered]@{
                    type = "MissingFallbackDevice"
                    deviceId = [int]$d.DeviceID
                    deviceName = $d.DeviceName
                    fallbackId = [int]$fid
                    message = "FallbackID set but fallback device could not be resolved (missing row?)."
                })
            }
        }
    }

    foreach ($d in $devices) {
        $fid = $d.FallbackID
        if ($fid -ne $null -and (""+$fid).Trim() -ne "") {
            if ([int]$d.DeviceID -eq [int]$fid) {
                $anoms.Add([ordered]@{
                    type = "SelfFallback"
                    deviceId = [int]$d.DeviceID
                    deviceName = $d.DeviceName
                    fallbackId = [int]$fid
                    message = "Device fallback references itself."
                })
            }
        }
    }

    return @($anoms)
}

function Apply-Fix_ClearInvalidFallbacks([System.Data.SqlClient.SqlConnection]$Conn, [object[]]$devices, [object[]]$anomalies) {
    $script:Report.results.fix.attempted = $true

    $fixable = $anomalies | Where-Object { $_.type -in @("SelfFallback","MissingFallbackDevice") }
    if (-not $fixable -or $fixable.Count -eq 0) {
        $script:Report.results.fix.reason = "No fixable anomalies found (v1 only clears invalid/self fallback IDs)."
        Out "Apply Fix: nothing to do."
        return
    }

    $msg = @"
Apply Fix will change SQL for fallback settings:

- Clears invalid/self FallbackID values (sets FallbackID = NULL)
- Does NOT assign new fallback printers

Tick the box and type YES to proceed.
"@

    $confirmed = Show-ConfirmFixDialog -Title "Printer Fallback Audit - Apply Fix" -Message $msg
    if (-not $confirmed) {
        $script:Report.results.fix.reason = "User cancelled."
        Out "Apply Fix cancelled."
        return
    }

    $script:Report.inputs.ApplyFix = $true

    foreach ($a in $fixable) {
        $deviceId = [int]$a.deviceId
        $deviceRow = $devices | Where-Object { [int]$_.DeviceID -eq $deviceId } | Select-Object -First 1

        $before = [ordered]@{
            deviceId = $deviceId
            deviceName = $deviceRow.DeviceName
            beforeFallbackId = if ($deviceRow.FallbackID -ne $null -and (""+$deviceRow.FallbackID).Trim() -ne "") { [int]$deviceRow.FallbackID } else { $null }
            beforeFallbackName = $deviceRow.FallbackDeviceName
            anomalyType = $a.type
        }

        $affected = Invoke-SqlNonQuery -Conn $Conn -Sql "UPDATE dbo.Device SET FallbackID = NULL WHERE DeviceID = @DeviceID;" -Params @{ DeviceID = $deviceId }

        $dtAfter = Invoke-SqlQuery -Conn $Conn -Sql "SELECT d.DeviceID, d.Name AS DeviceName, d.FallbackID, fb.Name AS FallbackDeviceName FROM dbo.Device d LEFT JOIN dbo.Device fb ON fb.DeviceID = d.FallbackID WHERE d.DeviceID = @DeviceID;" -Params @{ DeviceID = $deviceId }
        $afterObj = (DataTable-ToObjects $dtAfter | Select-Object -First 1)

        $after = [ordered]@{
            deviceId = [int]$afterObj.DeviceID
            deviceName = $afterObj.DeviceName
            afterFallbackId = if ($afterObj.FallbackID -ne $null -and (""+$afterObj.FallbackID).Trim() -ne "") { [int]$afterObj.FallbackID } else { $null }
            afterFallbackName = $afterObj.FallbackDeviceName
            rowsAffected = $affected
        }

        $script:Report.results.fix.beforeAfter += [ordered]@{ before = $before; after = $after }
        Out ("Cleared fallback for DeviceID {0} ({1}) [RowsAffected={2}]" -f $deviceId, $deviceRow.DeviceName, $affected)
    }

    $script:Report.results.fix.applied = $true
    $script:Report.results.fix.reason  = "Cleared invalid/self fallback IDs only."
}

# =========================
# Main
# =========================
Out "Printer Fallback Audit"
Out "RunDir: $($script:Report.runDir)"
Out ""

$conn = $null
try {
    $cfg = Get-SqlConfig

    Out "Opening SQL connection (local only)..."
    $conn = Open-SqlConnection -Server $cfg.Server -Database $cfg.Database
    Out "Connected."

    $rawVenue = Show-InputBox -Title "Printer Fallback Audit" -Prompt "Enter VenueID (integer) to audit:" -DefaultValue ""
    if ($null -eq $rawVenue) { Fail "User cancelled VenueID prompt." 7 }

    $rawVenue = $rawVenue.Trim()
    if ($rawVenue -notmatch '^\d+$') { Fail "VenueID must be an integer. You entered: '$rawVenue'." 7 }
    $venueId = [int]$rawVenue
    $script:Report.inputs.VenueID = $venueId

    $venue = Get-Venue -Conn $conn -VenueID $venueId
    if (-not $venue) { Fail "VenueID $venueId not found in dbo.Venue." 8 }
    $script:Report.results.venue = [ordered]@{ VenueID = [int]$venue.VenueID; Name = $venue.Name }

    Out ""
    Out "Auditing venue: [$venueId] $($venue.Name)"
    Out ""

    $devices = Get-VenuePrintingDevices -Conn $conn -VenueID $venueId

    $script:Report.results.counts.printers = ($devices | Measure-Object).Count
    $withFallback = ($devices | Where-Object { $_.FallbackID -ne $null -and (""+$_.FallbackID).Trim() -ne "" }).Count
    $script:Report.results.counts.withFallback = $withFallback
    $script:Report.results.counts.withoutFallback = $script:Report.results.counts.printers - $withFallback

    $anoms = Find-Anomalies -devices $devices
    $script:Report.results.anomalies = @($anoms)
    $script:Report.results.counts.anomalies = ($anoms | Measure-Object).Count

    $script:Report.results.printers = @(
        $devices | ForEach-Object {
            [ordered]@{
                StoreName = $_.StoreName
                WorkstationName = $_.WorkstationName
                DeviceID = [int]$_.DeviceID
                DeviceName = $_.DeviceName
                Disabled = [int]$_.Disabled
                DeviceType = $_.DeviceType
                SubType = $_.SubType
                PortName = $_.PortName
                IPAddress = $_.IPAddress
                TcpPort = $_.TcpPort
                FallbackID = if ($_.FallbackID -ne $null -and (""+$_.FallbackID).Trim() -ne "") { [int]$_.FallbackID } else { $null }
                FallbackDeviceName = $_.FallbackDeviceName
            }
        }
    )

    Out "Summary"
    Out "  Printers found:   $($script:Report.results.counts.printers)"
    Out "  With fallback:    $($script:Report.results.counts.withFallback)"
    Out "  Without fallback: $($script:Report.results.counts.withoutFallback)"
    Out "  Anomalies:        $($script:Report.results.counts.anomalies)"
    Out ""

    if ($devices.Count -gt 0) {
        Out "Store | Workstation | DeviceID | DeviceName | Port | Disabled | FallbackID | FallbackDeviceName"
        foreach ($d in $devices) {
            Out ("{0} | {1} | {2} | {3} | {4} | {5} | {6} | {7}" -f `
                ($d.StoreName ?? ""), ($d.WorkstationName ?? ""), $d.DeviceID, ($d.DeviceName ?? ""), ($d.PortName ?? ""), $d.Disabled, ($d.FallbackID ?? ""), ($d.FallbackDeviceName ?? ""))
        }
        Out ""
    } else {
        Out "NOTE: No devices matched the v1 printer heuristic for this venue."
        Out "      If incorrect, confirm printer DeviceType/SubType values and we will tighten the filter."
        Out ""
    }

    if ($anoms.Count -gt 0) {
        Out "Anomalies:"
        foreach ($a in $anoms) {
            Out ("- {0}: DeviceID {1} ({2}) -> {3}" -f $a.type, $a.deviceId, $a.deviceName, $a.message)
        }
        Out ""
    }

    # Optional fix (safe, limited)
    Apply-Fix_ClearInvalidFallbacks -Conn $conn -devices $devices -anomalies $anoms

    Succeed
}
catch {
    Fail $_.Exception.Message 1
}
finally {
    try { if ($conn) { $conn.Close(); $conn.Dispose() } } catch {}
    if ($script:Report.status -eq "running") {
        $script:Report.status = "failed"
        $script:Report.error = "Unexpected termination."
        Write-Report
        exit 1
    }
}
