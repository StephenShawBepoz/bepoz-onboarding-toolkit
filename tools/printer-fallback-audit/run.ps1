param([string]$RunDir)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# -------------------------
# Helpers: Output + Report
# -------------------------
function Out([string]$msg) { Write-Host $msg }

$script:Report = [ordered]@{
    toolId      = "printer-fallback-audit"
    toolVersion = "1.0.0"
    startedUtc  = [DateTime]::UtcNow.ToString("o")
    status      = "running"
    runDir      = $null
    reportPath  = $null
    warnings    = @()
    sql         = [ordered]@{ SQL_Server = $null; SQL_DSN = $null; LocalSqlOnlyVerified = $false }
    inputs      = [ordered]@{ VenueID = $null; ApplyFix = $false }
    results     = [ordered]@{
        venue      = $null
        printers   = @()
        counts     = [ordered]@{ printers = 0; withFallback = 0; withoutFallback = 0; disabled = 0; anomalies = 0 }
        anomalies  = @()
        fix        = [ordered]@{ supported = $true; attempted = $false; applied = $false; reason = $null; beforeAfter = @() }
    }
    finishedUtc = $null
    error       = $null
}

function Write-Report {
    try {
        $script:Report.finishedUtc = [DateTime]::UtcNow.ToString("o")
        $json = $script:Report | ConvertTo-Json -Depth 30
        $json | Out-File -LiteralPath $script:Report.reportPath -Encoding UTF8 -Force
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

# -------------------------
# RunDir rules
# -------------------------
try {
    if ([string]::IsNullOrWhiteSpace($RunDir)) { $RunDir = $env:BEPoz_TOOLKIT_RUNDIR }
    if ([string]::IsNullOrWhiteSpace($RunDir)) { Fail "RunDir was not provided and BEPoz_TOOLKIT_RUNDIR is not set." 2 }

    if (-not (Test-Path -LiteralPath $RunDir)) {
        New-Item -Path $RunDir -ItemType Directory -Force | Out-Null
    }
    $script:Report.runDir = (Resolve-Path -LiteralPath $RunDir).Path
    $script:Report.reportPath = Join-Path $script:Report.runDir "Report.json"
} catch {
    Fail "Failed to initialise RunDir: $($_.Exception.Message)" 2
}

# -------------------------
# Must run from C:\Bepoz\OnboardingToolkit (warn only)
# -------------------------
try {
    $expectedRoot = "C:\Bepoz\OnboardingToolkit"
    if ($PSScriptRoot -and ($PSScriptRoot -notlike "$expectedRoot*")) {
        Out "WARN: Not running from $expectedRoot (current: $PSScriptRoot). Continuing."
        $script:Report.warnings += "Not running from C:\Bepoz\OnboardingToolkit (current: $PSScriptRoot)."
    }
} catch {}

# -------------------------
# Only BackOffice / Server (not POS) heuristic
# -------------------------
try {
    $programs = "C:\Bepoz\Programs"
    $boExe  = Join-Path $programs "BackOffice.exe"
    $posExe = Join-Path $programs "SmartPOS.exe"

    $hasBO  = Test-Path -LiteralPath $boExe
    $hasPOS = Test-Path -LiteralPath $posExe

    if (-not $hasBO -and $hasPOS) {
        Fail "This tool must run on BackOffice PCs / Servers only. Detected SmartPOS.exe without BackOffice.exe." 3
    }
    if (-not $hasBO -and -not $hasPOS) {
        Out "WARN: Could not confirm BackOffice vs POS (BackOffice.exe/SmartPOS.exe not found). Continuing."
        $script:Report.warnings += "Could not confirm BackOffice vs POS (BackOffice.exe/SmartPOS.exe not found)."
    }
} catch {
    Out "WARN: BackOffice/POS heuristic check failed: $($_.Exception.Message)"
    $script:Report.warnings += "BackOffice/POS heuristic check failed: $($_.Exception.Message)"
}

# -------------------------
# GUI prompts (NO Read-Host)
# -------------------------
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
    $form.Size = New-Object System.Drawing.Size(620, 240)
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    $form.TopMost = $true

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.AutoSize = $false
    $lbl.Text = $Message
    $lbl.Location = New-Object System.Drawing.Point(12, 12)
    $lbl.Size = New-Object System.Drawing.Size(580, 80)
    $form.Controls.Add($lbl)

    $chk = New-Object System.Windows.Forms.CheckBox
    $chk.Text = "I understand this will change SQL data for fallback settings"
    $chk.Location = New-Object System.Drawing.Point(12, 96)
    $chk.Size = New-Object System.Drawing.Size(580, 22)
    $form.Controls.Add($chk)

    $lblYes = New-Object System.Windows.Forms.Label
    $lblYes.AutoSize = $false
    $lblYes.Text = "Type YES to proceed:"
    $lblYes.Location = New-Object System.Drawing.Point(12, 126)
    $lblYes.Size = New-Object System.Drawing.Size(200, 22)
    $form.Controls.Add($lblYes)

    $txt = New-Object System.Windows.Forms.TextBox
    $txt.Location = New-Object System.Drawing.Point(170, 126)
    $txt.Size = New-Object System.Drawing.Size(150, 22)
    $form.Controls.Add($txt)

    $ok = New-Object System.Windows.Forms.Button
    $ok.Text = "Apply Fix"
    $ok.Location = New-Object System.Drawing.Point(402, 160)
    $ok.Enabled = $false
    $form.Controls.Add($ok)

    $cancel = New-Object System.Windows.Forms.Button
    $cancel.Text = "Cancel"
    $cancel.Location = New-Object System.Drawing.Point(502, 160)
    $cancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.CancelButton = $cancel
    $form.Controls.Add($cancel)

    $handler = {
        $ok.Enabled = ($chk.Checked -and ($txt.Text.Trim().ToUpperInvariant() -eq "YES"))
    }
    $chk.add_CheckedChanged($handler)
    $txt.add_TextChanged($handler)

    $ok.add_Click({
        $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $form.Close()
    })

    $result = $form.ShowDialog()
    return ($result -eq [System.Windows.Forms.DialogResult]::OK)
}

# -------------------------
# DataTable normaliser (fixes Object[] / DataSet.Tables / DataRow[] issues)
# -------------------------
function Ensure-DataTable {
    param([object]$Value)

    if ($null -eq $Value) {
        throw "Expected DataTable but got: <null>"
    }

    if ($Value -is [System.Data.DataTable]) { return $Value }

    # DataSet -> first table
    if ($Value -is [System.Data.DataSet]) {
        if ($Value.Tables.Count -gt 0) { return $Value.Tables[0] }
        throw "Expected DataTable but got DataSet with 0 tables."
    }

    # Object[] that might contain DataTables or DataRows
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

# -------------------------
# SQL config from HKCU:\SOFTWARE\Backoffice
# LOCAL SQL ONLY enforcement
# -------------------------
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

    Out "SQL config:"
    Out "  SQL_Server = $server"
    Out "  SQL_DSN    = $db"

    # Local-only: validate host portion
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

function Invoke-Query {
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

    # Always fill a DataSet, then return ONLY the first DataTable (prevents Object[] issues)
    $da = New-Object System.Data.SqlClient.SqlDataAdapter $cmd
    $ds = New-Object System.Data.DataSet
    [void]$da.Fill($ds)

    if ($ds.Tables.Count -eq 0) {
        return (New-Object System.Data.DataTable)
    }

    return $ds.Tables[0]
}

function Invoke-NonQuery {
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

function DT {
    param([Parameter(Mandatory)][object]$dt)

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

# -------------------------
# Printer fallback audit logic
# - Uses dbo.Device.FallbackID
# - Venue -> Store -> Workstation -> Device
# -------------------------
function Get-VenueList([System.Data.SqlClient.SqlConnection]$Conn) {
    $dt = Invoke-Query -Conn $Conn -Sql "SELECT VenueID, Name FROM dbo.Venue ORDER BY Name;"
    return (DT $dt)
}

function Get-Venue([System.Data.SqlClient.SqlConnection]$Conn, [int]$VenueID) {
    $dt = Invoke-Query -Conn $Conn -Sql "SELECT VenueID, Name FROM dbo.Venue WHERE VenueID = @VenueID;" -Params @{ VenueID = $VenueID }
    return ((DT $dt) | Select-Object -First 1)
}

function Get-VenuePrintingDevices([System.Data.SqlClient.SqlConnection]$Conn, [int]$VenueID) {
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
    $dt = Invoke-Query -Conn $Conn -Sql $sql -Params @{ VenueID = $VenueID }
    return (DT $dt)
}

function Find-Anomalies([object[]]$devices) {
    $anoms = New-Object System.Collections.Generic.List[object]

    foreach ($d in $devices) {
        $fid = $d.FallbackID
        if ($fid -ne $null -and ("" + $fid).Trim() -ne "") {
            if ([string]::IsNullOrWhiteSpace("" + $d.FallbackDeviceName)) {
                $anoms.Add([ordered]@{
                    type = "MissingFallbackDevice"
                    deviceId = [int]$d.DeviceID
                    deviceName = $d.DeviceName
                    fallbackId = [int]$fid
                    message = "FallbackID is set but fallback device name could not be resolved."
                })
            }
        }
    }

    foreach ($d in $devices) {
        $fid = $d.FallbackID
        if ($fid -ne $null -and ("" + $fid).Trim() -ne "") {
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

    foreach ($d in $devices) {
        if ([int]$d.Disabled -eq 1) {
            $fid = $d.FallbackID
            if ($fid -ne $null -and ("" + $fid).Trim() -ne "") {
                $anoms.Add([ordered]@{
                    type = "DisabledHasFallback"
                    deviceId = [int]$d.DeviceID
                    deviceName = $d.DeviceName
                    fallbackId = [int]$fid
                    message = "Disabled device has a fallback set (review if intended)."
                })
            }
        }
    }

    return @($anoms)
}

# -------------------------
# Optional Fix (safe + limited)
# v1 fix policy:
# - Only clears invalid/self fallback IDs: sets FallbackID = NULL
# -------------------------
function Apply-Fix([System.Data.SqlClient.SqlConnection]$Conn, [int]$VenueID, [object[]]$devices, [object[]]$anomalies) {
    $script:Report.results.fix.attempted = $true

    $fixable = $anomalies | Where-Object { $_.type -in @("SelfFallback","MissingFallbackDevice") }
    if (-not $fixable -or $fixable.Count -eq 0) {
        $script:Report.results.fix.supported = $true
        $script:Report.results.fix.applied = $false
        $script:Report.results.fix.reason = "No fixable anomalies found (v1 only clears invalid/self fallback IDs)."
        Out ""
        Out "Apply Fix: nothing fixable found (v1 only clears invalid/self fallback IDs)."
        return
    }

    $msg = @"
Apply Fix will make SQL changes for this venue:

- For printers where FallbackID is invalid OR references itself:
  -> FallbackID will be cleared (set to NULL)

It will NOT automatically assign new fallback printers.

To proceed, tick the checkbox and type YES.
"@

    $confirmed = Show-ConfirmFixDialog -Title "Printer Fallback Audit - Apply Fix" -Message $msg
    if (-not $confirmed) {
        $script:Report.results.fix.applied = $false
        $script:Report.results.fix.reason = "User did not confirm Apply Fix."
        Out ""
        Out "Apply Fix cancelled by user."
        return
    }

    $script:Report.inputs.ApplyFix = $true

    Out ""
    Out "Applying fixes (clearing invalid/self fallback IDs)..."

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

        $sql = "UPDATE dbo.Device SET FallbackID = NULL WHERE DeviceID = @DeviceID;"
        $affected = Invoke-NonQuery -Conn $Conn -Sql $sql -Params @{ DeviceID = $deviceId }

        $dtAfter = Invoke-Query -Conn $Conn -Sql "SELECT d.DeviceID, d.Name AS DeviceName, d.FallbackID, fb.Name AS FallbackDeviceName FROM dbo.Device d LEFT JOIN dbo.Device fb ON fb.DeviceID = d.FallbackID WHERE d.DeviceID = @DeviceID;" -Params @{ DeviceID = $deviceId }
        $afterObj = (DT $dtAfter | Select-Object -First 1)

        $after = [ordered]@{
            deviceId = [int]$afterObj.DeviceID
            deviceName = $afterObj.DeviceName
            afterFallbackId = if ($afterObj.FallbackID -ne $null -and (""+$afterObj.FallbackID).Trim() -ne "") { [int]$afterObj.FallbackID } else { $null }
            afterFallbackName = $afterObj.FallbackDeviceName
            rowsAffected = $affected
        }

        $script:Report.results.fix.beforeAfter += [ordered]@{ before = $before; after = $after }

        Out ("  Cleared fallback for DeviceID {0} ({1}) [RowsAffected={2}]" -f $deviceId, $deviceRow.DeviceName, $affected)
    }

    $script:Report.results.fix.applied = $true
    $script:Report.results.fix.reason = "Cleared invalid/self fallback IDs only (v1 safe fix)."
    Out "Fix complete."
}

# -------------------------
# Main execution
# -------------------------
Out "Bepoz Onboarding Toolkit - Printer Fallback Audit"
Out "RunDir: $($script:Report.runDir)"
Out ""

$conn = $null
try {
    $cfg = Get-SqlConfig

    Out ""
    Out "Connecting to SQL..."
    $conn = Open-SqlConnection -Server $cfg.Server -Database $cfg.Database
    Out "Connected."

    $venues = Get-VenueList -Conn $conn
    if (-not $venues -or $venues.Count -eq 0) { Fail "No venues found in dbo.Venue." 6 }

    $venuePrompt = "Enter VenueID (integer) to audit.`r`nTip: there are $($venues.Count) venues; check dbo.Venue if unsure."
    $rawVenue = Show-InputBox -Title "Printer Fallback Audit" -Prompt $venuePrompt -DefaultValue ""
    if ($null -eq $rawVenue) { Fail "User cancelled VenueID prompt." 7 }

    $rawVenue = $rawVenue.Trim()
    if ($rawVenue -notmatch '^\d+$') { Fail "VenueID must be an integer. You entered: '$rawVenue'." 7 }
    $venueId = [int]$rawVenue
    $script:Report.inputs.VenueID = $venueId

    $venue = Get-Venue -Conn $conn -VenueID $venueId
    if (-not $venue) { Fail "VenueID $venueId not found in dbo.Venue." 8 }
    $script:Report.results.venue = [ordered]@{ VenueID = [int]$venue.VenueID; Name = $venue.Name }

    Out ""
    Out "Auditing Venue: [$venueId] $($venue.Name)"
    Out ""

    $devices = Get-VenuePrintingDevices -Conn $conn -VenueID $venueId

    $script:Report.results.counts.printers = ($devices | Measure-Object).Count
    $withFallback = ($devices | Where-Object { $_.FallbackID -ne $null -and (""+$_.FallbackID).Trim() -ne "" }).Count
    $withoutFallback = $script:Report.results.counts.printers - $withFallback
    $disabled = ($devices | Where-Object { [int]$_.Disabled -eq 1 }).Count

    $script:Report.results.counts.withFallback = $withFallback
    $script:Report.results.counts.withoutFallback = $withoutFallback
    $script:Report.results.counts.disabled = $disabled

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
    Out "  Printers found:        $($script:Report.results.counts.printers)"
    Out "  With fallback:         $withFallback"
    Out "  Without fallback:      $withoutFallback"
    Out "  Disabled printers:     $disabled"
    Out "  Anomalies:             $($script:Report.results.counts.anomalies)"
    Out ""

    if ($script:Report.results.counts.printers -eq 0) {
        Out "NOTE: No printers matched the v1 heuristic for this venue."
        Out "      If this is wrong, we should tighten the definition using known DeviceType/SubType values for printers."
        $script:Report.warnings += "No printers matched v1 heuristic; may need printer DeviceType/SubType list."
    } else {
        Out "Printers (venue-wide):"
        Out "Store | Workstation | DeviceID | DeviceName | Port | Disabled | FallbackID | FallbackDeviceName"
        foreach ($d in $devices) {
            $line = "{0} | {1} | {2} | {3} | {4} | {5} | {6} | {7}" -f `
                ($d.StoreName ?? ""), ($d.WorkstationName ?? ""), $d.DeviceID, ($d.DeviceName ?? ""), ($d.PortName ?? ""), $d.Disabled, ($d.FallbackID ?? ""), ($d.FallbackDeviceName ?? "")
            Out $line
        }
        Out ""
    }

    if ($script:Report.results.counts.anomalies -gt 0) {
        Out "Anomalies:"
        foreach ($a in $anoms) {
            Out ("- {0}: DeviceID {1} ({2}) -> {3}" -f $a.type, $a.deviceId, $a.deviceName, $a.message)
        }
        Out ""
    } else {
        Out "No anomalies detected (based on v1 rules)."
        Out ""
    }

    Apply-Fix -Conn $conn -VenueID $venueId -devices $devices -anomalies $anoms

    $script:Report.status = "success"
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
