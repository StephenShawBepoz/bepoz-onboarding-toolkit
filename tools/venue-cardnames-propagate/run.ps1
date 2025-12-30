<# 
  Tool: venue-cardnames-propagate
  Purpose:
    - AuditOnly (default): detect Bepoz apps under C:\Bepoz\Programs and preview SQL changes
    - ApplyFix: close those apps, copy cardname_1..16 from master VenueID to all other venues, reopen apps
  Notes:
    - SQL context must come from env/ToolkitContext.json; NO registry reads.
    - Uses Windows Integrated Auth via ODBC DSN (BEPOZ_SQL_DSN).
#>

[CmdletBinding()]
param(
  [Parameter(Mandatory = $true)]
  [string]$RunDir,

  # Safe-by-default switches
  [switch]$AuditOnly,
  [switch]$ApplyFix,
  [switch]$Force,

  # Optional: for non-interactive runs (no UI prompts)
  [int]$MasterVenueId,

  # Optional: where to detect Bepoz apps to close
  [string]$ProgramsRoot = "C:\Bepoz\Programs"
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$ToolId      = "venue-cardnames-propagate"
$ToolVersion = "1.0.0"

function Write-Info([string]$m) { Write-Host "[INFO] $m" }
function Write-Warn([string]$m) { Write-Host "[WARN] $m" }
function Write-Err ([string]$m) { Write-Host "[ERR ] $m" }

function Test-IsInteractive {
  try {
    return [Environment]::UserInteractive -and ($null -ne $Host.UI.RawUI)
  } catch {
    return $false
  }
}

function Get-MachineInfo {
  $osVersion = $null
  try { $osVersion = (Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop).Version } catch {}
  if (-not $osVersion) {
    try { $osVersion = [System.Environment]::OSVersion.Version.ToString() } catch { $osVersion = "" }
  }

  $user = ""
  try {
    $dom = [Environment]::UserDomainName
    $usr = [Environment]::UserName
    if ($dom) { $user = "$dom\$usr" } else { $user = $usr }
  } catch {}

  [pscustomobject]@{
    computerName = $env:COMPUTERNAME
    userName     = $user
    osVersion    = $osVersion
    psVersion    = $PSVersionTable.PSVersion.ToString()
  }
}

function Read-ToolkitContextFile([string]$path) {
  if (-not (Test-Path -LiteralPath $path)) { return $null }
  try {
    $raw = Get-Content -LiteralPath $path -Raw -Encoding UTF8
    if (-not $raw) { return $null }
    return ($raw | ConvertFrom-Json)
  } catch {
    return $null
  }
}

function Resolve-SqlContext([string]$runDir) {
  $server = ($env:BEPOZ_SQL_SERVER  | ForEach-Object { "$_".Trim() })
  $dsn    = ($env:BEPOZ_SQL_DSN     | ForEach-Object { "$_".Trim() })
  $regp   = ($env:BEPOZ_SQL_REGPATH | ForEach-Object { "$_".Trim() })

  $ctxPath = Join-Path $runDir "ToolkitContext.json"
  $ctx = Read-ToolkitContextFile -path $ctxPath

  if ((-not $server -or $server -eq "") -and $ctx) {
    try { $server = [string]$ctx.backOffice.sql.server } catch {}
  }
  if ((-not $dsn -or $dsn -eq "") -and $ctx) {
    try { $dsn = [string]$ctx.backOffice.sql.dsnName } catch {}
  }
  if ((-not $regp -or $regp -eq "") -and $ctx) {
    try { $regp = [string]$ctx.backOffice.registryPath } catch {}
  }

  [pscustomobject]@{
    server      = ($server | ForEach-Object { "$_".Trim() })
    dsn         = ($dsn    | ForEach-Object { "$_".Trim() })
    registryPath= ($regp   | ForEach-Object { "$_".Trim() })
    contextPath = $ctxPath
    contextOk   = [bool]$ctx
  }
}

function Write-Report([string]$path, [object]$reportObj) {
  $json = $reportObj | ConvertTo-Json -Depth 14
  Set-Content -LiteralPath $path -Value $json -Encoding UTF8
}

# --- CommandLine parsing (for restart) via CommandLineToArgvW ---
Add-Type -Language CSharp -Namespace Bepoz -Name CmdLine -MemberDefinition @"
using System;
using System.Runtime.InteropServices;

public static class CmdLine {
  [DllImport("shell32.dll", SetLastError=true)]
  static extern IntPtr CommandLineToArgvW([MarshalAs(UnmanagedType.LPWStr)] string lpCmdLine, out int pNumArgs);

  [DllImport("kernel32.dll")]
  static extern IntPtr LocalFree(IntPtr hMem);

  public static string[] Split(string commandLine) {
    if (string.IsNullOrWhiteSpace(commandLine)) return new string[0];
    int argc;
    IntPtr argv = CommandLineToArgvW(commandLine, out argc);
    if (argv == IntPtr.Zero) return new string[0];

    try {
      var args = new string[argc];
      for (int i = 0; i < argc; i++) {
        IntPtr p = Marshal.ReadIntPtr(argv, i * IntPtr.Size);
        args[i] = Marshal.PtrToStringUni(p);
      }
      return args;
    } finally {
      LocalFree(argv);
    }
  }
}
"@ | Out-Null

function Get-BepozProcesses([string]$root) {
  $rootNorm = $root.TrimEnd('\')
  $procs = @()
  try {
    $cim = Get-CimInstance Win32_Process -ErrorAction Stop
    foreach ($p in $cim) {
      if ($p.ExecutablePath -and $p.ExecutablePath.StartsWith($rootNorm, [System.StringComparison]::OrdinalIgnoreCase)) {
        $procs += [pscustomobject]@{
          ProcessId      = [int]$p.ProcessId
          Name           = [string]$p.Name
          ExecutablePath = [string]$p.ExecutablePath
          CommandLine    = [string]$p.CommandLine
        }
      }
    }
  } catch {
    throw "Failed to enumerate processes: $($_.Exception.Message)"
  }
  return $procs
}

function Stop-ProcessesGracefully([object[]]$procInfo, [int]$waitMs = 5000) {
  $stopped = New-Object System.Collections.Generic.List[object]
  foreach ($p in $procInfo) {
    $pid = $p.ProcessId
    try {
      $gp = Get-Process -Id $pid -ErrorAction Stop
      $closedMain = $false
      try {
        if ($gp.MainWindowHandle -ne 0) {
          $closedMain = $gp.CloseMainWindow()
        }
      } catch {}

      $sw = [System.Diagnostics.Stopwatch]::StartNew()
      while (-not $gp.HasExited -and $sw.ElapsedMilliseconds -lt $waitMs) {
        Start-Sleep -Milliseconds 200
        try { $gp.Refresh() } catch {}
      }

      if (-not $gp.HasExited) {
        Stop-Process -Id $pid -Force -ErrorAction Stop
      }

      $stopped.Add([pscustomobject]@{
        processId = $pid
        name = $p.Name
        executablePath = $p.ExecutablePath
        commandLine = $p.CommandLine
        result = "stopped"
        closedMainWindow = $closedMain
      }) | Out-Null
    } catch {
      $stopped.Add([pscustomobject]@{
        processId = $pid
        name = $p.Name
        executablePath = $p.ExecutablePath
        commandLine = $p.CommandLine
        result = "failed"
        error = $_.Exception.Message
      }) | Out-Null
    }
  }
  return @($stopped)
}

function Restart-Processes([object[]]$stoppedInfo) {
  $restarted = New-Object System.Collections.Generic.List[object]
  foreach ($p in $stoppedInfo) {
    if ($p.result -ne "stopped") {
      $restarted.Add([pscustomobject]@{
        processId = $p.processId
        name = $p.name
        result = "skipped"
        details = "Was not stopped successfully."
      }) | Out-Null
      continue
    }

    try {
      $exe = $p.executablePath
      $args = @()

      if ($p.commandLine) {
        $parts = [Bepoz.CmdLine]::Split($p.commandLine)
        if ($parts.Length -ge 1) {
          # parts[0] should be the exe path; remainder are args
          if ($parts.Length -gt 1) { $args = $parts[1..($parts.Length-1)] }
        }
      }

      if (-not (Test-Path -LiteralPath $exe)) {
        $restarted.Add([pscustomobject]@{
          name = $p.name
          result = "failed"
          error = "Executable not found: $exe"
        }) | Out-Null
        continue
      }

      $sp = $null
      if ($args.Count -gt 0) {
        $sp = Start-Process -FilePath $exe -ArgumentList $args -PassThru -ErrorAction Stop
      } else {
        $sp = Start-Process -FilePath $exe -PassThru -ErrorAction Stop
      }

      $restarted.Add([pscustomobject]@{
        name = $p.name
        executablePath = $exe
        arguments = @($args)
        result = "started"
        newProcessId = $sp.Id
      }) | Out-Null
    } catch {
      $restarted.Add([pscustomobject]@{
        name = $p.name
        executablePath = $p.executablePath
        result = "failed"
        error = $_.Exception.Message
      }) | Out-Null
    }
  }
  return @($restarted)
}

function New-OdbcConnection([string]$dsn) {
  # Integrated auth
  $cs = "DSN=$dsn;Trusted_Connection=Yes;"
  $conn = New-Object System.Data.Odbc.OdbcConnection($cs)
  return $conn
}

function Invoke-OdbcScalar([System.Data.Odbc.OdbcConnection]$conn, [string]$sql, [object[]]$params) {
  $cmd = $conn.CreateCommand()
  $cmd.CommandText = $sql
  $cmd.CommandTimeout = 10

  if ($params) {
    foreach ($p in $params) {
      $null = $cmd.Parameters.Add("@p", [System.Data.Odbc.OdbcType]::VarChar)
      $cmd.Parameters[$cmd.Parameters.Count-1].Value = $p
    }
  }

  return $cmd.ExecuteScalar()
}

function Invoke-OdbcQuery([System.Data.Odbc.OdbcConnection]$conn, [string]$sql, [object[]]$params) {
  $cmd = $conn.CreateCommand()
  $cmd.CommandText = $sql
  $cmd.CommandTimeout = 10

  if ($params) {
    foreach ($p in $params) {
      $null = $cmd.Parameters.Add("@p", [System.Data.Odbc.OdbcType]::VarChar)
      $cmd.Parameters[$cmd.Parameters.Count-1].Value = $p
    }
  }

  $da = New-Object System.Data.Odbc.OdbcDataAdapter($cmd)
  $dt = New-Object System.Data.DataTable
  [void]$da.Fill($dt)
  return $dt
}

function Invoke-OdbcNonQuery([System.Data.Odbc.OdbcConnection]$conn, [string]$sql, [object[]]$params) {
  $cmd = $conn.CreateCommand()
  $cmd.CommandText = $sql
  $cmd.CommandTimeout = 30

  if ($params) {
    foreach ($p in $params) {
      $null = $cmd.Parameters.Add("@p", [System.Data.Odbc.OdbcType]::VarChar)
      $cmd.Parameters[$cmd.Parameters.Count-1].Value = $p
    }
  }

  return $cmd.ExecuteNonQuery()
}

function Select-MasterVenueInteractive([object[]]$venues) {
  Add-Type -AssemblyName System.Windows.Forms
  Add-Type -AssemblyName System.Drawing

  $form = New-Object System.Windows.Forms.Form
  $form.Text = "Select Master Venue"
  $form.Width = 520
  $form.Height = 520
  $form.StartPosition = "CenterScreen"
  $form.TopMost = $true

  $label = New-Object System.Windows.Forms.Label
  $label.AutoSize = $true
  $label.Text = "Select the VenueID to use as the master for cardname_1..16:"
  $label.Location = New-Object System.Drawing.Point(10, 10)
  $form.Controls.Add($label)

  $list = New-Object System.Windows.Forms.ListBox
  $list.Location = New-Object System.Drawing.Point(10, 40)
  $list.Width = 480
  $list.Height = 380
  foreach ($v in $venues) {
    [void]$list.Items.Add(("{0} - {1}" -f $v.VenueID, $v.Name))
  }
  $form.Controls.Add($list)

  $ok = New-Object System.Windows.Forms.Button
  $ok.Text = "OK"
  $ok.Location = New-Object System.Drawing.Point(310, 435)
  $ok.Add_Click({
    if ($list.SelectedItem -eq $null) { return }
    $form.Tag = $list.SelectedItem
    $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.Close()
  })
  $form.Controls.Add($ok)

  $cancel = New-Object System.Windows.Forms.Button
  $cancel.Text = "Cancel"
  $cancel.Location = New-Object System.Drawing.Point(400, 435)
  $cancel.Add_Click({
    $form.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.Close()
  })
  $form.Controls.Add($cancel)

  $dr = $form.ShowDialog()
  if ($dr -ne [System.Windows.Forms.DialogResult]::OK) { return $null }

  $sel = [string]$form.Tag
  # Parse leading VenueID
  if ($sel -match '^\s*(\d+)\s*-') { return [int]$matches[1] }
  return $null
}

function Confirm-Apply([string]$message) {
  if ($Force) { return $true }

  if (-not (Test-IsInteractive)) {
    return $false
  }

  $ans = Read-Host "$message (type YES to continue)"
  return ($ans -eq "YES")
}

# ---------------------- Main ----------------------

# Ensure RunDir exists
try {
  if (-not (Test-Path -LiteralPath $RunDir)) {
    New-Item -ItemType Directory -Force -Path $RunDir | Out-Null
  }
} catch {
  Write-Err "RunDir could not be created/accessed: $RunDir"
  exit 2
}

$reportPath = Join-Path $RunDir "Report.json"

# Determine effective mode
$effectiveApplyFix = [bool]$ApplyFix
$effectiveAuditOnly = -not $effectiveApplyFix

if ($AuditOnly.IsPresent -and $effectiveApplyFix) {
  Write-Warn "Both -AuditOnly and -ApplyFix were provided. Continuing safely in audit mode."
  $effectiveApplyFix = $false
  $effectiveAuditOnly = $true
}

$warnings = New-Object System.Collections.Generic.List[string]
$errors   = New-Object System.Collections.Generic.List[string]
$findings = New-Object System.Collections.Generic.List[object]
$actions  = New-Object System.Collections.Generic.List[object]

$swTotal = [System.Diagnostics.Stopwatch]::StartNew()

Write-Info "ToolId: $ToolId v$ToolVersion"
Write-Info "RunDir: $RunDir"
Write-Info "Mode:   AuditOnly=$effectiveAuditOnly ApplyFix=$effectiveApplyFix"
Write-Info "Apps:   ProgramsRoot='$ProgramsRoot'"

$sqlContext = Resolve-SqlContext -runDir $RunDir
Write-Info "SQL:    Server='$($sqlContext.server)' DSN='$($sqlContext.dsn)' (context file present=$($sqlContext.contextOk))"

if (-not $sqlContext.dsn) {
  $errors.Add("BEPOZ_SQL_DSN is missing (env and ToolkitContext.json fallback).") | Out-Null
  $findings.Add([pscustomobject]@{
    severity = "critical"
    title   = "Missing SQL DSN context"
    details = "Cannot connect to SQL without BEPOZ_SQL_DSN."
    evidence= @{ env = @{ BEPOZ_SQL_DSN = $env:BEPOZ_SQL_DSN }; toolkitContextPath = $sqlContext.contextPath }
  }) | Out-Null
}

# Detect running Bepoz apps
$running = @()
try {
  $running = Get-BepozProcesses -root $ProgramsRoot
  Write-Info ("Detected {0} running Bepoz process(es) under {1}" -f $running.Count, $ProgramsRoot)
  $findings.Add([pscustomobject]@{
    severity = "info"
    title   = "Running Bepoz apps detected"
    details = "Processes with executable paths under the configured ProgramsRoot."
    evidence= @{ count = $running.Count; processes = @($running) }
  }) | Out-Null
} catch {
  $warnings.Add("Could not enumerate running apps: $($_.Exception.Message)") | Out-Null
}

# If DSN missing, write report and exit
if ($errors.Count -gt 0) {
  $swTotal.Stop()
  $report = [pscustomobject]([ordered]@{
    toolId      = $ToolId
    toolVersion = $ToolVersion
    runAt       = (Get-Date).ToString("o")
    machine     = Get-MachineInfo
    success     = $false
    mode        = @{ auditOnly = $effectiveAuditOnly; applyFix = $effectiveApplyFix }
    sqlContext  = @{
      server       = $sqlContext.server
      dsn          = $sqlContext.dsn
      registryPath = $sqlContext.registryPath
    }
    warnings    = @($warnings)
    errors      = @($errors)
    findings    = @($findings)
    actions     = @($actions)
    timingsMs   = @{ total = $swTotal.ElapsedMilliseconds }
  })
  Write-Report -path $reportPath -reportObj $report
  Write-Err "Cannot proceed. Report written: $reportPath"
  exit 2
}

# Connect to SQL (ODBC DSN)
$conn = $null
$swConn = [System.Diagnostics.Stopwatch]::StartNew()
try {
  $conn = New-OdbcConnection -dsn $sqlContext.dsn
  $conn.Open()
  $swConn.Stop()
  Write-Info ("SQL connection OK ({0}ms)" -f $swConn.ElapsedMilliseconds)

  $findings.Add([pscustomobject]@{
    severity = "info"
    title   = "SQL connection succeeded"
    details = "Connected using Windows Integrated Auth via ODBC DSN."
    evidence= @{ dsn = $sqlContext.dsn; connectMs = $swConn.ElapsedMilliseconds }
  }) | Out-Null
} catch {
  $swConn.Stop()
  $errors.Add("SQL connection failed: $($_.Exception.Message)") | Out-Null
}

if ($errors.Count -gt 0) {
  try { if ($conn) { $conn.Dispose() } } catch {}
  $swTotal.Stop()
  $report = [pscustomobject]([ordered]@{
    toolId      = $ToolId
    toolVersion = $ToolVersion
    runAt       = (Get-Date).ToString("o")
    machine     = Get-MachineInfo
    success     = $false
    mode        = @{ auditOnly = $effectiveAuditOnly; applyFix = $effectiveApplyFix }
    sqlContext  = @{
      server       = $sqlContext.server
      dsn          = $sqlContext.dsn
      registryPath = $sqlContext.registryPath
    }
    warnings    = @($warnings)
    errors      = @($errors)
    findings    = @($findings)
    actions     = @($actions)
    timingsMs   = @{ connect = $swConn.ElapsedMilliseconds; total = $swTotal.ElapsedMilliseconds }
  })
  Write-Report -path $reportPath -reportObj $report
  Write-Err "Connection failed. Report written: $reportPath"
  exit 1
}

# Load venue list
$venues = @()
try {
  $dtVenues = Invoke-OdbcQuery -conn $conn -sql "SELECT VenueID, Name FROM dbo.Venue ORDER BY VenueID;" -params @()
  foreach ($r in $dtVenues.Rows) {
    $venues += [pscustomobject]@{ VenueID = [int]$r["VenueID"]; Name = [string]$r["Name"] }
  }
  Write-Info ("Loaded {0} venue(s) from dbo.Venue" -f $venues.Count)
} catch {
  $errors.Add("Failed to read dbo.Venue: $($_.Exception.Message)") | Out-Null
}

# Select master venue
$selectedMaster = $null
if ($MasterVenueId -gt 0) {
  $selectedMaster = $MasterVenueId
} else {
  if (Test-IsInteractive) {
    $selectedMaster = Select-MasterVenueInteractive -venues $venues
  } else {
    $errors.Add("Non-interactive run: -MasterVenueId is required.") | Out-Null
  }
}

if (-not $selectedMaster) {
  if ($errors.Count -eq 0) { $errors.Add("No master venue selected (cancelled or invalid selection).") | Out-Null }
}

# Validate master exists
if ($selectedMaster) {
  $exists = $venues | Where-Object { $_.VenueID -eq $selectedMaster } | Select-Object -First 1
  if (-not $exists) {
    $errors.Add("Master VenueID $selectedMaster was not found in dbo.Venue.") | Out-Null
  }
}

# Discover table containing VenueID + cardname_1..16 (dynamic; avoids hardcoding schema)
$schemaName = $null
$tableName  = $null
$matches    = @()

try {
  $colList = @("VenueID")
  for ($i=1; $i -le 16; $i++) { $colList += ("cardname_{0}" -f $i) }
  $inList = ($colList | ForEach-Object { "'" + $_ + "'" }) -join ","

  $sqlDiscover = @"
SELECT s.name AS SchemaName, t.name AS TableName,
       SUM(CASE WHEN c.name = 'VenueID' THEN 1 ELSE 0 END) AS HasVenueID,
       SUM(CASE WHEN c.name IN ($inList) AND c.name <> 'VenueID' THEN 1 ELSE 0 END) AS CardCols
FROM sys.tables t
JOIN sys.schemas s ON t.schema_id = s.schema_id
JOIN sys.columns c ON t.object_id = c.object_id
WHERE c.name IN ($inList)
GROUP BY s.name, t.name
HAVING SUM(CASE WHEN c.name = 'VenueID' THEN 1 ELSE 0 END) = 1
   AND SUM(CASE WHEN c.name IN ($inList) AND c.name <> 'VenueID' THEN 1 ELSE 0 END) = 16
ORDER BY s.name, t.name;
"@

  $dtMatch = Invoke-OdbcQuery -conn $conn -sql $sqlDiscover -params @()
  foreach ($r in $dtMatch.Rows) {
    $matches += [pscustomobject]@{ SchemaName = [string]$r["SchemaName"]; TableName = [string]$r["TableName"] }
  }

  if ($matches.Count -ge 1) {
    $schemaName = $matches[0].SchemaName
    $tableName  = $matches[0].TableName
    Write-Info ("Discovered cardname table: [{0}].[{1}] (matches={2})" -f $schemaName, $tableName, $matches.Count)
  } else {
    $errors.Add("Could not find a table containing VenueID and cardname_1..cardname_16.") | Out-Null
  }

  $findings.Add([pscustomobject]@{
    severity = $(if ($matches.Count -ge 1) { "info" } else { "critical" })
    title   = "Cardname table discovery"
    details = $(if ($matches.Count -ge 1) { "Found candidate table(s)." } else { "No suitable table found." })
    evidence= @{ matches = @($matches) }
  }) | Out-Null
} catch {
  $errors.Add("Failed to discover cardname table: $($_.Exception.Message)") | Out-Null
}

# If errors, write report and exit
if ($errors.Count -gt 0) {
  try { if ($conn) { $conn.Dispose() } } catch {}
  $swTotal.Stop()
  $report = [pscustomobject]([ordered]@{
    toolId      = $ToolId
    toolVersion = $ToolVersion
    runAt       = (Get-Date).ToString("o")
    machine     = Get-MachineInfo
    success     = $false
    mode        = @{ auditOnly = $effectiveAuditOnly; applyFix = $effectiveApplyFix }
    sqlContext  = @{
      server       = $sqlContext.server
      dsn          = $sqlContext.dsn
      registryPath = $sqlContext.registryPath
    }
    warnings    = @($warnings)
    errors      = @($errors)
    findings    = @($findings)
    actions     = @($actions)
    timingsMs   = @{ connect = $swConn.ElapsedMilliseconds; total = $swTotal.ElapsedMilliseconds }
  })
  Write-Report -path $reportPath -reportObj $report
  Write-Err "Cannot proceed. Report written: $reportPath"
  exit 2
}

# Read master cardname values
$masterValues = $null
$swQuery = [System.Diagnostics.Stopwatch]::StartNew()
try {
  $cols = @()
  for ($i=1; $i -le 16; $i++) { $cols += ("cardname_{0}" -f $i) }
  $colSql = ($cols | ForEach-Object { "[" + $_ + "]" }) -join ", "

  # ODBC parameter placeholder is '?'
  $sqlMaster = "SELECT $colSql FROM [$schemaName].[$tableName] WHERE VenueID = ?;"
  $dtMaster = Invoke-OdbcQuery -conn $conn -sql $sqlMaster -params @($selectedMaster)

  if ($dtMaster.Rows.Count -ne 1) {
    throw "Expected 1 row for master VenueID $selectedMaster in [$schemaName].[$tableName], got $($dtMaster.Rows.Count)."
  }

  $row = $dtMaster.Rows[0]
  $masterValues = [ordered]@{}
  for ($i=1; $i -le 16; $i++) {
    $k = "cardname_{0}" -f $i
    $masterValues[$k] = [string]$row[$k]
  }

  $swQuery.Stop()
  Write-Info ("Read master cardnames from VenueID {0} ({1}ms)" -f $selectedMaster, $swQuery.ElapsedMilliseconds)

  $findings.Add([pscustomobject]@{
    severity = "info"
    title   = "Master values loaded"
    details = "Read cardname_1..16 for the selected master venue."
    evidence= @{ masterVenueId = $selectedMaster; table = "[$schemaName].[$tableName]"; queryMs = $swQuery.ElapsedMilliseconds }
  }) | Out-Null
} catch {
  $swQuery.Stop()
  $errors.Add("Failed to read master cardnames: $($_.Exception.Message)") | Out-Null
}

# Determine targets (all venues except master)
$targetVenueIds = @()
if ($venues) {
  $targetVenueIds = @($venues | Where-Object { $_.VenueID -ne $selectedMaster } | Select-Object -ExpandProperty VenueID)
}

$findings.Add([pscustomobject]@{
  severity = "info"
  title   = "Target venues"
  details = "All venues excluding the master will be updated (when ApplyFix is used)."
  evidence= @{ masterVenueId = $selectedMaster; targetCount = $targetVenueIds.Count; targetVenueIds = @($targetVenueIds) }
}) | Out-Null

# If master read failed, exit
if ($errors.Count -gt 0) {
  try { if ($conn) { $conn.Dispose() } } catch {}
  $swTotal.Stop()
  $report = [pscustomobject]([ordered]@{
    toolId      = $ToolId
    toolVersion = $ToolVersion
    runAt       = (Get-Date).ToString("o")
    machine     = Get-MachineInfo
    success     = $false
    mode        = @{ auditOnly = $effectiveAuditOnly; applyFix = $effectiveApplyFix }
    sqlContext  = @{
      server       = $sqlContext.server
      dsn          = $sqlContext.dsn
      registryPath = $sqlContext.registryPath
    }
    warnings    = @($warnings)
    errors      = @($errors)
    findings    = @($findings)
    actions     = @($actions)
    timingsMs   = @{ connect = $swConn.ElapsedMilliseconds; masterQuery = $swQuery.ElapsedMilliseconds; total = $swTotal.ElapsedMilliseconds }
  })
  Write-Report -path $reportPath -reportObj $report
  Write-Err "Failed. Report written: $reportPath"
  exit 1
}

# AuditOnly: no changes, just report
if ($effectiveAuditOnly) {
  try { if ($conn) { $conn.Dispose() } } catch {}
  $swTotal.Stop()

  $findings.Add([pscustomobject]@{
    severity = "info"
    title   = "Audit mode"
    details = "No applications were closed and no SQL updates were performed."
    evidence= @{ wouldCloseProcessCount = $running.Count; wouldUpdateTargetCount = $targetVenueIds.Count }
  }) | Out-Null

  $report = [pscustomobject]([ordered]@{
    toolId      = $ToolId
    toolVersion = $ToolVersion
    runAt       = (Get-Date).ToString("o")
    machine     = Get-MachineInfo
    success     = $true
    mode        = @{ auditOnly = $true; applyFix = $false }
    sqlContext  = @{
      server       = $sqlContext.server
      dsn          = $sqlContext.dsn
      registryPath = $sqlContext.registryPath
    }
    warnings    = @($warnings)
    errors      = @($errors)
    findings    = @($findings)
    actions     = @($actions)
    timingsMs   = @{
      connect    = $swConn.ElapsedMilliseconds
      masterQuery= $swQuery.ElapsedMilliseconds
      total      = $swTotal.ElapsedMilliseconds
    }
  })

  Write-Report -path $reportPath -reportObj $report
  Write-Info "Audit complete. Report written: $reportPath"
  exit 0
}

# ApplyFix requires confirmation unless -Force
if (-not (Confirm-Apply -message ("This will CLOSE {0} running app(s), UPDATE cardname_1..16 for {1} venue(s), then REOPEN apps. Continue?" -f $running.Count, $targetVenueIds.Count))) {
  try { if ($conn) { $conn.Dispose() } } catch {}
  $swTotal.Stop()
  $actions.Add([pscustomobject]@{
    title  = "User confirmation"
    details= "User declined the ApplyFix operation."
    result = "failed"
  }) | Out-Null

  $report = [pscustomobject]([ordered]@{
    toolId      = $ToolId
    toolVersion = $ToolVersion
    runAt       = (Get-Date).ToString("o")
    machine     = Get-MachineInfo
    success     = $false
    mode        = @{ auditOnly = $false; applyFix = $true }
    sqlContext  = @{
      server       = $sqlContext.server
      dsn          = $sqlContext.dsn
      registryPath = $sqlContext.registryPath
    }
    warnings    = @($warnings)
    errors      = @("User cancelled / declined confirmation.")
    findings    = @($findings)
    actions     = @($actions)
    timingsMs   = @{ total = $swTotal.ElapsedMilliseconds }
  })

  Write-Report -path $reportPath -reportObj $report
  Write-Warn "Cancelled. Report written: $reportPath"
  exit 3
}

# 1) Close apps
Write-Info "Closing running Bepoz apps..."
$closed = Stop-ProcessesGracefully -procInfo $running -waitMs 6000
$actions.Add([pscustomobject]@{
  title  = "Close Bepoz apps"
  details= "Stopped processes under ProgramsRoot prior to SQL update."
  result = "success"
}) | Out-Null

# 2) Update DB in a transaction
$rowsAffected = $null
$swUpdate = [System.Diagnostics.Stopwatch]::StartNew()
try {
  $tx = $conn.BeginTransaction()
  $cmd = $conn.CreateCommand()
  $cmd.Transaction = $tx
  $cmd.CommandTimeout = 60

  # Build update statement: set cardname_1..16 = ? ... where VenueID <> ?
  $sets = @()
  for ($i=1; $i -le 16; $i++) { $sets += ("[cardname_{0}] = ?" -f $i) }
  $setSql = $sets -join ", "

  $cmd.CommandText = "UPDATE [$schemaName].[$tableName] SET $setSql WHERE VenueID <> ?;"

  # Add parameters in order
  $cmd.Parameters.Clear() | Out-Null
  for ($i=1; $i -le 16; $i++) {
    $p = $cmd.Parameters.Add("@p", [System.Data.Odbc.OdbcType]::NVarChar, 255)
    $p.Value = [string]$masterValues[("cardname_{0}" -f $i)]
  }
  $pM = $cmd.Parameters.Add("@p", [System.Data.Odbc.OdbcType]::Int)
  $pM.Value = $selectedMaster

  $rowsAffected = $cmd.ExecuteNonQuery()
  $tx.Commit()

  $swUpdate.Stop()
  Write-Info ("SQL update committed. Rows affected: {0} ({1}ms)" -f $rowsAffected, $swUpdate.ElapsedMilliseconds)

  $actions.Add([pscustomobject]@{
    title  = "Propagate cardnames"
    details= "Updated cardname_1..16 from master VenueID to all other venues."
    result = "success"
  }) | Out-Null

  $findings.Add([pscustomobject]@{
    severity = "info"
    title   = "SQL update succeeded"
    details = "Cardnames were propagated to all venues except the master."
    evidence= @{ table = "[$schemaName].[$tableName]"; masterVenueId = $selectedMaster; affectedRows = $rowsAffected; updateMs = $swUpdate.ElapsedMilliseconds }
  }) | Out-Null
} catch {
  $swUpdate.Stop()
  try { if ($tx) { $tx.Rollback() } } catch {}
  $errors.Add("SQL update failed (rolled back): $($_.Exception.Message)") | Out-Null

  $actions.Add([pscustomobject]@{
    title  = "Propagate cardnames"
    details= "Attempted update but failed; transaction rolled back."
    result = "failed"
  }) | Out-Null
}

# 3) Reopen apps (best effort, even if SQL failed)
Write-Info "Restarting previously closed apps..."
$reopened = Restart-Processes -stoppedInfo $closed
$actions.Add([pscustomobject]@{
  title  = "Reopen Bepoz apps"
  details= "Restarted processes previously stopped under ProgramsRoot."
  result = "success"
}) | Out-Null

try { if ($conn) { $conn.Dispose() } } catch {}

$swTotal.Stop()
$success = ($errors.Count -eq 0)

$report = [pscustomobject]([ordered]@{
  toolId      = $ToolId
  toolVersion = $ToolVersion
  runAt       = (Get-Date).ToString("o")
  machine     = Get-MachineInfo
  success     = [bool]$success
  mode        = @{ auditOnly = $false; applyFix = $true }
  sqlContext  = @{
    server       = $sqlContext.server
    dsn          = $sqlContext.dsn
    registryPath = $sqlContext.registryPath
  }
  warnings    = @($warnings)
  errors      = @($errors)
  findings    = @($findings)
  actions     = @($actions)
  evidence    = @{
    masterVenueId = $selectedMaster
    masterCardnames = $masterValues  # NOTE: these are just labels; no secrets
    discoveredTable = "[$schemaName].[$tableName]"
    processesDetected = @($running)
    processesClosed = @($closed)
    processesRestarted = @($reopened)
    rowsAffected = $rowsAffected
  }
  timingsMs   = @{
    connect = $swConn.ElapsedMilliseconds
    masterQuery = $swQuery.ElapsedMilliseconds
    update = $swUpdate.ElapsedMilliseconds
    total = $swTotal.ElapsedMilliseconds
  }
})

Write-Report -path $reportPath -reportObj $report
Write-Info "Completed. Report written: $reportPath"

if ($success) { exit 0 }
exit 1
