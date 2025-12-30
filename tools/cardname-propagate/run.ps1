#requires -version 5.1
<#
Bepoz Toolkit Tool: cardname-propagate
- Audit-only by default.
- Apply requires -ApplyFix and confirmation (or -Force for unattended).
- Writes Report.json into RunDir.
#>

param(
  [Parameter(Mandatory = $true)]
  [string]$RunDir,

  [switch]$AuditOnly,
  [switch]$ApplyFix,
  [switch]$Force,

  [int]$MasterVenueId,

  # Optional: close Bepoz apps running from these paths before DB changes
  [switch]$CloseBepozApps,

  [string]$BepozProgramsRoot = "C:\Bepoz\Programs",
  [string]$BepozRoot = "C:\Bepoz"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# ----------------------------
# Exit codes
# 0 = success
# 1 = unexpected failure
# 2 = validation/input error
# 3 = user cancelled / declined confirmation
# ----------------------------

function Write-Log {
  param(
    [ValidateSet("INFO","WARN","ERROR")]
    [string]$Level,
    [string]$Message
  )
  $ts = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
  Write-Host "[$ts] [$Level] $Message"
}

function Test-IsInteractive {
  try {
    return [Environment]::UserInteractive -and ($Host.Name -eq "ConsoleHost")
  } catch {
    return $false
  }
}

function Ensure-RunDir {
  if (-not (Test-Path -LiteralPath $RunDir)) {
    New-Item -ItemType Directory -Path $RunDir -Force | Out-Null
  }
}

function New-ReportObject {
  param([hashtable]$SqlContext, [bool]$AuditMode, [bool]$ApplyMode)

  $os = (Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction SilentlyContinue)
  $osVer = if ($os) { $os.Version } else { [Environment]::OSVersion.Version.ToString() }

  return [ordered]@{
    toolId      = "cardname-propagate"
    toolVersion = "1.0.0"
    runAt       = (Get-Date).ToString("o")
    machine     = [ordered]@{
      computerName = $env:COMPUTERNAME
      userName     = "$($env:USERDOMAIN)\$($env:USERNAME)"
      osVersion    = $osVer
      psVersion    = $PSVersionTable.PSVersion.ToString()
    }
    success   = $true
    mode      = [ordered]@{
      auditOnly = $AuditMode
      applyFix  = $ApplyMode
    }
    sqlContext = [ordered]@{
      server       = $SqlContext.server
      dsn          = $SqlContext.dsn
      registryPath = $SqlContext.registryPath
      database     = $SqlContext.database
      provider     = $SqlContext.provider
    }
    warnings = @()
    errors   = @()
    findings = @()
    actions  = @()
  }
}

function Add-Finding {
  param(
    [hashtable]$Report,
    [ValidateSet("info","warning","critical")]
    [string]$Severity,
    [string]$Title,
    [string]$Details,
    $Evidence
  )
  $Report.findings += [ordered]@{
    severity = $Severity
    title    = $Title
    details  = $Details
    evidence = $Evidence
  }
}

function Add-Action {
  param(
    [hashtable]$Report,
    [string]$Title,
    [string]$Details,
    [ValidateSet("success","failed")]
    [string]$Result
  )
  $Report.actions += [ordered]@{
    title   = $Title
    details = $Details
    result  = $Result
  }
}

function Save-Report {
  param([hashtable]$Report)
  $path = Join-Path $RunDir "Report.json"
  $json = $Report | ConvertTo-Json -Depth 8
  [System.IO.File]::WriteAllText($path, $json, [System.Text.Encoding]::UTF8)
  Write-Log INFO "Wrote report: $path"
}

function Get-ToolkitSqlContext {
  $ctx = [ordered]@{
    server       = $null
    dsn          = $null
    registryPath = $null
    database     = $null
    provider     = $null
  }

  # Environment variables (preferred)
  if ($env:BEPOZ_SQL_SERVER)   { $ctx.server       = $env:BEPOZ_SQL_SERVER }
  if ($env:BEPOZ_SQL_DSN)      { $ctx.dsn          = $env:BEPOZ_SQL_DSN }
  if ($env:BEPOZ_SQL_REGPATH)  { $ctx.registryPath = $env:BEPOZ_SQL_REGPATH }
  if ($env:BEPOZ_SQL_DATABASE) { $ctx.database     = $env:BEPOZ_SQL_DATABASE }

  # ToolkitContext.json (if present)
  $ctxPath = Join-Path $RunDir "ToolkitContext.json"
  if (Test-Path -LiteralPath $ctxPath) {
    try {
      $raw = Get-Content -LiteralPath $ctxPath -Raw -Encoding UTF8
      $j = $raw | ConvertFrom-Json
      if (-not $ctx.server -and $j.sqlServer) { $ctx.server = [string]$j.sqlServer }
      if (-not $ctx.dsn -and $j.sqlDsn)       { $ctx.dsn    = [string]$j.sqlDsn }
      if (-not $ctx.database -and $j.sqlDatabase) { $ctx.database = [string]$j.sqlDatabase }
      if (-not $ctx.registryPath -and $j.sqlRegPath) { $ctx.registryPath = [string]$j.sqlRegPath }
    } catch {
      # Non-fatal, but note it
      Write-Log WARN "Failed to parse ToolkitContext.json (continuing). $($_.Exception.Message)"
    }
  }

  # Choose provider:
  # - Prefer DSN via ODBC when provided
  # - Otherwise use SqlClient with server (+ optional database)
  if ($ctx.dsn) {
    $ctx.provider = "odbc"
  } elseif ($ctx.server) {
    $ctx.provider = "sqlclient"
  } else {
    $ctx.provider = $null
  }

  return $ctx
}

function New-DbConnection {
  param([hashtable]$SqlContext)

  if (-not $SqlContext.provider) {
    throw "No SQL context found. Ensure BEPOZ_SQL_DSN or BEPOZ_SQL_SERVER is provided by the Toolkit (env vars and/or ToolkitContext.json)."
  }

  if ($SqlContext.provider -eq "odbc") {
    Add-Type -AssemblyName System.Data
    $cs = "DSN=$($SqlContext.dsn);Trusted_Connection=Yes;"
    $conn = New-Object System.Data.Odbc.OdbcConnection($cs)
    $conn.Open()
    return [ordered]@{ Kind="odbc"; Connection=$conn }
  }

  if ($SqlContext.provider -eq "sqlclient") {
    Add-Type -AssemblyName System.Data
    $db = if ($SqlContext.database) { $SqlContext.database } else { "Bepoz" }
    $cs = "Server=$($SqlContext.server);Database=$db;Integrated Security=SSPI;Application Name=BepozToolkit-cardname-propagate;"
    $conn = New-Object System.Data.SqlClient.SqlConnection($cs)
    $conn.Open()
    return [ordered]@{ Kind="sqlclient"; Connection=$conn }
  }

  throw "Unsupported provider: $($SqlContext.provider)"
}

function Invoke-DbQuery {
  param(
    [hashtable]$Db,
    [string]$Sql,
    [object[]]$Params = @(),
    [object]$Transaction = $null
  )

  $conn = $Db.Connection

  if ($Db.Kind -eq "sqlclient") {
    $cmd = $conn.CreateCommand()
    if ($Transaction) { $cmd.Transaction = $Transaction }
    $cmd.CommandText = $Sql
    for ($i=0; $i -lt $Params.Count; $i+=2) {
      $name = [string]$Params[$i]
      $val  = $Params[$i+1]
      $p = $cmd.Parameters.Add($name, [System.Data.SqlDbType]::Variant)
      $p.Value = $val
    }
    $da = New-Object System.Data.SqlClient.SqlDataAdapter($cmd)
    $dt = New-Object System.Data.DataTable
    [void]$da.Fill($dt)
    return $dt
  }

  if ($Db.Kind -eq "odbc") {
    $cmd = $conn.CreateCommand()
    if ($Transaction) { $cmd.Transaction = $Transaction }
    $cmd.CommandText = $Sql

    # ODBC uses positional parameters (?)
    foreach ($val in $Params) {
      $p = $cmd.Parameters.Add("p", [System.Data.Odbc.OdbcType]::VarChar, 0)
      $p.Value = $val
    }

    $da = New-Object System.Data.Odbc.OdbcDataAdapter($cmd)
    $dt = New-Object System.Data.DataTable
    [void]$da.Fill($dt)
    return $dt
  }

  throw "Unknown DB kind: $($Db.Kind)"
}

function Invoke-DbNonQuery {
  param(
    [hashtable]$Db,
    [string]$Sql,
    [object[]]$Params = @(),
    [object]$Transaction = $null
  )

  $conn = $Db.Connection

  if ($Db.Kind -eq "sqlclient") {
    $cmd = $conn.CreateCommand()
    if ($Transaction) { $cmd.Transaction = $Transaction }
    $cmd.CommandText = $Sql
    for ($i=0; $i -lt $Params.Count; $i+=2) {
      $name = [string]$Params[$i]
      $val  = $Params[$i+1]
      $p = $cmd.Parameters.Add($name, [System.Data.SqlDbType]::Variant)
      $p.Value = $val
    }
    return $cmd.ExecuteNonQuery()
  }

  if ($Db.Kind -eq "odbc") {
    $cmd = $conn.CreateCommand()
    if ($Transaction) { $cmd.Transaction = $Transaction }
    $cmd.CommandText = $Sql

    foreach ($val in $Params) {
      $p = $cmd.Parameters.Add("p", [System.Data.Odbc.OdbcType]::VarChar, 0)
      $p.Value = $val
    }

    return $cmd.ExecuteNonQuery()
  }

  throw "Unknown DB kind: $($Db.Kind)"
}

function Confirm-User {
  param(
    [string]$Title,
    [string]$Message,
    [switch]$RequireForceIfNonInteractive
  )

  if ($Force) { return $true }

  $interactive = Test-IsInteractive

  if (-not $interactive) {
    if ($RequireForceIfNonInteractive) {
      throw "Non-interactive run: confirmation required. Re-run with -Force to proceed unattended."
    } else {
      return $false
    }
  }

  # Ensure WinForms works (STA). If not STA, relaunch in STA once.
  try {
    $apt = [System.Threading.Thread]::CurrentThread.ApartmentState
    if ($apt -ne [System.Threading.ApartmentState]::STA) {
      Write-Log INFO "Relaunching in STA for WinForms prompts..."
      $argList = @("-NoProfile","-ExecutionPolicy","Bypass","-STA","-File",$PSCommandPath)
      foreach ($k in $PSBoundParameters.Keys) {
        $v = $PSBoundParameters[$k]
        if ($v -is [switch] -and $v.IsPresent) {
          $argList += "-$k"
        } elseif (-not ($v -is [switch])) {
          $argList += "-$k"
          $argList += "$v"
        }
      }
      $p = Start-Process -FilePath "powershell.exe" -ArgumentList $argList -Wait -PassThru
      exit $p.ExitCode
    }
  } catch {
    # If STA logic fails, continue without relaunch and attempt MessageBox anyway.
  }

  Add-Type -AssemblyName System.Windows.Forms
  Add-Type -AssemblyName System.Drawing
  [System.Windows.Forms.Application]::EnableVisualStyles() | Out-Null

  $res = [System.Windows.Forms.MessageBox]::Show(
    $Message,
    $Title,
    [System.Windows.Forms.MessageBoxButtons]::YesNo,
    [System.Windows.Forms.MessageBoxIcon]::Warning
  )
  return ($res -eq [System.Windows.Forms.DialogResult]::Yes)
}

function Select-MasterVenueIdWinForms {
  param([System.Data.DataTable]$Venues)

  Add-Type -AssemblyName System.Windows.Forms
  Add-Type -AssemblyName System.Drawing
  [System.Windows.Forms.Application]::EnableVisualStyles() | Out-Null

  $form = New-Object System.Windows.Forms.Form
  $form.Text = "Select Master Venue"
  $form.StartPosition = "CenterScreen"
  $form.Size = New-Object System.Drawing.Size(520, 420)
  $form.TopMost = $true

  $label = New-Object System.Windows.Forms.Label
  $label.Text = "Select the master VenueID. CardName_1..16 will be copied from this venue to all others."
  $label.AutoSize = $false
  $label.Size = New-Object System.Drawing.Size(480, 40)
  $label.Location = New-Object System.Drawing.Point(10, 10)
  $form.Controls.Add($label)

  $list = New-Object System.Windows.Forms.ListView
  $list.View = 'Details'
  $list.FullRowSelect = $true
  $list.MultiSelect = $false
  $list.Size = New-Object System.Drawing.Size(480, 280)
  $list.Location = New-Object System.Drawing.Point(10, 60)
  [void]$list.Columns.Add("VenueID", 80)
  [void]$list.Columns.Add("Name", 360)

  foreach ($row in $Venues.Rows) {
    $item = New-Object System.Windows.Forms.ListViewItem([string]$row.VenueID)
    [void]$item.SubItems.Add([string]$row.Name)
    [void]$list.Items.Add($item)
  }
  $form.Controls.Add($list)

  $ok = New-Object System.Windows.Forms.Button
  $ok.Text = "OK"
  $ok.Size = New-Object System.Drawing.Size(100, 30)
  $ok.Location = New-Object System.Drawing.Point(290, 350)
  $ok.DialogResult = [System.Windows.Forms.DialogResult]::OK
  $form.AcceptButton = $ok
  $form.Controls.Add($ok)

  $cancel = New-Object System.Windows.Forms.Button
  $cancel.Text = "Cancel"
  $cancel.Size = New-Object System.Drawing.Size(100, 30)
  $cancel.Location = New-Object System.Drawing.Point(390, 350)
  $cancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
  $form.CancelButton = $cancel
  $form.Controls.Add($cancel)

  $result = $form.ShowDialog()

  if ($result -ne [System.Windows.Forms.DialogResult]::OK) {
    return $null
  }

  if ($list.SelectedItems.Count -ne 1) {
    return $null
  }

  return [int]$list.SelectedItems[0].Text
}

function Get-BepozProcesses {
  param([string[]]$Roots)

  $procs = @()
  try {
    $wmi = Get-CimInstance -ClassName Win32_Process -ErrorAction Stop
  } catch {
    $wmi = Get-WmiObject -Class Win32_Process
  }

  foreach ($p in $wmi) {
    $path = $null
    try { $path = $p.ExecutablePath } catch { $path = $null }
    if (-not $path) { continue }

    foreach ($r in $Roots) {
      if ($path.StartsWith($r, [System.StringComparison]::OrdinalIgnoreCase)) {
        $procs += [ordered]@{
          Name = $p.Name
          ProcessId = [int]$p.ProcessId
          ExecutablePath = $path
        }
        break
      }
    }
  }

  return $procs
}

function Close-BepozProcesses {
  param([hashtable[]]$ProcList)

  $closed = @()
  $failed = @()

  foreach ($p in $ProcList) {
    try {
      $gp = Get-Process -Id $p.ProcessId -ErrorAction Stop
      if ($gp.MainWindowHandle -ne 0) {
        [void]$gp.CloseMainWindow()
        Start-Sleep -Seconds 3
      }
      if (-not $gp.HasExited) {
        try {
          Stop-Process -Id $p.ProcessId -Force -ErrorAction Stop
        } catch {
          # ignore here; report as failed below
        }
      }
      $closed += $p
    } catch {
      $failed += [ordered]@{
        process = $p
        error = $_.Exception.Message
      }
    }
  }

  return [ordered]@{ closed = $closed; failed = $failed }
}

# ----------------------------
# Main
# ----------------------------
try {
  Ensure-RunDir
  $applyMode = [bool]$ApplyFix.IsPresent
  $auditMode = -not $applyMode

  if ($AuditOnly.IsPresent -and $ApplyFix.IsPresent) {
    Write-Log ERROR "Invalid flags: -AuditOnly and -ApplyFix cannot both be specified."
    exit 2
  }

  # Get SQL context from Toolkit
  $sqlCtx = Get-ToolkitSqlContext

  $report = New-ReportObject -SqlContext $sqlCtx -AuditMode $auditMode -ApplyMode $applyMode

  Write-Log INFO "Mode: AuditOnly=$auditMode ApplyFix=$applyMode Force=$($Force.IsPresent)"
  Write-Log INFO "RunDir: $RunDir"

  if (-not $sqlCtx.provider) {
    $report.success = $false
    $report.errors += "No SQL context found (BEPOZ_SQL_DSN/BEPOZ_SQL_SERVER)."
    Save-Report $report
    exit 2
  }

  # 1) Detect running Bepoz apps under requested roots
  $roots = @()
  if ($BepozProgramsRoot) { $roots += $BepozProgramsRoot }
  if ($BepozRoot)         { $roots += $BepozRoot }

  $bepozProcs = Get-BepozProcesses -Roots $roots
  if ($bepozProcs.Count -gt 0) {
    Add-Finding -Report $report -Severity "warning" -Title "Bepoz apps running" -Details "Detected running processes under Bepoz folders. Consider closing before applying DB changes." -Evidence @{ count=$bepozProcs.Count; processes=$bepozProcs }
    Write-Log WARN "Detected $($bepozProcs.Count) running Bepoz process(es)."
  } else {
    Add-Finding -Report $report -Severity "info" -Title "No Bepoz apps running" -Details "No running processes detected under the configured Bepoz folders." -Evidence @{ roots=$roots }
    Write-Log INFO "No running Bepoz processes detected under configured roots."
  }

  # 2) Connect to DB
  Write-Log INFO "Connecting to database using provider: $($sqlCtx.provider)"
  $db = New-DbConnection -SqlContext $sqlCtx
  try {
    # 3) Find venues list
    $venues = Invoke-DbQuery -Db $db -Sql "SELECT VenueID, Name FROM dbo.Venue ORDER BY VenueID"
    if ($venues.Rows.Count -lt 1) {
      throw "No venues returned from dbo.Venue."
    }

    # 4) Choose master venue
    if (-not $PSBoundParameters.ContainsKey("MasterVenueId")) {
      if (-not (Test-IsInteractive)) {
        $report.success = $false
        $report.errors += "MasterVenueId not provided and run is non-interactive. Re-run with -MasterVenueId <id>."
        Save-Report $report
        exit 2
      }

      # Relaunch in STA if needed for WinForms (handled in Confirm-User too, but do it here for selection)
      try {
        $apt = [System.Threading.Thread]::CurrentThread.ApartmentState
        if ($apt -ne [System.Threading.ApartmentState]::STA) {
          Write-Log INFO "Relaunching in STA for WinForms venue picker..."
          $argList = @("-NoProfile","-ExecutionPolicy","Bypass","-STA","-File",$PSCommandPath)
          foreach ($k in $PSBoundParameters.Keys) {
            $v = $PSBoundParameters[$k]
            if ($v -is [switch] -and $v.IsPresent) {
              $argList += "-$k"
            } elseif (-not ($v -is [switch])) {
              $argList += "-$k"
              $argList += "$v"
            }
          }
          $p = Start-Process -FilePath "powershell.exe" -ArgumentList $argList -Wait -PassThru
          exit $p.ExitCode
        }
      } catch { }

      $picked = Select-MasterVenueIdWinForms -Venues $venues
      if (-not $picked) {
        $report.success = $false
        $report.errors += "User cancelled master venue selection."
        Save-Report $report
        exit 3
      }
      $MasterVenueId = $picked
    }

    # Validate master exists
    $masterRows = @($venues.Select("VenueID = $MasterVenueId"))
    if ($masterRows.Count -ne 1) {
      $report.success = $false
      $report.errors += "MasterVenueId '$MasterVenueId' not found in dbo.Venue."
      Save-Report $report
      exit 2
    }

    $masterName = [string]$masterRows[0].Name
    Add-Finding -Report $report -Severity "info" -Title "Master venue selected" -Details "Using VenueID $MasterVenueId ($masterName) as the master source for CardName_1..16." -Evidence @{ masterVenueId=$MasterVenueId; masterVenueName=$masterName }
    Write-Log INFO "Master VenueID: $MasterVenueId ($masterName)"

    # 5) Locate CardName_1..16 columns
    $colSql = @"
SELECT TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME
FROM INFORMATION_SCHEMA.COLUMNS
WHERE LOWER(COLUMN_NAME) LIKE 'cardname[_]%'
ORDER BY TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME
"@
    $cols = Invoke-DbQuery -Db $db -Sql $colSql

    # Group tables and check if all 16 numbered columns exist (case-insensitive)
    $tableMap = @{}
    foreach ($r in $cols.Rows) {
      $key = "$($r.TABLE_SCHEMA).$($r.TABLE_NAME)"
      if (-not $tableMap.ContainsKey($key)) {
        $tableMap[$key] = @()
      }
      $tableMap[$key] += [string]$r.COLUMN_NAME
    }

    function Get-NumberedSet {
      param([string[]]$Names)
      $set = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
      foreach ($n in $Names) { [void]$set.Add($n) }
      return $set
    }

    $targetTable = $null
    $targetCols  = $null

    # Prefer dbo.Venue if it matches
    foreach ($k in $tableMap.Keys) {
      $set = Get-NumberedSet -Names $tableMap[$k]
      $all = $true
      $resolved = @()
      for ($i=1; $i -le 16; $i++) {
        $c1 = "CardName_$i"
        $c2 = "cardname_$i"
        if ($set.Contains($c1)) { $resolved += $c1 }
        elseif ($set.Contains($c2)) { $resolved += $c2 }
        else { $all = $false; break }
      }
      if ($all -and ($k -ieq "dbo.Venue")) {
        $targetTable = $k
        $targetCols  = $resolved
        break
      }
    }

    # Otherwise pick the first table that has all 16 columns
    if (-not $targetTable) {
      foreach ($k in $tableMap.Keys) {
        $set = Get-NumberedSet -Names $tableMap[$k]
        $all = $true
        $resolved = @()
        for ($i=1; $i -le 16; $i++) {
          $c1 = "CardName_$i"
          $c2 = "cardname_$i"
          if ($set.Contains($c1)) { $resolved += $c1 }
          elseif ($set.Contains($c2)) { $resolved += $c2 }
          else { $all = $false; break }
        }
        if ($all) {
          $targetTable = $k
          $targetCols  = $resolved
          break
        }
      }
    }

    if (-not $targetTable) {
      $report.success = $false
      $report.errors += "Could not find a table containing CardName_1..16 (case-insensitive)."
      Add-Finding -Report $report -Severity "critical" -Title "CardName columns not found" -Details "Searched INFORMATION_SCHEMA.COLUMNS for cardname_* but did not find a table containing all 16 numbered columns." -Evidence @{ tablesChecked=$tableMap.Keys }
      Save-Report $report
      exit 2
    }

    Add-Finding -Report $report -Severity "info" -Title "CardName source table identified" -Details "Using $targetTable as the source/target for CardName_1..16 replication." -Evidence @{ table=$targetTable; columns=$targetCols }
    Write-Log INFO "Using CardName table: $targetTable"

    # 6) Fetch master values
    $schema, $table = $targetTable.Split(".")
    $selectCols = ($targetCols | ForEach-Object { "[$_]" }) -join ", "
    $sqlGetMaster = "SELECT $selectCols FROM [$schema].[$table] WHERE VenueID = " + [int]$MasterVenueId

    $masterDt = Invoke-DbQuery -Db $db -Sql $sqlGetMaster
    if ($masterDt.Rows.Count -ne 1) {
      $report.success = $false
      $report.errors += "Master venue row not found in $targetTable for VenueID $MasterVenueId."
      Save-Report $report
      exit 2
    }

    $masterVals = [ordered]@{}
    foreach ($c in $targetCols) {
      $masterVals[$c] = [string]$masterDt.Rows[0].Item($c)
    }

    # 7) Audit differences across venues
    $sqlAll = "SELECT VenueID, $selectCols FROM [$schema].[$table] ORDER BY VenueID"
    $allDt = Invoke-DbQuery -Db $db -Sql $sqlAll

    $diffs = @()
    foreach ($row in $allDt.Rows) {
      $vid = [int]$row.VenueID
      if ($vid -eq $MasterVenueId) { continue }

      $different = $false
      foreach ($c in $targetCols) {
        $v = [string]$row.Item($c)
        if ($v -ne [string]$masterVals[$c]) { $different = $true; break }
      }

      if ($different) {
        $diffs += $vid
      }
    }

    Add-Finding -Report $report -Severity "info" -Title "CardName audit summary" -Details "Compared CardName_1..16 across venues against master VenueID $MasterVenueId." -Evidence @{ masterVenueId=$MasterVenueId; totalVenues=$allDt.Rows.Count; venuesDifferent=$diffs; differentCount=$diffs.Count }

    if ($diffs.Count -gt 0) {
      Write-Log WARN "$($diffs.Count) venue(s) differ from the master and would be updated in ApplyFix mode."
    } else {
      Write-Log INFO "All venues already match the master CardName_1..16."
    }

    # 8) Apply changes if requested
    if ($applyMode) {

      # Optionally close apps first
      if ($CloseBepozApps -and $bepozProcs.Count -gt 0) {
        $okClose = Confirm-User -Title "Close Bepoz applications?" -Message "About to close $($bepozProcs.Count) running Bepoz process(es) under $($roots -join ', '). Continue?" -RequireForceIfNonInteractive
        if (-not $okClose) {
          $report.success = $false
          $report.errors += "User declined closing Bepoz applications."
          Save-Report $report
          exit 3
        }

        Write-Log INFO "Closing Bepoz processes..."
        $closeRes = Close-BepozProcesses -ProcList $bepozProcs
        Add-Action -Report $report -Title "Close Bepoz applications" -Details "Attempted to close running processes under Bepoz folders." -Result (if ($closeRes.failed.Count -eq 0) { "success" } else { "failed" })

        if ($closeRes.failed.Count -gt 0) {
          $report.warnings += "Some processes could not be closed."
          Add-Finding -Report $report -Severity "warning" -Title "Some Bepoz processes failed to close" -Details "Not all detected processes could be closed/killed." -Evidence $closeRes
          Write-Log WARN "Some processes failed to close. Continuing (DB update may still succeed)."
        } else {
          Write-Log INFO "Processes closed."
        }
      }

      $targets = $allDt.Rows.Count - 1
      $prompt = "This will overwrite CardName_1..16 for $targets venue(s) with values from master VenueID $MasterVenueId ($masterName). Continue?"
      $ok = Confirm-User -Title "Apply CardName propagation?" -Message $prompt -RequireForceIfNonInteractive
      if (-not $ok) {
        $report.success = $false
        $report.errors += "User declined CardName propagation."
        Save-Report $report
        exit 3
      }

      Write-Log INFO "Applying CardName propagation in a transaction..."

      $tran = $db.Connection.BeginTransaction()
      try {
        if ($db.Kind -eq "sqlclient") {
          $setParts = @()
          $paramPairs = @()
          $i = 1
          foreach ($c in $targetCols) {
            $pn = "@c$i"
            $setParts += "[$c] = $pn"
            $paramPairs += $pn
            $paramPairs += $masterVals[$c]
            $i++
          }
          $sqlUpd = "UPDATE [$schema].[$table] SET " + ($setParts -join ", ") + " WHERE VenueID <> @master"
          $paramPairs += "@master"
          $paramPairs += [int]$MasterVenueId

          $rows = Invoke-DbNonQuery -Db $db -Sql $sqlUpd -Params $paramPairs -Transaction $tran
        } else {
          # ODBC positional parameters
          $setParts = @()
          $vals = @()
          foreach ($c in $targetCols) {
            $setParts += "[$c] = ?"
            $vals += $masterVals[$c]
          }
          $sqlUpd = "UPDATE [$schema].[$table] SET " + ($setParts -join ", ") + " WHERE VenueID <> ?"
          $vals += [int]$MasterVenueId

          $rows = Invoke-DbNonQuery -Db $db -Sql $sqlUpd -Params $vals -Transaction $tran
        }

        $tran.Commit()
        Add-Action -Report $report -Title "Propagate CardName_1..16" -Details "Updated venues to match master VenueID $MasterVenueId ($masterName). Rows affected: $rows" -Result "success"
        Write-Log INFO "Propagation complete. Rows affected: $rows"
      } catch {
        try { $tran.Rollback() } catch { }
        Add-Action -Report $report -Title "Propagate CardName_1..16" -Details "Transaction failed and was rolled back. $($_.Exception.Message)" -Result "failed"
        throw
      }
    }

    Save-Report $report
    exit 0
  }
  finally {
    if ($db -and $db.Connection) {
      try { $db.Connection.Close() } catch { }
      try { $db.Connection.Dispose() } catch { }
    }
  }
}
catch {
  try {
    Write-Log ERROR $_.Exception.Message
  } catch { }

  try {
    if (-not $report) {
      $sqlCtx2 = Get-ToolkitSqlContext
      $report = New-ReportObject -SqlContext $sqlCtx2 -AuditMode $true -ApplyMode $false
    }
    $report.success = $false
    $report.errors += $_.Exception.Message
    Save-Report $report
  } catch { }

  exit 1
}
