#requires -version 5.1
<#
  Bepoz Onboarding Toolkit - Create POS Terminals
  - Creates POS terminals by cloning an existing dbo.Workstation row ("template") into a selected Venue + Store.
  - Safe by default: DryRun = true unless explicitly set to false AND user confirms.
  - Writes Report.json into RunDir.

  Toolkit passes -RunDir automatically. This script MUST accept it.

  Optional config file:
    create-pos-terminals.config.json placed in RunDir, or pass -ConfigPath.
#>

[CmdletBinding()]
param(
  [Parameter(Mandatory = $false)]
  [string]$RunDir,

  [Parameter(Mandatory = $false)]
  [string]$ConfigPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$ToolId = "create-pos-terminals"
$ToolVersion = "1.0.1"

function Resolve-RunDir {
  param([string]$ProvidedRunDir)

  if ($ProvidedRunDir -and $ProvidedRunDir.Trim() -ne '') {
    if (-not (Test-Path -LiteralPath $ProvidedRunDir)) {
      throw "RunDir path does not exist: $ProvidedRunDir"
    }
    return (Resolve-Path -LiteralPath $ProvidedRunDir).Path
  }

  if ($env:BEPoz_TOOLKIT_RUNDIR -and (Test-Path -LiteralPath $env:BEPoz_TOOLKIT_RUNDIR)) {
    return (Resolve-Path -LiteralPath $env:BEPoz_TOOLKIT_RUNDIR).Path
  }

  return $PSScriptRoot
}

$RunDir = Resolve-RunDir -ProvidedRunDir $RunDir
$env:BEPoz_TOOLKIT_RUNDIR = $RunDir

$ReportPath = Join-Path $RunDir 'Report.json'

# ---- Logging helpers ----
$script:LogLines = New-Object System.Collections.Generic.List[string]
$script:Errors   = New-Object System.Collections.Generic.List[object]

function Write-Log {
  param(
    [Parameter(Mandatory=$true)][string]$Message,
    [ValidateSet('INFO','WARN','ERROR')][string]$Level = 'INFO'
  )
  $line = "[{0}] [{1}] {2}" -f (Get-Date).ToString("yyyy-MM-dd HH:mm:ss"), $Level, $Message
  Write-Output $line
  [void]$script:LogLines.Add($line)
}

function Add-ErrorRecord {
  param([string]$Message, [System.Exception]$Exception)
  $errObj = [pscustomobject]@{
    message          = $Message
    exceptionType    = if ($Exception) { $Exception.GetType().FullName } else { $null }
    exceptionMessage = if ($Exception) { $Exception.Message } else { $null }
    stack            = if ($Exception) { $Exception.StackTrace } else { $null }
  }
  [void]$script:Errors.Add($errObj)
  Write-Log -Level 'ERROR' -Message $Message
  if ($Exception) { Write-Log -Level 'ERROR' -Message ("Exception: " + $Exception.Message) }
}

function Save-Report {
  param([Parameter(Mandatory=$true)][hashtable]$Report)
  try {
    $json = $Report | ConvertTo-Json -Depth 12
    $json | Set-Content -LiteralPath $ReportPath -Encoding UTF8
    Write-Log "Wrote report: $ReportPath"
  } catch {
    Write-Log -Level 'ERROR' -Message ("Failed to write report: " + $_.Exception.Message)
  }
}

function Test-Interactive {
  try {
    return [Environment]::UserInteractive
  } catch {
    return $false
  }
}

# ---- WinForms helpers ----
function Ensure-WinForms {
  Add-Type -AssemblyName System.Windows.Forms | Out-Null
  Add-Type -AssemblyName System.Drawing | Out-Null
  [System.Windows.Forms.Application]::EnableVisualStyles()
}

function Show-TextInputDialog {
  param([string]$Title,[string]$Prompt,[string]$DefaultValue = "")
  Ensure-WinForms

  $form = New-Object System.Windows.Forms.Form
  $form.Text = $Title
  $form.StartPosition = 'CenterScreen'
  $form.Width = 520
  $form.Height = 170
  $form.TopMost = $true

  $lbl = New-Object System.Windows.Forms.Label
  $lbl.AutoSize = $true
  $lbl.Left = 12
  $lbl.Top = 12
  $lbl.Text = $Prompt

  $txt = New-Object System.Windows.Forms.TextBox
  $txt.Left = 12
  $txt.Top = 40
  $txt.Width = 480
  $txt.Text = $DefaultValue

  $ok = New-Object System.Windows.Forms.Button
  $ok.Text = "OK"
  $ok.Left = 312
  $ok.Top = 80
  $ok.Width = 80
  $ok.DialogResult = [System.Windows.Forms.DialogResult]::OK

  $cancel = New-Object System.Windows.Forms.Button
  $cancel.Text = "Cancel"
  $cancel.Left = 412
  $cancel.Top = 80
  $cancel.Width = 80
  $cancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel

  $form.Controls.AddRange(@($lbl,$txt,$ok,$cancel))
  $form.AcceptButton = $ok
  $form.CancelButton = $cancel

  $result = $form.ShowDialog()
  if ($result -ne [System.Windows.Forms.DialogResult]::OK) { return $null }
  return $txt.Text
}

function Show-NumberInputDialog {
  param([string]$Title,[string]$Prompt,[int]$DefaultValue = 1,[int]$Min = 1,[int]$Max = 500)
  Ensure-WinForms

  $form = New-Object System.Windows.Forms.Form
  $form.Text = $Title
  $form.StartPosition = 'CenterScreen'
  $form.Width = 520
  $form.Height = 190
  $form.TopMost = $true

  $lbl = New-Object System.Windows.Forms.Label
  $lbl.AutoSize = $true
  $lbl.Left = 12
  $lbl.Top = 12
  $lbl.Text = $Prompt

  $num = New-Object System.Windows.Forms.NumericUpDown
  $num.Left = 12
  $num.Top = 40
  $num.Width = 200
  $num.Minimum = $Min
  $num.Maximum = $Max
  $num.Value = [Math]::Min([Math]::Max($DefaultValue, $Min), $Max)

  $ok = New-Object System.Windows.Forms.Button
  $ok.Text = "OK"
  $ok.Left = 312
  $ok.Top = 90
  $ok.Width = 80
  $ok.DialogResult = [System.Windows.Forms.DialogResult]::OK

  $cancel = New-Object System.Windows.Forms.Button
  $cancel.Text = "Cancel"
  $cancel.Left = 412
  $cancel.Top = 90
  $cancel.Width = 80
  $cancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel

  $form.Controls.AddRange(@($lbl,$num,$ok,$cancel))
  $form.AcceptButton = $ok
  $form.CancelButton = $cancel

  $result = $form.ShowDialog()
  if ($result -ne [System.Windows.Forms.DialogResult]::OK) { return $null }
  return [int]$num.Value
}

function Show-SelectDialog {
  param([string]$Title,[string]$Prompt,[Parameter(Mandatory=$true)][System.Collections.IEnumerable]$Items,[string]$DisplayMember = "Name")
  Ensure-WinForms

  $form = New-Object System.Windows.Forms.Form
  $form.Text = $Title
  $form.StartPosition = 'CenterScreen'
  $form.Width = 720
  $form.Height = 420
  $form.TopMost = $true

  $lbl = New-Object System.Windows.Forms.Label
  $lbl.AutoSize = $true
  $lbl.Left = 12
  $lbl.Top = 12
  $lbl.Text = $Prompt

  $list = New-Object System.Windows.Forms.ListBox
  $list.Left = 12
  $list.Top = 40
  $list.Width = 680
  $list.Height = 290
  $list.DisplayMember = $DisplayMember

  foreach ($i in $Items) { [void]$list.Items.Add($i) }
  if ($list.Items.Count -gt 0) { $list.SelectedIndex = 0 }

  $ok = New-Object System.Windows.Forms.Button
  $ok.Text = "OK"
  $ok.Left = 512
  $ok.Top = 340
  $ok.Width = 80
  $ok.DialogResult = [System.Windows.Forms.DialogResult]::OK

  $cancel = New-Object System.Windows.Forms.Button
  $cancel.Text = "Cancel"
  $cancel.Left = 612
  $cancel.Top = 340
  $cancel.Width = 80
  $cancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel

  $form.Controls.AddRange(@($lbl,$list,$ok,$cancel))
  $form.AcceptButton = $ok
  $form.CancelButton = $cancel

  $result = $form.ShowDialog()
  if ($result -ne [System.Windows.Forms.DialogResult]::OK) { return $null }
  return $list.SelectedItem
}

function Show-YesNo {
  param([string]$Title, [string]$Message)
  Ensure-WinForms
  $res = [System.Windows.Forms.MessageBox]::Show(
    $Message, $Title,
    [System.Windows.Forms.MessageBoxButtons]::YesNo,
    [System.Windows.Forms.MessageBoxIcon]::Question
  )
  return ($res -eq [System.Windows.Forms.DialogResult]::Yes)
}

# ---- SQL helpers ----
function New-SqlConnection {
  param([Parameter(Mandatory=$true)][string]$Server,[Parameter(Mandatory=$true)][string]$Database,[int]$TimeoutSeconds = 8)
  Add-Type -AssemblyName System.Data | Out-Null
  $cs = "Server=$Server;Database=$Database;Integrated Security=True;Connect Timeout=$TimeoutSeconds;Application Name=BepozToolkit-$ToolId"
  $conn = New-Object System.Data.SqlClient.SqlConnection $cs
  $conn.Open()
  return $conn
}

function Invoke-SqlQuery {
  param([Parameter(Mandatory=$true)][System.Data.SqlClient.SqlConnection]$Connection,[Parameter(Mandatory=$true)][string]$Sql,[hashtable]$Parameters)
  $cmd = $Connection.CreateCommand()
  $cmd.CommandText = $Sql
  $cmd.CommandTimeout = 30

  if ($Parameters) {
    foreach ($k in $Parameters.Keys) {
      $p = $cmd.Parameters.AddWithValue($k, $Parameters[$k])
      if ($null -eq $Parameters[$k]) { $p.Value = [DBNull]::Value }
    }
  }

  $da = New-Object System.Data.SqlClient.SqlDataAdapter $cmd
  $dt = New-Object System.Data.DataTable
  [void]$da.Fill($dt)
  return $dt
}

function Invoke-SqlScalar {
  param([Parameter(Mandatory=$true)][System.Data.SqlClient.SqlConnection]$Connection,[Parameter(Mandatory=$true)][string]$Sql,[hashtable]$Parameters)
  $cmd = $Connection.CreateCommand()
  $cmd.CommandText = $Sql
  $cmd.CommandTimeout = 30
  if ($Parameters) {
    foreach ($k in $Parameters.Keys) {
      $p = $cmd.Parameters.AddWithValue($k, $Parameters[$k])
      if ($null -eq $Parameters[$k]) { $p.Value = [DBNull]::Value }
    }
  }
  return $cmd.ExecuteScalar()
}

function Get-DefaultSqlServerGuess {
  try {
    $svc = Get-Service -ErrorAction SilentlyContinue | Where-Object { $_.Name -eq 'MSSQLSERVER' -and $_.Status -eq 'Running' }
    if ($svc) { return 'localhost' }

    $named = Get-Service -ErrorAction SilentlyContinue | Where-Object { $_.Name -like 'MSSQL$*' -and $_.Status -eq 'Running' } | Select-Object -First 1
    if ($named) {
      $inst = $named.Name.Substring(6)
      return "localhost\$inst"
    }
  } catch { }
  return 'localhost'
}

function Find-BepozCandidateDatabases {
  param([System.Data.SqlClient.SqlConnection]$MasterConnection)

  $dbs = Invoke-SqlQuery -Connection $MasterConnection -Sql @"
SELECT name
FROM sys.databases
WHERE database_id > 4
  AND state_desc = 'ONLINE'
ORDER BY name;
"@

  $candidates = New-Object System.Collections.Generic.List[object]
  foreach ($row in $dbs.Rows) {
    $dbName = [string]$row['name']
    try {
      $sql = @"
SELECT
  CASE WHEN
    EXISTS (SELECT 1 FROM [$dbName].sys.tables t JOIN [$dbName].sys.schemas s ON t.schema_id=s.schema_id WHERE s.name='dbo' AND t.name='Venue')
    AND EXISTS (SELECT 1 FROM [$dbName].sys.tables t JOIN [$dbName].sys.schemas s ON t.schema_id=s.schema_id WHERE s.name='dbo' AND t.name='Store')
    AND EXISTS (SELECT 1 FROM [$dbName].sys.tables t JOIN [$dbName].sys.schemas s ON t.schema_id=s.schema_id WHERE s.name='dbo' AND t.name='Workstation')
  THEN 1 ELSE 0 END AS IsCandidate;
"@
      $isCandidate = Invoke-SqlScalar -Connection $MasterConnection -Sql $sql
      if ([int]$isCandidate -eq 1) {
        [void]$candidates.Add([pscustomobject]@{ Name = $dbName })
      }
    } catch { }
  }

  return $candidates
}

function Get-InsertableWorkstationColumns {
  param([System.Data.SqlClient.SqlConnection]$Connection)

  $dt = Invoke-SqlQuery -Connection $Connection -Sql @"
SELECT c.name AS ColumnName, t.name AS TypeName, c.is_identity, c.is_computed, t.system_type_id
FROM sys.columns c
JOIN sys.types t ON c.user_type_id = t.user_type_id
WHERE c.object_id = OBJECT_ID('dbo.Workstation')
ORDER BY c.column_id;
"@

  $cols = @()
  foreach ($r in $dt.Rows) {
    $name = [string]$r['ColumnName']
    $isIdentity = [bool]$r['is_identity']
    $isComputed = [bool]$r['is_computed']
    $systemTypeId = [int]$r['system_type_id']

    if (-not $isIdentity -and -not $isComputed -and $systemTypeId -ne 189) {
      $cols += $name
    }
  }
  return $cols
}

function New-WorkstationInsertCommand {
  param([System.Data.SqlClient.SqlConnection]$Connection,[string[]]$Columns)

  $colList   = ($Columns | ForEach-Object { '[' + $_ + ']' }) -join ', '
  $paramList = ($Columns | ForEach-Object { '@' + $_ }) -join ', '
  $sql = "INSERT INTO dbo.Workstation ($colList) OUTPUT INSERTED.WorkstationID VALUES ($paramList);"

  $cmd = $Connection.CreateCommand()
  $cmd.CommandText = $sql
  $cmd.CommandTimeout = 30

  foreach ($c in $Columns) {
    [void]$cmd.Parameters.Add('@' + $c, [System.Data.SqlDbType]::Variant)
  }
  return $cmd
}

function Set-CommandParamValue {
  param([System.Data.SqlClient.SqlCommand]$Command,[string]$ParamName,$Value)
  if ($null -eq $Value) { $Command.Parameters[$ParamName].Value = [DBNull]::Value }
  else { $Command.Parameters[$ParamName].Value = $Value }
}

# ---- Main ----
$startUtc = (Get-Date).ToUniversalTime()
$report = @{
  toolId        = $ToolId
  toolVersion   = $ToolVersion
  startedAtUtc  = $startUtc.ToString("o")
  endedAtUtc    = $null
  runDir        = $RunDir
  machine       = $env:COMPUTERNAME
  user          = $env:USERNAME
  interactive   = (Test-Interactive)
  configSource  = $null
  sql           = @{ server = $null; database = $null }
  selection     = @{ venueId = $null; venueName = $null; storeId = $null; storeName = $null; templateWorkstationId = $null; templateWorkstationName = $null }
  plan          = @{ count = $null; namePrefix = $null; startNumber = $null; padding = $null; copyExportCodes = $false }
  dryRun        = $true
  confirmed     = $false
  results       = @()
  logs          = $null
  errors        = @()
  success       = $false
}

try {
  Write-Log "Tool starting. RunDir: $RunDir"

  $interactive = [bool]$report.interactive

  if (-not $ConfigPath -or $ConfigPath.Trim() -eq '') {
    $ConfigPath = Join-Path $RunDir 'create-pos-terminals.config.json'
  }

  $cfg = $null
  if (Test-Path -LiteralPath $ConfigPath) {
    Write-Log "Loading config: $ConfigPath"
    $cfg = (Get-Content -LiteralPath $ConfigPath -Raw -Encoding UTF8) | ConvertFrom-Json
    $report.configSource = "file"
  } else {
    $report.configSource = "gui"
  }

  # Inputs
  $sqlServer = $null
  $dbName = $null
  $venueId = $null
  $storeId = $null
  $templateWsId = $null
  $count = $null
  $prefix = $null
  $startNumber = 1
  $padding = 2
  $dryRun = $true
  $confirmFlag = $false
  $copyExportCodes = $false

  if ($cfg) {
    $sqlServer = [string]$cfg.SqlServer
    $dbName = [string]$cfg.Database
    $venueId = [int]$cfg.VenueId
    $storeId = [int]$cfg.StoreId
    $templateWsId = [int]$cfg.TemplateWorkstationId
    $count = [int]$cfg.Count
    $prefix = [string]$cfg.NamePrefix
    if ($cfg.PSObject.Properties.Name -contains 'StartNumber')     { $startNumber = [int]$cfg.StartNumber }
    if ($cfg.PSObject.Properties.Name -contains 'Padding')         { $padding = [int]$cfg.Padding }
    if ($cfg.PSObject.Properties.Name -contains 'DryRun')          { $dryRun = [bool]$cfg.DryRun }
    if ($cfg.PSObject.Properties.Name -contains 'Confirm')         { $confirmFlag = [bool]$cfg.Confirm }
    if ($cfg.PSObject.Properties.Name -contains 'CopyExportCodes') { $copyExportCodes = [bool]$cfg.CopyExportCodes }
  } else {
    if (-not $interactive) {
      throw "Non-interactive run detected, but no config file found. Create '$ConfigPath' in the RunDir and rerun."
    }

    $guess = Get-DefaultSqlServerGuess
    $sqlServer = Show-TextInputDialog -Title "SQL Server" -Prompt "Enter SQL Server instance (Windows auth):" -DefaultValue $guess
    if (-not $sqlServer) { throw "Cancelled by user." }

    Write-Log "Connecting to SQL Server (master): $sqlServer"
    $masterConn = New-SqlConnection -Server $sqlServer -Database 'master'
    try {
      $candidates = Find-BepozCandidateDatabases -MasterConnection $masterConn
      if (-not $candidates -or $candidates.Count -eq 0) {
        throw "No candidate databases found containing dbo.Venue, dbo.Store, and dbo.Workstation."
      }
    } finally {
      $masterConn.Dispose()
    }

    $pickedDb = Show-SelectDialog -Title "Database" -Prompt "Select the Bepoz venue database:" -Items $candidates -DisplayMember "Name"
    if (-not $pickedDb) { throw "Cancelled by user." }
    $dbName = [string]$pickedDb.Name

    $count = Show-NumberInputDialog -Title "How many POS terminals?" -Prompt "Enter the number of POS terminals to create:" -DefaultValue 1 -Min 1 -Max 200
    if (-not $count) { throw "Cancelled by user." }

    $prefix = Show-TextInputDialog -Title "POS naming" -Prompt "Name prefix (e.g. POS):" -DefaultValue "POS"
    if (-not $prefix) { throw "Cancelled by user." }

    $startNumber = Show-NumberInputDialog -Title "Starting number" -Prompt "Starting number (e.g. 1):" -DefaultValue 1 -Min 1 -Max 9999
    if (-not $startNumber) { throw "Cancelled by user." }

    $padding = Show-NumberInputDialog -Title "Number padding" -Prompt "Digits to pad (e.g. 2 => POS 01):" -DefaultValue 2 -Min 0 -Max 6
    if ($null -eq $padding) { throw "Cancelled by user." }

    $dryRun = Show-YesNo -Title "Dry Run?" -Message "Run in DRY RUN mode first?`n`nYes = plan only (recommended)`nNo = will insert after confirmation"

    $copyExportCodes = Show-YesNo -Title "Export codes" -Message "Copy ExportCode fields from the template workstation?`n`nNo is safer to avoid duplicates."
  }

  if (-not $sqlServer -or -not $dbName) { throw "Missing SQL server/database selection." }
  if (-not $prefix -or $prefix.Trim() -eq '') { throw "Missing NamePrefix." }
  if ($count -le 0) { throw "Count must be >= 1." }

  $report.sql.server = $sqlServer
  $report.sql.database = $dbName
  $report.plan.count = $count
  $report.plan.namePrefix = $prefix
  $report.plan.startNumber = $startNumber
  $report.plan.padding = $padding
  $report.plan.copyExportCodes = $copyExportCodes
  $report.dryRun = $dryRun

  Write-Log "Connecting to database: $sqlServer / $dbName"
  $conn = New-SqlConnection -Server $sqlServer -Database $dbName

  try {
    if (-not $venueId) {
      if (-not $interactive) { throw "VenueId is required in config for non-interactive runs." }
      $venuesDt = Invoke-SqlQuery -Connection $conn -Sql "SELECT VenueID, Name FROM dbo.Venue ORDER BY VenueID;"
      $venues = foreach ($r in $venuesDt.Rows) { [pscustomobject]@{ VenueID = [int]$r['VenueID']; Name = [string]$r['Name'] } }
      $pickedVenue = Show-SelectDialog -Title "Venue" -Prompt "Select the venue:" -Items $venues -DisplayMember "Name"
      if (-not $pickedVenue) { throw "Cancelled by user." }
      $venueId = [int]$pickedVenue.VenueID
    }

    $venueName = [string](Invoke-SqlScalar -Connection $conn -Sql "SELECT Name FROM dbo.Venue WHERE VenueID=@v" -Parameters @{ "@v" = $venueId })
    if (-not $venueName) { throw "VenueID $venueId not found." }

    if (-not $storeId) {
      if (-not $interactive) { throw "StoreId is required in config for non-interactive runs." }
      $storesDt = Invoke-SqlQuery -Connection $conn -Sql "SELECT StoreID, Name FROM dbo.Store WHERE VenueID=@v ORDER BY StoreID;" -Parameters @{ "@v" = $venueId }
      $stores = foreach ($r in $storesDt.Rows) { [pscustomobject]@{ StoreID = [int]$r['StoreID']; Name = [string]$r['Name'] } }
      if (-not $stores -or $stores.Count -eq 0) { throw "No stores found for VenueID $venueId." }
      $pickedStore = Show-SelectDialog -Title "Store" -Prompt "Select the store (within $venueName):" -Items $stores -DisplayMember "Name"
      if (-not $pickedStore) { throw "Cancelled by user." }
      $storeId = [int]$pickedStore.StoreID
    }

    $storeName = [string](Invoke-SqlScalar -Connection $conn -Sql "SELECT Name FROM dbo.Store WHERE StoreID=@s" -Parameters @{ "@s" = $storeId })
    if (-not $storeName) { throw "StoreID $storeId not found." }

    if (-not $templateWsId) {
      if (-not $interactive) { throw "TemplateWorkstationId is required in config for non-interactive runs." }
      $wsDt = Invoke-SqlQuery -Connection $conn -Sql "SELECT WorkstationID, Name FROM dbo.Workstation WHERE StoreID=@s ORDER BY WorkstationID;" -Parameters @{ "@s" = $storeId }
      $workstations = foreach ($r in $wsDt.Rows) { [pscustomobject]@{ WorkstationID = [int]$r['WorkstationID']; Name = [string]$r['Name'] } }
      if (-not $workstations -or $workstations.Count -eq 0) {
        throw "No workstations exist in StoreID $storeId. Create one manually first, then rerun to use it as a template."
      }
      $pickedWs = Show-SelectDialog -Title "Template Workstation" -Prompt "Select a template workstation to clone:" -Items $workstations -DisplayMember "Name"
      if (-not $pickedWs) { throw "Cancelled by user." }
      $templateWsId = [int]$pickedWs.WorkstationID
    }

    $templateWsName = [string](Invoke-SqlScalar -Connection $conn -Sql "SELECT Name FROM dbo.Workstation WHERE WorkstationID=@w" -Parameters @{ "@w" = $templateWsId })
    if (-not $templateWsName) { throw "Template WorkstationID $templateWsId not found." }

    $report.selection.venueId = $venueId
    $report.selection.venueName = $venueName
    $report.selection.storeId = $storeId
    $report.selection.storeName = $storeName
    $report.selection.templateWorkstationId = $templateWsId
    $report.selection.templateWorkstationName = $templateWsName

    Write-Log "Selected: Venue $venueId ($venueName) | Store $storeId ($storeName) | Template WS $templateWsId ($templateWsName)"
    Write-Log ("Plan: create {0} workstation(s) with prefix '{1}', start {2}, padding {3}. DryRun={4}" -f $count, $prefix, $startNumber, $padding, $dryRun)

    $tmplDt = Invoke-SqlQuery -Connection $conn -Sql "SELECT TOP 1 * FROM dbo.Workstation WHERE WorkstationID=@w;" -Parameters @{ "@w" = $templateWsId }
    if ($tmplDt.Rows.Count -ne 1) { throw "Failed to load template workstation row." }
    $tmplRow = $tmplDt.Rows[0]

    $cols = Get-InsertableWorkstationColumns -Connection $conn
    if (-not $cols -or $cols.Count -eq 0) { throw "Could not determine insertable columns for dbo.Workstation." }

    $existingDt = Invoke-SqlQuery -Connection $conn -Sql "SELECT Name FROM dbo.Workstation WHERE StoreID=@s;" -Parameters @{ "@s" = $storeId }
    $existingNames = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($r in $existingDt.Rows) { [void]$existingNames.Add([string]$r['Name']) }

    function Format-PosName([string]$Prefix, [int]$Number, [int]$Pad) {
      $n = if ($Pad -gt 0) { $Number.ToString().PadLeft($Pad, '0') } else { $Number.ToString() }
      return ("{0} {1}" -f $Prefix.Trim(), $n).Trim()
    }

    $namesToCreate = New-Object System.Collections.Generic.List[object]
    $n = $startNumber
    while ($namesToCreate.Count -lt $count) {
      $candidate = Format-PosName -Prefix $prefix -Number $n -Pad $padding
      if (-not $existingNames.Contains($candidate)) {
        [void]$namesToCreate.Add([pscustomobject]@{ Name = $candidate; Number = $n })
        [void]$existingNames.Add($candidate)
      }
      $n++
      if ($n -gt 999999) { throw "Name generation exceeded safe range. Check your prefix/padding/start number." }
    }

    Write-Log "Will create: $($namesToCreate.Count) workstation(s)."
    foreach ($item in $namesToCreate) { Write-Log ("  - {0}" -f $item.Name) }

    if (-not $dryRun) {
      if ($interactive) {
        $msg = "About to INSERT $count POS terminal(s) into:`n`nDatabase: $dbName`nVenue: $venueName (ID $venueId)`nStore: $storeName (ID $storeId)`nTemplate: $templateWsName (ID $templateWsId)`n`nProceed?"
        if (-not (Show-YesNo -Title "Confirm create POS terminals" -Message $msg)) {
          Write-Log -Level 'WARN' -Message "User declined. No changes made."
          $report.confirmed = $false
          $report.success = $true
          return
        }
        $report.confirmed = $true
      } else {
        if (-not $confirmFlag) {
          throw "Refusing to make changes in non-interactive mode. Set `"Confirm`": true in the config to proceed."
        }
        $report.confirmed = $true
      }
    }

    $results = New-Object System.Collections.Generic.List[object]

    if ($dryRun) {
      foreach ($item in $namesToCreate) {
        [void]$results.Add([pscustomobject]@{ name = $item.Name; workstationId = $null; status = "Planned"; error = $null })
      }
      Write-Log "Dry run complete (no inserts executed)."
    } else {
      $tx = $conn.BeginTransaction()
      try {
        $insertCmd = New-WorkstationInsertCommand -Connection $conn -Columns $cols
        $insertCmd.Transaction = $tx

        foreach ($item in $namesToCreate) {
          foreach ($c in $cols) {
            $paramName = '@' + $c
            $value = $tmplRow[$c]

            switch -Regex ($c) {
              '^Name$'          { $value = $item.Name; break }
              '^StoreID$'       { $value = $storeId; break }
              '^Disabled$'      { $value = 0; break }
              '^IsWorkStation$' { $value = 1; break }
              '^IPAddress$'     { $value = 0; break }
              '^DateUpdated$'   { $value = Get-Date; break }
              '^OtherID$'       { $value = 0; break }
              '^ExportCode_1$'  { if (-not $copyExportCodes) { $value = '' }; break }
              '^ExportCode_2$'  { if (-not $copyExportCodes) { $value = '' }; break }
            }

            if ($value -is [System.DBNull]) { $value = $null }
            Set-CommandParamValue -Command $insertCmd -ParamName $paramName -Value $value
          }

          $newId = $insertCmd.ExecuteScalar()
          Write-Log ("Inserted WorkstationID {0} for {1}" -f $newId, $item.Name)

          [void]$results.Add([pscustomobject]@{ name = $item.Name; workstationId = [int]$newId; status = "Created"; error = $null })
        }

        $tx.Commit()
        Write-Log "Transaction committed."
      } catch {
        try { $tx.Rollback() } catch { }
        throw
      }
    }

    $report.results = @($results)
    $report.success = $true
  }
  finally {
    if ($conn) { $conn.Dispose() }
  }
}
catch {
  Add-ErrorRecord -Message "Tool failed." -Exception $_.Exception
  $report.errors = @($script:Errors)
  $report.success = $false
}
finally {
  $endUtc = (Get-Date).ToUniversalTime()
  $report.endedAtUtc = $endUtc.ToString("o")
  $report.logs = @($script:LogLines)
  if (-not $report.errors -or $report.errors.Count -eq 0) {
    $report.errors = @($script:Errors)
  }
  Save-Report -Report $report
}
