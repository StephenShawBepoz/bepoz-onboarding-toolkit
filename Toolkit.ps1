<#
Bepoz Onboarding Toolkit - Toolkit.ps1
- Downloads manifest.json from GitHub
- Downloads missing tool files on-demand (raw GitHub URLs)
- Runs tools with a per-run RunDir (Output.log + ToolkitContext.json + ToolkitRun.json)
- Shares BackOffice SQL context (HKCU:\Software\BackOffice -> SQL_Server, SQL_DSN) with every tool run:
    - Env vars: BEPOZ_SQL_SERVER, BEPOZ_SQL_DSN, BEPOZ_SQL_REGPATH
    - File:     <RunDir>\ToolkitContext.json
- WinForms UI with Output Log viewer (tails Output.log)
#>

param(
  [ValidateSet("app","init","sync","list","run")]
  [string]$Command = "app",

  [string]$ToolId,
  [string]$Ui = "auto"
)

$ErrorActionPreference = "Stop"

# ---------------- CONFIG ----------------
$RepoOwner = "StephenShawBepoz"
$RepoName  = "bepoz-onboarding-toolkit"
$Branch    = "main"

$Root      = "C:\Bepoz\OnboardingToolkit"
$CatDir    = Join-Path $Root "catalogue"
$ToolsDir  = Join-Path $Root "tools"
$RunsDir   = Join-Path $Root "runs"
$LogsDir   = Join-Path $Root "logs"
$TempDir   = Join-Path $Root "temp"

$ManifestUrl = "https://raw.githubusercontent.com/$RepoOwner/$RepoName/$Branch/manifest.json"
# ----------------------------------------

# ---------------- BACK OFFICE SQL CONTEXT ----------------
# Requirement: only read BackOffice SQL settings from HKCU:\Software\BackOffice
$BackOfficeRegPath = "HKCU:\Software\BackOffice"
$Global:ToolkitContext = $null

function Get-BackOfficeSqlSettings {
  param([string]$RegistryPath = $BackOfficeRegPath)

  $r = [ordered]@{
    registryPath = $RegistryPath
    dsn          = $null
    server       = $null
    ok           = $false
    warnings     = @()
  }

  try {
    $p = Get-ItemProperty -Path $RegistryPath -ErrorAction Stop
  } catch {
    $r.warnings += "Registry key not found: $RegistryPath"
    return [pscustomobject]$r
  }

  $dsn    = $p.SQL_DSN
  $server = $p.SQL_Server

  if ([string]::IsNullOrWhiteSpace([string]$dsn)) {
    $r.warnings += "SQL_DSN is missing or blank in $RegistryPath"
  } else {
    $r.dsn = [string]$dsn
  }

  if ([string]::IsNullOrWhiteSpace([string]$server)) {
    $r.warnings += "SQL_Server is missing or blank in $RegistryPath"
  } else {
    $r.server = [string]$server
  }

  $r.ok = ($r.warnings.Count -eq 0)
  return [pscustomobject]$r
}

function New-ToolkitContext {
  $sql = Get-BackOfficeSqlSettings

  [pscustomobject]([ordered]@{
    schemaVersion = 3
    checkedAt     = (Get-Date).ToString("o")
    backOffice    = @{
      registryPath = $sql.registryPath
      sql          = @{
        dsnName  = $sql.dsn
        server   = $sql.server
        ok       = [bool]$sql.ok
        warnings = @($sql.warnings)
      }
    }
  })
}

function Set-ToolkitSqlEnvironment {
  param([object]$Context)

  if (-not $Context) { return }

  $env:BEPOZ_SQL_REGPATH = [string]$Context.backOffice.registryPath

  $dsn = [string]$Context.backOffice.sql.dsnName
  if (-not [string]::IsNullOrWhiteSpace($dsn)) { $env:BEPOZ_SQL_DSN = $dsn }

  $srv = [string]$Context.backOffice.sql.server
  if (-not [string]::IsNullOrWhiteSpace($srv)) { $env:BEPOZ_SQL_SERVER = $srv }
}

function Write-ToolkitContextFile {
  param(
    [Parameter(Mandatory=$true)][string]$RunDir,
    [Parameter(Mandatory=$true)][object]$Context
  )
  $path = Join-Path $RunDir "ToolkitContext.json"
  ($Context | ConvertTo-Json -Depth 10) | Set-Content -Path $path -Encoding UTF8
  return $path
}
# -----------------------------------------------------------

function Ensure-Tls12 {
  try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}
}

function Ensure-Dirs {
  New-Item -ItemType Directory -Force -Path $Root, $CatDir, $ToolsDir, $RunsDir, $LogsDir, $TempDir | Out-Null
}

function Write-Log([string]$msg) {
  $ts = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
  $line = "[$ts] $msg"
  Write-Host $line
  try {
    $path = Join-Path $LogsDir ("toolkit_{0}.log" -f (Get-Date -Format "yyyyMMdd"))
    Add-Content -Path $path -Value $line -Encoding UTF8
  } catch {}
}

function Download-File([string]$url, [string]$outFile) {
  Ensure-Tls12
  $dir = Split-Path -Parent $outFile
  if ($dir) { New-Item -ItemType Directory -Force -Path $dir | Out-Null }

  Write-Log "Downloading: $url"
  try {
    Invoke-WebRequest -Uri $url -OutFile $outFile -UseBasicParsing
  } catch {
    throw "Failed to download '$url' to '$outFile'. $($_.Exception.Message)"
  }
}

function Load-Manifest {
  Ensure-Dirs
  $localManifest = Join-Path $CatDir "manifest.json"
  Write-Log "Downloading manifest..."
  Write-Log $ManifestUrl
  Download-File -url $ManifestUrl -outFile $localManifest
  $json = Get-Content -Raw -Path $localManifest -Encoding UTF8 | ConvertFrom-Json
  if (-not $json.tools) { throw "manifest.json missing 'tools' array." }
  return $json
}

function Get-Tool([object]$manifest, [string]$toolId) {
  $t = $manifest.tools | Where-Object { $_.toolId -eq $toolId } | Select-Object -First 1
  if (-not $t) { throw "ToolId not found in manifest: $toolId" }
  return $t
}

function Ensure-ToolInstalled {
  param([object]$tool)

  $entry = [string]$tool.entryPoint
  if (-not $entry) { throw "Tool '$($tool.toolId)' missing entryPoint in manifest." }

  $files = @()
  if ($tool.PSObject.Properties.Name -contains "files" -and $tool.files) {
    $files += @($tool.files)
  } else {
    $files += $entry
  }

  # Tool folder derived from entryPoint (so toolId doesn't need to match folder name)
  $entryLocalPath = Join-Path $Root ($entry -replace "/","\")
  $toolFolder = Split-Path -Parent $entryLocalPath
  if (-not $toolFolder) { throw "Cannot determine tool folder for entryPoint: $entry" }

  $metaPath = Join-Path $toolFolder ".toolmeta.json"
  $manifestVersion = [string]$tool.toolVersion

  $force = $false
  if ($env:BEPOZ_TOOLKIT_FORCE_TOOL_UPDATE -eq "1") { $force = $true }

  # If meta exists, compare versions
  if (-not $force -and (Test-Path -LiteralPath $metaPath)) {
    try {
      $meta = Get-Content -LiteralPath $metaPath -Raw -Encoding UTF8 | ConvertFrom-Json
      if ($manifestVersion -and ($meta.toolVersion -ne $manifestVersion)) { $force = $true }
    } catch {
      $force = $true
    }
  }

  # If tool exists but has never had meta written, refresh once (prevents stale tool issues)
  if (-not $force -and $manifestVersion -and (Test-Path -LiteralPath $toolFolder) -and -not (Test-Path -LiteralPath $metaPath)) {
    $force = $true
  }

  if ($force -and (Test-Path -LiteralPath $toolFolder)) {
    Write-Log "Tool update required - removing cached folder: $toolFolder"
    try { Remove-Item -LiteralPath $toolFolder -Recurse -Force -ErrorAction Stop } catch {}
  }

  foreach ($rel in $files) {
    $rel = ([string]$rel) -replace "\\","/"         # raw URLs expect /
    $dest = Join-Path $Root ($rel -replace "/","\") # local path

    $dir = Split-Path -Parent $dest
    if ($dir) { New-Item -ItemType Directory -Force -Path $dir | Out-Null }

    if ($force -or -not (Test-Path -LiteralPath $dest)) {
      # Cache-bust to avoid stale CDN responses
      $v = if ($manifestVersion) { $manifestVersion } else { (Get-Date).ToUniversalTime().Ticks }
      $url = "https://raw.githubusercontent.com/$RepoOwner/$RepoName/$Branch/$rel?v=$v"
      Write-Log "Downloading tool file: $rel (force=$force)"
      Download-File -url $url -outFile $dest
    }
  }

  # Write tool metadata
  try {
    New-Item -ItemType Directory -Force -Path $toolFolder | Out-Null
    $metaObj = [pscustomobject]@{
      toolId          = [string]$tool.toolId
      toolVersion     = $manifestVersion
      entryPoint      = $entry
      downloadedAtUtc = (Get-Date).ToUniversalTime().ToString("o")
      repo            = "$RepoOwner/$RepoName"
      branch          = $Branch
    }
    ($metaObj | ConvertTo-Json -Depth 6) | Set-Content -LiteralPath $metaPath -Encoding UTF8
  } catch {}
}

function New-RunDir([string]$toolId) {
  Ensure-Dirs
  $stamp = (Get-Date).ToString("yyyyMMdd_HHmmss")
  $runDir = Join-Path $RunsDir ("{0}_{1}" -f $toolId, $stamp)
  New-Item -ItemType Directory -Force -Path $runDir | Out-Null
  return $runDir
}

function Start-ToolProcess([object]$tool, [string]$runDir) {
  $toolId    = [string]$tool.toolId
  $entryRel  = [string]$tool.entryPoint
  $entryPath = Join-Path $Root ($entryRel -replace "/","\")
  if (-not (Test-Path -LiteralPath $entryPath)) {
    throw "EntryPoint not found on disk: $entryPath"
  }

  # Refresh SQL context for every run and share with the tool
  $Global:ToolkitContext = New-ToolkitContext
  Set-ToolkitSqlEnvironment -Context $Global:ToolkitContext
  $ctxPath = Write-ToolkitContextFile -RunDir $runDir -Context $Global:ToolkitContext

  $outLog = Join-Path $runDir "Output.log"
  $runner = Join-Path $runDir "RunTool.ps1"

@"
`$ErrorActionPreference = 'Stop'
try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}

`$runDir = '$runDir'
`$outLog = '$outLog'
`$env:BEPOZ_TOOLKIT_RUNDIR = `$runDir

Write-Host "SQL_Server:  `$env:BEPOZ_SQL_SERVER"
Write-Host "SQL_DSN:     `$env:BEPOZ_SQL_DSN"
Write-Host "Context:     $ctxPath"
Write-Host "EntryPoint:  $entryRel"
Write-Host "RunDir:      `$runDir"
Write-Host ""

try { Start-Transcript -Path `$outLog -Append | Out-Null } catch {}

Push-Location '$Root'
try {
  & '$entryPath' -RunDir `$runDir
  exit `$LASTEXITCODE
} catch {
  Write-Host "ERROR: `$(`$_.Exception.Message)"
  Write-Host `$_.ScriptStackTrace
  exit 1
} finally {
  Pop-Location
  try { Stop-Transcript | Out-Null } catch {}
}
"@ | Set-Content -Path $runner -Encoding UTF8

  $psi = New-Object System.Diagnostics.ProcessStartInfo
  $psi.FileName = "powershell.exe"
  $psi.Arguments = "-NoProfile -ExecutionPolicy Bypass -STA -File `"$runner`""
  $psi.WorkingDirectory = $Root
  $psi.UseShellExecute = $false
  $psi.CreateNoWindow = $true

  # Pass SQL env vars into the child process explicitly
  $sql = $Global:ToolkitContext.backOffice.sql
  if ($sql -and $sql.dsnName)  { $psi.EnvironmentVariables["BEPOZ_SQL_DSN"]    = [string]$sql.dsnName }
  if ($sql -and $sql.server)   { $psi.EnvironmentVariables["BEPOZ_SQL_SERVER"] = [string]$sql.server }
  $psi.EnvironmentVariables["BEPOZ_SQL_REGPATH"] = [string]$Global:ToolkitContext.backOffice.registryPath

  $p = [System.Diagnostics.Process]::Start($psi)

  # Metadata (written immediately)
  $meta = @{
    toolId     = $toolId
    toolName   = [string]$tool.name
    startedAt  = (Get-Date).ToString("o")
    runDir     = $runDir
    outputLog  = $outLog
    entryPoint = $entryRel
    processId  = $p.Id
    context    = @{
      toolkitContext = "ToolkitContext.json"
    }
  } | ConvertTo-Json -Depth 8

  $metaPath = Join-Path $runDir "ToolkitRun.json"
  $meta | Set-Content -Path $metaPath -Encoding UTF8

  return $p
}

function Run-ToolConsole([string]$toolId) {
  $m = Load-Manifest
  $t = Get-Tool -manifest $m -toolId $toolId
  Ensure-ToolInstalled -tool $t

  $runDir = New-RunDir -toolId $toolId
  Write-Host "RunDir: $runDir"

  $p = Start-ToolProcess -tool $t -runDir $runDir
  $p.WaitForExit()

  Write-Host "ExitCode: $($p.ExitCode)"
  Write-Host "Output:   $(Join-Path $runDir 'Output.log')"
}

function Sync-AllTools {
  $m = Load-Manifest
  foreach ($t in $m.tools) {
    try {
      Ensure-ToolInstalled -tool $t
    } catch {
      Write-Log "Sync failed for $($t.toolId): $($_.Exception.Message)"
      throw
    }
  }
  Write-Log "Sync complete."
}

function Ensure-AppRunsInSta {
  # WinForms behaves best in STA. If user launches app from a non-STA host,
  # relaunch Toolkit.ps1 in STA automatically.
  if ([System.Threading.Thread]::CurrentThread.ApartmentState -ne "STA") {
    $psPath = $PSCommandPath
    if (-not $psPath) { return } # cannot relaunch safely

    Write-Log "Relaunching UI in STA..."
    $args = "-NoProfile -ExecutionPolicy Bypass -STA -File `"$psPath`" app"
    Start-Process -FilePath "powershell.exe" -ArgumentList $args
    exit 0
  }
}

function Show-Ui {
  Ensure-AppRunsInSta

  Add-Type -AssemblyName System.Windows.Forms
  Add-Type -AssemblyName System.Drawing
  [System.Windows.Forms.Application]::EnableVisualStyles()

  Ensure-Dirs

  # Read SQL settings on open and show them
  $Global:ToolkitContext = New-ToolkitContext
  Set-ToolkitSqlEnvironment -Context $Global:ToolkitContext

  $m = Load-Manifest

  $form = New-Object System.Windows.Forms.Form
  $form.Text = "Bepoz Onboarding Toolkit"
  $form.Width = 980
  $form.Height = 620
  $form.StartPosition = "CenterScreen"

  $list = New-Object System.Windows.Forms.ListBox
  $list.Width = 280
  $list.Dock = "Left"

  $panel = New-Object System.Windows.Forms.Panel
  $panel.Dock = "Fill"

  $lblName = New-Object System.Windows.Forms.Label
  $lblName.AutoSize = $false
  $lblName.Height = 26
  $lblName.Dock = "Top"
  $lblName.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
  $lblName.Padding = "8,6,8,0"

  $lblSql = New-Object System.Windows.Forms.Label
  $lblSql.AutoSize = $false
  $lblSql.Height = 42
  $lblSql.Dock = "Top"
  $lblSql.Padding = "8,0,8,0"
  $lblSql.Font = New-Object System.Drawing.Font("Segoe UI", 9)

  $sql = $Global:ToolkitContext.backOffice.sql
  if ($sql.ok) {
    $lblSql.Text = "SQL: {0}   DSN: {1}   (from {2})" -f $sql.server, $sql.dsnName, $Global:ToolkitContext.backOffice.registryPath
  } else {
    $lblSql.Text = "BackOffice SQL settings missing/incomplete (HKCU:\Software\BackOffice). Tools may fail to connect."
  }

  $txtDesc = New-Object System.Windows.Forms.TextBox
  $txtDesc.Multiline = $true
  $txtDesc.ReadOnly = $true
  $txtDesc.Height = 90
  $txtDesc.Dock = "Top"
  $txtDesc.ScrollBars = "Vertical"

  $tabs = New-Object System.Windows.Forms.TabControl
  $tabs.Dock = "Fill"

  $tabOut = New-Object System.Windows.Forms.TabPage
  $tabOut.Text = "Output Log"
  $outBox = New-Object System.Windows.Forms.TextBox
  $outBox.Multiline = $true
  $outBox.ReadOnly = $true
  $outBox.ScrollBars = "Vertical"
  $outBox.Dock = "Fill"
  $tabOut.Controls.Add($outBox)

  $tabInfo = New-Object System.Windows.Forms.TabPage
  $tabInfo.Text = "Toolkit Info"
  $infoBox = New-Object System.Windows.Forms.TextBox
  $infoBox.Multiline = $true
  $infoBox.ReadOnly = $true
  $infoBox.ScrollBars = "Vertical"
  $infoBox.Dock = "Fill"
  $infoBox.Text = @"
Tools root:
  $Root

Runs:
  $RunsDir

SQL context (auto-detected for the CURRENT Windows user):
  Registry:   $env:BEPOZ_SQL_REGPATH
  SQL Server: $env:BEPOZ_SQL_SERVER
  DSN name:   $env:BEPOZ_SQL_DSN

How the Toolkit finds SQL settings:
  1) HKCU\Software\BackOffice values:
       - SQL_Server  (SQL Server instance)
       - SQL_DSN     (BackOffice DSN name)

If SQL values are missing:
  - Make sure Back Office has been launched at least once as this user.
  - Check the registry key above contains SQL_Server and SQL_DSN.
"@
  $tabInfo.Controls.Add($infoBox)

  $tabs.TabPages.Add($tabOut) | Out-Null
  $tabs.TabPages.Add($tabInfo) | Out-Null

  $btnPanel = New-Object System.Windows.Forms.FlowLayoutPanel
  $btnPanel.Dock = "Bottom"
  $btnPanel.Height = 44
  $btnPanel.FlowDirection = "RightToLeft"
  $btnPanel.Padding = "8,8,8,8"

  $btnExit = New-Object System.Windows.Forms.Button
  $btnExit.Text = "Exit"
  $btnExit.Width = 90
  $btnExit.Add_Click({ $form.Close() })
  $btnPanel.Controls.Add($btnExit) | Out-Null

  $btnRefresh = New-Object System.Windows.Forms.Button
  $btnRefresh.Text = "Refresh"
  $btnRefresh.Width = 90
  $btnPanel.Controls.Add($btnRefresh) | Out-Null

  $btnRun = New-Object System.Windows.Forms.Button
  $btnRun.Text = "Run"
  $btnRun.Width = 90
  $btnPanel.Controls.Add($btnRun) | Out-Null

  $panel.Controls.Add($tabs)
  $panel.Controls.Add($txtDesc)
  $panel.Controls.Add($lblSql)
  $panel.Controls.Add($lblName)
  $panel.Controls.Add($btnPanel)

  $form.Controls.Add($panel)
  $form.Controls.Add($list)

  $toolMap = @{}
  foreach ($t in $m.tools) {
    $display = "{0}  (v{1})" -f $t.toolId, $t.toolVersion
    $toolMap[$display] = $t
    [void]$list.Items.Add($display)
  }

  $currentOutLog = $null
  $lastOutLen = 0

  $timer = New-Object System.Windows.Forms.Timer
  $timer.Interval = 600
  $timer.Add_Tick({
    if ($currentOutLog -and (Test-Path -LiteralPath $currentOutLog)) {
      try {
        $text = Get-Content -Raw -Path $currentOutLog -ErrorAction Stop
        if ($text.Length -ne $lastOutLen) {
          $outBox.Text = $text
          $outBox.SelectionStart = $outBox.TextLength
          $outBox.ScrollToCaret()
          $lastOutLen = $text.Length
        }
      } catch {}
    }
  })

  $list.Add_SelectedIndexChanged({
    $sel = $list.SelectedItem
    if (-not $sel) { return }
    $t = $toolMap[$sel]
    $lblName.Text = "{0}" -f $t.name
    $txtDesc.Text = "{0}`r`n`r`nToolId: {1}    Version: {2}`r`nEntryPoint: {3}" -f $t.description, $t.toolId, $t.toolVersion, $t.entryPoint
  })

  $btnRefresh.Add_Click({
    try {
      $m2 = Load-Manifest
      $list.Items.Clear()
      $toolMap.Clear()
      foreach ($t in $m2.tools) {
        $display = "{0}  (v{1})" -f $t.toolId, $t.toolVersion
        $toolMap[$display] = $t
        [void]$list.Items.Add($display)
      }
    } catch {
      [System.Windows.Forms.MessageBox]::Show("Refresh failed: $($_.Exception.Message)") | Out-Null
    }
  })

  $btnRun.Add_Click({
    try {
      $sel = $list.SelectedItem
      if (-not $sel) {
        [System.Windows.Forms.MessageBox]::Show("Pick a tool first.") | Out-Null
        return
      }

      $t = $toolMap[$sel]

      Ensure-ToolInstalled -tool $t

      $runDir = New-RunDir -toolId $t.toolId
      $currentOutLog = Join-Path $runDir "Output.log"
      $lastOutLen = 0

      $outBox.Text = "Running $($t.toolId)...`r`nRunDir: $runDir`r`n"
      $timer.Start()

      $null = Start-ToolProcess -tool $t -runDir $runDir

    } catch {
      $timer.Stop()
      [System.Windows.Forms.MessageBox]::Show("Run failed: $($_.Exception.Message)") | Out-Null
    }
  })

  if (-not $sql.ok) {
    $warn = "BackOffice SQL settings were not found or are incomplete for this Windows user.`r`n`r`nToolkit reads:`r`n  HKCU\Software\BackOffice (values SQL_Server and SQL_DSN)`r`n`r`nFix options:`r`n  1) Launch Back Office once as this user (often re-creates the registry values).`r`n  2) Confirm SQL_Server and SQL_DSN exist under HKCU\Software\BackOffice.`r`n`r`nTools that query SQL may fail until these are set."
    [System.Windows.Forms.MessageBox]::Show($warn, "Bepoz Toolkit - SQL settings", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
  }

  if ($list.Items.Count -gt 0) { $list.SelectedIndex = 0 }
  [void]$form.ShowDialog()
}

try {
  switch ($Command) {
    "init" {
      Ensure-Dirs
      Write-Log "Initialised at: $Root"
    }
    "sync" {
      Ensure-Dirs
      Sync-AllTools
    }
    "list" {
      $m = Load-Manifest
      $m.tools | ForEach-Object { "{0}`t{1}`t{2}" -f $_.toolId, $_.toolVersion, $_.name } | Write-Output
    }
    "run" {
      if (-not $ToolId) { throw "ToolId is required for 'run' command." }
      Run-ToolConsole -toolId $ToolId
    }
    default {
      Show-Ui
    }
  }
}
catch {
  Write-Host ""
  Write-Host "=== TOOLKIT ERROR ==="
  Write-Host $_.Exception.Message
  Write-Host ""
  Write-Host "Press ENTER to close:"
  Read-Host | Out-Null
}
