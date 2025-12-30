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
# All tools depend on SQL. The toolbox reads these values once and shares them with every tool run:
#   HKCU:\Software\BackOffice
#     - SQL_DSN    (database name)
#     - SQL_Server (server\instance)
$BackOfficeRegPath = "HKCU:\Software\BackOffice"

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

  $ctx = [ordered]@{
    schemaVersion = 1
    checkedAt     = (Get-Date).ToString("o")
    backOffice    = @{
      registryPath = $sql.registryPath
      sql          = @{
        dsn      = $sql.dsn
        server   = $sql.server
        ok       = [bool]$sql.ok
        warnings = @($sql.warnings)
      }
    }
  }

  return [pscustomobject]$ctx
}

function Set-ToolkitSqlEnvironment {
  param([object]$Context)

  if (-not $Context) { return }
  $env:BEPOZ_SQL_REGPATH = [string]$Context.backOffice.registryPath

  $dsn = [string]$Context.backOffice.sql.dsn
  if (-not [string]::IsNullOrWhiteSpace($dsn)) { $env:BEPOZ_SQL_DSN = $dsn }

  $srv = [string]$Context.backOffice.sql.server
  if (-not [string]::IsNullOrWhiteSpace($srv)) { $env:BEPOZ_SQL_SERVER = $srv }
}

# Initialise context on startup (refreshed again in the UI and per tool run).
$Global:ToolkitContext = New-ToolkitContext
Set-ToolkitSqlEnvironment -Context $Global:ToolkitContext
# -----------------------------------------------------------

function Ensure-Dirs {
  New-Item -ItemType Directory -Force -Path $Root, $CatDir, $ToolsDir, $RunsDir, $LogsDir, $TempDir | Out-Null
}

function Download-File([string]$url, [string]$outFile) {
  Write-Host "Downloading: $url"
  $wc = New-Object System.Net.WebClient
  try {
    $wc.Headers["User-Agent"] = "BepozOnboardingToolkit"
    $wc.DownloadFile($url, $outFile)
  } finally {
    $wc.Dispose()
  }
}

function Load-Manifest {
  Ensure-Dirs
  $localManifest = Join-Path $CatDir "manifest.json"
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

function Download-File([string]$url, [string]$outFile) {
  Ensure-Tls12
  $dir = Split-Path -Parent $outFile
  if ($dir) { New-Item -ItemType Directory -Force -Path $dir | Out-Null }
  Invoke-WebRequest -Uri $url -OutFile $outFile -UseBasicParsing
}

function Ensure-ToolInstalled([object]$tool) {
  $entry = [string]$tool.entryPoint
  if (-not $entry) { throw "Tool '$($tool.toolId)' missing entryPoint in manifest." }

  $files = @()
  if ($tool.PSObject.Properties.Name -contains "files" -and $tool.files) {
    $files += @($tool.files)
  } else {
    $files += $entry
  }

  foreach ($rel in $files) {
    $rel = $rel -replace "\\","/"
    $dest = Join-Path $Root ($rel -replace "/","\")
    if (-not (Test-Path $dest)) {
      $url = "https://raw.githubusercontent.com/$RepoOwner/$RepoName/$Branch/$rel"
      Write-Log "Downloading tool file: $rel"
      Download-File -url $url -outFile $dest
    }
  }
}

function New-RunDir([string]$toolId) {
  Ensure-Dirs
  $stamp = (Get-Date).ToString("yyyyMMdd_HHmmss")
  $runDir = Join-Path $RunsDir ("{0}_{1}" -f $toolId, $stamp)
  New-Item -ItemType Directory -Force -Path $runDir | Out-Null
  return $runDir
}

function Clear-ToolkitCache {
  param(
    [switch]$OnOpen,
    [int]$PruneRunsOlderThanDays = 30
  )

  if ($OnOpen) {
    # Prune old runs
    $cut = (Get-Date).AddDays(-$PruneRunsOlderThanDays)
    Get-ChildItem -Path $RunsDir -Directory -ErrorAction SilentlyContinue |
      Where-Object { $_.LastWriteTime -lt $cut } |
      ForEach-Object { try { Remove-Item -Recurse -Force -Path $_.FullName } catch {} }
  }
}

function Start-ToolProcess([object]$tool, [string]$runDir) {
  $toolId    = [string]$tool.toolId
  $entryRel  = [string]$tool.entryPoint
  $entryPath = Join-Path $Root ($entryRel -replace "/","\")
  if (-not (Test-Path $entryPath)) { throw "EntryPoint not found on disk: $entryPath" }

  $outLog = Join-Path $runDir "Output.log"
  $runner = Join-Path $runDir "RunTool.ps1"

  # Refresh SQL context for this run (reads HKCU:\Software\BackOffice)
  $Global:ToolkitContext = New-ToolkitContext
  Set-ToolkitSqlEnvironment -Context $Global:ToolkitContext
  $ctxJson = $Global:ToolkitContext | ConvertTo-Json -Depth 10

  # This runner captures ALL output (including Write-Host) via transcript,
  # and runs in STA so WinForms dialogs can display.
@"
`$ErrorActionPreference = 'Stop'
try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}

`$runDir = '$runDir'
`$outLog = '$outLog'
`$env:BEPoz_TOOLKIT_RUNDIR = `$runDir

New-Item -ItemType Directory -Force -Path `$runDir | Out-Null
# Share SQL context with tools (available via env vars and ToolkitContext.json)
`$ctxPath = Join-Path `$runDir "ToolkitContext.json"
@'
$ctxJson
'@ | Set-Content -Path `$ctxPath -Encoding UTF8

Write-Host "SQL_Server:  $env:BEPOZ_SQL_SERVER"
Write-Host "SQL_DSN:     $env:BEPOZ_SQL_DSN"
Write-Host "Context:     `$ctxPath"


try {
  Start-Transcript -Path `$outLog -Append | Out-Null
} catch {}

Push-Location '$Root'
try {
  Write-Host "Running tool: $toolId"
  Write-Host "EntryPoint:   $entryRel"
  Write-Host "RunDir:       `$runDir"
  & '$entryPath' -RunDir `$runDir
  exit `$LASTEXITCODE
} catch {
  Write-Host "ERROR: $($_.Exception.Message)"
  Write-Host $_.ScriptStackTrace
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

  # Pass SQL settings to the child PowerShell process
  $sql = $Global:ToolkitContext.backOffice.sql
  if ($sql -and $sql.dsn)    { $psi.EnvironmentVariables["BEPOZ_SQL_DSN"]    = [string]$sql.dsn }
  if ($sql -and $sql.server) { $psi.EnvironmentVariables["BEPOZ_SQL_SERVER"] = [string]$sql.server }
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
  } | ConvertTo-Json -Depth 6

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

function Show-Ui {
  Add-Type -AssemblyName System.Windows.Forms
  Add-Type -AssemblyName System.Drawing
  [System.Windows.Forms.Application]::EnableVisualStyles()

  Ensure-Dirs
  Clear-ToolkitCache -OnOpen -PruneRunsOlderThanDays 30

  # Read BackOffice SQL settings on open and expose them to every tool
  $Global:ToolkitContext = New-ToolkitContext
  Set-ToolkitSqlEnvironment -Context $Global:ToolkitContext

  if (-not $Global:ToolkitContext.backOffice.sql.ok) {
    $warn = "BackOffice SQL settings were not found or are incomplete.`r`n`r`nExpected in HKCU:\Software\BackOffice:`r`n  SQL_Server  (server\instance)`r`n  SQL_DSN     (database name)`r`n`r`nTools that query SQL may fail until these are set."
    [System.Windows.Forms.MessageBox]::Show($warn, "Bepoz Toolkit - SQL settings", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
  }

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
  $lblSql.Text = ""

  # Display BackOffice SQL settings (shared with tools via env vars + ToolkitContext.json)
  $sql = $Global:ToolkitContext.backOffice.sql
  if ($sql.ok) {
    $lblSql.Text = "SQL: {0}   DB: {1}   (from {2})" -f $sql.server, $sql.dsn, $Global:ToolkitContext.backOffice.registryPath
  } else {
    $lblSql.Text = "SQL settings missing/incomplete (HKCU:\Software\BackOffice). Tools may fail to connect."
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
  $infoBox.Text = "Tools run from: $Root`r`nRuns folder: $RunsDir`r`n`r`nSQL context:`r`n  BEPOZ_SQL_SERVER=$env:BEPOZ_SQL_SERVER`r`n  BEPOZ_SQL_DSN=$env:BEPOZ_SQL_DSN"
  $infoBox.ScrollBars = "Vertical"
  $infoBox.Dock = "Fill"
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
    if ($currentOutLog -and (Test-Path $currentOutLog)) {
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
      [System.Windows.Forms.MessageBox]::Show("Refresh failed: $($_.Exception.Message)")
    }
  })

  $btnRun.Add_Click({
    try {
      $sel = $list.SelectedItem
      if (-not $sel) {
        [System.Windows.Forms.MessageBox]::Show("Pick a tool first.")
        return
      }

      $t = $toolMap[$sel]
      Ensure-ToolInstalled -tool $t

      $runDir = New-RunDir -toolId $t.toolId
      $currentOutLog = Join-Path $runDir "Output.log"
      $lastOutLen = 0

      $outBox.Text = "Running $($t.toolId)...`r`nRunDir: $runDir`r`n"
      $timer.Start()

      # Start tool detached (STA, no console window)
      $null = Start-ToolProcess -tool $t -runDir $runDir

    } catch {
      $timer.Stop()
      [System.Windows.Forms.MessageBox]::Show("Run failed: $($_.Exception.Message)")
    }
  })

  if ($list.Items.Count -gt 0) { $list.SelectedIndex = 0 }
  [void]$form.ShowDialog()
}

try {
  switch ($Command) {
    "init" {
      Ensure-Dirs
      Write-Host "Initialised at: $Root"
    }
    "sync" {
      Ensure-Dirs
      Write-Host "Sync is repo-based in this toolkit layout. Ensure the repo contents exist under: $Root"
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
