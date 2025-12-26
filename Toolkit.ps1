<# 
Bepoz Onboarding Toolkit - Toolkit.ps1
- Windows PowerShell 5.1
- Root: C:\Bepoz\OnboardingToolkit
- Private GitHub via $env:GITHUB_TOKEN (fine-grained read-only token)

Modes:
- GUI WinForms (interactive)
- Console menu (interactive console host)
- Headless fallback (non-interactive; auto-runs Doctor)

Automation:
-Command run -ToolId <id> -ExitAfterRun
#>

[CmdletBinding()]
param(
  [ValidateSet("init","sync","list-remote","install","list-local","run","app")]
  [string]$Command = "app",

  [string]$ToolId,

  [string]$RepoOwner = "StephenShawBepoz",
  [string]$RepoName  = "bepoz-onboarding-toolkit",
  [string]$Branch    = "main",

  [string]$GitHubToken = $env:GITHUB_TOKEN,

  [switch]$ExitAfterRun,

  [ValidateSet("auto","console","winforms")]
  [string]$Ui = "auto"
)

$ErrorActionPreference = "Stop"

$Root      = "C:\Bepoz\OnboardingToolkit"
$ToolsRoot = Join-Path $Root "tools"
$RunsRoot  = Join-Path $Root "runs"
$TempRoot  = Join-Path $Root "temp"
$CatRoot   = Join-Path $Root "catalogue"
$LogsRoot  = Join-Path $Root "logs"

$ManifestLocal = Join-Path $CatRoot "manifest.json"
$ManifestUrl   = "https://raw.githubusercontent.com/$RepoOwner/$RepoName/$Branch/manifest.json"

function Ensure-Tls12 { try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {} }
function Ensure-Folders { New-Item -ItemType Directory -Force -Path $Root,$ToolsRoot,$RunsRoot,$TempRoot,$CatRoot,$LogsRoot | Out-Null }

function Log-Info([string]$msg) {
  Ensure-Folders
  $p = Join-Path $LogsRoot ("toolkit_{0}.log" -f (Get-Date -Format "yyyyMMdd"))
  ("[{0}] {1}" -f (Get-Date -Format "s"), $msg) | Out-File -FilePath $p -Append -Encoding utf8
}

function Invoke-DownloadFile {
  param([Parameter(Mandatory=$true)][string]$Url,[Parameter(Mandatory=$true)][string]$OutFile)
  Ensure-Tls12
  $headers = @{ "User-Agent" = "BepozOnboardingToolkit" }
  if ($GitHubToken) { $headers["Authorization"] = "Bearer $GitHubToken" }
  Invoke-WebRequest -Uri $Url -Headers $headers -OutFile $OutFile -UseBasicParsing
}

function Read-Manifest { if (-not (Test-Path $ManifestLocal)) { return $null }; Get-Content $ManifestLocal -Raw | ConvertFrom-Json }
function Sync-Manifest  { Ensure-Folders; Invoke-DownloadFile -Url $ManifestUrl -OutFile $ManifestLocal }

function Get-LocalInstalledToolIds {
  if (-not (Test-Path $ToolsRoot)) { return @() }
  @(Get-ChildItem $ToolsRoot -Directory | Select-Object -ExpandProperty Name)
}

function Get-ToolsModel {
  $m = Read-Manifest
  $local = Get-LocalInstalledToolIds
  $tools = @()

  if ($m -and $m.tools) {
    foreach ($t in $m.tools) {
      $tools += [pscustomobject]@{
        toolId      = [string]$t.toolId
        name        = [string]$t.name
        description = [string]$t.description
        toolVersion = [string]$t.toolVersion
        entryPoint  = [string]$t.entryPoint
        installed   = ($local -contains [string]$t.toolId)
      }
    }
  } else {
    foreach ($id in $local) {
      $tools += [pscustomobject]@{
        toolId      = [string]$id
        name        = [string]$id
        description = "(local only)"
        toolVersion = ""
        entryPoint  = ""
        installed   = $true
      }
    }
  }
  return $tools
}

function Install-ToolFromRepo([string]$toolId) {
  Ensure-Folders
  $m = Read-Manifest
  if (-not $m) { throw "No manifest found. Run sync first." }
  $t = $m.tools | Where-Object { $_.toolId -eq $toolId } | Select-Object -First 1
  if (-not $t) { throw "ToolId '$toolId' not found in manifest." }
  if ($t.packageType -ne "repo") { throw "Only packageType 'repo' supported." }
  if (-not $t.entryPoint) { throw "Tool '$toolId' missing entryPoint." }

  $rawUrl = "https://raw.githubusercontent.com/$RepoOwner/$RepoName/$Branch/$($t.entryPoint)"
  $toolDir = Join-Path $ToolsRoot $toolId
  New-Item -ItemType Directory -Force -Path $toolDir | Out-Null

  $dest = Join-Path $toolDir "run.ps1"
  Invoke-DownloadFile -Url $rawUrl -OutFile $dest

  $meta = [ordered]@{
    toolId       = $toolId
    toolVersion  = [string]$t.toolVersion
    installedUtc = (Get-Date).ToUniversalTime().ToString("s")+"Z"
    source       = "repo"
    entryPoint   = [string]$t.entryPoint
    rawUrl       = $rawUrl
    repo         = "$RepoOwner/$RepoName@$Branch"
  }
  ($meta | ConvertTo-Json -Depth 6) | Out-File (Join-Path $toolDir "installed.json") -Encoding utf8

  if (-not (Test-Path $dest)) { throw "Install failed: run.ps1 not written." }
}

function New-RunDir([string]$name) {
  $stamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
  $runId = ([Guid]::NewGuid().ToString("N")).Substring(0,8)
  $dir = Join-Path $RunsRoot "${stamp}_${name}_${runId}"
  New-Item -ItemType Directory -Force -Path $dir | Out-Null
  return $dir
}

function Run-Tool([string]$toolId) {
  Ensure-Folders
  $entry = Join-Path (Join-Path $ToolsRoot $toolId) "run.ps1"
  if (-not (Test-Path $entry)) { throw "Tool not installed: $toolId" }

  $runDir = New-RunDir $toolId
  $stdout = Join-Path $runDir "stdout.log"
  $stderr = Join-Path $runDir "stderr.log"

  $args = @(
    "-NoProfile",
    "-ExecutionPolicy","Bypass",
    "-File", "`"$entry`"",
    "-RunDir", "`"$runDir`""
  ) -join " "

  $p = Start-Process -FilePath "powershell.exe" -ArgumentList $args -Wait -PassThru -NoNewWindow `
        -RedirectStandardOutput $stdout -RedirectStandardError $stderr

  $report = [ordered]@{
    ToolId      = $toolId
    RunDir      = $runDir
    MachineName = $env:COMPUTERNAME
    UserName    = $env:USERNAME
    ExitCode    = $p.ExitCode
    StdOutPath  = $stdout
    StdErrPath  = $stderr
  }
  ($report | ConvertTo-Json -Depth 6) | Out-File (Join-Path $runDir "RunReport.json") -Encoding utf8

  return [pscustomobject]@{
    ExitCode   = $p.ExitCode
    RunDir     = $runDir
    ReportPath = (Join-Path $runDir "RunReport.json")
    StdOutPath = $stdout
    StdErrPath = $stderr
  }
}

function Is-InteractiveDesktop {
  try {
    $sid = [System.Diagnostics.Process]::GetCurrentProcess().SessionId
    return ([Environment]::UserInteractive -and $sid -ne 0)
  } catch { return $false }
}

function Can-UseWinForms {
  if ($Ui -eq "console") { return $false }
  if (-not (Is-InteractiveDesktop)) { return $false }
  try {
    Add-Type -AssemblyName System.Windows.Forms | Out-Null
    Add-Type -AssemblyName System.Drawing | Out-Null
    return $true
  } catch { return $false }
}

function Invoke-HeadlessDefault {
  Log-Info "Headless mode detected; running Doctor."
  try { Sync-Manifest } catch { Log-Info ("Sync failed: " + $_.Exception.Message) }

  try { Install-ToolFromRepo "doctor" } catch { Log-Info ("Install doctor failed: " + $_.Exception.Message) }
  $r = Run-Tool "doctor"

  Log-Info ("Doctor exit code: {0}; report: {1}" -f $r.ExitCode, $r.ReportPath)
  Write-Output ("Headless complete. Report: {0}" -f $r.ReportPath)

  if ($ExitAfterRun) { exit $r.ExitCode }
}

function Append-Log([System.Windows.Forms.TextBox]$tb, [string]$line) {
  if ($null -eq $tb) { return }
  $tb.AppendText(($line + [Environment]::NewLine))
  $tb.SelectionStart = $tb.Text.Length
  $tb.ScrollToCaret()
}

function Invoke-AppConsole {
  # Only safe when a console is actually attached
  if ($Host.Name -ne "ConsoleHost") { Invoke-HeadlessDefault; return }

  while ($true) {
    Clear-Host
    Write-Host "=== Bepoz Onboarding Toolkit ==="
    Write-Host "Root: $Root"
    Write-Host "Repo: $RepoOwner/$RepoName@$Branch"
    Write-Host ""

    try { Sync-Manifest; Write-Host "Catalogue synced." } catch { Write-Host "Sync warning: $($_.Exception.Message)" }

    $tools = Get-ToolsModel
    if ($tools.Count -eq 0) { Read-Host "No tools available. ENTER to quit" | Out-Null; return }

    Write-Host ""
    for ($i=0; $i -lt $tools.Count; $i++) {
      $t = $tools[$i]
      $idx = $i + 1
      $ver = if ($t.toolVersion) { "v$($t.toolVersion)" } else { "" }
      $inst = if ($t.installed) { "Yes" } else { "No" }
      Write-Host ("[{0}] {1} {2} (Installed: {3})" -f $idx, $t.toolId, $ver, $inst)
      if ($t.description) { Write-Host ("     {0}" -f $t.description) }
    }

    Write-Host ""
    $choice = Read-Host "Enter number, or Q to quit"
    if ($choice -match '^(q|quit)$') { return }

    $n=0
    if (-not [int]::TryParse($choice,[ref]$n)) { continue }
    if ($n -lt 1 -or $n -gt $tools.Count) { continue }

    $sel = $tools[$n-1]
    try {
      Install-ToolFromRepo $sel.toolId
      $r = Run-Tool $sel.toolId
      Write-Host ""
      Write-Host "ExitCode: $($r.ExitCode)"
      Write-Host "Report:   $($r.ReportPath)"
      Read-Host "ENTER to return to menu" | Out-Null
    } catch {
      Write-Host "ERROR: $($_.Exception.Message)"
      Read-Host "ENTER to return" | Out-Null
    }
  }
}

function Invoke-AppWinForms {
  Ensure-Folders
  [System.Windows.Forms.Application]::EnableVisualStyles()

  $form = New-Object System.Windows.Forms.Form
  $form.Text = "Bepoz Onboarding Toolkit"
  $form.Width = 980
  $form.Height = 640
  $form.StartPosition = "CenterScreen"

  $status = New-Object System.Windows.Forms.StatusStrip
  $statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
  $statusLabel.Text = "Ready"
  $status.Items.Add($statusLabel) | Out-Null
  $form.Controls.Add($status)

  $mainSplit = New-Object System.Windows.Forms.SplitContainer
  $mainSplit.Dock = "Fill"
  $mainSplit.Orientation = "Vertical"
  $mainSplit.SplitterDistance = 340
  $form.Controls.Add($mainSplit)

  $lblTools = New-Object System.Windows.Forms.Label
  $lblTools.Text = "Available Tools"
  $lblTools.Dock = "Top"
  $lblTools.Height = 24
  $lblTools.Padding = "8,6,0,0"
  $mainSplit.Panel1.Controls.Add($lblTools)

  $lst = New-Object System.Windows.Forms.ListBox
  $lst.Dock = "Fill"
  $lst.IntegralHeight = $false
  $mainSplit.Panel1.Controls.Add($lst)

  $rightSplit = New-Object System.Windows.Forms.SplitContainer
  $rightSplit.Dock = "Fill"
  $rightSplit.Orientation = "Horizontal"
  $rightSplit.SplitterDistance = 190
  $mainSplit.Panel2.Controls.Add($rightSplit)

  $details = New-Object System.Windows.Forms.TextBox
  $details.Dock = "Fill"
  $details.Multiline = $true
  $details.ReadOnly = $true
  $details.ScrollBars = "Vertical"
  $details.Font = New-Object System.Drawing.Font("Consolas", 10)
  $rightSplit.Panel1.Controls.Add($details)

  $output = New-Object System.Windows.Forms.TextBox
  $output.Dock = "Fill"
  $output.Multiline = $true
  $output.ReadOnly = $true
  $output.ScrollBars = "Vertical"
  $output.Font = New-Object System.Drawing.Font("Consolas", 10)
  $rightSplit.Panel2.Controls.Add($output)

  $btnPanel = New-Object System.Windows.Forms.Panel
  $btnPanel.Dock = "Bottom"
  $btnPanel.Height = 46
  $btnPanel.Padding = "8,8,8,8"
  $rightSplit.Panel2.Controls.Add($btnPanel)
  $btnPanel.BringToFront()

  $btnRefresh = New-Object System.Windows.Forms.Button
  $btnRefresh.Text = "Refresh"
  $btnRefresh.Width = 110
  $btnRefresh.Left = 8
  $btnRefresh.Top = 8

  $btnInstall = New-Object System.Windows.Forms.Button
  $btnInstall.Text = "Install / Update"
  $btnInstall.Width = 140
  $btnInstall.Left = 128
  $btnInstall.Top = 8

  $btnRun = New-Object System.Windows.Forms.Button
  $btnRun.Text = "Run"
  $btnRun.Width = 110
  $btnRun.Left = 278
  $btnRun.Top = 8

  $btnOpenRuns = New-Object System.Windows.Forms.Button
  $btnOpenRuns.Text = "Open Runs"
  $btnOpenRuns.Width = 120
  $btnOpenRuns.Left = 398
  $btnOpenRuns.Top = 8

  $btnExit = New-Object System.Windows.Forms.Button
  $btnExit.Text = "Exit"
  $btnExit.Width = 110
  $btnExit.Left = 528
  $btnExit.Top = 8

  $btnPanel.Controls.AddRange(@($btnRefresh,$btnInstall,$btnRun,$btnOpenRuns,$btnExit))

  $toolsModel = @()

  function Set-Status([string]$t) { $statusLabel.Text = $t }

  function Reload-Tools {
    $lst.Items.Clear()
    $details.Clear()
    $toolsModel = Get-ToolsModel
    foreach ($t in $toolsModel) {
      $ver = if ($t.toolVersion) { "v$($t.toolVersion)" } else { "" }
      $inst = if ($t.installed) { "Installed" } else { "Not installed" }
      $lst.Items.Add("$($t.toolId)  $ver  [$inst]") | Out-Null
    }
    Append-Log $output ("Loaded {0} tool(s)." -f $toolsModel.Count)
  }

  function Show-SelectedDetails([int]$idx) {
    if ($idx -lt 0 -or $idx -ge $toolsModel.Count) { return }
    $t = $toolsModel[$idx]
    $details.Text =
@"
ToolId:      $($t.toolId)
Name:        $($t.name)
Version:     $($t.toolVersion)
Installed:   $($t.installed)
EntryPoint:  $($t.entryPoint)

Description:
$($t.description)

Root:        $Root
Repo:        $RepoOwner/$RepoName@$Branch
"@
  }

  function Run-Background([string]$title,[scriptblock]$work,[scriptblock]$done) {
    $btnRefresh.Enabled = $false
    $btnInstall.Enabled = $false
    $btnRun.Enabled = $false
    $btnOpenRuns.Enabled = $false
    $btnExit.Enabled = $false
    Set-Status $title

    $bw = New-Object System.ComponentModel.BackgroundWorker
    $bw.DoWork += { param($s,$e) try { $e.Result = & $work } catch { $e.Result = $_ } }
    $bw.RunWorkerCompleted += {
      param($s,$e)
      try { & $done $e.Result } catch {}
      $btnRefresh.Enabled = $true
      $btnInstall.Enabled = $true
      $btnRun.Enabled = $true
      $btnOpenRuns.Enabled = $true
      $btnExit.Enabled = $true
      Set-Status "Ready"
    }
    $bw.RunWorkerAsync() | Out-Null
  }

  $lst.Add_SelectedIndexChanged({ Show-SelectedDetails $lst.SelectedIndex })
  $btnExit.Add_Click({ $form.Close() })

  $btnOpenRuns.Add_Click({
    try { Start-Process -FilePath "explorer.exe" -ArgumentList "`"$RunsRoot`"" | Out-Null } catch {}
  })

  $btnRefresh.Add_Click({
    Append-Log $output "Refreshing catalogue..."
    Run-Background "Refreshing..." { Sync-Manifest; $true } {
      param($r)
      if ($r -is [System.Management.Automation.ErrorRecord] -or $r -is [Exception]) {
        Append-Log $output ("Refresh failed: {0}" -f $r.Exception.Message)
      } else { Append-Log $output "Catalogue refreshed." }
      Reload-Tools
      if ($lst.Items.Count -gt 0 -and $lst.SelectedIndex -lt 0) { $lst.SelectedIndex = 0 }
    }
  })

  $btnInstall.Add_Click({
    if ($lst.SelectedIndex -lt 0) { return }
    $t = $toolsModel[$lst.SelectedIndex]
    Append-Log $output ("Installing/updating: {0}" -f $t.toolId)
    Run-Background ("Installing {0}..." -f $t.toolId) { Install-ToolFromRepo $t.toolId; $true } {
      param($r)
      if ($r -is [System.Management.Automation.ErrorRecord] -or $r -is [Exception]) {
        Append-Log $output ("Install failed: {0}" -f $r.Exception.Message)
      } else { Append-Log $output "Install/update complete." }
      Reload-Tools
      Show-SelectedDetails $lst.SelectedIndex
    }
  })

  $btnRun.Add_Click({
    if ($lst.SelectedIndex -lt 0) { return }
    $t = $toolsModel[$lst.SelectedIndex]

    Run-Background ("Running {0}..." -f $t.toolId) {
      Install-ToolFromRepo $t.toolId
      Run-Tool $t.toolId
    } {
      param($r)
      if ($r -is [System.Management.Automation.ErrorRecord] -or $r -is [Exception]) {
        Append-Log $output ("Run failed: {0}" -f $r.Exception.Message)
        return
      }
      Append-Log $output ("ExitCode: {0}" -f $r.ExitCode)
      Append-Log $output ("Report:   {0}" -f $r.ReportPath)
      Append-Log $output ("RunDir:   {0}" -f $r.RunDir)

      if ($r.ExitCode -ne 0 -and (Test-Path $r.StdErrPath)) {
        Append-Log $output "---- stderr (first 30 lines) ----"
        Get-Content $r.StdErrPath -TotalCount 30 | ForEach-Object { Append-Log $output $_ }
      }

      try { Start-Process -FilePath "explorer.exe" -ArgumentList "`"$($r.RunDir)`"" | Out-Null } catch {}
    }
  })

  Append-Log $output "Starting..."
  Append-Log $output ("Root: {0}" -f $Root)
  Append-Log $output ("Repo: {0}/{1}@{2}" -f $RepoOwner,$RepoName,$Branch)

  Run-Background "Syncing..." { try { Sync-Manifest; $true } catch { $_ } } {
    param($r)
    if ($r -is [System.Management.Automation.ErrorRecord] -or $r -is [Exception]) {
      Append-Log $output ("Sync warning: {0}" -f $r.Exception.Message)
    } else { Append-Log $output "Catalogue synced." }
    Reload-Tools
    if ($lst.Items.Count -gt 0) { $lst.SelectedIndex = 0 }
  }

  [void]$form.ShowDialog()
}

# ---- ENTRY ----
Ensure-Tls12
Ensure-Folders
Log-Info ("Start: Command={0} Ui={1} UserInteractive={2} SessionId={3} Host={4}" -f $Command,$Ui,[Environment]::UserInteractive,([System.Diagnostics.Process]::GetCurrentProcess().SessionId),$Host.Name)

try {
  switch ($Command) {
    "init" { Write-Host "Initialised: $Root" }

    "sync" { Sync-Manifest; Write-Host "Saved: $ManifestLocal" }

    "list-remote" { (Read-Manifest).tools | Select-Object toolId,name,toolVersion,description,entryPoint }

    "install" {
      if (-not $ToolId) { throw "Usage: Toolkit.ps1 -Command install -ToolId <toolId>" }
      Install-ToolFromRepo $ToolId
      Write-Host "Installed: $ToolId"
    }

    "list-local" { Get-LocalInstalledToolIds }

    "run" {
      if (-not $ToolId) { throw "Usage: Toolkit.ps1 -Command run -ToolId <toolId>" }
      $r = Run-Tool $ToolId
      Write-Host "RunDir:   $($r.RunDir)"
      Write-Host "Report:   $($r.ReportPath)"
      Write-Host "ExitCode: $($r.ExitCode)"
      if ($ExitAfterRun) { exit $r.ExitCode }
    }

    "app" {
      # If we cannot show UI and cannot prompt, do something useful headlessly
      if (-not (Is-InteractiveDesktop)) {
        Invoke-HeadlessDefault
        return
      }

      $useWinForms = $false
      if ($Ui -eq "winforms") { $useWinForms = Can-UseWinForms }
      elseif ($Ui -eq "console") { $useWinForms = $false }
      else { $useWinForms = Can-UseWinForms }

      if ($useWinForms) { Invoke-AppWinForms } else { Invoke-AppConsole }
    }
  }
}
catch {
  Log-Info ("ERROR: " + $_.Exception.Message)
  Write-Host ""
  Write-Host "=== TOOLKIT ERROR ==="
  Write-Host $_.Exception.Message
}
