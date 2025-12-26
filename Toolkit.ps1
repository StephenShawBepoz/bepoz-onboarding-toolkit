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

function Ensure-Tls12 {
  try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}
}

function Ensure-Dirs {
  New-Item -ItemType Directory -Force -Path $Root,$CatDir,$ToolsDir,$RunsDir,$LogsDir,$TempDir | Out-Null
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
  Invoke-WebRequest -Uri $url -OutFile $outFile -UseBasicParsing
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
  $stamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
  $guid  = ([Guid]::NewGuid().ToString("N")).Substring(0,8)
  $dir   = Join-Path $RunsDir ("{0}_{1}_{2}" -f $stamp,$toolId,$guid)
  New-Item -ItemType Directory -Force -Path $dir | Out-Null
  return $dir
}

# ---------------- Cache / housekeeping ----------------
function Clear-ToolkitCache {
  param(
    [switch]$OnOpen,
    [switch]$OnClose,
    [int]$PruneRunsOlderThanDays = 30
  )

  Write-Host ("[Cache] Clearing cache ({0})..." -f ($(if($OnOpen){"open"}elseif($OnClose){"close"}else{"manual"})))

  $manifestPath = Join-Path $CatDir "manifest.json"
  if (Test-Path $manifestPath) {
    Remove-Item -Force $manifestPath -ErrorAction SilentlyContinue
    Write-Host "[Cache] Deleted cached manifest.json"
  }

  if (Test-Path $TempDir) {
    Get-ChildItem -Path $TempDir -Force -ErrorAction SilentlyContinue | Remove-Item -Force -Recurse -ErrorAction SilentlyContinue
    Write-Host "[Cache] Cleared temp folder"
  }

  if ($PruneRunsOlderThanDays -gt 0 -and (Test-Path $RunsDir)) {
    $cutoff = (Get-Date).AddDays(-$PruneRunsOlderThanDays)
    Get-ChildItem -Path $RunsDir -Directory -ErrorAction SilentlyContinue |
      Where-Object { $_.LastWriteTime -lt $cutoff } |
      ForEach-Object {
        try {
          Remove-Item -LiteralPath $_.FullName -Force -Recurse -ErrorAction Stop
          Write-Host "[Cache] Pruned old run: $($_.Name)"
        } catch {
          Write-Host "[Cache] Could not prune $($_.Name): $($_.Exception.Message)"
        }
      }
  }
}
# -----------------------------------------------------

function Run-ToolConsole([string]$toolId) {
  $m = Load-Manifest
  $t = Get-Tool -manifest $m -toolId $toolId
  Ensure-ToolInstalled -tool $t

  $runDir = New-RunDir -toolId $toolId
  $outLog = Join-Path $runDir "Output.log"
  $runner = Join-Path $runDir "RunTool.ps1"

  @"
`$ErrorActionPreference = 'Stop'
`$runDir = '$runDir'
`$env:BEPoz_TOOLKIT_RUNDIR = `$runDir
Push-Location '$Root'
try {
  & '$(Join-Path $Root ($t.entryPoint -replace "/","\"))' -RunDir `$runDir *>> '$outLog'
  exit `$LASTEXITCODE
} finally { Pop-Location }
"@ | Set-Content -Path $runner -Encoding UTF8

  Write-Host "RunDir: $runDir"
  $p = Start-Process -FilePath "powershell.exe" -ArgumentList "-NoProfile -ExecutionPolicy Bypass -STA -File `"$runner`"" -WorkingDirectory $Root -PassThru
  $p.WaitForExit()
  Write-Host "ExitCode: $($p.ExitCode)"
  Write-Host "Output:  $outLog"
}

function Show-Ui {
  Add-Type -AssemblyName System.Windows.Forms
  Add-Type -AssemblyName System.Drawing
  [System.Windows.Forms.Application]::EnableVisualStyles()

  Ensure-Dirs
  Clear-ToolkitCache -OnOpen -PruneRunsOlderThanDays 30

  $m = Load-Manifest

  $form = New-Object System.Windows.Forms.Form
  $form.Text = "Bepoz Onboarding Toolkit"
  $form.Width = 980
  $form.Height = 620
  $form.StartPosition = "CenterScreen"

  $form.Add_FormClosed({
    try { Clear-ToolkitCache -OnClose -PruneRunsOlderThanDays 30 } catch {}
  })

  $list = New-Object System.Windows.Forms.ListBox
  $list.Dock = "Left"
  $list.Width = 260

  $panel = New-Object System.Windows.Forms.Panel
  $panel.Dock = "Fill"

  $lblName = New-Object System.Windows.Forms.Label
  $lblName.AutoSize = $false
  $lblName.Height = 26
  $lblName.Dock = "Top"
  $lblName.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
  $lblName.Padding = "8,6,8,0"

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
  $outBox.ScrollBars = "Both"
  $outBox.Dock = "Fill"
  $outBox.Font = New-Object System.Drawing.Font("Consolas", 9)
  $tabOut.Controls.Add($outBox)

  $tabInfo = New-Object System.Windows.Forms.TabPage
  $tabInfo.Text = "Details"
  $infoBox = New-Object System.Windows.Forms.TextBox
  $infoBox.Multiline = $true
  $infoBox.ReadOnly = $true
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

  $btnRun = New-Object System.Windows.Forms.Button
  $btnRun.Text = "Run"
  $btnRun.Width = 90

  $btnRefresh = New-Object System.Windows.Forms.Button
  $btnRefresh.Text = "Refresh"
  $btnRefresh.Width = 90

  $btnOpenRuns = New-Object System.Windows.Forms.Button
  $btnOpenRuns.Text = "Open Runs Folder"
  $btnOpenRuns.Width = 140

  $btnPanel.Controls.AddRange(@($btnExit,$btnOpenRuns,$btnRun,$btnRefresh))

  $panel.Controls.Add($tabs)
  $panel.Controls.Add($txtDesc)
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

  $btnExit.Add_Click({ $form.Close() })

  $btnOpenRuns.Add_Click({
    try { Start-Process "explorer.exe" $RunsDir | Out-Null } catch {}
  })

  $btnRefresh.Add_Click({
    try {
      Clear-ToolkitCache -OnOpen -PruneRunsOlderThanDays 30

      $outBox.Clear()
      $infoBox.Clear()
      $lblName.Text = ""
      $txtDesc.Text = ""

      $list.Items.Clear()
      $toolMap.Clear()

      $script:m = Load-Manifest
      foreach ($t in $script:m.tools) {
        $display = "{0}  (v{1})" -f $t.toolId, $t.toolVersion
        $toolMap[$display] = $t
        [void]$list.Items.Add($display)
      }
    } catch {
      [System.Windows.Forms.MessageBox]::Show("Refresh failed: $($_.Exception.Message)")
    }
  })

  $list.Add_SelectedIndexChanged({
    $sel = $list.SelectedItem
    if (-not $sel) { return }
    $t = $toolMap[$sel]
    $lblName.Text = "{0}" -f $t.name
    $txtDesc.Text = "{0}`r`n`r`nToolId: {1}    Version: {2}`r`nEntryPoint: {3}" -f $t.description, $t.toolId, $t.toolVersion, $t.entryPoint

    $infoBox.Text = ($t | ConvertTo-Json -Depth 8)
    $outBox.Clear()
    $currentOutLog = $null
    $lastOutLen = 0
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
      $outLog = Join-Path $runDir "Output.log"
      $currentOutLog = $outLog
      $lastOutLen = 0

      $outBox.Text = "Running $($t.toolId)...`r`nRunDir: $runDir`r`n"
      $timer.Start()

      $runner = Join-Path $runDir "RunTool.ps1"
      @"
`$ErrorActionPreference = 'Stop'
`$runDir = '$runDir'
`$env:BEPoz_TOOLKIT_RUNDIR = `$runDir
Push-Location '$Root'
try {
  & '$(Join-Path $Root ($t.entryPoint -replace "/","\"))' -RunDir `$runDir *>> '$outLog'
  exit `$LASTEXITCODE
} finally { Pop-Location }
"@ | Set-Content -Path $runner -Encoding UTF8

      # IMPORTANT: -STA so WinForms dialogs (InputBox/MessageBox/etc) can display
      # IMPORTANT: WindowStyle Normal so you don't lose dialogs behind hidden windows
      Start-Process -FilePath "powershell.exe" -ArgumentList "-NoProfile -ExecutionPolicy Bypass -STA -File `"$runner`"" -WorkingDirectory $Root -WindowStyle Normal | Out-Null

    } catch {
      $timer.Stop()
      [System.Windows.Forms.MessageBox]::Show("Run failed: $($_.Exception.Message)")
    }
  })

  if ($list.Items.Count -gt 0) { $list.SelectedIndex = 0 }
  [void]$form.ShowDialog()
}

# ---------------- Main ----------------
Ensure-Dirs

try {
  switch ($Command) {
    "init" { Write-Host "Initialised: $Root" }
    "sync" { $null = Load-Manifest; Write-Host "Synced manifest to: $(Join-Path $CatDir 'manifest.json')" }
    "list" { $m = Load-Manifest; $m.tools | ForEach-Object { $_.toolId } }
    "run"  { if (-not $ToolId) { throw "Use: -Command run -ToolId <id>" }; Run-ToolConsole -toolId $ToolId }
    default { Show-Ui }
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
