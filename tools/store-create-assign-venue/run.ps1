#requires -Version 5.1
Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

param(
    [string]$ConfigPath,

    [string]$ToolkitRoot,
    [string]$ToolId,

    [Nullable[bool]]$DryRun,
    [switch]$Force,

    [string]$RepoZipUrl
)

function Test-PathWritable {
    param([string]$Path)
    try {
        if (-not (Test-Path -LiteralPath $Path)) { return $false }
        $probe = Join-Path $Path ("._writeprobe_" + [Guid]::NewGuid().ToString("N") + ".tmp")
        "probe" | Set-Content -LiteralPath $probe -Encoding UTF8 -ErrorAction Stop
        Remove-Item -LiteralPath $probe -Force -ErrorAction SilentlyContinue
        return $true
    } catch {
        return $false
    }
}

function Resolve-RunDir {
    # Prefer an explicit run directory if the launcher provides it
    foreach ($envName in @("TOOLKIT_RUNDIR","BEPOZ_RUNDIR","RUN_DIR","RUNDIR")) {
        $v = [Environment]::GetEnvironmentVariable($envName)
        if (-not [string]::IsNullOrWhiteSpace($v) -and (Test-Path -LiteralPath $v)) {
            if (Test-PathWritable -Path $v) { return (Resolve-Path -LiteralPath $v).Path }
        }
    }

    # Next best: current working directory (what the launcher likely sets per-job)
    try {
        $cwd = (Get-Location).Path
        if (-not [string]::IsNullOrWhiteSpace($cwd) -and (Test-Path -LiteralPath $cwd)) {
            if (Test-PathWritable -Path $cwd) { return (Resolve-Path -LiteralPath $cwd).Path }
        }
    } catch {}

    # Fallback: tool folder
    return (Resolve-Path -LiteralPath $PSScriptRoot).Path
}

$RunDir = Resolve-RunDir
$ReportPath = Join-Path $RunDir "Report.json"
$ToolLogPath = Join-Path $RunDir "Tool.log"
$MarkerPath = Join-Path $RunDir "Started.marker"

function Write-Log {
    param(
        [ValidateSet("INFO","WARN","ERROR","DEBUG")]
        [string]$Level,
        [string]$Message
    )
    $ts = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $line = "[$ts][$Level] $Message"
    Write-Host $line
    try { Add-Content -LiteralPath $ToolLogPath -Value $line -Encoding UTF8 -ErrorAction SilentlyContinue } catch {}
}

function Save-Report {
    param([hashtable]$Report)
    try {
        ($Report | ConvertTo-Json -Depth 15) | Set-Content -Path $ReportPath -Encoding UTF8
        Write-Log INFO "Wrote report: $ReportPath"
    } catch {
        Write-Log ERROR "Failed to write Report.json: $($_.Exception.Message)"
    }
}

function Test-IsInteractive {
    try {
        if (-not [Environment]::UserInteractive) { return $false }
        if ($Host -and $Host.Name -eq "ConsoleHost") { return $true }
        return $false
    } catch { return $false }
}

function Read-JsonFile {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return $null }
    if (-not (Test-Path -LiteralPath $Path)) { return $null }
    $raw = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop
    if ([string]::IsNullOrWhiteSpace($raw)) { return $null }
    return ($raw | ConvertFrom-Json -ErrorAction Stop)
}

function Get-EnvBool([string]$Name) {
    $v = [Environment]::GetEnvironmentVariable($Name)
    if ([string]::IsNullOrWhiteSpace($v)) { return $null }
    switch -Regex ($v.Trim().ToLowerInvariant()) {
        '^(1|true|yes|y)$' { return $true }
        '^(0|false|no|n)$' { return $false }
        default { return $null }
    }
}

function Get-EnvString([string]$Name) {
    $v = [Environment]::GetEnvironmentVariable($Name)
    if ([string]::IsNullOrWhiteSpace($v)) { return $null }
    return $v.Trim()
}

function Confirm-Apply([string]$Message) {
    if ($Force.IsPresent) { return $true }
    if (-not (Test-IsInteractive)) { return $false }

    Write-Host ""
    Write-Host $Message
    Write-Host "Type REPAIR to proceed, anything else to cancel: " -NoNewline
    $resp = Read-Host
    return ($resp -eq "REPAIR")
}

function Resolve-ToolkitRoot {
    param([string]$ProvidedRoot)

    if (-not [string]::IsNullOrWhiteSpace($ProvidedRoot)) {
        return (Resolve-Path -LiteralPath $ProvidedRoot).Path
    }

    # Heuristic: tools\<this tool>\run.ps1 -> ToolkitRoot is parent of "tools"
    $p = (Resolve-Path -LiteralPath $PSScriptRoot).Path
    $dir = New-Object System.IO.DirectoryInfo($p)
    while ($dir -ne $null) {
        if ($dir.Name -ieq "tools") {
            return $dir.Parent.FullName
        }
        $dir = $dir.Parent
    }

    $default = "C:\Bepoz\OnboardingToolkit"
    if (Test-Path -LiteralPath $default) { return $default }

    throw "Could not resolve ToolkitRoot. Provide -ToolkitRoot or set BEPOZ_TOOLKIT_ROOT."
}

function Try-GitPull {
    param([string]$Root)

    $gitDir = Join-Path $Root ".git"
    if (-not (Test-Path -LiteralPath $gitDir)) {
        return @{ attempted = $false; ok = $false; message = "Not a git clone (no .git folder)." }
    }

    $git = Get-Command git.exe -ErrorAction SilentlyContinue
    if (-not $git) {
        return @{ attempted = $false; ok = $false; message = "git.exe not found in PATH." }
    }

    try {
        Write-Log INFO "Attempting git pull in $Root ..."
        $p = Start-Process -FilePath $git.Source -ArgumentList @("-C", $Root, "pull", "--ff-only") -NoNewWindow -Wait -PassThru
        if ($p.ExitCode -ne 0) {
            return @{ attempted = $true; ok = $false; message = "git pull failed with exit code $($p.ExitCode)." }
        }
        return @{ attempted = $true; ok = $true; message = "git pull completed." }
    } catch {
        return @{ attempted = $true; ok = $false; message = "git pull threw: $($_.Exception.Message)" }
    }
}

function Try-ToolkitSync {
    param([string]$Root)

    $toolkitPs1 = Join-Path $Root "Toolkit.ps1"
    if (-not (Test-Path -LiteralPath $toolkitPs1)) {
        return @{ attempted = $false; ok = $false; message = "Toolkit.ps1 not found at root." }
    }

    try {
        Write-Log INFO "Attempting Toolkit.ps1 -Sync ..."
        $args = @("-NoProfile","-ExecutionPolicy","Bypass","-File",$toolkitPs1,"-Sync")
        $p = Start-Process -FilePath "powershell.exe" -ArgumentList $args -NoNewWindow -Wait -PassThru
        if ($p.ExitCode -ne 0) {
            return @{ attempted = $true; ok = $false; message = "Toolkit.ps1 -Sync failed with exit code $($p.ExitCode)." }
        }
        return @{ attempted = $true; ok = $true; message = "Toolkit.ps1 -Sync completed." }
    } catch {
        return @{ attempted = $true; ok = $false; message = "Toolkit.ps1 -Sync threw: $($_.Exception.Message)" }
    }
}

function Try-ZipRepair {
    param(
        [string]$Root,
        [string]$ZipUrl,
        [string]$ToolFolderRelative
    )

    if ([string]::IsNullOrWhiteSpace($ZipUrl)) {
        return @{ attempted = $false; ok = $false; message = "RepoZipUrl not provided." }
    }

    try {
        $tmp = Join-Path $env:TEMP ("BepozToolkitZip_" + [Guid]::NewGuid().ToString("N"))
        New-Item -ItemType Directory -Path $tmp -Force | Out-Null

        $zipPath = Join-Path $tmp "repo.zip"
        Write-Log INFO "Downloading repo ZIP..."
        Invoke-WebRequest -Uri $ZipUrl -OutFile $zipPath -UseBasicParsing -ErrorAction Stop

        $extract = Join-Path $tmp "extract"
        New-Item -ItemType Directory -Path $extract -Force | Out-Null

        Add-Type -AssemblyName System.IO.Compression.FileSystem
        [System.IO.Compression.ZipFile]::ExtractToDirectory($zipPath, $extract)

        $top = Get-ChildItem -LiteralPath $extract -Directory | Select-Object -First 1
        if (-not $top) { throw "ZIP extract did not contain a top-level folder." }

        $srcToolPath = Join-Path $top.FullName $ToolFolderRelative
        if (-not (Test-Path -LiteralPath $srcToolPath)) {
            throw "Tool folder not found in ZIP at: $ToolFolderRelative"
        }

        $dstToolPath = Join-Path $Root $ToolFolderRelative
        if (-not (Test-Path -LiteralPath $dstToolPath)) {
            New-Item -ItemType Directory -Path $dstToolPath -Force | Out-Null
        }

        Write-Log INFO "Copying missing tool folder from ZIP..."
        Copy-Item -Path (Join-Path $srcToolPath "*") -Destination $dstToolPath -Recurse -Force -ErrorAction Stop

        return @{ attempted = $true; ok = $true; message = "ZIP repair copied tool folder successfully." }
    } catch {
        return @{ attempted = $true; ok = $false; message = "ZIP repair failed: $($_.Exception.Message)" }
    } finally {
        try { if ($tmp -and (Test-Path $tmp)) { Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue } } catch {}
    }
}

$report = @{
    toolId    = "toolkit-sync-verify"
    startedAt = (Get-Date).ToString("o")
    runDir    = $RunDir
    resolved  = @{}
    checks    = @()
    repairs   = @()
    result    = @{
        status  = "Unknown"
        message = ""
        toolId  = $null
        entryPoint = $null
        entryPointFullPath = $null
    }
}

try {
    # Marker proves the script actually started in the intended RunDir
    "started $(Get-Date -Format o)" | Set-Content -LiteralPath $MarkerPath -Encoding UTF8 -Force

    Write-Log INFO "Toolkit Sync + Verify starting..."
    Write-Log INFO "Resolved RunDir: $RunDir"

    if ([string]::IsNullOrWhiteSpace($ConfigPath)) {
        $ConfigPath = Join-Path $PSScriptRoot "ToolkitSyncVerify.config.json"
    }
    $cfg = Read-JsonFile -Path $ConfigPath

    $envRoot  = Get-EnvString "BEPOZ_TOOLKIT_ROOT"
    $envTool  = Get-EnvString "BEPOZ_VERIFY_TOOLID"
    $envDry   = Get-EnvBool   "BEPOZ_DRYRUN"
    $envForce = Get-EnvBool   "BEPOZ_FORCE"
    $envZip   = Get-EnvString "BEPOZ_REPO_ZIP_URL"

    if ([string]::IsNullOrWhiteSpace($ToolkitRoot)) {
        if ($cfg -and $cfg.ToolkitRoot) { $ToolkitRoot = [string]$cfg.ToolkitRoot }
        elseif ($envRoot) { $ToolkitRoot = $envRoot }
    }

    if ([string]::IsNullOrWhiteSpace($ToolId)) {
        if ($cfg -and $cfg.ToolId) { $ToolId = [string]$cfg.ToolId }
        elseif ($envTool) { $ToolId = $envTool }
        else { $ToolId = "store-create-assign-venue" }
    }

    if ($DryRun -eq $null) {
        if ($cfg -and $cfg.DryRun -ne $null) { $DryRun = [bool]$cfg.DryRun }
        elseif ($envDry -ne $null) { $DryRun = [bool]$envDry }
        else { $DryRun = $true }
    }

    if (-not $Force.IsPresent) {
        if ($cfg -and $cfg.Force -eq $true) { $Force = $true }
        elseif ($envForce -eq $true) { $Force = $true }
    }

    if ([string]::IsNullOrWhiteSpace($RepoZipUrl)) {
        if ($cfg -and $cfg.RepoZipUrl) { $RepoZipUrl = [string]$cfg.RepoZipUrl }
        elseif ($envZip) { $RepoZipUrl = $envZip }
    }

    $ToolkitRoot = Resolve-ToolkitRoot -ProvidedRoot $ToolkitRoot

    $report.resolved = @{
        toolkitRoot = $ToolkitRoot
        toolId = $ToolId
        dryRun = [bool]$DryRun
        force = [bool]$Force.IsPresent
        hasRepoZipUrl = -not [string]::IsNullOrWhiteSpace($RepoZipUrl)
        configPath = $ConfigPath
    }

    Write-Log INFO "ToolkitRoot: $ToolkitRoot"
    Write-Log INFO "ToolId: $ToolId"
    Write-Log INFO ("DryRun: " + $DryRun.ToString() + " (Force: " + ([bool]$Force.IsPresent).ToString() + ")")

    $manifestPath = Join-Path $ToolkitRoot "manifest.json"
    if (-not (Test-Path -LiteralPath $manifestPath)) {
        throw "manifest.json not found at ToolkitRoot: $manifestPath"
    }

    $manifest = Read-JsonFile -Path $manifestPath
    if (-not $manifest -or -not $manifest.tools) {
        throw "manifest.json is missing 'tools' array or could not be parsed."
    }

    $tool = $manifest.tools | Where-Object { $_.toolId -eq $ToolId } | Select-Object -First 1
    $report.checks += @{ name = "ToolExistsInManifest"; ok = [bool]$tool }
    if (-not $tool) {
        $report.result.status = "Error"
        $report.result.message = "ToolId '$ToolId' not found in manifest.json."
        Write-Log ERROR $report.result.message
        return
    }

    $entryPoint = [string]$tool.entryPoint
    $report.result.toolId = $ToolId
    $report.result.entryPoint = $entryPoint

    if ([string]::IsNullOrWhiteSpace($entryPoint)) {
        throw "Tool '$ToolId' has an empty entryPoint in manifest.json."
    }

    $entryFull = Join-Path $ToolkitRoot $entryPoint
    $report.result.entryPointFullPath = $entryFull

    $exists = Test-Path -LiteralPath $entryFull
    $report.checks += @{
        name = "EntryPointExistsOnDisk"
        entryPoint = $entryPoint
        fullPath = $entryFull
        ok = [bool]$exists
    }

    if ($exists) {
        $report.result.status = "Success"
        $report.result.message = "EntryPoint exists on disk."
        Write-Log INFO $report.result.message
        return
    }

    Write-Log WARN "EntryPoint is missing on disk: $entryFull"

    if ($DryRun) {
        $report.result.status = "Missing"
        $report.result.message = "Missing entryPoint. DryRun enabled; no repair attempted."
        Write-Log WARN $report.result.message
        Write-Log INFO "Re-run with -DryRun:`$false (and -Force for non-interactive) to attempt repair."
        return
    }

    $okToRepair = Confirm-Apply -Message "Repair will try to sync/pull and/or copy files. Proceed?"
    if (-not $okToRepair) {
        $report.result.status = "Cancelled"
        $report.result.message = "Repair not confirmed (or non-interactive without -Force)."
        Write-Log WARN $report.result.message
        return
    }

    $toolFolderRel = Split-Path -Path $entryPoint -Parent

    $repair1 = Try-GitPull -Root $ToolkitRoot
    $report.repairs += @{ name="gitPull"; details=$repair1 }

    if (-not (Test-Path -LiteralPath $entryFull)) {
        $repair2 = Try-ToolkitSync -Root $ToolkitRoot
        $report.repairs += @{ name="toolkitSync"; details=$repair2 }
    }

    if (-not (Test-Path -LiteralPath $entryFull)) {
        $repair3 = Try-ZipRepair -Root $ToolkitRoot -ZipUrl $RepoZipUrl -ToolFolderRelative $toolFolderRel
        $report.repairs += @{ name="zipRepair"; details=$repair3 }
    }

    $existsAfter = Test-Path -LiteralPath $entryFull
    $report.checks += @{ name = "EntryPointExistsAfterRepair"; ok = [bool]$existsAfter }

    if ($existsAfter) {
        $report.result.status = "Success"
        $report.result.message = "EntryPoint is now present after repair."
        Write-Log INFO $report.result.message
    } else {
        $report.result.status = "Error"
        $report.result.message = "EntryPoint still missing after repair attempts."
        Write-Log ERROR $report.result.message
    }

} catch {
    $report.result.status = "Error"
    $report.result.message = $_.Exception.Message
    Write-Log ERROR $report.result.message
} finally {
    $report.finishedAt = (Get-Date).ToString("o")
    Save-Report -Report $report
}

if ($report.result.status -eq "Error") { exit 1 } else { exit 0 }
